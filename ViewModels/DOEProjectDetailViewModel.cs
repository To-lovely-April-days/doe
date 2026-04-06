using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Models;
using Newtonsoft.Json;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media;

namespace MaxChemical.Modules.DOE.ViewModels
{
    public class DOEProjectDetailViewModel : BindableBase
    {
        private readonly IDOERepository _repository;

        // ── 基础信息 ──
        public string ProjectName { get; set; } = "";
        public string PhaseText { get; set; } = "";
        public string PhaseColor { get; set; } = "#7F8C8D";
        public string StatusText { get; set; } = "";
        public string BestValueText { get; set; } = "—";
        public string ObjectiveDescription { get; set; } = "";
        public int TotalExperiments { get; set; }
        public int CompletedRounds { get; set; }
        public DateTime CreatedTime { get; set; }
        public DateTime UpdatedTime { get; set; }
        public string PhaseProgressText { get; set; } = "";

        // ── 进度条 ──
        public ObservableCollection<PhaseSegmentItem> PhaseSegments { get; set; } = new();

        // ── 最优因子组合 ──
        public ObservableCollection<BestFactorItem> BestFactorItems { get; set; } = new();

        // ── 因子池表格 ──
        public ObservableCollection<FactorTableItem> FactorTableItems { get; set; } = new();

        // ── 模型列表 ──
        public ObservableCollection<ModelListItem> ModelItems { get; set; } = new();

        // ── 批次列表 ──
        public ObservableCollection<BatchListItem> BatchItems { get; set; } = new();

        public DOEProjectDetailViewModel(IDOERepository repository)
        {
            _repository = repository;
        }

        public async Task LoadAsync(string projectId)
        {
            var project = await _repository.GetProjectWithDetailsAsync(projectId);
            if (project == null) return;

            // 基础信息
            ProjectName = project.ProjectName;
            ObjectiveDescription = string.IsNullOrEmpty(project.ObjectiveDescription) ? "—" : project.ObjectiveDescription;
            TotalExperiments = project.TotalExperiments;
            CompletedRounds = project.CompletedRounds;
            CreatedTime = project.CreatedTime;
            UpdatedTime = project.UpdatedTime;
            BestValueText = project.BestResponseValue.HasValue ? $"{project.BestResponseValue:F2}" : "—";

            PhaseText = project.CurrentPhase switch
            {
                DOEProjectPhase.Screening => "筛选",
                DOEProjectPhase.PathSearch => "爬坡",
                DOEProjectPhase.RSM => "响应面优化",
                DOEProjectPhase.Augmenting => "补点",
                DOEProjectPhase.Confirmation => "验证",
                DOEProjectPhase.Completed => "已完成",
                _ => project.CurrentPhase.ToString()
            };

            PhaseColor = project.CurrentPhase switch
            {
                DOEProjectPhase.Screening => "#E67E22",
                DOEProjectPhase.PathSearch => "#3498DB",
                DOEProjectPhase.RSM => "#9B59B6",
                DOEProjectPhase.Augmenting => "#1ABC9C",
                DOEProjectPhase.Confirmation => "#27AE60",
                DOEProjectPhase.Completed => "#95A5A6",
                _ => "#7F8C8D"
            };

            StatusText = project.Status switch
            {
                DOEProjectStatus.Active => "活跃",
                DOEProjectStatus.Paused => "暂停",
                DOEProjectStatus.Completed => "已完成",
                DOEProjectStatus.Archived => "已归档",
                _ => ""
            };

            // 进度条
            PhaseSegments = BuildPhaseSegments(project.CurrentPhase);
            PhaseProgressText = project.CurrentPhase == DOEProjectPhase.Completed
                ? "已完成全部优化流程"
                : "筛选 → 爬坡 → RSM → 验证 → 完成";

            // 最优因子组合
            BestFactorItems = ParseBestFactors(project.BestFactorsJson);

            // 因子池
            FactorTableItems = BuildFactorTable(project.ProjectFactors);

            // 模型列表
            await LoadModelsAsync(project);

            // 批次列表
            BatchItems = await BuildBatchListAsync(project.Batches, project.RoundSummaries);

            RaisePropertyChanged(string.Empty); // 刷新所有属性
        }

        private ObservableCollection<BestFactorItem> ParseBestFactors(string? json)
        {
            var items = new ObservableCollection<BestFactorItem>();
            if (string.IsNullOrEmpty(json)) return items;

            try
            {
                var dict = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                if (dict != null)
                {
                    foreach (var kv in dict)
                    {
                        items.Add(new BestFactorItem { Name = kv.Key, Value = kv.Value?.ToString() ?? "—" });
                    }
                }
            }
            catch { }
            return items;
        }

        private ObservableCollection<FactorTableItem> BuildFactorTable(List<DOEProjectFactor> factors)
        {
            var items = new ObservableCollection<FactorTableItem>();
            if (factors == null) return items;

            // 活跃因子排前面
            foreach (var f in factors.OrderBy(f => f.FactorStatus).ThenBy(f => f.SortOrder))
            {
                string range;
                if (f.IsCategorical)
                    range = "{" + (f.CategoryLevels ?? "—") + "}";
                else
                    range = $"[{f.CurrentLowerBound}, {f.CurrentUpperBound}]";

                string status;
                bool isActive;
                switch (f.FactorStatus)
                {
                    case ProjectFactorStatus.Active:
                        status = "活跃";
                        isActive = true;
                        break;
                    case ProjectFactorStatus.ScreenedOut:
                        status = f.StatusReason ?? "已淘汰";
                        isActive = false;
                        break;
                    case ProjectFactorStatus.Fixed:
                        var fixVal = f.FixedCategoryLevel ?? f.FixedValue?.ToString("F1") ?? "";
                        status = $"固定={fixVal}";
                        isActive = false;
                        break;
                    default:
                        status = f.FactorStatus.ToString();
                        isActive = true;
                        break;
                }

                items.Add(new FactorTableItem
                {
                    Name = f.FactorName,
                    Range = isActive ? range : "—",
                    Status = status,
                    IsActive = isActive
                });
            }
            return items;
        }

        private async Task LoadModelsAsync(DOEProject project)
        {
            var items = new ObservableCollection<ModelListItem>();
            try
            {
                // ★ 按项目 ID 查所有 GPR 模型
                var models = await _repository.GetGPRModelsByProjectAsync(project.ProjectId);

                foreach (var m in models)
                {
                    var name = m.ModelName;
                    if (string.IsNullOrEmpty(name))
                    {
                        var sigParts = m.FactorSignature?.Split("::") ?? Array.Empty<string>();
                        name = sigParts.Length > 1 ? sigParts[1] : "主响应模型";
                    }

                    var factorSig = m.FactorSignature ?? "";
                    if (factorSig.Contains("::"))
                        factorSig = factorSig.Substring(0, factorSig.IndexOf("::"));

                    items.Add(new ModelListItem
                    {
                        Name = name,
                        Description = $"{factorSig} · {m.DataCount} 组数据",
                        RSquaredText = m.RSquared.HasValue ? $"{m.RSquared:F3}" : "—"
                    });
                }

                // 如果也有 FlowId，补充查 FlowId 下没有 ProjectId 的旧模型
                if (!string.IsNullOrEmpty(project.FlowId))
                {
                    var flowModels = await _repository.GetGPRModelsByFlowAsync(project.FlowId);
                    foreach (var m in flowModels)
                    {
                        // 跳过已经按 ProjectId 查到的（避免重复）
                        if (!string.IsNullOrEmpty(m.ProjectId)) continue;

                        var name = m.ModelName;
                        if (string.IsNullOrEmpty(name))
                        {
                            var sigParts = m.FactorSignature?.Split("::") ?? Array.Empty<string>();
                            name = sigParts.Length > 1 ? sigParts[1] : "主响应模型";
                        }

                        var factorSig = m.FactorSignature ?? "";
                        if (factorSig.Contains("::"))
                            factorSig = factorSig.Substring(0, factorSig.IndexOf("::"));

                        items.Add(new ModelListItem
                        {
                            Name = name,
                            Description = $"{factorSig} · {m.DataCount} 组数据",
                            RSquaredText = m.RSquared.HasValue ? $"{m.RSquared:F3}" : "—"
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"加载 GPR 模型列表失败: {ex.Message}");
            }

            if (items.Count == 0)
            {
                items.Add(new ModelListItem
                {
                    Name = "暂无模型",
                    Description = "执行实验后将自动创建 GPR 模型",
                    RSquaredText = "—"
                });
            }

            ModelItems = items;
        }


        private async Task<ObservableCollection<BatchListItem>> BuildBatchListAsync(
      List<DOEBatch> batches, List<DOERoundSummary> summaries)
        {
            var items = new ObservableCollection<BatchListItem>();
            if (batches == null) return items;

            foreach (var b in batches.OrderByDescending(b => b.RoundNumber ?? 0))
            {
                var summary = summaries?.FirstOrDefault(s => s.BatchId == b.BatchId);

                // ★ 查询实际 Runs 数量
                int runCount = 0;
                try
                {
                    var runs = await _repository.GetRunsAsync(b.BatchId);
                    runCount = runs?.Count ?? 0;
                }
                catch { }

                var phase = b.ProjectPhase ?? DOEProjectPhase.Screening;
                var phaseText = phase switch
                {
                    DOEProjectPhase.Screening => "筛选",
                    DOEProjectPhase.PathSearch => "爬坡",
                    DOEProjectPhase.RSM => "RSM",
                    DOEProjectPhase.Augmenting => "补点",
                    DOEProjectPhase.Confirmation => "验证",
                    _ => phase.ToString()
                };
                var phaseColor = phase switch
                {
                    DOEProjectPhase.Screening => "#E67E22",
                    DOEProjectPhase.PathSearch => "#3498DB",
                    DOEProjectPhase.RSM => "#9B59B6",
                    DOEProjectPhase.Augmenting => "#1ABC9C",
                    DOEProjectPhase.Confirmation => "#27AE60",
                    _ => "#7F8C8D"
                };

                var methodText = b.DesignMethod switch
                {
                    DOEDesignMethod.FullFactorial => "全因子",
                    DOEDesignMethod.FractionalFactorial => "部分因子",
                    DOEDesignMethod.Taguchi => "Taguchi",
                    DOEDesignMethod.CCD => "CCD",
                    DOEDesignMethod.BoxBehnken => "BBD",
                    DOEDesignMethod.DOptimal => "D-Optimal",
                    DOEDesignMethod.SteepestAscent => "最速上升",
                    DOEDesignMethod.AugmentedDesign => "增强设计",
                    DOEDesignMethod.ConfirmationRuns => "验证实验",
                    _ => b.DesignMethod.ToString()
                };

                items.Add(new BatchListItem
                {
                    RoundLabel = $"R{b.RoundNumber ?? 0}",
                    PhaseText = phaseText,
                    PhaseColor = phaseColor,
                    MethodAndRuns = $"{methodText} · {runCount} 组",
                    RSquaredText = summary?.RSquared.HasValue == true ? $"R²={summary.RSquared:F3}" : "—",
                    DateText = b.CreatedTime.ToString("MM-dd")
                });
            }
            return items;
        }

        private static ObservableCollection<PhaseSegmentItem> BuildPhaseSegments(DOEProjectPhase currentPhase)
        {
            int currentIndex = currentPhase switch
            {
                DOEProjectPhase.Screening => 0,
                DOEProjectPhase.PathSearch => 1,
                DOEProjectPhase.RSM => 2,
                DOEProjectPhase.Augmenting => 2,
                DOEProjectPhase.Confirmation => 3,
                DOEProjectPhase.Completed => 5,
                _ => 0
            };
            var active = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0078D4"));
            var current = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0078D4")) { Opacity = 0.35 };
            var inactive = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8EAED"));
            var segments = new ObservableCollection<PhaseSegmentItem>();
            for (int i = 0; i < 5; i++)
            {
                if (i < currentIndex) segments.Add(new PhaseSegmentItem { Color = active });
                else if (i == currentIndex) segments.Add(new PhaseSegmentItem { Color = current });
                else segments.Add(new PhaseSegmentItem { Color = inactive });
            }
            return segments;
        }
    }

    // ── 数据项类 ──

    public class BestFactorItem
    {
        public string Name { get; set; } = "";
        public string Value { get; set; } = "";
    }

    public class FactorTableItem
    {
        public string Name { get; set; } = "";
        public string Range { get; set; } = "";
        public string Status { get; set; } = "";
        public bool IsActive { get; set; } = true;
    }

    public class ModelListItem
    {
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
        public string RSquaredText { get; set; } = "—";
    }

    public class BatchListItem
    {
        public string RoundLabel { get; set; } = "";
        public string PhaseText { get; set; } = "";
        public string PhaseColor { get; set; } = "#7F8C8D";
        public string MethodAndRuns { get; set; } = "";
        public string RSquaredText { get; set; } = "—";
        public string DateText { get; set; } = "";
    }
}