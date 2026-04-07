using MaxChemical.Infrastructure.DOE;
using MaxChemical.Infrastructure.Services;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Models;
using MaxChemical.Modules.DOE.Services;
using Newtonsoft.Json;
using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace MaxChemical.Modules.DOE.ViewModels
{
    /// <summary>
    /// DOE 历史管理 ViewModel — 项目模式
    /// 
    /// ★ 重写: 从批次平铺改为 项目列表 → 展开轮次 两级结构
    /// 
    /// 左侧: 项目列表（含状态、阶段、最优值）
    /// 右侧: 选中项目的轮次列表 + 轮次详情 + 轮次总结
    /// </summary>
    public class DOEHistoryViewModel : BindableBase
    {
        private readonly IDOERepository _repository;
        private readonly IGPRModelService _gprService;
        private readonly DOEExportService _exportService;
        private readonly IFlowParameterProvider _paramProvider;
        private readonly IDialogService _dialogService;
        private readonly ILogService _logger;

        // ── 项目列表（左侧）──
        private ObservableCollection<DOEProjectSummary> _projects = new();
        private DOEProjectSummary? _selectedProject;

        // ── 轮次列表（右侧）──
        private ObservableCollection<RoundListItem> _rounds = new();
        private RoundListItem? _selectedRound;
        private DOEBatch? _roundDetail;

        // ── 轮次总结 ──
        private DOERoundSummary? _roundSummary;

        private bool _isLoading;
        private string _statusMessage = "";

        // GPR 模型状态
        private string _gprStatusText = "未初始化";
        private bool _gprIsActive;
        private int _gprDataCount;

        private bool _hasSelectedRound;
        private string _selectedRoundTitle = "";
        private string _selectedRoundR2Text = "—";
        private int _selectedRoundRunCount;
        private string _selectedRoundRecommendation = "";
        private bool _hasSelectedRoundRecommendation;
        private DataTable _selectedRoundRuns = new();
        public DOEHistoryViewModel(
            IDOERepository repository,
            IGPRModelService gprService,
            DOEExportService exportService,
            IFlowParameterProvider paramProvider,
            IDialogService dialogService,
            ILogService logger)
        {
            _repository = repository;
            _gprService = gprService;
            _exportService = exportService;
            _paramProvider = paramProvider;
            _dialogService = dialogService;
            _logger = logger?.ForContext<DOEHistoryViewModel>() ?? throw new ArgumentNullException(nameof(logger));

            RefreshCommand = new DelegateCommand(async () => await LoadProjectsAsync());
            DeleteProjectCommand = new DelegateCommand<DOEProjectSummary>(async s => await DeleteProjectAsync(s));
            DeleteRoundCommand = new DelegateCommand<RoundListItem>(async r => await DeleteRoundAsync(r));
            ExportRoundCommand = new DelegateCommand<RoundListItem>(async r => await ExportRoundAsync(r));
            NavigateToExecutionCommand = new DelegateCommand(
                () => { if (_selectedRound != null) RequestExecuteBatch?.Invoke(this, _selectedRound.BatchId); },
                () => _selectedRound != null);
            NavigateToAnalysisCommand = new DelegateCommand(
                () => { if (_selectedRound != null) RequestAnalyzeBatch?.Invoke(this, _selectedRound.BatchId); },
                () => _selectedRound != null);
            SelectRoundCommand = new DelegateCommand<RoundListItem>(r => { if (r != null) SelectedRound = r; });
            // 初始加载
            _ = LoadProjectsAsync();
        }

        // ══════════════ Properties ══════════════

        // 项目列表
        public ObservableCollection<DOEProjectSummary> Projects { get => _projects; set => SetProperty(ref _projects, value); }
        public DOEProjectSummary? SelectedProject
        {
            get => _selectedProject;
            set { SetProperty(ref _selectedProject, value); _ = OnSelectedProjectChangedAsync(); }
        }

        // 轮次列表
        public ObservableCollection<RoundListItem> Rounds { get => _rounds; set => SetProperty(ref _rounds, value); }
        public RoundListItem? SelectedRound
        {
            get => _selectedRound;
            set
            {
                SetProperty(ref _selectedRound, value);
                _ = OnSelectedRoundChangedAsync();
                NavigateToExecutionCommand.RaiseCanExecuteChanged();
                NavigateToAnalysisCommand.RaiseCanExecuteChanged();
            }
        }

        // 轮次详情
        public DOEBatch? RoundDetail { get => _roundDetail; set => SetProperty(ref _roundDetail, value); }
        public DOERoundSummary? RoundSummary { get => _roundSummary; set => SetProperty(ref _roundSummary, value); }

        public bool IsLoading { get => _isLoading; set => SetProperty(ref _isLoading, value); }
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
        public string GPRStatusText { get => _gprStatusText; set => SetProperty(ref _gprStatusText, value); }
        public bool GPRIsActive { get => _gprIsActive; set => SetProperty(ref _gprIsActive, value); }
        public int GPRDataCount { get => _gprDataCount; set => SetProperty(ref _gprDataCount, value); }

        /// <summary>是否有选中的项目（控制右侧面板显示）</summary>
        public bool HasSelectedProject => SelectedProject != null;

        /// <summary>是否有轮次总结（控制总结面板显示）</summary>
        public bool HasRoundSummary => RoundSummary != null;
        public bool HasSelectedRound { get => _hasSelectedRound; set => SetProperty(ref _hasSelectedRound, value); }
        public string SelectedRoundTitle { get => _selectedRoundTitle; set => SetProperty(ref _selectedRoundTitle, value); }
        public string SelectedRoundR2Text { get => _selectedRoundR2Text; set => SetProperty(ref _selectedRoundR2Text, value); }
        public int SelectedRoundRunCount { get => _selectedRoundRunCount; set => SetProperty(ref _selectedRoundRunCount, value); }
        public string SelectedRoundRecommendation { get => _selectedRoundRecommendation; set => SetProperty(ref _selectedRoundRecommendation, value); }
        public bool HasSelectedRoundRecommendation { get => _hasSelectedRoundRecommendation; set => SetProperty(ref _hasSelectedRoundRecommendation, value); }
        public DataTable SelectedRoundRuns { get => _selectedRoundRuns; set => SetProperty(ref _selectedRoundRuns, value); }

        // ══════════════ Commands ══════════════

        public DelegateCommand RefreshCommand { get; }
        public DelegateCommand<DOEProjectSummary> DeleteProjectCommand { get; }
        public DelegateCommand<RoundListItem> DeleteRoundCommand { get; }
        public DelegateCommand<RoundListItem> ExportRoundCommand { get; }
        public DelegateCommand NavigateToExecutionCommand { get; }
        public DelegateCommand NavigateToAnalysisCommand { get; }
        public DelegateCommand<RoundListItem> SelectRoundCommand { get; }
        // ══════════════ Events ══════════════

        public event EventHandler<string>? RequestExecuteBatch;
        public event EventHandler<string>? RequestAnalyzeBatch;

        // ══════════════ 加载 ══════════════

        /// <summary>加载项目列表（原 LoadBatchesAsync 重写）</summary>
        public async Task LoadProjectsAsync()
        {
            try
            {
                IsLoading = true;
                var projects = await _repository.GetAllProjectSummariesAsync(50);
                Projects = new ObservableCollection<DOEProjectSummary>(projects);
                StatusMessage = $"共 {projects.Count} 个项目";

                // GPR 状态
                var flowId = _paramProvider.GetCurrentFlowId();
                if (!string.IsNullOrEmpty(flowId))
                {
                    var state = await _repository.GetGPRModelStateAsync(flowId);
                    if (state != null)
                    {
                        GPRIsActive = state.IsActive;
                        GPRDataCount = state.DataCount;
                        GPRStatusText = state.IsActive
                            ? $"已激活 | R²={state.RSquared:F3} | {state.DataCount} 组数据"
                            : $"未激活 | {state.DataCount} 组数据";
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载项目列表失败");
                StatusMessage = $"加载失败: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        /// <summary>兼容旧调用点（DOEMainView.xaml.cs 中 case 3 调用的）</summary>
        public async Task LoadBatchesAsync() => await LoadProjectsAsync();

        private async Task OnSelectedProjectChangedAsync()
        {
            RaisePropertyChanged(nameof(HasSelectedProject));
            RoundDetail = null;
            RoundSummary = null;
            RaisePropertyChanged(nameof(HasRoundSummary));

            if (SelectedProject == null)
            {
                Rounds = new ObservableCollection<RoundListItem>();
                return;
            }

            try
            {
                IsLoading = true;
                var batches = await _repository.GetBatchesByProjectAsync(SelectedProject.ProjectId);
                var summaries = await _repository.GetRoundSummariesAsync(SelectedProject.ProjectId);

                // ★ 修复: 为每个批次查询 Runs 统计（GetBatchesByProjectAsync 不含 Runs）
                var roundItems = new List<RoundListItem>();
                foreach (var b in batches)
                {
                    var summary = summaries.FirstOrDefault(s => s.BatchId == b.BatchId);

                    // 查询该批次的 Runs 统计
                    int totalRuns = 0;
                    int completedRuns = 0;
                    try
                    {
                        var runs = await _repository.GetRunsAsync(b.BatchId);
                        totalRuns = runs.Count;
                        completedRuns = runs.Count(r => r.Status == DOERunStatus.Completed);
                    }
                    catch { }

                    roundItems.Add(new RoundListItem
                    {
                        BatchId = b.BatchId,
                        BatchName = b.BatchName,
                        RoundNumber = b.RoundNumber ?? 0,
                        Phase = b.ProjectPhase ?? DOEProjectPhase.Screening,
                        PhaseText = (b.ProjectPhase ?? DOEProjectPhase.Screening) switch
                        {
                            DOEProjectPhase.Screening => "筛选",
                            DOEProjectPhase.PathSearch => "爬坡",
                            DOEProjectPhase.RSM => "RSM",
                            DOEProjectPhase.Augmenting => "补点",
                            DOEProjectPhase.Confirmation => "验证",
                            _ => b.ProjectPhase?.ToString() ?? ""
                        },
                        DesignMethod = b.DesignMethod,
                        DesignMethodText = b.DesignMethod switch
                        {
                            DOEDesignMethod.CCD => "CCD",
                            DOEDesignMethod.BoxBehnken => "BBD",
                            DOEDesignMethod.DOptimal => "D-Optimal",
                            DOEDesignMethod.FullFactorial => "全因子",
                            DOEDesignMethod.FractionalFactorial => "部分因子",
                            DOEDesignMethod.Taguchi => "Taguchi",
                            DOEDesignMethod.SteepestAscent => "最速上升",
                            DOEDesignMethod.AugmentedDesign => "增强设计",
                            DOEDesignMethod.ConfirmationRuns => "验证实验",
                            DOEDesignMethod.PlackettBurman => "PB",
                            _ => b.DesignMethod.ToString()
                        },
                        Status = b.Status,
                        TotalRuns = totalRuns,
                        CompletedRuns = completedRuns,
                        ProgressText = $"{completedRuns}/{totalRuns}",
                        RSquared = summary?.RSquared,
                        RSquaredText = summary?.RSquared.HasValue == true ? $"R²={summary.RSquared:F3}" : "",
                        Recommendation = summary?.RecommendationReason ?? "",
                        CreatedTime = b.CreatedTime
                    });
                }

                Rounds = new ObservableCollection<RoundListItem>(roundItems);
                StatusMessage = $"项目 \"{SelectedProject.ProjectName}\" — {roundItems.Count} 轮";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载项目轮次失败");
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task OnSelectedRoundChangedAsync()
        {
            // 清理旧选中状态
            foreach (var r in Rounds)
                r.IsSelected = false;

            RoundSummary = null;
            RaisePropertyChanged(nameof(HasRoundSummary));

            if (_selectedRound == null)
            {
                RoundDetail = null;
                HasSelectedRound = false;
                SelectedRoundRuns = new DataTable();
                return;
            }

            // 标记选中
            _selectedRound.IsSelected = true;
            HasSelectedRound = true;

            try
            {
                // 加载批次详情
                RoundDetail = await _repository.GetBatchWithDetailsAsync(_selectedRound.BatchId);
                RoundSummary = await _repository.GetRoundSummaryByBatchAsync(_selectedRound.BatchId);
                RaisePropertyChanged(nameof(HasRoundSummary));

                // 更新概况文字
                SelectedRoundTitle = $"R{_selectedRound.RoundNumber} · {_selectedRound.DesignMethodText} · {_selectedRound.PhaseText}";
                SelectedRoundR2Text = _selectedRound.RSquaredText ?? "—";
                SelectedRoundRecommendation = RoundSummary?.RecommendationReason ?? "";
                HasSelectedRoundRecommendation = !string.IsNullOrEmpty(SelectedRoundRecommendation);

                // ★ 加载实验数据表
                await LoadRunDataTableAsync(_selectedRound.BatchId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载轮次详情失败");
            }
        }
        private async Task LoadRunDataTableAsync(string batchId)
        {
            try
            {
                var batch = RoundDetail ?? await _repository.GetBatchWithDetailsAsync(batchId);
                if (batch == null) { SelectedRoundRuns = new DataTable(); return; }

                var runs = await _repository.GetRunsAsync(batchId);
                SelectedRoundRunCount = runs.Count;

                // 获取因子名和响应名
                var factors = batch.Factors ?? await _repository.GetFactorsAsync(batchId);
                var responses = batch.Responses ?? await _repository.GetResponsesAsync(batchId);

                var factorNames = factors.OrderBy(f => f.SortOrder).Select(f => f.FactorName).ToList();
                var responseNames = responses.OrderBy(r => r.SortOrder).Select(r => r.ResponseName).ToList();

                // 构建 DataTable
                var dt = new DataTable();
                dt.Columns.Add("#", typeof(int));
                foreach (var fn in factorNames)
                    dt.Columns.Add(fn, typeof(string));
                foreach (var rn in responseNames)
                    dt.Columns.Add(rn, typeof(string));
                dt.Columns.Add("状态", typeof(string));

                // 找最优值（用于标记最优行）
                double? bestValue = null;
                int bestIndex = -1;
                if (responseNames.Count > 0)
                {
                    foreach (var run in runs.Where(r => r.Status == DOERunStatus.Completed))
                    {
                        try
                        {
                            var respValues = JsonConvert.DeserializeObject<Dictionary<string, double>>(run.ResponseValuesJson ?? "{}");
                            if (respValues != null && respValues.TryGetValue(responseNames[0], out var val))
                            {
                                if (!bestValue.HasValue || val > bestValue.Value)
                                {
                                    bestValue = val;
                                    bestIndex = run.RunIndex;
                                }
                            }
                        }
                        catch { }
                    }
                }

                // 填充行
                foreach (var run in runs.OrderBy(r => r.RunIndex))
                {
                    var row = dt.NewRow();
                    row["#"] = run.RunIndex + 1;

                    // 因子值
                    try
                    {
                        var factorValues = JsonConvert.DeserializeObject<Dictionary<string, object>>(run.FactorValuesJson ?? "{}");
                        if (factorValues != null)
                        {
                            foreach (var fn in factorNames)
                            {
                                if (factorValues.TryGetValue(fn, out var val))
                                    row[fn] = val?.ToString() ?? "";
                            }
                        }
                    }
                    catch { }

                    // 响应值
                    try
                    {
                        var respValues = JsonConvert.DeserializeObject<Dictionary<string, double>>(run.ResponseValuesJson ?? "{}");
                        if (respValues != null)
                        {
                            foreach (var rn in responseNames)
                            {
                                if (respValues.TryGetValue(rn, out var val))
                                    row[rn] = val.ToString("F2");
                            }
                        }
                    }
                    catch { }

                    // 状态
                    row["状态"] = run.Status switch
                    {
                        DOERunStatus.Completed => "✓",
                        DOERunStatus.Failed => "✗",
                        DOERunStatus.Pending => "待执行",
                        DOERunStatus.Running => "执行中",
                        _ => run.Status.ToString()
                    };

                    dt.Rows.Add(row);
                }

                SelectedRoundRuns = dt;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载实验数据表失败: {BatchId}", batchId);
                SelectedRoundRuns = new DataTable();
            }
        }
        // ══════════════ 操作 ══════════════

        private async Task DeleteProjectAsync(DOEProjectSummary? project)
        {
            if (project == null) return;
            if (!_dialogService.ShowConfirmation(
                $"确定删除项目 \"{project.ProjectName}\"？\n" +
                $"该项目包含 {project.TotalBatches} 轮、{project.TotalExperiments} 组实验数据。\n" +
                $"项目下的批次将被解除关联（不删除实验数据）。\n\n此操作不可恢复。",
                "删除项目"))
                return;

            try
            {
                await _repository.DeleteProjectWithChildrenAsync(project.ProjectId);
                Projects.Remove(project);
                StatusMessage = $"已删除项目: {project.ProjectName}";

                if (SelectedProject?.ProjectId == project.ProjectId)
                {
                    SelectedProject = null;
                    Rounds.Clear();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "删除项目失败");
                _dialogService.ShowError($"删除失败: {ex.Message}", "错误");
            }
        }

        private async Task DeleteRoundAsync(RoundListItem? round)
        {
            if (round == null || SelectedProject == null) return;
            if (!_dialogService.ShowConfirmation(
                $"确定删除第 {round.RoundNumber} 轮 \"{round.BatchName}\"？\n此操作不可恢复。",
                "删除轮次"))
                return;

            try
            {
                await _repository.DeleteBatchWithChildrenAsync(round.BatchId);
                Rounds.Remove(round);
                StatusMessage = $"已删除轮次: {round.BatchName}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "删除轮次失败");
                _dialogService.ShowError($"删除失败: {ex.Message}", "错误");
            }
        }

        private async Task ExportRoundAsync(RoundListItem? round)
        {
            if (round == null) return;

            var path = _dialogService.ShowSaveFileDialog(
                $"DOE报告_{round.BatchName}_{DateTime.Now:yyyyMMdd}.xlsx",
                "Excel 文件|*.xlsx");
            if (string.IsNullOrEmpty(path)) return;

            try
            {
                IsLoading = true;
                StatusMessage = "正在导出报告...";
                await _exportService.ExportBatchReportAsync(round.BatchId, path);
                StatusMessage = $"报告已导出: {path}";
                _dialogService.ShowInfo($"DOE 报告已导出至:\n{path}", "导出成功");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "导出报告失败");
                _dialogService.ShowError($"导出失败: {ex.Message}", "错误");
            }
            finally
            {
                IsLoading = false;
            }
        }
    }

    /// <summary>
    /// 轮次列表项（历史页右侧列表绑定）
    /// </summary>
    public class RoundListItem
    {
        public string BatchId { get; set; } = "";
        public string BatchName { get; set; } = "";
        public int RoundNumber { get; set; }
        public DOEProjectPhase Phase { get; set; }
        public string PhaseText { get; set; } = "";
        public DOEDesignMethod DesignMethod { get; set; }
        public string DesignMethodText { get; set; } = "";
        public DOEBatchStatus Status { get; set; }
        public int TotalRuns { get; set; }
        public int CompletedRuns { get; set; }
        public string ProgressText { get; set; } = "";
        public double? RSquared { get; set; }
        public string RSquaredText { get; set; } = "";
        public string Recommendation { get; set; } = "";
        public DateTime CreatedTime { get; set; }

        /// <summary>状态显示文本</summary>
        public string StatusText => Status switch
        {
            DOEBatchStatus.Ready => "就绪",
            DOEBatchStatus.Running => "执行中",
            DOEBatchStatus.Paused => "已暂停",
            DOEBatchStatus.Completed => "已完成",
            DOEBatchStatus.Aborted => "已中止",
            _ => Status.ToString()
        };

        private bool _isSelected;
        public bool IsSelected { get => _isSelected; set { _isSelected = value; } }
        public string PhaseColor => Phase switch
        {
            DOEProjectPhase.Screening => "#E67E22",
            DOEProjectPhase.PathSearch => "#3498DB",
            DOEProjectPhase.RSM => "#9B59B6",
            DOEProjectPhase.Augmenting => "#1ABC9C",
            DOEProjectPhase.Confirmation => "#27AE60",
            _ => "#7F8C8D"
        };
    }
}