using MaxChemical.Infrastructure.DOE;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Models;
using MaxChemical.Modules.DOE.Views;
using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace MaxChemical.Modules.DOE.ViewModels
{
    /// <summary>
    /// 概览页 ViewModel — 项目模式
    /// 
    /// ★ 重写: 从"最近批次列表"改为"项目卡片列表"
    /// 
    /// 展示内容:
    ///   - 活跃项目卡片（项目名、阶段、最优值、总实验数、轮次数）
    ///   - 统计摘要（总项目数、总实验数、GPR 状态）
    ///   - 快捷操作: 新建项目、继续优化、查看历史
    /// </summary>
    public class DOEOverviewViewModel : BindableBase
    {
        private readonly IDOERepository _repository;
        private readonly IFlowParameterProvider _paramProvider;
        private readonly ILogService _logger;

        // ── 统计数据 ──
        private int _totalProjectCount;
        private int _activeProjectCount;
        private int _totalExperiments;
        private string _gprStatusSummary = "未初始化";

        // ── 项目列表 ──
        private ObservableCollection<ProjectCardItem> _projectCards = new();
        private ProjectCardItem? _selectedProject;

        // ── 可恢复批次（某个项目中正在执行的批次）──
        private bool _hasResumableBatch;
        private string _resumableBatchSummary = "";
        private string _resumableBatchId = "";

        public DOEOverviewViewModel(
            IDOERepository repository,
            IFlowParameterProvider paramProvider,
            ILogService logger)
        {
            _repository = repository;
            _paramProvider = paramProvider;
            _logger = logger?.ForContext<DOEOverviewViewModel>() ?? throw new ArgumentNullException(nameof(logger));

            ResumeLastBatchCommand = new DelegateCommand(
                () => RequestResumeBatch?.Invoke(this, _resumableBatchId),
                () => HasResumableBatch);
            GoToHistoryCommand = new DelegateCommand(
                () => RequestGoToHistory?.Invoke(this, EventArgs.Empty));
            ContinueProjectCommand = new DelegateCommand<string>(
                id => { if (!string.IsNullOrEmpty(id)) RequestContinueProject?.Invoke(this, id); });
            ViewProjectCommand = new DelegateCommand<string>(
                id => { if (!string.IsNullOrEmpty(id)) RequestViewProject?.Invoke(this, id); });
            ViewProjectDetailCommand = new DelegateCommand<string>(async id => await ShowProjectDetailAsync(id));
            ViewProjectAnalysisCommand = new DelegateCommand<string>(id => { if (!string.IsNullOrEmpty(id)) RequestViewAnalysis?.Invoke(this, id); });
            ArchiveProjectCommand = new DelegateCommand<string>(async id => await ArchiveProjectAsync(id));
            DeleteProjectCommand = new DelegateCommand<string>(async id => await DeleteProjectAsync(id));
        }

        // ══════════════ Properties ══════════════

        public int TotalProjectCount { get => _totalProjectCount; set => SetProperty(ref _totalProjectCount, value); }
        public int ActiveProjectCount { get => _activeProjectCount; set => SetProperty(ref _activeProjectCount, value); }
        public int TotalExperiments { get => _totalExperiments; set => SetProperty(ref _totalExperiments, value); }
        public string GPRStatusSummary { get => _gprStatusSummary; set => SetProperty(ref _gprStatusSummary, value); }

        public ObservableCollection<ProjectCardItem> ProjectCards
        {
            get => _projectCards;
            set => SetProperty(ref _projectCards, value);
        }

        public ProjectCardItem? SelectedProject
        {
            get => _selectedProject;
            set => SetProperty(ref _selectedProject, value);
        }

        public bool HasResumableBatch
        {
            get => _hasResumableBatch;
            set { SetProperty(ref _hasResumableBatch, value); ResumeLastBatchCommand.RaiseCanExecuteChanged(); }
        }
        public string ResumableBatchSummary { get => _resumableBatchSummary; set => SetProperty(ref _resumableBatchSummary, value); }

        /// <summary>是否有项目（控制空状态引导显示）</summary>
        public bool HasProjects => ProjectCards.Count > 0;

        // ══════════════ Commands ══════════════

        public DelegateCommand ResumeLastBatchCommand { get; }
        public DelegateCommand GoToHistoryCommand { get; }
        public DelegateCommand<string> ContinueProjectCommand { get; }
        public DelegateCommand<string> ViewProjectCommand { get; }
        public DelegateCommand<string> ViewProjectDetailCommand { get; }
        public DelegateCommand<string> ViewProjectAnalysisCommand { get; }
        public DelegateCommand<string> ArchiveProjectCommand { get; }
        public DelegateCommand<string> DeleteProjectCommand { get; }
        // ══════════════ Events ══════════════

        public event EventHandler<string>? RequestResumeBatch;
        public event EventHandler? RequestGoToHistory;
        /// <summary>请求继续某个项目的优化（触发下一轮）</summary>
        public event EventHandler<string>? RequestContinueProject;
        /// <summary>请求查看某个项目的详情（跳转历史页）</summary>
        public event EventHandler<string>? RequestViewProject;
        /// <summary>请求创建新项目</summary>
        public event EventHandler? RequestCreateProject;
        // 事件：
        public event EventHandler<string>? RequestViewAnalysis;
        // ══════════════ 加载 ══════════════

        public async Task LoadAsync()
        {
            try
            {
                var projects = await _repository.GetAllProjectSummariesAsync(20);

                TotalProjectCount = projects.Count;
                ActiveProjectCount = projects.Count(p => p.Status == DOEProjectStatus.Active);
                TotalExperiments = projects.Sum(p => p.TotalExperiments);

                var cards = new List<ProjectCardItem>();

                foreach (var p in projects)
                {
                    var card = new ProjectCardItem
                    {
                        ProjectId = p.ProjectId,
                        ProjectName = p.ProjectName,
                        FlowName = p.FlowName ?? "",
                        CurrentPhase = p.CurrentPhase,
                        PhaseText = GetShortPhaseText(p.CurrentPhase),
                        Status = p.Status,
                        TotalBatches = p.TotalBatches,
                        TotalExperiments = p.TotalExperiments,
                        BestResponseValue = p.BestResponseValue,
                        CreatedTime = p.CreatedTime,
                        UpdatedTime = p.UpdatedTime,
                        PhaseColor = GetPhaseColor(p.CurrentPhase),
                        StatusIcon = p.Status == DOEProjectStatus.Active ? "▶" : "■",
                        BestValueText = p.BestResponseValue.HasValue ? $"{p.BestResponseValue:F2}" : "—",
                        SummaryText = $"{p.TotalBatches} 轮 · {p.TotalExperiments} 组实验",
                        PhaseSegments = BuildPhaseSegments(p.CurrentPhase),
                        PhaseProgressText = BuildPhaseProgressText(p.CurrentPhase),
                    };

                    // ★ 查询最新批次的因子和响应，构建标签和详细摘要
                    await BuildCardDetails(card, p);

                    cards.Add(card);
                }

                ProjectCards = new ObservableCollection<ProjectCardItem>(cards);
                RaisePropertyChanged(nameof(HasProjects));

                await FindResumableBatchAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载概览数据失败");
            }
        }

        private async Task BuildCardDetails(ProjectCardItem card, DOEProjectSummary p)
        {
            try
            {
                // 查询该项目最新批次
                var batches = await _repository.GetBatchesByProjectAsync(p.ProjectId);
                var latestBatch = batches?.OrderByDescending(b => b.RoundNumber ?? 0).FirstOrDefault();

                if (latestBatch != null)
                {
                    // 设计方法文字
                    var methodText = GetDesignMethodText(latestBatch.DesignMethod);

                    // 详细摘要
                    card.DetailSummaryText = $"第 {latestBatch.RoundNumber ?? card.TotalBatches} 轮 · " +
                                             $"{methodText} · " +
                                             $"{p.TotalExperiments} 组实验 · " +
                                             $"{p.UpdatedTime:MM-dd HH:mm}";

                    // 因子标签
                    var factors = latestBatch.Factors;
                    if (factors == null || factors.Count == 0)
                    {
                        factors = (await _repository.GetFactorsAsync(latestBatch.BatchId))?.ToList()
                                  ?? new List<DOEFactor>();
                    }

                    var tags = new ObservableCollection<TagItem>();

                    // 因子标签（灰底）
                    foreach (var f in factors)
                    {
                        tags.Add(new TagItem
                        {
                            Text = f.FactorName,
                            Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#5A5F6D")),
                            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F2F4F7"))
                        });
                    }

                    // 响应标签（蓝底）
                    var responses = latestBatch.Responses;
                    if (responses == null || responses.Count == 0)
                    {
                        responses = (await _repository.GetResponsesAsync(latestBatch.BatchId))?.ToList()
                                    ?? new List<DOEResponse>();
                    }

                    foreach (var r in responses)
                    {
                        tags.Add(new TagItem
                        {
                            Text = r.ResponseName,
                            Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0078D4")),
                            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EBF3FE"))
                        });
                    }

                    card.TagItems = tags;
                }
                else
                {
                    card.DetailSummaryText = $"{p.TotalBatches} 轮 · {p.TotalExperiments} 组实验 · {p.UpdatedTime:MM-dd HH:mm}";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "构建项目卡片详情失败: {ProjectId}", p.ProjectId);
                card.DetailSummaryText = $"{p.TotalBatches} 轮 · {p.TotalExperiments} 组实验 · {p.UpdatedTime:MM-dd HH:mm}";
            }
        }

        private static string GetShortPhaseText(DOEProjectPhase phase) => phase switch
        {
            DOEProjectPhase.Screening => "筛选",
            DOEProjectPhase.PathSearch => "爬坡",
            DOEProjectPhase.RSM => "RSM",
            DOEProjectPhase.Augmenting => "补点",
            DOEProjectPhase.Confirmation => "验证",
            DOEProjectPhase.Completed => "完成",
            _ => phase.ToString()
        };

        private static string GetPhaseColor(DOEProjectPhase phase) => phase switch
        {
            DOEProjectPhase.Screening => "#E67E22",
            DOEProjectPhase.PathSearch => "#3498DB",
            DOEProjectPhase.RSM => "#9B59B6",
            DOEProjectPhase.Augmenting => "#1ABC9C",
            DOEProjectPhase.Confirmation => "#27AE60",
            DOEProjectPhase.Completed => "#95A5A6",
            _ => "#7F8C8D"
        };

        private static string GetDesignMethodText(DOEDesignMethod method) => method switch
        {
            DOEDesignMethod.FullFactorial => "全因子设计",
            DOEDesignMethod.FractionalFactorial => "部分因子设计",
            DOEDesignMethod.Taguchi => "Taguchi 设计",
            DOEDesignMethod.CCD => "CCD 设计",
            DOEDesignMethod.BoxBehnken => "Box-Behnken 设计",
            DOEDesignMethod.DOptimal => "D-Optimal 设计",
            DOEDesignMethod.SteepestAscent => "最速上升",
            DOEDesignMethod.AugmentedDesign => "增强设计",
            DOEDesignMethod.ConfirmationRuns => "验证实验",
            _ => method.ToString()
        };

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
            var current = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0078D4")) { Opacity = 0.45 };
            var inactive = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8EAED"));

            var segments = new ObservableCollection<PhaseSegmentItem>();
            for (int i = 0; i < 5; i++)
            {
                if (i < currentIndex)
                    segments.Add(new PhaseSegmentItem { Color = active });
                else if (i == currentIndex)
                    segments.Add(new PhaseSegmentItem { Color = current });
                else
                    segments.Add(new PhaseSegmentItem { Color = inactive });
            }
            return segments;
        }

        private static string BuildPhaseProgressText(DOEProjectPhase currentPhase) => currentPhase switch
        {
            DOEProjectPhase.Completed => "已完成全部优化流程",
            _ => "筛选 → 爬坡 → RSM → 验证 → 完成"
        };
        private async Task FindResumableBatchAsync()
        {
            try
            {
                var batches = await _repository.GetRecentBatchesAsync(50);
                var resumable = batches.FirstOrDefault(b =>
                    b.Status == DOEBatchStatus.Running || b.Status == DOEBatchStatus.Paused);

                if (resumable != null)
                {
                    HasResumableBatch = true;
                    _resumableBatchId = resumable.BatchId;
                    var done = resumable.Runs.Count(r => r.Status == DOERunStatus.Completed);
                    ResumableBatchSummary = $"{resumable.BatchName} — 已完成 {done}/{resumable.Runs.Count} 组";
                }
                else
                {
                    HasResumableBatch = false;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "查找可恢复批次失败");
                HasResumableBatch = false;
            }
        }
        private async Task ShowProjectDetailAsync(string? projectId)
        {
            if (string.IsNullOrEmpty(projectId)) return;
            try
            {
                var vm = new DOEProjectDetailViewModel(_repository);
                await vm.LoadAsync(projectId);
                var dialog = new DOEProjectDetailDialog
                {
                    DataContext = vm,
                    Owner = System.Windows.Application.Current.Windows
                        .OfType<Window>()
                        .FirstOrDefault(w => w.IsActive)
                        ?? System.Windows.Application.Current.MainWindow
                };
                dialog.ShowDialog();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "打开项目详情失败");
            }
        }

        private async Task ArchiveProjectAsync(string? projectId)
        {
            if (string.IsNullOrEmpty(projectId)) return;
            var result = System.Windows.MessageBox.Show("确定要归档此项目吗？归档后不会删除数据。",
                "归档确认", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question);
            if (result != System.Windows.MessageBoxResult.Yes) return;
            try
            {
                await _repository.UpdateProjectStatusAsync(projectId, DOEProjectStatus.Archived);
                await LoadAsync();
            }
            catch (Exception ex) { _logger.LogError(ex, "归档项目失败"); }
        }

        private async Task DeleteProjectAsync(string? projectId)
        {
            if (string.IsNullOrEmpty(projectId)) return;
            var result = System.Windows.MessageBox.Show("确定要删除此项目吗？\n批次数据将保留但不再关联到项目。",
                "删除确认", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Warning);
            if (result != System.Windows.MessageBoxResult.Yes) return;
            try
            {
                await _repository.DeleteProjectWithChildrenAsync(projectId);
                await LoadAsync();
            }
            catch (Exception ex) { _logger.LogError(ex, "删除项目失败"); }
        }
    }

    /// <summary>
    /// 项目卡片展示数据（供概览页 ItemsControl 绑定）
    /// </summary>
    public class ProjectCardItem : BindableBase
    {
        public string ProjectId { get; set; } = "";
        public string ProjectName { get; set; } = "";
        public string FlowName { get; set; } = "";
        public DOEProjectPhase CurrentPhase { get; set; }
        public string PhaseText { get; set; } = "";
        public DOEProjectStatus Status { get; set; }
        public int TotalBatches { get; set; }
        public int TotalExperiments { get; set; }
        public double? BestResponseValue { get; set; }
        public DateTime CreatedTime { get; set; }
        public DateTime UpdatedTime { get; set; }

        // 展示辅助
        public string PhaseColor { get; set; } = "#7F8C8D";
        public string StatusIcon { get; set; } = "▶";
        public string BestValueText { get; set; } = "—";
        public string SummaryText { get; set; } = "";

        /// <summary>是否可继续优化</summary>
        public bool CanContinue => Status == DOEProjectStatus.Active;
        /// <summary>阶段进度条分段数据（5段：筛选/爬坡/RSM/验证/完成）</summary>
        public ObservableCollection<PhaseSegmentItem> PhaseSegments { get; set; } = new();

        /// <summary>阶段进度文字（如 "筛选 → 爬坡 → RSM → 验证 → 完成"）</summary>
        public string PhaseProgressText { get; set; } = "";

        /// <summary>系统推荐文字（有推荐时显示）</summary>
        public string? RecommendationText { get; set; }

        /// <summary>是否有系统推荐</summary>
        public bool HasRecommendation => !string.IsNullOrEmpty(RecommendationText);
        /// <summary>详细摘要文字（如 "第 4 轮 · CCD 设计 · 72 组实验 · 04-06 13:53"）</summary>
        public string DetailSummaryText { get; set; } = "";

        /// <summary>因子+响应标签列表</summary>
        public ObservableCollection<TagItem> TagItems { get; set; } = new();



    }
    /// <summary>
    /// 阶段进度条分段项
    /// </summary>
    public class PhaseSegmentItem
    {
        /// <summary>分段颜色画刷</summary>
        public SolidColorBrush Color { get; set; } = new SolidColorBrush(Colors.LightGray);
    }
    /// <summary>
    /// 因子/响应标签项（供概览页 ItemsControl 绑定）
    /// </summary>
    public class TagItem
    {
        public string Text { get; set; } = "";
        public SolidColorBrush Foreground { get; set; } = new(Colors.Gray);
        public SolidColorBrush Background { get; set; } = new(Colors.LightGray);
    }
}