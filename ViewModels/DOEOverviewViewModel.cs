using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using MaxChemical.Infrastructure.DOE;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Models;
using Prism.Commands;
using Prism.Mvvm;

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

        // ══════════════ Events ══════════════

        public event EventHandler<string>? RequestResumeBatch;
        public event EventHandler? RequestGoToHistory;
        /// <summary>请求继续某个项目的优化（触发下一轮）</summary>
        public event EventHandler<string>? RequestContinueProject;
        /// <summary>请求查看某个项目的详情（跳转历史页）</summary>
        public event EventHandler<string>? RequestViewProject;
        /// <summary>请求创建新项目</summary>
        public event EventHandler? RequestCreateProject;

        // ══════════════ 加载 ══════════════

        public async Task LoadAsync()
        {
            try
            {
                // 加载项目列表
                var projects = await _repository.GetAllProjectSummariesAsync(20);

                TotalProjectCount = projects.Count;
                ActiveProjectCount = projects.Count(p => p.Status == DOEProjectStatus.Active);
                TotalExperiments = projects.Sum(p => p.TotalExperiments);

                var cards = projects.Select(p => new ProjectCardItem
                {
                    ProjectId = p.ProjectId,
                    ProjectName = p.ProjectName,
                    FlowName = p.FlowName ?? "",
                    CurrentPhase = p.CurrentPhase,
                    PhaseText = p.PhaseText,
                    Status = p.Status,
                    TotalBatches = p.TotalBatches,
                    TotalExperiments = p.TotalExperiments,
                    BestResponseValue = p.BestResponseValue,
                    CreatedTime = p.CreatedTime,
                    UpdatedTime = p.UpdatedTime,
                    // 状态颜色
                    PhaseColor = p.CurrentPhase switch
                    {
                        DOEProjectPhase.Screening => "#E67E22",
                        DOEProjectPhase.PathSearch => "#3498DB",
                        DOEProjectPhase.RSM => "#9B59B6",
                        DOEProjectPhase.Augmenting => "#1ABC9C",
                        DOEProjectPhase.Confirmation => "#27AE60",
                        DOEProjectPhase.Completed => "#95A5A6",
                        _ => "#7F8C8D"
                    },
                    StatusIcon = p.Status == DOEProjectStatus.Active ? "▶" : "■",
                    BestValueText = p.BestResponseValue.HasValue ? $"{p.BestResponseValue:F2}" : "—",
                    SummaryText = $"{p.TotalBatches} 轮 · {p.TotalExperiments} 组实验"
                }).ToList();

                ProjectCards = new ObservableCollection<ProjectCardItem>(cards);
                RaisePropertyChanged(nameof(HasProjects));

                // 查找正在执行的批次（跨所有项目）
                await FindResumableBatchAsync();

                // GPR 状态
                var flowId = _paramProvider.GetCurrentFlowId();
                if (!string.IsNullOrEmpty(flowId))
                {
                    var state = await _repository.GetGPRModelStateAsync(flowId);
                    GPRStatusSummary = state?.IsActive == true ? $"已激活 R²={state.RSquared:F2}" : "未初始化";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载概览数据失败");
            }
        }

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
    }
}