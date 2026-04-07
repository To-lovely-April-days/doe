using System;
using System.Linq;
using System.Threading.Tasks;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Events;
using MaxChemical.Modules.DOE.Models;
using MaxChemical.Modules.DOE.Services;
using MaxChemical.Modules.DOE.Views;
using Prism.Commands;
using Prism.Events;
using Prism.Ioc;
using Prism.Mvvm;

namespace MaxChemical.Modules.DOE.ViewModels
{
    /// <summary>
    /// DOE 主 ViewModel — 项目模式
    /// 
    /// ★ 重写: "新建方案"改为"新建项目"，所有操作在项目上下文中进行
    /// </summary>
    public class DOEMainViewModel : BindableBase
    {
        private readonly IContainerProvider _container;
        private readonly IEventAggregator _eventAggregator;
        private readonly IDOERepository _repository;
        private readonly IProjectDecisionEngine _decisionEngine;
        private readonly ILogService _logger;

        private int _selectedTabIndex;
        private string _currentBatchId = "";
        private string? _currentProjectId;
        private DOEProject? _currentProject;

        public DOEMainViewModel(
            IContainerProvider container,
            IEventAggregator eventAggregator,
            IDOERepository repository,
            IProjectDecisionEngine decisionEngine,
            ILogService logger)
        {
            _container = container;
            _eventAggregator = eventAggregator;
            _repository = repository;
            _decisionEngine = decisionEngine;
            _logger = logger?.ForContext<DOEMainViewModel>() ?? throw new ArgumentNullException(nameof(logger));

            CreateProjectCommand = new DelegateCommand(async () => await CreateProjectAsync());
            ContinueProjectCommand = new DelegateCommand<string>(async id => await ContinueProjectAsync(id));
            NavigateToExecutionCommand = new DelegateCommand<string>(NavigateToExecution);
            NavigateToAnalysisCommand = new DelegateCommand<string>(s => NavigateToAnalysis(s));

            // ★ 修复: 保存订阅 Token，窗口关闭时取消订阅防止重复弹窗
            _roundCompletedToken = _eventAggregator.GetEvent<DOERoundCompletedEvent>().Subscribe(
                async payload => await OnRoundCompletedAsync(payload),
                ThreadOption.UIThread);
        }

        private readonly SubscriptionToken _roundCompletedToken;

        /// <summary>★ 窗口关闭时调用，取消事件订阅</summary>
        public void Dispose()
        {
            _eventAggregator.GetEvent<DOERoundCompletedEvent>().Unsubscribe(_roundCompletedToken);
        }

        // ══════════════ Tab 导航 ══════════════

        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set
            {
                if (SetProperty(ref _selectedTabIndex, value))
                {
                    RaisePropertyChanged(nameof(IsTab0));
                    RaisePropertyChanged(nameof(IsTab1));
                    RaisePropertyChanged(nameof(IsTab2));
                    RaisePropertyChanged(nameof(IsTab3));
                }
            }
        }

        public bool IsTab0 => SelectedTabIndex == 0;
        public bool IsTab1 => SelectedTabIndex == 1;
        public bool IsTab2 => SelectedTabIndex == 2;
        public bool IsTab3 => SelectedTabIndex == 3;
        public string CurrentBatchId { get => _currentBatchId; set => SetProperty(ref _currentBatchId, value); }

        // ══════════════ 项目状态 ══════════════

        private string _currentProjectName = "";
        public string CurrentProjectName { get => _currentProjectName; set => SetProperty(ref _currentProjectName, value); }

        private bool _hasActiveProject;
        public bool HasActiveProject { get => _hasActiveProject; set => SetProperty(ref _hasActiveProject, value); }

        private string _projectPhaseText = "";
        public string ProjectPhaseText { get => _projectPhaseText; set => SetProperty(ref _projectPhaseText, value); }

        private string _projectStatusText = "";
        public string ProjectStatusText { get => _projectStatusText; set => SetProperty(ref _projectStatusText, value); }

        // ══════════════ Commands ══════════════

        public DelegateCommand CreateProjectCommand { get; }
        public DelegateCommand<string> ContinueProjectCommand { get; }
        public DelegateCommand<string> NavigateToExecutionCommand { get; }
        public DelegateCommand<string> NavigateToAnalysisCommand { get; }

        // ══════════════ Events ══════════════

        public event EventHandler<string>? RequestLoadExecution;
        public event EventHandler<string>? RequestLoadAnalysis;
        public event EventHandler? RequestRefreshHistory;
        public event EventHandler? RequestRefreshOverview;

        // ══════════════ 创建项目 ══════════════

        private async Task CreateProjectAsync()
        {
            try
            {
                var projectName = Microsoft.VisualBasic.Interaction.InputBox(
                    "请输入项目名称（如：XX产品产率优化）",
                    "创建优化项目", $"项目_{DateTime.Now:yyyyMMdd}");
                if (string.IsNullOrWhiteSpace(projectName)) return;

                var objective = Microsoft.VisualBasic.Interaction.InputBox(
                    "请输入优化目标描述（可选）\n如：最大化产率，约束纯度≥99%",
                    "优化目标", "");

                var project = new DOEProject
                {
                    ProjectName = projectName.Trim(),
                    ObjectiveDescription = objective?.Trim() ?? "",
                    CurrentPhase = DOEProjectPhase.Screening,
                    Status = DOEProjectStatus.Active
                };

                var projectId = await _repository.CreateProjectAsync(project);
                _currentProjectId = projectId;
                _currentProject = await _repository.GetProjectAsync(projectId);
                UpdateProjectDisplay();
                _logger.LogInformation("项目已创建: {ProjectId} - {Name}", projectId, projectName);

                OpenDesignWizardForProject(projectId, 1, DOEProjectPhase.Screening);
                RequestRefreshOverview?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex) { _logger.LogError(ex, "创建项目失败"); }
        }

        // ══════════════ 继续项目 ══════════════

        public async Task ContinueProjectAsync(string? projectId)
        {
            if (string.IsNullOrEmpty(projectId)) return;
            try
            {
                _currentProjectId = projectId;
                _currentProject = await _repository.GetProjectAsync(projectId);
                if (_currentProject == null) return;
                UpdateProjectDisplay();
                OpenDesignWizardForProject(projectId, _currentProject.CompletedRounds + 1, _currentProject.CurrentPhase);
            }
            catch (Exception ex) { _logger.LogError(ex, "继续项目失败"); }
        }

        // ══════════════ 轮次完成 → 决策 ══════════════

        private string? _lastAnalyzedBatchId;  // ★ 防重入: 记录已分析的批次ID

        private async Task OnRoundCompletedAsync(DOERoundCompletedPayload payload)
        {
            // ★ 防重入: 同一个批次只弹一次
            if (_lastAnalyzedBatchId == payload.BatchId)
            {
                _logger.LogWarning("轮次分析已处理过，跳过重复触发: Batch={BatchId}", payload.BatchId);
                return;
            }
            _lastAnalyzedBatchId = payload.BatchId;

            try
            {
                var summary = await _decisionEngine.AnalyzeRoundAsync(payload.ProjectId, payload.BatchId);
                _currentProjectId = payload.ProjectId;
                _currentProject = await _repository.GetProjectAsync(payload.ProjectId);
                UpdateProjectDisplay();

                var r2Text = summary.RSquared.HasValue ? $"R² = {summary.RSquared:F4}" : "模型未拟合";
                var bestText = summary.BestResponseValue.HasValue ? $"当前最优 = {summary.BestResponseValue:F2}" : "";

                var message = $"第 {summary.RoundNumber} 轮分析完成\n\n" +
                              $"{r2Text}\n{bestText}\n\n" +
                              $"推荐: {summary.RecommendationReason}\n\n" +
                              (summary.Recommendation == NextStepRecommendation.Complete
                                  ? "项目已达到优化目标，是否标记为完成？"
                                  : "是否按推荐创建下一轮实验？");

                var result = System.Windows.MessageBox.Show(message, "轮次分析",
                    System.Windows.MessageBoxButton.YesNoCancel, System.Windows.MessageBoxImage.Question);

                if (result == System.Windows.MessageBoxResult.Yes)
                {
                    if (summary.Recommendation == NextStepRecommendation.Complete)
                    {
                        await _decisionEngine.AdvancePhaseAsync(payload.ProjectId, DOEProjectPhase.Completed);
                        await _repository.UpdateProjectStatusAsync(payload.ProjectId, DOEProjectStatus.Completed);
                        _currentProject = await _repository.GetProjectAsync(payload.ProjectId);
                        UpdateProjectDisplay();
                        System.Windows.MessageBox.Show("项目已标记为完成！", "优化完成",
                            System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                    }
                    else
                    {
                        var nextPhase = RecommendationToPhase(summary.Recommendation);
                        await _decisionEngine.AdvancePhaseAsync(payload.ProjectId, nextPhase);
                        OpenDesignWizardForProject(payload.ProjectId, summary.RoundNumber + 1, nextPhase);
                    }
                }
                else if (result == System.Windows.MessageBoxResult.No)
                {
                    OpenDesignWizardForProject(payload.ProjectId, summary.RoundNumber + 1,
                        _currentProject?.CurrentPhase ?? DOEProjectPhase.RSM);
                }

                RequestRefreshOverview?.Invoke(this, EventArgs.Empty);
                RequestRefreshHistory?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex) { _logger.LogError(ex, "轮次分析失败"); }
        }

        // ══════════════ 导航 ══════════════

        public void NavigateToExecution(string? batchId)
        {
            if (string.IsNullOrEmpty(batchId)) return;
            CurrentBatchId = batchId;
            SelectedTabIndex = 1;
            RequestLoadExecution?.Invoke(this, batchId);
        }

        private void NavigateToAnalysis(string? batchId)
        {
            if (string.IsNullOrEmpty(batchId)) return;
            CurrentBatchId = batchId;
            SelectedTabIndex = 2;
            RequestLoadAnalysis?.Invoke(this, batchId);
        }

        // ══════════════ 内部方法 ══════════════

        private void OpenDesignWizardForProject(string projectId, int roundNumber, DOEProjectPhase phase)
        {
            try
            {
                var wizardView = _container.Resolve<DOEDesignWizardView>();
                string? savedBatchId = null;

                if (wizardView.DataContext is DOEDesignWizardViewModel wizardVm)
                {
                    var factors = Task.Run(async () =>
                        await _decisionEngine.GetNextRoundFactorsAsync(projectId)).Result;

                    var method = _decisionEngine.RecommendDesignMethod(
                        NextStepRecommendation.UserDecision, factors.Count, phase);

                    // ★ 优化2: 查询上一轮的响应变量
                    var previousResponses = new System.Collections.Generic.List<DOEResponse>();
                    try
                    {
                        var batches = Task.Run(async () =>
                            await _repository.GetBatchesByProjectAsync(projectId)).Result;
                        var lastBatch = batches.OrderByDescending(b => b.RoundNumber ?? 0).FirstOrDefault();
                        if (lastBatch != null)
                        {
                            previousResponses = Task.Run(async () =>
                                await _repository.GetResponsesAsync(lastBatch.BatchId)).Result;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "获取上一轮响应变量失败，用户需手动添加");
                    }

                    wizardVm.InitializeFromProject(new NextRoundPayload
                    {
                        ProjectId = projectId,
                        RoundNumber = roundNumber,
                        RecommendedPhase = phase,
                        RecommendedMethod = method,
                        PrefilledFactors = factors,
                        PrefilledResponses = previousResponses
                    });
                    // ★ 新增：设置项目名称供顶部栏显示
                    wizardVm.ProjectName = _currentProject?.ProjectName ?? "";
                    wizardVm.BatchSaved += (s, id) => savedBatchId = id;
                }

                wizardView.ShowDialog();

                if (!string.IsNullOrEmpty(savedBatchId))
                    NavigateToExecution(savedBatchId);

                RequestRefreshOverview?.Invoke(this, EventArgs.Empty);
                RequestRefreshHistory?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex) { _logger.LogError(ex, "打开设计向导失败"); }
        }

        private void UpdateProjectDisplay()
        {
            if (_currentProject != null)
            {
                HasActiveProject = true;
                CurrentProjectName = _currentProject.ProjectName;
                ProjectPhaseText = _currentProject.CurrentPhase switch
                {
                    DOEProjectPhase.Screening => "筛选阶段",
                    DOEProjectPhase.PathSearch => "路径探索",
                    DOEProjectPhase.RSM => "响应面优化",
                    DOEProjectPhase.Augmenting => "增强补点",
                    DOEProjectPhase.Confirmation => "验证阶段",
                    DOEProjectPhase.Completed => "已完成",
                    _ => _currentProject.CurrentPhase.ToString()
                };
                ProjectStatusText = _currentProject.BestResponseValue.HasValue
                    ? $"最优: {_currentProject.BestResponseValue:F2} | {_currentProject.TotalExperiments} 组"
                    : $"{_currentProject.TotalExperiments} 组实验";
            }
            else
            {
                HasActiveProject = false;
                CurrentProjectName = "";
                ProjectPhaseText = "";
                ProjectStatusText = "";
            }
        }
        public async Task NavigateToAnalysisByProjectAsync(string projectId)
        {
            try
            {
                var batches = await _repository.GetBatchesByProjectAsync(projectId);
                var latest = batches?.OrderByDescending(b => b.RoundNumber ?? 0).FirstOrDefault();
                if (latest != null)
                {
                    CurrentBatchId = latest.BatchId;
                    SelectedTabIndex = 2;
                    RequestLoadAnalysis?.Invoke(this, latest.BatchId);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "跳转项目分析页失败: {ProjectId}", projectId);
            }
        }
        private static DOEProjectPhase RecommendationToPhase(NextStepRecommendation rec) => rec switch
        {
            NextStepRecommendation.ContinueScreening => DOEProjectPhase.Screening,
            NextStepRecommendation.SteepestAscent => DOEProjectPhase.PathSearch,
            NextStepRecommendation.StartRSM => DOEProjectPhase.RSM,
            NextStepRecommendation.AugmentDesign => DOEProjectPhase.Augmenting,
            NextStepRecommendation.ExpandRange => DOEProjectPhase.RSM,
            NextStepRecommendation.ConfirmationRuns => DOEProjectPhase.Confirmation,
            NextStepRecommendation.Complete => DOEProjectPhase.Completed,
            _ => DOEProjectPhase.RSM
        };
    }
}