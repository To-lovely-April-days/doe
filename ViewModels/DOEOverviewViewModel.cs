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
    public class DOEOverviewViewModel : BindableBase
    {
        private readonly IDOERepository _repository;
        private readonly IFlowParameterProvider _paramProvider;
        private readonly ILogService _logger;

        private int _totalBatchCount;
        private int _totalCompletedRuns;
        private string _gprStatusSummary = "未初始化";
        private string _bestResult = "-";
        private bool _hasResumableBatch;
        private string _resumableBatchSummary = "";
        private string _resumableBatchId = "";
        private DOEBatchSummary? _selectedBatch;
        private ObservableCollection<DOEBatchSummary> _recentBatches = new();

        public DOEOverviewViewModel(IDOERepository repository, IFlowParameterProvider paramProvider, ILogService logger)
        {
            _repository = repository;
            _paramProvider = paramProvider;
            _logger = logger?.ForContext<DOEOverviewViewModel>() ?? throw new ArgumentNullException(nameof(logger));

            ResumeLastBatchCommand = new DelegateCommand(() => RequestResumeBatch?.Invoke(this, _resumableBatchId),
                () => HasResumableBatch);
            GoToHistoryCommand = new DelegateCommand(() => RequestGoToHistory?.Invoke(this, EventArgs.Empty));
            EditBatchCommand = new DelegateCommand(EditSelectedBatch);  //  新增
        }

        public int TotalBatchCount { get => _totalBatchCount; set => SetProperty(ref _totalBatchCount, value); }
        public int TotalCompletedRuns { get => _totalCompletedRuns; set => SetProperty(ref _totalCompletedRuns, value); }
        public string GPRStatusSummary { get => _gprStatusSummary; set => SetProperty(ref _gprStatusSummary, value); }
        public string BestResult { get => _bestResult; set => SetProperty(ref _bestResult, value); }
        public bool HasResumableBatch { get => _hasResumableBatch; set { SetProperty(ref _hasResumableBatch, value); ResumeLastBatchCommand.RaiseCanExecuteChanged(); } }
        public string ResumableBatchSummary { get => _resumableBatchSummary; set => SetProperty(ref _resumableBatchSummary, value); }
        public ObservableCollection<DOEBatchSummary> RecentBatches { get => _recentBatches; set => SetProperty(ref _recentBatches, value); }
        public DOEBatchSummary? SelectedBatch { get => _selectedBatch; set => SetProperty(ref _selectedBatch, value); }  //  新增

        public DelegateCommand ResumeLastBatchCommand { get; }
        public DelegateCommand GoToHistoryCommand { get; }
        public DelegateCommand EditBatchCommand { get; }  //  新增

        public event EventHandler<string>? RequestResumeBatch;
        public event EventHandler? RequestGoToHistory;
        public event EventHandler<string>? RequestEditBatch;  //  新增

        //  编辑方案（仅 Ready 状态）
        private void EditSelectedBatch()
        {
            if (SelectedBatch == null) return;
            if (SelectedBatch.Status != DOEBatchStatus.Ready)
            {
                // 静默忽略，非 Ready 状态不可编辑
                return;
            }
            RequestEditBatch?.Invoke(this, SelectedBatch.BatchId);
        }

        public async Task LoadAsync()
        {
            try
            {
                var batches = await _repository.GetRecentBatchesAsync(20);
                TotalBatchCount = batches.Count;

                var summaries = batches.Select(b => new DOEBatchSummary
                {
                    BatchId = b.BatchId,
                    BatchName = b.BatchName,
                    FlowName = b.FlowName,
                    DesignMethod = b.DesignMethod,
                    Status = b.Status,
                    CreatedTime = b.CreatedTime,
                    TotalRuns = b.Runs.Count,
                    CompletedRuns = b.Runs.Count(r => r.Status == DOERunStatus.Completed)
                }).ToList();

                TotalCompletedRuns = summaries.Sum(s => s.CompletedRuns);
                RecentBatches = new ObservableCollection<DOEBatchSummary>(summaries.Take(5));

                // 查找可恢复批次
                var resumable = batches.FirstOrDefault(b => b.Status == DOEBatchStatus.Running || b.Status == DOEBatchStatus.Paused);
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

                // GPR 状态
                var flowId = _paramProvider.GetCurrentFlowId();
                if (!string.IsNullOrEmpty(flowId))
                {
                    var state = await _repository.GetGPRModelStateAsync(flowId);
                    GPRStatusSummary = state?.IsActive == true ? $"已激活 R²={state.RSquared:F2}" : "未激活";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载概览数据失败");
            }
        }
    }
}