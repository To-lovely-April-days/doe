using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using MaxChemical.Infrastructure.DOE;
using MaxChemical.Infrastructure.Services;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Models;
using MaxChemical.Modules.DOE.Services;
using Prism.Commands;
using Prism.Mvvm;

namespace MaxChemical.Modules.DOE.ViewModels
{
    /// <summary>
    /// DOE 历史管理 ViewModel — 批次浏览、详情、复用方案、对比、GPR 模型管理、Excel 导出
    /// </summary>
    public class DOEHistoryViewModel : BindableBase
    {
        private readonly IDOERepository _repository;
        private readonly IGPRModelService _gprService;
        private readonly DOEExportService _exportService;
        private readonly IFlowParameterProvider _paramProvider;
        private readonly IDialogService _dialogService;
        private readonly ILogService _logger;

        private ObservableCollection<DOEBatchSummary> _batches = new();
        private DOEBatchSummary? _selectedBatch;
        private DOEBatch? _batchDetail;
        private bool _isLoading;
        private string _statusMessage = "";

        // GPR 模型状态
        private string _gprStatusText = "未初始化";
        private bool _gprIsActive;
        private int _gprDataCount;

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

            RefreshCommand = new DelegateCommand(async () => await LoadBatchesAsync());
            ViewDetailCommand = new DelegateCommand<DOEBatchSummary>(async s => await ViewBatchDetailAsync(s));
            DeleteBatchCommand = new DelegateCommand<DOEBatchSummary>(async s => await DeleteBatchAsync(s));
            ExportBatchCommand = new DelegateCommand<DOEBatchSummary>(async s => await ExportBatchAsync(s));
            ExportComparisonCommand = new DelegateCommand(async () => await ExportComparisonAsync(), () => Batches.Count >= 2);
            EditBatchCommand = new DelegateCommand<DOEBatchSummary>(s => EditBatch(s));  //  新增
            ResetGPRModelCommand = new DelegateCommand(async () => await ResetGPRModelAsync());
            RetrainGPRModelCommand = new DelegateCommand(async () => await RetrainGPRModelAsync());

            // 初始加载
            _ = LoadBatchesAsync();
        }

        // ══════════════ Properties ══════════════

        public ObservableCollection<DOEBatchSummary> Batches { get => _batches; set => SetProperty(ref _batches, value); }
        public DOEBatchSummary? SelectedBatch { get => _selectedBatch; set { SetProperty(ref _selectedBatch, value); _ = OnSelectedBatchChanged(); } }
        public DOEBatch? BatchDetail { get => _batchDetail; set => SetProperty(ref _batchDetail, value); }
        public bool IsLoading { get => _isLoading; set => SetProperty(ref _isLoading, value); }
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
        public string GPRStatusText { get => _gprStatusText; set => SetProperty(ref _gprStatusText, value); }
        public bool GPRIsActive { get => _gprIsActive; set => SetProperty(ref _gprIsActive, value); }
        public int GPRDataCount { get => _gprDataCount; set => SetProperty(ref _gprDataCount, value); }

        // ══════════════ Commands ══════════════

        public DelegateCommand RefreshCommand { get; }
        public DelegateCommand<DOEBatchSummary> ViewDetailCommand { get; }
        public DelegateCommand<DOEBatchSummary> DeleteBatchCommand { get; }
        public DelegateCommand<DOEBatchSummary> ExportBatchCommand { get; }
        public DelegateCommand ExportComparisonCommand { get; }
        public DelegateCommand<DOEBatchSummary> EditBatchCommand { get; }  //  新增
        public DelegateCommand ResetGPRModelCommand { get; }
        public DelegateCommand RetrainGPRModelCommand { get; }

        /// <summary>
        /// 请求打开设计向导（复用方案）事件
        /// </summary>
        public event EventHandler<string>? RequestReuseBatch;

        /// <summary>
        /// 请求打开执行仪表盘事件
        /// </summary>
        public event EventHandler<string>? RequestExecuteBatch;

        /// <summary>
        /// 请求打开分析视图事件
        /// </summary>
        public event EventHandler<string>? RequestAnalyzeBatch;

        /// <summary>
        ///  请求编辑方案事件（历史页右键菜单触发）
        /// </summary>
        public event EventHandler<string>? RequestEditBatch;

        // ══════════════ 加载 ══════════════

        public async Task LoadBatchesAsync()
        {
            try
            {
                IsLoading = true;

                //  直接查所有批次的摘要（不依赖 FlowId）
                var summaries = await _repository.GetAllBatchSummariesAsync(50);

                Batches = new ObservableCollection<DOEBatchSummary>(summaries);
                StatusMessage = $"共 {summaries.Count} 个 DOE 批次";
                ExportComparisonCommand.RaiseCanExecuteChanged();

                // 加载 GPR 模型状态
                var flowId = _paramProvider.GetCurrentFlowId();
                await LoadGPRStatusAsync(flowId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载批次列表失败");
                StatusMessage = $"加载失败: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task OnSelectedBatchChanged()
        {
            if (SelectedBatch == null) { BatchDetail = null; return; }
            await ViewBatchDetailAsync(SelectedBatch);
        }

        private async Task ViewBatchDetailAsync(DOEBatchSummary? summary)
        {
            if (summary == null) return;
            try
            {
                BatchDetail = await _repository.GetBatchWithDetailsAsync(summary.BatchId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载批次详情失败: {BatchId}", summary.BatchId);
            }
        }

        // ══════════════ 操作 ══════════════

        /// <summary>
        ///  编辑方案（仅 Ready 状态）
        /// </summary>
        private void EditBatch(DOEBatchSummary? summary)
        {
            if (summary == null) return;

            if (summary.Status != DOEBatchStatus.Ready)
            {
                _dialogService.ShowError("只有「就绪」状态的方案才能编辑。\n已执行或已完成的方案不可修改。", "无法编辑");
                return;
            }

            RequestEditBatch?.Invoke(this, summary.BatchId);
        }

        /// <summary>
        ///  删除方案（带 GPR 模型同步清理）
        /// </summary>
        private async Task DeleteBatchAsync(DOEBatchSummary? summary)
        {
            if (summary == null) return;
            if (!_dialogService.ShowConfirmation($"确定删除批次 '{summary.BatchName}'？\n此操作不可恢复。", "删除确认"))
                return;

            try
            {
                // 获取方案的因子信息（用于判断模型关联）
                var batch = await _repository.GetBatchWithDetailsAsync(summary.BatchId);
                string? flowId = batch?.FlowId;
                string? signature = null;

                if (batch?.Factors?.Count > 0)
                {
                    var factorNames = batch.Factors.Select(f => f.FactorName);
                    signature = GPRModelState.BuildSignature(factorNames);
                }

                // 删除批次及所有子表
                await _repository.DeleteBatchWithChildrenAsync(summary.BatchId);
                Batches.Remove(summary);
                StatusMessage = $"已删除批次: {summary.BatchName}";

                //  GPR 模型同步清理逻辑
                if (!string.IsNullOrEmpty(flowId) && !string.IsNullOrEmpty(signature))
                {
                    var remaining = await _repository.GetBatchCountBySignatureAsync(flowId, signature);

                    if (remaining == 0)
                    {
                        // 该签名下没有其他方案了，询问用户是否删除模型
                        var state = await _repository.GetGPRModelStateAsync(flowId, signature);
                        if (state != null)
                        {
                            var deleteModel = _dialogService.ShowConfirmation(
                                $"该方案关联的 GPR 模型（{state.ModelName}，{state.DataCount} 组数据）\n" +
                                $"不再被任何方案使用。是否一并删除？\n\n" +
                                $"选择「否」将保留模型数据以备后用。",
                                "清理关联模型");

                            if (deleteModel)
                            {
                                await _repository.DeleteGPRModelAsync(state.Id);
                                StatusMessage += "，已删除关联 GPR 模型";
                                _logger.LogInformation("已删除孤立 GPR 模型: {ModelName}, Sig={Sig}", state.ModelName, signature);
                            }
                        }
                    }
                }

                await LoadGPRStatusAsync(flowId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "删除批次失败");
                _dialogService.ShowError($"删除失败: {ex.Message}", "错误");
            }
        }

        private async Task ExportBatchAsync(DOEBatchSummary? summary)
        {
            if (summary == null) return;

            var path = _dialogService.ShowSaveFileDialog($"DOE报告_{summary.BatchName}_{DateTime.Now:yyyyMMdd}.xlsx",
                "Excel 文件|*.xlsx");
            if (string.IsNullOrEmpty(path)) return;

            try
            {
                IsLoading = true;
                StatusMessage = "正在导出报告...";
                await _exportService.ExportBatchReportAsync(summary.BatchId, path);
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

        private async Task ExportComparisonAsync()
        {
            var path = _dialogService.ShowSaveFileDialog($"DOE批次对比_{DateTime.Now:yyyyMMdd}.xlsx", "Excel 文件|*.xlsx");
            if (string.IsNullOrEmpty(path)) return;

            try
            {
                IsLoading = true;
                var batchIds = Batches.Select(b => b.BatchId).ToList();
                await _exportService.ExportComparisonReportAsync(batchIds, path);
                StatusMessage = $"对比报告已导出: {path}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "导出对比报告失败");
            }
            finally
            {
                IsLoading = false;
            }
        }

        // ══════════════ GPR 模型管理 ══════════════

        private async Task LoadGPRStatusAsync(string? flowId)
        {
            if (string.IsNullOrEmpty(flowId))
            {
                GPRStatusText = "未选择流程";
                return;
            }

            try
            {
                var state = await _repository.GetGPRModelStateAsync(flowId);
                if (state != null)
                {
                    GPRIsActive = state.IsActive;
                    GPRDataCount = state.DataCount;
                    GPRStatusText = state.IsActive
                        ? $"已激活 | 数据量: {state.DataCount} | R²: {state.RSquared:F4} | 上次训练: {state.LastTrainedTime:HH:mm}"
                        : $"未激活 | 数据量: {state.DataCount}/{6}（冷启动阈值）";
                }
                else
                {
                    GPRStatusText = "无模型记录";
                    GPRIsActive = false;
                    GPRDataCount = 0;
                }
            }
            catch (Exception ex)
            {
                GPRStatusText = $"加载失败: {ex.Message}";
            }
        }

        private async Task ResetGPRModelAsync()
        {
            var flowId = _paramProvider.GetCurrentFlowId();
            if (string.IsNullOrEmpty(flowId)) return;

            if (!_dialogService.ShowConfirmation("重置 GPR 模型将清除所有训练数据和模型状态，确定继续？", "重置确认"))
                return;

            try
            {
                await _gprService.ResetModelAsync(flowId, keepData: false);
                await LoadGPRStatusAsync(flowId);
                StatusMessage = "GPR 模型已重置";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "重置 GPR 模型失败");
            }
        }

        private async Task RetrainGPRModelAsync()
        {
            var flowId = _paramProvider.GetCurrentFlowId();
            if (string.IsNullOrEmpty(flowId)) return;

            try
            {
                IsLoading = true;
                StatusMessage = "正在重新训练 GPR 模型...";

                // 重置模型但保留数据
                await _gprService.ResetModelAsync(flowId, keepData: true);

                // 重新训练
                var result = await _gprService.TrainModelAsync();
                await _gprService.SaveStateAsync(flowId);
                await LoadGPRStatusAsync(flowId);

                StatusMessage = result.IsActive
                    ? $"重新训练完成: R²={result.RSquared:F4}, RMSE={result.RMSE:F4}"
                    : "数据量不足，模型未激活";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "重新训练 GPR 模型失败");
                StatusMessage = $"训练失败: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }
    }
}