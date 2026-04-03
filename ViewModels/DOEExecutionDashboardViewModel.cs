using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using MaxChemical.Infrastructure.Services;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Events;
using MaxChemical.Modules.DOE.Models;
using MaxChemical.Modules.DOE.Services;
using MaxChemical.Modules.DOE.Views;
using Newtonsoft.Json;
using Prism.Commands;
using Prism.Events;
using Prism.Mvvm;

namespace MaxChemical.Modules.DOE.ViewModels
{
    public class DOEExecutionDashboardViewModel : BindableBase
    {
        private readonly IDOEExecutionService _executionService;
        private readonly IDOERepository _repository;
        private readonly IEventAggregator _eventAggregator;
        private readonly IDialogService _dialogService;
        private readonly ILogService _logger;

        // ── 批次基本信息 ──
        private string _batchId = "";
        private string _batchName = "";
        private DOEBatchStatus _batchStatus = DOEBatchStatus.Designing;

        // ── 执行状态 ──
        private int _currentRun;
        private int _totalRuns;
        private double _progressPercent;
        private bool _isRunning;
        private bool _isPaused;
        private bool _isPausing;

        // ── 指标卡片 ──
        private string _statusMessage = "就绪";
        private string _bestRunText = "—";
        private string _bestValueText = "暂无结果";
        private string _gprStatusText = "未创建";
        private string _gprDetailText = "";
        private string _etaText = "—";
        private string _etaDetailText = "";

        // ── 动态实验矩阵表 ──
        private DataTable _matrixTable = new();

        // ── 缓存 ──
        private List<DOEResponse> _responses = new();
        private List<DOEFactor> _factors = new();
        private List<string> _factorNames = new();
        private List<string> _responseNames = new();
        private DateTime? _firstRunStartTime;

        // ── 统计（按钮状态用）──
        private int _pendingCount;
        private int _failedCount;

        // ──  迷你模式 ──
        private bool _isMiniMode;
        private ObservableCollection<ProgressSegment> _progressSegments = new();
        private ObservableCollection<FactorDisplayItem> _currentFactorItems = new();
        private bool _hasCurrentFactors;

        public DOEExecutionDashboardViewModel(
            IDOEExecutionService executionService,
            IDOERepository repository,
            IEventAggregator eventAggregator,
            IDialogService dialogService,
            ILogService logger)
        {
            _executionService = executionService;
            _repository = repository;
            _eventAggregator = eventAggregator;
            _dialogService = dialogService;
            _logger = logger?.ForContext<DOEExecutionDashboardViewModel>()
                      ?? throw new ArgumentNullException(nameof(logger));

            InitializeCommands();
            SubscribeEvents();
        }

        // ══════════════ Properties ══════════════

        public string BatchId { get => _batchId; set => SetProperty(ref _batchId, value); }
        public string BatchName { get => _batchName; set => SetProperty(ref _batchName, value); }
        public DOEBatchStatus BatchStatus { get => _batchStatus; set => SetProperty(ref _batchStatus, value); }

        public int CurrentRun
        {
            get => _currentRun;
            set { if (SetProperty(ref _currentRun, value)) { RaisePropertyChanged(nameof(ProgressText)); RaisePropertyChanged(nameof(ProgressBarWidth)); RaisePropertyChanged(nameof(StartButtonText)); UpdateETA(); UpdateProgressSegments(); } }
        }

        public int TotalRuns
        {
            get => _totalRuns;
            set { if (SetProperty(ref _totalRuns, value)) { RaisePropertyChanged(nameof(ProgressText)); RaisePropertyChanged(nameof(ProgressBarWidth)); UpdateETA(); UpdateProgressSegments(); } }
        }

        public double ProgressPercent
        {
            get => _progressPercent;
            set { if (SetProperty(ref _progressPercent, value)) RaisePropertyChanged(nameof(ProgressBarWidth)); }
        }

        public bool IsRunning { get => _isRunning; set { SetProperty(ref _isRunning, value); RaiseCommandCanExecute(); } }
        public bool IsPaused { get => _isPaused; set { SetProperty(ref _isPaused, value); RaiseCommandCanExecute(); UpdateProgressSegments(); } }
        public bool IsPausing { get => _isPausing; set => SetProperty(ref _isPausing, value); }

        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
        public string ProgressText => $"{CurrentRun} / {TotalRuns}";
        public double ProgressBarWidth => TotalRuns > 0 ? ProgressPercent / 100.0 * 600 : 0;
        public string BestRunText { get => _bestRunText; set => SetProperty(ref _bestRunText, value); }
        public string BestValueText { get => _bestValueText; set => SetProperty(ref _bestValueText, value); }
        public string GPRStatusText { get => _gprStatusText; set => SetProperty(ref _gprStatusText, value); }
        public string GPRDetailText { get => _gprDetailText; set => SetProperty(ref _gprDetailText, value); }
        public string ETAText { get => _etaText; set => SetProperty(ref _etaText, value); }
        public string ETADetailText { get => _etaDetailText; set => SetProperty(ref _etaDetailText, value); }

        public DataTable MatrixTable { get => _matrixTable; set => SetProperty(ref _matrixTable, value); }

        public int PendingCount { get => _pendingCount; set { SetProperty(ref _pendingCount, value); RaiseCommandCanExecute(); } }
        public int FailedCount { get => _failedCount; set { SetProperty(ref _failedCount, value); RaiseCommandCanExecute(); } }
        public bool HasFailedRuns => FailedCount > 0;
        public string StartButtonText => CurrentRun > 0 ? "继续执行" : "开始执行";

        // ──  迷你模式属性 ──

        /// <summary>是否处于迷你监控模式</summary>
        public bool IsMiniMode
        {
            get => _isMiniMode;
            set => SetProperty(ref _isMiniMode, value);
        }

        /// <summary>分段进度条数据（用于 ItemsControl 绑定）</summary>
        public ObservableCollection<ProgressSegment> ProgressSegments
        {
            get => _progressSegments;
            set => SetProperty(ref _progressSegments, value);
        }

        /// <summary>当前因子参数列表（用于迷你面板三列展示）</summary>
        public ObservableCollection<FactorDisplayItem> CurrentFactorItems
        {
            get => _currentFactorItems;
            set => SetProperty(ref _currentFactorItems, value);
        }

        /// <summary>是否有当前因子数据可显示</summary>
        public bool HasCurrentFactors
        {
            get => _hasCurrentFactors;
            set => SetProperty(ref _hasCurrentFactors, value);
        }

        // ══════════════ Commands ══════════════

        public DelegateCommand ContinueCommand { get; private set; } = null!;
        public DelegateCommand RetryFailedCommand { get; private set; } = null!;
        public DelegateCommand PauseCommand { get; private set; } = null!;
        public DelegateCommand ResumeCommand { get; private set; } = null!;
        public DelegateCommand AbortCommand { get; private set; } = null!;
        public DelegateCommand ToggleMiniModeCommand { get; private set; } = null!;

        public DelegateCommand StartCommand => ContinueCommand;

        private void InitializeCommands()
        {
            ContinueCommand = new DelegateCommand(async () => await ContinueExecutionAsync(),
                () => !IsRunning && !string.IsNullOrEmpty(BatchId) && PendingCount > 0);
            RetryFailedCommand = new DelegateCommand(async () => await RetryFailedAsync(),
                () => !IsRunning && !string.IsNullOrEmpty(BatchId) && FailedCount > 0);
            PauseCommand = new DelegateCommand(() =>
            {
                _executionService.PauseBatch();
                IsPausing = true;
                StatusMessage = "正在等待当前组完成后暂停...";
            }, () => IsRunning && !IsPaused && !IsPausing);
            ResumeCommand = new DelegateCommand(() =>
            {
                _executionService.ResumeBatch();
                IsPaused = false;
                IsPausing = false;
                StatusMessage = "已恢复执行...";
                IsMiniMode = true;
            }, () => IsPaused);
            AbortCommand = new DelegateCommand(async () => await AbortExecutionAsync(), () => IsRunning || IsPaused);
            ToggleMiniModeCommand = new DelegateCommand(() => IsMiniMode = !IsMiniMode);
        }

        private void RaiseCommandCanExecute()
        {
            Application.Current?.Dispatcher.InvokeAsync(() =>
            {
                ContinueCommand.RaiseCanExecuteChanged();
                RetryFailedCommand.RaiseCanExecuteChanged();
                PauseCommand.RaiseCanExecuteChanged();
                ResumeCommand.RaiseCanExecuteChanged();
                AbortCommand.RaiseCanExecuteChanged();
                RaisePropertyChanged(nameof(HasFailedRuns));
                RaisePropertyChanged(nameof(IsPausing));
            });
        }

        // ══════════════ 加载批次 ══════════════

        public async Task LoadBatchAsync(string batchId)
        {
            try
            {
                var batch = await _repository.GetBatchWithDetailsAsync(batchId);
                if (batch == null)
                {
                    _dialogService.ShowError($"未找到批次: {batchId}", "错误");
                    return;
                }

                BatchId = batchId;
                BatchName = batch.BatchName;
                BatchStatus = batch.Status;
                _responses = batch.Responses;
                _factors = batch.Factors;
                _factorNames = _factors.Select(f => f.FactorName).ToList();
                _responseNames = _responses.Select(r => r.ResponseName).ToList();

                var activeRuns = batch.Runs.Where(r => r.Status != DOERunStatus.Skipped).ToList();
                var completedCount = activeRuns.Count(r => r.Status == DOERunStatus.Completed);
                PendingCount = activeRuns.Count(r => r.Status == DOERunStatus.Pending || r.Status == DOERunStatus.Suspended);
                FailedCount = activeRuns.Count(r => r.Status == DOERunStatus.Failed);
                TotalRuns = activeRuns.Count;
                CurrentRun = completedCount;
                ProgressPercent = TotalRuns > 0 ? (double)completedCount / TotalRuns * 100 : 0;

                BuildMatrixTable(batch.Runs);
                UpdateBestResult();
                await UpdateGPRStatusAsync(batch.FlowId);
                UpdateETA();
                UpdateProgressSegments();

                StatusMessage = $"已加载: {batch.BatchName}, {PendingCount} 组待执行, {completedCount} 组已完成"
                    + (FailedCount > 0 ? $", {FailedCount} 组失败" : "");
                ContinueCommand.RaiseCanExecuteChanged();
                RetryFailedCommand.RaiseCanExecuteChanged();

                _logger.LogInformation("加载 DOE 批次: {BatchId}, 总计 {Total} 组, 已完成 {Done} 组",
                    batchId, TotalRuns, completedCount);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载批次失败: {BatchId}", batchId);
                _dialogService.ShowError($"加载失败: {ex.Message}", "错误");
            }
        }

        // ══════════════ 动态表格构建 ══════════════

        private void BuildMatrixTable(List<DOERunRecord> runs)
        {
            var dt = new DataTable();

            dt.Columns.Add("#", typeof(int));
            dt.Columns.Add("_RunIndex", typeof(int));
            foreach (var name in _factorNames)
                dt.Columns.Add(name, typeof(string));
            foreach (var name in _responseNames)
                dt.Columns.Add(name, typeof(string));
            dt.Columns.Add("状态", typeof(string));
            dt.Columns.Add("耗时", typeof(string));

            var visibleRuns = runs
                .Where(r => r.Status != DOERunStatus.Skipped)
                .OrderBy(r => r.RunIndex);

            int displayIndex = 1;
            foreach (var run in visibleRuns)
            {
                var row = dt.NewRow();
                row["#"] = displayIndex++;
                row["_RunIndex"] = run.RunIndex;

                var factors = JsonConvert.DeserializeObject<Dictionary<string, object>>(run.FactorValuesJson)
                              ?? new Dictionary<string, object>();
                foreach (var name in _factorNames)
                    row[name] = factors.TryGetValue(name, out var v) ? v?.ToString() ?? "—" : "—";

                var responses = !string.IsNullOrEmpty(run.ResponseValuesJson)
                    ? JsonConvert.DeserializeObject<Dictionary<string, double>>(run.ResponseValuesJson) ?? new()
                    : new Dictionary<string, double>();
                foreach (var name in _responseNames)
                    row[name] = responses.TryGetValue(name, out var rv) ? rv.ToString("F2") : "—";

                row["状态"] = run.Status switch
                {
                    DOERunStatus.Pending => "待执行",
                    DOERunStatus.Running => "执行中",
                    DOERunStatus.WaitingResponse => "等待输入",
                    DOERunStatus.Completed => "已完成",
                    DOERunStatus.Failed => "失败",
                    DOERunStatus.Skipped => "已跳过",
                    DOERunStatus.Suspended => "已暂停",
                    _ => run.Status.ToString()
                };

                row["耗时"] = run.StartTime.HasValue && run.EndTime.HasValue
                    ? (run.EndTime.Value - run.StartTime.Value).ToString(@"mm\:ss")
                    : "—";

                dt.Rows.Add(row);
            }

            MatrixTable = dt;
        }

        // ══════════════ 指标更新 ══════════════

        private void UpdateBestResult()
        {
            if (MatrixTable.Rows.Count == 0 || _responseNames.Count == 0)
            {
                BestRunText = "—";
                BestValueText = "暂无结果";
                return;
            }

            var primaryResp = _responseNames[0];
            double bestVal = double.MinValue;
            int bestIndex = -1;

            foreach (DataRow row in MatrixTable.Rows)
            {
                var valStr = row[primaryResp]?.ToString();
                if (valStr != null && valStr != "—" && double.TryParse(valStr, out var val) && val > bestVal)
                {
                    bestVal = val;
                    bestIndex = Convert.ToInt32(row["#"]);
                }
            }

            if (bestIndex > 0)
            {
                BestRunText = $"第 {bestIndex} 组";
                BestValueText = $"{primaryResp} = {bestVal:F2}";
            }
            else
            {
                BestRunText = "—";
                BestValueText = "暂无结果";
            }
        }

        private async Task UpdateGPRStatusAsync(string flowId)
        {
            try
            {
                var signature = GPRModelState.BuildSignature(_factorNames);
                var state = await _repository.GetGPRModelStateAsync(flowId, signature);

                if (state == null)
                {
                    GPRStatusText = "未创建";
                    GPRDetailText = "";
                }
                else if (state.IsActive)
                {
                    GPRStatusText = $"R² = {state.RSquared:F4}";
                    GPRDetailText = $"{state.DataCount} 组 · RMSE {state.RMSE:F2}";
                }
                else
                {
                    GPRStatusText = $"冷启动 {state.DataCount}/6";
                    GPRDetailText = $"需 {Math.Max(0, 6 - state.DataCount)} 组激活";
                }
            }
            catch
            {
                GPRStatusText = "—";
                GPRDetailText = "";
            }
        }

        private void UpdateETA()
        {
            if (CurrentRun == 0 || !_firstRunStartTime.HasValue)
            {
                ETAText = "—";
                ETADetailText = $"{TotalRuns} 组实验";
                return;
            }

            var elapsed = DateTime.Now - _firstRunStartTime.Value;
            var avgPerRun = elapsed.TotalSeconds / CurrentRun;
            var remaining = (TotalRuns - CurrentRun) * avgPerRun;

            if (remaining < 60)
                ETAText = $"~{remaining:F0} 秒";
            else
                ETAText = $"~{remaining / 60:F0} 分钟";

            ETADetailText = $"均 {avgPerRun:F0}s/组 · 剩 {TotalRuns - CurrentRun} 组";
        }

        // ══════════════  迷你模式辅助方法 ══════════════

        /// <summary>
        /// 根据当前进度生成分段进度条数据
        /// </summary>
        private void UpdateProgressSegments()
        {
            if (TotalRuns <= 0) return;

            // 最多显示 9 段，超过时按比例分组
            int segCount = Math.Min(TotalRuns, 9);
            var segments = new ObservableCollection<ProgressSegment>();

            double runsPerSeg = (double)TotalRuns / segCount;

            for (int i = 0; i < segCount; i++)
            {
                double segEnd = (i + 1) * runsPerSeg;
                if (CurrentRun >= segEnd)
                    segments.Add(ProgressSegment.Completed(IsPaused));
                else if (CurrentRun > i * runsPerSeg)
                    segments.Add(ProgressSegment.Running(IsPaused));
                else
                    segments.Add(ProgressSegment.Pending());
            }

            ProgressSegments = segments;
        }

        /// <summary>
        /// 更新当前因子参数显示（从最新执行/正在执行的 run 中获取）
        /// </summary>
        private void UpdateCurrentFactors(DOERunRecord? run)
        {
            if (run == null || string.IsNullOrEmpty(run.FactorValuesJson))
            {
                HasCurrentFactors = false;
                return;
            }

            try
            {
                var factors = JsonConvert.DeserializeObject<Dictionary<string, object>>(run.FactorValuesJson);
                if (factors == null || factors.Count == 0)
                {
                    HasCurrentFactors = false;
                    return;
                }

                var items = new ObservableCollection<FactorDisplayItem>();
                foreach (var kv in factors)
                {
                    // 尝试找到对应因子的单位信息
                    var factorDef = _factors.FirstOrDefault(f => f.FactorName == kv.Key);
                    string unit = "";
                    // 简单的单位推断：根据因子名
                    if (kv.Key.Contains("温度") || kv.Key.Contains("Temperature")) unit = "℃";
                    else if (kv.Key.Contains("流量") || kv.Key.Contains("Flow")) unit = kv.Key.Contains("H2") || kv.Key.Contains("氢") ? "sccm" : "mL/min";
                    else if (kv.Key.Contains("压力") || kv.Key.Contains("Pressure")) unit = "bar";

                    // 缩短因子名用于显示
                    string shortName = kv.Key
                        .Replace("Target", "")
                        .Replace("target", "");

                    // ★ 修复 (v3): 类别因子显示标签，连续因子显示数值
                    string displayValue;
                    if (kv.Value is double d)
                        displayValue = d.ToString("F2");
                    else if (double.TryParse(kv.Value?.ToString(), out var parsed))
                        displayValue = parsed.ToString("F2");
                    else
                        displayValue = kv.Value?.ToString() ?? "—";

                    items.Add(new FactorDisplayItem
                    {
                        Label = shortName,
                        Value = displayValue,
                        Unit = unit
                    });
                }

                CurrentFactorItems = items;
                HasCurrentFactors = items.Count > 0;
            }
            catch
            {
                HasCurrentFactors = false;
            }
        }

        // ══════════════ 执行控制 ══════════════

        private async Task ContinueExecutionAsync()
        {
            try
            {
                if (CurrentRun > 0)
                {
                    var confirm = _dialogService.ShowConfirmation(
                        $"当前已完成 {CurrentRun}/{TotalRuns} 组。\n将继续执行剩余 {PendingCount} 组实验。\n\n确定继续？",
                        "继续执行");
                    if (!confirm) return;
                }

                var batch = await _repository.GetBatchWithDetailsAsync(BatchId);
                if (batch != null)
                {
                    var suspendedRuns = batch.Runs.Where(r => r.Status == DOERunStatus.Suspended).ToList();
                    foreach (var run in suspendedRuns)
                    {
                        run.Status = DOERunStatus.Pending;
                        await _repository.UpdateRunAsync(run);
                    }
                    if (suspendedRuns.Count > 0)
                        _logger.LogInformation("恢复 {Count} 组 Suspended 为 Pending", suspendedRuns.Count);
                }

                IsRunning = true;
                IsPaused = false;
                _firstRunStartTime = DateTime.Now;
                StatusMessage = $"正在执行... 剩余 {PendingCount} 组";

                //  切换到迷你模式
                IsMiniMode = true;

                _executionService.RunCompleted += OnRunCompleted;
                _executionService.ProgressChanged += OnProgressChanged;
                _executionService.BatchCompleted += OnBatchCompleted;
                _executionService.PauseConfirmed += OnPauseConfirmed;

                await _executionService.StartBatchAsync(BatchId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "启动执行失败");
                _dialogService.ShowError($"启动失败: {ex.Message}", "错误");
                IsRunning = false;
                IsMiniMode = false;
            }
        }

        private void OnPauseConfirmed(object? sender, EventArgs e)
        {
            Application.Current?.Dispatcher.InvokeAsync(() =>
            {
                IsPausing = false;
                IsPaused = true;
                StatusMessage = "已暂停。点击「恢复」继续执行。";
            });
        }

        private async Task RetryFailedAsync()
        {
            try
            {
                var confirm = _dialogService.ShowConfirmation(
                    $"将重新执行 {FailedCount} 组失败的实验。\n失败组的状态将重置为「待执行」。\n\n确定重试？",
                    "重试失败组");
                if (!confirm) return;

                IsRunning = true;
                IsPaused = false;
                _firstRunStartTime = DateTime.Now;
                StatusMessage = $"正在重置失败组...";

                var batch = await _repository.GetBatchWithDetailsAsync(BatchId);
                if (batch == null) return;

                var failedRuns = batch.Runs.Where(r => r.Status == DOERunStatus.Failed).ToList();
                foreach (var run in failedRuns)
                {
                    run.Status = DOERunStatus.Pending;
                    run.ResponseValuesJson = null;
                    run.StartTime = null;
                    run.EndTime = null;
                    await _repository.UpdateRunAsync(run);
                }

                await LoadBatchAsync(BatchId);

                StatusMessage = $"已重置 {failedRuns.Count} 组，正在执行...";

                //  切换到迷你模式
                IsMiniMode = true;

                _executionService.RunCompleted += OnRunCompleted;
                _executionService.ProgressChanged += OnProgressChanged;
                _executionService.BatchCompleted += OnBatchCompleted;
                _executionService.PauseConfirmed += OnPauseConfirmed;

                await _executionService.StartBatchAsync(BatchId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "重试失败组出错");
                _dialogService.ShowError($"重试失败: {ex.Message}", "错误");
                IsRunning = false;
            }
        }

        private async Task AbortExecutionAsync()
        {
            var confirm = _dialogService.ShowConfirmation("确定要终止当前批次执行？", "终止确认");
            if (!confirm) return;

            try
            {
                await _executionService.AbortBatchAsync();
                StatusMessage = "批次已终止";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "终止执行失败");
            }
        }

        // ══════════════ 事件处理 ══════════════

        private void SubscribeEvents()
        {
            _eventAggregator.GetEvent<DOERunCompletedEvent>().Subscribe(OnDOERunNeedsResponse, ThreadOption.UIThread);
            _eventAggregator.GetEvent<DOEProgressChangedEvent>().Subscribe(OnDOEProgressChanged, ThreadOption.UIThread);
        }

        private void OnDOERunNeedsResponse(DOERunCompletedPayload payload)
        {
            if (payload.BatchId != BatchId) return;

            try
            {
                var factorValues = JsonConvert.DeserializeObject<Dictionary<string, object>>(payload.FactorValuesJson)
                                   ?? new Dictionary<string, object>();

                var responseDefs = _responses.Select(r => (r.ResponseName, r.Unit)).ToList();

                var dialog = new DOEResponseCollectionDialog(
                    payload.RunIndex, TotalRuns, factorValues, responseDefs);
                dialog.Owner = Application.Current.MainWindow;

                var result = dialog.ShowDialog();

                if (result == true && dialog.CollectedValues != null)
                    _executionService.SubmitRunResponse(payload.RunIndex, dialog.CollectedValues);
                else
                    _executionService.SubmitRunResponse(payload.RunIndex, new Dictionary<string, double>());
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "弹框采集响应值失败");
                _executionService.SubmitRunResponse(payload.RunIndex, new Dictionary<string, double>());
            }
        }

        private void OnDOEProgressChanged(DOEProgressPayload payload)
        {
            if (payload.BatchId != BatchId) return;
            StatusMessage = payload.StatusMessage;
        }

        private void OnRunCompleted(object? sender, DOERunCompletedEventArgs e)
        {
            Application.Current?.Dispatcher.InvokeAsync(async () =>
            {
                try
                {
                    var run = await _repository.GetRunAsync(BatchId, e.RunIndex);
                    if (run != null)
                    {
                        UpdateTableRow(run);
                        UpdateBestResult();
                        RecalculateProgress();

                        //  更新迷你面板的因子参数显示
                        UpdateCurrentFactors(run);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "更新结果表失败");
                }
            });
        }

        private void OnProgressChanged(object? sender, DOEProgressEventArgs e)
        {
            Application.Current?.Dispatcher.InvokeAsync(() =>
            {
                StatusMessage = e.StatusMessage;
            });
        }

        private void OnBatchCompleted(object? sender, DOEBatchCompletedEventArgs e)
        {
            Application.Current?.Dispatcher.InvokeAsync(() =>
            {
                IsRunning = false;
                IsPaused = false;
                IsPausing = false;
                BatchStatus = e.FinalStatus;
                StatusMessage = $"批次完成: {e.StopReason} (完成 {e.CompletedRuns}/{e.TotalRuns} 组)";

                //  执行完成，恢复大窗口
                IsMiniMode = false;

                _executionService.RunCompleted -= OnRunCompleted;
                _executionService.ProgressChanged -= OnProgressChanged;
                _executionService.BatchCompleted -= OnBatchCompleted;
                _executionService.PauseConfirmed -= OnPauseConfirmed;

                RecalculateProgress();
                _dialogService.ShowInfo(StatusMessage, "DOE 执行完成");
            });
        }

        // ══════════════ 表格行更新 ══════════════

        private void UpdateTableRow(DOERunRecord run)
        {
            foreach (DataRow row in MatrixTable.Rows)
            {
                if (Convert.ToInt32(row["_RunIndex"]) == run.RunIndex)
                {
                    var responses = !string.IsNullOrEmpty(run.ResponseValuesJson)
                        ? JsonConvert.DeserializeObject<Dictionary<string, double>>(run.ResponseValuesJson) ?? new()
                        : new Dictionary<string, double>();
                    foreach (var name in _responseNames)
                        row[name] = responses.TryGetValue(name, out var rv) ? rv.ToString("F2") : "—";

                    row["状态"] = run.Status switch
                    {
                        DOERunStatus.Completed => "已完成",
                        DOERunStatus.Failed => "失败",
                        DOERunStatus.Running => "执行中",
                        DOERunStatus.WaitingResponse => "等待输入",
                        DOERunStatus.Suspended => "已暂停",
                        _ => run.Status.ToString()
                    };

                    row["耗时"] = run.StartTime.HasValue && run.EndTime.HasValue
                        ? (run.EndTime.Value - run.StartTime.Value).ToString(@"mm\:ss")
                        : "—";

                    break;
                }
            }
        }

        private void RecalculateProgress()
        {
            int completed = 0, pending = 0, failed = 0;
            foreach (DataRow row in MatrixTable.Rows)
            {
                var status = row["状态"]?.ToString();
                if (status == "已完成") completed++;
                else if (status == "待执行" || status == "已暂停") pending++;
                else if (status == "失败") failed++;
            }

            CurrentRun = completed;
            PendingCount = pending;
            FailedCount = failed;
            ProgressPercent = TotalRuns > 0 ? (double)completed / TotalRuns * 100 : 0;
        }
    }
}