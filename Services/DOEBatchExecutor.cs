using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using MaxChemical.Infrastructure.DOE;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Events;
using MaxChemical.Modules.DOE.Models;
using MaxChemical.Modules.DOE.Views;
using Newtonsoft.Json;
using Prism.Events;

namespace MaxChemical.Modules.DOE.Services
{
    /// <summary>
    /// DOE 批量执行器 — 修复版 v2：
    /// 
    ///  修复内容:
    /// 1. [严重] 多响应 GPR: 使用 IMultiResponseGPRService 同时喂所有响应变量的数据
    /// 2. [严重] 停止条件 bestObserved 方向性: 根据 Desirability 配置决定 max/min
    /// 3. [中等] CollectResponseFromDialog 空返回保护: 用户关闭对话框时不丢失 run 状态
    /// 4. [中等] GPR 超时泄漏监控: 计数器跟踪超时任务数量
    /// 5. [小] SubmitRunResponse 空方法: 标注 deprecated
    /// 
    /// 原有修复保留:
    /// - GIL 超时保护 (30s)
    /// - 每组 try-catch 容错
    /// - 每步详细日志
    /// </summary>
    public class DOEBatchExecutor : IDOEExecutionService
    {
        private readonly IFlowExecutionService _flowExecution;
        private readonly IFlowParameterProvider _paramProvider;
        private readonly IDOERepository _repository;
        private readonly IGPRModelService _gprService;
        private readonly IMultiResponseGPRService _multiGprService;  //  新增
        private readonly IEventAggregator _eventAggregator;
        private readonly ILogService _logger;

        private DOEBatchStatus _batchState = DOEBatchStatus.Designing;
        private int _currentRunIndex = -1;
        private string _currentBatchId = "";
        private CancellationTokenSource? _cts;
        private readonly ManualResetEventSlim _pauseEvent = new(true);

        private DOEBatch? _currentBatch;
        private List<DOERunRecord> _runs = new();
        private List<DOEStopCondition> _stopConditions = new();
        private List<DOEResponse> _responses = new();

        /// <summary>GPR 操作超时（秒）— 防止 Python GIL 死锁阻塞循环</summary>
        private const int GPR_TIMEOUT_SECONDS = 30;

        /// <summary> 新增: 跟踪因超时而泄漏的 GPR 任务数量</summary>
        private int _leakedGprTaskCount = 0;
        private const int LEAKED_TASK_WARNING_THRESHOLD = 5;

        /// <summary>
        ///  新增: Desirability 配置缓存，用于确定停止条件中 bestObserved 的方向
        /// </summary>
        private List<DesirabilityResponseConfig> _desirabilityConfigs = new();

        public DOEBatchExecutor(
            IFlowExecutionService flowExecution,
            IFlowParameterProvider paramProvider,
            IDOERepository repository,
            IGPRModelService gprService,
            IMultiResponseGPRService multiGprService,  //  新增
            IEventAggregator eventAggregator,
            ILogService logger)
        {
            _flowExecution = flowExecution ?? throw new ArgumentNullException(nameof(flowExecution));
            _paramProvider = paramProvider ?? throw new ArgumentNullException(nameof(paramProvider));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _gprService = gprService ?? throw new ArgumentNullException(nameof(gprService));
            _multiGprService = multiGprService ?? throw new ArgumentNullException(nameof(multiGprService));
            _eventAggregator = eventAggregator ?? throw new ArgumentNullException(nameof(eventAggregator));
            _logger = logger?.ForContext<DOEBatchExecutor>() ?? throw new ArgumentNullException(nameof(logger));
        }

        public DOEBatchStatus BatchState
        {
            get => _batchState;
            private set { _batchState = value; _logger.LogInformation("DOE 批次状态: {State}", value); }
        }

        public int CurrentRunIndex => _currentRunIndex;

        public event EventHandler<DOERunCompletedEventArgs>? RunCompleted;
        public event EventHandler<DOEBatchCompletedEventArgs>? BatchCompleted;
        public event EventHandler<DOEProgressEventArgs>? ProgressChanged;
        public event EventHandler? PauseConfirmed;

        // ══════════════ 启动 ══════════════

        public async Task StartBatchAsync(string batchId, CancellationToken cancellationToken = default)
        {
            if (BatchState == DOEBatchStatus.Running)
                throw new InvalidOperationException("已有批次正在执行中");

            try
            {
                _currentBatch = await _repository.GetBatchWithDetailsAsync(batchId);
                if (_currentBatch == null) throw new InvalidOperationException($"未找到批次: {batchId}");

                _currentBatchId = batchId;
                _runs = _currentBatch.Runs
                    .Where(r => r.DataSource == DOEDataSource.Measured && r.Status == DOERunStatus.Pending)
                    .OrderBy(r => r.RunIndex).ToList();
                _stopConditions = _currentBatch.StopConditions;
                _responses = _currentBatch.Responses;

                if (_runs.Count == 0) throw new InvalidOperationException("没有待执行的实验组");

                // 初始化 GPR 模型（主响应）
                try { await _gprService.InitializeModelAsync(_currentBatch.FlowId, _currentBatch.Factors); }
                catch (Exception ex) { _logger.LogError(ex, "GPR 模型初始化失败，无模型模式执行"); }

                //  新增: 初始化多响应 GPR 模型
                try
                {
                    await _multiGprService.InitializeAsync(_currentBatch.FlowId, _currentBatch.Factors, _responses);
                    _logger.LogInformation("多响应 GPR 初始化成功: {Count} 个响应", _responses.Count);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "多响应 GPR 初始化失败，使用主响应模式继续");
                }

                //  新增: 加载 Desirability 配置（用于停止条件方向判断）
                try
                {
                    var configJson = await _repository.GetDesirabilityConfigJsonAsync(batchId);
                    if (!string.IsNullOrEmpty(configJson))
                    {
                        _desirabilityConfigs = JsonConvert.DeserializeObject<List<DesirabilityResponseConfig>>(configJson)
                                               ?? new();
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "加载 Desirability 配置失败，停止条件使用默认方向（maximize）");
                }

                _cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                _pauseEvent.Set();
                _currentRunIndex = -1;
                _leakedGprTaskCount = 0;

                BatchState = DOEBatchStatus.Running;
                await _repository.UpdateBatchStatusAsync(batchId, DOEBatchStatus.Running);
                _eventAggregator.GetEvent<DOEBatchStartedEvent>().Publish(batchId);

                _logger.LogInformation("DOE 批次开始: {BatchId}, {Count} 组待执行", batchId, _runs.Count);

                await Task.Run(async () => await ExecuteBatchLoopAsync());
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "DOE 批次启动失败");
                BatchState = DOEBatchStatus.Aborted;
                await _repository.UpdateBatchStatusAsync(batchId, DOEBatchStatus.Aborted);
                throw;
            }
        }

        public void PauseBatch()
        {
            if (BatchState != DOEBatchStatus.Running) return;
            _pauseEvent.Reset();
            BatchState = DOEBatchStatus.Paused;
        }

        public void ResumeBatch()
        {
            if (BatchState != DOEBatchStatus.Paused) return;
            _pauseEvent.Set();
            BatchState = DOEBatchStatus.Running;
        }

        public async Task AbortBatchAsync()
        {
            _cts?.Cancel();
            _pauseEvent.Set();
            await _flowExecution.StopExecutionAsync();
            BatchState = DOEBatchStatus.Aborted;
            await _repository.UpdateBatchStatusAsync(_currentBatchId, DOEBatchStatus.Aborted);
            await SaveGPRModelSafe();
            await SaveMultiGPRModelSafe();  //  新增
            NotifyBatchCompleted(DOEBatchStatus.Aborted, "用户手动终止");
        }

        /// <summary>
        /// [Deprecated] 该方法原为异步提交响应值的入口，但当前执行循环使用同步弹框模式。
        /// 保留接口兼容性，不做任何操作。如需异步模式，请重构为 TaskCompletionSource 模式。
        /// </summary>
        [Obsolete("当前使用 CollectResponseFromDialog 同步模式，此方法不执行任何操作")]
        public void SubmitRunResponse(int runIndex, Dictionary<string, double> responseValues) { }

        // ══════════════ 主循环 ══════════════

        private async Task ExecuteBatchLoopAsync()
        {
            string? stopReason = null;
            int completedCount = 0;

            var factorNameToParamId = new Dictionary<string, string>();
            if (_currentBatch?.Factors != null)
            {
                foreach (var factor in _currentBatch.Factors)
                {
                    string paramId = !string.IsNullOrEmpty(factor.SourceNodeId) && !string.IsNullOrEmpty(factor.SourceParamName)
                        ? $"{factor.SourceNodeId}.{factor.SourceParamName}"
                        : factor.FactorName;
                    factorNameToParamId[factor.FactorName] = paramId;
                }
            }

            try
            {
                for (int i = 0; i < _runs.Count; i++)
                {
                    _cts!.Token.ThrowIfCancellationRequested();

                    var run = _runs[i];
                    _currentRunIndex = run.RunIndex;

                    _logger.LogInformation("═══════ DOE 开始第 {Index}/{Total} 组 (RunIndex={RunIndex}) ═══════",
                        i + 1, _runs.Count, run.RunIndex);

                    NotifyProgress(i, _runs.Count, $"正在执行第 {i + 1}/{_runs.Count} 组实验...");

                    try
                    {
                        // ── Step 1: 注入参数 ──
                        // ★ 修复 (v3): 反序列化为 Dictionary<string, object> 以保留类别因子的字符串标签
                        // 原来的 bug: Dictionary<string, double> 会丢失类别因子标签，
                        // GPR 的 _encode_factors() 收到数字后 str(0.0) 不匹配任何水平标签
                        var factorValuesRaw = JsonConvert.DeserializeObject<Dictionary<string, object>>(run.FactorValuesJson)
                                           ?? new Dictionary<string, object>();

                        // 构建类别因子名称集合
                        var categoricalFactorNames = _currentBatch?.Factors?
                            .Where(f => f.IsCategorical).Select(f => f.FactorName).ToHashSet()
                            ?? new HashSet<string>();

                        // 提取连续因子的 double 值（用于参数注入和事件通知）
                        var factorValues = new Dictionary<string, double>();
                        foreach (var kv in factorValuesRaw)
                        {
                            if (categoricalFactorNames.Contains(kv.Key))
                                continue; // 类别因子不参与连续因子的 double dict
                            if (kv.Value is double d) factorValues[kv.Key] = d;
                            else if (kv.Value is long l) factorValues[kv.Key] = l;
                            else if (double.TryParse(kv.Value?.ToString(), out var parsed)) factorValues[kv.Key] = parsed;
                        }

                        // 注入连续因子参数到流程
                        var injectParams = new Dictionary<string, object>();

                        foreach (var kv in factorValues)
                        {
                            var paramId = factorNameToParamId.TryGetValue(kv.Key, out var pid) ? pid : kv.Key;
                            injectParams[paramId] = kv.Value;
                        }

                        // 记录类别因子信息（供日志和 UI 显示）
                        foreach (var catName in categoricalFactorNames)
                        {
                            if (factorValuesRaw.TryGetValue(catName, out var catVal))
                            {
                                _logger.LogDebug("第 {Index} 组: 类别因子 {Name}={Value}（手动操作）",
                                    i + 1, catName, catVal);
                            }
                        }

                        _logger.LogInformation("第 {Index} 组: 注入参数（{Count} 个连续因子）...", i + 1, injectParams.Count);
                        bool injected = injectParams.Count == 0 || _paramProvider.InjectParameters(injectParams);
                        if (!injected)
                        {
                            _logger.LogWarning("第 {Index} 组: 参数注入失败，跳过", i + 1);
                            run.Status = DOERunStatus.Failed;
                            await _repository.UpdateRunAsync(run);
                            continue;
                        }
                        _logger.LogInformation("第 {Index} 组: 参数注入成功", i + 1);

                        // ── Step 2: 执行流程 ──
                        run.Status = DOERunStatus.Running;
                        run.StartTime = DateTime.Now;
                        await _repository.UpdateRunAsync(run);

                        _logger.LogInformation("第 {Index} 组: 调用 ExecuteCurrentFlowAsync...", i + 1);
                        var flowResult = await _flowExecution.ExecuteCurrentFlowAsync(_cts.Token);
                        _logger.LogInformation("第 {Index} 组: 流程返回 Success={Success}", i + 1, flowResult.IsSuccessful);

                        if (!flowResult.IsSuccessful)
                        {
                            _logger.LogWarning("第 {Index} 组: 流程执行失败: {Error}", i + 1, flowResult.ErrorMessage);
                            run.Status = DOERunStatus.Failed;
                            run.EndTime = DateTime.Now;
                            await _repository.UpdateRunAsync(run);
                            RunCompleted?.Invoke(this, new DOERunCompletedEventArgs
                            {
                                RunIndex = run.RunIndex,
                                FactorValues = factorValues,
                                IsSuccessful = false
                            });
                            continue;
                        }

                        run.ExperimentId = flowResult.ExperimentId;

                        // ── Step 3: 同步弹框采集响应值 ──
                        run.Status = DOERunStatus.WaitingResponse;
                        await _repository.UpdateRunAsync(run);
                        NotifyProgress(i, _runs.Count, $"第 {i + 1} 组流程完成，等待输入结果...");

                        _logger.LogInformation("第 {Index} 组: 弹框等待用户输入...", i + 1);
                        var responseValues = CollectResponseFromDialog(run.RunIndex, _runs.Count, factorValuesRaw);
                        _logger.LogInformation("第 {Index} 组: 用户输入完成，响应值数量={Count}", i + 1, responseValues.Count);

                        //  修复: 用户关闭对话框（未输入）时标记为 Failed 而不是静默丢失
                        if (responseValues.Count == 0)
                        {
                            _logger.LogWarning("第 {Index} 组: 用户未输入响应值，标记为 Failed", i + 1);
                            run.Status = DOERunStatus.Failed;
                            run.EndTime = DateTime.Now;
                            await _repository.UpdateRunAsync(run);
                            RunCompleted?.Invoke(this, new DOERunCompletedEventArgs
                            {
                                RunIndex = run.RunIndex,
                                FactorValues = factorValues,
                                IsSuccessful = false
                            });
                            continue;
                        }

                        // ── Step 4: 保存结果 ──
                        run.ResponseValuesJson = JsonConvert.SerializeObject(responseValues);
                        run.Status = DOERunStatus.Completed;
                        run.EndTime = DateTime.Now;
                        await _repository.UpdateRunAsync(run);
                        completedCount++;

                        _logger.LogInformation("第 {Index} 组完成: 响应={Responses}", i + 1, run.ResponseValuesJson);

                        RunCompleted?.Invoke(this, new DOERunCompletedEventArgs
                        {
                            RunIndex = run.RunIndex,
                            FactorValues = factorValues,
                            ResponseValues = responseValues,
                            IsSuccessful = true
                        });

                        // ── Step 4.5:  多响应 GPR 训练（加超时保护）──
                        // ★ 修复 (v3): 传完整因子值（含类别因子标签）给 GPR
                        _logger.LogInformation("第 {Index} 组: 开始多响应 GPR 训练...", i + 1);
                        await FeedGPRModelWithTimeoutAsync(factorValuesRaw, responseValues);
                        _logger.LogInformation("第 {Index} 组: GPR 训练完成", i + 1);

                        // ── Step 5:  停止条件（加超时保护）──
                        _logger.LogInformation("第 {Index} 组: 评估停止条件...", i + 1);
                        var stopDecision = EvaluateAllStopConditionsWithTimeout(responseValues, i);
                        _logger.LogInformation("第 {Index} 组: 停止条件: ShouldStop={Stop}", i + 1, stopDecision.ShouldStop);

                        if (stopDecision.ShouldStop)
                        {
                            stopReason = stopDecision.Reason;
                            _logger.LogInformation("停止条件满足: {Reason}", stopReason);
                            for (int j = i + 1; j < _runs.Count; j++)
                            {
                                _runs[j].Status = DOERunStatus.Suspended;
                                await _repository.UpdateRunAsync(_runs[j]);

                            }
                            break;

                        }
                        await CheckPauseAsync(i);
                        _logger.LogInformation("═══════ DOE 第 {Index}/{Total} 组结束，继续下一组 ═══════", i + 1, _runs.Count);
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception runEx)
                    {
                        _logger.LogError(runEx, "第 {Index} 组异常: {Message}", i + 1, runEx.Message);
                        run.Status = DOERunStatus.Failed;
                        run.EndTime = DateTime.Now;
                        try { await _repository.UpdateRunAsync(run); } catch { }
                        RunCompleted?.Invoke(this, new DOERunCompletedEventArgs
                        {
                            RunIndex = run.RunIndex,
                            IsSuccessful = false
                        });
                    }
                }

                BatchState = DOEBatchStatus.Completed;
                await _repository.UpdateBatchStatusAsync(_currentBatchId, DOEBatchStatus.Completed);
                NotifyBatchCompleted(DOEBatchStatus.Completed, stopReason ?? "所有实验组已执行完毕");

                // ★ 新增: 项目模式下发布轮次完成事件，触发决策分析弹窗
                if (_currentBatch?.BelongsToProject == true)
                {
                    try
                    {
                        _eventAggregator.GetEvent<DOERoundCompletedEvent>().Publish(
                            new DOERoundCompletedPayload
                            {
                                ProjectId = _currentBatch.ProjectId!,
                                BatchId = _currentBatchId,
                                RoundNumber = _currentBatch.RoundNumber ?? 0,
                                FinalStatus = DOEBatchStatus.Completed
                            });
                        _logger.LogInformation(
                            "项目轮次完成事件已发布: Project={ProjectId}, Round={Round}",
                            _currentBatch.ProjectId, _currentBatch.RoundNumber);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "发布项目轮次完成事件失败（不影响批次完成）");
                    }
                }
            }
            catch (OperationCanceledException)
            {
                _logger.LogInformation("DOE 批次被取消");
                BatchState = DOEBatchStatus.Aborted;
                await _repository.UpdateBatchStatusAsync(_currentBatchId, DOEBatchStatus.Aborted);
                NotifyBatchCompleted(DOEBatchStatus.Aborted, "用户取消");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "DOE 批次执行异常");
                BatchState = DOEBatchStatus.Aborted;
                await _repository.UpdateBatchStatusAsync(_currentBatchId, DOEBatchStatus.Aborted);
                NotifyBatchCompleted(DOEBatchStatus.Aborted, $"异常: {ex.Message}");
            }
            finally
            {
                await SaveGPRModelSafe();
                await SaveMultiGPRModelSafe();  //  新增
                _currentRunIndex = -1;
                _cts?.Dispose();
                _cts = null;
                if (BatchState == DOEBatchStatus.Running || BatchState == DOEBatchStatus.Paused)
                    BatchState = DOEBatchStatus.Completed;
            }
        }

        // ══════════════  GPR 训练（带超时保护）══════════════

        /// <summary>
        /// ★ 修复 (v3): factorValues 改为 Dictionary&lt;string, object&gt; 以透传类别因子标签
        /// </summary>
        private async Task FeedGPRModelWithTimeoutAsync(Dictionary<string, object> factorValues, Dictionary<string, double> responseValues)
        {
            if (responseValues.Count == 0) return;

            //  新增: 检查泄漏任务数量
            if (_leakedGprTaskCount >= LEAKED_TASK_WARNING_THRESHOLD)
            {
                _logger.LogWarning("已有 {Count} 个 GPR 训练任务因超时泄漏，Python GIL 可能存在严重竞争问题。" +
                    "建议重启应用或检查设备回调线程。", _leakedGprTaskCount);
            }

            try
            {
                using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(GPR_TIMEOUT_SECONDS));

                var gprTask = Task.Run(async () => await FeedGPRModelAsync(factorValues, responseValues));

                var completedTask = await Task.WhenAny(gprTask, Task.Delay(GPR_TIMEOUT_SECONDS * 1000, timeoutCts.Token));

                if (completedTask == gprTask)
                {
                    timeoutCts.Cancel();  // 取消超时计时器
                    await gprTask;        // 传播异常（如果有）
                }
                else
                {
                    Interlocked.Increment(ref _leakedGprTaskCount);
                    _logger.LogWarning("GPR 训练超时（{Timeout}s），跳过本次训练，继续执行下一组。泄漏任务数: {Leaked}",
                        GPR_TIMEOUT_SECONDS, _leakedGprTaskCount);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GPR 训练异常（已跳过）");
            }
        }

        /// <summary>
        /// ★ 修复 (v3): factorValues 改为 Dictionary&lt;string, object&gt; 以透传类别因子标签
        /// </summary>
        private async Task FeedGPRModelAsync(Dictionary<string, object> factorValues, Dictionary<string, double> responseValues)
        {
            if (responseValues.Count == 0) return;
            try
            {
                //  修改: 同时喂所有响应变量（原来只喂主响应）
                _multiGprService.AppendAllResponses(
                    factorValues, responseValues,
                    "measured", _currentBatch!.BatchName,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                // 训练所有模型
                var trainResults = await _multiGprService.TrainAllAsync();

                // 取主响应的训练结果用于事件通知和保存
                var primaryResponseName = _responses.FirstOrDefault()?.ResponseName ?? "";
                var primaryResult = trainResults.TryGetValue(primaryResponseName, out var pr)
                    ? pr : new GPRTrainResult { IsActive = false };

                //  保存主模型（与原来一致）
                try
                {
                    await _gprService.SaveStateAsync(_currentBatch!.FlowId);
                    _logger.LogInformation("GPR 主模型已保存: DataCount={Count}, IsActive={Active}, R²={R2}",
                        primaryResult.DataCount, primaryResult.IsActive, primaryResult.RSquared);
                }
                catch (Exception saveEx)
                {
                    _logger.LogError(saveEx, "GPR 主模型保存失败（数据仍在内存中）");
                }

                //  新增: 保存次要模型
                try
                {
                    await _multiGprService.SaveAllAsync(_currentBatch!.FlowId);
                }
                catch (Exception saveEx)
                {
                    _logger.LogError(saveEx, "GPR 次要模型保存失败（数据仍在内存中）");
                }

                // 日志：各响应训练状态
                foreach (var kv in trainResults)
                {
                    _logger.LogInformation("GPR [{Response}]: Active={Active}, Data={Count}, R²={R2}",
                        kv.Key, kv.Value.IsActive, kv.Value.DataCount, kv.Value.RSquared);
                }

                _eventAggregator.GetEvent<GPRModelUpdatedEvent>().Publish(new GPRModelUpdatedPayload
                {
                    FlowId = _currentBatch!.FlowId,
                    IsActive = primaryResult.IsActive,
                    RSquared = primaryResult.RSquared,
                    DataCount = primaryResult.DataCount
                });
            }
            catch (Exception ex) { _logger.LogError(ex, "GPR 更新失败"); }
        }

        // ══════════════  停止条件（带超时保护）══════════════

        private StopEvaluationResult EvaluateAllStopConditionsWithTimeout(Dictionary<string, double> responseValues, int currentIndex)
        {
            try
            {
                // 阈值条件不涉及 Python，直接评估
                var thresholdResult = EvaluateThresholdConditions(responseValues);
                if (thresholdResult.ShouldStop) return thresholdResult;

                // GPR 模型条件涉及 Python GIL，加超时
                var modelTask = Task.Run(() => EvaluateModelStopCondition(currentIndex));
                if (modelTask.Wait(TimeSpan.FromSeconds(GPR_TIMEOUT_SECONDS)))
                {
                    return modelTask.Result.ShouldStop ? modelTask.Result : new StopEvaluationResult { ShouldStop = false };
                }
                else
                {
                    _logger.LogWarning("GPR 停止条件评估超时（{Timeout}s），跳过模型判断", GPR_TIMEOUT_SECONDS);
                    return new StopEvaluationResult { ShouldStop = false };
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "停止条件评估异常");
                return new StopEvaluationResult { ShouldStop = false };
            }
        }

        // ══════════════ 弹框采集响应值 ══════════════

        private Dictionary<string, double> CollectResponseFromDialog(
            int runIndex, int totalRuns, Dictionary<string, object> factorValues)
        {
            Dictionary<string, double> result = new();

            Application.Current.Dispatcher.Invoke(() =>
            {
                var responseDefs = _responses.Select(r => (r.ResponseName, r.Unit)).ToList();
                var dialog = new DOEResponseCollectionDialog(runIndex, totalRuns, factorValues, responseDefs);
                dialog.Owner = Application.Current.MainWindow;
                dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;

                var dialogResult = dialog.ShowDialog();

                if (dialogResult == true && dialog.CollectedValues != null)
                {
                    result = dialog.CollectedValues;
                }
            });

            return result;
        }

        // ══════════════ 停止条件（内部） ══════════════

        private StopEvaluationResult EvaluateThresholdConditions(Dictionary<string, double> responseValues)
        {
            var conditions = _stopConditions.Where(c => c.ConditionType == DOEStopConditionType.Threshold).ToList();
            if (conditions.Count == 0) return new StopEvaluationResult { ShouldStop = false };

            bool useAnd = conditions.Any(c => c.LogicGroup == DOELogicGroup.And);
            var results = new List<bool>();

            foreach (var cond in conditions)
            {
                if (string.IsNullOrEmpty(cond.ResponseName)) continue;
                if (!responseValues.TryGetValue(cond.ResponseName, out var actual)) continue;
                results.Add(EvaluateOperator(actual, cond.Operator, cond.TargetValue));
            }

            if (results.Count == 0) return new StopEvaluationResult { ShouldStop = false };
            bool stop = useAnd ? results.All(r => r) : results.Any(r => r);
            return new StopEvaluationResult { ShouldStop = stop, Reason = stop ? "阈值条件满足" : "" };
        }

        /// <summary>
        ///  修复: bestObserved 方向性
        /// 原来硬编码取 max，但如果目标是 minimize（如成本最低）应该取 min。
        /// 修复: 根据 Desirability 配置的 Goal 决定取 max 还是 min。
        /// </summary>
        private StopEvaluationResult EvaluateModelStopCondition(int currentIndex)
        {
            var modelConds = _stopConditions.Where(c => c.ConditionType == DOEStopConditionType.ModelPrediction).ToList();
            if (modelConds.Count == 0 || !_gprService.IsActive)
                return new StopEvaluationResult { ShouldStop = false };

            try
            {
                // ★ 修复 (v3): 用 Dictionary<string, object> 保留类别因子标签
                var remaining = _runs.Where(r => r.Status == DOERunStatus.Pending)
                    .Select(r => JsonConvert.DeserializeObject<Dictionary<string, object>>(r.FactorValuesJson)!)
                    .Where(d => d != null).ToList();

                if (remaining.Count == 0) return new StopEvaluationResult { ShouldStop = false };

                var primaryResp = _responses.FirstOrDefault()?.ResponseName;
                if (primaryResp == null) return new StopEvaluationResult { ShouldStop = false };

                //  修复: 根据 Desirability 配置确定优化方向
                bool isMinimize = _desirabilityConfigs
                    .Any(c => c.ResponseName == primaryResp && c.Goal == DesirabilityGoal.Minimize);

                var completedValues = _runs
                    .Where(r => r.Status == DOERunStatus.Completed && !string.IsNullOrEmpty(r.ResponseValuesJson))
                    .Select(r =>
                    {
                        var d = JsonConvert.DeserializeObject<Dictionary<string, double>>(r.ResponseValuesJson!);
                        return d != null && d.TryGetValue(primaryResp, out var v) ? v : (double?)null;
                    })
                    .Where(v => v.HasValue)
                    .Select(v => v!.Value)
                    .ToList();

                if (completedValues.Count == 0)
                    return new StopEvaluationResult { ShouldStop = false };

                //  修复: maximize 时取 max，minimize 时取 min
                double bestObs = isMinimize ? completedValues.Min() : completedValues.Max();

                _logger.LogDebug("停止条件: 主响应=\"{Name}\", 方向={Dir}, bestObs={Best}",
                    primaryResp, isMinimize ? "minimize" : "maximize", bestObs);

                // ★ 修复: 传入 maximize 参数（与 Python 端 EI 方向匹配）
                var decision = _gprService.ShouldStop(remaining, bestObs, maximize: !isMinimize);
                return new StopEvaluationResult
                {
                    ShouldStop = decision.ShouldStop,
                    Reason = decision.ShouldStop ? $"GPR: {decision.Reason}" : ""
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GPR 停止判断失败");
                return new StopEvaluationResult { ShouldStop = false };
            }
        }

        private static bool EvaluateOperator(double actual, string op, double target) => op switch
        {
            "Equals" => Math.Abs(actual - target) < 1e-9,
            "NotEquals" => Math.Abs(actual - target) >= 1e-9,
            "GreaterThan" => actual > target,
            "GreaterThanOrEqual" => actual >= target,
            "LessThan" => actual < target,
            "LessThanOrEqual" => actual <= target,
            _ => false
        };

        private async Task SaveGPRModelSafe()
        {
            _logger.LogInformation(" SaveGPRModelSafe 执行: batch={Batch}, dataCount={Count}",
                _currentBatch?.BatchId, _gprService.DataCount);
            try
            {
                if (_currentBatch != null && _gprService.DataCount > 0)
                {
                    await _gprService.SaveStateAsync(_currentBatch.FlowId);
                    _logger.LogInformation("SaveGPRModelSafe 成功: FlowId={FlowId}", _currentBatch.FlowId);
                }
                else
                {
                    _logger.LogWarning(" SaveGPRModelSafe 跳过: batch={Batch}, count={Count}",
                        _currentBatch?.BatchId, _gprService.DataCount);
                }
            }
            catch (Exception ex) { _logger.LogError(ex, " SaveGPRModelSafe 失败"); }
        }

        /// <summary> 新增: 保存所有次要 GPR 模型</summary>
        private async Task SaveMultiGPRModelSafe()
        {
            try
            {
                if (_currentBatch != null)
                {
                    await _multiGprService.SaveAllAsync(_currentBatch.FlowId);
                    _logger.LogInformation("SaveMultiGPRModelSafe 成功");
                }
            }
            catch (Exception ex) { _logger.LogError(ex, " SaveMultiGPRModelSafe 失败"); }
        }

        private async Task CheckPauseAsync(int currentIndex)
        {
            if (_pauseEvent.IsSet) return;

            _logger.LogInformation("暂停生效：在第 {Index} 组后暂停", currentIndex + 1);
            NotifyProgress(currentIndex, _runs.Count, "已暂停，等待恢复...");
            PauseConfirmed?.Invoke(this, EventArgs.Empty);

            await Task.Run(() => _pauseEvent.Wait(_cts!.Token));

            _logger.LogInformation("恢复执行");
            _cts!.Token.ThrowIfCancellationRequested();
        }

        private void NotifyProgress(int current, int total, string message)
        {
            ProgressChanged?.Invoke(this, new DOEProgressEventArgs { CurrentRun = current + 1, TotalRuns = total, StatusMessage = message });
            _eventAggregator.GetEvent<DOEProgressChangedEvent>().Publish(new DOEProgressPayload
            { BatchId = _currentBatchId, CurrentRun = current + 1, TotalRuns = total, StatusMessage = message });
        }

        private void NotifyBatchCompleted(DOEBatchStatus status, string reason)
        {
            var completed = _runs.Count(r => r.Status == DOERunStatus.Completed);
            BatchCompleted?.Invoke(this, new DOEBatchCompletedEventArgs
            { BatchId = _currentBatchId, FinalStatus = status, TotalRuns = _runs.Count, CompletedRuns = completed, StopReason = reason });
            _eventAggregator.GetEvent<DOEBatchCompletedEvent>().Publish(new DOEBatchCompletedPayload
            { BatchId = _currentBatchId, FinalStatus = status, CompletedRuns = completed, StopReason = reason });
        }

        private class StopEvaluationResult { public bool ShouldStop { get; set; } public string Reason { get; set; } = ""; }
    }
}