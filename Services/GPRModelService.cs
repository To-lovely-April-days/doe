using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Models;
using Newtonsoft.Json;
using Python.Runtime;

namespace MaxChemical.Modules.DOE.Services
{
    /// <summary>
    /// GPR 模型服务 — 管理高斯过程回归模型的完整生命周期。
    /// 通过 PythonEnvironmentManager 调用 doe_gpr.py 模块。
    /// </summary>
    public class GPRModelService : IGPRModelService
    {
        private readonly IDOERepository _repository;
        private readonly ILogService _logger;
        private dynamic? _pythonModel;
        private bool _isPythonInitialized;
        private string _currentFlowId = "";
        private string _currentSignature = "";   //  新增
        private string _currentModelName = "";   //  新增
        public GPRModelService(IDOERepository repository, ILogService logger)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _logger = logger?.ForContext<GPRModelService>() ?? throw new ArgumentNullException(nameof(logger));
        }

        public bool IsActive => _pythonModel != null && _isPythonInitialized &&
                                (bool)GetPythonProperty("is_active", false);

        public int DataCount => _pythonModel != null && _isPythonInitialized
            ? (int)GetPythonProperty("data_count", 0) : 0;

        // ══════════════ 初始化 ══════════════

        public async Task InitializeModelAsync(string flowId, List<DOEFactor> factors)
        {
            _currentFlowId = flowId;
            var factorNames = factors.Select(f => f.FactorName).ToList();

            //  计算因子签名
            _currentSignature = GPRModelState.BuildSignature(factorNames);

            try
            {
                //  按签名查模型（替换原来的按 FlowId 查）
                var state = await _repository.GetGPRModelStateAsync(flowId, _currentSignature);

                if (state?.ModelStateBytes != null && state.ModelStateBytes.Length > 0)
                {
                    // 恢复已有模型
                    _currentModelName = state.ModelName;
                    EnsurePythonEnv();

                    using (Py.GIL())
                    {
                        dynamic doeGpr = Py.Import("doe_gpr");
                        _pythonModel = doeGpr.load_model(state.ModelStateBytes);
                    }

                    _isPythonInitialized = true;
                    _logger.LogInformation("GPR 模型恢复: FlowId={FlowId}, Sig={Sig}, Data={Count}",
                        flowId, _currentSignature, state.DataCount);
                    return;
                }

                // 新建模型
                var factorNamesJson = JsonConvert.SerializeObject(factorNames);

                // ★ 修改: bounds 区分连续因子 [lower, upper] 和类别因子 ["A", "B", "C"]
                var boundsDict = new Dictionary<string, object>();
                foreach (var f in factors)
                {
                    if (f.IsCategorical)
                        boundsDict[f.FactorName] = f.GetCategoryLevelList();
                    else
                        boundsDict[f.FactorName] = new[] { f.LowerBound, f.UpperBound };
                }
                var boundsJson = JsonConvert.SerializeObject(boundsDict);

                // ★ 新增: 因子类型字典
                var factorTypesJson = JsonConvert.SerializeObject(
                    factors.ToDictionary(f => f.FactorName,
                        f => f.IsCategorical ? "categorical" : "continuous"));

                //  自动生成 ModelName
                _currentModelName = $"{factorNames.Count}因子-{string.Join(",", factorNames.Take(3))}";
                if (factorNames.Count > 3) _currentModelName += "...";

                EnsurePythonEnv();

                using (Py.GIL())
                {
                    dynamic doeGpr = Py.Import("doe_gpr");
                    _pythonModel = doeGpr.create_model(factorNamesJson, boundsJson, 6, factorTypesJson);
                }

                _isPythonInitialized = true;
                _logger.LogInformation("GPR 新模型: FlowId={FlowId}, Sig={Sig}, Name={Name}",
                    flowId, _currentSignature, _currentModelName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GPR 初始化失败: FlowId={FlowId}", flowId);
                _isPythonInitialized = false;
                throw;
            }
        }

        // ══════════════ 数据追加 ══════════════

        /// <summary>
        /// ★ 修复 (v3): factorValues 改为 Dictionary&lt;string, object&gt; 以支持类别因子标签
        /// </summary>
        public void AppendData(Dictionary<string, object> factorValues, double responseValue,
                        string source = "measured", string batchName = "", string timestamp = "")
        {
            EnsureModelReady();
            var factorsJson = JsonConvert.SerializeObject(factorValues);
            if (string.IsNullOrEmpty(timestamp))
                timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            using (Py.GIL())
            {
                int count = (int)_pythonModel!.append_data(factorsJson, responseValue, source, batchName, timestamp);
                _logger.LogDebug("GPR 数据追加: count={Count}, source={Source}, batch={Batch}", count, source, batchName);
            }
        }
        /// <summary>
        /// 获取内部 Python GPR 模型对象（供 DOEAnalysisService.set_gpr_model 使用）
        /// </summary>
        public dynamic? GetPythonModel()
        {
            return _isPythonInitialized ? _pythonModel : null;
        }
        // ══════════════ 训练 ══════════════

        public async Task<GPRTrainResult> TrainModelAsync()
        {
            EnsureModelReady();

            using (Py.GIL())
            {
                string resultJson = _pythonModel!.train().ToString();
                var pyResult = JsonConvert.DeserializeObject<PythonTrainResult>(resultJson)!;

                return new GPRTrainResult
                {
                    IsActive = pyResult.IsActive,
                    RSquared = pyResult.RSquared,
                    RMSE = pyResult.RMSE,
                    Lengthscales = pyResult.Lengthscales ?? new Dictionary<string, double>(),
                    DataCount = pyResult.DataCount,
                    RSquaredType = pyResult.RSquaredType ?? "LOO-CV"  // ★ 新增
                };
            }
        }

        // ══════════════ 预测 ══════════════

        /// <summary>
        /// ★ 修复 (v3): factorValues 改为 object 以支持类别因子标签
        /// </summary>
        public GPRPrediction Predict(Dictionary<string, object> factorValues)
        {
            EnsureModelReady();

            var factorsJson = JsonConvert.SerializeObject(factorValues);

            using (Py.GIL())
            {
                string resultJson = _pythonModel!.predict(factorsJson).ToString();
                var pyResult = JsonConvert.DeserializeObject<PythonPrediction>(resultJson)!;
                return new GPRPrediction { Mean = pyResult.Mean, Std = pyResult.Std };
            }
        }

        /// <summary>便利重载: 纯连续因子场景直接传 double dict</summary>
        public GPRPrediction Predict(Dictionary<string, double> factorValues)
        {
            return Predict(factorValues.ToDictionary(kv => kv.Key, kv => (object)kv.Value));
        }

        /// <summary>
        /// ★ 修复 (v3): factorValuesList 改为 object 以支持类别因子标签
        /// </summary>
        public List<GPRPrediction> PredictBatch(List<Dictionary<string, object>> factorValuesList)
        {
            EnsureModelReady();

            var matrixJson = JsonConvert.SerializeObject(factorValuesList);

            using (Py.GIL())
            {
                string resultJson = _pythonModel!.predict_batch(matrixJson).ToString();
                var pyResults = JsonConvert.DeserializeObject<List<PythonPrediction>>(resultJson) ?? new();
                return pyResults.Select(p => new GPRPrediction { Mean = p.Mean, Std = p.Std }).ToList();
            }
        }

        /// <summary>便利重载: 纯连续因子场景直接传 double dict list</summary>
        public List<GPRPrediction> PredictBatch(List<Dictionary<string, double>> factorValuesList)
        {
            return PredictBatch(factorValuesList.Select(d => d.ToDictionary(kv => kv.Key, kv => (object)kv.Value)).ToList());
        }

        // ══════════════ 最优搜索 ══════════════

        public GPROptimalResult FindOptimal(bool maximize = true)
        {
            EnsureModelReady();

            using (Py.GIL())
            {
                // ★ 修复: 传入 maximize 参数，支持最小化搜索
                string resultJson = _pythonModel!.find_optimal("EI", maximize).ToString();
                var pyResult = JsonConvert.DeserializeObject<PythonOptimalResult>(resultJson);

                if (pyResult?.OptimalFactors == null)
                    return new GPROptimalResult();

                return new GPROptimalResult
                {
                    OptimalFactors = pyResult.OptimalFactors,
                    PredictedResponse = pyResult.PredictedResponse,
                    PredictionStd = pyResult.PredictionStd,
                    Direction = pyResult.Direction ?? (maximize ? "maximize" : "minimize")  // ★ 新增
                };
            }
        }

        // ══════════════ 智能停止判断 ══════════════

        /// <summary>
        /// ★ 修复 (v3): 参数类型从 Dictionary&lt;string, double&gt; 改为 Dictionary&lt;string, object&gt;
        /// 以支持类别因子标签（字符串值）透传到 Python 的 _encode_factors()
        /// </summary>
        public GPRStopDecision ShouldStop(List<Dictionary<string, object>> remainingRuns, double bestObserved,
                                          bool maximize = true, double eiThreshold = 0.01, double minRunsRatio = 1.5)
        {
            if (!IsActive)
                return new GPRStopDecision { ShouldStop = false, Reason = "模型未激活" };

            var remainingJson = JsonConvert.SerializeObject(remainingRuns);

            using (Py.GIL())
            {
                // ★ 修复: 传入 maximize, ei_threshold, min_runs_ratio 参数
                string resultJson = _pythonModel!.should_stop(
                    remainingJson, bestObserved, maximize, eiThreshold, minRunsRatio).ToString();
                var pyResult = JsonConvert.DeserializeObject<PythonStopDecision>(resultJson)!;

                return new GPRStopDecision
                {
                    ShouldStop = pyResult.ShouldStop,
                    Reason = pyResult.Reason ?? "",
                    BestPredicted = pyResult.BestPredicted,
                    BestPredictedStd = pyResult.BestPredictedStd,
                    MaxEI = pyResult.MaxEI  // ★ 新增
                };
            }
        }

        /// <summary>便利重载: 纯连续因子场景直接传 double dict list</summary>
        public GPRStopDecision ShouldStop(List<Dictionary<string, double>> remainingRuns, double bestObserved,
                                          bool maximize = true, double eiThreshold = 0.01, double minRunsRatio = 1.5)
        {
            return ShouldStop(
                remainingRuns.Select(d => d.ToDictionary(kv => kv.Key, kv => (object)kv.Value)).ToList(),
                bestObserved, maximize, eiThreshold, minRunsRatio);
        }

        // ══════════════ 敏感性 ══════════════

        public Dictionary<string, double> GetSensitivity()
        {
            if (!IsActive) return new Dictionary<string, double>();

            using (Py.GIL())
            {
                string resultJson = _pythonModel!.get_sensitivity().ToString();
                return JsonConvert.DeserializeObject<Dictionary<string, double>>(resultJson) ?? new();
            }
        }

        // ══════════════ 持久化 ══════════════

        public async Task SaveStateAsync(string flowId)
        {
            if (_pythonModel == null) return;

            try
            {
                byte[] modelBytes;
                string trainingDataJson;
                string evolutionJson;
                int dataCount;
                bool isActive;

                // ★ 修复: 不再调用 TrainModelAsync()，避免 FeedGPRModelAsync 中已训练过后的重复训练
                // 原来的 bug: 每组实验完成后 GPR 被训练两次（MultiGPR.TrainAllAsync 一次 + SaveState 一次）
                // 浪费 Python GIL 时间，在 30s 超时保护下可能导致不必要的超时
                using (Py.GIL())
                {
                    dynamic pyBytes = _pythonModel.serialize();
                    modelBytes = (byte[])pyBytes;
                    trainingDataJson = _pythonModel.get_training_data().ToString();
                    evolutionJson = _pythonModel.get_evolution_data().ToString();
                    dataCount = (int)_pythonModel.data_count;
                    isActive = (bool)_pythonModel.is_active;
                }

                // 从 evolution 历史中提取最近的 R²/RMSE（避免重复训练）
                double? rSquared = null;
                double? rmse = null;
                string? hyperparamsJson = null;

                if (isActive)
                {
                    try
                    {
                        var evolution = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(evolutionJson);
                        if (evolution != null && evolution.Count > 0)
                        {
                            var last = evolution[evolution.Count - 1];
                            if (last.TryGetValue("r_squared", out var r2Obj))
                                rSquared = Convert.ToDouble(r2Obj);
                            if (last.TryGetValue("rmse", out var rmseObj))
                                rmse = Convert.ToDouble(rmseObj);
                        }

                        // 提取 lengthscales
                        var sensitivity = GetSensitivity();
                        if (sensitivity.Count > 0)
                            hyperparamsJson = JsonConvert.SerializeObject(sensitivity);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "GPR 读取 evolution/sensitivity 失败（不影响保存）");
                    }
                }

                var state = new GPRModelState
                {
                    FlowId = flowId,
                    FactorSignature = _currentSignature,
                    ModelName = _currentModelName,
                    ModelStateBytes = modelBytes,
                    TrainingDataJson = trainingDataJson,
                    EvolutionHistoryJson = evolutionJson,
                    DataCount = dataCount,
                    RSquared = rSquared,
                    RMSE = rmse,
                    IsActive = isActive,
                    HyperparamsJson = hyperparamsJson,
                    LastTrainedTime = DateTime.Now
                };

                await _repository.SaveGPRModelStateAsync(state);
                _logger.LogInformation("GPR 保存: FlowId={FlowId}, Sig={Sig}, Data={Count}",
                    flowId, _currentSignature, state.DataCount);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存 GPR 失败: FlowId={FlowId}", flowId);
            }
        }

        /// <summary>
        ///  新增: 保存模型初始状态（创建 DOE 方案时调用）
        /// 与 SaveStateAsync 的区别：
        /// 1. 空模型（0条数据）时直接序列化保存，不调用 TrainModelAsync
        /// 2. 有数据（导入历史数据）时才尝试训练
        /// 3. 训练失败也不影响保存（降级为保存未训练状态）
        /// </summary>
        public async Task SaveInitialStateAsync(string flowId)
        {
            if (_pythonModel == null) return;

            try
            {
                byte[] modelBytes;
                string trainingDataJson;
                string evolutionJson;
                int dataCount;
                bool isActive;

                using (Py.GIL())
                {
                    dynamic pyBytes = _pythonModel.serialize();
                    modelBytes = (byte[])pyBytes;
                    trainingDataJson = _pythonModel.get_training_data().ToString();
                    evolutionJson = _pythonModel.get_evolution_data().ToString();
                    dataCount = (int)_pythonModel.data_count;
                    isActive = (bool)_pythonModel.is_active;
                }

                // 有数据时尝试训练（比如导入了历史数据）
                double? rSquared = null;
                double? rmse = null;
                string? hyperparamsJson = null;

                if (dataCount > 0)
                {
                    try
                    {
                        var trainResult = await TrainModelAsync();
                        isActive = trainResult.IsActive;
                        dataCount = trainResult.DataCount;

                        if (trainResult.IsActive)
                        {
                            rSquared = trainResult.RSquared;
                            rmse = trainResult.RMSE;
                            if (trainResult.Lengthscales.Count > 0)
                                hyperparamsJson = JsonConvert.SerializeObject(trainResult.Lengthscales);
                        }

                        // 训练后重新序列化（模型内部状态已变化）
                        using (Py.GIL())
                        {
                            dynamic pyBytes2 = _pythonModel.serialize();
                            modelBytes = (byte[])pyBytes2;
                            trainingDataJson = _pythonModel.get_training_data().ToString();
                            evolutionJson = _pythonModel.get_evolution_data().ToString();
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "GPR 初始训练失败（不影响模型保存）");
                    }
                }

                var state = new GPRModelState
                {
                    FlowId = flowId,
                    FactorSignature = _currentSignature,
                    ModelName = _currentModelName,
                    ModelStateBytes = modelBytes,
                    TrainingDataJson = trainingDataJson,
                    EvolutionHistoryJson = evolutionJson,
                    DataCount = dataCount,
                    RSquared = rSquared,
                    RMSE = rmse,
                    IsActive = isActive,
                    HyperparamsJson = hyperparamsJson,
                    LastTrainedTime = dataCount > 0 ? DateTime.Now : null
                };

                await _repository.SaveGPRModelStateAsync(state);
                _logger.LogInformation("GPR 初始保存: FlowId={FlowId}, Sig={Sig}, Data={Count}, Active={Active}",
                    flowId, _currentSignature, dataCount, isActive);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存 GPR 初始状态失败: FlowId={FlowId}", flowId);
            }
        }

        public async Task<bool> LoadStateAsync(string flowId)
        {
            try
            {
                //  如果有签名就按签名查，没有就按 FlowId 查最新的
                var state = !string.IsNullOrEmpty(_currentSignature)
                    ? await _repository.GetGPRModelStateAsync(flowId, _currentSignature)
                    : await _repository.GetGPRModelStateAsync(flowId);

                if (state?.ModelStateBytes == null || state.ModelStateBytes.Length == 0)
                    return false;

                _currentModelName = state.ModelName;
                _currentSignature = state.FactorSignature;
                EnsurePythonEnv();

                using (Py.GIL())
                {
                    dynamic doeGpr = Py.Import("doe_gpr");
                    _pythonModel = doeGpr.load_model(state.ModelStateBytes);
                }

                _isPythonInitialized = true;
                _currentFlowId = flowId;
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "恢复 GPR 失败: FlowId={FlowId}", flowId);
                return false;
            }
        }

        public async Task ResetModelAsync(string flowId, bool keepData)
        {
            if (_pythonModel != null)
            {
                using (Py.GIL())
                {
                    _pythonModel.reset(keepData);
                }
            }

            if (!keepData)
            {
                await _repository.DeleteGPRModelStateAsync(flowId);
            }

            _logger.LogInformation("GPR 模型已重置: FlowId={FlowId}, keepData={KeepData}", flowId, keepData);
        }

        // ══════════════ 内部辅助 ══════════════

        private void EnsurePythonEnv()
        {
            var envManager = MaxChemical.Modules.Designer.Services.PythonEnvironmentManager.Instance;
            if (!envManager.IsInitialized)
            {
                if (!envManager.Initialize())
                    throw new InvalidOperationException("Python 环境初始化失败");
            }
        }

        private void EnsureModelReady()
        {
            if (_pythonModel == null || !_isPythonInitialized)
                throw new InvalidOperationException("GPR 模型未初始化");
        }

        private object GetPythonProperty(string propName, object defaultValue)
        {
            try
            {
                using (Py.GIL())
                {
                    return propName switch
                    {
                        "is_active" => (bool)_pythonModel!.is_active,
                        "data_count" => (int)_pythonModel!.data_count,
                        _ => defaultValue
                    };
                }
            }
            catch
            {
                return defaultValue;
            }
        }
        // ★ 修复 (v3): 便利重载 — 纯连续因子场景的兼容入口
        // 内部转发到 Dictionary<string, object> 版本
        public void AppendData(Dictionary<string, double> factorValues, double responseValue,
                               string source = "measured", string batchName = "", string timestamp = "")
        {
            var objectDict = factorValues.ToDictionary(kv => kv.Key, kv => (object)kv.Value);
            AppendData(objectDict, responseValue, source, batchName, timestamp);
        }
        // ── Python JSON 映射 DTO ──

        private class PythonTrainResult
        {
            [JsonProperty("is_active")] public bool IsActive { get; set; }
            [JsonProperty("r_squared")] public double RSquared { get; set; }
            [JsonProperty("rmse")] public double RMSE { get; set; }
            [JsonProperty("lengthscales")] public Dictionary<string, double>? Lengthscales { get; set; }
            [JsonProperty("data_count")] public int DataCount { get; set; }
            [JsonProperty("r_squared_type")] public string? RSquaredType { get; set; }  // ★ 新增
        }

        private class PythonPrediction
        {
            [JsonProperty("mean")] public double Mean { get; set; }
            [JsonProperty("std")] public double Std { get; set; }
        }

        private class PythonOptimalResult
        {
            [JsonProperty("optimal_factors")] public Dictionary<string, double>? OptimalFactors { get; set; }
            [JsonProperty("predicted_response")] public double PredictedResponse { get; set; }
            [JsonProperty("prediction_std")] public double PredictionStd { get; set; }
            [JsonProperty("direction")] public string? Direction { get; set; }  // ★ 新增
        }

        private class PythonStopDecision
        {
            [JsonProperty("should_stop")] public bool ShouldStop { get; set; }
            [JsonProperty("reason")] public string? Reason { get; set; }
            [JsonProperty("best_predicted")] public double BestPredicted { get; set; }
            [JsonProperty("best_predicted_std")] public double BestPredictedStd { get; set; }
            [JsonProperty("max_ei")] public double MaxEI { get; set; }  // ★ 新增
        }
    }
}