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
    ///  新增: 多响应 GPR 管理器实现
    /// 
    /// 架构说明:
    ///   - _primaryService: 主响应使用已注入的 IGPRModelService 单例（即原来的 GPRModelService）
    ///     所有现有代码（停止条件、ModelAnalysis 页面等）通过 IGPRModelService 接口仍然只看到主响应模型，完全兼容。
    ///   - _secondaryModels: 非主响应各自有一个轻量级的 Python GPR 模型实例
    ///     这些模型直接通过 pythonnet 调用 doe_gpr.py，不经过 GPRModelService 包装
    ///     这样做的原因:
    ///       a. 避免为每个响应注册一个 GPRModelService 单例（Prism DI 不支持 keyed 注册）
    ///       b. 次要模型不需要完整的 GPRModelService 功能（不需要 FindOptimal、ShouldStop 等）
    ///       c. 持久化使用独立的 GPRModelState 记录，FactorSignature 加上响应名后缀区分
    /// 
    /// 持久化策略:
    ///   主模型: FlowId + FactorSignature（与原来一致）
    ///   次要模型: FlowId + FactorSignature + "::" + ResponseName
    ///   例: 主响应"转化率" → sig = "Temperature,Pressure,Catalyst"
    ///       次要"选择性"   → sig = "Temperature,Pressure,Catalyst::选择性"
    ///       次要"成本"     → sig = "Temperature,Pressure,Catalyst::成本"
    /// </summary>
    public class MultiResponseGPRService : IMultiResponseGPRService
    {
        private readonly IGPRModelService _primaryService;
        private readonly IDOERepository _repository;
        private readonly ILogService _logger;

        // 次要响应模型: ResponseName → Python GPR 对象
        private readonly Dictionary<string, dynamic> _secondaryModels = new();
        private readonly Dictionary<string, bool> _secondaryInitialized = new();

        // 配置缓存
        private string _primaryResponseName = "";
        private List<string> _allResponseNames = new();
        private string _flowId = "";
        private string? _projectId;
        private string _baseSignature = "";
        private List<DOEFactor> _factors = new();

        public MultiResponseGPRService(
            IGPRModelService primaryService,
            IDOERepository repository,
            ILogService logger)
        {
            _primaryService = primaryService ?? throw new ArgumentNullException(nameof(primaryService));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _logger = logger?.ForContext<MultiResponseGPRService>() ?? throw new ArgumentNullException(nameof(logger));
        }

        public List<string> ResponseNames => _allResponseNames.ToList();

        public bool AnyModelActive
        {
            get
            {
                if (_primaryService.IsActive) return true;
                return _secondaryModels.Any(kv => IsSecondaryActive(kv.Key));
            }
        }

        public bool AllModelsActive
        {
            get
            {
                if (!_primaryService.IsActive) return false;
                return _secondaryModels.All(kv => IsSecondaryActive(kv.Key));
            }
        }

        // ══════════════ 初始化 ══════════════

        public async Task InitializeAsync(string flowId, List<DOEFactor> factors, List<DOEResponse> responses, string? projectId = null)
        {
            _flowId = flowId;
            _factors = factors;
            _projectId = projectId;
            if (responses.Count == 0)
            {
                _logger.LogWarning("MultiResponseGPR: 没有响应变量，跳过初始化");
                return;
            }

            _primaryResponseName = responses[0].ResponseName;
            _allResponseNames = responses.Select(r => r.ResponseName).ToList();
            _baseSignature = GPRModelState.BuildSignature(factors.Select(f => f.FactorName));

            // 主响应: 使用已有的 IGPRModelService（BatchExecutor 已经在调用它了）
            // 这里不重复初始化，由 BatchExecutor.StartBatchAsync 负责
            _logger.LogInformation("MultiResponseGPR: 主响应=\"{Primary}\", 次要响应={Count}",
                _primaryResponseName, responses.Count - 1);

            // 次要响应: 各自创建独立的 Python GPR 模型
            var factorNamesJson = JsonConvert.SerializeObject(factors.Select(f => f.FactorName).ToList());

            // ★ 修复 (v3): bounds 区分连续因子 [lower, upper] 和类别因子 ["A", "B", "C"]
            // 原来的 bug: 所有因子都传 [LowerBound, UpperBound]，类别因子无法正确初始化
            var boundsDict = new Dictionary<string, object>();
            foreach (var f in factors)
            {
                if (f.IsCategorical)
                    boundsDict[f.FactorName] = f.GetCategoryLevelList();
                else
                    boundsDict[f.FactorName] = new[] { f.LowerBound, f.UpperBound };
            }
            var boundsJson = JsonConvert.SerializeObject(boundsDict);

            // ★ 修复 (v3): 传入因子类型字典
            var factorTypesJson = JsonConvert.SerializeObject(
                factors.ToDictionary(f => f.FactorName,
                    f => f.IsCategorical ? "categorical" : "continuous"));

            for (int i = 1; i < responses.Count; i++)
            {
                var respName = responses[i].ResponseName;
                try
                {
                    var secondarySig = BuildSecondarySignature(respName);

                    // 尝试从数据库恢复
                    var state = await _repository.GetGPRModelStateAsync(flowId, secondarySig);

                    EnsurePythonEnv();

                    using (Py.GIL())
                    {
                        dynamic doeGpr = Py.Import("doe_gpr");

                        if (state?.ModelStateBytes != null && state.ModelStateBytes.Length > 0)
                        {
                            _secondaryModels[respName] = doeGpr.load_model(state.ModelStateBytes);
                            _logger.LogInformation("MultiResponseGPR: 恢复次要模型 \"{Name}\", data={Count}",
                                respName, state.DataCount);
                        }
                        else
                        {
                            _secondaryModels[respName] = doeGpr.create_model(factorNamesJson, boundsJson, 6, factorTypesJson);
                            _logger.LogInformation("MultiResponseGPR: 新建次要模型 \"{Name}\"", respName);
                        }
                    }

                    _secondaryInitialized[respName] = true;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "MultiResponseGPR: 初始化次要模型 \"{Name}\" 失败", respName);
                    _secondaryInitialized[respName] = false;
                }
            }
        }

        // ══════════════ 数据追加 ══════════════

        /// <summary>
        /// ★ 修复 (v3): factorValues 改为 Dictionary&lt;string, object&gt; 以透传类别因子标签
        /// </summary>
        public void AppendAllResponses(
            Dictionary<string, object> factorValues,
            Dictionary<string, double> responseValues,
            string source = "measured",
            string batchName = "",
            string timestamp = "")
        {
            if (string.IsNullOrEmpty(timestamp))
                timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // 主响应: 通过 IGPRModelService 追加
            if (responseValues.TryGetValue(_primaryResponseName, out var primaryVal))
            {
                try
                {
                    _primaryService.AppendData(factorValues, primaryVal, source, batchName, timestamp);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "MultiResponseGPR: 主响应 \"{Name}\" 数据追加失败", _primaryResponseName);
                }
            }

            // 次要响应: 直接调用 Python 模型
            var factorsJson = JsonConvert.SerializeObject(factorValues);
            foreach (var kv in _secondaryModels)
            {
                if (!responseValues.TryGetValue(kv.Key, out var respVal)) continue;
                if (_secondaryInitialized.TryGetValue(kv.Key, out var ok) && !ok) continue;

                try
                {
                    using (Py.GIL())
                    {
                        kv.Value.append_data(factorsJson, respVal, source, batchName, timestamp);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "MultiResponseGPR: 次要响应 \"{Name}\" 数据追加失败", kv.Key);
                }
            }
        }

        // ══════════════ 训练 ══════════════

        public async Task<Dictionary<string, GPRTrainResult>> TrainAllAsync()
        {
            var results = new Dictionary<string, GPRTrainResult>();

            // 主响应
            try
            {
                var primaryResult = await _primaryService.TrainModelAsync();
                results[_primaryResponseName] = primaryResult;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "MultiResponseGPR: 主响应训练失败");
                results[_primaryResponseName] = new GPRTrainResult { IsActive = false };
            }

            // 次要响应
            foreach (var kv in _secondaryModels)
            {
                if (_secondaryInitialized.TryGetValue(kv.Key, out var ok) && !ok)
                {
                    results[kv.Key] = new GPRTrainResult { IsActive = false };
                    continue;
                }

                try
                {
                    using (Py.GIL())
                    {
                        string resultJson = kv.Value.train().ToString();
                        var pyResult = JsonConvert.DeserializeObject<SecondaryTrainResult>(resultJson);
                        results[kv.Key] = new GPRTrainResult
                        {
                            IsActive = pyResult?.IsActive ?? false,
                            RSquared = pyResult?.RSquared ?? 0,
                            RMSE = pyResult?.RMSE ?? 0,
                            Lengthscales = pyResult?.Lengthscales ?? new(),
                            DataCount = pyResult?.DataCount ?? 0
                        };
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "MultiResponseGPR: 次要响应 \"{Name}\" 训练失败", kv.Key);
                    results[kv.Key] = new GPRTrainResult { IsActive = false };
                }
            }

            return results;
        }

        // ══════════════ 多响应预测（Desirability 核心） ══════════════

        /// <summary>
        /// ★ 修复 (v3): factorValues 改为 Dictionary&lt;string, object&gt; 以支持类别因子标签
        /// </summary>
        public Dictionary<string, GPRPrediction> PredictAll(Dictionary<string, object> factorValues)
        {
            var predictions = new Dictionary<string, GPRPrediction>();

            // 主响应
            try
            {
                if (_primaryService.IsActive)
                {
                    predictions[_primaryResponseName] = _primaryService.Predict(factorValues);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "MultiResponseGPR: 主响应预测失败");
            }

            // 次要响应
            var factorsJson = JsonConvert.SerializeObject(factorValues);
            foreach (var kv in _secondaryModels)
            {
                if (!IsSecondaryActive(kv.Key)) continue;

                try
                {
                    using (Py.GIL())
                    {
                        string resultJson = kv.Value.predict(factorsJson).ToString();
                        var pyResult = JsonConvert.DeserializeObject<SecondaryPrediction>(resultJson);
                        if (pyResult != null)
                        {
                            predictions[kv.Key] = new GPRPrediction
                            {
                                Mean = pyResult.Mean,
                                Std = pyResult.Std
                            };
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "MultiResponseGPR: 次要响应 \"{Name}\" 预测失败", kv.Key);
                }
            }

            return predictions;
        }
        /// <summary>
        /// ★ 新增: GIL 安全版本的多响应预测
        /// 调用者必须已持有 GIL（例如在 Python 回调委托中）
        /// </summary>
        public Dictionary<string, GPRPrediction> PredictAllInsideGIL(Dictionary<string, object> factorValues)
        {
            var predictions = new Dictionary<string, GPRPrediction>();

            // 主响应: 需要确认 IGPRModelService.Predict 内部是否获取 GIL
            // 如果是，同样需要一个 PredictInsideGIL 版本
            // 目前先直接调用（pythonnet 的 Py.GIL() 在同一线程上是可重入的）
            try
            {
                if (_primaryService.IsActive)
                {
                    predictions[_primaryResponseName] = _primaryService.Predict(factorValues);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "MultiResponseGPR: 主响应预测失败 (InsideGIL)");
            }

            // 次要响应: 直接调用 Python 对象，不再获取 GIL
            var factorsJson = JsonConvert.SerializeObject(factorValues);
            foreach (var kv in _secondaryModels)
            {
                if (!IsSecondaryActive(kv.Key)) continue;

                try
                {
                    // ★ 关键修复: 不使用 using(Py.GIL())
                    string resultJson = kv.Value.predict(factorsJson).ToString();
                    var pyResult = JsonConvert.DeserializeObject<SecondaryPrediction>(resultJson);
                    if (pyResult != null)
                    {
                        predictions[kv.Key] = new GPRPrediction
                        {
                            Mean = pyResult.Mean,
                            Std = pyResult.Std
                        };
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "MultiResponseGPR: 次要响应 \"{Name}\" 预测失败 (InsideGIL)", kv.Key);
                }
            }

            return predictions;
        }

        // ══════════════ 持久化 ══════════════

        public async Task SaveAllAsync(string flowId)
        {
            // 主响应: 由 IGPRModelService.SaveStateAsync 负责（BatchExecutor 已经在调用了）
            // 这里只保存次要模型

            foreach (var kv in _secondaryModels)
            {
                if (_secondaryInitialized.TryGetValue(kv.Key, out var ok) && !ok) continue;

                try
                {
                    var secondarySig = BuildSecondarySignature(kv.Key);

                    byte[] modelBytes;
                    string trainingDataJson;
                    string evolutionJson;
                    int dataCount;
                    bool isActive;

                    using (Py.GIL())
                    {
                        dynamic pyBytes = kv.Value.serialize();
                        modelBytes = (byte[])pyBytes;
                        trainingDataJson = kv.Value.get_training_data().ToString();
                        evolutionJson = kv.Value.get_evolution_data().ToString();
                        dataCount = (int)kv.Value.data_count;
                        isActive = (bool)kv.Value.is_active;
                    }

                    double? rSquared = null;
                    double? rmse = null;

                    if (isActive)
                    {
                        // ★ 修复 (v3): 不再重复训练，直接从 evolution 历史中提取最近的 R²/RMSE
                        // 原来的 bug: 每次保存时调 train() 导致次要模型被训练两次（执行循环中一次+保存时一次）
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
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "MultiResponseGPR: 次要模型 \"{Name}\" 读取 evolution 失败", kv.Key);
                        }
                    }

                    var state = new GPRModelState
                    {
                        FlowId = flowId,
                        ProjectId = _projectId,             // ★ 新增
                        FactorSignature = secondarySig,
                        ModelName = $"次要模型-{kv.Key}",
                        ModelStateBytes = modelBytes,
                        TrainingDataJson = trainingDataJson,
                        EvolutionHistoryJson = evolutionJson,
                        DataCount = dataCount,
                        RSquared = rSquared,
                        RMSE = rmse,
                        IsActive = isActive,
                        LastTrainedTime = dataCount > 0 ? DateTime.Now : null
                    };

                    await _repository.SaveGPRModelStateAsync(state);
                    _logger.LogInformation("MultiResponseGPR: 保存次要模型 \"{Name}\": data={Count}, active={Active}",
                        kv.Key, dataCount, isActive);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "MultiResponseGPR: 保存次要模型 \"{Name}\" 失败", kv.Key);
                }
            }
        }

        // ══════════════ 状态查询 ══════════════

        public List<ResponseModelStatus> GetModelStatuses()
        {
            var statuses = new List<ResponseModelStatus>();

            // 主响应
            statuses.Add(new ResponseModelStatus
            {
                ResponseName = _primaryResponseName,
                IsPrimary = true,
                IsActive = _primaryService.IsActive,
                DataCount = _primaryService.DataCount,
                RSquared = null // 需要额外调用才能获取，这里先简化
            });

            // 次要响应
            foreach (var kv in _secondaryModels)
            {
                int dataCount = 0;
                bool isActive = false;

                try
                {
                    if (_secondaryInitialized.TryGetValue(kv.Key, out var ok) && ok)
                    {
                        using (Py.GIL())
                        {
                            dataCount = (int)kv.Value.data_count;
                            isActive = (bool)kv.Value.is_active;
                        }
                    }
                }
                catch { /* 安全忽略 */ }

                statuses.Add(new ResponseModelStatus
                {
                    ResponseName = kv.Key,
                    IsPrimary = false,
                    IsActive = isActive,
                    DataCount = dataCount
                });
            }

            return statuses;
        }

        // ══════════════ 内部辅助 ══════════════

        private string BuildSecondarySignature(string responseName)
        {
            return $"{_baseSignature}::{responseName}";
        }

        private bool IsSecondaryActive(string responseName)
        {
            if (!_secondaryInitialized.TryGetValue(responseName, out var ok) || !ok) return false;
            if (!_secondaryModels.TryGetValue(responseName, out var model)) return false;

            try
            {
                using (Py.GIL())
                {
                    return (bool)model.is_active;
                }
            }
            catch
            {
                return false;
            }
        }

        private void EnsurePythonEnv()
        {
            var envManager = MaxChemical.Modules.Designer.Services.PythonEnvironmentManager.Instance;
            if (!envManager.IsInitialized)
            {
                if (!envManager.Initialize())
                    throw new InvalidOperationException("Python 环境初始化失败");
            }
        }

        // ── Python JSON DTO ──

        private class SecondaryTrainResult
        {
            [JsonProperty("is_active")] public bool IsActive { get; set; }
            [JsonProperty("r_squared")] public double RSquared { get; set; }
            [JsonProperty("rmse")] public double RMSE { get; set; }
            [JsonProperty("lengthscales")] public Dictionary<string, double>? Lengthscales { get; set; }
            [JsonProperty("data_count")] public int DataCount { get; set; }
        }

        private class SecondaryPrediction
        {
            [JsonProperty("mean")] public double Mean { get; set; }
            [JsonProperty("std")] public double Std { get; set; }
        }
    }
}