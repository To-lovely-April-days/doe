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
    ///  修复: Desirability 多响应优化服务
    /// 
    /// 修复内容:
    /// 1. [严重] OptimizeAsync: 使用 IMultiResponseGPRService.PredictAll 替代单一 GPR 预测
    ///    原来的 bug: 把主响应的 GPR 预测值赋给所有响应，导致优化结果完全错误
    ///    修复后: 每个响应使用自己独立的 GPR 模型预测，Desirability 搜索才是正确的
    /// 
    /// 2. [严重] GetProfileDataAsync: 同样使用 PredictAll
    /// 
    /// 3. [新增] OptimizeAsync 前置检查: 验证所有响应模型是否都已激活
    /// </summary>
    public class DesirabilityService : IDesirabilityService
    {
        private readonly IDOERepository _repository;
        private readonly IGPRModelService _gprService;
        private readonly IMultiResponseGPRService _multiGprService;  //  新增
        private readonly ILogService _logger;
        private dynamic? _pythonEngine;
        private bool _isPythonInitialized;
        private dynamic? _olsPredictor;
        // 当前配置缓存
        private List<DesirabilityResponseConfig> _currentConfigs = new();
        private List<DOEFactor> _currentFactors = new();

        public DesirabilityService(
            IDOERepository repository,
            IGPRModelService gprService,
            IMultiResponseGPRService multiGprService,  //  新增
            ILogService logger)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _gprService = gprService ?? throw new ArgumentNullException(nameof(gprService));
            _multiGprService = multiGprService ?? throw new ArgumentNullException(nameof(multiGprService));
            _logger = logger?.ForContext<DesirabilityService>() ?? throw new ArgumentNullException(nameof(logger));
        }

        private void EnsurePythonReady()
        {
            if (_isPythonInitialized && _pythonEngine != null) return;
            try
            {
                var envManager = MaxChemical.Modules.Designer.Services.PythonEnvironmentManager.Instance;
                if (!envManager.IsInitialized && !envManager.Initialize())
                    throw new InvalidOperationException("Python 环境初始化失败");

                using (Py.GIL())
                {
                    dynamic doeDesirability = Py.Import("doe_desirability");
                    _pythonEngine = doeDesirability.create_engine();
                }
                _isPythonInitialized = true;
                _logger.LogInformation(" Desirability 服务初始化成功");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, " Desirability 服务初始化失败");
                throw;
            }
        }

        /// <summary>
        /// 配置多响应优化参数并初始化 Python 引擎
        /// </summary>
        public Task ConfigureAsync(string batchId, List<DesirabilityResponseConfig> configs, List<DOEFactor> factors)
        {
            EnsurePythonReady();

            _currentConfigs = configs;
            _currentFactors = factors;

            // 构建 Python 侧配置 JSON
            var configsList = configs.Select(c => new
            {
                name = c.ResponseName,
                goal = c.Goal.ToString().ToLower(),
                lower = c.Lower,
                upper = c.Upper,
                target = c.Target,
                weight = c.Weight,
                importance = c.Importance,
                shape = c.Shape,
                shape_lower = c.ShapeLower,
                shape_upper = c.ShapeUpper
            }).ToList();

            var factorNames = factors.Select(f => f.FactorName).ToList();

            // ★ 修复 (v3): bounds 区分连续因子 [lower, upper] 和类别因子 ["A", "B", "C"]
            // 原来的 bug: 所有因子都传 [LowerBound, UpperBound]，类别因子在 differential_evolution
            // 中会被搜索为连续值（如 1.37），完全没有意义。
            // 修复: 连续因子传 [lower, upper]，类别因子传水平标签列表。
            // Python 端的 DesirabilityEngine.optimize() 也需要相应处理（见 doe_desirability.py 修复）。
            var boundsDict = new Dictionary<string, object>();
            foreach (var f in factors)
            {
                if (f.IsCategorical)
                    boundsDict[f.FactorName] = f.GetCategoryLevelList();
                else
                    boundsDict[f.FactorName] = new[] { f.LowerBound, f.UpperBound };
            }
            var boundsJson = JsonConvert.SerializeObject(boundsDict);

            var configsJson = JsonConvert.SerializeObject(configsList);
            var factorNamesJson = JsonConvert.SerializeObject(factorNames);

            using (Py.GIL())
            {
                _pythonEngine!.configure(configsJson, factorNamesJson, boundsJson);
            }

            _logger.LogInformation(" Desirability 已配置: {Count} 个响应变量", configs.Count);
            return Task.CompletedTask;
        }

        /// <summary>
        /// 计算给定预测值的综合 Desirability
        /// </summary>
        public Task<DesirabilityResult> ComputeDesirabilityAsync(Dictionary<string, double> predictions)
        {
            EnsurePythonReady();

            var predsJson = JsonConvert.SerializeObject(predictions);

            using (Py.GIL())
            {
                string resultJson = _pythonEngine!.composite_desirability(predsJson).ToString();
                var pyResult = JsonConvert.DeserializeObject<PyDesirabilityResult>(resultJson)!;

                var result = new DesirabilityResult
                {
                    CompositeD = pyResult.CompositeD,
                    Success = true,
                    IndividualD = new()
                };

                if (pyResult.IndividualD != null)
                {
                    foreach (var kv in pyResult.IndividualD)
                    {
                        var config = _currentConfigs.FirstOrDefault(c => c.ResponseName == kv.Key);
                        result.IndividualD.Add(new IndividualDesirability
                        {
                            ResponseName = kv.Key,
                            PredictedValue = kv.Value.Value,
                            D = kv.Value.D,
                            Goal = config?.Goal ?? DesirabilityGoal.Maximize
                        });
                    }
                }

                return Task.FromResult(result);
            }
        }

        /// <summary>
        ///  修复: 搜索最优因子组合
        /// 
        /// 原来的 bug:
        ///   所有响应变量都用同一个 GPR (主响应) 的预测值 → Desirability 搜索结果完全错误
        /// 
        /// 修复:
        ///   使用 IMultiResponseGPRService.PredictAll 获取每个响应的独立预测值
        /// </summary>
        public Task<DesirabilityResult> OptimizeAsync()
        {
            EnsurePythonReady();

            //  修复: 改用多响应 GPR 检查
            if (!_multiGprService.AnyModelActive)
            {
                return Task.FromResult(new DesirabilityResult
                {
                    Success = false,
                    ErrorMessage = "GPR 模型未激活，无法进行多响应优化"
                });
            }

            //  新增: 检查是否所有响应都有激活的模型
            if (!_multiGprService.AllModelsActive)
            {
                var statuses = _multiGprService.GetModelStatuses();
                var inactiveNames = statuses.Where(s => !s.IsActive).Select(s => s.ResponseName).ToList();
                _logger.LogWarning("部分响应模型未激活: {Names}，优化结果可能不完整", string.Join(", ", inactiveNames));
            }

            _logger.LogInformation(" Desirability 优化: 开始搜索最优因子组合");

            using (Py.GIL())
            {
                //  修复: 使用多响应 GPR 预测
                Func<string, string> predictor = (factorsJson) =>
                {
                    try
                    {
                        // ★ 修复 (v3): 用 object 反序列化，保留类别因子标签
                        var factors = JsonConvert.DeserializeObject<Dictionary<string, object>>(factorsJson) ?? new();

                        //  核心修复: 用 PredictAll 获取每个响应的独立预测
                        var allPredictions = _multiGprService.PredictAllInsideGIL(factors);

                        var predictions = new Dictionary<string, double>();
                        foreach (var config in _currentConfigs)
                        {
                            if (allPredictions.TryGetValue(config.ResponseName, out var pred))
                            {
                                predictions[config.ResponseName] = pred.Mean;
                            }
                            else
                            {
                                // 该响应模型未激活，使用中间值作为回退
                                predictions[config.ResponseName] = (config.Lower + config.Upper) / 2.0;
                                _logger.LogDebug("Desirability: 响应 \"{Name}\" 无 GPR 预测，使用回退值",
                                    config.ResponseName);
                            }
                        }

                        return JsonConvert.SerializeObject(predictions);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Desirability 预测回调异常");
                        var fallback = _currentConfigs.ToDictionary(c => c.ResponseName, c => (c.Lower + c.Upper) / 2.0);
                        return JsonConvert.SerializeObject(fallback);
                    }
                };

                // 将 C# 委托转为 Python callable
                dynamic pyPredictor = predictor.ToPython();
                string resultJson = _pythonEngine!.optimize(pyPredictor).ToString();

                var pyResult = JsonConvert.DeserializeObject<PyOptimizeResult>(resultJson)!;

                var result = new DesirabilityResult
                {
                    OptimalFactors = pyResult.OptimalFactors ?? new(),
                    CompositeD = pyResult.CompositeD,
                    Success = pyResult.Success,
                    ErrorMessage = pyResult.Error,
                    IndividualD = new()
                };

                if (pyResult.IndividualD != null)
                {
                    foreach (var kv in pyResult.IndividualD)
                    {
                        var config = _currentConfigs.FirstOrDefault(c => c.ResponseName == kv.Key);
                        result.IndividualD.Add(new IndividualDesirability
                        {
                            ResponseName = kv.Key,
                            PredictedValue = kv.Value.Value,
                            D = kv.Value.D,
                            Goal = config?.Goal ?? DesirabilityGoal.Maximize
                        });
                    }
                }

                _logger.LogInformation(" Desirability 优化完成: D={D}, 成功={Success}", result.CompositeD, result.Success);
                return Task.FromResult(result);
            }
        }

        /// <summary>
        ///  修复: 获取 Profiler 图数据（同样使用多响应预测）
        /// </summary>
        public Task<string> GetProfileDataAsync(int gridSize = 50)
        {
            EnsurePythonReady();

            using (Py.GIL())
            {
                Func<string, string> predictor = (factorsJson) =>
                {
                    try
                    {
                        // ★ 修复 (v3): 用 object 反序列化，保留类别因子标签
                        var factors = JsonConvert.DeserializeObject<Dictionary<string, object>>(factorsJson) ?? new();

                        //  修复: 使用多响应预测
                        var allPredictions = _multiGprService.PredictAllInsideGIL(factors);

                        var predictions = new Dictionary<string, double>();
                        foreach (var config in _currentConfigs)
                        {
                            if (allPredictions.TryGetValue(config.ResponseName, out var pred))
                            {
                                predictions[config.ResponseName] = pred.Mean;
                            }
                            else
                            {
                                predictions[config.ResponseName] = (config.Lower + config.Upper) / 2.0;
                            }
                        }
                        return JsonConvert.SerializeObject(predictions);
                    }
                    catch
                    {
                        var fallback = _currentConfigs.ToDictionary(c => c.ResponseName, c => (c.Lower + c.Upper) / 2.0);
                        return JsonConvert.SerializeObject(fallback);
                    }
                };

                dynamic pyPredictor = predictor.ToPython();
                string resultJson = _pythonEngine!.profile_plot_data(pyPredictor, gridSize).ToString();
                return Task.FromResult(resultJson);
            }
        }

        /// <summary>
        /// 保存配置到数据库
        /// </summary>
        public async Task SaveConfigAsync(string batchId, List<DesirabilityResponseConfig> configs)
        {
            var configJson = JsonConvert.SerializeObject(configs);
            await _repository.SaveDesirabilityConfigAsync(batchId, configJson);
            _logger.LogInformation(" Desirability 配置已保存: batchId={BatchId}", batchId);
        }

        /// <summary>
        /// 从数据库加载配置
        /// </summary>
        public async Task<List<DesirabilityResponseConfig>> LoadConfigAsync(string batchId)
        {
            var configJson = await _repository.GetDesirabilityConfigJsonAsync(batchId);
            if (string.IsNullOrEmpty(configJson))
                return new List<DesirabilityResponseConfig>();

            return JsonConvert.DeserializeObject<List<DesirabilityResponseConfig>>(configJson)
                   ?? new List<DesirabilityResponseConfig>();
        }
        /// <summary>
        /// ★ v2 新增: 基于 OLS 模型的多响应 Desirability 优化
        /// 对标 Minitab Response Optimizer
        /// </summary>
        public async Task<DesirabilityResult> OptimizeWithOlsAsync(
            string batchId,
            List<DesirabilityResponseConfig> configs,
            List<DOEFactor> factors,
            IDOEAnalysisService analysisService)
        {
            EnsurePythonReady();

            _currentConfigs = configs;
            _currentFactors = factors;

            try
            {
                using (Py.GIL())
                {
                    // 1. 创建 OLS 多响应预测器
                    dynamic doeDesirability = Py.Import("doe_desirability");
                    _olsPredictor = doeDesirability.create_ols_predictor();

                    // 2. 为每个响应变量创建独立的 DOEAnalyzer 并拟合 OLS 模型
                    dynamic doeEngine = Py.Import("doe_engine");

                    var batch = await _repository.GetBatchWithDetailsAsync(batchId);
                    if (batch == null)
                        return new DesirabilityResult { Success = false, ErrorMessage = "未找到批次" };

                    var completedRuns = batch.Runs
                        .Where(r => r.Status == DOERunStatus.Completed)
                        .OrderBy(r => r.RunIndex)
                        .ToList();

                    if (completedRuns.Count == 0)
                        return new DesirabilityResult { Success = false, ErrorMessage = "无已完成实验数据" };

                    // 解析因子值
                    var factorsList = completedRuns
                        .Select(r => JsonConvert.DeserializeObject<Dictionary<string, object>>(r.FactorValuesJson))
                        .Where(f => f != null)
                        .ToList();

                    var factorsJsonArray = JsonConvert.SerializeObject(factorsList);

                    // 因子类型 JSON
                    var factorTypes = new Dictionary<string, string>();
                    foreach (var f in factors)
                        factorTypes[f.FactorName] = f.IsCategorical ? "categorical" : "continuous";
                    var factorTypesJson = JsonConvert.SerializeObject(factorTypes);

                    // 为每个响应变量创建独立的 analyzer
                    foreach (var cfg in configs)
                    {
                        var responseName = cfg.ResponseName;

                        // 提取该响应的值
                        var responseValues = completedRuns
                            .Select(r =>
                            {
                                var respDict = JsonConvert.DeserializeObject<Dictionary<string, double>>(
                                    r.ResponseValuesJson ?? "{}");
                                return respDict?.GetValueOrDefault(responseName, 0.0) ?? 0.0;
                            })
                            .ToList();

                        var responsesJson = JsonConvert.SerializeObject(responseValues);

                        // 创建独立的 analyzer 并拟合
                        dynamic analyzer = doeEngine.create_analyzer();
                        analyzer.load_data(factorsJsonArray, responsesJson, responseName, factorTypesJson);
                        analyzer.fit_ols("quadratic");

                        // 注册到预测器
                        _olsPredictor.add_response(responseName, analyzer);

                        _logger.LogInformation("OLS Desirability: 响应 '{Name}' 模型已拟合", responseName);
                    }

                    // 3. 配置 DesirabilityEngine
                    var configsList = configs.Select(c => new
                    {
                        name = c.ResponseName,
                        goal = c.Goal.ToString().ToLower(),
                        lower = c.Lower,
                        upper = c.Upper,
                        target = c.Target,
                        weight = c.Weight,
                        importance = c.Importance,
                        shape = c.Shape,
                        shape_lower = c.ShapeLower,
                        shape_upper = c.ShapeUpper
                    }).ToList();

                    var factorNames = factors.Select(f => f.FactorName).ToList();
                    var boundsDict = new Dictionary<string, object>();
                    foreach (var f in factors)
                    {
                        if (f.IsCategorical)
                            boundsDict[f.FactorName] = f.CategoryLevels?.Split(',').Select(s => s.Trim()).ToList()
                                                       ?? new List<string>();
                        else
                            boundsDict[f.FactorName] = new[] { f.LowerBound, f.UpperBound };
                    }

                    _pythonEngine!.configure(
                        JsonConvert.SerializeObject(configsList),
                        JsonConvert.SerializeObject(factorNames),
                        JsonConvert.SerializeObject(boundsDict)
                    );

                    // 4. 用 OLS 预测器的 predict_json 方法做优化
                    // ★ 关键: 这里传 _olsPredictor.predict_json 而非 GPR 的预测
                    string resultJson = _pythonEngine!.optimize(_olsPredictor.predict_json).ToString();

                    var pyResult = JsonConvert.DeserializeObject<PyOptimizeResult>(resultJson)!;

                    var result = new DesirabilityResult
                    {
                        OptimalFactors = pyResult.OptimalFactors ?? new(),
                        CompositeD = pyResult.CompositeD,
                        Success = pyResult.Success,
                        ErrorMessage = pyResult.Error,
                        IndividualD = new()
                    };

                    if (pyResult.IndividualD != null)
                    {
                        foreach (var kv in pyResult.IndividualD)
                        {
                            var config = configs.FirstOrDefault(c => c.ResponseName == kv.Key);
                            result.IndividualD.Add(new IndividualDesirability
                            {
                                ResponseName = kv.Key,
                                PredictedValue = kv.Value.Value,
                                D = kv.Value.D,
                                Goal = config?.Goal ?? DesirabilityGoal.Maximize
                            });
                        }
                    }

                    _logger.LogInformation("OLS Desirability 优化完成: D={D}, 成功={Success}",
                        result.CompositeD, result.Success);
                    return result;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OLS Desirability 优化失败");
                return new DesirabilityResult
                {
                    Success = false,
                    ErrorMessage = $"OLS 优化失败: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// ★ v2 新增: 获取 OLS 模式的 Profile 图数据
        /// </summary>
        public Task<string> GetOlsProfileDataAsync(int gridSize = 50)
        {
            EnsurePythonReady();

            if (_olsPredictor == null)
                return Task.FromResult(JsonConvert.SerializeObject(new { error = "OLS 预测器未初始化，请先运行优化" }));

            using (Py.GIL())
            {
                string resultJson = _pythonEngine!.profile_plot_data(_olsPredictor.predict_json, gridSize).ToString();
                return Task.FromResult(resultJson);
            }
        }
        // ── Python 结果反序列化类 ──

        private class PyDesirabilityResult
        {
            [JsonProperty("composite_d")] public double CompositeD { get; set; }
            [JsonProperty("individual_d")] public Dictionary<string, PyIndividualD>? IndividualD { get; set; }
        }

        private class PyIndividualD
        {
            [JsonProperty("value")] public double Value { get; set; }
            [JsonProperty("d")] public double D { get; set; }
            [JsonProperty("goal")] public string? Goal { get; set; }
        }

        private class PyOptimizeResult
        {
            /// <summary>
            /// ★ 修复 (Bug#1): 从 Dictionary&lt;string, double&gt; 改为 Dictionary&lt;string, object&gt;
            /// Python 端 optimize() 返回类别因子为字符串值，double 反序列化会失败
            /// </summary>
            [JsonProperty("optimal_factors")] public Dictionary<string, object>? OptimalFactors { get; set; }
            [JsonProperty("composite_d")] public double CompositeD { get; set; }
            [JsonProperty("individual_d")] public Dictionary<string, PyIndividualD>? IndividualD { get; set; }
            [JsonProperty("success")] public bool Success { get; set; }
            [JsonProperty("error")] public string? Error { get; set; }
        }
    }
}