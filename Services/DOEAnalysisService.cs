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
    /// DOE 数据分析服务
    /// 
    ///  修改: 新增 FitOlsAsync / GetResidualDiagnosticsAsync / GetEffectsParetoAsync
    ///   这三个方法调用 Python DOEAnalyzer 的新方法: fit_ols() / residual_diagnostics() / effects_pareto()
    /// </summary>
    public class DOEAnalysisService : IDOEAnalysisService
    {
        private readonly IDOERepository _repository;
        private readonly IGPRModelService _gprService;
        private readonly ILogService _logger;
        private dynamic? _analyzer;
        private dynamic? _plotter;
        private bool _isPythonInitialized;

        public DOEAnalysisService(IDOERepository repository, IGPRModelService gprService, ILogService logger)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _gprService = gprService ?? throw new ArgumentNullException(nameof(gprService));
            _logger = logger?.ForContext<DOEAnalysisService>() ?? throw new ArgumentNullException(nameof(logger));
        }

        private void EnsureReady()
        {
            if (_isPythonInitialized && _analyzer != null) return;
            try
            {
                var envManager = MaxChemical.Modules.Designer.Services.PythonEnvironmentManager.Instance;
                if (!envManager.IsInitialized && !envManager.Initialize())
                    throw new InvalidOperationException("Python 环境初始化失败");
                using (Py.GIL())
                {
                    dynamic doeEngine = Py.Import("doe_engine");
                    _analyzer = doeEngine.create_analyzer();
                    _plotter = doeEngine.create_surface_plotter();
                }
                _isPythonInitialized = true;
                _logger.LogInformation("DOE 分析服务初始化成功");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "DOE 分析服务初始化失败");
                throw;
            }
        }

        /// <summary>
        /// 加载批次数据到 Python 分析器
        ///  修改: 支持指定 responseName（多响应时逐个分析）
        /// </summary>
        private async Task LoadBatchDataAsync(string batchId, string? responseName = null)
        {
            var batch = await _repository.GetBatchWithDetailsAsync(batchId);
            if (batch == null) throw new InvalidOperationException($"未找到批次: {batchId}");

            var completedRuns = batch.Runs
                .Where(r => r.Status == DOERunStatus.Completed)
                .OrderBy(r => r.RunIndex)
                .ToList();

            if (completedRuns.Count == 0)
                throw new InvalidOperationException("批次没有已完成的实验数据");

            //  修改: 支持指定响应变量名
            var primaryResponse = responseName ?? batch.Responses.FirstOrDefault()?.ResponseName ?? "Y";

            var factorsData = completedRuns
                .Select(r => JsonConvert.DeserializeObject<Dictionary<string, object>>(r.FactorValuesJson ?? "{}") ?? new())
                .ToList();

            var responsesData = completedRuns
                .Select(r =>
                {
                    var dict = JsonConvert.DeserializeObject<Dictionary<string, double>>(r.ResponseValuesJson ?? "{}") ?? new();
                    return dict.TryGetValue(primaryResponse, out var v) ? v : 0.0;
                })
                .ToList();

            var factorsJson = JsonConvert.SerializeObject(factorsData);
            var responsesJson = JsonConvert.SerializeObject(responsesData);

            // ★ 新增: 构建因子类型字典，传给 Python 区分连续/类别因子
            var factorTypesDict = new Dictionary<string, string>();
            if (batch.Factors != null)
            {
                foreach (var f in batch.Factors)
                {
                    factorTypesDict[f.FactorName] = f.IsCategorical ? "categorical" : "continuous";
                }
            }
            var factorTypesJson = JsonConvert.SerializeObject(factorTypesDict);

            using (Py.GIL())
            {
                _analyzer!.load_data(factorsJson, responsesJson, primaryResponse, factorTypesJson);
            }
        }

        // ═══════════════════════════════════════════════════════
        //  新增: 完整 OLS 回归分析
        // ═══════════════════════════════════════════════════════

        public async Task<OLSAnalysisResult> FitOlsAsync(string batchId, string responseName, string modelType = "quadratic")
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);

            _logger.LogInformation(" FitOLS: batchId={BatchId}, response={Response}, model={Model}",
                batchId, responseName, modelType);

            using (Py.GIL())
            {
                string resultJson = _analyzer!.fit_ols(modelType).ToString();
                var pyResult = JsonConvert.DeserializeObject<PythonOLSResult>(resultJson);

                if (pyResult == null || pyResult.Error != null)
                {
                    _logger.LogError("FitOLS 失败: {Error}", pyResult?.Error ?? "未知错误");
                    return new OLSAnalysisResult();
                }

                var result = new OLSAnalysisResult();

                // 解析 ANOVA 表
                if (pyResult.AnovaTable != null)
                {
                    result.AnovaTable = pyResult.AnovaTable.Select(row => new AnovaRow
                    {
                        Source = row.Source ?? "",
                        DF = row.DF,
                        SS = row.SS,
                        MS = row.MS,
                        FValue = row.FValue,
                        PValue = row.PValue
                    }).ToList();
                }

                // 解析系数表
                if (pyResult.Coefficients != null)
                {
                    result.Coefficients = pyResult.Coefficients.Select(c => new CoefficientRow
                    {
                        Term = c.Term ?? "",
                        Coefficient = c.Coeff,
                        StdError = c.SE,
                        TValue = c.TValue,
                        PValue = c.PValue,
                        VIF = c.VIF
                    }).ToList();
                }

                // 解析模型摘要
                if (pyResult.ModelSummary != null)
                {
                    result.ModelSummary = new OLSModelSummary
                    {
                        RSquared = pyResult.ModelSummary.RSquared,
                        RSquaredAdj = pyResult.ModelSummary.RSquaredAdj,
                        RSquaredPred = pyResult.ModelSummary.RSquaredPred,
                        RMSE = pyResult.ModelSummary.RMSE,
                        AdequatePrecision = pyResult.ModelSummary.AdequatePrecision,
                        PRESS = pyResult.ModelSummary.PRESS,
                        LackOfFitP = pyResult.ModelSummary.LackOfFitP,
                        ModelP = pyResult.ModelSummary.ModelP,
                        Equation = pyResult.ModelSummary.Equation ?? "",
                        Equations = pyResult.ModelSummary.Equations
                    };
                }

                _logger.LogInformation(" FitOLS 完成: R²={R2}, R²adj={R2adj}, R²pred={R2pred}",
                    result.ModelSummary.RSquared, result.ModelSummary.RSquaredAdj, result.ModelSummary.RSquaredPred);

                return result;
            }
        }
        public async Task<OLSAnalysisResult> FitOlsCustomAsync(string batchId, string responseName, List<string> terms)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);

            var termsJson = JsonConvert.SerializeObject(terms);
            _logger.LogInformation("★ FitOLS Custom: batchId={BatchId}, response={Response}, terms={Terms}",
                batchId, responseName, termsJson);

            using (Py.GIL())
            {
                // 调用 Python: fit_ols("custom", terms_json)
                string resultJson = _analyzer!.fit_ols("custom", termsJson).ToString();
                var pyResult = JsonConvert.DeserializeObject<PythonOLSResult>(resultJson);

                if (pyResult == null || pyResult.Error != null)
                {
                    _logger.LogError("FitOLS Custom 失败: {Error}", pyResult?.Error ?? "未知错误");
                    return new OLSAnalysisResult();
                }

                var result = new OLSAnalysisResult();

                if (pyResult.AnovaTable != null)
                {
                    result.AnovaTable = pyResult.AnovaTable.Select(row => new AnovaRow
                    {
                        Source = row.Source ?? "",
                        DF = row.DF,
                        SS = row.SS,
                        MS = row.MS,
                        FValue = row.FValue,
                        PValue = row.PValue
                    }).ToList();
                }

                if (pyResult.Coefficients != null)
                {
                    result.Coefficients = pyResult.Coefficients.Select(c => new CoefficientRow
                    {
                        Term = c.Term ?? "",
                        Coefficient = c.Coeff,
                        StdError = c.SE,
                        TValue = c.TValue,
                        PValue = c.PValue,
                        VIF = c.VIF
                    }).ToList();
                }

                if (pyResult.ModelSummary != null)
                {
                    result.ModelSummary = new OLSModelSummary
                    {
                        RSquared = pyResult.ModelSummary.RSquared,
                        RSquaredAdj = pyResult.ModelSummary.RSquaredAdj,
                        RSquaredPred = pyResult.ModelSummary.RSquaredPred,
                        RMSE = pyResult.ModelSummary.RMSE,
                        AdequatePrecision = pyResult.ModelSummary.AdequatePrecision,
                        PRESS = pyResult.ModelSummary.PRESS,
                        LackOfFitP = pyResult.ModelSummary.LackOfFitP,
                        ModelP = pyResult.ModelSummary.ModelP,
                        Equation = pyResult.ModelSummary.Equation ?? "",
                        Equations = pyResult.ModelSummary.Equations
                    };
                }

                _logger.LogInformation("★ FitOLS Custom 完成: R²={R2}, R²adj={R2adj}, 项数={Terms}",
                    result.ModelSummary.RSquared, result.ModelSummary.RSquaredAdj, result.Coefficients.Count);

                return result;
            }
        }
        // ═══════════════════════════════════════════════════════
        //  新增: 残差诊断四图
        // ═══════════════════════════════════════════════════════

        public async Task<string> GetResidualDiagnosticsAsync(string batchId, string responseName)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);

            _logger.LogInformation(" 残差诊断: batchId={BatchId}, response={Response}", batchId, responseName);

            using (Py.GIL())
            {
                string resultJson = _analyzer!.residual_diagnostics().ToString();
                return resultJson;
            }
        }

        // ═══════════════════════════════════════════════════════
        //  新增: 效应 Pareto 图（基于 T 值）
        // ═══════════════════════════════════════════════════════

        public async Task<string> GetEffectsParetoAsync(string batchId, string responseName, double alpha = 0.05)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);

            _logger.LogInformation(" 效应 Pareto: batchId={BatchId}, response={Response}, alpha={Alpha}",
                batchId, responseName, alpha);

            using (Py.GIL())
            {
                string resultJson = _analyzer!.effects_pareto(alpha).ToString();
                return resultJson;
            }
        }

        // ═══════════════════════════════════════════════════════
        // ★ 新增 v5: 预测刻画器 + 最优化
        // ═══════════════════════════════════════════════════════

        public async Task<string> GetPredictionProfilerAsync(string batchId, string responseName,
            int gridSize = 50, Dictionary<string, object>? fixedValues = null)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);

            // 确保模型已拟合
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");

                string fixedJson = fixedValues != null
                    ? JsonConvert.SerializeObject(fixedValues)
                    : "";

                string resultJson = _analyzer!.prediction_profiler(gridSize, fixedJson).ToString();
                return resultJson;
            }
        }

        public async Task<string> FindOptimalAsync(string batchId, string responseName, bool maximize = true)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);

            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                string resultJson = _analyzer!.find_optimal(maximize).ToString();
                return resultJson;
            }
        }

        // ═══════════════════════════════════════════════════════
        // 原有方法（保留不变）
        // ═══════════════════════════════════════════════════════

        public async Task<string> GetMainEffectsAsync(string batchId)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.main_effects().ToString(); }
        }

        public async Task<string> GetInteractionEffectsAsync(string batchId)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.interaction_effects().ToString(); }
        }

        public async Task<string> GetParetoChartAsync(string batchId)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.pareto_chart().ToString(); }
        }

        public async Task<string> GetAnovaTableAsync(string batchId)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.anova_table().ToString(); }
        }

        public async Task<string> GetRegressionSummaryAsync(string batchId)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.regression_summary().ToString(); }
        }

        public async Task<string> GetResidualAnalysisAsync(string batchId)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.residual_analysis().ToString(); }
        }

        public async Task<string> GetActualVsPredictedAsync(string batchId)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.actual_vs_predicted().ToString(); }
        }

        public async Task<string> GetResponseSurfaceDataAsync(string batchId, string factor1, string factor2)
        {
            EnsureReady();
            await SetupPlotterForBatch(batchId);
            using (Py.GIL()) { return _plotter!.response_surface_data(factor1, factor2).ToString(); }
        }

        public async Task<string> GetContourDataAsync(string batchId, string factor1, string factor2)
        {
            EnsureReady();
            await SetupPlotterForBatch(batchId);
            using (Py.GIL()) { return _plotter!.contour_data(factor1, factor2).ToString(); }
        }

        public async Task<byte[]> GetResponseSurfaceImageAsync(string batchId, string factor1, string factor2)
        {
            EnsureReady();
            await SetupPlotterForBatch(batchId);
            using (Py.GIL())
            {
                string base64 = _plotter!.response_surface_image(factor1, factor2).ToString();
                return string.IsNullOrEmpty(base64) ? Array.Empty<byte>() : Convert.FromBase64String(base64);
            }
        }

        public Task<byte[]> GetResponseSurfaceImageFromGPRAsync(string factor1, string factor2)
        {
            EnsureReady();
            var gprModel = _gprService.GetPythonModel();
            if (gprModel == null) return Task.FromResult(Array.Empty<byte>());

            using (Py.GIL())
            {
                string factorNamesJson = gprModel.get_factor_names().ToString();
                string boundsJson = gprModel.get_bounds().ToString();
                _plotter!.set_gpr_model(gprModel, factorNamesJson, boundsJson);
                string base64 = _plotter!.response_surface_image(factor1, factor2).ToString();
                return Task.FromResult(string.IsNullOrEmpty(base64) ? Array.Empty<byte>() : Convert.FromBase64String(base64));
            }
        }

        public Task<string> GetGPRConfidenceBandAsync(string flowId)
        {
            EnsureReady();
            var gprModel = _gprService.GetPythonModel();
            if (gprModel == null) return Task.FromResult("{}");
            using (Py.GIL())
            {
                string factorNamesJson = gprModel.get_factor_names().ToString();
                string boundsJson = gprModel.get_bounds().ToString();
                _plotter!.set_gpr_model(gprModel, factorNamesJson, boundsJson);
                var factorNames = JsonConvert.DeserializeObject<List<string>>(factorNamesJson) ?? new();
                if (factorNames.Count == 0) return Task.FromResult("{}");
                return Task.FromResult(_plotter!.gpr_confidence_band_data(factorNames[0]).ToString());
            }
        }

        public Task<string> GetAcquisitionSurfaceAsync(string flowId, string factor1, string factor2)
        {
            EnsureReady();
            var gprModel = _gprService.GetPythonModel();
            if (gprModel == null) return Task.FromResult("{}");
            using (Py.GIL())
            {
                string factorNamesJson = gprModel.get_factor_names().ToString();
                string boundsJson = gprModel.get_bounds().ToString();
                _plotter!.set_gpr_model(gprModel, factorNamesJson, boundsJson);
                return Task.FromResult(_plotter!.response_surface_data(factor1, factor2).ToString());
            }
        }

        public Task<string> GetModelEvolutionAsync(string flowId)
        {
            var gprModel = _gprService.GetPythonModel();
            if (gprModel == null) return Task.FromResult("[]");
            using (Py.GIL()) { return Task.FromResult(gprModel.get_evolution_data().ToString()); }
        }

        // ── 内部辅助 ──

        private async Task SetupPlotterForBatch(string batchId)
        {
            await LoadBatchDataAsync(batchId);
            var batch = await _repository.GetBatchWithDetailsAsync(batchId);
            if (batch == null) return;

            // ★ 修复 (Bug#3): 区分连续因子 [lower, upper] 和类别因子 ["A", "B", "C"]
            // 原来的 bug: 所有因子都用 [LowerBound, UpperBound]，类别因子的 0/0 会导致绘图范围错误
            var boundsDict = new Dictionary<string, object>();
            foreach (var f in batch.Factors)
            {
                if (f.IsCategorical)
                    boundsDict[f.FactorName] = f.GetCategoryLevelList();
                else
                    boundsDict[f.FactorName] = new[] { f.LowerBound, f.UpperBound };
            }
            var boundsJson = JsonConvert.SerializeObject(boundsDict);

            var gprModel = _gprService.GetPythonModel();
            using (Py.GIL())
            {
                if (gprModel != null)
                {
                    string factorNamesJson = gprModel.get_factor_names().ToString();
                    _plotter!.set_gpr_model(gprModel, factorNamesJson, boundsJson);
                }
                else
                {
                    _plotter!.set_analyzer(_analyzer, boundsJson);
                }
            }
        }

        // ──  新增: Python 结果反序列化类 ──

        private class PythonOLSResult
        {
            [JsonProperty("error")] public string? Error { get; set; }
            [JsonProperty("anova_table")] public List<PyAnovaRow>? AnovaTable { get; set; }
            [JsonProperty("coefficients")] public List<PyCoeffRow>? Coefficients { get; set; }
            [JsonProperty("model_summary")] public PyModelSummary? ModelSummary { get; set; }
        }

        private class PyAnovaRow
        {
            [JsonProperty("source")] public string? Source { get; set; }
            [JsonProperty("df")] public int DF { get; set; }
            [JsonProperty("ss")] public double SS { get; set; }
            [JsonProperty("ms")] public double? MS { get; set; }
            [JsonProperty("f_value")] public double? FValue { get; set; }
            [JsonProperty("p_value")] public double? PValue { get; set; }
        }

        private class PyCoeffRow
        {
            [JsonProperty("term")] public string? Term { get; set; }
            [JsonProperty("coeff")] public double Coeff { get; set; }
            [JsonProperty("se")] public double SE { get; set; }
            [JsonProperty("t_value")] public double TValue { get; set; }
            [JsonProperty("p_value")] public double PValue { get; set; }
            [JsonProperty("vif")] public double? VIF { get; set; }
        }

        private class PyModelSummary
        {
            [JsonProperty("r_squared")] public double RSquared { get; set; }
            [JsonProperty("r_squared_adj")] public double RSquaredAdj { get; set; }
            [JsonProperty("r_squared_pred")] public double RSquaredPred { get; set; }
            [JsonProperty("rmse")] public double RMSE { get; set; }
            [JsonProperty("adeq_precision")] public double AdequatePrecision { get; set; }
            [JsonProperty("press")] public double PRESS { get; set; }
            [JsonProperty("lack_of_fit_p")] public double? LackOfFitP { get; set; }
            [JsonProperty("model_p")] public double? ModelP { get; set; }
            [JsonProperty("equation")] public string? Equation { get; set; }

            [JsonProperty("equations")] public EquationsInfo? Equations { get; set; }
        }
    }
}