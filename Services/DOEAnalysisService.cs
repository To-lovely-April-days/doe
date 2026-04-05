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
    /// ★ v13: 所有 OLS 解析统一使用 ParseOLSResult，支持 uncoded 系数/方程
    /// </summary>
    public class DOEAnalysisService : IDOEAnalysisService
    {
        private readonly IDOERepository _repository;
        private readonly IGPRModelService _gprService;
        private readonly ILogService _logger;
        private dynamic? _analyzer;
        private dynamic? _plotter;
        private bool _isPythonInitialized;
        private string _loadedBatchId = "";
        private string _loadedResponseName = "";

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

            var factorTypesDict = new Dictionary<string, string>();
            if (batch.Factors != null)
            {
                foreach (var f in batch.Factors)
                    factorTypesDict[f.FactorName] = f.IsCategorical ? "categorical" : "continuous";
            }
            var factorTypesJson = JsonConvert.SerializeObject(factorTypesDict);

            using (Py.GIL())
            {
                _analyzer!.load_data(factorsJson, responsesJson, primaryResponse, factorTypesJson);
            }
        }

        private async Task<string> BuildBoundsJsonAsync(string batchId)
        {
            var batch = await _repository.GetBatchWithDetailsAsync(batchId);
            if (batch == null) return "{}";

            var boundsDict = new Dictionary<string, object>();
            foreach (var f in batch.Factors)
            {
                if (f.IsCategorical)
                    boundsDict[f.FactorName] = f.GetCategoryLevelList();
                else
                    boundsDict[f.FactorName] = new[] { f.LowerBound, f.UpperBound };
            }
            return JsonConvert.SerializeObject(boundsDict);
        }

        // ═══════════════════════════════════════════════════════
        // ★ v13: 统一 OLS 结果解析方法
        // ═══════════════════════════════════════════════════════

        /// <summary>
        /// 从 Python OLS 结果统一解析为 C# 对象
        /// 包含编码值系数、未编码系数、编码/未编码方程、编码信息
        /// </summary>
        private OLSAnalysisResult ParseOLSResult(PythonOLSResult pyResult, string defaultModelType = "quadratic")
        {
            var result = new OLSAnalysisResult();

            // ANOVA 表
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

            // 编码值系数（含 SE/T/P/VIF）
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

            // ★ v13: 未编码系数（只有 term + coeff）
            if (pyResult.UncodedCoefficients != null)
            {
                result.UncodedCoefficients = pyResult.UncodedCoefficients.Select(c => new UncCodedCoefficientRow
                {
                    Term = c.Term ?? "",
                    Coefficient = c.Coeff,
                    StdError = c.SE,
                    TValue = c.TValue,
                    PValue = c.PValue,
                    VIF = c.VIF
                }).ToList();
            }

            // 模型摘要
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

            // ★ v13: 未编码方程 + 编码信息
            result.UncodedEquation = pyResult.UncodedEquation ?? "";
            result.UncodedEquations = pyResult.UncodedEquations;
            result.CodingInfo = pyResult.CodingInfo ?? new();

            // v7: 不可估计项信息
            result.DroppedTerms = pyResult.DroppedTerms ?? new();
            result.InestimableWarning = pyResult.InestimableWarning ?? "";
            result.OriginalModelType = pyResult.OriginalModelType ?? defaultModelType;

            result.XtxInvFlat = pyResult.XtxInvFlat;
            result.Sigma = pyResult.Sigma;
            result.TCrit = pyResult.TCrit;
            result.ParamCount = pyResult.ParamCount;


            return result;
        }

        // ═══════════════════════════════════════════════════════
        // OLS 回归分析
        // ═══════════════════════════════════════════════════════

        public async Task<OLSAnalysisResult> FitOlsAsync(string batchId, string responseName, string modelType = "quadratic")
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);

            _logger.LogInformation("FitOLS: batchId={BatchId}, response={Response}, model={Model}",
                batchId, responseName, modelType);

            using (Py.GIL())
            {
                string resultJson = _analyzer!.fit_ols(modelType).ToString();
                var pyResult = JsonConvert.DeserializeObject<PythonOLSResult>(resultJson);

                if (pyResult == null || pyResult.Error != null)
                {
                    _logger.LogError("FitOLS 失败: {Error}", pyResult?.Error ?? "未知错误");
                    return new OLSAnalysisResult
                    {
                        DroppedTerms = pyResult?.DroppedTerms ?? new(),
                        InestimableWarning = pyResult?.InestimableWarning ?? ""
                    };
                }

                var result = ParseOLSResult(pyResult, modelType);

                if (result.DroppedTerms.Count > 0)
                    _logger.LogWarning("FitOLS 自动剔除不可估计项: {Terms}", string.Join(", ", result.DroppedTerms));

                _logger.LogInformation("FitOLS 完成: R²={R2}, R²adj={R2adj}, Dropped={Dropped}",
                    result.ModelSummary.RSquared, result.ModelSummary.RSquaredAdj, result.DroppedTerms.Count);

                return result;
            }
        }

        public async Task<OLSAnalysisResult> FitOlsCustomAsync(string batchId, string responseName, List<string> terms)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);

            var termsJson = JsonConvert.SerializeObject(terms);
            _logger.LogInformation("FitOLS Custom: batchId={BatchId}, response={Response}, terms={Terms}",
                batchId, responseName, termsJson);

            using (Py.GIL())
            {
                string resultJson = _analyzer!.fit_ols("custom", termsJson).ToString();
                var pyResult = JsonConvert.DeserializeObject<PythonOLSResult>(resultJson);

                if (pyResult == null || pyResult.Error != null)
                {
                    _logger.LogError("FitOLS Custom 失败: {Error}", pyResult?.Error ?? "未知错误");
                    return new OLSAnalysisResult
                    {
                        DroppedTerms = pyResult?.DroppedTerms ?? new(),
                        InestimableWarning = pyResult?.InestimableWarning ?? ""
                    };
                }

                var result = ParseOLSResult(pyResult, "custom");

                _logger.LogInformation("FitOLS Custom 完成: R²={R2}, R²adj={R2adj}, 项数={Terms}, Dropped={Dropped}",
                    result.ModelSummary.RSquared, result.ModelSummary.RSquaredAdj,
                    result.Coefficients.Count, result.DroppedTerms.Count);

                return result;
            }
        }

        // ═══════════════════════════════════════════════════════
        // 残差诊断四图
        // ═══════════════════════════════════════════════════════

        public async Task<string> GetResidualDiagnosticsAsync(string batchId, string responseName)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            using (Py.GIL()) { return _analyzer!.residual_diagnostics().ToString(); }
        }

        // ═══════════════════════════════════════════════════════
        // 效应 Pareto 图
        // ═══════════════════════════════════════════════════════

        public async Task<string> GetEffectsParetoAsync(string batchId, string responseName, double alpha = 0.05)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            using (Py.GIL()) { return _analyzer!.effects_pareto(alpha).ToString(); }
        }

        // ═══════════════════════════════════════════════════════
        // 预测刻画器 + 最优化
        // ═══════════════════════════════════════════════════════

        public async Task<string> GetPredictionProfilerAsync(string batchId, string responseName,
            int gridSize = 50, Dictionary<string, object>? fixedValues = null)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                string fixedJson = fixedValues != null ? JsonConvert.SerializeObject(fixedValues) : "";
                return _analyzer!.prediction_profiler(gridSize, fixedJson).ToString();
            }
        }

        public async Task<string> FindOptimalAsync(string batchId, string responseName, bool maximize = true)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                return _analyzer!.find_optimal(maximize).ToString();
            }
        }

        // ═══════════════════════════════════════════════════════
        // OLS 响应曲面 + 等高线
        // ═══════════════════════════════════════════════════════

        public async Task<string> GetOlsResponseSurfaceDataAsync(string batchId, string responseName,
            string factor1, string factor2, int gridSize = 30)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            var boundsJson = await BuildBoundsJsonAsync(batchId);
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                return _analyzer!.response_surface_data_ols(factor1, factor2, gridSize, boundsJson).ToString();
            }
        }

        public async Task<byte[]> GetOlsResponseSurfaceImageAsync(string batchId, string responseName,
            string factor1, string factor2)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            var boundsJson = await BuildBoundsJsonAsync(batchId);
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                string base64 = _analyzer!.response_surface_image_ols(factor1, factor2, 50, 800, 600, boundsJson).ToString();
                return string.IsNullOrEmpty(base64) ? Array.Empty<byte>() : Convert.FromBase64String(base64);
            }
        }

        public async Task<string> GetOlsContourDataAsync(string batchId, string responseName,
            string factor1, string factor2, int gridSize = 30)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            var boundsJson = await BuildBoundsJsonAsync(batchId);
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                return _analyzer!.contour_data_ols(factor1, factor2, gridSize, boundsJson).ToString();
            }
        }

        // ═══════════════════════════════════════════════════════
        // 异常点分析 + 排除重拟合
        // ═══════════════════════════════════════════════════════

        public async Task<string> GetOutlierAnalysisAsync(string batchId, string responseName)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                return _analyzer!.outlier_analysis().ToString();
            }
        }

        public async Task<OLSAnalysisResult> RefitExcludingAsync(string batchId, string responseName,
            List<int> excludeIndices, string modelType = "quadratic")
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);

            var excludeJson = JsonConvert.SerializeObject(excludeIndices);
            _logger.LogInformation("排除重拟合: batch={BatchId}, exclude={Indices}", batchId, excludeJson);

            using (Py.GIL())
            {
                string resultJson = _analyzer!.refit_excluding(excludeJson, modelType).ToString();
                var pyResult = JsonConvert.DeserializeObject<PythonOLSResult>(resultJson);

                if (pyResult == null || pyResult.Error != null)
                {
                    _logger.LogError("排除重拟合失败: {Error}", pyResult?.Error ?? "未知错误");
                    return new OLSAnalysisResult();
                }

                var result = ParseOLSResult(pyResult, modelType);

                _logger.LogInformation("排除重拟合完成: R²={R2}, 排除={Excluded}组",
                    result.ModelSummary.RSquared, excludeIndices.Count);

                return result;
            }
        }

        // ═══════════════════════════════════════════════════════
        // Tukey HSD + 主效应 / 交互效应
        // ═══════════════════════════════════════════════════════

        public async Task<string> GetTukeyHSDAsync(string batchId, string responseName, string factorName = "")
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            using (Py.GIL()) { return _analyzer!.tukey_hsd(factorName).ToString(); }
        }

        public async Task<string> GetMainEffectsForResponseAsync(string batchId, string responseName)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            using (Py.GIL()) { return _analyzer!.main_effects().ToString(); }
        }

        public async Task<string> GetInteractionEffectsForResponseAsync(string batchId, string responseName)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            using (Py.GIL()) { return _analyzer!.interaction_effects().ToString(); }
        }

        // ═══════════════════════════════════════════════════════
        // Box-Cox + 报告导出
        // ═══════════════════════════════════════════════════════

        public async Task<string> GetBoxCoxAnalysisAsync(string batchId, string responseName)
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                return _analyzer!.box_cox_analysis().ToString();
            }
        }

        public async Task<OLSAnalysisResult> ApplyBoxCoxAsync(string batchId, string responseName,
            double lambdaValue, string modelType = "quadratic")
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            using (Py.GIL())
            {
                string resultJson = _analyzer!.apply_box_cox(lambdaValue, modelType).ToString();
                var pyResult = JsonConvert.DeserializeObject<PythonOLSResult>(resultJson);

                if (pyResult == null || pyResult.Error != null)
                    return new OLSAnalysisResult();

                return ParseOLSResult(pyResult, modelType);
            }
        }

        public async Task<string> ExportOlsReportAsync(string batchId, string responseName,
            string outputPath, string title = "OLS 回归分析报告")
        {
            EnsureReady();
            await LoadBatchDataAsync(batchId, responseName);
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                return _analyzer!.export_ols_report(outputPath, title).ToString();
            }
        }

        // ═══════════════════════════════════════════════════════
        // 原有方法（保留不变）
        // ═══════════════════════════════════════════════════════

        public async Task<string> GetMainEffectsAsync(string batchId)
        {
            EnsureReady(); await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.main_effects().ToString(); }
        }

        public async Task<string> GetInteractionEffectsAsync(string batchId)
        {
            EnsureReady(); await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.interaction_effects().ToString(); }
        }

        public async Task<string> GetParetoChartAsync(string batchId)
        {
            EnsureReady(); await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.pareto_chart().ToString(); }
        }

        public async Task<string> GetAnovaTableAsync(string batchId)
        {
            EnsureReady(); await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.anova_table().ToString(); }
        }

        public async Task<string> GetRegressionSummaryAsync(string batchId)
        {
            EnsureReady(); await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.regression_summary().ToString(); }
        }

        public async Task<string> GetResidualAnalysisAsync(string batchId)
        {
            EnsureReady(); await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.residual_analysis().ToString(); }
        }

        public async Task<string> GetActualVsPredictedAsync(string batchId)
        {
            EnsureReady(); await LoadBatchDataAsync(batchId);
            using (Py.GIL()) { return _analyzer!.actual_vs_predicted().ToString(); }
        }

        public async Task<string> GetResponseSurfaceDataAsync(string batchId, string factor1, string factor2)
        {
            EnsureReady(); await SetupPlotterForBatch(batchId);
            using (Py.GIL()) { return _plotter!.response_surface_data(factor1, factor2).ToString(); }
        }

        public async Task<string> GetContourDataAsync(string batchId, string factor1, string factor2)
        {
            EnsureReady(); await SetupPlotterForBatch(batchId);
            using (Py.GIL()) { return _plotter!.contour_data(factor1, factor2).ToString(); }
        }

        public async Task<byte[]> GetResponseSurfaceImageAsync(string batchId, string factor1, string factor2)
        {
            EnsureReady(); await SetupPlotterForBatch(batchId);
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
        /// <summary>
        /// ★ 新增: 直接数据导入 OLS 分析（不通过批次）
        /// 数据由调用方从 Excel 读取后传入，绕过 LoadBatchDataAsync
        /// </summary>
        public Task<OLSAnalysisResult> FitOlsDirectAsync(
            List<Dictionary<string, object>> factorsData,
            List<double> responsesData,
            string responseName,
            Dictionary<string, string>? factorTypes = null,
            string modelType = "quadratic")
        {
            EnsureReady();

            var factorsJson = JsonConvert.SerializeObject(factorsData);
            var responsesJson = JsonConvert.SerializeObject(responsesData);
            var factorTypesJson = factorTypes != null
                ? JsonConvert.SerializeObject(factorTypes)
                : "{}";

            _logger.LogInformation("FitOLS Direct: response={Response}, rows={Count}, model={Model}",
                responseName, factorsData.Count, modelType);

            using (Py.GIL())
            {
                // 直接加载数据（不走批次）
                _analyzer!.load_data(factorsJson, responsesJson, responseName, factorTypesJson);

                string resultJson = _analyzer!.fit_ols(modelType).ToString();
                var pyResult = JsonConvert.DeserializeObject<PythonOLSResult>(resultJson);

                if (pyResult == null || pyResult.Error != null)
                {
                    _logger.LogError("FitOLS Direct 失败: {Error}", pyResult?.Error ?? "未知错误");
                    return Task.FromResult(new OLSAnalysisResult
                    {
                        DroppedTerms = pyResult?.DroppedTerms ?? new(),
                        InestimableWarning = pyResult?.InestimableWarning ?? ""
                    });
                }

                var result = ParseOLSResult(pyResult, modelType);
                _logger.LogInformation("FitOLS Direct 完成: R²={R2}, 系数数={Count}",
                    result.ModelSummary?.RSquared, result.Coefficients?.Count);

                // 清除批次缓存标记，防止后续方法误用
                _loadedBatchId = "";
                _loadedResponseName = responseName;

                return Task.FromResult(result);
            }
        }
        // ═══════════════════════════════════════════════════════
        // ★ 新增: Direct 模式方法（analyzer 已有数据和模型）
        // ═══════════════════════════════════════════════════════

        public string GetEffectsParetoDirectAsync(string responseName, double alpha = 0.05)
        {
            EnsureReady();
            using (Py.GIL()) { return _analyzer!.effects_pareto(alpha).ToString(); }
        }

        public string GetResidualDiagnosticsDirectAsync(string responseName)
        {
            EnsureReady();
            using (Py.GIL()) { return _analyzer!.residual_diagnostics().ToString(); }
        }

        public string GetMainEffectsDirectAsync(string responseName)
        {
            EnsureReady();
            using (Py.GIL()) { return _analyzer!.main_effects().ToString(); }
        }

        public string GetInteractionEffectsDirectAsync(string responseName)
        {
            EnsureReady();
            using (Py.GIL()) { return _analyzer!.interaction_effects().ToString(); }
        }

        public string GetOutlierAnalysisDirectAsync(string responseName)
        {
            EnsureReady();
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                return _analyzer!.outlier_analysis().ToString();
            }
        }

        public string GetBoxCoxAnalysisDirectAsync(string responseName)
        {
            EnsureReady();
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                return _analyzer!.box_cox_analysis().ToString();
            }
        }

        public string GetOlsSurfaceDataDirectAsync(string factor1, string factor2, string boundsJson, int gridSize = 30)
        {
            EnsureReady();
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                return _analyzer!.response_surface_data_ols(factor1, factor2, gridSize, boundsJson).ToString();
            }
        }

        public byte[] GetOlsSurfaceImageDirectAsync(string factor1, string factor2, string boundsJson)
        {
            EnsureReady();
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                string base64 = _analyzer!.response_surface_image_ols(factor1, factor2, 50, 800, 600, boundsJson).ToString();
                return string.IsNullOrEmpty(base64) ? Array.Empty<byte>() : Convert.FromBase64String(base64);
            }
        }

        public string GetOlsContourDataDirectAsync(string factor1, string factor2, string boundsJson, int gridSize = 30)
        {
            EnsureReady();
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                return _analyzer!.contour_data_ols(factor1, factor2, gridSize, boundsJson).ToString();
            }
        }

        public string GetPredictionProfilerDirectAsync(string responseName, int gridSize = 50, string fixedValuesJson = "")
        {
            EnsureReady();
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                return _analyzer!.prediction_profiler(gridSize, fixedValuesJson).ToString();
            }
        }

        public string FindOptimalDirectAsync(string responseName, bool maximize = true)
        {
            EnsureReady();
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                return _analyzer!.find_optimal(maximize).ToString();
            }
        }
        public string GetTukeyHSDDirectAsync(string responseName, string categoricalFactorName)
        {
            EnsureReady();
            using (Py.GIL())
            {
                _analyzer!.fit_ols("quadratic");
                return _analyzer!.tukey_hsd(categoricalFactorName).ToString();
            }
        }
        // ═══════════════════════════════════════════════════════
        // Python 结果反序列化类
        // ═══════════════════════════════════════════════════════

        private class PythonOLSResult
        {
            [JsonProperty("error")] public string? Error { get; set; }
            [JsonProperty("anova_table")] public List<PyAnovaRow>? AnovaTable { get; set; }
            [JsonProperty("coefficients")] public List<PyCoeffRow>? Coefficients { get; set; }
            [JsonProperty("model_summary")] public PyModelSummary? ModelSummary { get; set; }

            // v7
            [JsonProperty("dropped_terms")] public List<string>? DroppedTerms { get; set; }
            [JsonProperty("inestimable_warning")] public string? InestimableWarning { get; set; }
            [JsonProperty("original_model_type")] public string? OriginalModelType { get; set; }

            // ★ v13 新增
            [JsonProperty("uncoded_coefficients")] public List<PyUncodedCoeffRow>? UncodedCoefficients { get; set; }
            [JsonProperty("uncoded_equation")] public string? UncodedEquation { get; set; }
            [JsonProperty("uncoded_equations")] public EquationsInfo? UncodedEquations { get; set; }
            [JsonProperty("coding_info")] public Dictionary<string, CodingInfoItem>? CodingInfo { get; set; }

            [JsonProperty("xtx_inv_flat")] public double[]? XtxInvFlat { get; set; }
            [JsonProperty("sigma")] public double Sigma { get; set; }
            [JsonProperty("t_crit")] public double TCrit { get; set; }
            [JsonProperty("param_count")] public int ParamCount { get; set; }
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

        private class PyUncodedCoeffRow
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