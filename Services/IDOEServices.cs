using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using MaxChemical.Modules.DOE.Models;

namespace MaxChemical.Modules.DOE.Services
{
    // ═══════════════════════════════════════════════════════
    // IDOEDesignService — 实验设计服务
    // ═══════════════════════════════════════════════════════

    public interface IDOEDesignService
    {
        Task<DOEDesignMatrix> GenerateFullFactorialAsync(List<DOEFactor> factors);
        Task<DOEDesignMatrix> GenerateFractionalFactorialAsync(List<DOEFactor> factors, int resolution);
        Task<DOEDesignMatrix> GenerateTaguchiAsync(List<DOEFactor> factors, string tableType);
        Task<DOEDesignMatrix> ImportCustomMatrixAsync(string excelPath, List<DOEFactor> factors);
        Task<DataValidationResult> ValidateImportedDataAsync(string excelPath, List<DOEFactor> factors, List<DOEResponse> responses);
        Task<List<DOERunRecord>> ImportHistoricalDataAsync(string excelPath, string batchId, List<DOEFactor> factors, List<DOEResponse> responses);
        Task<DOEDesignMatrix> GenerateCCDAsync(List<DOEFactor> factors, string alphaType = "rotatable", int centerCount = -1);
        Task<DOEDesignMatrix> GenerateBoxBehnkenAsync(List<DOEFactor> factors, int centerCount = -1);
        Task<DOEDesignMatrix> GenerateDOptimalAsync(List<DOEFactor> factors, int numRuns = -1, string modelType = "quadratic");
        Task<DOEDesignQuality> GetDesignQualityAsync(List<DOEFactor> factors, DOEDesignMatrix matrix, string modelType = "quadratic");
        Task<DOEDesignMatrix> GeneratePlackettBurmanAsync(List<DOEFactor> factors);
    }

    public class DataValidationResult
    {
        public bool IsValid { get; set; }
        public List<string> Errors { get; set; } = new();
        public List<string> Warnings { get; set; } = new();
        public int ValidRowCount { get; set; }
    }

    // ═══════════════════════════════════════════════════════
    // IDOEExecutionService — 批量执行服务（保留不变）
    // ═══════════════════════════════════════════════════════

    public interface IDOEExecutionService
    {
        DOEBatchStatus BatchState { get; }
        int CurrentRunIndex { get; }
        Task StartBatchAsync(string batchId, CancellationToken cancellationToken = default);
        void PauseBatch();
        void ResumeBatch();
        Task AbortBatchAsync();
        void SubmitRunResponse(int runIndex, Dictionary<string, double> responseValues);
        event EventHandler<DOERunCompletedEventArgs>? RunCompleted;
        event EventHandler<DOEBatchCompletedEventArgs>? BatchCompleted;
        event EventHandler<DOEProgressEventArgs>? ProgressChanged;
        event EventHandler? PauseConfirmed;
    }

    public class DOERunCompletedEventArgs : EventArgs
    {
        public int RunIndex { get; set; }
        public Dictionary<string, double> FactorValues { get; set; } = new();
        public Dictionary<string, double> ResponseValues { get; set; } = new();
        public bool IsSuccessful { get; set; }
    }

    public class DOEBatchCompletedEventArgs : EventArgs
    {
        public string BatchId { get; set; } = string.Empty;
        public DOEBatchStatus FinalStatus { get; set; }
        public int TotalRuns { get; set; }
        public int CompletedRuns { get; set; }
        public string? StopReason { get; set; }
    }

    public class DOEProgressEventArgs : EventArgs
    {
        public int CurrentRun { get; set; }
        public int TotalRuns { get; set; }
        public double ProgressPercent => TotalRuns > 0 ? (double)CurrentRun / TotalRuns * 100 : 0;
        public string StatusMessage { get; set; } = string.Empty;
    }

    // ═══════════════════════════════════════════════════════
    // IGPRModelService — GPR 预测模型服务（保留不变）
    // ═══════════════════════════════════════════════════════

    public interface IGPRModelService
    {
        Task InitializeModelAsync(string flowId, List<DOEFactor> factors);
        dynamic? GetPythonModel();
        Task<GPRTrainResult> TrainModelAsync();
        GPRPrediction Predict(Dictionary<string, object> factorValues);
        GPRPrediction Predict(Dictionary<string, double> factorValues);
        List<GPRPrediction> PredictBatch(List<Dictionary<string, object>> factorValuesList);
        List<GPRPrediction> PredictBatch(List<Dictionary<string, double>> factorValuesList);
        GPROptimalResult FindOptimal(bool maximize = true);
        GPRStopDecision ShouldStop(List<Dictionary<string, object>> remainingRuns, double bestObserved,
                                    bool maximize = true, double eiThreshold = 0.01, double minRunsRatio = 1.5);
        GPRStopDecision ShouldStop(List<Dictionary<string, double>> remainingRuns, double bestObserved,
                                    bool maximize = true, double eiThreshold = 0.01, double minRunsRatio = 1.5);
        Dictionary<string, double> GetSensitivity();
        Task SaveStateAsync(string flowId);
        Task<bool> LoadStateAsync(string flowId);
        Task ResetModelAsync(string flowId, bool keepData);
        bool IsActive { get; }
        int DataCount { get; }
        Task SaveInitialStateAsync(string flowId);
        void AppendData(Dictionary<string, object> factorValues, double responseValue,
                        string source = "measured", string batchName = "", string timestamp = "");
        void AppendData(Dictionary<string, double> factorValues, double responseValue,
                        string source = "measured", string batchName = "", string timestamp = "");
        /// <summary>设置当前关联的项目 ID</summary>
        void SetProjectId(string? projectId);

    }

    public class GPRTrainResult
    {
        public bool IsActive { get; set; }
        public double RSquared { get; set; }
        public double RMSE { get; set; }
        public Dictionary<string, double> Lengthscales { get; set; } = new();
        public int DataCount { get; set; }
        public string RSquaredType { get; set; } = "LOO-CV";
    }

    public class GPRPrediction
    {
        public double Mean { get; set; }
        public double Std { get; set; }
        public double LowerBound => Mean - 2 * Std;
        public double UpperBound => Mean + 2 * Std;
    }

    public class GPROptimalResult
    {
        public Dictionary<string, double> OptimalFactors { get; set; } = new();
        public double PredictedResponse { get; set; }
        public double PredictionStd { get; set; }
        public double Confidence
        {
            get
            {
                if (Math.Abs(PredictedResponse) < 1e-12)
                    return PredictionStd < 1e-12 ? 1.0 : 0.5;
                double cv = PredictionStd / Math.Max(Math.Abs(PredictedResponse), 1e-12);
                return 1.0 / (1.0 + cv);
            }
        }
        public string Direction { get; set; } = "maximize";
    }

    public class GPRStopDecision
    {
        public bool ShouldStop { get; set; }
        public string Reason { get; set; } = string.Empty;
        public double BestPredicted { get; set; }
        public double BestPredictedStd { get; set; }
        public double MaxEI { get; set; }
    }

    // ═══════════════════════════════════════════════════════
    // IDOEAnalysisService — 数据分析服务
    // ★ v7 新增: GetOlsResponseSurfaceDataAsync / GetOlsResponseSurfaceImageAsync / GetOlsContourDataAsync
    // ═══════════════════════════════════════════════════════

    public interface IDOEAnalysisService
    {
        // ── 原有方法 ──
        Task<string> GetMainEffectsAsync(string batchId);
        Task<string> GetInteractionEffectsAsync(string batchId);
        Task<string> GetParetoChartAsync(string batchId);
        Task<string> GetResponseSurfaceDataAsync(string batchId, string factor1, string factor2);
        Task<string> GetContourDataAsync(string batchId, string factor1, string factor2);
        Task<byte[]> GetResponseSurfaceImageAsync(string batchId, string factor1, string factor2);
        Task<byte[]> GetResponseSurfaceImageFromGPRAsync(string factor1, string factor2);
        Task<string> GetResidualAnalysisAsync(string batchId);
        Task<string> GetActualVsPredictedAsync(string batchId);
        Task<string> GetAnovaTableAsync(string batchId);
        Task<string> GetRegressionSummaryAsync(string batchId);
        Task<string> GetGPRConfidenceBandAsync(string flowId);
        Task<string> GetAcquisitionSurfaceAsync(string flowId, string factor1, string factor2);
        Task<string> GetModelEvolutionAsync(string flowId);

        // ── OLS 分析 ──
        Task<OLSAnalysisResult> FitOlsAsync(string batchId, string responseName, string modelType = "quadratic");
        Task<OLSAnalysisResult> FitOlsCustomAsync(string batchId, string responseName, List<string> terms);
        Task<string> GetResidualDiagnosticsAsync(string batchId, string responseName);
        Task<string> GetEffectsParetoAsync(string batchId, string responseName, double alpha = 0.05);
        Task<string> GetPredictionProfilerAsync(string batchId, string responseName,
            int gridSize = 50, Dictionary<string, object>? fixedValues = null);
        Task<string> FindOptimalAsync(string batchId, string responseName, bool maximize = true);

        // ── ★ v7 新增: OLS 响应曲面 + 等高线 ──

        /// <summary>
        /// ★ v7 新增: 获取 OLS 模型的响应曲面数据（3D 网格）
        /// 返回 JSON: { "x": [...], "y": [...], "z": [[...]], "x_label": "...", "y_label": "..." }
        /// </summary>
        /// <param name="batchId">批次 ID</param>
        /// <param name="responseName">响应变量名</param>
        /// <param name="factor1">X 轴因子名</param>
        /// <param name="factor2">Y 轴因子名</param>
        /// <param name="gridSize">网格密度（默认30）</param>
        Task<string> GetOlsResponseSurfaceDataAsync(string batchId, string responseName,
            string factor1, string factor2, int gridSize = 30);

        /// <summary>
        /// ★ v7 新增: 获取 OLS 模型的响应曲面 PNG 图片
        /// 返回 PNG 图片字节数组（由 Python matplotlib 生成）
        /// </summary>
        Task<byte[]> GetOlsResponseSurfaceImageAsync(string batchId, string responseName,
            string factor1, string factor2);

        /// <summary>
        /// ★ v7 新增: 获取 OLS 模型的等高线数据（与响应曲面格式相同）
        /// </summary>
        Task<string> GetOlsContourDataAsync(string batchId, string responseName,
            string factor1, string factor2, int gridSize = 30);
        // ── ★ v8 新增: 异常点分析 ──

        /// <summary>
        /// ★ v8 新增: 综合异常点分析
        /// 返回基于 Cook's D、杠杆值、标准化残差的异常点列表
        /// </summary>
        Task<string> GetOutlierAnalysisAsync(string batchId, string responseName);

        /// <summary>
        /// ★ v8 新增: 排除指定异常点后重新拟合 OLS 模型
        /// </summary>
        /// <param name="excludeIndices">要排除的观测序号列表 (1-based)</param>
        Task<OLSAnalysisResult> RefitExcludingAsync(string batchId, string responseName,
            List<int> excludeIndices, string modelType = "quadratic");

        // ── ★ v8 新增: Tukey HSD 多重比较 ──

        /// <summary>
        /// ★ v8 新增: Tukey HSD 事后多重比较检验
        /// 针对类别因子，检验各水平间均值差异的显著性
        /// </summary>
        /// <param name="factorName">类别因子名（空字符串=第一个类别因子）</param>
        Task<string> GetTukeyHSDAsync(string batchId, string responseName, string factorName = "");

        // ── ★ v8 新增: OLS 上下文的主效应图 + 交互效应图 ──
        // 注: 这些方法在 GPR 上下文中已有 (GetMainEffectsAsync/GetInteractionEffectsAsync)，
        // 但那些版本使用默认响应变量。v8 新增带 responseName 参数的版本用于 OLS Tab 多响应场景。

        /// <summary>
        /// ★ v8 新增: 获取主效应图数据（指定响应变量）
        /// </summary>
        Task<string> GetMainEffectsForResponseAsync(string batchId, string responseName);

        /// <summary>
        /// ★ v8 新增: 获取交互效应图数据（指定响应变量）
        /// </summary>
        Task<string> GetInteractionEffectsForResponseAsync(string batchId, string responseName);

        /// <summary>★ v9: Box-Cox 变换分析</summary>
        Task<string> GetBoxCoxAnalysisAsync(string batchId, string responseName);

        /// <summary>★ v9: 应用 Box-Cox 变换后重拟合</summary>
        Task<OLSAnalysisResult> ApplyBoxCoxAsync(string batchId, string responseName,
            double lambdaValue, string modelType = "quadratic");

        /// <summary>★ v9: 导出 OLS 报告为 Word 文件</summary>
        Task<string> ExportOlsReportAsync(string batchId, string responseName, string outputPath,
            string title = "OLS 回归分析报告");

        // ═══ 新增: 直接数据导入 OLS 分析（不依赖批次） ═══

        /// <summary>
        /// 直接用因子+响应数据做 OLS 分析，不依赖 DOE 批次
        /// </summary>
        /// <param name="factorsData">每组实验的因子值列表</param>
        /// <param name="responsesData">对应的响应值列表</param>
        /// <param name="responseName">响应变量名称</param>
        /// <param name="factorTypes">因子类型 (continuous/categorical)，可选</param>
        /// <param name="modelType">模型类型，默认 quadratic</param>
        Task<OLSAnalysisResult> FitOlsDirectAsync(
            List<Dictionary<string, object>> factorsData,
            List<double> responsesData,
            string responseName,
            Dictionary<string, string>? factorTypes = null,
            string modelType = "quadratic");

        /// <summary>直接模式下获取 Pareto（analyzer 已有数据和模型）</summary>
        string GetEffectsParetoDirectAsync(string responseName, double alpha = 0.05);

        /// <summary>直接模式下获取残差诊断（analyzer 已有数据和模型）</summary>
        string GetResidualDiagnosticsDirectAsync(string responseName);
     
        string GetMainEffectsDirectAsync(string responseName);
        string GetInteractionEffectsDirectAsync(string responseName);
        string GetOutlierAnalysisDirectAsync(string responseName);
        string GetBoxCoxAnalysisDirectAsync(string responseName);
        string GetOlsSurfaceDataDirectAsync(string factor1, string factor2, string boundsJson, int gridSize = 30);
        byte[] GetOlsSurfaceImageDirectAsync(string factor1, string factor2, string boundsJson);
        string GetOlsContourDataDirectAsync(string factor1, string factor2, string boundsJson, int gridSize = 30);
        string GetPredictionProfilerDirectAsync(string responseName, int gridSize = 50, string fixedValuesJson = "");
        string FindOptimalDirectAsync(string responseName, bool maximize = true);
        string GetTukeyHSDDirectAsync(string responseName, string categoricalFactorName);
    }

    // ═══════════════════════════════════════════════════════
    // IDesirabilityService — 多响应优化服务（保留不变）
    // ═══════════════════════════════════════════════════════

    public interface IDesirabilityService
    {
        Task ConfigureAsync(string batchId, List<DesirabilityResponseConfig> configs, List<DOEFactor> factors);
        Task<DesirabilityResult> ComputeDesirabilityAsync(Dictionary<string, double> predictions);
        Task<DesirabilityResult> OptimizeAsync();
        Task<string> GetProfileDataAsync(int gridSize = 50);
        Task SaveConfigAsync(string batchId, List<DesirabilityResponseConfig> configs);
        Task<List<DesirabilityResponseConfig>> LoadConfigAsync(string batchId);

        // ★ v2 新增: OLS 模式
        Task<DesirabilityResult> OptimizeWithOlsAsync(
            string batchId,
            List<DesirabilityResponseConfig> configs,
            List<DOEFactor> factors,
            IDOEAnalysisService analysisService);
        Task<string> GetOlsProfileDataAsync(int gridSize = 50);
    }

    // ═══════════════════════════════════════════════════════
    // IModelRouter（保留不变）
    // ═══════════════════════════════════════════════════════

    public enum AnalysisStrategy { OLS, GPR, Insufficient }

    public interface IModelRouter
    {
        AnalysisStrategy SelectStrategy(DOEDesignMethod method, int factorCount, int dataCount);
        string GetStrategyReason(AnalysisStrategy strategy, DOEDesignMethod method, int factorCount, int dataCount);
    }


}