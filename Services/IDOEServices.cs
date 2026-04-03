using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using MaxChemical.Modules.DOE.Models;

namespace MaxChemical.Modules.DOE.Services
{
    // ═══════════════════════════════════════════════════════
    // IDOEDesignService — 实验设计服务
    //  修改: 新增 CCD/BBD/DOptimal 设计方法 + 设计质量评估
    // ═══════════════════════════════════════════════════════

    /// <summary>
    /// DOE 实验设计服务 — 调用 Python doe_engine.pyd 生成参数矩阵
    /// </summary>
    public interface IDOEDesignService
    {
        // ── 原有方法（保留） ──

        Task<DOEDesignMatrix> GenerateFullFactorialAsync(List<DOEFactor> factors);
        Task<DOEDesignMatrix> GenerateFractionalFactorialAsync(List<DOEFactor> factors, int resolution);
        Task<DOEDesignMatrix> GenerateTaguchiAsync(List<DOEFactor> factors, string tableType);
        Task<DOEDesignMatrix> ImportCustomMatrixAsync(string excelPath, List<DOEFactor> factors);
        Task<DataValidationResult> ValidateImportedDataAsync(string excelPath, List<DOEFactor> factors, List<DOEResponse> responses);
        Task<List<DOERunRecord>> ImportHistoricalDataAsync(string excelPath, string batchId, List<DOEFactor> factors, List<DOEResponse> responses);

        // ──  新增: RSM 设计方法 ──

        /// <summary>
        ///  新增: 生成 CCD 中心复合设计矩阵
        /// </summary>
        /// <param name="factors">因子列表（需要 LowerBound/UpperBound）</param>
        /// <param name="alphaType">alpha 类型: "rotatable"(推荐) / "face" / "orthogonal"</param>
        /// <param name="centerCount">中心点数，-1 自动</param>
        Task<DOEDesignMatrix> GenerateCCDAsync(List<DOEFactor> factors, string alphaType = "rotatable", int centerCount = -1);

        /// <summary>
        ///  新增: 生成 Box-Behnken 设计矩阵
        /// </summary>
        /// <param name="factors">因子列表（至少3个）</param>
        /// <param name="centerCount">中心点数，-1 自动</param>
        Task<DOEDesignMatrix> GenerateBoxBehnkenAsync(List<DOEFactor> factors, int centerCount = -1);

        /// <summary>
        ///  新增: 生成 D-Optimal 设计矩阵
        /// </summary>
        /// <param name="factors">因子列表</param>
        /// <param name="numRuns">实验次数，-1 自动</param>
        /// <param name="modelType">模型类型: "linear"/"interaction"/"quadratic"</param>
        Task<DOEDesignMatrix> GenerateDOptimalAsync(List<DOEFactor> factors, int numRuns = -1, string modelType = "quadratic");

        // ──  新增: 设计质量评估 ──

        /// <summary>
        ///  新增: 评估设计矩阵的统计质量
        /// 返回 D/A/G 效率、VIF、条件数、Power 分析等
        /// </summary>
        Task<DOEDesignQuality> GetDesignQualityAsync(List<DOEFactor> factors, DOEDesignMatrix matrix, string modelType = "quadratic");
    }

    /// <summary>
    /// 数据校验结果
    /// </summary>
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

    /// <summary>
    /// DOE 批量执行服务 — 编排多组实验的执行循环
    /// </summary>
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

    /// <summary>
    /// GPR 模型服务 — 管理高斯过程回归模型的生命周期
    /// </summary>
    public interface IGPRModelService
    {
        Task InitializeModelAsync(string flowId, List<DOEFactor> factors);
        dynamic? GetPythonModel();
        Task<GPRTrainResult> TrainModelAsync();
        /// <summary>★ 修复 (v3): factorValues 改为 object 以支持类别因子标签</summary>
        GPRPrediction Predict(Dictionary<string, object> factorValues);
        /// <summary>便利重载: 纯连续因子场景</summary>
        GPRPrediction Predict(Dictionary<string, double> factorValues);
        /// <summary>★ 修复 (v3): factorValuesList 改为 object 以支持类别因子标签</summary>
        List<GPRPrediction> PredictBatch(List<Dictionary<string, object>> factorValuesList);
        /// <summary>便利重载: 纯连续因子场景</summary>
        List<GPRPrediction> PredictBatch(List<Dictionary<string, double>> factorValuesList);
        GPROptimalResult FindOptimal(bool maximize = true);
        /// <summary>★ 修复 (v3): 参数类型改为 object 以支持类别因子标签</summary>
        GPRStopDecision ShouldStop(List<Dictionary<string, object>> remainingRuns, double bestObserved,
                                    bool maximize = true, double eiThreshold = 0.01, double minRunsRatio = 1.5);
        /// <summary>便利重载: 纯连续因子场景</summary>
        GPRStopDecision ShouldStop(List<Dictionary<string, double>> remainingRuns, double bestObserved,
                                    bool maximize = true, double eiThreshold = 0.01, double minRunsRatio = 1.5);
        Dictionary<string, double> GetSensitivity();
        Task SaveStateAsync(string flowId);
        Task<bool> LoadStateAsync(string flowId);
        Task ResetModelAsync(string flowId, bool keepData);
        bool IsActive { get; }
        int DataCount { get; }
        Task SaveInitialStateAsync(string flowId);
        /// <summary>★ 修复 (v3): factorValues 改为 object 以支持类别因子标签</summary>
        void AppendData(Dictionary<string, object> factorValues, double responseValue,
                        string source = "measured", string batchName = "", string timestamp = "");
        /// <summary>便利重载: 纯连续因子场景</summary>
        void AppendData(Dictionary<string, double> factorValues, double responseValue,
                        string source = "measured", string batchName = "", string timestamp = "");
    }

    public class GPRTrainResult
    {
        public bool IsActive { get; set; }
        public double RSquared { get; set; }
        public double RMSE { get; set; }
        public Dictionary<string, double> Lengthscales { get; set; } = new();
        public int DataCount { get; set; }
        /// <summary>★ 新增: R² 类型标注 — "LOO-CV"(留一交叉验证) 而非训练 R²</summary>
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
        /// <summary>
        /// ★ 修复 (v3): 归一化置信度 — 基于预测标准差相对于预测值的比例
        /// 原来的 bug: 1.0 - PredictionStd 当 Std > 1 时为负数，完全无意义
        /// 修复: 使用 1 / (1 + CV) 其中 CV = Std / |Mean|，映射到 (0, 1] 区间
        /// CV=0 → 1.0 (完全确定), CV=1 → 0.5 (不确定性等于均值), CV→∞ → 0.0
        /// </summary>
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
        /// <summary>★ 新增: 优化方向 — "maximize" 或 "minimize"</summary>
        public string Direction { get; set; } = "maximize";
    }

    public class GPRStopDecision
    {
        public bool ShouldStop { get; set; }
        public string Reason { get; set; } = string.Empty;
        public double BestPredicted { get; set; }
        public double BestPredictedStd { get; set; }
        /// <summary>★ 新增: 剩余组中的最大 Expected Improvement</summary>
        public double MaxEI { get; set; }
    }

    // ═══════════════════════════════════════════════════════
    // IDOEAnalysisService — 数据分析服务
    //  修改: 新增 FitOlsAsync / GetResidualDiagnosticsAsync / GetEffectsParetoAsync
    // ═══════════════════════════════════════════════════════

    /// <summary>
    /// DOE 数据分析服务
    /// </summary>
    public interface IDOEAnalysisService
    {
        // ── 原有方法（保留） ──
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

        // ──  新增: 完整 OLS 分析 ──

        /// <summary>
        ///  新增: 完整 OLS 回归分析 — 返回 ANOVA 表 + 系数表 + 模型摘要
        /// 对标 JMP/Minitab 的分析报表
        /// </summary>
        /// <param name="batchId">批次 ID</param>
        /// <param name="responseName">响应变量名（支持多响应时逐个分析）</param>
        /// <param name="modelType">"linear"/"interaction"/"quadratic"</param>
        Task<OLSAnalysisResult> FitOlsAsync(string batchId, string responseName, string modelType = "quadratic");
        /// <summary>
        /// ★ 新增 v4: 自定义项列表的 OLS 回归（Model Reduction）
        /// 用户从 Pareto 图选择保留的显著项，重新拟合精简模型
        /// </summary>
        /// <param name="batchId">批次 ID</param>
        /// <param name="responseName">响应变量名</param>
        /// <param name="terms">保留的项列表，如 ["温度", "压力", "温度×压力", "温度²"]</param>
        Task<OLSAnalysisResult> FitOlsCustomAsync(string batchId, string responseName, List<string> terms);
        /// <summary>
        ///  新增: 残差诊断四图数据
        /// 正态概率图 / 残差vs拟合 / 残差vs顺序 / 残差直方图 + Cook's距离
        /// </summary>
        Task<string> GetResidualDiagnosticsAsync(string batchId, string responseName);

        /// <summary>
        ///  新增: 效应 Pareto 图（基于 T 值）
        /// 比原有的 GetParetoChartAsync 更准确（基于回归 T 值而非效应差值）
        /// </summary>
        Task<string> GetEffectsParetoAsync(string batchId, string responseName, double alpha = 0.05);

        /// <summary>
        /// ★ 新增 v5: 预测刻画器数据 — JMP Prediction Profiler 风格
        /// 对每个因子在其范围内扫描，返回一维剖面曲线数据
        /// </summary>
        Task<string> GetPredictionProfilerAsync(string batchId, string responseName,
            int gridSize = 50, Dictionary<string, object>? fixedValues = null);

        /// <summary>
        /// ★ 新增 v5: OLS 最优化求解 — 寻找使响应最大/最小的最优因子条件
        /// </summary>
        Task<string> FindOptimalAsync(string batchId, string responseName, bool maximize = true);
    }

    // ═══════════════════════════════════════════════════════
    //  新增: IDesirabilityService — 多响应优化服务
    // ═══════════════════════════════════════════════════════

    /// <summary>
    ///  新增: Desirability 多响应优化服务
    /// 调用 Python doe_desirability.py 实现 Derringer-Suich 方法
    /// </summary>
    public interface IDesirabilityService
    {
        /// <summary>
        /// 配置多响应优化参数
        /// </summary>
        Task ConfigureAsync(string batchId, List<DesirabilityResponseConfig> configs, List<DOEFactor> factors);

        /// <summary>
        /// 计算给定预测值的综合 Desirability
        /// </summary>
        Task<DesirabilityResult> ComputeDesirabilityAsync(Dictionary<string, double> predictions);

        /// <summary>
        /// 搜索最优因子组合（使综合 D 最大）
        /// </summary>
        Task<DesirabilityResult> OptimizeAsync();

        /// <summary>
        /// 获取 Profiler 图数据
        /// </summary>
        Task<string> GetProfileDataAsync(int gridSize = 50);

        /// <summary>
        /// 保存配置到数据库
        /// </summary>
        Task SaveConfigAsync(string batchId, List<DesirabilityResponseConfig> configs);

        /// <summary>
        /// 从数据库加载配置
        /// </summary>
        Task<List<DesirabilityResponseConfig>> LoadConfigAsync(string batchId);
    }

    // ═══════════════════════════════════════════════════════
    //  新增: IModelRouter — 模型策略路由器
    // ═══════════════════════════════════════════════════════

    /// <summary>
    ///  新增: 分析策略类型
    /// </summary>
    public enum AnalysisStrategy
    {
        /// <summary>
        /// OLS 回归 — RSM 设计 + 因子数 ≤6 + 数据量充足
        /// 输出: ANOVA 表、系数表、回归方程（可解释性强）
        /// </summary>
        OLS,

        /// <summary>
        /// GPR 代理模型 — 筛选设计 / 因子多 / 非线性强
        /// 输出: 敏感性排名、预测+不确定度、贝叶斯优化
        /// </summary>
        GPR,

        /// <summary>
        /// 数据不足，无法分析
        /// </summary>
        Insufficient
    }

    /// <summary>
    ///  新增: 模型策略路由器 — 根据设计方法和数据特征自动选择 OLS 或 GPR
    /// </summary>
    public interface IModelRouter
    {
        /// <summary>
        /// 根据批次信息决定分析策略
        /// </summary>
        AnalysisStrategy SelectStrategy(DOEDesignMethod method, int factorCount, int dataCount);

        /// <summary>
        /// 获取策略选择的理由（UI 提示用）
        /// </summary>
        string GetStrategyReason(AnalysisStrategy strategy, DOEDesignMethod method, int factorCount, int dataCount);
    }
}