using Newtonsoft.Json;
using System.Collections.Generic;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// OLS 回归分析完整结果
    /// 由 Python DOEAnalyzer.fit_ols() 返回，C# 侧解析展示
    /// 对标 JMP/Minitab 的分析报表格式
    /// 
    /// ★ v7 新增: DroppedTerms / InestimableWarning / OriginalModelType
    /// ★ v13 新增: UncCodedCoefficients / UncodedEquation / UncodedEquations / CodingInfo
    /// </summary>
    public class OLSAnalysisResult
    {
        /// <summary>
        /// ANOVA 方差分析表
        /// </summary>
        public List<AnovaRow> AnovaTable { get; set; } = new();

        /// <summary>
        /// 回归系数表（编码值 [-1,+1] 系数）
        /// </summary>
        public List<CoefficientRow> Coefficients { get; set; } = new();

        /// <summary>
        /// ★ v13 新增: 未编码（原始值）回归系数表
        /// 对标 Minitab "以未编码单位表示的回归方程" 的系数
        /// C# 端用户选择"未编码"时展示此系数表
        /// </summary>
        public List<UncCodedCoefficientRow> UncodedCoefficients { get; set; } = new();

        /// <summary>
        /// 模型摘要统计
        /// </summary>
        public OLSModelSummary ModelSummary { get; set; } = new();

        /// <summary>
        /// ★ v13 新增: 因子编码信息（center / half_range）
        /// C# 端可用于显示编码参考或其他用途
        /// </summary>
        public Dictionary<string, CodingInfoItem> CodingInfo { get; set; } = new();

        /// <summary>
        /// ★ v13 新增: 未编码回归方程字符串（含哑变量项的完整方程）
        /// </summary>
        public string UncodedEquation { get; set; } = string.Empty;

        /// <summary>
        /// ★ v13 新增: 未编码回归方程 — 按类别因子水平展开
        /// </summary>
        public EquationsInfo? UncodedEquations { get; set; }

        /// <summary>
        /// ★ v7 新增: 因数据不足或共线性被自动剔除的模型项列表
        /// </summary>
        public List<string> DroppedTerms { get; set; } = new();

        /// <summary>
        /// ★ v7 新增: 不可估计项警告信息（人类可读）
        /// </summary>
        public string InestimableWarning { get; set; } = string.Empty;

        /// <summary>
        /// ★ v7 新增: 用户请求的原始模型类型
        /// </summary>
        public string OriginalModelType { get; set; } = string.Empty;
        [JsonProperty("xtx_inv_flat")]
        public double[]? XtxInvFlat { get; set; }

        [JsonProperty("sigma")]
        public double Sigma { get; set; }

        [JsonProperty("t_crit")]
        public double TCrit { get; set; }

        [JsonProperty("param_count")]
        public int ParamCount { get; set; }
        /// <summary>
        /// ★ v17 新增: 弯曲检验结果 (全因子设计 + 中心点时有值)
        /// 对标 Minitab 的 Ct Pt 行和弯曲检验
        /// </summary>
        public CurvatureTestResult? CurvatureTest { get; set; }
    }
    /// <summary>
    /// ★ v17 新增: 弯曲检验结果 — 对标 Minitab 全因子分析的 Ct Pt
    /// 
    /// 当设计方法为 FullFactorial/FractionalFactorial 且数据含中心点时,
    /// Python 端会自动执行弯曲检验:
    ///   1. 仅用角点拟合模型
    ///   2. 中心点做弯曲检验 (Ct Pt)
    ///   3. 误差 = 失拟 + 纯误差
    /// 
    /// 此时 ANOVA 表和系数表已经包含弯曲行,
    /// 本类提供额外的检验元数据供 UI 层使用。
    /// </summary>
    public class CurvatureTestResult
    {
        /// <summary>是否执行了弯曲检验</summary>
        [JsonProperty("has_curvature_test")]
        public bool HasCurvatureTest { get; set; }

        /// <summary>角点(因子点)数量</summary>
        [JsonProperty("n_corner")]
        public int CornerCount { get; set; }

        /// <summary>中心点数量</summary>
        [JsonProperty("n_center")]
        public int CenterCount { get; set; }

        /// <summary>角点响应均值 (= Minitab 截距)</summary>
        [JsonProperty("corner_mean")]
        public double CornerMean { get; set; }

        /// <summary>中心点响应均值</summary>
        [JsonProperty("center_mean")]
        public double CenterMean { get; set; }

        /// <summary>Ct Pt 系数 = 中心点均值 - 角点均值</summary>
        [JsonProperty("ct_pt_coeff")]
        public double CtPtCoefficient { get; set; }

        /// <summary>Ct Pt 标准误</summary>
        [JsonProperty("ct_pt_se")]
        public double CtPtStdError { get; set; }

        /// <summary>Ct Pt T值</summary>
        [JsonProperty("ct_pt_t")]
        public double CtPtTValue { get; set; }

        /// <summary>Ct Pt P值</summary>
        [JsonProperty("ct_pt_p")]
        public double CtPtPValue { get; set; }

        /// <summary>弯曲 SS</summary>
        [JsonProperty("curvature_ss")]
        public double CurvatureSS { get; set; }

        /// <summary>弯曲 DF</summary>
        [JsonProperty("curvature_df")]
        public int CurvatureDF { get; set; }

        /// <summary>弯曲 F值</summary>
        [JsonProperty("curvature_f")]
        public double CurvatureF { get; set; }

        /// <summary>弯曲 P值</summary>
        [JsonProperty("curvature_p")]
        public double CurvatureP { get; set; }

        /// <summary>弯曲是否显著 (P < 0.05 表示存在曲率，需要做 RSM)</summary>
        public bool IsSignificant => CurvatureP < 0.05;
    }
    /// <summary>
    /// ANOVA 表行
    /// </summary>
    public class AnovaRow
    {
        public string Source { get; set; } = string.Empty;
        public int DF { get; set; }
        public double SS { get; set; }
        public double? MS { get; set; }
        public double? FValue { get; set; }
        public double? PValue { get; set; }
        public bool IsSignificant => PValue.HasValue && PValue.Value < 0.05;
    }

    /// <summary>
    /// 回归系数表行（编码值系数）
    /// </summary>
    public class CoefficientRow
    {
        public string Term { get; set; } = string.Empty;
        public double Coefficient { get; set; }
        public double StdError { get; set; }
        public double TValue { get; set; }
        public double PValue { get; set; }
        public double? VIF { get; set; }
        /// <summary>
        /// ★ v17 新增: 效应值 (Effect = 2 × 编码系数)
        /// 仅全因子设计时有值，RSM 时为 null
        /// UI 层判断: 如果 CurvatureTest != null，则显示效应列
        /// </summary>
        public double? Effect { get; set; }
    }

    // 新代码:
    /// <summary>
    /// ★ v14 改进: 未编码系数行 — 现在包含完整的 SE/T/P/VIF
    /// 切换到"未编码单位"时，系数表可以显示完整统计量
    /// </summary>
    public class UncCodedCoefficientRow
    {
        public string Term { get; set; } = string.Empty;
        public double Coefficient { get; set; }
        public double StdError { get; set; }
        public double TValue { get; set; }
        public double PValue { get; set; }
        public double? VIF { get; set; }
    }
    /// <summary>
    /// ★ v13 新增: 因子编码信息
    /// </summary>
    public class CodingInfoItem
    {
        [JsonProperty("type")]
        public string Type { get; set; } = "continuous";

        [JsonProperty("center")]
        public double? Center { get; set; }

        [JsonProperty("half_range")]
        public double? HalfRange { get; set; }
    }

    /// <summary>
    /// OLS 模型摘要
    /// </summary>
    public class OLSModelSummary
    {
        public double RSquared { get; set; }
        public double RSquaredAdj { get; set; }
        public double RSquaredPred { get; set; }
        public double RMSE { get; set; }
        public double AdequatePrecision { get; set; }
        public double PRESS { get; set; }
        public double? LackOfFitP { get; set; }
        public double? ModelP { get; set; }
        public string Equation { get; set; } = string.Empty;

        /// <summary>
        /// ★ v4: Minitab 风格方程 — 按类别因子水平展开（编码值）
        /// </summary>
        public EquationsInfo? Equations { get; set; }
    }

    /// <summary>
    /// ★ v4: Minitab 风格方程展示信息
    /// </summary>
    public class EquationsInfo
    {
        [JsonProperty("has_categorical")]
        public bool HasCategorical { get; set; }

        [JsonProperty("categorical_factor")]
        public string? CategoricalFactor { get; set; }

        [JsonProperty("equations")]
        public Dictionary<string, string> EquationsByLevel { get; set; } = new();

        [JsonProperty("common_equation")]
        public string CommonEquation { get; set; } = string.Empty;
    }
}