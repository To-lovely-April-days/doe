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