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
    /// </summary>
    public class OLSAnalysisResult
    {
        /// <summary>
        /// ANOVA 方差分析表
        /// </summary>
        public List<AnovaRow> AnovaTable { get; set; } = new();

        /// <summary>
        /// 回归系数表
        /// </summary>
        public List<CoefficientRow> Coefficients { get; set; } = new();

        /// <summary>
        /// 模型摘要统计
        /// </summary>
        public OLSModelSummary ModelSummary { get; set; } = new();

        /// <summary>
        /// ★ v7 新增: 因数据不足或共线性被自动剔除的模型项列表
        /// 如 ["温度²", "温度×压力"]
        /// 空列表表示所有请求的项均可估计
        /// </summary>
        public List<string> DroppedTerms { get; set; } = new();

        /// <summary>
        /// ★ v7 新增: 不可估计项警告信息（人类可读）
        /// 如 "以下项因数据不足或共线性被自动剔除: 温度², 温度×压力"
        /// 空字符串表示无警告
        /// </summary>
        public string InestimableWarning { get; set; } = string.Empty;

        /// <summary>
        /// ★ v7 新增: 用户请求的原始模型类型
        /// 如用户请求 "quadratic" 但因数据不足被降阶为 "linear"，
        /// 此字段记录原始请求 "quadratic"，实际拟合的模型从系数表推断
        /// </summary>
        public string OriginalModelType { get; set; } = string.Empty;
    }

    /// <summary>
    /// ANOVA 表行
    /// 
    /// ANOVA 分解原理:
    ///   SS_Total = SS_Model + SS_Residual
    ///   SS_Residual = SS_LackOfFit + SS_PureError
    ///   F = MS_Source / MS_Residual ~ F(df_source, df_residual)
    /// </summary>
    public class AnovaRow
    {
        /// <summary>
        /// 变异来源 (如: "温度", "温度×压力", "温度²", "残差", "失拟", "纯误差", "总计")
        /// </summary>
        public string Source { get; set; } = string.Empty;

        /// <summary>
        /// 自由度
        /// </summary>
        public int DF { get; set; }

        /// <summary>
        /// 平方和 (Sum of Squares)
        /// </summary>
        public double SS { get; set; }

        /// <summary>
        /// 均方 (Mean Square) = SS / DF
        /// </summary>
        public double? MS { get; set; }

        /// <summary>
        /// F 统计量
        /// </summary>
        public double? FValue { get; set; }

        /// <summary>
        /// P 值 — &lt;0.05 表示显著
        /// </summary>
        public double? PValue { get; set; }
    }

    /// <summary>
    /// 回归系数表行
    /// 
    /// 每一行对应模型中的一个项（截距、主效应、交互项、二次项）
    /// </summary>
    public class CoefficientRow
    {
        /// <summary>
        /// 项名称 (如: "截距", "温度", "温度×压力", "温度²")
        /// </summary>
        public string Term { get; set; } = string.Empty;

        /// <summary>
        /// 回归系数 β
        /// </summary>
        public double Coefficient { get; set; }

        /// <summary>
        /// 标准误差 SE(β)
        /// </summary>
        public double StdError { get; set; }

        /// <summary>
        /// T 统计量 = β / SE(β)
        /// </summary>
        public double TValue { get; set; }

        /// <summary>
        /// P 值 — &lt;0.05 表示该项显著
        /// </summary>
        public double PValue { get; set; }

        /// <summary>
        /// 方差膨胀因子 (截距无 VIF)
        /// VIF &lt; 5 优秀, 5-10 可接受, &gt;10 严重共线性
        /// </summary>
        public double? VIF { get; set; }
    }

    /// <summary>
    /// OLS 模型摘要
    /// </summary>
    public class OLSModelSummary
    {
        /// <summary>
        /// 决定系数 R² = 1 - SS_Res/SS_Total
        /// </summary>
        public double RSquared { get; set; }

        /// <summary>
        /// 调整决定系数 R²_adj = 1 - (SS_Res/(n-p)) / (SS_Total/(n-1))
        /// 惩罚过多的参数，防止过拟合
        /// </summary>
        public double RSquaredAdj { get; set; }

        /// <summary>
        /// 预测决定系数 R²_pred = 1 - PRESS/SS_Total
        /// 基于留一交叉验证，衡量模型的预测能力
        /// R²_pred 和 R²_adj 差距 &gt; 0.2 表示可能过拟合
        /// </summary>
        public double RSquaredPred { get; set; }

        /// <summary>
        /// 均方根误差 RMSE = √(MS_Residual)
        /// </summary>
        public double RMSE { get; set; }

        /// <summary>
        /// 适当精度 (Adequate Precision) = signal / noise
        /// &gt;4 表示信噪比足够，模型可用于预测
        /// </summary>
        public double AdequatePrecision { get; set; }

        /// <summary>
        /// PRESS 统计量（预测残差平方和）
        /// </summary>
        public double PRESS { get; set; }

        /// <summary>
        /// 失拟检验 P 值
        /// P &gt; 0.05 表示模型无显著失拟（好事）
        /// P &lt; 0.05 表示模型可能不充分
        /// null 表示无法计算（没有重复实验）
        /// </summary>
        public double? LackOfFitP { get; set; }

        /// <summary>
        /// 整体模型 F 检验 P 值
        /// P &lt; 0.05 表示模型整体显著
        /// </summary>
        public double? ModelP { get; set; }

        /// <summary>
        /// 回归方程字符串
        /// 如: "转化率 = 85.2 + 12.3×温度 + 5.1×压力 - 2.5×温度²"
        /// </summary>
        public string Equation { get; set; } = string.Empty;

        /// <summary>
        /// ★ v4: Minitab 风格方程 — 按类别因子水平展开
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