using MaxChemical.Modules.DOE.Models;

namespace MaxChemical.Modules.DOE.Services
{
    /// <summary>
    ///  新增: 模型策略路由器 — 根据设计方法和数据特征自动选择 OLS 或 GPR
    /// 
    /// 路由规则:
    /// ┌─────────────────────┬──────────────┬──────────────────────────────────────┐
    /// │ 设计方法            │ 分析策略     │ 原因                                  │
    /// ├─────────────────────┼──────────────┼──────────────────────────────────────┤
    /// │ CCD / BoxBehnken    │ OLS          │ RSM 设计天生适配二阶多项式模型        │
    /// │ DOptimal            │ OLS          │ 同上（D-最优就是为 OLS 设计的）        │
    /// │ FullFactorial (≤6)  │ OLS          │ 全因子数据完整，OLS 可给出 ANOVA 表    │
    /// │ FullFactorial (>6)  │ GPR          │ 因子太多，OLS 参数爆炸                │
    /// │ FractionalFactorial │ GPR          │ 筛选设计，数据量不够拟合二阶模型       │
    /// │ Taguchi             │ GPR          │ 正交设计不保证 OLS 模型可估            │
    /// │ Custom              │ GPR          │ 不确定设计结构，GPR 更鲁棒             │
    /// │ 任何 (数据 < p+2)   │ Insufficient │ 数据量不足以拟合任何模型               │
    /// └─────────────────────┴──────────────┴──────────────────────────────────────┘
    /// 
    /// 其中 p = 二阶模型参数数 = 1 + 2k + C(k,2)
    /// </summary>
    public class ModelRouter : IModelRouter
    {
        /// <summary>
        /// 根据批次信息决定分析策略
        /// </summary>
        public AnalysisStrategy SelectStrategy(DOEDesignMethod method, int factorCount, int dataCount)
        {
            // 所有设计方法都走 OLS，只是模型类型不同（由 olsModelType 控制）
            // 
            // 路由规则 v2 (对齐 Minitab):
            // ┌─────────────────────┬──────────────┬──────────────────┐
            // │ 设计方法            │ 分析策略     │ OLS 模型类型     │
            // ├─────────────────────┼──────────────┼──────────────────┤
            // │ PlackettBurman      │ OLS          │ linear           │
            // │ FractionalFactorial │ OLS          │ linear/interact  │
            // │ FullFactorial       │ OLS          │ interact/linear  │
            // │ CCD / BoxBehnken    │ OLS          │ quadratic        │
            // │ DOptimal            │ OLS          │ 用户选择         │
            // │ Custom              │ OLS          │ 用户选择         │
            // └─────────────────────┴──────────────┴──────────────────┘

            // 根据模型类型计算所需参数数
            // linear:      p = 1 + k
            // interaction: p = 1 + k + C(k,2)
            // quadratic:   p = 1 + 2k + C(k,2)
            // 这里用最保守的 quadratic 判断数据是否充足
            int p_quadratic = 1 + 2 * factorCount + factorCount * (factorCount - 1) / 2;
            int p_linear = 1 + factorCount;

            // 数据不足：连一阶主效应模型都拟合不了
            if (dataCount < p_linear + 1)
            {
                return AnalysisStrategy.Insufficient;
            }

            // 全部走 OLS
            return AnalysisStrategy.OLS;
        }

        /// <summary>
        /// 获取策略选择的理由（供 UI 提示）
        /// </summary>
        public string GetStrategyReason(AnalysisStrategy strategy, DOEDesignMethod method, int factorCount, int dataCount)
        {
            int p_linear = 1 + factorCount;
            int p_interaction = 1 + factorCount + factorCount * (factorCount - 1) / 2;
            int p_quadratic = 1 + 2 * factorCount + factorCount * (factorCount - 1) / 2;

            return strategy switch
            {
                AnalysisStrategy.OLS => method switch
                {
                    DOEDesignMethod.PlackettBurman =>
                        $"Plackett-Burman 筛选设计，使用一阶主效应 OLS 回归，识别显著因子。",
                    DOEDesignMethod.FractionalFactorial =>
                        $"部分因子设计，使用 OLS 回归分析主效应" +
                        (dataCount >= p_interaction + 2 ? "和交互效应。" : "。"),
                    DOEDesignMethod.FullFactorial =>
                        $"全因子设计 ({factorCount}因子)，使用 OLS 回归给出完整 ANOVA 表。",
                    DOEDesignMethod.CCD =>
                        $"CCD 设计适配二阶多项式模型，{dataCount} 组数据拟合 {p_quadratic} 参数 OLS 回归。",
                    DOEDesignMethod.BoxBehnken =>
                        $"Box-Behnken 设计适配二阶多项式模型，提供 ANOVA 表和回归方程。",
                    DOEDesignMethod.DOptimal =>
                        $"D-最优设计，{dataCount} 组数据拟合 OLS 回归模型。",
                    DOEDesignMethod.Custom =>
                        $"自定义设计，按用户选择的分析类型进行 OLS 回归。",
                    _ => "使用 OLS 回归分析。"
                },
                AnalysisStrategy.Insufficient =>
                    $"当前 {dataCount} 组数据不足以拟合 {factorCount} 因子的模型" +
                    $"（至少需要 {p_linear + 1} 组）。请继续执行更多实验。",
                _ => ""
            };
        }

    }
}