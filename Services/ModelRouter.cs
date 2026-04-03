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
            // 计算二阶模型参数数: p = 1 + 2k + C(k,2)
            int p_quadratic = 1 + 2 * factorCount + factorCount * (factorCount - 1) / 2;

            // 数据不足判断: 至少需要 p+2 个数据点（p个参数 + 2个自由度做误差估计）
            if (dataCount < p_quadratic + 2 && dataCount < factorCount + 2)
            {
                return AnalysisStrategy.Insufficient;
            }

            // RSM 设计 → OLS
            if (method == DOEDesignMethod.CCD ||
                method == DOEDesignMethod.BoxBehnken ||
                method == DOEDesignMethod.DOptimal)
            {
                // 数据量够拟合二阶模型才走 OLS，否则降级到 GPR
                if (dataCount >= p_quadratic + 2)
                    return AnalysisStrategy.OLS;
                else
                    return AnalysisStrategy.GPR;
            }

            // 全因子设计: 因子 ≤6 且数据量够 → OLS
            if (method == DOEDesignMethod.FullFactorial)
            {
                if (factorCount <= 6 && dataCount >= p_quadratic + 2)
                    return AnalysisStrategy.OLS;
                else
                    return AnalysisStrategy.GPR;
            }

            // 筛选类设计 → GPR（部分因子、Taguchi 数据结构不适合二阶 OLS）
            // Custom → GPR（不确定设计结构）
            return AnalysisStrategy.GPR;
        }

        /// <summary>
        /// 获取策略选择的理由（供 UI 提示）
        /// </summary>
        public string GetStrategyReason(AnalysisStrategy strategy, DOEDesignMethod method, int factorCount, int dataCount)
        {
            int p_quadratic = 1 + 2 * factorCount + factorCount * (factorCount - 1) / 2;

            return strategy switch
            {
                AnalysisStrategy.OLS => method switch
                {
                    DOEDesignMethod.CCD => $"CCD 设计适配二阶多项式模型，{dataCount} 组数据可拟合 {p_quadratic} 参数的 OLS 回归，提供 ANOVA 表和回归方程。",
                    DOEDesignMethod.BoxBehnken => $"Box-Behnken 设计适配二阶多项式模型，提供 ANOVA 表和回归方程。",
                    DOEDesignMethod.DOptimal => $"D-最优设计针对 OLS 回归优化，{dataCount} 组数据可拟合完整模型。",
                    DOEDesignMethod.FullFactorial => $"全因子设计 ({factorCount}因子) 数据完整，使用 OLS 回归给出 ANOVA 表。",
                    _ => "使用 OLS 回归分析。"
                },
                AnalysisStrategy.GPR => method switch
                {
                    DOEDesignMethod.FractionalFactorial => $"部分因子设计为筛选型，使用 GPR 代理模型进行敏感性分析和预测。",
                    DOEDesignMethod.Taguchi => $"正交设计使用 GPR 代理模型，提供参数敏感性排名和最优搜索。",
                    DOEDesignMethod.FullFactorial => $"全因子设计因子数较多 ({factorCount}个)，使用 GPR 模型避免参数爆炸。",
                    _ => $"使用 GPR 代理模型进行分析，提供敏感性排名、预测和贝叶斯优化。"
                },
                AnalysisStrategy.Insufficient => $"当前 {dataCount} 组数据不足以拟合 {factorCount} 因子的模型（至少需要 {Math.Min(p_quadratic + 2, factorCount + 2)} 组）。请继续执行更多实验。",
                _ => ""
            };
        }
    }
}