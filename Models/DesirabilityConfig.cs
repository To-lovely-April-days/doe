using System.Collections.Generic;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    ///  新增: Desirability 目标类型
    /// </summary>
    public enum DesirabilityGoal
    {
        /// <summary>
        /// 最大化 — d(y) 随 y 增大而增大
        /// </summary>
        Maximize,

        /// <summary>
        /// 最小化 — d(y) 随 y 减小而增大
        /// </summary>
        Minimize,

        /// <summary>
        /// 达标 — d(y) 在 target 附近最大，偏离时减小
        /// </summary>
        Target
    }

    /// <summary>
    ///  新增: 单个响应变量的 Desirability 配置
    /// 
    /// Derringer-Suich (1980) 变换函数:
    ///   Maximize: d(y) = 0 if y≤L, ((y-L)/(T-L))^s if L&lt;y&lt;T, 1 if y≥T
    ///   Minimize: d(y) = 1 if y≤T, ((U-y)/(U-T))^s if T&lt;y&lt;U, 0 if y≥U
    ///   Target:   d(y) = ((y-L)/(T-L))^s1 if L≤y≤T, ((U-y)/(U-T))^s2 if T&lt;y≤U
    /// </summary>
    public class DesirabilityResponseConfig
    {
        /// <summary>
        /// 响应变量名称（必须与 DOEResponse.ResponseName 匹配）
        /// </summary>
        public string ResponseName { get; set; } = string.Empty;

        /// <summary>
        /// 优化目标类型
        /// </summary>
        public DesirabilityGoal Goal { get; set; } = DesirabilityGoal.Maximize;

        /// <summary>
        /// 可接受下界 L
        /// </summary>
        public double Lower { get; set; }

        /// <summary>
        /// 可接受上界 U
        /// </summary>
        public double Upper { get; set; }

        /// <summary>
        /// 目标值 T（Maximize 时 T=Upper, Minimize 时 T=Lower）
        /// </summary>
        public double Target { get; set; }

        /// <summary>
        /// 权重 wⱼ — 控制该响应在综合 D 中的相对重要性
        /// </summary>
        public double Weight { get; set; } = 1.0;

        /// <summary>
        /// 重要度 1-5（UI 展示用，内部转为 Weight 的缩放因子）
        /// 5=最重要 3=标准 1=不太重要
        /// </summary>
        public int Importance { get; set; } = 3;

        /// <summary>
        /// 形状参数 s — 控制变换曲线的凹凸
        /// s=1: 线性, s&lt;1: 凸形（容易满足）, s&gt;1: 凹形（难以满足）
        /// </summary>
        public double Shape { get; set; } = 1.0;

        /// <summary>
        /// Target 类型下半段形状参数（L→T 段）
        /// </summary>
        public double ShapeLower { get; set; } = 1.0;

        /// <summary>
        /// Target 类型上半段形状参数（T→U 段）
        /// </summary>
        public double ShapeUpper { get; set; } = 1.0;
    }

    /// <summary>
    ///  新增: 批次级 Desirability 配置（存入数据库）
    /// </summary>
    public class DesirabilityBatchConfig
    {
        public int Id { get; set; }
        public string BatchId { get; set; } = string.Empty;

        /// <summary>
        /// 各响应变量的配置列表
        /// </summary>
        public List<DesirabilityResponseConfig> ResponseConfigs { get; set; } = new();
    }

    /// <summary>
    ///  新增: Desirability 优化结果
    /// </summary>
    public class DesirabilityResult
    {
        /// <summary>
        /// 最优因子组合
        /// ★ 修复 (Bug#1): 从 Dictionary&lt;string, double&gt; 改为 Dictionary&lt;string, object&gt;
        /// 连续因子存 double (如 135.2)，类别因子存 string (如 "催化剂B")
        /// 原来的 bug: Python 端 optimize() 返回的 optimal_factors 包含字符串值（类别因子），
        /// JSON 反序列化为 double 时失败导致整个字典变 null
        /// </summary>
        public Dictionary<string, object> OptimalFactors { get; set; } = new();

        /// <summary>
        /// 综合 Desirability 值 (0-1)
        /// D = (d₁^w₁ × d₂^w₂ × ...)^(1/Σwⱼ)
        /// </summary>
        public double CompositeD { get; set; }

        /// <summary>
        /// 各响应的个体 Desirability
        /// </summary>
        public List<IndividualDesirability> IndividualD { get; set; } = new();

        /// <summary>
        /// 优化是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 错误信息（失败时）
        /// </summary>
        public string? ErrorMessage { get; set; }
    }

    /// <summary>
    ///  新增: 单个响应的 Desirability 结果
    /// </summary>
    public class IndividualDesirability
    {
        /// <summary>
        /// 响应变量名
        /// </summary>
        public string ResponseName { get; set; } = string.Empty;

        /// <summary>
        /// 在最优因子下的预测值
        /// </summary>
        public double PredictedValue { get; set; }

        /// <summary>
        /// 个体 Desirability (0-1)
        /// </summary>
        public double D { get; set; }

        /// <summary>
        /// 优化目标
        /// </summary>
        public DesirabilityGoal Goal { get; set; }
    }
}