namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// DOE 停止条件
    /// </summary>
    public class DOEStopCondition
    {
        public int Id { get; set; }
        public string BatchId { get; set; } = string.Empty;

        /// <summary>
        /// 条件类型：阈值判断 / 模型智能判断
        /// </summary>
        public DOEStopConditionType ConditionType { get; set; } = DOEStopConditionType.Threshold;

        /// <summary>
        /// 响应变量名称（阈值类型时使用）
        /// </summary>
        public string? ResponseName { get; set; }

        /// <summary>
        /// 运算符（复用现有 IF 节点的运算符体系）
        /// </summary>
        public string Operator { get; set; } = "GreaterThanOrEqual";

        /// <summary>
        /// 目标值（阈值类型时使用）
        /// </summary>
        public double TargetValue { get; set; }

        /// <summary>
        /// 逻辑组标识（多个条件之间 AND / OR）
        /// </summary>
        public DOELogicGroup LogicGroup { get; set; } = DOELogicGroup.And;
    }

    /// <summary>
    /// 停止条件类型
    /// </summary>
    public enum DOEStopConditionType
    {
        /// <summary>
        /// 用户硬阈值（如 转化率 >= 95%）
        /// </summary>
        Threshold,

        /// <summary>
        /// GPR 模型智能停止（预测后续无法超越当前最优）
        /// </summary>
        ModelPrediction
    }

    /// <summary>
    /// 条件逻辑关系
    /// </summary>
    public enum DOELogicGroup
    {
        And,
        Or
    }
}
