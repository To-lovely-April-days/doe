using System;
using System.Collections.Generic;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// ★ 新增: 轮次总结 — 记录每个批次完成后的分析结论和决策
    /// 
    /// 每个批次执行完毕后，系统自动生成一份轮次总结：
    ///   - 模型质量指标
    ///   - 因子显著性排名
    ///   - 当前最优点及其位置（是否在边界上）
    ///   - 系统推荐的下一步操作
    ///   - 用户的实际决策
    /// 
    /// 这些记录构成项目的"决策链"，用户可以回顾每一步为什么这么做。
    /// </summary>
    public class DOERoundSummary
    {
        public int Id { get; set; }

        /// <summary>关联的项目 ID</summary>
        public string ProjectId { get; set; } = string.Empty;

        /// <summary>关联的批次 ID</summary>
        public string BatchId { get; set; } = string.Empty;

        /// <summary>轮次序号（从 1 开始）</summary>
        public int RoundNumber { get; set; }

        /// <summary>该轮的优化阶段</summary>
        public DOEProjectPhase Phase { get; set; }

        /// <summary>该轮使用的设计方法</summary>
        public DOEDesignMethod DesignMethod { get; set; }

        /// <summary>该轮的活跃因子数</summary>
        public int ActiveFactorCount { get; set; }

        /// <summary>该轮的实验组数</summary>
        public int RunCount { get; set; }

        // ── 模型质量 ──

        /// <summary>R²</summary>
        public double? RSquared { get; set; }

        /// <summary>R² adjusted</summary>
        public double? RSquaredAdj { get; set; }

        /// <summary>R² predicted</summary>
        public double? RSquaredPred { get; set; }

        /// <summary>Lack-of-Fit p 值</summary>
        public double? LackOfFitP { get; set; }

        /// <summary>GPR 模型 R²（如果使用 GPR）</summary>
        public double? GprRSquared { get; set; }

        // ── 因子分析 ──

        /// <summary>
        /// 因子显著性排名 JSON
        /// 格式: [{"name": "温度", "p_value": 0.001, "sensitivity": 0.85, "significant": true}, ...]
        /// </summary>
        public string? FactorRankingJson { get; set; }

        /// <summary>
        /// 本轮淘汰的因子列表（逗号分隔）
        /// </summary>
        public string? ScreenedOutFactors { get; set; }

        // ── 最优点 ──

        /// <summary>该轮找到的最优响应值</summary>
        public double? BestResponseValue { get; set; }

        /// <summary>最优因子组合 JSON</summary>
        public string? BestFactorsJson { get; set; }

        /// <summary>
        /// 最优点是否在因子边界上
        /// JSON: {"温度": "at_upper", "压力": "interior", ...}
        /// 值: "at_lower" | "at_upper" | "interior"
        /// </summary>
        public string? OptimalBoundaryStatusJson { get; set; }

        /// <summary>
        /// 最大 Expected Improvement（GPR 模型的改进空间指标）
        /// </summary>
        public double? MaxEI { get; set; }

        // ── 决策 ──

        /// <summary>
        /// 系统推荐的下一步操作
        /// </summary>
        public NextStepRecommendation Recommendation { get; set; }

        /// <summary>
        /// 推荐理由
        /// </summary>
        public string? RecommendationReason { get; set; }

        /// <summary>
        /// 用户实际选择的下一步（可能与推荐不同）
        /// </summary>
        public NextStepRecommendation? UserDecision { get; set; }

        /// <summary>
        /// 用户的决策备注
        /// </summary>
        public string? UserNotes { get; set; }

        public DateTime CreatedTime { get; set; } = DateTime.Now;
    }

    /// <summary>
    /// 系统推荐的下一步操作
    /// </summary>
    public enum NextStepRecommendation
    {
        /// <summary>继续筛选（因子仍然太多）</summary>
        ContinueScreening,

        /// <summary>执行最速上升/下降</summary>
        SteepestAscent,

        /// <summary>进入 RSM 精细优化</summary>
        StartRSM,

        /// <summary>增强设计补点（模型不够好）</summary>
        AugmentDesign,

        /// <summary>扩展因子范围（最优点在边界上）</summary>
        ExpandRange,

        /// <summary>进行验证实验</summary>
        ConfirmationRuns,

        /// <summary>项目完成</summary>
        Complete,

        /// <summary>无法自动判断，需要用户决策</summary>
        UserDecision
    }
}
