using System;
using System.Collections.Generic;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// ★ 新增: DOE 优化项目 — 跨批次的优化目标容器
    /// 
    /// 数据层级: DOEProject → DOEBatch → DOERunRecord
    /// 
    /// 项目串联多轮迭代:
    ///   Round 1 (Screening)  → 部分因子设计, 筛掉不显著因子
    ///   Round 2 (PathSearch) → 最速上升, 移动到高坡区
    ///   Round 3 (RSM)        → CCD/BBD, 拟合二阶模型
    ///   Round 4 (Augmenting) → 增强设计补点, 改善模型
    ///   Round 5 (Confirmation) → 验证实验, 确认最优
    /// </summary>
    public class DOEProject
    {
        public int Id { get; set; }

        /// <summary>
        /// 项目唯一标识（GUID 短码，如 "PROJ_20260406_a3b5c7"）
        /// </summary>
        public string ProjectId { get; set; } = string.Empty;

        /// <summary>
        /// 项目名称（用户可读，如 "XX产品产率优化"）
        /// </summary>
        public string ProjectName { get; set; } = string.Empty;

        /// <summary>
        /// 优化目标描述（如 "最大化产率，约束纯度≥99%"）
        /// </summary>
        public string ObjectiveDescription { get; set; } = string.Empty;

        /// <summary>
        /// 关联的流程 ID（可选，项目可以跨流程）
        /// </summary>
        public string? FlowId { get; set; }

        /// <summary>
        /// 关联的流程名称
        /// </summary>
        public string? FlowName { get; set; }

        /// <summary>
        /// 当前优化阶段
        /// </summary>
        public DOEProjectPhase CurrentPhase { get; set; } = DOEProjectPhase.Screening;

        /// <summary>
        /// 项目状态
        /// </summary>
        public DOEProjectStatus Status { get; set; } = DOEProjectStatus.Active;

        /// <summary>
        /// 当前最优响应值
        /// </summary>
        public double? BestResponseValue { get; set; }

        /// <summary>
        /// 当前最优因子组合 JSON
        /// 如 {"温度": 135, "压力": 22, "催化剂类型": "B"}
        /// </summary>
        public string? BestFactorsJson { get; set; }

        /// <summary>
        /// 总实验次数（跨所有批次累计）
        /// </summary>
        public int TotalExperiments { get; set; }

        /// <summary>
        /// 已完成的轮次（批次）数
        /// </summary>
        public int CompletedRounds { get; set; }

        public DateTime CreatedTime { get; set; } = DateTime.Now;
        public DateTime UpdatedTime { get; set; } = DateTime.Now;

        // ── 导航属性（非数据库字段）──

        /// <summary>项目下所有因子（含已淘汰的）</summary>
        public List<DOEProjectFactor> ProjectFactors { get; set; } = new();

        /// <summary>项目下所有批次</summary>
        public List<DOEBatch> Batches { get; set; } = new();

        /// <summary>轮次决策历史</summary>
        public List<DOERoundSummary> RoundSummaries { get; set; } = new();
    }

    /// <summary>
    /// 优化阶段
    /// </summary>
    public enum DOEProjectPhase
    {
        /// <summary>筛选 — 因子多，识别关键因子</summary>
        Screening,
        /// <summary>路径探索 — 沿最速上升/下降方向移动</summary>
        PathSearch,
        /// <summary>响应面优化 — CCD/BBD 拟合二阶模型</summary>
        RSM,
        /// <summary>增强补点 — 模型不够好，补充实验</summary>
        Augmenting,
        /// <summary>验证 — 在预测最优点做重复实验确认</summary>
        Confirmation,
        /// <summary>已完成 — 优化目标达成</summary>
        Completed
    }

    /// <summary>
    /// 项目状态
    /// </summary>
    public enum DOEProjectStatus
    {
        /// <summary>活跃 — 正在优化中</summary>
        Active,
        /// <summary>暂停 — 用户暂停优化</summary>
        Paused,
        /// <summary>完成 — 优化目标达成</summary>
        Completed,
        /// <summary>归档 — 不再使用</summary>
        Archived
    }
}
