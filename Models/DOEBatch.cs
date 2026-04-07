using System;
using System.Collections.Generic;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// DOE 批次主实体
    /// 
    /// ★ 修改: 新增 ProjectId 和 RoundNumber 字段（可空，向后兼容）
    ///   - ProjectId 为 null 时，行为与旧版完全一致（单批次模式）
    ///   - ProjectId 不为 null 时，该批次属于某个优化项目的某一轮
    /// </summary>
    public class DOEBatch
    {
        public int Id { get; set; }
        public string BatchId { get; set; } = string.Empty;
        public string FlowId { get; set; } = string.Empty;
        public string FlowName { get; set; } = string.Empty;
        public string BatchName { get; set; } = string.Empty;
        public DOEDesignMethod DesignMethod { get; set; } = DOEDesignMethod.FullFactorial;
        public DOEBatchStatus Status { get; set; } = DOEBatchStatus.Designing;
        public string? DesignConfigJson { get; set; }
        public DateTime CreatedTime { get; set; } = DateTime.Now;
        public DateTime UpdatedTime { get; set; } = DateTime.Now;

        // ★ 新增: 项目关联（可空，向后兼容）
        /// <summary>
        /// 关联的优化项目 ID（null 表示独立批次，不属于任何项目）
        /// </summary>
        public string? ProjectId { get; set; }

        /// <summary>
        /// 在项目中的轮次序号（从 1 开始，仅当 ProjectId 不为 null 时有意义）
        /// </summary>
        public int? RoundNumber { get; set; }

        /// <summary>
        /// 该轮的优化阶段（仅当 ProjectId 不为 null 时有意义）
        /// </summary>
        public DOEProjectPhase? ProjectPhase { get; set; }

        // 导航属性（非数据库字段，从关联查询加载）
        public List<DOEFactor> Factors { get; set; } = new();
        public List<DOEResponse> Responses { get; set; } = new();
        public List<DOERunRecord> Runs { get; set; } = new();
        public List<DOEStopCondition> StopConditions { get; set; } = new();

        // ── 便利属性 ──

        /// <summary>是否属于某个优化项目</summary>
        public bool BelongsToProject => !string.IsNullOrEmpty(ProjectId);
    }

    /// <summary>
    /// DOE 设计方法
    /// ★ 修改: 新增 SteepestAscent / AugmentedDesign / ConfirmationRuns
    /// </summary>
    public enum DOEDesignMethod
    {
        FullFactorial,
        FractionalFactorial,
        Taguchi,
        Custom,
        CCD,
        BoxBehnken,
        DOptimal,
        PlackettBurman,         // ★ 新增: Plackett-Burman 筛选设计
        // ★ 新增: 项目迭代专用设计方法
        SteepestAscent,     // 最速上升/下降
        AugmentedDesign,    // 增强设计（在已有数据基础上补点）
        ConfirmationRuns    // 验证实验（最优点附近重复实验）
    }

    /// <summary>
    /// DOE 批次状态
    /// </summary>
    public enum DOEBatchStatus
    {
        Designing,
        Ready,
        Running,
        Paused,
        Completed,
        Aborted
    }
}