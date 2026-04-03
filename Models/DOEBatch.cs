using System;
using System.Collections.Generic;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// DOE 批次主实体
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

        // 导航属性（非数据库字段，从关联查询加载）
        public List<DOEFactor> Factors { get; set; } = new();
        public List<DOEResponse> Responses { get; set; } = new();
        public List<DOERunRecord> Runs { get; set; } = new();
        public List<DOEStopCondition> StopConditions { get; set; } = new();
    }

    /// <summary>
    /// DOE 设计方法
    ///  修改: 新增 CCD / BoxBehnken / DOptimal 三个枚举值
    /// </summary>
    public enum DOEDesignMethod
    {
        FullFactorial,
        FractionalFactorial,
        Taguchi,
        Custom,
        CCD,            //  新增: 中心复合设计 (Central Composite Design)
        BoxBehnken,     //  新增: Box-Behnken 设计
        DOptimal        //  新增: D-最优设计
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