using System;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// DOE 单组实验记录
    /// </summary>
    public class DOERunRecord
    {
        public int Id { get; set; }
        public string BatchId { get; set; } = string.Empty;
        public int RunIndex { get; set; }

        /// <summary>
        /// 因子值 JSON，如 {"温度": 120, "压力": 15, "催化剂": 3.5}
        /// </summary>
        public string FactorValuesJson { get; set; } = "{}";

        /// <summary>
        /// 响应值 JSON，如 {"转化率": 87.5, "选择性": 92.1}
        /// </summary>
        public string? ResponseValuesJson { get; set; }

        /// <summary>
        /// 数据来源：measured（DOE 实测）/ imported（历史导入）
        /// </summary>
        public DOEDataSource DataSource { get; set; } = DOEDataSource.Measured;

        /// <summary>
        /// 关联的实验记录 ID（执行时由 FlowExecutionEngine 创建）
        /// </summary>
        public string? ExperimentId { get; set; }

        /// <summary>
        /// 运行状态
        /// </summary>
        public DOERunStatus Status { get; set; } = DOERunStatus.Pending;

        public DateTime? StartTime { get; set; }
        public DateTime? EndTime { get; set; }
    }

    /// <summary>
    /// 数据来源
    /// </summary>
    public enum DOEDataSource
    {
        Measured,
        Imported
    }

    /// <summary>
    /// 单组运行状态
    /// </summary>
    public enum DOERunStatus
    {
        Pending,
        Running,
        WaitingResponse,
        Completed,
        Failed,
        Skipped,
        Suspended   //  停止条件触发后，剩余未执行的组
    }
}