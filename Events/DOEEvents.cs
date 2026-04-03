using Prism.Events;
using MaxChemical.Modules.DOE.Models;

namespace MaxChemical.Modules.DOE.Events
{
    /// <summary>
    /// DOE 批次开始执行事件
    /// </summary>
    public class DOEBatchStartedEvent : PubSubEvent<string> { }

    /// <summary>
    /// DOE 单组实验完成事件（携带因子值 + 响应值）
    /// </summary>
    public class DOERunCompletedEvent : PubSubEvent<DOERunCompletedPayload> { }

    public class DOERunCompletedPayload
    {
        public string BatchId { get; set; } = string.Empty;
        public int RunIndex { get; set; }
        public string FactorValuesJson { get; set; } = "{}";
        public string ResponseValuesJson { get; set; } = "{}";
    }

    /// <summary>
    /// DOE 批次完成事件
    /// </summary>
    public class DOEBatchCompletedEvent : PubSubEvent<DOEBatchCompletedPayload> { }

    public class DOEBatchCompletedPayload
    {
        public string BatchId { get; set; } = string.Empty;
        public DOEBatchStatus FinalStatus { get; set; }
        public int CompletedRuns { get; set; }
        public string? StopReason { get; set; }
    }

    /// <summary>
    /// DOE 进度变化事件
    /// </summary>
    public class DOEProgressChangedEvent : PubSubEvent<DOEProgressPayload> { }

    public class DOEProgressPayload
    {
        public string BatchId { get; set; } = string.Empty;
        public int CurrentRun { get; set; }
        public int TotalRuns { get; set; }
        public string StatusMessage { get; set; } = string.Empty;
    }

    /// <summary>
    /// 请求打开 DOE 模块事件（从菜单栏/工具栏触发）
    /// </summary>
    public class OpenDOEModuleEvent : PubSubEvent { }

    /// <summary>
    /// GPR 模型状态更新事件
    /// </summary>
    public class GPRModelUpdatedEvent : PubSubEvent<GPRModelUpdatedPayload> { }

    public class GPRModelUpdatedPayload
    {
        public string FlowId { get; set; } = string.Empty;
        public bool IsActive { get; set; }
        public double? RSquared { get; set; }
        public int DataCount { get; set; }
    }
}
