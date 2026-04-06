using Prism.Events;
using MaxChemical.Modules.DOE.Models;

namespace MaxChemical.Modules.DOE.Events
{
    /// <summary>
    /// ★ 新增: 项目轮次完成事件 — 批次完成后触发决策分析
    /// 
    /// 发布时机: DOEBatchExecutor 在批次 Completed 后发布
    /// 订阅方: DOEProjectDashboardViewModel (弹出决策面板)
    /// </summary>
    public class DOERoundCompletedEvent : PubSubEvent<DOERoundCompletedPayload> { }

    public class DOERoundCompletedPayload
    {
        public string ProjectId { get; set; } = string.Empty;
        public string BatchId { get; set; } = string.Empty;
        public int RoundNumber { get; set; }
        public DOEBatchStatus FinalStatus { get; set; }
    }

    /// <summary>
    /// ★ 新增: 项目阶段变更事件
    /// </summary>
    public class DOEProjectPhaseChangedEvent : PubSubEvent<DOEProjectPhasePayload> { }

    public class DOEProjectPhasePayload
    {
        public string ProjectId { get; set; } = string.Empty;
        public DOEProjectPhase OldPhase { get; set; }
        public DOEProjectPhase NewPhase { get; set; }
    }

    /// <summary>
    /// ★ 新增: 请求创建下一轮批次事件
    /// 
    /// 用户在决策面板确认推荐后触发，
    /// DOEDesignWizardViewModel 订阅此事件，预填因子和设计方法。
    /// </summary>
    public class RequestNextRoundEvent : PubSubEvent<NextRoundPayload> { }

    public class NextRoundPayload
    {
        public string ProjectId { get; set; } = string.Empty;

        /// <summary>推荐的设计方法</summary>
        public DOEDesignMethod RecommendedMethod { get; set; }

        /// <summary>推荐的项目阶段</summary>
        public DOEProjectPhase RecommendedPhase { get; set; }

        /// <summary>预填的因子列表（仅活跃因子）</summary>
        public System.Collections.Generic.List<DOEFactor> PrefilledFactors { get; set; } = new();

        /// <summary>★ 新增: 预填的响应变量列表（从上一轮继承）</summary>
        public System.Collections.Generic.List<DOEResponse> PrefilledResponses { get; set; } = new();

        /// <summary>轮次序号</summary>
        public int RoundNumber { get; set; }
    }
}