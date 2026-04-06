using System.Collections.Generic;
using System.Threading.Tasks;
using MaxChemical.Modules.DOE.Models;

namespace MaxChemical.Modules.DOE.Services
{
    /// <summary>
    /// ★ 新增: 项目决策引擎接口 — 每轮结束后分析状态，推荐下一步
    /// 
    /// 职责:
    ///   1. 分析当前轮次的实验结果（OLS/GPR 模型质量、因子显著性、最优点位置）
    ///   2. 生成轮次总结 DOERoundSummary
    ///   3. 推荐下一步操作 NextStepRecommendation
    ///   4. 为下一轮生成预填的因子定义（排除已淘汰因子，继承范围）
    /// </summary>
    public interface IProjectDecisionEngine
    {
        /// <summary>
        /// 分析当前轮次结果，生成轮次总结和推荐
        /// 
        /// 在批次 Completed 后调用。内部流程:
        ///   1. 加载该批次的 OLS/GPR 分析结果
        ///   2. 评估模型质量（R², Lack-of-Fit, Power）
        ///   3. 评估因子显著性（p值 / GPR lengthscale）
        ///   4. 检测最优点是否在边界上
        ///   5. 计算 EI 改进空间
        ///   6. 根据规则推荐下一步
        /// </summary>
        Task<DOERoundSummary> AnalyzeRoundAsync(string projectId, string batchId);

        /// <summary>
        /// 获取下一轮的预填因子列表
        /// 
        /// 根据当前项目状态（活跃因子、淘汰因子、范围变更），
        /// 生成下一轮设计向导应该预填的因子定义。
        /// </summary>
        Task<List<DOEFactor>> GetNextRoundFactorsAsync(string projectId);

        /// <summary>
        /// 获取推荐的设计方法
        /// 
        /// 根据推荐的下一步操作和当前因子数，推荐最合适的设计方法。
        /// </summary>
        DOEDesignMethod RecommendDesignMethod(NextStepRecommendation recommendation,
            int activeFactorCount, DOEProjectPhase currentPhase);

        /// <summary>
        /// 执行因子淘汰
        /// 
        /// 将指定因子标记为 ScreenedOut，更新项目因子池。
        /// </summary>
        Task ScreenOutFactorsAsync(string projectId, List<string> factorNames,
            string batchId, string reason);

        /// <summary>
        /// 执行因子固定
        /// 
        /// 将指定因子固定在某个值，从后续优化中排除。
        /// </summary>
        Task FixFactorAsync(string projectId, string factorName,
            double? fixedValue, string? fixedCategoryLevel,
            string batchId, string reason);

        /// <summary>
        /// 更新因子范围
        /// 
        /// 当最优点在边界上时，扩展该因子的范围。
        /// </summary>
        Task UpdateFactorBoundsAsync(string projectId, string factorName,
            double newLower, double newUpper, string batchId, string reason);

        /// <summary>
        /// 推进项目阶段
        /// 
        /// 用户确认推荐（或选择其他操作）后，更新项目阶段。
        /// </summary>
        Task AdvancePhaseAsync(string projectId, DOEProjectPhase newPhase);
    }
}
