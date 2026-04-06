using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using MaxChemical.Modules.DOE.Models;

namespace MaxChemical.Modules.DOE.Data
{
    /// <summary>
    /// DOE 数据仓储接口 — 6 张表的 CRUD + 复合查询
    /// 
    /// ★ 修改: 新增 Project / ProjectFactor / RoundSummary 相关方法
    ///   所有原有方法签名不变，保持完全兼容。
    /// </summary>
    public interface IDOERepository
    {
        // ══════════════════════════════════════════════════════
        // ★ 新增: Project 项目
        // ══════════════════════════════════════════════════════

        /// <summary>创建项目</summary>
        Task<string> CreateProjectAsync(DOEProject project);

        /// <summary>获取项目（不含子实体）</summary>
        Task<DOEProject?> GetProjectAsync(string projectId);

        /// <summary>获取项目（含 ProjectFactors + Batches + RoundSummaries）</summary>
        Task<DOEProject?> GetProjectWithDetailsAsync(string projectId);

        /// <summary>获取所有活跃项目摘要</summary>
        Task<List<DOEProjectSummary>> GetActiveProjectsAsync();

        /// <summary>获取所有项目摘要（含已完成和归档）</summary>
        Task<List<DOEProjectSummary>> GetAllProjectSummariesAsync(int limit = 50);

        /// <summary>更新项目基础信息</summary>
        Task UpdateProjectAsync(DOEProject project);

        /// <summary>更新项目阶段</summary>
        Task UpdateProjectPhaseAsync(string projectId, DOEProjectPhase phase);

        /// <summary>更新项目状态</summary>
        Task UpdateProjectStatusAsync(string projectId, DOEProjectStatus status);

        /// <summary>更新项目最优值和统计</summary>
        Task UpdateProjectBestAsync(string projectId, double bestValue, string bestFactorsJson, int totalExperiments, int completedRounds);

        /// <summary>删除项目及所有关联数据</summary>
        Task DeleteProjectWithChildrenAsync(string projectId);

        // ══════════════════════════════════════════════════════
        // ★ 新增: ProjectFactor 项目因子池
        // ══════════════════════════════════════════════════════

        /// <summary>保存项目因子列表（覆盖写入）</summary>
        Task SaveProjectFactorsAsync(string projectId, List<DOEProjectFactor> factors);

        /// <summary>获取项目所有因子</summary>
        Task<List<DOEProjectFactor>> GetProjectFactorsAsync(string projectId);

        /// <summary>获取项目活跃因子</summary>
        Task<List<DOEProjectFactor>> GetActiveProjectFactorsAsync(string projectId);

        /// <summary>更新单个因子状态</summary>
        Task UpdateProjectFactorStatusAsync(int factorId, ProjectFactorStatus status,
            string? reason = null, string? batchId = null);

        /// <summary>更新因子范围</summary>
        Task UpdateProjectFactorBoundsAsync(int factorId, double lower, double upper,
            string boundsHistoryJson);

        // ══════════════════════════════════════════════════════
        // ★ 新增: RoundSummary 轮次总结
        // ══════════════════════════════════════════════════════

        /// <summary>保存轮次总结</summary>
        Task SaveRoundSummaryAsync(DOERoundSummary summary);

        /// <summary>获取项目所有轮次总结</summary>
        Task<List<DOERoundSummary>> GetRoundSummariesAsync(string projectId);

        /// <summary>获取特定批次的轮次总结</summary>
        Task<DOERoundSummary?> GetRoundSummaryByBatchAsync(string batchId);

        /// <summary>更新用户决策</summary>
        Task UpdateRoundDecisionAsync(int summaryId, NextStepRecommendation decision, string? notes);

        // ══════════════════════════════════════════════════════
        // ★ 新增: 按项目查询批次和数据
        // ══════════════════════════════════════════════════════

        /// <summary>获取项目下所有批次</summary>
        Task<List<DOEBatch>> GetBatchesByProjectAsync(string projectId);

        /// <summary>获取项目下所有已完成实验数据（用于跨批次 GPR 训练）</summary>
        Task<List<DOERunRecord>> GetAllCompletedRunsByProjectAsync(string projectId);

        /// <summary>
        /// 按项目 ID 查找 GPR 模型
        /// </summary>
        Task<GPRModelState?> GetGPRModelByProjectAsync(string projectId, string factorSignature);

        // ══════════════════════════════════════════════════════
        // 原有接口（完全不变）
        // ══════════════════════════════════════════════════════

        // ── Batch ──────────────────────────────────

        Task<string> CreateBatchAsync(DOEBatch batch);
        Task<DOEBatch?> GetBatchAsync(string batchId);
        Task<DOEBatch?> GetBatchWithDetailsAsync(string batchId);
        Task<List<DOEBatch>> GetBatchesByFlowAsync(string flowId);
        Task<List<DOEBatch>> GetRecentBatchesAsync(int count = 20);
        Task UpdateBatchStatusAsync(string batchId, DOEBatchStatus status);
        Task UpdateBatchAsync(DOEBatch batch);
        Task DeleteBatchAsync(string batchId);

        // ── Factors ────────────────────────────────

        Task SaveFactorsAsync(string batchId, List<DOEFactor> factors);
        Task<List<DOEFactor>> GetFactorsAsync(string batchId);

        // ── Responses ──────────────────────────────

        Task SaveResponsesAsync(string batchId, List<DOEResponse> responses);
        Task<List<DOEResponse>> GetResponsesAsync(string batchId);

        // ── Runs ───────────────────────────────────

        Task SaveRunAsync(DOERunRecord run);
        Task SaveRunsAsync(List<DOERunRecord> runs);
        Task UpdateRunAsync(DOERunRecord run);
        Task<DOERunRecord?> GetRunAsync(string batchId, int runIndex);
        Task<List<DOERunRecord>> GetRunsAsync(string batchId);
        Task<List<DOERunRecord>> GetCompletedRunsAsync(string batchId);

        // ── Stop Conditions ────────────────────────

        Task SaveStopConditionsAsync(string batchId, List<DOEStopCondition> conditions);
        Task<List<DOEStopCondition>> GetStopConditionsAsync(string batchId);

        // ── GPR Model State ────────────────────────

        Task<GPRModelState?> GetGPRModelStateAsync(string flowId);
        Task UpdateGPRModelStateAsync(GPRModelState state);
        Task DeleteGPRModelStateAsync(string flowId);

        // ── 复合查询 ──────────────────────────────

        Task<List<DOEBatchSummary>> GetBatchSummariesByFlowAsync(string flowId);
        Task<List<DOEBatchSummary>> GetAllBatchSummariesAsync(int limit = 50);
        Task<List<DOERunRecord>> GetAllTrainingDataByFlowAsync(string flowId);
        Task<GPRModelState?> GetGPRModelStateAsync(string flowId, string factorSignature);
        Task<List<GPRModelState>> GetGPRModelsByFlowAsync(string flowId);
        Task SaveGPRModelStateAsync(GPRModelState state);
        Task DeleteGPRModelAsync(int modelId);
        Task<int> GetBatchCountBySignatureAsync(string flowId, string factorSignature);
        Task DeleteBatchWithChildrenAsync(string batchId);
        Task<List<GPRModelState>> GetAllGPRModelsAsync();
        Task<GPRModelState?> GetGPRModelByIdAsync(int modelId);

        // ── Desirability ──────────────────────────

        Task SaveDesirabilityConfigAsync(string batchId, string configJson);
        Task<string?> GetDesirabilityConfigJsonAsync(string batchId);
        Task SaveOlsResultAsync(string batchId, string responseName, string resultJson);
        Task<string?> GetOlsResultJsonAsync(string batchId, string responseName);
    }

    /// <summary>
    /// 批次摘要信息（列表展示用）
    /// </summary>
    public class DOEBatchSummary
    {
        public string BatchId { get; set; } = string.Empty;
        public string BatchName { get; set; } = string.Empty;
        public string FlowName { get; set; } = string.Empty;
        public DOEDesignMethod DesignMethod { get; set; }
        public DOEBatchStatus Status { get; set; }
        public int TotalRuns { get; set; }
        public int CompletedRuns { get; set; }
        public double? BestResponseValue { get; set; }
        public DateTime CreatedTime { get; set; }

        // ★ 新增: 项目关联
        public string? ProjectId { get; set; }
        public int? RoundNumber { get; set; }

        /// <summary>进度显示文本（XAML 绑定用）</summary>
        public string ProgressText => $"{CompletedRuns}/{TotalRuns}";

        /// <summary>是否可编辑（仅 Ready 状态）</summary>
        public bool CanEdit => Status == DOEBatchStatus.Ready;
    }

    /// <summary>
    /// ★ 新增: 项目摘要信息（列表展示用）
    /// </summary>
    public class DOEProjectSummary
    {
        public string ProjectId { get; set; } = string.Empty;
        public string ProjectName { get; set; } = string.Empty;
        public string? FlowName { get; set; }
        public DOEProjectPhase CurrentPhase { get; set; }
        public DOEProjectStatus Status { get; set; }
        public int TotalBatches { get; set; }
        public int TotalExperiments { get; set; }
        public double? BestResponseValue { get; set; }
        public DateTime CreatedTime { get; set; }
        public DateTime UpdatedTime { get; set; }

        /// <summary>阶段显示文本</summary>
        public string PhaseText => CurrentPhase switch
        {
            DOEProjectPhase.Screening => "筛选",
            DOEProjectPhase.PathSearch => "路径探索",
            DOEProjectPhase.RSM => "响应面优化",
            DOEProjectPhase.Augmenting => "增强补点",
            DOEProjectPhase.Confirmation => "验证",
            DOEProjectPhase.Completed => "已完成",
            _ => CurrentPhase.ToString()
        };
    }
}