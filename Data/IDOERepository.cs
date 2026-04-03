using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using MaxChemical.Modules.DOE.Models;

namespace MaxChemical.Modules.DOE.Data
{
    /// <summary>
    /// DOE 数据仓储接口 — 6 张表的 CRUD + 复合查询
    /// </summary>
    public interface IDOERepository
    {
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

        /// <summary>
        /// 获取某流程下所有批次的汇总信息（用于历史列表展示）
        /// </summary>
        Task<List<DOEBatchSummary>> GetBatchSummariesByFlowAsync(string flowId);

        /// <summary>
        /// 获取所有批次摘要（不限流程，按时间倒序，带 Runs 统计）
        /// </summary>
        Task<List<DOEBatchSummary>> GetAllBatchSummariesAsync(int limit = 50);

        /// <summary>
        /// 获取某流程下所有实测+导入的数据（用于 GPR 训练）
        /// </summary>
        Task<List<DOERunRecord>> GetAllTrainingDataByFlowAsync(string flowId);

        /// <summary>
        /// 按 FlowId + FactorSignature 获取模型
        /// </summary>
        Task<GPRModelState?> GetGPRModelStateAsync(string flowId, string factorSignature);

        /// <summary>
        /// 获取某个流程的所有 GPR 模型
        /// </summary>
        Task<List<GPRModelState>> GetGPRModelsByFlowAsync(string flowId);

        /// <summary>
        /// 保存/更新 GPR 模型（按 FlowId + FactorSignature upsert）
        /// </summary>
        Task SaveGPRModelStateAsync(GPRModelState state);

        /// <summary>
        /// 删除指定模型
        /// </summary>
        Task DeleteGPRModelAsync(int modelId);

        /// <summary>
        /// 查询同 FlowId + FactorSignature 下有多少个方案（用于删除时判断模型是否还被其他方案使用）
        /// </summary>
        Task<int> GetBatchCountBySignatureAsync(string flowId, string factorSignature);

        /// <summary>
        /// 删除批次及所有子表数据（因子、响应、Runs、停止条件）
        /// </summary>
        Task DeleteBatchWithChildrenAsync(string batchId);

        /// <summary>
        /// 获取所有 GPR 模型（不限流程）
        /// </summary>
        Task<List<GPRModelState>> GetAllGPRModelsAsync();

        Task<GPRModelState?> GetGPRModelByIdAsync(int modelId);

        // ──  新增: Desirability 配置持久化 ──

        /// <summary>
        ///  新增: 保存 Desirability 配置 JSON
        /// </summary>
        /// <param name="batchId">批次 ID</param>
        /// <param name="configJson">配置 JSON 字符串</param>
        Task SaveDesirabilityConfigAsync(string batchId, string configJson);

        /// <summary>
        ///  新增: 获取 Desirability 配置 JSON
        /// </summary>
        Task<string?> GetDesirabilityConfigJsonAsync(string batchId);

        /// <summary>
        ///  新增: 保存 OLS 分析结果 JSON
        /// </summary>
        Task SaveOlsResultAsync(string batchId, string responseName, string resultJson);

        /// <summary>
        ///  新增: 获取 OLS 分析结果 JSON
        /// </summary>
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

        /// <summary>进度显示文本（XAML 绑定用）</summary>
        public string ProgressText => $"{CompletedRuns}/{TotalRuns}";

        /// <summary>是否可编辑（仅 Ready 状态）</summary>
        public bool CanEdit => Status == DOEBatchStatus.Ready;
    }
}