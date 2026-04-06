using System;
using System.Collections.Generic;
using System.Linq;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// ★ 新增: 项目级因子 — 跨批次的因子池条目
    /// 
    /// 与 DOEFactor（批次级因子）的关系:
    ///   DOEProjectFactor 是"因子池"中的一条记录，跟踪该因子在整个项目中的生命周期。
    ///   DOEFactor 是某个具体批次中该因子的快照（含该轮的范围、水平数等）。
    ///   创建新批次时，从 DOEProjectFactor 中复制活跃因子生成 DOEFactor。
    /// </summary>
    public class DOEProjectFactor
    {
        public int Id { get; set; }

        /// <summary>关联的项目 ID</summary>
        public string ProjectId { get; set; } = string.Empty;

        /// <summary>因子名称</summary>
        public string FactorName { get; set; } = string.Empty;

        /// <summary>因子类型</summary>
        public DOEFactorType FactorType { get; set; } = DOEFactorType.Continuous;

        /// <summary>因子在项目中的状态</summary>
        public ProjectFactorStatus FactorStatus { get; set; } = ProjectFactorStatus.Active;

        /// <summary>
        /// 当前下界（随轮次更新）
        /// </summary>
        public double CurrentLowerBound { get; set; }

        /// <summary>
        /// 当前上界（随轮次更新）
        /// </summary>
        public double CurrentUpperBound { get; set; }

        /// <summary>
        /// 类别因子的水平标签（逗号分隔）
        /// </summary>
        public string? CategoryLevels { get; set; }

        /// <summary>
        /// 固定值（当 FactorStatus == Fixed 时有效）
        /// </summary>
        public double? FixedValue { get; set; }

        /// <summary>
        /// 固定的类别水平（当类别因子被固定时）
        /// </summary>
        public string? FixedCategoryLevel { get; set; }

        /// <summary>
        /// 参数绑定: 来源节点 ID
        /// </summary>
        public string? SourceNodeId { get; set; }

        /// <summary>
        /// 参数绑定: 来源参数名
        /// </summary>
        public string? SourceParamName { get; set; }

        /// <summary>
        /// 淘汰/固定的原因（如 "p=0.82, 不显著" 或 "用户手动固定"）
        /// </summary>
        public string? StatusReason { get; set; }

        /// <summary>
        /// 状态变更发生在哪个批次（可追溯）
        /// </summary>
        public string? StatusChangedInBatchId { get; set; }

        /// <summary>
        /// 范围变更历史 JSON
        /// 格式: [{"batch_id": "xxx", "lower": 80, "upper": 160, "reason": "初始范围"},
        ///        {"batch_id": "yyy", "lower": 80, "upper": 200, "reason": "最优点在上界，扩展范围"}]
        /// </summary>
        public string? BoundsHistoryJson { get; set; }

        /// <summary>排序</summary>
        public int SortOrder { get; set; }

        public DateTime CreatedTime { get; set; } = DateTime.Now;
        public DateTime UpdatedTime { get; set; } = DateTime.Now;

        // ── 便利方法 ──

        /// <summary>是否活跃（参与当前优化）</summary>
        public bool IsActive => FactorStatus == ProjectFactorStatus.Active;

        /// <summary>是否类别因子</summary>
        public bool IsCategorical => FactorType == DOEFactorType.Categorical;

        /// <summary>获取类别水平列表</summary>
        public List<string> GetCategoryLevelList()
        {
            if (string.IsNullOrWhiteSpace(CategoryLevels))
                return new List<string>();
            return CategoryLevels.Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();
        }

        /// <summary>
        /// 转为批次级 DOEFactor（创建新批次时使用）
        /// </summary>
        public DOEFactor ToBatchFactor(string batchId)
        {
            return new DOEFactor
            {
                BatchId = batchId,
                FactorName = FactorName,
                FactorType = FactorType,
                LowerBound = CurrentLowerBound,
                UpperBound = CurrentUpperBound,
                LevelCount = IsCategorical ? GetCategoryLevelList().Count : 3,
                CategoryLevels = CategoryLevels,
                SourceNodeId = SourceNodeId,
                SourceParamName = SourceParamName,
                FactorSource = !string.IsNullOrEmpty(SourceNodeId)
                    ? Infrastructure.DOE.FactorSourceType.ParameterOverride
                    : Infrastructure.DOE.FactorSourceType.ParameterOverride,
                SortOrder = SortOrder
            };
        }

        /// <summary>
        /// 添加范围变更记录
        /// </summary>
        public void AddBoundsHistory(string batchId, double newLower, double newUpper, string reason)
        {
            var history = DeserializeBoundsHistory();
            history.Add(new BoundsHistoryEntry
            {
                BatchId = batchId,
                Lower = newLower,
                Upper = newUpper,
                Reason = reason,
                Timestamp = DateTime.Now
            });
            BoundsHistoryJson = Newtonsoft.Json.JsonConvert.SerializeObject(history);
            CurrentLowerBound = newLower;
            CurrentUpperBound = newUpper;
            UpdatedTime = DateTime.Now;
        }

        /// <summary>反序列化范围历史</summary>
        public List<BoundsHistoryEntry> DeserializeBoundsHistory()
        {
            if (string.IsNullOrEmpty(BoundsHistoryJson))
                return new List<BoundsHistoryEntry>();
            try
            {
                return Newtonsoft.Json.JsonConvert.DeserializeObject<List<BoundsHistoryEntry>>(BoundsHistoryJson)
                       ?? new List<BoundsHistoryEntry>();
            }
            catch { return new List<BoundsHistoryEntry>(); }
        }
    }

    /// <summary>
    /// 因子在项目中的状态
    /// </summary>
    public enum ProjectFactorStatus
    {
        /// <summary>活跃 — 参与当前优化</summary>
        Active,
        /// <summary>已筛除 — 筛选轮后标记为不显著</summary>
        ScreenedOut,
        /// <summary>已固定 — 固定在某个值，不再变化</summary>
        Fixed
    }

    /// <summary>
    /// 因子范围变更记录
    /// </summary>
    public class BoundsHistoryEntry
    {
        public string BatchId { get; set; } = string.Empty;
        public double Lower { get; set; }
        public double Upper { get; set; }
        public string Reason { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; }
    }
}
