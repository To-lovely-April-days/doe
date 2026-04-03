using System;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// GPR 模型状态实体 — 按 FlowId + FactorSignature 联合唯一，支持多模型
    /// </summary>
    public class GPRModelState
    {
        public int Id { get; set; }

        /// <summary>
        /// 关联的流程 ID
        /// </summary>
        public string FlowId { get; set; } = string.Empty;

        /// <summary>
        /// 因子签名（因子名按字母排序后用逗号拼接，如 "Flow,H2Flow,Temperature"）
        /// FlowId + FactorSignature 联合唯一
        /// </summary>
        public string FactorSignature { get; set; } = string.Empty;

        /// <summary>
        /// 模型显示名称（用户可识别，如 "3因子-温压H2模型"）
        /// </summary>
        public string ModelName { get; set; } = string.Empty;

        /// <summary>
        /// 序列化的 GPR 模型（Python pickle bytes）
        /// </summary>
        public byte[]? ModelStateBytes { get; set; }

        /// <summary>
        /// 训练数据集 JSON（因子值 + 响应值 + 来源标记）
        /// </summary>
        public string? TrainingDataJson { get; set; }

        /// <summary>
        /// 超参数 JSON（核函数参数、lengthscales 等）
        /// </summary>
        public string? HyperparamsJson { get; set; }

        /// <summary>
        /// 训练数据量
        /// </summary>
        public int DataCount { get; set; }

        /// <summary>
        /// 决定系数
        /// </summary>
        public double? RSquared { get; set; }

        /// <summary>
        /// 均方根误差
        /// </summary>
        public double? RMSE { get; set; }

        /// <summary>
        /// 模型是否已激活（数据量 >= 冷启动阈值）
        /// </summary>
        public bool IsActive { get; set; }

        /// <summary>
        /// 模型演进历史 JSON
        /// </summary>
        public string? EvolutionHistoryJson { get; set; }

        /// <summary>
        /// 上次训练时间
        /// </summary>
        public DateTime? LastTrainedTime { get; set; }

        public DateTime CreatedTime { get; set; } = DateTime.Now;
        public DateTime UpdatedTime { get; set; } = DateTime.Now;

        /// <summary>
        /// 生成因子签名（静态工具方法）
        /// </summary>
        public static string BuildSignature(IEnumerable<string> factorNames)
        {
            var sorted = factorNames.OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
            return string.Join(",", sorted);
        }
    }
}
