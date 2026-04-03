using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using MaxChemical.Modules.DOE.Models;

namespace MaxChemical.Modules.DOE.Services
{
    // ═══════════════════════════════════════════════════════════════════
    // IMultiResponseGPRService — 多响应 GPR 管理器
    //
    //  新增文件
    // 
    // 设计思路:
    //   原来的 IGPRModelService 管理「单个 GPR 模型」，用于主响应的在线学习和预测。
    //   IMultiResponseGPRService 在其上层管理「一组 GPR 模型」，每个响应变量一个。
    //
    //   关键约束:
    //   1. 不修改 IGPRModelService 接口 — 所有现有调用方（BatchExecutor、ModelAnalysisVM 等）零改动
    //   2. 不修改 doe_gpr.py — Python 层已经是单响应模型，天然适配
    //   3. GPRModelService 仍然是主响应的「主模型」，注册为 IGPRModelService 单例
    //   4. MultiResponseGPRService 内部为每个非主响应创建额外的 GPRModelService 实例
    //
    //   这种设计让改动面最小:
    //   - BatchExecutor 的 FeedGPRModelAsync 改为调用 MultiResponseGPRService.AppendAllResponses
    //   - DesirabilityService.OptimizeAsync 改为调用 MultiResponseGPRService.PredictAll
    //   - 其他所有地方（停止条件、ModelAnalysis 页面等）完全不变
    // ═══════════════════════════════════════════════════════════════════

    /// <summary>
    ///  新增: 多响应 GPR 管理器 — 为每个响应变量维护独立的 GPR 模型
    /// </summary>
    public interface IMultiResponseGPRService
    {
        /// <summary>
        /// 初始化多响应模型集合。
        /// 主响应使用已有的 IGPRModelService（保持兼容），
        /// 其他响应各自创建独立的 GPR 模型实例。
        /// </summary>
        /// <param name="flowId">流程 ID</param>
        /// <param name="factors">因子列表</param>
        /// <param name="responses">响应变量列表（第一个为主响应）</param>
        Task InitializeAsync(string flowId, List<DOEFactor> factors, List<DOEResponse> responses);

        /// <summary>
        /// 向所有响应模型追加一组实验数据。
        /// ★ 修复 (v3): factorValues 改为 object 以支持类别因子标签
        /// </summary>
        void AppendAllResponses(Dictionary<string, object> factorValues,
                                Dictionary<string, double> responseValues,
                                string source = "measured",
                                string batchName = "",
                                string timestamp = "");

        /// <summary>
        /// 训练所有响应模型。
        /// 返回每个响应的训练结果。
        /// </summary>
        Task<Dictionary<string, GPRTrainResult>> TrainAllAsync();

        /// <summary>
        /// 对给定因子组合，预测所有响应的值。
        /// ★ 修复 (v3): factorValues 改为 object 以支持类别因子标签
        /// </summary>
        Dictionary<string, GPRPrediction> PredictAll(Dictionary<string, object> factorValues);

        /// <summary>
        /// 保存所有模型状态到数据库。
        /// </summary>
        Task SaveAllAsync(string flowId);

        /// <summary>
        /// 所有模型中是否至少有一个已激活。
        /// </summary>
        bool AnyModelActive { get; }

        /// <summary>
        /// 所有模型是否全部激活（Desirability 优化的前置条件）。
        /// </summary>
        bool AllModelsActive { get; }

        /// <summary>
        /// 获取各响应模型的状态摘要（UI 展示用）。
        /// </summary>
        List<ResponseModelStatus> GetModelStatuses();

        /// <summary>
        /// 获取已初始化的响应变量名列表。
        /// </summary>
        List<string> ResponseNames { get; }

        Dictionary<string, GPRPrediction> PredictAllInsideGIL(Dictionary<string, object> factorValues);
    }

    /// <summary>
    ///  新增: 单个响应模型的状态摘要
    /// </summary>
    public class ResponseModelStatus
    {
        public string ResponseName { get; set; } = string.Empty;
        public bool IsActive { get; set; }
        public int DataCount { get; set; }
        public double? RSquared { get; set; }
        public bool IsPrimary { get; set; }
    }
}