using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Models;
using Newtonsoft.Json;

namespace MaxChemical.Modules.DOE.Services
{
    /// <summary>
    /// ★ 新增: 项目决策引擎实现
    /// 
    /// 决策规则总结:
    ///   1. 因子 ≥6 且没做过筛选        → Screening
    ///   2. 一阶模型无曲率 (LOF p>0.05)  → SteepestAscent
    ///   3. 最优点在边界上               → ExpandRange
    ///   4. R² < 0.8 或 Power < 0.6      → AugmentDesign
    ///   5. R² ≥ 0.9 且 EI 低            → ConfirmationRuns
    ///   6. 验证通过                      → Complete
    /// </summary>
    public class ProjectDecisionEngine : IProjectDecisionEngine
    {
        private readonly IDOERepository _repository;
        private readonly IGPRModelService _gprService;
        private readonly IDOEAnalysisService _analysisService;  // ★ 新增
        private readonly ILogService _logger;

        public ProjectDecisionEngine(
            IDOERepository repository,
            IGPRModelService gprService,
            IDOEAnalysisService analysisService,                // ★ 新增
            ILogService logger)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _gprService = gprService ?? throw new ArgumentNullException(nameof(gprService));
            _analysisService = analysisService ?? throw new ArgumentNullException(nameof(analysisService));
            _logger = logger?.ForContext<ProjectDecisionEngine>()
                      ?? throw new ArgumentNullException(nameof(logger));
        }

        // ══════════════ 轮次分析 ══════════════

        public async Task<DOERoundSummary> AnalyzeRoundAsync(string projectId, string batchId)
        {
            var project = await _repository.GetProjectWithDetailsAsync(projectId);
            if (project == null) throw new InvalidOperationException($"项目不存在: {projectId}");

            var batch = await _repository.GetBatchWithDetailsAsync(batchId);
            if (batch == null) throw new InvalidOperationException($"批次不存在: {batchId}");

            var completedRuns = batch.Runs
                .Where(r => r.Status == DOERunStatus.Completed)
                .ToList();

            var activeFactors = project.ProjectFactors
                .Where(f => f.IsActive)
                .ToList();

            // 构建轮次总结
            var summary = new DOERoundSummary
            {
                ProjectId = projectId,
                BatchId = batchId,
                RoundNumber = batch.RoundNumber ?? (project.CompletedRounds + 1),
                Phase = batch.ProjectPhase ?? project.CurrentPhase,
                DesignMethod = batch.DesignMethod,
                ActiveFactorCount = activeFactors.Count,
                RunCount = completedRuns.Count
            };

            // ★ 优化: 自动触发 OLS 分析（不依赖缓存）
            var primaryResponse = batch.Responses.FirstOrDefault()?.ResponseName;
            if (!string.IsNullOrEmpty(primaryResponse) && completedRuns.Count >= 6)
            {
                try
                {
                    _logger.LogInformation("轮次分析: 自动触发 OLS 拟合, Batch={BatchId}, Response={Resp}",
                        batchId, primaryResponse);

                    var olsResult = await _analysisService.FitOlsAsync(batchId, primaryResponse, "quadratic");

                    if (olsResult?.ModelSummary != null)
                    {
                        summary.RSquared = olsResult.ModelSummary.RSquared;
                        summary.RSquaredAdj = olsResult.ModelSummary.RSquaredAdj;
                        summary.RSquaredPred = olsResult.ModelSummary.RSquaredPred;
                        summary.LackOfFitP = olsResult.ModelSummary.LackOfFitP;

                        // 提取因子排名（从系数表）
                        if (olsResult.Coefficients != null)
                        {
                            var ranking = olsResult.Coefficients
                                .Where(c => c.Term != "Intercept")
                                .Select(c => new
                                {
                                    name = c.Term,
                                    p_value = c.PValue,
                                    significant = c.PValue < 0.05
                                })
                                .OrderBy(x => x.p_value)
                                .ToList();
                            summary.FactorRankingJson = JsonConvert.SerializeObject(ranking);
                        }

                        // 缓存 OLS 结果供模型分析页使用
                        var resultJson = JsonConvert.SerializeObject(olsResult);
                        await _repository.SaveOlsResultAsync(batchId, primaryResponse, resultJson);

                        _logger.LogInformation("轮次 OLS 自动分析完成: R²={R2:F4}", summary.RSquared);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "自动 OLS 分析失败，继续使用其他指标");
                }
            }

            // GPR 模型指标
            if (_gprService.IsActive)
            {
                try
                {
                    PopulateFromGprModel(summary, activeFactors, batch);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "获取 GPR 指标失败，跳过");
                }
            }

            // 计算最优点和边界状态
            ComputeOptimalAndBoundary(summary, completedRuns, activeFactors, primaryResponse);

            // 更新项目最优
            if (summary.BestResponseValue.HasValue)
            {
                var allExperiments = project.TotalExperiments + completedRuns.Count;
                var roundNumber = batch.RoundNumber ?? (project.CompletedRounds + 1);
                await _repository.UpdateProjectBestAsync(
                    projectId,
                    summary.BestResponseValue.Value,
                    summary.BestFactorsJson ?? "{}",
                    allExperiments,
                    roundNumber);
            }

            // 生成推荐
            summary.Recommendation = GenerateRecommendation(summary, project, activeFactors);
            summary.RecommendationReason = GenerateRecommendationReason(summary);

            // 持久化
            await _repository.SaveRoundSummaryAsync(summary);

            _logger.LogInformation(
                "轮次分析完成: Project={ProjectId}, Round={Round}, R²={R2}, 推荐={Rec}",
                projectId, summary.RoundNumber, summary.RSquared, summary.Recommendation);

            return summary;
        }

        // ══════════════ 下一轮因子 ══════════════

        public async Task<List<DOEFactor>> GetNextRoundFactorsAsync(string projectId)
        {
            var activeFactors = await _repository.GetActiveProjectFactorsAsync(projectId);

            return activeFactors.Select(pf => pf.ToBatchFactor("")).ToList();
        }

        // ══════════════ 推荐设计方法 ══════════════

        public DOEDesignMethod RecommendDesignMethod(NextStepRecommendation recommendation,
            int activeFactorCount, DOEProjectPhase currentPhase)
        {
            return recommendation switch
            {
                NextStepRecommendation.ContinueScreening =>
                    activeFactorCount > 7
                        ? DOEDesignMethod.FractionalFactorial
                        : DOEDesignMethod.Taguchi,

                NextStepRecommendation.SteepestAscent =>
                    DOEDesignMethod.SteepestAscent,

                NextStepRecommendation.StartRSM =>
                    activeFactorCount >= 3
                        ? DOEDesignMethod.BoxBehnken
                        : DOEDesignMethod.CCD,

                NextStepRecommendation.AugmentDesign =>
                    DOEDesignMethod.AugmentedDesign,

                NextStepRecommendation.ExpandRange =>
                    activeFactorCount <= 4
                        ? DOEDesignMethod.CCD
                        : DOEDesignMethod.DOptimal,

                NextStepRecommendation.ConfirmationRuns =>
                    DOEDesignMethod.ConfirmationRuns,

                _ => DOEDesignMethod.DOptimal
            };
        }

        // ══════════════ 因子管理 ══════════════

        public async Task ScreenOutFactorsAsync(string projectId, List<string> factorNames,
            string batchId, string reason)
        {
            var factors = await _repository.GetProjectFactorsAsync(projectId);
            foreach (var pf in factors.Where(f => factorNames.Contains(f.FactorName)))
            {
                await _repository.UpdateProjectFactorStatusAsync(
                    pf.Id, ProjectFactorStatus.ScreenedOut, reason, batchId);
            }

            _logger.LogInformation("因子淘汰: Project={ProjectId}, Factors=[{Factors}], Reason={Reason}",
                projectId, string.Join(",", factorNames), reason);
        }

        public async Task FixFactorAsync(string projectId, string factorName,
            double? fixedValue, string? fixedCategoryLevel,
            string batchId, string reason)
        {
            var factors = await _repository.GetProjectFactorsAsync(projectId);
            var pf = factors.FirstOrDefault(f => f.FactorName == factorName);
            if (pf == null) return;

            pf.FactorStatus = ProjectFactorStatus.Fixed;
            pf.FixedValue = fixedValue;
            pf.FixedCategoryLevel = fixedCategoryLevel;
            pf.StatusReason = reason;
            pf.StatusChangedInBatchId = batchId;

            await _repository.UpdateProjectFactorStatusAsync(
                pf.Id, ProjectFactorStatus.Fixed, reason, batchId);

            _logger.LogInformation("因子固定: Project={ProjectId}, Factor={Factor}, Value={Value}",
                projectId, factorName, fixedValue ?? (object?)fixedCategoryLevel ?? "null");
        }

        public async Task UpdateFactorBoundsAsync(string projectId, string factorName,
            double newLower, double newUpper, string batchId, string reason)
        {
            var factors = await _repository.GetProjectFactorsAsync(projectId);
            var pf = factors.FirstOrDefault(f => f.FactorName == factorName);
            if (pf == null) return;

            pf.AddBoundsHistory(batchId, newLower, newUpper, reason);
            await _repository.UpdateProjectFactorBoundsAsync(
                pf.Id, newLower, newUpper, pf.BoundsHistoryJson ?? "[]");

            _logger.LogInformation("因子范围更新: Factor={Factor}, [{Lower}, {Upper}], Reason={Reason}",
                factorName, newLower, newUpper, reason);
        }

        public async Task AdvancePhaseAsync(string projectId, DOEProjectPhase newPhase)
        {
            await _repository.UpdateProjectPhaseAsync(projectId, newPhase);
            _logger.LogInformation("项目阶段推进: Project={ProjectId}, Phase={Phase}", projectId, newPhase);
        }

        // ══════════════ 内部方法 ══════════════

        private void PopulateFromOlsResult(DOERoundSummary summary, string olsJson,
            List<DOEProjectFactor> activeFactors)
        {
            var olsResult = JsonConvert.DeserializeObject<Dictionary<string, object>>(olsJson);
            if (olsResult == null) return;

            // 提取模型摘要
            if (olsResult.TryGetValue("model_summary", out var summaryObj))
            {
                var ms = JsonConvert.DeserializeObject<Dictionary<string, double>>(summaryObj.ToString()!);
                if (ms != null)
                {
                    summary.RSquared = ms.GetValueOrDefault("r_squared");
                    summary.RSquaredAdj = ms.GetValueOrDefault("r_squared_adj");
                    summary.RSquaredPred = ms.GetValueOrDefault("r_squared_pred");
                    summary.LackOfFitP = ms.GetValueOrDefault("lack_of_fit_p");
                }
            }

            // 提取因子排名
            if (olsResult.TryGetValue("coefficients", out var coeffsObj))
            {
                try
                {
                    var coeffs = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(coeffsObj.ToString()!);
                    if (coeffs != null)
                    {
                        var ranking = coeffs
                            .Where(c => c.TryGetValue("term", out var t) && t?.ToString() != "Intercept")
                            .Select(c => new
                            {
                                name = c.GetValueOrDefault("term")?.ToString() ?? "",
                                p_value = Convert.ToDouble(c.GetValueOrDefault("p_value", 1.0)),
                                significant = Convert.ToDouble(c.GetValueOrDefault("p_value", 1.0)) < 0.05
                            })
                            .OrderBy(x => x.p_value)
                            .ToList();

                        summary.FactorRankingJson = JsonConvert.SerializeObject(ranking);
                    }
                }
                catch { /* 忽略解析失败 */ }
            }
        }

        private void PopulateFromGprModel(DOERoundSummary summary,
            List<DOEProjectFactor> activeFactors, DOEBatch batch)
        {
            // GPR R²: 从演进历史中获取最新值
            var sensitivity = _gprService.GetSensitivity();
            if (sensitivity.Count > 0 && string.IsNullOrEmpty(summary.FactorRankingJson))
            {
                var ranking = sensitivity
                    .OrderByDescending(kv => kv.Value)
                    .Select(kv => new
                    {
                        name = kv.Key,
                        sensitivity = kv.Value,
                        significant = kv.Value > 0.1  // lengthscale 归一化后 > 0.1 视为显著
                    })
                    .ToList();

                summary.FactorRankingJson = JsonConvert.SerializeObject(ranking);
            }
        }

        private void ComputeOptimalAndBoundary(DOERoundSummary summary,
            List<DOERunRecord> completedRuns, List<DOEProjectFactor> activeFactors,
            string? primaryResponse)
        {
            if (completedRuns.Count == 0 || string.IsNullOrEmpty(primaryResponse)) return;

            // 找最优组
            double bestVal = double.MinValue;
            Dictionary<string, object>? bestFactors = null;

            foreach (var run in completedRuns)
            {
                if (string.IsNullOrEmpty(run.ResponseValuesJson)) continue;
                var responses = JsonConvert.DeserializeObject<Dictionary<string, double>>(run.ResponseValuesJson);
                if (responses == null || !responses.TryGetValue(primaryResponse, out var val)) continue;

                if (val > bestVal)
                {
                    bestVal = val;
                    bestFactors = JsonConvert.DeserializeObject<Dictionary<string, object>>(run.FactorValuesJson);
                }
            }

            if (bestFactors != null)
            {
                summary.BestResponseValue = bestVal;
                summary.BestFactorsJson = JsonConvert.SerializeObject(bestFactors);

                // 检测边界状态
                var boundaryStatus = new Dictionary<string, string>();
                foreach (var pf in activeFactors.Where(f => !f.IsCategorical))
                {
                    if (!bestFactors.TryGetValue(pf.FactorName, out var valObj)) continue;
                    if (!double.TryParse(valObj.ToString(), out var v)) continue;

                    double range = pf.CurrentUpperBound - pf.CurrentLowerBound;
                    double tolerance = range * 0.05;  // 5% 容差

                    if (Math.Abs(v - pf.CurrentLowerBound) < tolerance)
                        boundaryStatus[pf.FactorName] = "at_lower";
                    else if (Math.Abs(v - pf.CurrentUpperBound) < tolerance)
                        boundaryStatus[pf.FactorName] = "at_upper";
                    else
                        boundaryStatus[pf.FactorName] = "interior";
                }

                summary.OptimalBoundaryStatusJson = JsonConvert.SerializeObject(boundaryStatus);
            }
        }

        private NextStepRecommendation GenerateRecommendation(DOERoundSummary summary,
            DOEProject project, List<DOEProjectFactor> activeFactors)
        {
            // 规则 1: 验证阶段已完成 → Complete
            if (summary.Phase == DOEProjectPhase.Confirmation)
            {
                // 简单判断: 验证实验完成即可确认
                return NextStepRecommendation.Complete;
            }

            // 规则 2: 最优点在边界上 → ExpandRange
            if (!string.IsNullOrEmpty(summary.OptimalBoundaryStatusJson))
            {
                var status = JsonConvert.DeserializeObject<Dictionary<string, string>>(
                    summary.OptimalBoundaryStatusJson);
                if (status != null && status.Values.Any(v => v == "at_lower" || v == "at_upper"))
                {
                    return NextStepRecommendation.ExpandRange;
                }
            }

            // 规则 3: 因子多且在筛选阶段 → 继续筛选
            if (activeFactors.Count >= 6 && summary.Phase == DOEProjectPhase.Screening)
            {
                // 检查是否有不显著因子可以淘汰
                if (!string.IsNullOrEmpty(summary.FactorRankingJson))
                {
                    return NextStepRecommendation.ContinueScreening;
                }
            }

            // 规则 4: 模型质量不好 → 补点
            if (summary.RSquared.HasValue && summary.RSquared < 0.80)
            {
                return NextStepRecommendation.AugmentDesign;
            }

            // 规则 5: 一阶模型无曲率 (LOF 不显著) 且处于筛选后阶段 → 爬坡
            if (summary.LackOfFitP.HasValue && summary.LackOfFitP > 0.05
                && summary.Phase <= DOEProjectPhase.PathSearch
                && activeFactors.Count <= 5)
            {
                return NextStepRecommendation.SteepestAscent;
            }

            // 规则 6: 模型好且 EI 低 → 验证
            if (summary.RSquared.HasValue && summary.RSquared >= 0.90)
            {
                if (summary.MaxEI.HasValue && summary.MaxEI < 0.01)
                {
                    return NextStepRecommendation.ConfirmationRuns;
                }
            }

            // 规则 7: 因子少，还没做 RSM → RSM
            if (activeFactors.Count <= 5 && summary.Phase < DOEProjectPhase.RSM)
            {
                return NextStepRecommendation.StartRSM;
            }

            // 规则 8: 模型还行但可以改善 → 补点
            if (summary.RSquared.HasValue && summary.RSquared < 0.90)
            {
                return NextStepRecommendation.AugmentDesign;
            }

            // 默认: 让用户决定
            return NextStepRecommendation.UserDecision;
        }

        private string GenerateRecommendationReason(DOERoundSummary summary)
        {
            return summary.Recommendation switch
            {
                NextStepRecommendation.ContinueScreening =>
                    $"当前 {summary.ActiveFactorCount} 个因子，建议继续筛选，淘汰不显著因子后再进入 RSM。",

                NextStepRecommendation.SteepestAscent =>
                    $"一阶模型 Lack-of-Fit p={summary.LackOfFitP:F3} > 0.05（无显著曲率），" +
                    $"建议沿最速上升方向移动到更优区域。",

                NextStepRecommendation.StartRSM =>
                    $"{summary.ActiveFactorCount} 个活跃因子，适合使用 CCD 或 Box-Behnken 做响应面优化。",

                NextStepRecommendation.AugmentDesign =>
                    $"模型 R²={summary.RSquared:F3}，拟合质量不足，建议在已有数据基础上补充实验点。",

                NextStepRecommendation.ExpandRange =>
                    "当前最优点位于因子边界上，建议扩展该因子的范围后重新优化。",

                NextStepRecommendation.ConfirmationRuns =>
                    $"模型 R²={summary.RSquared:F3}，EI={summary.MaxEI:E2}，改进空间很小，" +
                    $"建议在预测最优点做验证实验确认。",

                NextStepRecommendation.Complete =>
                    "验证实验已完成，优化目标达成。",

                _ => "系统无法自动判断，请根据专业经验选择下一步操作。"
            };
        }
    }
}