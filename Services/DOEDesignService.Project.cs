using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MaxChemical.Modules.DOE.Models;
using Newtonsoft.Json;
using Python.Runtime;

namespace MaxChemical.Modules.DOE.Services
{
    /// <summary>
    /// ★ 新增: DOEDesignService 的项目迭代扩展方法
    /// 
    /// 将以下方法添加到现有 DOEDesignService 类中。
    /// 不修改任何已有方法，只新增。
    /// </summary>
    public partial class DOEDesignService
    {
        // ══════════════════════════════════════════════════════
        // ★ 新增: 项目迭代专用设计方法
        // ══════════════════════════════════════════════════════

        /// <summary>
        /// 增强设计 — 在已有实验数据基础上补充新实验点
        /// </summary>
        public Task<DOEDesignMatrix> GenerateAugmentedDesignAsync(
            List<DOEFactor> factors,
            List<Dictionary<string, object>> existingPoints,
            int numAdditional = -1,
            string modelType = "quadratic")
        {
            EnsurePythonReady();
            var factorsJson = BuildFactorsJsonWithBounds(factors);
            var existingJson = JsonConvert.SerializeObject(existingPoints);

            _logger.LogInformation(
                "★ 生成增强设计: {FactorCount} 个因子, 已有 {Existing} 点, 补充 {Additional} 点",
                factors.Count, existingPoints.Count, numAdditional);

            using (Py.GIL())
            {
                string matrixJson = _pythonDesigner!.augment_design(
                    factorsJson, existingJson, numAdditional, modelType).ToString();

                var result = ParseDesignMatrix(matrixJson, factors);
                _logger.LogInformation("★ 增强设计完成: 新增 {RunCount} 组实验", result.RunCount);
                return Task.FromResult(result);
            }
        }

        /// <summary>
        /// 最速上升/下降法 — 沿一阶模型梯度方向生成实验点
        /// </summary>
        public Task<DOEDesignMatrix> GenerateSteepestAscentAsync(
            List<DOEFactor> factors,
            Dictionary<string, double> coefficients,
            int stepCount = 5,
            double stepMultiplier = 1.0,
            bool descend = false)
        {
            EnsurePythonReady();
            var factorsJson = BuildFactorsJsonWithBounds(factors);
            var coefficientsJson = JsonConvert.SerializeObject(coefficients);

            _logger.LogInformation(
                "★ 生成最速{Direction}: {FactorCount} 个因子, {Steps} 步",
                descend ? "下降" : "上升", factors.Count, stepCount);

            using (Py.GIL())
            {
                string matrixJson = _pythonDesigner!.steepest_ascent(
                    factorsJson, coefficientsJson, stepCount, stepMultiplier, descend).ToString();

                var result = ParseDesignMatrix(matrixJson, factors);
                _logger.LogInformation("★ 最速{Direction}完成: {RunCount} 组实验",
                    descend ? "下降" : "上升", result.RunCount);
                return Task.FromResult(result);
            }
        }

        /// <summary>
        /// 验证实验 — 在预测最优点附近生成重复实验
        /// </summary>
        public Task<DOEDesignMatrix> GenerateConfirmationRunsAsync(
            List<DOEFactor> factors,
            Dictionary<string, object> optimalFactors,
            int repeatCount = 5,
            double perturbation = 0.0)
        {
            EnsurePythonReady();
            var factorsJson = BuildFactorsJsonWithBounds(factors);
            var optimalJson = JsonConvert.SerializeObject(optimalFactors);

            _logger.LogInformation(
                "★ 生成验证实验: {Repeat} 次重复, 扰动={Perturbation}",
                repeatCount, perturbation);

            using (Py.GIL())
            {
                string matrixJson = _pythonDesigner!.confirmation_runs(
                    optimalJson, factorsJson, repeatCount, perturbation).ToString();

                var result = ParseDesignMatrix(matrixJson, factors);
                _logger.LogInformation("★ 验证实验完成: {RunCount} 组", result.RunCount);
                return Task.FromResult(result);
            }
        }
    }
}
