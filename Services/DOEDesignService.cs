using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Models;
using Newtonsoft.Json;
using Python.Runtime;

namespace MaxChemical.Modules.DOE.Services
{
    /// <summary>
    /// DOE 实验设计服务 — 调用 Python doe_engine.py 生成参数矩阵
    /// 
    ///  修改: 新增 GenerateCCDAsync / GenerateBoxBehnkenAsync / GenerateDOptimalAsync / GetDesignQualityAsync
    /// </summary>
    public partial class DOEDesignService : IDOEDesignService
    {
        private readonly IDOERepository _repository;
        private readonly ILogService _logger;
        private dynamic? _pythonDesigner;
        private dynamic? _pythonImporter;
        private bool _isPythonInitialized;

        public DOEDesignService(IDOERepository repository, ILogService logger)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _logger = logger?.ForContext<DOEDesignService>() ?? throw new ArgumentNullException(nameof(logger));
        }

        private void EnsurePythonReady()
        {
            if (_isPythonInitialized && _pythonDesigner != null) return;

            try
            {
                var envManager = MaxChemical.Modules.Designer.Services.PythonEnvironmentManager.Instance;
                if (!envManager.IsInitialized)
                {
                    var success = envManager.Initialize();
                    if (!success)
                        throw new InvalidOperationException("Python 环境初始化失败");
                }

                using (Py.GIL())
                {
                    dynamic doeEngine = Py.Import("doe_engine");
                    _pythonDesigner = doeEngine.create_designer();
                    _pythonImporter = doeEngine.create_importer();
                }

                _isPythonInitialized = true;
                _logger.LogInformation("DOE 设计服务 Python 模块初始化成功");
            }
            catch (Exception ex)
            {
                _isPythonInitialized = false;
                _logger.LogError(ex, "DOE 设计服务 Python 模块初始化失败");
                throw;
            }
        }

        // ══════════════ 原有方法（保留） ══════════════

        public Task<DOEDesignMatrix> GenerateFullFactorialAsync(List<DOEFactor> factors)
        {
            EnsurePythonReady();
            var factorsJson = BuildFactorsJsonWithLevels(factors);
            _logger.LogInformation("生成全因子设计: {FactorCount} 个因子", factors.Count);
            using (Py.GIL())
            {
                string matrixJson = _pythonDesigner!.full_factorial(factorsJson).ToString();
                var result = ParseDesignMatrix(matrixJson, factors);
                return Task.FromResult(result);
            }
        }

        public Task<DOEDesignMatrix> GenerateFractionalFactorialAsync(List<DOEFactor> factors, int resolution)
        {
            EnsurePythonReady();
            var factorsJson = BuildFactorsJsonWithBounds(factors);
            _logger.LogInformation("生成部分因子设计: {FactorCount} 个因子, 分辨度 {Resolution}", factors.Count, resolution);
            using (Py.GIL())
            {
                string matrixJson = _pythonDesigner!.fractional_factorial(factorsJson, resolution).ToString();
                var result = ParseDesignMatrix(matrixJson, factors);
                return Task.FromResult(result);
            }
        }

        public Task<DOEDesignMatrix> GenerateTaguchiAsync(List<DOEFactor> factors, string tableType)
        {
            EnsurePythonReady();
            var factorsJson = BuildFactorsJsonWithLevels(factors);
            _logger.LogInformation("生成正交设计: {FactorCount} 个因子, 表类型 {TableType}", factors.Count, tableType);
            using (Py.GIL())
            {
                string matrixJson = _pythonDesigner!.taguchi(factorsJson, tableType).ToString();
                var result = ParseDesignMatrix(matrixJson, factors);
                return Task.FromResult(result);
            }
        }

        public Task<DOEDesignMatrix> ImportCustomMatrixAsync(string excelPath, List<DOEFactor> factors)
        {
            EnsurePythonReady();
            var factorNames = factors.Select(f => f.FactorName).ToList();
            var factorNamesJson = JsonConvert.SerializeObject(factorNames);
            var emptyResponsesJson = "[]";
            _logger.LogInformation("导入自定义参数矩阵: {Path}", excelPath);
            using (Py.GIL())
            {
                string dataJson = _pythonImporter!.import_excel(excelPath, factorNamesJson, emptyResponsesJson).ToString();
                var records = JsonConvert.DeserializeObject<List<ImportedRecord>>(dataJson) ?? new();
                var matrix = new DOEDesignMatrix
                {
                    FactorNames = factorNames,
                    // ★ 修复 (Bug#4): ImportedRecord.Factors 已是 Dict<string,object>，直接使用
                    Rows = records.Select(r => new Dictionary<string, object>(r.Factors)).ToList()
                };
                return Task.FromResult(matrix);
            }
        }

        public Task<DataValidationResult> ValidateImportedDataAsync(string excelPath, List<DOEFactor> factors, List<DOEResponse> responses)
        {
            EnsurePythonReady();
            var factorNamesJson = JsonConvert.SerializeObject(factors.Select(f => f.FactorName).ToList());
            var responseNamesJson = JsonConvert.SerializeObject(responses.Select(r => r.ResponseName).ToList());
            using (Py.GIL())
            {
                string resultJson = _pythonImporter!.validate_excel(excelPath, factorNamesJson, responseNamesJson).ToString();
                var pyResult = JsonConvert.DeserializeObject<PythonValidationResult>(resultJson)!;
                return Task.FromResult(new DataValidationResult
                {
                    IsValid = pyResult.IsValid,
                    Errors = pyResult.Errors ?? new(),
                    Warnings = pyResult.Warnings ?? new(),
                    ValidRowCount = pyResult.ValidRowCount
                });
            }
        }

        public Task<List<DOERunRecord>> ImportHistoricalDataAsync(string excelPath, string batchId, List<DOEFactor> factors, List<DOEResponse> responses)
        {
            EnsurePythonReady();
            var factorNamesJson = JsonConvert.SerializeObject(factors.Select(f => f.FactorName).ToList());
            var responseNamesJson = JsonConvert.SerializeObject(responses.Select(r => r.ResponseName).ToList());
            using (Py.GIL())
            {
                string dataJson = _pythonImporter!.import_excel(excelPath, factorNamesJson, responseNamesJson).ToString();
                var records = JsonConvert.DeserializeObject<List<ImportedRecord>>(dataJson) ?? new();
                var runs = records.Select((r, idx) => new DOERunRecord
                {
                    BatchId = batchId,
                    RunIndex = 10000 + idx,
                    FactorValuesJson = JsonConvert.SerializeObject(r.Factors),
                    ResponseValuesJson = JsonConvert.SerializeObject(r.Responses),
                    DataSource = DOEDataSource.Imported,
                    Status = DOERunStatus.Completed,
                    StartTime = DateTime.Now,
                    EndTime = DateTime.Now
                }).ToList();
                return Task.FromResult(runs);
            }
        }

        // ═══════════════════════════════════════════════════════
        //  新增: RSM 设计方法
        // ═══════════════════════════════════════════════════════

        public Task<DOEDesignMatrix> GenerateCCDAsync(List<DOEFactor> factors, string alphaType = "rotatable", int centerCount = -1)
        {
            EnsurePythonReady();
            var factorsJson = BuildFactorsJsonWithBounds(factors);
            _logger.LogInformation(" 生成 CCD 设计: {FactorCount} 个因子, alpha={Alpha}, 中心点={Center}",
                factors.Count, alphaType, centerCount);

            using (Py.GIL())
            {
                string matrixJson = _pythonDesigner!.ccd(factorsJson, alphaType, centerCount).ToString();
                var result = ParseDesignMatrix(matrixJson, factors);
                _logger.LogInformation(" CCD 生成完成: {RunCount} 组实验", result.RunCount);
                return Task.FromResult(result);
            }
        }

        public Task<DOEDesignMatrix> GenerateBoxBehnkenAsync(List<DOEFactor> factors, int centerCount = -1)
        {
            EnsurePythonReady();
            var factorsJson = BuildFactorsJsonWithBounds(factors);
            _logger.LogInformation(" 生成 Box-Behnken 设计: {FactorCount} 个因子, 中心点={Center}",
                factors.Count, centerCount);

            using (Py.GIL())
            {
                string matrixJson = _pythonDesigner!.box_behnken(factorsJson, centerCount).ToString();
                var result = ParseDesignMatrix(matrixJson, factors);
                _logger.LogInformation(" BBD 生成完成: {RunCount} 组实验", result.RunCount);
                return Task.FromResult(result);
            }
        }

        public Task<DOEDesignMatrix> GenerateDOptimalAsync(List<DOEFactor> factors, int numRuns = -1, string modelType = "quadratic")
        {
            EnsurePythonReady();
            var factorsJson = BuildFactorsJsonWithBounds(factors);
            _logger.LogInformation(" 生成 D-Optimal 设计: {FactorCount} 个因子, 组数={Runs}, 模型={Model}",
                factors.Count, numRuns, modelType);

            using (Py.GIL())
            {
                string matrixJson = _pythonDesigner!.d_optimal(factorsJson, numRuns, modelType).ToString();
                var result = ParseDesignMatrix(matrixJson, factors);
                _logger.LogInformation(" D-Optimal 生成完成: {RunCount} 组实验", result.RunCount);
                return Task.FromResult(result);
            }
        }

        // ═══════════════════════════════════════════════════════
        //  新增: 设计质量评估
        // ═══════════════════════════════════════════════════════

        public Task<DOEDesignQuality> GetDesignQualityAsync(List<DOEFactor> factors, DOEDesignMatrix matrix, string modelType = "quadratic")
        {
            EnsurePythonReady();
            var factorsJson = BuildFactorsJsonWithBounds(factors);
            var matrixJson = matrix.ToJson();
            _logger.LogInformation(" 评估设计质量: {FactorCount} 因子, {RunCount} 组, 模型={Model}",
                factors.Count, matrix.RunCount, modelType);

            using (Py.GIL())
            {
                string resultJson = _pythonDesigner!.get_design_quality(factorsJson, matrixJson, modelType).ToString();
                var pyResult = JsonConvert.DeserializeObject<PythonDesignQuality>(resultJson)!;

                var quality = new DOEDesignQuality
                {
                    DEfficiency = pyResult.DEfficiency,
                    AEfficiency = pyResult.AEfficiency,
                    GEfficiency = pyResult.GEfficiency,
                    VIF = pyResult.VIF ?? new(),
                    ConditionNumber = pyResult.ConditionNumber,
                    PowerAnalysis = pyResult.PowerAnalysis ?? new(),
                    RunCount = pyResult.RunCount,
                    ParameterCount = pyResult.ParameterCount,
                    DegreesOfFreedom = pyResult.DegreesOfFreedom
                };

                _logger.LogInformation(" 设计质量: D-eff={DEff}, A-eff={AEff}, G-eff={GEff}",
                    quality.DEfficiency, quality.AEfficiency, quality.GEfficiency);

                return Task.FromResult(quality);
            }
        }

        // ══════════════ 内部辅助方法 ══════════════

        private string BuildFactorsJsonWithLevels(List<DOEFactor> factors)
        {
            var list = factors.Select(f =>
            {
                if (f.IsCategorical)
                {
                    // ★ 修复 v4: 类别因子直接传字符串标签作为 levels
                    var catLevels = f.GetCategoryLevelList();
                    return new Dictionary<string, object>
            {
                { "name", f.FactorName },
                { "type", "categorical" },
                { "levels", catLevels }
            };
                }
                else
                {
                    return new Dictionary<string, object>
            {
                { "name", f.FactorName },
                { "type", "continuous" },
                { "levels", f.GetLevels().Select(v => (object)v).ToList() }
            };
                }
            }).ToList();
            return JsonConvert.SerializeObject(list);
        }

        private string BuildFactorsJsonWithBounds(List<DOEFactor> factors)
        {
            var list = factors.Select(f =>
            {
                if (f.IsCategorical)
                {
                    return new Dictionary<string, object>
                    {
                        { "name", f.FactorName },
                        { "type", "categorical" },
                        { "levels", f.GetCategoryLevelList() }
                    };
                }
                else
                {
                    return new Dictionary<string, object>
                    {
                        { "name", f.FactorName },
                        { "type", "continuous" },
                        { "lower", f.LowerBound },
                        { "upper", f.UpperBound }
                    };
                }
            }).ToList();
            return JsonConvert.SerializeObject(list);
        }

        private DOEDesignMatrix ParseDesignMatrix(string matrixJson, List<DOEFactor> factors)
        {
            // ★ 修复 (v3): 类别因子保留字符串标签，不再转为数值索引
            // 这样标签可以从设计矩阵一直透传到 GPR 的 _encode_factors()
            var rawRows = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(matrixJson)
                       ?? new List<Dictionary<string, object>>();

            // 构建类别因子名称集合（用于判断是否需要保留字符串）
            var categoricalFactorNames = new HashSet<string>(
                factors.Where(f => f.IsCategorical).Select(f => f.FactorName));

            var rows = rawRows.Select(rawRow =>
            {
                var row = new Dictionary<string, object>();
                foreach (var kv in rawRow)
                {
                    if (categoricalFactorNames.Contains(kv.Key))
                    {
                        // 类别因子: 保留字符串标签（如 "催化剂A"）
                        row[kv.Key] = kv.Value?.ToString() ?? "";
                    }
                    else
                    {
                        // 连续因子: 转为 double
                        if (kv.Value is double d)
                            row[kv.Key] = d;
                        else if (kv.Value is long l)
                            row[kv.Key] = (double)l;
                        else if (double.TryParse(kv.Value?.ToString(), out var parsed))
                            row[kv.Key] = parsed;
                        else
                            row[kv.Key] = 0.0;
                    }
                }
                return row;
            }).ToList();

            return new DOEDesignMatrix
            {
                FactorNames = factors.Select(f => f.FactorName).ToList(),
                Rows = rows
            };
        }

        // ── Python 结果反序列化辅助类 ──

        private class ImportedRecord
        {
            /// <summary>
            /// ★ 修复 (Bug#4): 从 Dictionary&lt;string, double&gt; 改为 Dictionary&lt;string, object&gt;
            /// 类别因子在 Excel 中是字符串值，double 反序列化会失败
            /// </summary>
            [JsonProperty("factors")] public Dictionary<string, object> Factors { get; set; } = new();
            [JsonProperty("responses")] public Dictionary<string, double> Responses { get; set; } = new();
        }

        private class PythonValidationResult
        {
            [JsonProperty("is_valid")] public bool IsValid { get; set; }
            [JsonProperty("errors")] public List<string>? Errors { get; set; }
            [JsonProperty("warnings")] public List<string>? Warnings { get; set; }
            [JsonProperty("valid_row_count")] public int ValidRowCount { get; set; }
        }

        //  新增: 设计质量反序列化类
        private class PythonDesignQuality
        {
            [JsonProperty("d_efficiency")] public double DEfficiency { get; set; }
            [JsonProperty("a_efficiency")] public double AEfficiency { get; set; }
            [JsonProperty("g_efficiency")] public double GEfficiency { get; set; }
            [JsonProperty("vif")] public Dictionary<string, double>? VIF { get; set; }
            [JsonProperty("condition_number")] public double ConditionNumber { get; set; }
            [JsonProperty("power_analysis")] public Dictionary<string, double>? PowerAnalysis { get; set; }
            [JsonProperty("run_count")] public int RunCount { get; set; }
            [JsonProperty("parameter_count")] public int ParameterCount { get; set; }
            [JsonProperty("degrees_of_freedom")] public int DegreesOfFreedom { get; set; }
        }
    }
}