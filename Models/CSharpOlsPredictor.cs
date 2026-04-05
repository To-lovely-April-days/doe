using System;
using System.Collections.Generic;
using System.Linq;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// ★ v9: C# 端 OLS 轻量预测器（含 CI 计算）
    /// 
    /// 预测刻画器拖动时，完全在 C# 端计算预测值 + 置信区间，不调 Python。
    /// </summary>
    public class CSharpOlsPredictor
    {
        // ── 模型系数 ──
        private double _intercept;
        private readonly List<string> _termNames = new();
        private readonly List<double> _termCoeffs = new();

        // ── CI 计算所需 ──
        private double[][]? _xtxInv;    // (XᵀX)⁻¹
        private double _sigma;          // σ̂ = √MSE
        private double _tCrit;          // t(α/2, df_error)
        private int _paramCount;
        private bool _hasCi;

        // ── 因子信息 ──
        private readonly List<string> _factorNames = new();
        private readonly Dictionary<string, CodingInfo> _codingInfos = new();
        private readonly Dictionary<string, List<string>> _categoricalLevels = new();
        private readonly HashSet<string> _categoricalFactors = new();
        private readonly Dictionary<string, (double min, double max)> _factorRanges = new();
        private readonly Dictionary<string, string> _safeNames = new();

        public bool IsReady { get; private set; }

        public void Initialize(
            List<CoefficientRow> coefficients,
            Dictionary<string, CodingInfoItem>? codingInfo,
            List<string> factorNames,
            List<string> categoricalFactors,
            Dictionary<string, List<string>> categoricalLevels,
            Dictionary<string, (double min, double max)>? factorRanges = null)
        {
            _factorNames.Clear();
            _factorNames.AddRange(factorNames);

            _categoricalFactors.Clear();
            foreach (var cf in categoricalFactors) _categoricalFactors.Add(cf);

            _categoricalLevels.Clear();
            foreach (var kv in categoricalLevels) _categoricalLevels[kv.Key] = kv.Value;

            _factorRanges.Clear();
            if (factorRanges != null)
                foreach (var kv in factorRanges) _factorRanges[kv.Key] = kv.Value;

            _safeNames.Clear();
            for (int i = 0; i < factorNames.Count; i++)
                _safeNames[factorNames[i]] = $"X{i}";

            _codingInfos.Clear();
            if (codingInfo != null)
                foreach (var kv in codingInfo)
                    _codingInfos[kv.Key] = new CodingInfo
                    {
                        IsCategorical = kv.Value.Type == "categorical",
                        Center = kv.Value.Center ?? 0,
                        HalfRange = kv.Value.HalfRange ?? 1
                    };

            _termNames.Clear();
            _termCoeffs.Clear();
            _intercept = 0;

            foreach (var coeff in coefficients)
            {
                if (coeff.Term == "截距" || coeff.Term == "Intercept")
                    _intercept = coeff.Coefficient;
                else
                {
                    _termNames.Add(coeff.Term);
                    _termCoeffs.Add(coeff.Coefficient);
                }
            }

            _paramCount = _termNames.Count + 1;
            _hasCi = false;
            IsReady = _termNames.Count > 0;
        }

        /// <summary>
        /// 设置 CI 参数（从 Python 一次性导出）
        /// xtxInvFlat: 展平的 (XᵀX)⁻¹，长度 = p*p
        /// </summary>
        public void SetCIParameters(double[] xtxInvFlat, double sigma, double tCrit, int paramCount)
        {
            _sigma = sigma;
            _tCrit = tCrit;
            _paramCount = paramCount;
            _xtxInv = new double[paramCount][];
            for (int i = 0; i < paramCount; i++)
            {
                _xtxInv[i] = new double[paramCount];
                for (int j = 0; j < paramCount; j++)
                    _xtxInv[i][j] = xtxInvFlat[i * paramCount + j];
            }
            _hasCi = true;
        }

        public (double predicted, double ciLower, double ciUpper) PredictWithCI(
            Dictionary<string, object> factorValues)
        {
            if (!IsReady) return (0, 0, 0);

            var coded = EncodeFactors(factorValues);
            double predicted = _intercept;
            for (int i = 0; i < _termNames.Count; i++)
                predicted += _termCoeffs[i] * EvaluateTerm(_termNames[i], coded);

            if (!_hasCi || _xtxInv == null)
                return (predicted, predicted, predicted);

            // x₀ = [1, term0, term1, ...]
            var x0 = new double[_paramCount];
            x0[0] = 1.0;
            for (int i = 0; i < _termNames.Count && i + 1 < _paramCount; i++)
                x0[i + 1] = EvaluateTerm(_termNames[i], coded);

            // h = x₀ᵀ(XᵀX)⁻¹x₀
            double h = 0;
            for (int i = 0; i < _paramCount; i++)
            {
                double sum = 0;
                for (int j = 0; j < _paramCount; j++)
                    sum += _xtxInv[i][j] * x0[j];
                h += x0[i] * sum;
            }

            double seMean = _sigma * Math.Sqrt(Math.Max(0, h));
            double margin = _tCrit * seMean;
            return (predicted, predicted - margin, predicted + margin);
        }

        public double Predict(Dictionary<string, object> factorValues)
            => PredictWithCI(factorValues).predicted;

        public SweepResult SweepFactor(string factorName,
            Dictionary<string, object> fixedValues, int gridSize = 50)
        {
            var result = new SweepResult();

            if (_categoricalFactors.Contains(factorName))
            {
                if (_categoricalLevels.TryGetValue(factorName, out var levels))
                {
                    for (int i = 0; i < levels.Count; i++)
                    {
                        var point = new Dictionary<string, object>(fixedValues);
                        point[factorName] = levels[i];
                        var (pred, lo, hi) = PredictWithCI(point);
                        result.X.Add(i);
                        result.Y.Add(pred);
                        result.YLower.Add(lo);
                        result.YUpper.Add(hi);
                    }
                }
                result.IsCategorical = true;
            }
            else
            {
                var (fMin, fMax) = GetFactorRange(factorName);
                for (int i = 0; i < gridSize; i++)
                {
                    double x = fMin + (fMax - fMin) * i / (gridSize - 1);
                    var point = new Dictionary<string, object>(fixedValues);
                    point[factorName] = x;
                    var (pred, lo, hi) = PredictWithCI(point);
                    result.X.Add(x);
                    result.Y.Add(pred);
                    result.YLower.Add(lo);
                    result.YUpper.Add(hi);
                }
                result.IsCategorical = false;
            }
            return result;
        }

        public Dictionary<string, SweepResult> SweepAllFactors(
            Dictionary<string, object> currentValues, int gridSize = 50)
        {
            var results = new Dictionary<string, SweepResult>();
            foreach (var fname in _factorNames)
                results[fname] = SweepFactor(fname,
                    new Dictionary<string, object>(currentValues), gridSize);
            return results;
        }

        public (double min, double max) GetFactorRange(string factorName)
        {
            if (_factorRanges.TryGetValue(factorName, out var range)) return range;
            if (_codingInfos.TryGetValue(factorName, out var info) && !info.IsCategorical)
                return (info.Center - info.HalfRange, info.Center + info.HalfRange);
            return (0, 1);
        }

        // ── 内部 ──

        private Dictionary<string, double> EncodeFactors(Dictionary<string, object> factorValues)
        {
            var coded = new Dictionary<string, double>();
            foreach (var fname in _factorNames)
            {
                if (!factorValues.TryGetValue(fname, out var rawVal)) continue;
                var safeName = _safeNames.GetValueOrDefault(fname, fname);
                if (_categoricalFactors.Contains(fname))
                {
                    // ★ 不假设编码方式，为每个水平生成 indicator
                    // 系数表中的项名格式可能是:
                    //   Sum coding: "Catalyst[A]", "Catalyst[B]"
                    //   Treatment:  "Catalyst[T.B]", "Catalyst[T.C]"
                    // 这里两种都生成，EvaluateTerm 会自动匹配到存在的那个
                    var strVal = rawVal?.ToString() ?? "";
                    if (_categoricalLevels.TryGetValue(fname, out var levels))
                    {
                        foreach (var lv in levels)
                        {
                            // Sum coding 格式: "Catalyst[A]"
                            coded[$"{fname}[{lv}]"] = (strVal == lv) ? 1.0 : 0.0;
                            // Treatment coding 格式: "Catalyst[T.B]"
                            coded[$"{fname}[T.{lv}]"] = (strVal == lv) ? 1.0 : 0.0;
                        }
                    }
                }
                else
                {
                    double numVal = Convert.ToDouble(rawVal);
                    var info = _codingInfos.GetValueOrDefault(fname);
                    if (info != null && !info.IsCategorical)
                    {
                        double hr = info.HalfRange > 1e-12 ? info.HalfRange : 1.0;
                        coded[safeName] = (numVal - info.Center) / hr;
                    }
                    else coded[safeName] = numVal;
                }
            }
            return coded;
        }

        private double EvaluateTerm(string termName, Dictionary<string, double> coded)
        {
            if (termName.EndsWith("²"))
            {
                var safeName = _safeNames.GetValueOrDefault(termName[..^1], termName[..^1]);
                return coded.TryGetValue(safeName, out var v) ? v * v : 0;
            }
            if (termName.Contains('×'))
            {
                double product = 1.0;
                foreach (var part in termName.Split('×'))
                    product *= GetSingleTermValue(part.Trim(), coded);
                return product;
            }
            return GetSingleTermValue(termName, coded);
        }

        private double GetSingleTermValue(string termName, Dictionary<string, double> coded)
        {
            // ★ 直接用系数表中的项名去 coded 字典查找
            // 支持 "Catalyst[A]"、"Catalyst[T.B]"、"Temperature" 等格式
            if (termName.Contains('['))
                return coded.TryGetValue(termName, out var v) ? v : 0;
            var safeName = _safeNames.GetValueOrDefault(termName, termName);
            return coded.TryGetValue(safeName, out var val) ? val : 0;
        }

        private class CodingInfo
        {
            public bool IsCategorical { get; set; }
            public double Center { get; set; }
            public double HalfRange { get; set; }
        }
    }

    public class SweepResult
    {
        public List<double> X { get; set; } = new();
        public List<double> Y { get; set; } = new();
        public List<double> YLower { get; set; } = new();
        public List<double> YUpper { get; set; } = new();
        public bool IsCategorical { get; set; }
    }
}