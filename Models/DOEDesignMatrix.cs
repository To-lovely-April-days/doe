using System.Collections.Generic;
using System.Linq;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// DOE 参数矩阵 — 由设计方法生成的完整实验参数组合表
    /// 
    /// ★ 修复 (v3): Rows 类型从 Dictionary&lt;string, double&gt; 改为 Dictionary&lt;string, object&gt;
    /// 连续因子存 double 值 (如 120.0)，类别因子存 string 标签 (如 "催化剂A")
    /// 这样标签可以从设计矩阵一直透传到 GPR 的 _encode_factors()
    /// </summary>
    public class DOEDesignMatrix
    {
        /// <summary>
        /// 因子名称列表（列头）
        /// </summary>
        public List<string> FactorNames { get; set; } = new();

        /// <summary>
        /// 参数矩阵行（每行是一组实验的因子值）
        /// ★ 修复 (v3): 类型从 Dictionary&lt;string, double&gt; 改为 Dictionary&lt;string, object&gt;
        /// 连续因子: double (如 120.0)
        /// 类别因子: string (如 "催化剂A")
        /// </summary>
        public List<Dictionary<string, object>> Rows { get; set; } = new();

        /// <summary>
        /// 总实验组数
        /// </summary>
        public int RunCount => Rows.Count;

        /// <summary>
        /// 因子数量
        /// </summary>
        public int FactorCount => FactorNames.Count;

        /// <summary>
        /// 矩阵是否为空
        /// </summary>
        public bool IsEmpty => Rows.Count == 0;

        /// <summary>
        /// 获取指定组的因子值
        /// </summary>
        public Dictionary<string, object>? GetRun(int index)
        {
            return index >= 0 && index < Rows.Count ? Rows[index] : null;
        }

        /// <summary>
        /// 添加中心点实验（仅连续因子有效，类别因子取第一个水平）
        /// </summary>
        public void AddCenterPoint(Dictionary<string, object> centerValues)
        {
            Rows.Add(new Dictionary<string, object>(centerValues));
        }

        /// <summary>
        /// 随机化排列
        /// </summary>
        public void Randomize(int? seed = null)
        {
            var rng = seed.HasValue ? new System.Random(seed.Value) : new System.Random();
            var n = Rows.Count;
            for (int i = n - 1; i > 0; i--)
            {
                int j = rng.Next(i + 1);
                (Rows[i], Rows[j]) = (Rows[j], Rows[i]);
            }
        }

        /// <summary>
        /// 转为 JSON 字符串（供 Python 调用）
        /// </summary>
        public string ToJson()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(Rows);
        }

        /// <summary>
        /// 从 JSON 字符串还原
        /// </summary>
        public static DOEDesignMatrix FromJson(string json, List<string> factorNames)
        {
            var rows = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(json)
                       ?? new List<Dictionary<string, object>>();
            return new DOEDesignMatrix { FactorNames = factorNames, Rows = rows };
        }

        /// <summary>
        /// ★ 新增 (v3): 获取指定因子的 double 值（连续因子用）
        /// 类别因子返回 0.0
        /// </summary>
        public double GetDoubleValue(int rowIndex, string factorName)
        {
            if (rowIndex < 0 || rowIndex >= Rows.Count) return 0.0;
            if (!Rows[rowIndex].TryGetValue(factorName, out var val)) return 0.0;
            if (val is double d) return d;
            if (val is long l) return l;
            if (val is int i) return i;
            if (double.TryParse(val?.ToString(), out var parsed)) return parsed;
            return 0.0;
        }

        /// <summary>
        /// ★ 新增 (v3): 获取指定因子的 string 值（类别因子用）
        /// </summary>
        public string GetStringValue(int rowIndex, string factorName)
        {
            if (rowIndex < 0 || rowIndex >= Rows.Count) return "";
            if (!Rows[rowIndex].TryGetValue(factorName, out var val)) return "";
            return val?.ToString() ?? "";
        }
    }
}