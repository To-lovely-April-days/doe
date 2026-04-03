using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MaxChemical.Modules.DOE.Models
{
    public class DOEDesignQuality
    {
        /// <summary>
        /// D-效率 (0-1)，越接近 1 越好
        /// 物理含义: 系数估计的联合精度
        /// 公式: (|X'X|^(1/p)) / n
        /// </summary>
        public double DEfficiency { get; set; }

        /// <summary>
        /// A-效率 (0-1)
        /// 物理含义: 系数估计方差之和的倒数
        /// 公式: p / (trace((X'X)^-1) × n)
        /// </summary>
        public double AEfficiency { get; set; }

        /// <summary>
        /// G-效率 (0-1)
        /// 物理含义: 最大预测方差与平均预测方差的比值
        /// 公式: p / (n × max(leverage))
        /// </summary>
        public double GEfficiency { get; set; }

        /// <summary>
        /// 各因子的方差膨胀因子 (VIF)
        /// VIF &lt; 5 优秀, 5-10 可接受, &gt;10 严重多重共线性
        /// </summary>
        public Dictionary<string, double> VIF { get; set; } = new();

        /// <summary>
        /// 条件数 — &gt;30 表示矩阵病态（设计不稳定）
        /// </summary>
        public double ConditionNumber { get; set; }

        /// <summary>
        /// 各因子的检验功效 (Power)
        /// 0-1 之间，&gt;0.8 表示有足够能力检测到效应
        /// </summary>
        public Dictionary<string, double> PowerAnalysis { get; set; } = new();

        /// <summary>
        /// 混杂/别名结构（部分因子设计时重要）
        /// </summary>
        public List<AliasEntry> AliasStructure { get; set; } = new();

        /// <summary>
        /// 实验组数
        /// </summary>
        public int RunCount { get; set; }

        /// <summary>
        /// 模型参数数
        /// </summary>
        public int ParameterCount { get; set; }

        /// <summary>
        /// 自由度 (n - p)
        /// </summary>
        public int DegreesOfFreedom { get; set; }
    }

    /// <summary>
    /// 别名结构条目
    /// </summary>
    public class AliasEntry
    {
        public string Term { get; set; } = string.Empty;
        public string AliasedWith { get; set; } = string.Empty;
    }
}
