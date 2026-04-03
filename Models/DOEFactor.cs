using MaxChemical.Infrastructure.DOE;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// DOE 因子定义
    /// </summary>
    public class DOEFactor
    {
        public int Id { get; set; }
        public string BatchId { get; set; } = string.Empty;
        public string FactorName { get; set; } = string.Empty;
        public FactorSourceType FactorSource { get; set; } = FactorSourceType.ParameterOverride;
        public string? SourceNodeId { get; set; }
        public string? SourceParamName { get; set; }
        public double LowerBound { get; set; }
        public double UpperBound { get; set; }
        public int LevelCount { get; set; } = 3;
        public DOEFactorType FactorType { get; set; } = DOEFactorType.Continuous;
        public int SortOrder { get; set; }

        /// <summary>
        /// ★ 新增: 类别因子的水平标签（逗号分隔，如 "催化剂A,催化剂B,催化剂C"）
        /// 仅当 FactorType == Categorical 时有效
        /// </summary>
        public string? CategoryLevels { get; set; }

        /// <summary>
        /// ★ 新增: 获取类别因子的水平标签列表
        /// </summary>
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
        /// ★ 新增: 是否为类别因子
        /// </summary>
        public bool IsCategorical => FactorType == DOEFactorType.Categorical;

        /// <summary>
        /// ★ 新增: 是否为连续因子
        /// </summary>
        public bool IsContinuous => FactorType == DOEFactorType.Continuous;

        /// <summary>
        /// 计算步长（仅连续因子有效）
        /// </summary>
        public double StepSize => IsContinuous && LevelCount > 1
            ? (UpperBound - LowerBound) / (LevelCount - 1) : 0;

        /// <summary>
        /// 计算中心点（仅连续因子有效）
        /// ★ 修复 (Bug#8): 类别因子返回 NaN 而不是无意义的 0.0
        /// </summary>
        public double CenterPoint => IsContinuous
            ? (UpperBound + LowerBound) / 2.0
            : double.NaN;

        /// <summary>
        /// 生成水平值列表（仅连续因子有效）
        /// ★ 修复 (Bug#8): 类别因子返回空数组，调用方应使用 GetCategoryLevelList()
        /// </summary>
        public double[] GetLevels()
        {
            if (IsCategorical) return Array.Empty<double>();
            if (LevelCount <= 1) return new[] { CenterPoint };
            var levels = new double[LevelCount];
            var step = StepSize;
            for (int i = 0; i < LevelCount; i++)
            {
                levels[i] = LowerBound + step * i;
            }
            return levels;
        }
    }

    /// <summary>
    /// 因子类型
    /// ★ 修改: Discrete 改名为 Categorical，语义更准确
    /// </summary>
    public enum DOEFactorType
    {
        /// <summary>连续因子 — 有上下界，参与二次项</summary>
        Continuous,
        /// <summary>类别因子 — 有水平标签（如催化剂A/B/C），不参与二次项</summary>
        Categorical
    }
}