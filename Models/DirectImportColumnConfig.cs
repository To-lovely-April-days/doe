using System.Collections.Generic;
using System.Linq;
using Prism.Mvvm;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// 单列配置（用户在确认弹窗中编辑）
    /// </summary>
    public class ColumnConfigItem : BindableBase
    {
        /// <summary>Excel 列名</summary>
        public string ColumnName { get; set; } = "";

        /// <summary>前5行预览值</summary>
        public List<string> PreviewValues { get; set; } = new();

        /// <summary>预览文本（逗号拼接）</summary>
        public string PreviewText => string.Join(", ", PreviewValues.Take(5));

        /// <summary>是否全为数值</summary>
        public bool IsAllNumeric { get; set; }

        /// <summary>唯一值数量（帮助判断是否类别）</summary>
        public int UniqueCount { get; set; }

        /// <summary>角色: Factor / Response / Ignore</summary>
        private string _role = "Factor";
        public string Role
        {
            get => _role;
            set
            {
                if (SetProperty(ref _role, value))
                {
                    // 响应和忽略不需要选类型
                    RaisePropertyChanged(nameof(IsFactorTypeEnabled));
                    if (value != "Factor") FactorType = "";
                }
            }
        }

        /// <summary>因子类型: Continuous / Categorical（仅角色=Factor时有效）</summary>
        private string _factorType = "Continuous";
        public string FactorType
        {
            get => _factorType;
            set => SetProperty(ref _factorType, value);
        }

        /// <summary>类型下拉是否可选</summary>
        public bool IsFactorTypeEnabled => Role == "Factor";

        /// <summary>角色选项</summary>
        public static List<string> RoleOptions => new() { "Factor", "Response", "Ignore" };

        /// <summary>因子类型选项</summary>
        public static List<string> FactorTypeOptions => new() { "Continuous", "Categorical" };
    }
}