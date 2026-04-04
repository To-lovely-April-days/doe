using MaxChemical.Modules.DOE.Models;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace MaxChemical.Modules.DOE.Views
{
    public partial class OlsImportConfirmDialog : Window
    {
        public ObservableCollection<ColumnConfigItem> Columns { get; set; }
        public string FilePath { get; set; } = "";
        public int TotalRows { get; set; }

        /// <summary>用户确认后为 true</summary>
        public bool Confirmed { get; private set; }

        public OlsImportConfirmDialog(ObservableCollection<ColumnConfigItem> columns, string filePath, int totalRows)
        {
            Columns = columns;
            FilePath = filePath;
            TotalRows = totalRows;
            DataContext = this;
            InitializeComponent();
            UpdateSummary();

            // 监听每列角色变化刷新摘要
            foreach (var col in Columns)
                col.PropertyChanged += (s, e) => UpdateSummary();
        }

        private void UpdateSummary()
        {
            int factors = Columns.Count(c => c.Role == "Factor");
            int responses = Columns.Count(c => c.Role == "Response");
            int ignored = Columns.Count(c => c.Role == "Ignore");
            int continuous = Columns.Count(c => c.Role == "Factor" && c.FactorType == "Continuous");
            int categorical = Columns.Count(c => c.Role == "Factor" && c.FactorType == "Categorical");
            SummaryText.Text = $"共 {TotalRows} 行数据  |  {factors} 个因子（{continuous} 连续 + {categorical} 类别）  |  {responses} 个响应  |  {ignored} 列忽略";
        }

        private void Confirm_Click(object sender, RoutedEventArgs e)
        {
            // 校验
            int responseCount = Columns.Count(c => c.Role == "Response");
            int factorCount = Columns.Count(c => c.Role == "Factor");

            if (responseCount == 0)
            {
                MessageBox.Show("请至少指定 1 个响应变量列。", "校验失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (factorCount == 0)
            {
                MessageBox.Show("请至少指定 1 个因子列。", "校验失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 响应列必须全为数值
            var nonNumericResponses = Columns.Where(c => c.Role == "Response" && !c.IsAllNumeric).ToList();
            if (nonNumericResponses.Count > 0)
            {
                var names = string.Join(", ", nonNumericResponses.Select(c => c.ColumnName));
                MessageBox.Show($"响应变量列必须为纯数值，以下列包含非数值数据：\n{names}", "校验失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Confirmed = true;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Confirmed = false;
            DialogResult = false;
            Close();
        }
    }
}