using MaxChemical.Modules.DOE.ViewModels;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Media;

namespace MaxChemical.Modules.DOE.Views
{
    public partial class DOEExecutionDashboardView : UserControl
    {
        public DOEExecutionDashboardView()
        {
            InitializeComponent();
        }

        /// <summary>
        /// DataGrid 自动生成列时的样式定制：
        /// - # 列：窄宽度、左对齐、灰色
        /// - 因子列：等宽字体、右对齐
        /// - 响应列：蓝色加粗、右对齐
        /// - 状态列：居中
        /// - 耗时列：右对齐、灰色
        /// </summary>
        private void OnAutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            //  隐藏内部匹配用的列
            if (e.PropertyName == "_RunIndex")
            {
                e.Cancel = true;
                return;
            }

            var col = e.Column as DataGridTextColumn;
            if (col == null) return;

            switch (e.PropertyName)
            {
                case "#":
                    col.Width = new DataGridLength(48);
                    col.ElementStyle = CreateCellStyle(
                        hAlign: HorizontalAlignment.Left,
                        foreground: "#A0A4B0",
                        fontFamily: "Segoe UI",
                        fontWeight: FontWeights.Normal);
                    col.HeaderStyle = CreateHeaderStyle(HorizontalAlignment.Left);
                    break;

                case "状态":
                    col.Width = new DataGridLength(80);
                    col.ElementStyle = CreateCellStyle(
                        hAlign: HorizontalAlignment.Center,
                        foreground: "#4A5066",
                        fontFamily: "Segoe UI",
                        fontSize: 11);
                    col.HeaderStyle = CreateHeaderStyle(HorizontalAlignment.Center);
                    break;

                case "耗时":
                    col.Width = new DataGridLength(60);
                    col.ElementStyle = CreateCellStyle(
                        hAlign: HorizontalAlignment.Right,
                        foreground: "#A0A4B0",
                        fontFamily: "Consolas",
                        fontSize: 11);
                    col.HeaderStyle = CreateHeaderStyle(HorizontalAlignment.Right);
                    break;

                default:
                    // 判断是否是响应列（ViewModel 里的 _responseNames）
                    // 通过 DataContext 获取 ViewModel 来判断
                    bool isResponseCol = false;
                    if (DataContext is DOEExecutionDashboardViewModel vm)
                    {
                        // 响应列名不会是 #/状态/耗时，且不在因子列表中
                        // 简单判断: 如果列名不在 DataTable 的前几个固定列中
                        // 实际上我们可以通过检查 MatrixTable 的列顺序来判断
                    }

                    // 响应列用蓝色高亮 — 通过排除法：不是 #、不是 状态、不是 耗时
                    // 因子列和响应列都是动态的，我们用列索引判断
                    // 因子列在前，响应列在后（在 # 之后，状态/耗时之前）
                    col.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
                    col.ElementStyle = CreateCellStyle(
                        hAlign: HorizontalAlignment.Right,
                        foreground: "#1A1F36",
                        fontFamily: "Consolas",
                        fontSize: 12);
                    col.HeaderStyle = CreateHeaderStyle(HorizontalAlignment.Right);
                    break;
            }
        }

        /// <summary>
        /// 创建单元格样式
        /// </summary>
        private static Style CreateCellStyle(
            HorizontalAlignment hAlign = HorizontalAlignment.Right,
            string foreground = "#1A1F36",
            string fontFamily = "Consolas",
            FontWeight? fontWeight = null,
            double fontSize = 12)
        {
            var style = new Style(typeof(TextBlock));
            style.Setters.Add(new Setter(TextBlock.HorizontalAlignmentProperty, hAlign));
            style.Setters.Add(new Setter(TextBlock.ForegroundProperty,
                new BrushConverter().ConvertFromString(foreground) as Brush));
            style.Setters.Add(new Setter(TextBlock.FontFamilyProperty, new FontFamily(fontFamily)));
            style.Setters.Add(new Setter(TextBlock.FontSizeProperty, fontSize));
            style.Setters.Add(new Setter(TextBlock.PaddingProperty, new Thickness(8, 0, 8, 0)));
            style.Setters.Add(new Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center));
            if (fontWeight.HasValue)
                style.Setters.Add(new Setter(TextBlock.FontWeightProperty, fontWeight.Value));
            return style;
        }

        /// <summary>
        /// 创建列头样式
        /// </summary>
        private static Style CreateHeaderStyle(HorizontalAlignment hAlign)
        {
            var style = new Style(typeof(DataGridColumnHeader));
            style.Setters.Add(new Setter(DataGridColumnHeader.BackgroundProperty,
                new BrushConverter().ConvertFromString("#FAFBFC") as Brush));
            style.Setters.Add(new Setter(DataGridColumnHeader.ForegroundProperty,
                new BrushConverter().ConvertFromString("#8A8F9C") as Brush));
            style.Setters.Add(new Setter(DataGridColumnHeader.FontSizeProperty, 11.0));
            style.Setters.Add(new Setter(DataGridColumnHeader.FontWeightProperty, FontWeights.SemiBold));
            style.Setters.Add(new Setter(DataGridColumnHeader.PaddingProperty, new Thickness(12, 0, 12, 0)));
            style.Setters.Add(new Setter(DataGridColumnHeader.BorderBrushProperty,
                new BrushConverter().ConvertFromString("#E8EAED") as Brush));
            style.Setters.Add(new Setter(DataGridColumnHeader.BorderThicknessProperty, new Thickness(0, 0, 0, 1)));
            style.Setters.Add(new Setter(DataGridColumnHeader.HorizontalContentAlignmentProperty, hAlign));
            return style;
        }
    }
}