using System.Windows.Media;

namespace MaxChemical.Modules.DOE.Models
{
    /// <summary>
    /// 分段进度条的单个段
    /// </summary>
    public class ProgressSegment
    {
        public Brush Background { get; set; } = Brushes.Transparent;

        /// <summary>已完成的段（蓝色实心）</summary>
        public static ProgressSegment Completed(bool isPaused = false) => new()
        {
            Background = new SolidColorBrush(isPaused
                ? Color.FromRgb(0xF0, 0xA0, 0x30)    // 橙色
                : Color.FromRgb(0x3B, 0x9E, 0xFF))    // 蓝色
        };

        /// <summary>正在执行的段（半透明）</summary>
        public static ProgressSegment Running(bool isPaused = false) => new()
        {
            Background = new SolidColorBrush(isPaused
                ? Color.FromArgb(0x59, 0xF0, 0xA0, 0x30)   // 橙色 35%
                : Color.FromArgb(0x59, 0x3B, 0x9E, 0xFF))   // 蓝色 35%
        };

        /// <summary>未执行的段（深色底）</summary>
        public static ProgressSegment Pending() => new()
        {
            Background = new SolidColorBrush(Color.FromArgb(0x0F, 0xFF, 0xFF, 0xFF))  // 白色 6%
        };
    }

    /// <summary>
    /// 因子参数显示项
    /// </summary>
    public class FactorDisplayItem
    {
        public string Label { get; set; } = "";
        public string Value { get; set; } = "";
        public string Unit { get; set; } = "";
    }
}