using MaxChemical.Modules.DOE.Models;
using MaxChemical.Modules.DOE.Services;
using OxyPlot;
using OxyPlot.Annotations;
using OxyPlot.Axes;
using OxyPlot.Series;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace MaxChemical.Modules.DOE.Views
{
    /// <summary>
    /// DesirabilityConfigDialog.xaml 的交互逻辑
    /// </summary>
    public partial class DesirabilityConfigDialog : Window
    {
        public bool Confirmed { get; private set; }
        public List<DesirabilityResponseConfig> ResultConfigs { get; private set; } = new();

        private readonly IDesirabilityService _desirabilityService;
        private ObservableCollection<DesirabilityConfigItem> _configs;

        public DesirabilityConfigDialog(List<DesirabilityResponseConfig> configs,
                                         IDesirabilityService service)
        {
            InitializeComponent();
            _desirabilityService = service;

            _configs = new ObservableCollection<DesirabilityConfigItem>(
                configs.Select(c => new DesirabilityConfigItem(c)));

            DataContext = new { Configs = _configs };

            // 为每个配置生成预览曲线
            foreach (var cfg in _configs)
                cfg.UpdatePreview();
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            Confirmed = true;
            ResultConfigs = _configs.Select(c => c.ToConfig()).ToList();
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Confirmed = false;
            Close();
        }
    }

    /// <summary>
    /// 弹窗中每个响应的配置项（含实时预览）
    /// </summary>
    public class DesirabilityConfigItem : BindableBase
    {
        public string ResponseName { get; set; } = "";

        private string _goal = "maximize";
        private double _lower, _upper, _target;
        private int _importance = 3;
        private double _shape = 1.0, _shapeLower = 1.0, _shapeUpper = 1.0;
        private PlotModel? _previewPlot;

        public string Goal { get => _goal; set { if (SetProperty(ref _goal, value)) UpdatePreview(); } }
        public double Lower { get => _lower; set { if (SetProperty(ref _lower, value)) UpdatePreview(); } }
        public double Upper { get => _upper; set { if (SetProperty(ref _upper, value)) UpdatePreview(); } }
        public double Target { get => _target; set { if (SetProperty(ref _target, value)) UpdatePreview(); } }
        public int Importance { get => _importance; set { if (SetProperty(ref _importance, value)) UpdatePreview(); } }
        public double Shape { get => _shape; set { if (SetProperty(ref _shape, value)) UpdatePreview(); } }
        public double ShapeLower { get => _shapeLower; set => SetProperty(ref _shapeLower, value); }
        public double ShapeUpper { get => _shapeUpper; set => SetProperty(ref _shapeUpper, value); }
        public PlotModel? PreviewPlot { get => _previewPlot; set => SetProperty(ref _previewPlot, value); }
        // 选项列表
        public static List<string> GoalOptions => new() { "最大化", "最小化", "望目" };
        public DesirabilityConfigItem(DesirabilityResponseConfig cfg)
        {
            ResponseName = cfg.ResponseName;
            _goal = cfg.Goal.ToString().ToLower();
            _lower = cfg.Lower;
            _upper = cfg.Upper;
            _target = cfg.Target;
            _importance = cfg.Importance;
            _shape = cfg.Shape;
            _shapeLower = cfg.ShapeLower;
            _shapeUpper = cfg.ShapeUpper;
            ResponseName = cfg.ResponseName;
            _goal = cfg.Goal switch
            {
                DesirabilityGoal.Maximize => "最大化",
                DesirabilityGoal.Minimize => "最小化",
                DesirabilityGoal.Target => "望目",
                _ => "最大化"
            };
        }

        /// <summary>
        /// 实时预览 d(y) 曲线 — 纯 C# 端计算（不走 Python，避免延迟）
        /// </summary>
        public void UpdatePreview()
        {
            var model = new PlotModel
            {
                PlotMargins = new OxyThickness(40, 5, 10, 30)
            };
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "y", FontSize = 9 });
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "d(y)",
                Minimum = -0.05,
                Maximum = 1.05,
                FontSize = 9
            });

            double margin = (Upper - Lower) * 0.1;
            double yMin = Lower - margin;
            double yMax = Upper + margin;
            int n = 100;

            var series = new LineSeries { Color = OxyColors.SteelBlue, StrokeThickness = 2 };

            for (int i = 0; i <= n; i++)
            {
                double y = yMin + (yMax - yMin) * i / n;
                double d = ComputeD(y);
                series.Points.Add(new DataPoint(y, d));
            }
            model.Series.Add(series);

            // 标注 L, T, U 竖线
            model.Annotations.Add(new LineAnnotation
            {
                Type = LineAnnotationType.Vertical,
                X = Lower,
                Color = OxyColors.Green,
                LineStyle = LineStyle.Dot,
                Text = "L"
            });
            model.Annotations.Add(new LineAnnotation
            {
                Type = LineAnnotationType.Vertical,
                X = Target,
                Color = OxyColors.Red,
                LineStyle = LineStyle.Dash,
                Text = "T"
            });
            model.Annotations.Add(new LineAnnotation
            {
                Type = LineAnnotationType.Vertical,
                X = Upper,
                Color = OxyColors.Blue,
                LineStyle = LineStyle.Dot,
                Text = "U"
            });

            PreviewPlot = model;
        }

        private double ComputeD(double y)
        {
            if (Goal == "最大化")
            {
                if (y <= Lower) return 0;
                if (y >= Target) return 1;
                var denom = Target - Lower;
                return denom < 1e-12 ? 1 : Math.Pow((y - Lower) / denom, Shape);
            }
            else if (Goal == "最小化")
            {
                if (y >= Upper) return 0;
                if (y <= Target) return 1;
                var denom = Upper - Target;
                return denom < 1e-12 ? 1 : Math.Pow((Upper - y) / denom, Shape);
            }
            else // 望目
            {
                if (y < Lower || y > Upper) return 0;
                if (y <= Target)
                { var denom = Target - Lower; return denom < 1e-12 ? 1 : Math.Pow((y - Lower) / denom, ShapeLower); }
                else
                { var denom = Upper - Target; return denom < 1e-12 ? 1 : Math.Pow((Upper - y) / denom, ShapeUpper); }
            }
        }

        public DesirabilityResponseConfig ToConfig() => new()
        {
            ResponseName = ResponseName,
            Goal = Goal switch
            {
                "最大化" => DesirabilityGoal.Maximize,
                "最小化" => DesirabilityGoal.Minimize,
                "望目" => DesirabilityGoal.Target,
                _ => DesirabilityGoal.Maximize
            },
            Lower = Lower,
            Upper = Upper,
            Target = Target,
            Weight = 1.0,
            Importance = Importance,
            Shape = Shape,
            ShapeLower = ShapeLower,
            ShapeUpper = ShapeUpper
        };
    }

}
