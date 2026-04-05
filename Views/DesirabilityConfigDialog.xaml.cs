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
using System.Windows;

namespace MaxChemical.Modules.DOE.Views
{
    public partial class DesirabilityConfigDialog : Window
    {
        public bool Confirmed { get; private set; }
        public List<DesirabilityResponseConfig> ResultConfigs { get; private set; } = new();

        private readonly IDesirabilityService _desirabilityService;
        private ObservableCollection<DesirabilityConfigItem> _configs;

        public DesirabilityConfigDialog(
            List<DesirabilityResponseConfig> configs,
            IDesirabilityService service,
            Dictionary<string, (double min, double max)>? dataRanges = null)
        {
            InitializeComponent();
            _desirabilityService = service;

            _configs = new ObservableCollection<DesirabilityConfigItem>(
                configs.Select(c =>
                {
                    var item = new DesirabilityConfigItem(c);
                    if (dataRanges != null && dataRanges.TryGetValue(c.ResponseName, out var range))
                    {
                        item.DataMin = range.min;
                        item.DataMax = range.max;
                    }
                    if (item.IsDefaultValues && !double.IsNaN(item.DataMin))
                        item.AutoFillFromDataRange();
                    return item;
                }));

            DataContext = new { Configs = _configs };
            foreach (var cfg in _configs) cfg.UpdatePreview();
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cfg in _configs)
            {
                var err = cfg.Validate();
                if (err != null)
                {
                    MessageBox.Show($"响应 '{cfg.ResponseName}': {err}", "校验失败",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }
            Confirmed = true;
            ResultConfigs = _configs.Select(c => c.ToConfig()).ToList();
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Confirmed = false;
            Close();
        }

        private void ResetDefaults_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cfg in _configs)
            {
                cfg.AutoFillFromDataRange();
                cfg.UpdatePreview();
            }
        }
    }

    /// <summary>
    /// 意愿配置项 — 用户直觉的属性命名
    /// 最大化: 用户填 TargetValue(目标) + LowerBound(最低可接受)
    /// 最小化: 用户填 TargetValue(目标) + UpperBound(最高可接受)
    /// 望目:   用户填 LowerBound(下限) + TargetValue(目标) + UpperBound(上限)
    /// </summary>
    public class DesirabilityConfigItem : BindableBase
    {
        public string ResponseName { get; set; } = "";
        public double DataMin { get; set; } = double.NaN;
        public double DataMax { get; set; } = double.NaN;

        private string _goal = "最大化";
        private double _targetValue;   // 目标值 (d=1 的位置)
        private double _lowerBound;    // 下界 (最大化: d=0; 望目: d=0)
        private double _upperBound;    // 上界 (最小化: d=0; 望目: d=0)
        private int _importance = 3;
        private double _shape = 1.0;
        private PlotModel? _previewPlot;

        // ── 绑定属性 ──

        public string Goal
        {
            get => _goal;
            set
            {
                if (SetProperty(ref _goal, value))
                {
                    RaisePropertyChanged(nameof(IsMaximize));
                    RaisePropertyChanged(nameof(IsMinimize));
                    RaisePropertyChanged(nameof(IsTarget));
                    UpdatePreview();
                }
            }
        }

        public double TargetValue
        {
            get => _targetValue;
            set { if (SetProperty(ref _targetValue, value)) UpdatePreview(); }
        }

        public double LowerBound
        {
            get => _lowerBound;
            set { if (SetProperty(ref _lowerBound, value)) UpdatePreview(); }
        }

        public double UpperBound
        {
            get => _upperBound;
            set { if (SetProperty(ref _upperBound, value)) UpdatePreview(); }
        }

        public int Importance
        {
            get => _importance;
            set => SetProperty(ref _importance, value);
        }

        public double Shape
        {
            get => _shape;
            set { if (SetProperty(ref _shape, value)) UpdatePreview(); }
        }

        public PlotModel? PreviewPlot
        {
            get => _previewPlot;
            set => SetProperty(ref _previewPlot, value);
        }

        // ── 可见性 ──

        public bool IsMaximize => Goal == "最大化";
        public bool IsMinimize => Goal == "最小化";
        public bool IsTarget => Goal == "望目";

        // ── 显示文本 ──

        public string DataRangeText =>
            !double.IsNaN(DataMin) ? $"数据范围: [{DataMin:F2}, {DataMax:F2}]" : "";

        public string ShapeHint =>
            _shape < 0.95 ? "凸形（容易满足）" :
            _shape > 1.05 ? "凹形（严格要求）" : "线性";

        public bool IsDefaultValues =>
            Math.Abs(_lowerBound) < 1e-6 && Math.Abs(_upperBound - 100) < 1e-6;

        public static List<string> GoalOptions => new() { "最大化", "最小化", "望目" };

        // ── 构造 ──

        public DesirabilityConfigItem(DesirabilityResponseConfig cfg)
        {
            ResponseName = cfg.ResponseName;
            _importance = cfg.Importance;
            _shape = cfg.Shape;

            _goal = cfg.Goal switch
            {
                DesirabilityGoal.Maximize => "最大化",
                DesirabilityGoal.Minimize => "最小化",
                DesirabilityGoal.Target => "望目",
                _ => "最大化"
            };

            _lowerBound = cfg.Lower;
            _upperBound = cfg.Upper;
            _targetValue = cfg.Target;
        }

        // ── 智能默认值 ──

        public void AutoFillFromDataRange()
        {
            if (double.IsNaN(DataMin) || double.IsNaN(DataMax) || DataMax <= DataMin)
            {
                _lowerBound = 0; _upperBound = 100; _targetValue = 50;
                _shape = 1.0;
                NotifyAll();
                return;
            }

            double range = DataMax - DataMin;
            double margin = range * 0.05;

            if (Goal == "最大化")
            {
                _lowerBound = Math.Round(DataMin - margin, 4);
                _targetValue = Math.Round(DataMax + margin, 4);
                _upperBound = _targetValue;
            }
            else if (Goal == "最小化")
            {
                _targetValue = Math.Round(DataMin - margin, 4);
                _upperBound = Math.Round(DataMax + margin, 4);
                _lowerBound = _targetValue;
            }
            else
            {
                _lowerBound = Math.Round(DataMin - margin, 4);
                _upperBound = Math.Round(DataMax + margin, 4);
                _targetValue = Math.Round((DataMin + DataMax) / 2.0, 4);
            }

            _shape = 1.0;
            NotifyAll();
        }

        private void NotifyAll()
        {
            RaisePropertyChanged(nameof(TargetValue));
            RaisePropertyChanged(nameof(LowerBound));
            RaisePropertyChanged(nameof(UpperBound));
            RaisePropertyChanged(nameof(Shape));
            RaisePropertyChanged(nameof(ShapeHint));
        }

        // ── 校验 ──

        public string? Validate()
        {
            if (Goal == "最大化")
            {
                if (_lowerBound >= _targetValue)
                    return "最低可接受值必须小于目标值";
            }
            else if (Goal == "最小化")
            {
                if (_targetValue >= _upperBound)
                    return "目标值必须小于最高可接受值";
            }
            else
            {
                if (_lowerBound >= _targetValue)
                    return "下限必须小于目标值";
                if (_targetValue >= _upperBound)
                    return "目标值必须小于上限";
            }
            return null;
        }

        // ── 预览曲线 ──

        public void UpdatePreview()
        {
            double lower, upper, target;
            GetEffectiveBounds(out lower, out upper, out target);

            var model = new PlotModel { PlotMargins = new OxyThickness(40, 5, 10, 30) };
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "y", FontSize = 9 });
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "d(y)",
                Minimum = -0.05,
                Maximum = 1.05,
                FontSize = 9
            });

            double range = upper - lower;
            if (range < 1e-6) range = 1.0;
            double margin = range * 0.15;
            double yMin = lower - margin;
            double yMax = upper + margin;

            var series = new LineSeries { Color = OxyColors.SteelBlue, StrokeThickness = 2 };
            for (int i = 0; i <= 100; i++)
            {
                double y = yMin + (yMax - yMin) * i / 100.0;
                series.Points.Add(new DataPoint(y, ComputeD(y, lower, upper, target)));
            }
            model.Series.Add(series);

            // 标注线
            model.Annotations.Add(new LineAnnotation
            {
                Type = LineAnnotationType.Vertical,
                X = lower,
                Color = OxyColors.Red,
                LineStyle = LineStyle.Dot,
                StrokeThickness = 1,
                Text = "L",
                TextColor = OxyColors.Red,
                FontSize = 9
            });
            if (Goal == "望目")
            {
                model.Annotations.Add(new LineAnnotation
                {
                    Type = LineAnnotationType.Vertical,
                    X = target,
                    Color = OxyColors.Orange,
                    LineStyle = LineStyle.Dash,
                    StrokeThickness = 1.5,
                    Text = "T",
                    TextColor = OxyColors.Orange,
                    FontSize = 9
                });
            }
            model.Annotations.Add(new LineAnnotation
            {
                Type = LineAnnotationType.Vertical,
                X = upper,
                Color = OxyColors.Blue,
                LineStyle = LineStyle.Dot,
                StrokeThickness = 1,
                Text = "U",
                TextColor = OxyColors.Blue,
                FontSize = 9
            });

            PreviewPlot = model;
        }

        private void GetEffectiveBounds(out double lower, out double upper, out double target)
        {
            if (Goal == "最大化")
            {
                lower = _lowerBound;
                target = _targetValue;
                upper = _targetValue;
            }
            else if (Goal == "最小化")
            {
                lower = _targetValue;
                target = _targetValue;
                upper = _upperBound;
            }
            else
            {
                lower = _lowerBound;
                target = _targetValue;
                upper = _upperBound;
            }
        }

        private double ComputeD(double y, double lower, double upper, double target)
        {
            if (Goal == "最大化")
            {
                if (y <= lower) return 0;
                if (y >= target) return 1;
                double denom = target - lower;
                return denom < 1e-12 ? 1 : Math.Pow((y - lower) / denom, _shape);
            }
            else if (Goal == "最小化")
            {
                if (y >= upper) return 0;
                if (y <= target) return 1;
                double denom = upper - target;
                return denom < 1e-12 ? 1 : Math.Pow((upper - y) / denom, _shape);
            }
            else
            {
                if (y < lower || y > upper) return 0;
                if (y <= target)
                {
                    double denom = target - lower;
                    return denom < 1e-12 ? 1 : Math.Pow((y - lower) / denom, _shape);
                }
                else
                {
                    double denom = upper - target;
                    return denom < 1e-12 ? 1 : Math.Pow((upper - y) / denom, _shape);
                }
            }
        }

        // ── 转为 Config ──

        public DesirabilityResponseConfig ToConfig()
        {
            GetEffectiveBounds(out double lower, out double upper, out double target);

            return new DesirabilityResponseConfig
            {
                ResponseName = ResponseName,
                Goal = Goal switch
                {
                    "最大化" => DesirabilityGoal.Maximize,
                    "最小化" => DesirabilityGoal.Minimize,
                    "望目" => DesirabilityGoal.Target,
                    _ => DesirabilityGoal.Maximize
                },
                Lower = lower,
                Upper = upper,
                Target = target,
                Weight = 1.0,
                Importance = Importance,
                Shape = _shape,
                ShapeLower = _shape,
                ShapeUpper = _shape
            };
        }
    }
}