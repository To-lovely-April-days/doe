using MaxChemical.Modules.DOE.Models;
using MaxChemical.Modules.DOE.ViewModels;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace MaxChemical.Modules.DOE.Views
{
    public partial class DOEModelAnalysisView : UserControl
    {
        private bool _isDraggingProfiler;
        private string? _draggingFactorName;
        private OxyPlot.Wpf.PlotView? _draggingPlotView;
        private bool _draggingIsCategorical;
        // ── 意愿行拖拽状态 ──
        private bool _isDraggingDesirability;
        private string? _draggingDesirabilityFactorName;
        private OxyPlot.Wpf.PlotView? _draggingDesirabilityPlotView;
        public DOEModelAnalysisView() { InitializeComponent(); }
        // ── 意愿图控制点拖拽状态 ──
        private bool _isDraggingControlPoint;
        private int _draggingControlPointIndex = -1; // 0=Lower, 1=Mid, 2=Target
        private OxyPlot.Wpf.PlotView? _controlPointPlotView;
        private void GprTab_Click(object sender, MouseButtonEventArgs e)
        { if (DataContext is DOEModelAnalysisViewModel vm) vm.SwitchToGprTabCommand.Execute(); }
        private void OlsTab_Click(object sender, MouseButtonEventArgs e)
        { if (DataContext is DOEModelAnalysisViewModel vm) vm.SwitchToOlsTabCommand.Execute(); }
        private void ModelTab_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is GPRModelTabItem tab)
                if (DataContext is DOEModelAnalysisViewModel vm) vm.SelectModel(tab.ModelId);
        }
        private void OlsBatchItem_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is OlsBatchItem item)
                if (DataContext is DOEModelAnalysisViewModel vm) vm.SelectOlsBatchCommand.Execute(item);
        }
        private void TrainingData_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e.Column is DataGridTextColumn col)
                switch (e.PropertyName)
                {
                    case "#": col.Width = new DataGridLength(40); break;
                    case "来源": col.Width = new DataGridLength(70); break;
                    case "批次名称": col.Width = new DataGridLength(160); break;
                    case "时间": col.Width = new DataGridLength(130); break;
                    case "响应值": col.Width = new DataGridLength(70); break;
                    default: col.Width = new DataGridLength(1, DataGridLengthUnitType.Star); break;
                }
        }
        private void OlsPareto_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (DataContext is not DOEModelAnalysisViewModel vm) return;
            if (sender is not OxyPlot.Wpf.PlotView plotView) return;
            var model = plotView.Model;
            if (model == null) return;
            var pos = e.GetPosition(plotView);
            var barSeries = model.Series.OfType<BarSeries>().FirstOrDefault();
            if (barSeries == null) return;
            var result = barSeries.GetNearestPoint(new ScreenPoint(pos.X, pos.Y), false);
            if (result != null)
            {
                int index = (int)Math.Round(result.DataPoint.X);
                if (index >= 0) { vm.ToggleParetoTerm(index); e.Handled = true; }
            }
        }

        // ═══════════════════════════════════════════════════════
        // ★ v6: 预测刻画器拖动 — 连续和类别因子统一拖动
        // ═══════════════════════════════════════════════════════

        private void Profiler_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (DataContext is not DOEModelAnalysisViewModel vm) return;
            if (sender is not OxyPlot.Wpf.PlotView plotView) return;
            if (!vm.IsProfilerLoaded) return;

            var factorName = plotView.Model?.Title;
            if (string.IsNullOrEmpty(factorName)) return;

            // ★ 连续和类别因子统一走拖动逻辑
            _isDraggingProfiler = true;
            _draggingFactorName = factorName;
            _draggingPlotView = plotView;
            _draggingIsCategorical = vm.IsProfilerFactorCategorical(factorName);
            plotView.CaptureMouse();

            // 立即移动十字线到点击位置
            MoveCrosshair(plotView, factorName, e, vm);
            e.Handled = true;
        }

        private void Profiler_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDraggingProfiler || _draggingFactorName == null || _draggingPlotView == null) return;
            if (DataContext is not DOEModelAnalysisViewModel vm) return;
            MoveCrosshair(_draggingPlotView, _draggingFactorName, e, vm);
        }

        private async void Profiler_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isDraggingProfiler) return;
            if (DataContext is not DOEModelAnalysisViewModel vm) return;

            var factorName = _draggingFactorName!;
            var plotView = _draggingPlotView!;
            var isCat = _draggingIsCategorical;

            plotView.ReleaseMouseCapture();
            _isDraggingProfiler = false;
            _draggingFactorName = null;
            _draggingPlotView = null;

            if (isCat)
            {
                // ★ 类别因子: X 四舍五入到最近的水平索引 → 传水平名
                var levels = vm.GetProfilerCategoryLevels(factorName);
                if (levels == null || levels.Count == 0) return;

                var rawX = GetRawDataX(plotView, e);
                if (!rawX.HasValue) return;

                int index = (int)Math.Round(rawX.Value);
                index = Math.Max(0, Math.Min(levels.Count - 1, index));

                await vm.UpdateProfilerCategoryAsync(factorName, levels[index]);
            }
            else
            {
                // ★ 连续因子: 传数值
                var newX = ClampXToRange(plotView, e, vm, factorName);
                if (!newX.HasValue) return;

                await vm.UpdateProfilerValueAsync(factorName, newX.Value);
            }
        }

        /// <summary>
        /// 移动十字线 — 连续因子和类别因子统一处理
        /// </summary>
        private void MoveCrosshair(OxyPlot.Wpf.PlotView plotView, string factorName,
            MouseEventArgs e, DOEModelAnalysisViewModel vm)
        {
            var model = plotView.Model;
            if (model == null) return;

            double? newX;
            double? newY;

            if (_draggingIsCategorical)
            {
                // 类别因子: X 吸附到最近的整数水平
                var rawX = GetRawDataX(plotView, e);
                if (!rawX.HasValue) return;

                var levels = vm.GetProfilerCategoryLevels(factorName);
                if (levels == null) return;

                int index = (int)Math.Round(rawX.Value);
                index = Math.Max(0, Math.Min(levels.Count - 1, index));
                newX = index;

                // Y = 该水平对应的预测值
                newY = vm.GetProfilerCategoryYAtIndex(factorName, index);
            }
            else
            {
                // 连续因子: X 限因子范围
                newX = ClampXToRange(plotView, e, vm, factorName);
                if (!newX.HasValue) return;

                // Y = 曲线插值
                newY = vm.GetProfilerYAtX(factorName, newX.Value);
            }

            // 更新红色竖线
            var vline = model.Annotations
                .OfType<OxyPlot.Annotations.LineAnnotation>()
                .FirstOrDefault(a => a.Type == OxyPlot.Annotations.LineAnnotationType.Vertical
                                     && a.Color == OxyColors.Red);
            if (vline != null)
            {
                vline.X = newX.Value;
                if (!_draggingIsCategorical)
                    vline.Text = $"{newX.Value:F1}";
            }

            // 更新红色横线
            var hline = model.Annotations
                .OfType<OxyPlot.Annotations.LineAnnotation>()
                .FirstOrDefault(a => a.Type == OxyPlot.Annotations.LineAnnotationType.Horizontal);
            if (hline != null && newY.HasValue)
                hline.Y = newY.Value;

            // 更新顶部预测值
            if (newY.HasValue)
            {
                if (_draggingIsCategorical)
                {
                    var levels = vm.GetProfilerCategoryLevels(factorName);
                    int idx = (int)Math.Round(newX.Value);
                    var levelName = (levels != null && idx >= 0 && idx < levels.Count) ? levels[idx] : "?";
                    vm.ProfilerCurrentPredicted = $"{factorName}={levelName}  →  预测值={newY.Value:F4}";
                }
                else
                {
                    vm.ProfilerCurrentPredicted = $"{factorName}={newX.Value:F2}  →  预测值={newY.Value:F4}";
                }
            }

            model.InvalidatePlot(false);
        }

        /// <summary>
        /// 获取原始 X 数据坐标（不限范围）
        /// </summary>
        private double? GetRawDataX(OxyPlot.Wpf.PlotView plotView, MouseEventArgs e)
        {
            var model = plotView.Model;
            if (model == null) return null;
            var xAxis = model.Axes.OfType<LinearAxis>()
                .FirstOrDefault(a => a.Position == AxisPosition.Bottom);
            if (xAxis == null) return null;
            var pos = e.GetPosition(plotView);
            return xAxis.InverseTransform(pos.X);
        }

        /// <summary>
        /// 获取 X 数据坐标，限制在因子范围内（连续因子用）
        /// </summary>
        private double? ClampXToRange(OxyPlot.Wpf.PlotView plotView, MouseEventArgs e,
            DOEModelAnalysisViewModel vm, string factorName)
        {
            var rawX = GetRawDataX(plotView, e);
            if (!rawX.HasValue) return null;
            var (fMin, fMax) = vm.GetProfilerFactorRange(factorName);
            double dataX = Math.Max(fMin, Math.Min(fMax, rawX.Value));
            return Math.Round(dataX, 4);
        }

        private void ToggleEquationCoding_Click(object sender, MouseButtonEventArgs e)
        {
            if (DataContext is DOEModelAnalysisViewModel vm)
            {
                vm.IsCodedEquation = !vm.IsCodedEquation;
            }
        }
        private void DesirabilityMenu_Click(object sender, RoutedEventArgs e)
        {
            var menu = new ContextMenu();
            var vm = DataContext as DOEModelAnalysisViewModel;
            if (vm == null) return;

            var item1 = new MenuItem { Header = "意愿函数", IsCheckable = true, IsChecked = vm.IsDesirabilityVisible };
            item1.Click += (s, _) => vm.ToggleDesirabilityCommand.Execute();
            menu.Items.Add(item1);

            menu.Items.Add(new Separator());

            var item2 = new MenuItem { Header = "最大化意愿" };
            item2.Click += (s, _) => vm.MaximizeDesirabilityCommand.Execute();
            menu.Items.Add(item2);

            var item3 = new MenuItem { Header = "最大化并记住" };
            item3.Click += (s, _) => vm.MaximizeAndRememberCommand.Execute();
            menu.Items.Add(item3);

            menu.Items.Add(new Separator());

            var item6 = new MenuItem { Header = "保存意愿" };
            item6.Click += (s, _) => vm.SaveDesirabilityCommand.Execute();
            menu.Items.Add(item6);

            var item7 = new MenuItem { Header = "设置意愿..." };
            item7.Click += (s, _) => vm.SetDesirabilityCommand.Execute();
            menu.Items.Add(item7);

            var item8 = new MenuItem { Header = "保存意愿公式" };
            item8.Click += (s, _) => vm.SaveDesirabilityFormulaCommand.Execute();
            menu.Items.Add(item8);

            menu.IsOpen = true;
        }

        /// <summary>
        /// ★ Feature 2: 意愿行鼠标按下 — 判断是否点击了红色竖线附近
        /// </summary>
        private void DesirabilityProfiler_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (DataContext is not DOEModelAnalysisViewModel vm) return;
            if (sender is not OxyPlot.Wpf.PlotView plotView) return;

            var parent = FindParentItemsControl(plotView);
            if (parent == null) return;
            int index = FindPlotViewIndex(parent, plotView);
            if (index < 0) return;

            var factorName = vm.GetDesirabilityFactorName(index);

            // ★ Feature 3: 最后一个图是意愿图，双击弹出设置弹窗
            if (factorName == null)
            {
                if (e.ClickCount == 2)
                {
                    vm.SetDesirabilityCommand.Execute();
                }
                return;
            }

            // 检查是否为类别因子
            bool isCat = vm.IsProfilerFactorCategorical(factorName);
            if (isCat)
            {
                // 类别因子: 点击切换水平（跟响应行一样）
                var levels = vm.GetProfilerCategoryLevels(factorName);
                if (levels != null && levels.Count > 0)
                {
                    var rawX = GetRawDataX(plotView, e);
                    if (rawX.HasValue)
                    {
                        int idx = Math.Max(0, Math.Min(levels.Count - 1, (int)Math.Round(rawX.Value)));
                        _ = vm.UpdateProfilerCategoryAsync(factorName, levels[idx]);
                    }
                }
                return;
             
            }

            // 连续因子: 开始拖拽
            _isDraggingDesirability = true;
            _draggingDesirabilityFactorName = factorName;
            _draggingDesirabilityPlotView = plotView;
            plotView.CaptureMouse();
            e.Handled = true;
        }

        private void DesirabilityProfiler_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDraggingDesirability || _draggingDesirabilityFactorName == null) return;
            if (_draggingDesirabilityPlotView == null) return;
            if (DataContext is not DOEModelAnalysisViewModel vm) return;

            var position = e.GetPosition(_draggingDesirabilityPlotView);
            var plotModel = _draggingDesirabilityPlotView.Model;
            if (plotModel == null) return;

            var xAxis = plotModel.Axes.FirstOrDefault(a => a.Position == OxyPlot.Axes.AxisPosition.Bottom);
            if (xAxis == null) return;

            double newX = xAxis.InverseTransform(position.X);
            var (min, max) = vm.GetProfilerFactorRange(_draggingDesirabilityFactorName);
            newX = Math.Max(min, Math.Min(max, newX));

            // ★ 移动红线 + 更新文字
            var vline = plotModel.Annotations
                .OfType<OxyPlot.Annotations.LineAnnotation>()
                .FirstOrDefault(a => a.Type == OxyPlot.Annotations.LineAnnotationType.Vertical
                                     && a.Color == OxyColors.Red);
            if (vline != null)
            {
                vline.X = newX;
            }

            // ★ 更新 X 轴 Title（需要用 true 触发完整重绘）
            xAxis.Title = $"{newX:F2}";

            plotModel.InvalidatePlot(true);
        }


        private async void DesirabilityProfiler_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isDraggingDesirability) return;
            if (DataContext is not DOEModelAnalysisViewModel vm) return;

            var factorName = _draggingDesirabilityFactorName!;
            var plotView = _draggingDesirabilityPlotView!;

            plotView.ReleaseMouseCapture();
            _isDraggingDesirability = false;
            _draggingDesirabilityFactorName = null;
            _draggingDesirabilityPlotView = null;

            // ★ 松开时才调用 Python 更新
            var rawX = GetRawDataX(plotView, e);
            if (!rawX.HasValue) return;

            var (min, max) = vm.GetProfilerFactorRange(factorName);
            double newX = Math.Max(min, Math.Min(max, rawX.Value));

            await vm.UpdateProfilerValueAsync(factorName, newX);
        }

        /// <summary>
        /// 找到 PlotView 所在的 ItemsControl
        /// </summary>
        private ItemsControl? FindParentItemsControl(DependencyObject child)
        {
            var parent = VisualTreeHelper.GetParent(child);
            while (parent != null)
            {
                if (parent is ItemsControl ic && ic != ProfilerItemsControl)
                    return ic;
                parent = VisualTreeHelper.GetParent(parent);
            }
            return null;
        }

        /// <summary>
        /// 找到 PlotView 在 ItemsControl 中的索引
        /// </summary>
        private int FindPlotViewIndex(ItemsControl itemsControl, OxyPlot.Wpf.PlotView plotView)
        {
            for (int i = 0; i < itemsControl.Items.Count; i++)
            {
                var container = itemsControl.ItemContainerGenerator.ContainerFromIndex(i) as ContentPresenter;
                if (container != null)
                {
                    var pv = FindChild<OxyPlot.Wpf.PlotView>(container);
                    if (pv == plotView) return i;
                }
            }
            return -1;
        }

        private static T? FindChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T t) return t;
                var result = FindChild<T>(child);
                if (result != null) return result;
            }
            return null;
        }

        private void HandleDesirabilityFunctionMouseDown(DOEModelAnalysisViewModel vm,
    OxyPlot.Wpf.PlotView plotView, MouseButtonEventArgs e)
        {
            var position = e.GetPosition(plotView);
            var plotModel = plotView.Model;
            if (plotModel == null) return;

            var xAxis = plotModel.Axes.FirstOrDefault(a => a.Position == AxisPosition.Bottom);
            var yAxis = plotModel.Axes.FirstOrDefault(a => a.Position == AxisPosition.Left);
            if (xAxis == null || yAxis == null) return;

            // 检查鼠标是否在控制点附近（20像素容差）
            var cfg = vm.GetCurrentDesirabilityConfig();
            if (cfg == null) return;

            double[][] controlPts = new[]
            {
        new[] { cfg.Lower, 0.0 },
        new[] { (cfg.Lower + cfg.Target) / 2.0, 0.5 },
        new[] { cfg.Target, cfg.Goal == DesirabilityGoal.Maximize ? 1.0 : 0.0 }
    };

            for (int i = 0; i < controlPts.Length; i++)
            {
                var screenPt = xAxis.Transform(controlPts[i][0], controlPts[i][1], yAxis);
                double dist = Math.Sqrt(Math.Pow(position.X - screenPt.X, 2) + Math.Pow(position.Y - screenPt.Y, 2));
                if (dist < 20)
                {
                    _isDraggingControlPoint = true;
                    _draggingControlPointIndex = i;
                    _controlPointPlotView = plotView;
                    plotView.CaptureMouse();
                    e.Handled = true;
                    return;
                }
            }
        }

        private void HandleDesirabilityFunctionMouseMove(DOEModelAnalysisViewModel vm, MouseEventArgs e)
        {
            if (!_isDraggingControlPoint || _controlPointPlotView == null) return;

            var position = e.GetPosition(_controlPointPlotView);
            var plotModel = _controlPointPlotView.Model;
            if (plotModel == null) return;

            var xAxis = plotModel.Axes.FirstOrDefault(a => a.Position == AxisPosition.Bottom);
            if (xAxis == null) return;

            // 计算新的 Y 值（响应值）
            double newY = xAxis.InverseTransform(position.X);

            // 根据拖拽的是哪个点，更新配置
            switch (_draggingControlPointIndex)
            {
                case 0: // Lower
                    vm.UpdateDesirabilityLower(newY);
                    break;
                case 1: // Mid（调整 shape 参数）
                    vm.UpdateDesirabilityMidPoint(newY);
                    break;
                case 2: // Target
                    vm.UpdateDesirabilityTarget(newY);
                    break;
            }

            // 刷新意愿图和意愿行
            vm.RefreshDesirabilityFunctionPlot();
            vm.RefreshDesirabilityPlotsInPlace();
        }
    }


}