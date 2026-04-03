using MaxChemical.Modules.DOE.ViewModels;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace MaxChemical.Modules.DOE.Views
{
    public partial class DOEModelAnalysisView : UserControl
    {
        private bool _isDraggingProfiler;
        private string? _draggingFactorName;
        private OxyPlot.Wpf.PlotView? _draggingPlotView;
        private bool _draggingIsCategorical;

        public DOEModelAnalysisView() { InitializeComponent(); }

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
    }
}