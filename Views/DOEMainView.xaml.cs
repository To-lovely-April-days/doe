using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Animation;
using MaxChemical.Modules.DOE.Events;
using MaxChemical.Modules.DOE.ViewModels;
using Prism.Events;
using Prism.Ioc;

namespace MaxChemical.Modules.DOE.Views
{
    public partial class DOEMainView : Window
    {
        private DOEOverviewViewModel? _overviewVm;
        private DOEExecutionDashboardViewModel? _execVm;
        private DOEModelAnalysisViewModel? _modelAnalysisVm;
        private DOEHistoryViewModel? _historyVm;
        private DOEMainViewModel? _mainVm;

        // ──  迷你模式状态保存 ──
        private double _savedWidth, _savedHeight, _savedLeft, _savedTop;
        private WindowState _savedWindowState;
        private const double MINI_WIDTH = 320;
        private const double MINI_HEIGHT = 240;
        private static readonly Duration AnimDuration = new(TimeSpan.FromMilliseconds(300));
        private static readonly CubicEase AnimEase = new() { EasingMode = EasingMode.EaseInOut };

        public DOEMainView(DOEMainViewModel mainVm, IContainerProvider container)
        {
            InitializeComponent();
            DataContext = mainVm;
            _mainVm = mainVm;

            mainVm.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(DOEMainViewModel.SelectedTabIndex))
                {
                    int idx = mainVm.SelectedTabIndex;
                    foreach (var child in TabBar.Children)
                    {
                        if (child is RadioButton rb && rb.Tag is string t && int.TryParse(t, out int i))
                            rb.IsChecked = (i == idx);
                    }
                }
            };

            Loaded += async (s, e) => await InitAsync(mainVm, container);
        }

        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2) return;
            DragMove();
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e) => Close();

        private async void TabRadio_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton rb && rb.Tag is string tagStr && int.TryParse(tagStr, out int index))
            {
                if (_mainVm != null)
                    _mainVm.SelectedTabIndex = index;

                try
                {
                    switch (index)
                    {
                        case 0:
                            if (_overviewVm != null) await _overviewVm.LoadAsync();
                            break;
                        case 1:
                            if (_execVm != null && !string.IsNullOrEmpty(_mainVm?.CurrentBatchId))
                                await _execVm.LoadBatchAsync(_mainVm.CurrentBatchId);
                            break;
                        case 2:
                            if (_modelAnalysisVm != null)
                                await _modelAnalysisVm.LoadAsync();
                            break;
                        case 3:
                            if (_historyVm != null) await _historyVm.LoadBatchesAsync();
                            break;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"TabSwitch error: {ex.Message}");
                }
            }
        }

        private async Task InitAsync(DOEMainViewModel mainVm, IContainerProvider container)
        {
            try
            {
                LoadingOverlay.Visibility = Visibility.Visible;

                await Task.Run(() =>
                {
                    try
                    {
                        var env = MaxChemical.Modules.Designer.Services.PythonEnvironmentManager.Instance;
                        if (!env.IsInitialized) env.Initialize();
                    }
                    catch { }
                });

                _overviewVm = container.Resolve<DOEOverviewViewModel>();
                _execVm = container.Resolve<DOEExecutionDashboardViewModel>();
                _modelAnalysisVm = container.Resolve<DOEModelAnalysisViewModel>();
                _historyVm = container.Resolve<DOEHistoryViewModel>();

                OverviewView.DataContext = _overviewVm;
                ExecutionView.DataContext = _execVm;
                ModelAnalysisView.DataContext = _modelAnalysisVm;
                HistoryView.DataContext = _historyVm;

                //  迷你面板绑定到执行 ViewModel
                MiniPanel.DataContext = _execVm;

                //  监听迷你模式切换
                _execVm.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(DOEExecutionDashboardViewModel.IsMiniMode))
                    {
                        if (_execVm.IsMiniMode)
                            SwitchToMiniMode();
                        else
                            SwitchToNormalMode();
                    }
                };

                mainVm.RequestLoadExecution += async (s, id) =>
                {
                    if (_execVm != null)
                        await _execVm.LoadBatchAsync(id);
                };

                mainVm.RequestLoadAnalysis += async (s, id) =>
                {
                    if (_modelAnalysisVm != null)
                        await _modelAnalysisVm.LoadBatchAsync(id);
                };

                mainVm.RequestRefreshHistory += async (s, e) =>
                {
                    if (_historyVm != null) await _historyVm.LoadBatchesAsync();
                };

                if (_historyVm != null)
                {
                    _historyVm.RequestExecuteBatch += (s, id) => mainVm.NavigateToExecution(id);
                    _historyVm.RequestAnalyzeBatch += (s, id) =>
                    {
                        mainVm.CurrentBatchId = id;
                        mainVm.SelectedTabIndex = 2;
                        _ = _modelAnalysisVm?.LoadBatchAsync(id);
                    };
                }

                if (_overviewVm != null)
                {
                    _overviewVm.RequestResumeBatch += (s, id) => mainVm.NavigateToExecution(id);
                    _overviewVm.RequestGoToHistory += (s, e) => mainVm.SelectedTabIndex = 3;
                }

                LoadingOverlay.Visibility = Visibility.Collapsed;

                mainVm.SelectedTabIndex = 0;
                await _overviewVm.LoadAsync();
            }
            catch (Exception ex)
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
                MessageBox.Show($"初始化失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ══════════════  迷你模式切换 ══════════════

        /// <summary>
        /// 切换到迷你模式：保存窗口状态 → 隐藏主内容 → 动画缩小 → 置顶
        /// </summary>
        private void SwitchToMiniMode()
        {
            // 保存当前窗口状态
            _savedWindowState = WindowState;
            if (WindowState == WindowState.Maximized)
                WindowState = WindowState.Normal;

            _savedWidth = ActualWidth;
            _savedHeight = ActualHeight;
            _savedLeft = Left;
            _savedTop = Top;

            // 切换内容
            MainContent.Visibility = Visibility.Collapsed;
            MiniPanel.Visibility = Visibility.Visible;
            MiniPanel.Opacity = 0;

            // 瞬切尺寸和位置
            SizeToContent = SizeToContent.Height;
            Width = MINI_WIDTH;
            var screen = SystemParameters.WorkArea;
            Left = screen.Right - MINI_WIDTH - 20;
            Top = screen.Bottom - 260 - 20;  // 预估高度

            Topmost = true;
            ResizeMode = ResizeMode.NoResize;

            // 淡入
            MiniPanel.BeginAnimation(OpacityProperty,
                new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(150))));
        }

        /// <summary>
        /// 恢复到正常模式：清除动画 → 显示主内容 → 动画放大 → 取消置顶
        /// </summary>
        private void SwitchToNormalMode()
        {

            // 淡出迷你面板
            var fadeOut = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(120)));
            fadeOut.Completed += (s, e) =>
            {
                // 淡出完成后瞬切回大窗口
                MiniPanel.Visibility = Visibility.Collapsed;
                MainContent.Visibility = Visibility.Visible;
                MainContent.Opacity = 0;

                SizeToContent = SizeToContent.Manual;
                Topmost = false;
                ResizeMode = ResizeMode.CanResizeWithGrip;

                Width = _savedWidth;
                Height = _savedHeight;
                Left = _savedLeft;
                Top = _savedTop;

                if (_savedWindowState == WindowState.Maximized)
                    WindowState = WindowState.Maximized;

                // 淡入主内容
                MainContent.BeginAnimation(OpacityProperty,
                    new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(150))));
            };

            MiniPanel.BeginAnimation(OpacityProperty, fadeOut);
        }

        private static DoubleAnimation CreateAnim(double to)
        {
            return new DoubleAnimation(to, AnimDuration) { EasingFunction = AnimEase };
        }
    }
}