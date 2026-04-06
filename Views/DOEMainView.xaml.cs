using MaxChemical.Modules.DOE.Events;
using MaxChemical.Modules.DOE.ViewModels;
using Prism.Events;
using Prism.Ioc;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace MaxChemical.Modules.DOE.Views
{
    public partial class DOEMainView : Window
    {
        private DOEOverviewViewModel? _overviewVm;
        private DOEExecutionDashboardViewModel? _execVm;
        private DOEModelAnalysisViewModel? _modelAnalysisVm;
        private DOEHistoryViewModel? _historyVm;
        private DOEMainViewModel? _mainVm;

        // ── 迷你模式状态保存 ──
        private double _savedWidth, _savedHeight, _savedLeft, _savedTop;
        private WindowState _savedWindowState;
        private const double MINI_WIDTH = 320;
        private const double MINI_HEIGHT = 240;
        private static readonly Duration AnimDuration = new(TimeSpan.FromMilliseconds(300));
        private static readonly CubicEase AnimEase = new() { EasingMode = EasingMode.EaseInOut };
        private static readonly string[] LoadingMessages = new[]
{
    "正在启动 Python 运行时...",
    "正在加载统计分析引擎...",
    "正在初始化 GPR 模型服务...",
    "正在加载项目数据..."
};
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

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            // ★ 修复: 关闭窗口时取消事件订阅，防止下次打开时重复弹窗
            if (_mainVm != null)
                _mainVm.Dispose();
            Close();
        }

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
                            // ★ 如果是从历史页跳转过来（CurrentBatchId 已设置），
                            //    不重复加载，NavigateToOlsBatchAsync 已经处理了
                            //    只有手动点 Tab 且没有 CurrentBatchId 时才 LoadAsync
                            if (_modelAnalysisVm != null && string.IsNullOrEmpty(_mainVm?.CurrentBatchId))
                            {
                                await _modelAnalysisVm.LoadAsync();
                            }
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

                // ── Step 1: Python 运行时 ──
                UpdateLoadingStep(0);
                await Task.Run(() =>
                {
                    try
                    {
                        var env = MaxChemical.Modules.Designer.Services.PythonEnvironmentManager.Instance;
                        if (!env.IsInitialized) env.Initialize();
                    }
                    catch { }
                });

                // ── Step 2: 统计分析引擎 ──
                UpdateLoadingStep(1);
                _overviewVm = container.Resolve<DOEOverviewViewModel>();
                _execVm = container.Resolve<DOEExecutionDashboardViewModel>();

                // ── Step 3: GPR 模型服务 ──
                UpdateLoadingStep(2);
                _modelAnalysisVm = container.Resolve<DOEModelAnalysisViewModel>();
                _historyVm = container.Resolve<DOEHistoryViewModel>();

                // ── Step 4: 项目数据 ──
                UpdateLoadingStep(3);

                OverviewView.DataContext = _overviewVm;
                ExecutionView.DataContext = _execVm;
                ModelAnalysisView.DataContext = _modelAnalysisVm;
                HistoryView.DataContext = _historyVm;
                MiniPanel.DataContext = _execVm;

                _execVm.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(DOEExecutionDashboardViewModel.IsMiniMode))
                    {
                        if (_execVm.IsMiniMode) SwitchToMiniMode();
                        else SwitchToNormalMode();
                    }
                };

                mainVm.RequestLoadExecution += async (s, id) =>
                {
                    if (_execVm != null) await _execVm.LoadBatchAsync(id);
                };
                mainVm.RequestLoadAnalysis += async (s, id) =>
                {
                    if (_modelAnalysisVm != null) await _modelAnalysisVm.LoadBatchAsync(id);
                };
                mainVm.RequestRefreshHistory += async (s, e) =>
                {
                    if (_historyVm != null) await _historyVm.LoadBatchesAsync();
                };
                mainVm.RequestRefreshOverview += async (s, e) =>
                {
                    if (_overviewVm != null) await _overviewVm.LoadAsync();
                };

                if (_historyVm != null)
                {
                    _historyVm.RequestExecuteBatch += (s, id) => mainVm.NavigateToExecution(id);
                    _historyVm.RequestAnalyzeBatch += (s, id) =>
                    {
                        mainVm.CurrentBatchId = id;
                        mainVm.SelectedTabIndex = 2;
                        // ★ 直接跳到 OLS Tab 并加载该批次的分析
                        _ = _modelAnalysisVm?.NavigateToOlsBatchAsync(id);
                    };
                }

                if (_overviewVm != null)
                {
                    _overviewVm.RequestResumeBatch += (s, id) => mainVm.NavigateToExecution(id);
                    _overviewVm.RequestGoToHistory += (s, e) => mainVm.SelectedTabIndex = 3;
                    _overviewVm.RequestContinueProject += (s, projectId) =>
                    {
                        _ = mainVm.ContinueProjectAsync(projectId);
                    };
                    _overviewVm.RequestViewProject += (s, projectId) =>
                    {
                        mainVm.SelectedTabIndex = 3;
                    };
                }
                _overviewVm.RequestViewAnalysis += (s, projectId) =>
                {
                    _ = mainVm.NavigateToAnalysisByProjectAsync(projectId);
                };
                // ── 完成 ──
                UpdateLoadingStep(4, "准备就绪");
                await Task.Delay(200);

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
      
        private void UpdateLoadingStep(int completedSteps, string? overrideText = null)
        {
            // 更新进度条宽度（总宽 220）
            var targetWidth = 220.0 * completedSteps / 4.0;
            LoadingProgressBar.Width = targetWidth;

            // 更新状态文字
            if (overrideText != null)
            {
                LoadingStatusText.Text = overrideText;
            }
            else if (completedSteps < LoadingMessages.Length)
            {
                LoadingStatusText.Text = LoadingMessages[completedSteps];
            }
            else
            {
                LoadingStatusText.Text = "准备就绪";
            }
        }
        private void ProjectMenuBtn_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.ContextMenu != null)
            {
                btn.ContextMenu.PlacementTarget = btn;
                btn.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
                btn.ContextMenu.DataContext = DataContext;
                btn.ContextMenu.IsOpen = true;
            }
        }

        // ══════════════ 迷你模式切换 ══════════════

        private void SwitchToMiniMode()
        {
            _savedWindowState = WindowState;
            if (WindowState == WindowState.Maximized)
                WindowState = WindowState.Normal;

            _savedWidth = ActualWidth;
            _savedHeight = ActualHeight;
            _savedLeft = Left;
            _savedTop = Top;

            MainContent.Visibility = Visibility.Collapsed;
            MiniPanel.Visibility = Visibility.Visible;
            MiniPanel.Opacity = 0;

            SizeToContent = SizeToContent.Height;
            Width = MINI_WIDTH;
            var screen = SystemParameters.WorkArea;
            Left = screen.Right - MINI_WIDTH - 20;
            Top = screen.Bottom - 260 - 20;

            Topmost = true;
            ResizeMode = ResizeMode.NoResize;

            MiniPanel.BeginAnimation(OpacityProperty,
                new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(150))));
        }

        private void SwitchToNormalMode()
        {
            var fadeOut = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(120)));
            fadeOut.Completed += (s, e) =>
            {
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