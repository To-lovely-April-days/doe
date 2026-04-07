using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MaxChemical.Modules.DOE.ViewModels;

namespace MaxChemical.Modules.DOE.Views
{
    public partial class DOEDesignWizardView : Window
    {
        public DOEDesignWizardView(DOEDesignWizardViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;

            // Esc 键关闭窗口
            PreviewKeyDown += (s, e) =>
            {
                if (e.Key == Key.Escape)
                {
                    e.Handled = true;
                    TryClose();
                }
            };

            // 订阅保存完成事件，自动关闭窗口
            viewModel.BatchSaved += (sender, batchId) =>
            {
                MessageBox.Show($"DOE 方案已保存成功！\n批次ID: {batchId}", "保存成功",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                DialogResult = true;
                Close();
            };
        }

        /// <summary>
        /// 关闭按钮点击
        /// </summary>
        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            TryClose();
        }

        /// <summary>
        /// 确认关闭（有数据时提示）
        /// </summary>
        private void TryClose()
        {
            if (DataContext is DOEDesignWizardViewModel vm && vm.CurrentStep > 1)
            {
                var result = MessageBox.Show("确定要关闭向导吗？未保存的配置将丢失。",
                    "确认关闭", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result != MessageBoxResult.Yes)
                    return;
            }
            DialogResult = false;
            Close();
        }

        /// <summary>
        /// 拖拽标题栏移动窗口
        /// </summary>
        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                DragMove();
        }

        /// <summary>
        /// 左侧导航步骤点击跳转
        /// </summary>
        private void NavStep_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.Tag is string tagStr && int.TryParse(tagStr, out int step))
            {
                if (DataContext is DOEDesignWizardViewModel vm)
                {
                    // 只允许跳到已完成的步骤或当前步骤+1
                    if (step <= vm.CurrentStep)
                    {
                        vm.CurrentStep = step;
                    }
                }
            }
        }
    }
}