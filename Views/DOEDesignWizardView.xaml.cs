using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MaxChemical.Infrastructure.DOE;
using MaxChemical.Modules.DOE.ViewModels;

namespace MaxChemical.Modules.DOE.Views
{
    public partial class DOEDesignWizardView : Window
    {
        public DOEDesignWizardView(DOEDesignWizardViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;

            // 订阅保存完成事件，自动关闭窗口
            viewModel.BatchSaved += (sender, batchId) =>
            {
                MessageBox.Show($"DOE 方案已保存成功！\n批次ID: {batchId}", "保存成功",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                DialogResult = true;
                Close();
            };
        }

       
    }
}
