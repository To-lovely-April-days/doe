using MaxChemical.Modules.DOE.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace MaxChemical.Modules.DOE.Views
{
    public partial class DOEHistoryView : UserControl
    {
        public DOEHistoryView() { InitializeComponent(); }

        private void BatchList_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is DataGrid dg && dg.SelectedItem is DOEBatchSummary summary)
            {
                if (Window.GetWindow(this)?.DataContext is ViewModels.DOEMainViewModel mainVm)
                    mainVm.NavigateToExecution(summary.BatchId);
            }
        }
    }
}