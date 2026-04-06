using System.Windows;
using System.Windows.Input;

namespace MaxChemical.Modules.DOE.Views
{
    public partial class DOEProjectDetailDialog : Window
    {
        public DOEProjectDetailDialog()
        {
            InitializeComponent();
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

    

        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                DragMove();
        }
    }
}