using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace MaxChemical.Modules.DOE.Views
{
    public partial class DOEMiniPanel : UserControl
    {
        public DOEMiniPanel()
        {
            InitializeComponent();
        }

        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                var window = Window.GetWindow(this);
                window?.DragMove();
            }
        }
    }
}