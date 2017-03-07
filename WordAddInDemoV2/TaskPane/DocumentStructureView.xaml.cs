using System.Windows;
using System.Windows.Controls;

namespace WordAddInDemoV2.TaskPane
{
    /// <summary>
    /// Interaction logic for DocumentStructureView.xaml
    /// </summary>
    public partial class DocumentStructureView : UserControl
    {
        public DocumentStructureView()
        {
            InitializeComponent();
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Button is clicked!");
        }
    }
}
