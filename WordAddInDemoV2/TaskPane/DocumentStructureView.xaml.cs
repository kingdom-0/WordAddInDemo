using WordAddInDemoV2.DataContainers;

namespace WordAddInDemoV2.TaskPane
{
    /// <summary>
    /// Interaction logic for DocumentStructureView.xaml
    /// </summary>
    public partial class DocumentStructureView
    {
        public DocumentStructureView()
        {
            InitializeComponent();
            DataContext = ControlsContainer.Instance.ControlItems;
        }
    }
}
