using System.Windows;
using WordAddInDemoV2.DataContainers;

namespace WordAddInDemoV2.Bookmark
{
    /// <summary>
    /// Interaction logic for BookmarksView.xaml
    /// </summary>
    public partial class BookmarksView
    {
        public BookmarksView()
        {
            InitializeComponent();
            DataContext = BookmarksConatiner.Instance.Bookmarks;
        }

        private void OnAddButtonClick(object sender, RoutedEventArgs e)
        {
            
        }

        public void OnNavigateButtonClick(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
