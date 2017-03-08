using System.Collections.ObjectModel;

namespace WordAddInDemoV2.DataContainers
{
    internal class BookmarksConatiner
    {
        private static BookmarksConatiner _instance;

        private BookmarksConatiner()
        {
            Bookmarks = new ObservableCollection<string>();
        }

        public ObservableCollection<string> Bookmarks { get; private set; }

        public static BookmarksConatiner Instance => _instance ?? (_instance = new BookmarksConatiner());

        public void AddBookmark()
        {
            
        }

        public void RemoveBookmark()
        {
            
        }

    }
}
