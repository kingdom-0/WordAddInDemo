using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace WordAddInDemoV2.DataContainers
{
    public enum BookmarkOrderType
    {
        Name,
        Position
    }

    internal class BookmarksConatiner
    {
        private static BookmarksConatiner _instance;

        private BookmarksConatiner()
        {
            Bookmarks = new ObservableCollection<BookmarkItem>();
        }

        public ObservableCollection<BookmarkItem> Bookmarks { get; }

        public static BookmarksConatiner Instance => _instance ?? (_instance = new BookmarksConatiner());

        public void Reset()
        {
            var document = Globals.ThisAddIn.Application.ActiveDocument;

            Bookmarks.Clear();

            var bookmarkList = new List<Microsoft.Office.Interop.Word.Bookmark>();

            foreach (Microsoft.Office.Interop.Word.Bookmark bookmark in document.Bookmarks)
            {
                bookmarkList.Add(bookmark);
            }

            foreach (var bookmark in bookmarkList.OrderBy(x => x.Start))
            {
                Bookmarks.Add(new BookmarkItem(bookmark.Name, bookmark.Start));
            }
        }

        public bool Contains(string bookmarkName)
        {
            foreach (var bookmarkItem in Bookmarks)
            {
                if (string.Equals(bookmarkName, bookmarkItem.Name))
                {
                    return true;
                }
            }

            return false;
        }

        public void OrderBookmark(BookmarkOrderType orderType)
        {
            switch (orderType)
            {
                case BookmarkOrderType.Name:
                    OrderByName();
                    break;
                case BookmarkOrderType.Position:
                    OrderByPosition();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(orderType), orderType, null);
            }
        }

        private void OrderByName()
        {
            var orderedBookmarks = Bookmarks.ToList().OrderBy(x => x.Name).ToList();
            Bookmarks.Clear();
            orderedBookmarks.ForEach(item => Bookmarks.Add(item));
        }

        private void OrderByPosition()
        {
            var orderedBookmarks = Bookmarks.ToList().OrderBy(x => x.Index).ToList();
            Bookmarks.Clear();
            orderedBookmarks.ForEach(item => Bookmarks.Add(item));
        }
    }

    public class BookmarkItem
    {
        public BookmarkItem(string name, int index)
        {
            Name = name;
            Index = index;
        }

        public int Index { get; }

        public string Name { get; }
    }
}
