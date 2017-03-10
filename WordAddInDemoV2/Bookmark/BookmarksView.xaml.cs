using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Interop.Word;
using WordAddInDemoV2.ConstantDatas;
using WordAddInDemoV2.DataContainers;
using WordAddInDemoV2.Helpers;

namespace WordAddInDemoV2.Bookmark
{
    /// <summary>
    /// Interaction logic for BookmarksView.xaml
    /// </summary>
    public partial class BookmarksView
    {
        private readonly Document _activeDocument;

        public BookmarksView()
        {
            InitializeComponent();
            DataContext = BookmarksConatiner.Instance.Bookmarks;
            _activeDocument = Globals.ThisAddIn.Application.ActiveDocument;
        }

        private void OnAddButtonClick(object sender, RoutedEventArgs e)
        {
            var bookmarkName = TxtBlkBookmark.Text;
            DeleteBookmarkIfExist(bookmarkName);
            _activeDocument.Bookmarks.Add(bookmarkName, ApplicationHelper.GetCurrentSelectionRange());

            FormHelper.Instance.Close();
        }

        public void OnNavigateButtonClick(object sender, RoutedEventArgs e)
        {
            var bookmarkName = TxtBlkBookmark.Text;
            if (_activeDocument.Bookmarks.Exists(bookmarkName))
            {
                _activeDocument.Bookmarks[bookmarkName].Select();
            }
        }

        private void DeleteBookmarkIfExist(string bookmarkName)
        {
            if (_activeDocument.Bookmarks.Exists(bookmarkName))
            {
                _activeDocument.Bookmarks[bookmarkName].Delete();
            }
        }

        private void OnOrderTypeChecked(object sender, RoutedEventArgs e)
        {
            var radioButton = sender as RadioButton;
            if (radioButton == null)
            {
                return;
            }

            switch (radioButton.Name)
            {
                case ConstantControlNames.RBtnSortByName:
                    BookmarksConatiner.Instance.OrderBookmark(BookmarkOrderType.Name);
                    break;
                case ConstantControlNames.RBtnSortByLocation:
                    BookmarksConatiner.Instance.OrderBookmark(BookmarkOrderType.Position);
                    break;
            }

            DataContext = BookmarksConatiner.Instance.Bookmarks;
        }

        private void OnBookmarkSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var bookmarkItem = BookmarkList.SelectedItem as BookmarkItem;
            if (bookmarkItem != null)
            {
                TxtBlkBookmark.Text = bookmarkItem.Name;
            }
        }

        private void OnBookmarkTextChanged(object sender, TextChangedEventArgs e)
        {
            var bookmarkName = TxtBlkBookmark.Text;
            if (string.IsNullOrEmpty(bookmarkName))
            {
                AddBookmark.IsEnabled = false;
                NavigateToBookmark.IsEnabled = false;
                return;
            }

            AddBookmark.IsEnabled = true;

            if (!BookmarksConatiner.Instance.Contains(bookmarkName))
            {
                NavigateToBookmark.IsEnabled = false;
                return;
            }
            NavigateToBookmark.IsEnabled = true;
            BookmarkList.SelectedItem = bookmarkName;
        }

        private void OnCancalButtonClick(object sender, RoutedEventArgs e)
        {
            FormHelper.Instance.Close();
        }

        private void OnHideBookmarkControlClick(object sender, RoutedEventArgs e)
        {
            var checkBox = sender as System.Windows.Controls.CheckBox;
            if(checkBox?.IsChecked != null)
            { 
                Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.ShowHidden = (bool)checkBox.IsChecked;
            }
        }
    }
}
