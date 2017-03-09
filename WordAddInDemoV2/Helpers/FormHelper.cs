using System.Windows.Forms;
using WordAddInDemoV2.Bookmark;

namespace WordAddInDemoV2.Helpers
{
    internal class FormHelper
    {
        private BookmarksForm _bookmarksForm;
        private static FormHelper _instance;

        private FormHelper()
        {
            
        }

        public static FormHelper Instance => _instance ?? (_instance = new FormHelper());

        public void Show()
        {
            _bookmarksForm = new BookmarksForm
            {
                StartPosition = FormStartPosition.CenterScreen
            };
            _bookmarksForm.Closed += OnBookmarksFormClosed;
            _bookmarksForm.ShowDialog();
        }

        public void Close()
        {
            _bookmarksForm?.Close();
        }

        private void OnBookmarksFormClosed(object sender, System.EventArgs e)
        {
            _bookmarksForm.Closed -= OnBookmarksFormClosed;
            _bookmarksForm = null;
        }
    }
}
