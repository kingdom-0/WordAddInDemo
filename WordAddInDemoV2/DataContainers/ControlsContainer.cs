using System.Collections.ObjectModel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordAddInDemoV2.Ribbons;

namespace WordAddInDemoV2.DataContainers
{
    internal class ControlsContainer
    {
        private static ControlsContainer _instance;

        private ControlsContainer()
        {
            ControlItems = new ObservableCollection<ControlItem>();
        }

        public static ControlsContainer Instance => _instance ?? (_instance = new ControlsContainer());

        public ObservableCollection<ControlItem> ControlItems { get; }

        public void Reset()
        {
            ControlItems.Clear();
            var document = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveDocument);
            
            foreach (Control control in document.Controls)
            {
                ControlItems.Add(ControlItem.GetNewInstance(ElementType.Control,
                    control.GetType().Name, control.Text));
            }

            foreach (Comment comment in document.Comments)
            {
                ControlItems.Add(ControlItem.GetNewInstance(ElementType.Comment, comment.GetType().Name,
                    comment.Author));
            }

            foreach (Table table in document.Tables)
            {
                ControlItems.Add(ControlItem.GetNewInstance(ElementType.Table,
                    table.GetType().Name, table.Title));
            }

            foreach (ContentControl contentControl in document.ContentControls)
            {
                ControlItems.Add(ControlItem.GetNewInstance(ElementType.ContentControl,
                    contentControl.GetType().Name, contentControl.Title));
            }

            foreach (var bookmark in document.Bookmarks)
            {
                ControlItems.Add(ControlItem.GetNewInstance(ElementType.Bookmark, 
                    bookmark.GetType().Name, bookmark.ToString()));
            }
        }
    }

    public class ControlItem
    {
        private ControlItem(ElementType elementType, string typeName, string description)
        {
            ElementType = elementType;
            TypeName = typeName;
            Description = description;
        }

        public static ControlItem GetNewInstance(ElementType elementType, string typeName, 
            string description)
        {
            return new ControlItem(elementType, typeName, description);
        }

        public string Description { get; private set; }

        public ElementType ElementType { get; private set; }

        public string TypeName { get; private set; }
    }
}
