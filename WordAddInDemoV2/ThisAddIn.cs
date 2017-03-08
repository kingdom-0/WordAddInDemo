using Microsoft.Office.Core;
using WordAddInDemoV2.Bookmark;
using WordAddInDemoV2.Ribbons;
using WordAddInDemoV2.TaskPane;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddInDemoV2
{
    public partial class ThisAddIn
    {
        private RibbonTab _ribbonTab;
        internal Microsoft.Office.Tools.CustomTaskPane CurrentTaskPane { get; private set; }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonTab();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            AddTaskPane();
            HandleWordEvents(true);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            HandleWordEvents(false);
        }

        private void AddTaskPane()
        {
            if (CurrentTaskPane != null)
            {
                return;
            }

            CurrentTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(new TaskPaneView(), "文档元素结构");
            CurrentTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            CurrentTaskPane.Width = 300;
        }

        private void HandleWordEvents(bool register)
        {
            if (register)
            {
                Application.DocumentOpen += OnDocumentOpen;
                Application.DocumentChange += OnDocumentChange;
                Application.DocumentBeforeClose += OnDocumentBeforeClose;
            }
            else
            {
                Application.DocumentOpen -= OnDocumentOpen;
                Application.DocumentChange -= OnDocumentChange;
                Application.DocumentBeforeClose -= OnDocumentBeforeClose;
            }
        }

        private void OnDocumentBeforeClose(Word.Document doc, ref bool cancel)
        {
            int i = 0;
        }

        private void OnDocumentChange()
        {
            int i = 0;
        }

        private void OnDocumentOpen(Word.Document doc)
        {
            
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
