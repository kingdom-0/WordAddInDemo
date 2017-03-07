using Microsoft.Office.Core;
using WordAddInDemoV2.TaskPane;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddInDemoV2
{
    public partial class ThisAddIn
    {
        private Word.ApplicationEvents4_Event _wordEvent;
        private RibbonTab _ribbonTab;
        internal Microsoft.Office.Tools.CustomTaskPane CurrentTaskPane { get; private set; }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonTab();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            AddTaskPane();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            
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
