using System;
using Microsoft.Office.Core;
using WordAddInDemoV2.Helpers;
using WordAddInDemoV2.Ribbons;
using WordAddInDemoV2.TaskPane;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddInDemoV2
{
    public partial class ThisAddIn
    {
        private bool _needDisplayRecentFile;

        internal Microsoft.Office.Tools.CustomTaskPane CurrentTaskPane { get; private set; }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new MyRibbonTab();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            AddTaskPane();
            HandleWordEvents(true);
            InitializeDisplayRectFilesProperty();
            //Application.CommandBars["File"].Controls.Add(typeof(CommandBarButton), "MyId",);
            //var controls = Application.CommandBars["File"].Controls;

            //foreach (CommandBarControl commandBarControl in controls)
            //{
            //    Console.WriteLine(commandBarControl.Caption);
            //}
            //var menuItem = controls.Add(MsoControlType.msoControlButton, Temporary: true);
            //menuItem.Caption = "自定义项";
            //menuItem.Width = 100;
            //menuItem.Height = 40;
            //menuItem.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            HandleWordEvents(false);
            SetDisplayRecentFilesProperty(_needDisplayRecentFile);
        }

        private void InitializeDisplayRectFilesProperty()
        {
            _needDisplayRecentFile = Globals.ThisAddIn.Application.DisplayRecentFiles;
            SetDisplayRecentFilesProperty(false);
        }

        private static void SetDisplayRecentFilesProperty(bool needDisplay)
        {
            Globals.ThisAddIn.Application.DisplayRecentFiles = needDisplay;
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
                Application.DocumentBeforeSave += OnDocumentBeforeSave;
            }
            else
            {
                Application.DocumentOpen -= OnDocumentOpen;
                Application.DocumentChange -= OnDocumentChange;
                Application.DocumentBeforeClose -= OnDocumentBeforeClose;
                Application.DocumentBeforeSave -= OnDocumentBeforeSave;
            }
        }

        private void OnDocumentBeforeSave(Word.Document doc, ref bool saveAsUi, ref bool cancel)
        {
            cancel = true;
            var dialog = doc.Application.FileDialog[MsoFileDialogType.msoFileDialogSaveAs];
            dialog.InitialFileName = GuidGenerator.NewGuid();
            var result = dialog.Show();
            if (result == -1)
            {
                //TODO.
                Globals.ThisAddIn.Application.Documents.Open("filepath");
            }
        }

        private void OnDocumentBeforeClose(Word.Document doc, ref bool cancel)
        {
            //TODO.
        }

        private void OnDocumentChange()
        {
            //TODO.
        }

        private void OnDocumentOpen(Word.Document doc)
        {
            //TODO.
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
