using System;
using Microsoft.Office.Core;
using WordAddInDemoV2.ConstantDatas;
using WordAddInDemoV2.DataContainers;
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

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            AddTaskPane();
            HandleApplicationEvents(true);
            InitializeDisplayRectFilesProperty();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            HandleApplicationEvents(false);
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

            CurrentTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(new TaskPaneView(), 
                ConstantControlNames.TaskPaneTitle);
            CurrentTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            CurrentTaskPane.Width = 300;
        }

        private void HandleApplicationEvents(bool register)
        {
            if (register)
            {
                Application.WindowSelectionChange += OnWindowSelectionChange;
                Application.DocumentBeforeSave += OnDocumentBeforeSave;
                Application.DocumentChange += OnDocumentChange;
            }
            else
            {
                Application.DocumentBeforeSave -= OnDocumentBeforeSave;
                Application.WindowSelectionChange -= OnWindowSelectionChange;
            }
        }

        private void OnDocumentChange()
        {
            AddCustomXmlPartToActiveDocument(Globals.ThisAddIn.Application.ActiveDocument);
        }

        private void OnWindowSelectionChange(Word.Selection sel)
        {
            
            ControlsContainer.Instance.Reset();
        }

        // ReSharper disable once RedundantAssignment
        private static void OnDocumentBeforeSave(Word.Document doc, ref bool saveAsUi, ref bool cancel)
        {
            saveAsUi = false;
            var dialog = doc.Application.FileDialog[MsoFileDialogType.msoFileDialogSaveAs];
            dialog.InitialFileName = GuidGenerator.NewGuid();
            dialog.Title = ConstantControlNames.DialogTitle;
            var result = dialog.Show();
            if (result == -1)
            {
                var selectedPath = default(string);
                foreach (var fileDialogSelectedItem in dialog.SelectedItems)
                {
                    selectedPath = fileDialogSelectedItem.ToString();
                }
                Globals.ThisAddIn.Application.ActiveDocument.SaveAs(
                    $"{selectedPath}.{ConstantControlNames.CustomDocumentExtension}");
            }
            else
            {
                cancel = true;
            }
        }

        private void AddCustomXmlPartToActiveDocument(Word.Document document)
        {
            string xmlString =
                "<?xml version=\"1.0\" encoding=\"utf-8\" ?>" +
                "<employees xmlns=\"http://schemas.microsoft.com/vsto/samples\">" +
                    "<employee>" +
                        "<name>Karina Leal</name>" +
                        "<hireDate>1999-04-01</hireDate>" +
                        "<title>Manager</title>" +
                    "</employee>" +
                "</employees>";

            document.CustomXMLParts.Add(xmlString, missing);
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
