using System;
using System.IO;
using System.Windows.Forms;
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
        private const int TaskPaneWidth = 400;
        private const int ConfirmedDialogCode = -1;
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
            CurrentTaskPane.Width = TaskPaneWidth;
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
                Application.DocumentChange -= OnDocumentChange;
            }
        }

        private void OnDocumentChange()
        {
            //AddCustomXmlPartToActiveDocument(Globals.ThisAddIn.Application.ActiveDocument);
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
            if (result == ConfirmedDialogCode)
            {
                var selectedPath = dialog.SelectedItems.Item(1);
                SaveDocumentWithCustomExtension(doc, selectedPath);
                SaveDocumentXmlContent(doc, selectedPath);
            }
            else
            {
                cancel = true;
            }
        }

        private static void SaveDocumentWithCustomExtension(Word.Document doc, string selectedPath)
        {
            doc.SaveAs($"{selectedPath}.{ConstantControlNames.CustomDocumentExtension}");
        }

        private static void SaveDocumentXmlContent(Word._Document doc, string selectedPath)
        {
            try
            {
                StreamWriter writer;
                using (writer = File.CreateText($@"{selectedPath}.xml"))
                {
                    writer.Write(doc.WordOpenXML);
                }
            }
            catch (Exception ex)
            {
                // ReSharper disable once LocalizableElement
                MessageBox.Show($"Save document to xml failed, {ex.Message}.");
            }
        }

        //private void AddCustomXmlPartToActiveDocument(Word.Document document)
        //{
        //    //TODO: add logic.
        //    //document.CustomXMLParts.Add(xmlString, missing);
        //}

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
