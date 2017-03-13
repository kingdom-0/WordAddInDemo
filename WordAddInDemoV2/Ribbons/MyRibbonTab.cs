using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using WordAddInDemoV2.ConstantDatas;
using WordAddInDemoV2.DataContainers;
using WordAddInDemoV2.Helpers;
using Office = Microsoft.Office.Core;

namespace WordAddInDemoV2.Ribbons
{
    [ComVisible(true)]
    public class MyRibbonTab : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        private bool _addInEnabled;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("WordAddInDemoV2.Ribbons.MyRibbonTab.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ui)
        {
            _ribbon = ui;
        }

        public void HandleInertBookmarkCommand(Office.IRibbonControl control, ref bool cancelDefault)
        {
            if (!_addInEnabled)
            {
                cancelDefault = false;
                return;
            }

            BookmarksConatiner.Instance.Reset();
            FormHelper.Instance.Show();
        }

        public void OnSelectedWinformControlChanged(Office.IRibbonControl control,
            string selectedItemId, int selectdedItemIndex)
        {
            const int large = ConstantControlSize.Large;
            const int middle = ConstantControlSize.Middle;
            const int small = ConstantControlSize.Small;

            switch (selectedItemId)
            {
                case ConstantControlNames.TestButton:
                    AddControl(WinformControlType.Button, large, middle);
                    break;
                case ConstantControlNames.TestCheckBox:
                    AddControl(WinformControlType.CheckBox, middle, small);
                    break;
                case ConstantControlNames.TestDatePicker:
                    AddControl(WinformControlType.DateTimePicker, large, small);
                    break;
                case ConstantControlNames.TestGroupBox:
                    AddControl(WinformControlType.GroupBox, large, large);
                    break;
                case ConstantControlNames.TestLabel:
                    AddControl(WinformControlType.Label, middle, small);
                    break;
            }

            ApplicationHelper.MoveCursorToEnd();
        }

        public string GetImage(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case ConstantControlNames.AddInController:
                    return ConstantControlNames.DarkShading;
                case ConstantControlNames.TaskPaneController:
                    return ConstantControlNames.Bullets;
            }

            return string.Empty;
        }

        public void OnToggleButtonClick(Office.IRibbonControl control, bool isPressed)
        {
            switch (control.Id)
            {
                case ConstantControlNames.AddInController:
                    SetAddInUsability(isPressed);
                    break;
                case ConstantControlNames.TaskPaneController:
                    SetTaskPaneVisibility(isPressed);
                    break;
            }
        }

        public void OnButtonClick(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case ConstantControlNames.SaveAsButton:
                    SaveDocument();
                    break;
            }
        }

        public bool GetControlEnabled(Office.IRibbonControl control)
        {
            return _addInEnabled;
        }

        private static void SaveDocument()
        {
            try
            {
                Globals.ThisAddIn.Application.ActiveDocument.SaveAs();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Save document failed, {ex.Message}.");
            }
        }

        private void SetAddInUsability(bool enable)
        {
            _addInEnabled = enable;
            if (!_addInEnabled)
            {
                SetTaskPaneVisibility(false);
            }

            _ribbon.Invalidate();
        }

        private static void SetTaskPaneVisibility(bool visible)
        {
            Globals.ThisAddIn.CurrentTaskPane.Visible = visible;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            for (var i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i],
                    StringComparison.OrdinalIgnoreCase) != 0)
                {
                    continue;
                }

                // ReSharper disable once AssignNullToNotNullAttribute
                var stream = asm.GetManifestResourceStream(resourceNames[i]);
                if (stream == null)
                {
                    continue;
                }
                using (var resourceReader = new StreamReader(stream))
                {
                    return resourceReader.ReadToEnd();
                }
            }
            return null;
        }

        private static void AddControl(WinformControlType controlType, int width, int height)
        {
            var range = ApplicationHelper.GetCurrentSelectionRange();
            var controlId = GuidGenerator.NewGuid();
            var document = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveDocument);
            var control = ControlGenerator.Instance.Generate(controlType, width, height);
            document.Controls.AddControl(control, range, control.Width, 
                control.Height, controlId);
        }

        #endregion
    }
}
