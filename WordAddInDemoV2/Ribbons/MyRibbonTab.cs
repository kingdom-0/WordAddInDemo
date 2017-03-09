using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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
        private static int _controlIndex;

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

        public void HandleSaveAsCommand(Office.IRibbonControl control, ref bool cancelDefault)
        {
            //TODO: Why this command is not executed???
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
                case ConstantNameData.TestButton:
                    AddControl(WinformControlType.Button, large, middle);
                    break;
                case ConstantNameData.TestCheckBox:
                    AddControl(WinformControlType.CheckBox, middle, small);
                    break;
                case ConstantNameData.TestDatePicker:
                    AddControl(WinformControlType.DateTimePicker, large, small);
                    break;
                case ConstantNameData.TestGroupBox:
                    AddControl(WinformControlType.GroupBox, large, large);
                    break;
                case ConstantNameData.TestLabel:
                    AddControl(WinformControlType.Label, middle, small);
                    break;
            }

            ApplicationHelper.MoveCursorToEnd();
        }

        public string GetImage(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case ConstantNameData.AddInController:
                    return "DarkShading";
                case ConstantNameData.TaskPaneController:
                    return "Bullets";
            }

            return string.Empty;
        }

        public void OnToggleButtonClick(Office.IRibbonControl control, bool isPressed)
        {
            switch (control.Id)
            {
                case ConstantNameData.AddInController:
                    SetAddInUsability(isPressed);
                    break;
                case ConstantNameData.TaskPaneController:
                    SetTaskPaneVisibility(isPressed);
                    break;
            }
        }

        public bool GetControlEnabled(Office.IRibbonControl control)
        {
            return _addInEnabled;
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
            var customTaskPane = Globals.ThisAddIn.CurrentTaskPane;
            customTaskPane.Visible = visible;
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

        private static void AddControl(WinformControlType controlType, float width, float height)
        {
            var range = ApplicationHelper.GetCurrentSelectionRange();
            var controlId = GuidGenerator.NewGuid();
            var document = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveDocument);
            document.Controls.AddControl(GetControl(controlType), range, width, 
                height, controlId);
            var controlItem = new ControlItem(controlId, controlType, range, width, height);
            ControlsContainer.Instance.ControlItems.Add(controlItem);
        }

        private static Control GetControl(WinformControlType controlType)
        {
            Control control;
            var index = _controlIndex++;
            switch (controlType)
            {
                case WinformControlType.Button:
                    control = new Button {Text = $@"Button{index}"};
                    break;
                case WinformControlType.CheckBox:
                    control = new CheckBox {Text = $@"CheckBox{index}"};
                    break;
                case WinformControlType.DateTimePicker:
                    control = new DateTimePicker();
                    break;
                case WinformControlType.GroupBox:
                    control = new GroupBox {Text = $@"GroupBox{index}"};
                    break;
                case WinformControlType.Label:
                    control = new Label {Text = $@"Label{index}"};
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(controlType), controlType, null);
            }

            return control;
        }

        #endregion
    }
}
