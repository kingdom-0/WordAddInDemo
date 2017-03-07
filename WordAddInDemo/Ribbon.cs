using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordAddInDemo.Properties;
using Office = Microsoft.Office.Core;

namespace WordAddInDemo
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private const string TestImageUrl = @"D:\Menu.png";

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("WordAddInDemo.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ui)
        {
        }

        public void OnToggleButtonActionCallback(Office.IRibbonControl control, bool isPressed)
        {
            if (control == null)
            {
                return;
            }

            switch (control.Id)
            {
                case ConstantControlName.ToggleBtnSwitchFont:
                    UpdateRangeTextFont(isPressed);
                    break;
                case ConstantControlName.ToggleButtonSelectAll:
                    SelectOrUnselectDocumentContent(isPressed);
                    break;
            }

            PrintTraceMessage(control.Id, "OnToggleButtonActionCallback");
        }

        public void OnButtonActionCallback(Office.IRibbonControl control)
        {
            if (control == null)
            {
                return;
            }

            var range = GetCurrentSelectionRange();

            var document = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveDocument);

            switch (control.Id)
            {
                case ConstantControlName.BtnAddText:
                    range.Text = "Text from code behind.";
                    break;
                case ConstantControlName.BtnAddTable:
                    Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(range, 3, 4);
                    break;
                case ConstantControlName.BtnAddImage:
                    Resources.Menu.Save(TestImageUrl);
                    Globals.ThisAddIn.Application.ActiveDocument.InlineShapes.AddPicture(TestImageUrl);
                    break;
                case ConstantControlName.BtnDisplayAreaElements:
                    DisplayAreaElements();
                    break;
                case ConstantControlName.BtnAddDatePicker:
                    document.Controls.AddDatePickerContentControl(range, $"MyDatePicker{Guid.NewGuid()}");
                    break;
                case ConstantControlName.BtnAddRichText:
                    document.Controls.AddRichTextContentControl(range, $"MyRichText{Guid.NewGuid()}");
                    break;
                case ConstantControlName.BtnAddDropDownList:
                    document.Controls.AddDropDownListContentControl(range, $"MyDropDownList{Guid.NewGuid()}");
                    break;
            }

            PrintTraceMessage(control.Id, "OnButtonActionCallback");

            MoveCursorToEnd();
        }

        public string GetButtonLabel(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetButtonLabel");
            return "添加文本";
        }

        public string GetMenuLabel(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetMenuLabel");
            return "添加WinForm控件";
        }

        public string GetGroupLabel(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetGroupLabel");
            return "添加基本数据";
        }

        public string GetEditBoxLabel(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetInputField");
            return "文本:";
        }

        public string GetToggleButtonLabel(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetToggleButtonLabel");
            return "选择全部";
        }

        public string GetCheckBoxLabel(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetCheckBoxLabel");
            return "测试CheckBox";
        }

        public string GetLabelControlText(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetLabelControlText");
            return "文本信息";
        }

        public string GetComboBoxItemImage(Office.IRibbonControl control, int itemIndex)
        {
            return "Bullets";
        }

        public string GetComboBoxLabel(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetComboBoxLabel");
            return "测试ComboBox";
        }

        public void OnComboBoxChange(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "OnComboBoxChange");
        }

        public string GetDropdownItemImage(Office.IRibbonControl control, int itemIndex)
        {
            PrintTraceMessage(control.Id, "GetComboBoxLabel");
            return "Numbering";
        }

        public string GetDropDownListLabel(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetDropDownListLabel");
            return "测试DropDown";
        }

        public string GetGalleryItemImage(Office.IRibbonControl control, int itemIndex)
        {
            PrintTraceMessage(control.Id, "GetGalleryItemImage");
            return "PenComment";
        }

        public string GetGalleryLabel(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetGalleryLabel");
            return "测试Gallery";
        }

        public string GetGalleryItemWidth(Office.IRibbonControl control, int itemIndex)
        {
            return "30";
        }

        public string GetGalleryItemHeight(Office.IRibbonControl control, int itemIndex)
        {
            return "8";
        }

        public void OnCheckboxActionCallback(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "OnCheckboxActionCallback");
        }

        public string GetControlDescription(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlDescription");
            return string.Empty;
        }

        public bool GetControlEnabled(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlEnabled");
            return true;
        }

        public Bitmap GetImage(Office.IRibbonControl control)
        {
            var result = default(Bitmap);

            switch (control.Id)
            {
                case ConstantControlName.MenuAddBasicInfo:
                case ConstantControlName.MenuAddWinFormControl:
                    result = Resources.Menu;
                    break;
                case ConstantControlName.BtnAddText:
                    result = Resources.Text;
                    break;
                case ConstantControlName.BtnAddTable:
                    result = Resources.Table;
                    break;
                case ConstantControlName.BtnAddImage:
                    result = Resources.Image;
                    break;
                case ConstantControlName.BtnAddDatePicker:
                case ConstantControlName.BtnAddDropDownList:
                case ConstantControlName.BtnAddRichText:
                case ConstantControlName.BtnAddInputText:
                    result = Resources.Add;
                    break;
                case ConstantControlName.MyComboBox:
                    result = Resources.List0;
                    break;
                case ConstantControlName.MyDropDown:
                    result = Resources.List1;
                    break;
                case ConstantControlName.BtnInMySplitButton:
                    result = Resources.List2;
                    break;
                case ConstantControlName.MyGallery:
                    result = Resources.List3;
                    break;
                case ConstantControlName.GroupBtn1:
                case ConstantControlName.GroupBtn2:
                case ConstantControlName.GroupBtn3:
                    result = Resources.MenuItem;
                    break;
                case ConstantControlName.ToggleButtonSelectAll:
                    result = Resources.Select;
                    break;
                case ConstantControlName.ToggleBtnSwitchFont:
                    result = Resources.Pen;
                    break;
                case ConstantControlName.BtnDisplayAreaElements:
                    result = Resources.Map;
                    break;
            }

            return result;
        }

        public string GetControlKeytip(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlKeytip");
            return string.Empty;
        }

        public string GetControlScreentip(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlScreentip");
            return $"{control.Id} screen tip.";
        }

        public bool GetControlShowImage(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlShowImage");
            return true;
        }

        public bool GetControlShowLabel(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlShowLabel");
            return true;
        }

        public string GetControlSupertip(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlSupertip");
            return $"{control.Id} super tip.";
        }

        public bool GetControlVisible(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlVisible");
            return true;
        }

        public int GetControlItemCount(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlItemCount");
            return 5;
        }

        public string GetControlItemId(Office.IRibbonControl control, int itemIndex)
        {
            PrintTraceMessage(control.Id, "GetControlItemId");
            return string.Concat("Item", itemIndex);
        }

        public string GetControlItemLabel(Office.IRibbonControl control, int itemIndex)
        {
            PrintTraceMessage(control.Id, "GetControlItemLabel");
            return string.Concat("Item", itemIndex);
        }

        public string GetControlItemScreentip(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlItemScreentip");
            return "Item screen tip from code behind.";
        }

        public string GetControlItemSupertip(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlItemSupertip");
            return "Tip from code behind.";
        }

        public string GetControlText(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlText");
            return string.Empty;
        }

        public Office.RibbonControlSize GetControlSize(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlSize");
            return default(Office.RibbonControlSize);
        }

        public bool GetControlPressed(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlPressed");

            return false;
        }

        public string GetControlSelectedItemId(Office.IRibbonControl control)
        {
            PrintTraceMessage(control.Id, "GetControlSelectedItemId");
            return string.Empty;
        }

        public void OnEditBoxTextChange(Office.IRibbonControl control, string text)
        {
            PrintTraceMessage(control.Id, "OnEditBoxTextChange");

            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var range = GetCurrentSelectionRange();
            range.Text = text;

            MoveCursorToEnd();
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();

            var resourceNames = asm.GetManifestResourceNames();
            for (var i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) != 0)
                {
                    continue;
                }

                // ReSharper disable once AssignNullToNotNullAttribute
                using (var resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                {
                    return resourceReader.ReadToEnd();
                }
            }

            return null;
        }

        #endregion

        #region Inner methods

        private static void DisplayAreaElements()
        {
            var selection = Globals.ThisAddIn.Application.Selection;
            var elementBuilder = new StringBuilder();

            elementBuilder.AppendLine($"Character count:{selection.Characters.Count}");

            elementBuilder.AppendLine($"Image count:{selection.InlineShapes.Count}");

            elementBuilder.AppendLine($"Shape count:{selection.ShapeRange.Count}");
            var shapeIndex = 0;
            foreach (Shape shape in selection.ShapeRange)
            {
                elementBuilder.AppendLine($"\t Shape{shapeIndex} type is {shape.GetType()}");
            }

            elementBuilder.AppendLine($"Table count:{selection.Tables.Count}");

            var tableIndex = 0;
            foreach (Table selectionTable in selection.Tables)
            {
                var rows = selectionTable.Rows;
                var columns = selectionTable.Columns;
                elementBuilder.AppendLine($"\t Table{tableIndex}:Rows {rows.Count}, Columns {columns.Count}");
                tableIndex++;
            }

            elementBuilder.AppendLine($"ContentControl count:{selection.ContentControls.Count}");

            elementBuilder.AppendLine($"Bookmarks count:{selection.Bookmarks.Count}");

            MessageBox.Show(elementBuilder.ToString());
        }

        private void UpdateRangeTextFont(bool reset)
        {
            var range = GetCurrentSelectionRange();
            if (reset)
            {
                range.Font.Size = 18;
                range.Font.Bold = 1;
                range.Font.Animation = WdAnimation.wdAnimationSparkleText;
                range.Font.Color = WdColor.wdColorBlue;
                range.Font.Italic = 1;
            }
            else
            {
                range.Font.Size = 10.5f;
                range.Font.Bold = 0;
                range.Font.Animation = WdAnimation.wdAnimationNone;
                range.Font.Color = WdColor.wdColorBlack;
                range.Font.Italic = 0;
            }
        }

        private static void SelectOrUnselectDocumentContent(bool select)
        {
            if (select)
            {
                Globals.ThisAddIn.Application.ActiveDocument.Content.Select();
            }
            else
            {
                MoveCursorToEnd();
            }
        }

        private static void MoveCursorToEnd()
        {
            object story = WdUnits.wdStory;
            Globals.ThisAddIn.Application.Selection.EndKey(ref story);
        }

        private Range GetCurrentSelectionRange()
        {
            object start = Globals.ThisAddIn.Application.Selection.Start;
            object end = Globals.ThisAddIn.Application.Selection.End;
            return Globals.ThisAddIn.Application.ActiveDocument.Range(ref start, ref end);
        }

        private static void PrintTraceMessage(string controlId, string methodName)
        {
            Trace.WriteLine($"Method {methodName} belongs to {controlId} is called.");
        }

        #endregion
    }
}
