using System;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddInDemo
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            BtnAddText.Click += OnAddTextButtonClick;
            BtnAddTable.Click += OnAddTableButtonClick;
            BtnAddImage.Click += OnAddImageButtonClick;
            BtnAddInputText.Click += OnButtonClick;
            BtnAddDatePicker.Click += OnButtonClick;
            BtnAddRichText.Click += OnButtonClick;
            BtnAddDropDownList.Click += OnButtonClick;
            ToggleBtnSwitchFont.Click += OnToggleButtonClick;
            ToggleButtonSelectAll.Click += OnToggleButtonClick;
        }

        private void OnButtonClick(object sender, RibbonControlEventArgs e)
        {
            var button = sender as RibbonButton;
            if (button == null)
            {
                return;
            }

            var range = GetCurrentSelectionRange();
            var document = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveDocument);
            switch (button.Name)
            {
                case "BtnAddInputText":
                    range.Text = TextInputField.Text;
                    break;
                case "BtnAddDatePicker":
                    document.Controls.AddDatePickerContentControl(range, $"MyDatePicker{Guid.NewGuid()}");
                    break;
                case "BtnAddRichText":
                    document.Controls.AddRichTextContentControl(range, $"MyRichText{Guid.NewGuid()}");
                    break;
                case "BtnAddDropDownList":
                    document.Controls.AddDropDownListContentControl(range, $"MyDropDownList{Guid.NewGuid()}");
                    break;
            }

            MoveCursorToEnd();
        }

        private void OnToggleButtonClick(object sender, RibbonControlEventArgs e)
        {
            var toggleButton = sender as RibbonToggleButton;
            if (toggleButton == null)
            {
                return;
            }

            switch (toggleButton.Name)
            {
                case "ToggleBtnSwitchFont":
                    UpdateRangeTextFont(toggleButton.Checked);
                    break;
                case "ToggleButtonSelectAll":
                    SelectOrUnselectDocumentContent(toggleButton.Checked);
                    break;
            }

            
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

        private void SelectOrUnselectDocumentContent(bool select)
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

        private void OnAddTextButtonClick(object sender, RibbonControlEventArgs e)
        {
            var range = GetCurrentSelectionRange();
            range.Text = "Text from code behind.";
            MoveCursorToEnd();
        }

        private void OnAddImageButtonClick(object sender, RibbonControlEventArgs e)
        {
            const string imageUrl = @"D:\Menu.png";
            Globals.ThisAddIn.Application.ActiveDocument.InlineShapes.AddPicture(imageUrl);
            MoveCursorToEnd();
        }

        private void OnAddTableButtonClick(object sender, RibbonControlEventArgs e)
        {
            var range = GetCurrentSelectionRange();
            Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(range, 3, 4);
            MoveCursorToEnd();
        }

        private Range GetCurrentSelectionRange()
        {
            object start = Globals.ThisAddIn.Application.Selection.Start;
            object end = Globals.ThisAddIn.Application.Selection.End;
            return Globals.ThisAddIn.Application.ActiveDocument.Range(ref start, ref end);
        }

        private static void MoveCursorToEnd()
        {
            object story = WdUnits.wdStory;
            Globals.ThisAddIn.Application.Selection.EndKey(ref story);
        }
    }
}
