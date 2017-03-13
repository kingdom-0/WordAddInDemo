using System;
using System.Windows.Forms;
using WordAddInDemoV2.Ribbons;

namespace WordAddInDemoV2.Helpers
{
    internal class ControlGenerator
    {
        private static ControlGenerator _instance;

        private ControlGenerator()
        {
            
        }

        public static ControlGenerator Instance => _instance ?? (_instance = new ControlGenerator());

        public Control Generate(WinformControlType controlType, int width, int height)
        {
            Control control;
            switch (controlType)
            {
                case WinformControlType.Button:
                    control = ButtonControl.Instance.GetInstance(width, height);
                    break;
                case WinformControlType.CheckBox:
                    control = CheckBoxControl.Instance.GetInstance(width, height);
                    break;
                case WinformControlType.DateTimePicker:
                    control = DateTimePickerControl.Instance.GetInstance(width, height);
                    break;
                case WinformControlType.GroupBox:
                    control = GroupBoxControl.Instance.GetInstance(width, height);
                    break;
                case WinformControlType.Label:
                    control = LabelControl.Instance.GetInstance(width, height);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(controlType), controlType, null);
            }

            return control;
        }
    }
}
