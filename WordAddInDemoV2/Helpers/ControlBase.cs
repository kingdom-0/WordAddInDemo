using System.Windows.Forms;

namespace WordAddInDemoV2.Helpers
{
    internal abstract class ControlBase
    {
        protected static int Index { get; set; }

        static ControlBase()
        {
            Index = 0;
        }

        public abstract Control GetInstance(int width, int height);
    }

    internal class ButtonControl : ControlBase
    {
        private const string ButtonPrefix = "Btn";
        private static ButtonControl _instance;

        private ButtonControl() { }

        public static ButtonControl Instance => _instance ?? (_instance = new ButtonControl());

        public override Control GetInstance(int width, int height)
        {
            return new Button {Text = $@"{ButtonPrefix}-{Index++}", Width = width, Height = height};
        }
    }

    internal class CheckBoxControl : ControlBase
    {
        private const string CheckBoxPrefix = "CheckBox";
        private static CheckBoxControl _instance;

        private CheckBoxControl() { }

        public static CheckBoxControl Instance => _instance ?? (_instance = new CheckBoxControl());

        public override Control GetInstance(int width, int height)
        {
            return new CheckBox {Text = $@"{CheckBoxPrefix}-{Index++}", Width = width, Height = height};
        }
    }

    internal class DateTimePickerControl : ControlBase
    {
        private static DateTimePickerControl _instance;

        private DateTimePickerControl() { }

        public static DateTimePickerControl Instance => _instance ?? 
            (_instance = new DateTimePickerControl());

        public override Control GetInstance(int width, int height)
        {
            Index++;
            return new DateTimePicker() {Width = width, Height = height};
        }
    }

    internal class GroupBoxControl : ControlBase
    {
        private const string GroupBoxPrefix = "GroupBox";
        private static GroupBoxControl _instance;

        private GroupBoxControl() { }

        public static GroupBoxControl Instance => _instance ?? (_instance = new GroupBoxControl());

        public override Control GetInstance(int width, int height)
        {
            return new GroupBox {Text = $@"{GroupBoxPrefix}-{Index++}", Width = width, Height = height};
        }
    }

    internal class LabelControl : ControlBase
    {
        private const string LabelPrefix = "Label";
        private static LabelControl _instance;

        private LabelControl() { }

        public static LabelControl Instance => _instance ?? (_instance = new LabelControl());

        public override Control GetInstance(int width, int height)
        {
            return new Label {Text = $@"{LabelPrefix}-{Index++}", Width = width, Height = height};
        }
    }
}
