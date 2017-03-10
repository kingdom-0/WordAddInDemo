using System.Windows.Forms;

namespace WordAddInDemoV2.Helpers
{
    internal abstract class ControlGenerator
    {
        protected static int Index { get; }

        static ControlGenerator()
        {
            Index = 0;
        }

        public abstract Control Generate(int width, int height);
    }

    internal class ButtonGenerator : ControlGenerator
    {
        private const string ButtonPrefix = "Btn";
        private static ButtonGenerator _instance;

        private ButtonGenerator() { }

        public static ButtonGenerator Instance => _instance ?? (_instance = new ButtonGenerator());

        public override Control Generate(int width, int height)
        {
            return new Button {Text = $@"{ButtonPrefix}-{Index}", Width = width, Height = height};
        }
    }

    internal class CheckBoxGenerator : ControlGenerator
    {
        private const string CheckBoxPrefix = "CheckBox";
        private static CheckBoxGenerator _instance;

        private CheckBoxGenerator() { }

        public static CheckBoxGenerator Instance => _instance ?? (_instance = new CheckBoxGenerator());

        public override Control Generate(int width, int height)
        {
            return new CheckBox {Text = $@"{CheckBoxPrefix}-{Index}", Width = width, Height = height};
        }
    }

    internal class DateTimePickerGenerator : ControlGenerator
    {
        private static DateTimePickerGenerator _instance;

        private DateTimePickerGenerator() { }

        public static DateTimePickerGenerator Instance => _instance ?? 
            (_instance = new DateTimePickerGenerator());

        public override Control Generate(int width, int height)
        {
            return new DateTimePicker() {Width = width, Height = height};
        }
    }

    internal class GroupBoxGenerator : ControlGenerator
    {
        private const string GroupBoxPrefix = "GroupBox";
        private static GroupBoxGenerator _instance;

        private GroupBoxGenerator() { }

        public static GroupBoxGenerator Instance => _instance ?? (_instance = new GroupBoxGenerator());

        public override Control Generate(int width, int height)
        {
            return new GroupBox {Text = $@"{GroupBoxPrefix}-{Index}", Width = width, Height = height};
        }
    }

    internal class LabelGenerator : ControlGenerator
    {
        private const string LabelPrefix = "Label";
        private static LabelGenerator _instance;

        private LabelGenerator() { }

        public static LabelGenerator Instance => _instance ?? (_instance = new LabelGenerator());

        public override Control Generate(int width, int height)
        {
            return new Label {Text = $@"{LabelPrefix}-{Index}", Width = width, Height = height};
        }
    }
}
