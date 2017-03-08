using System.Collections.ObjectModel;
using Microsoft.Office.Interop.Word;
using WordAddInDemoV2.Ribbons;

namespace WordAddInDemoV2.DataContainers
{
    internal class ControlsContainer
    {
        private static ControlsContainer _instance;

        private ControlsContainer()
        {
            ControlItems = new ObservableCollection<ControlItem>();
        }

        public static ControlsContainer Instance => _instance ?? (_instance = new ControlsContainer());

        public ObservableCollection<ControlItem> ControlItems { get; private set; }

        public void Add(ControlItem controlItem)
        {
            
        }
    }

    public class ControlItem
    {
        public ControlItem(string name, WinformControlType controlType, Range range,
            double width, double height)
        {
            Name = name;
            ControlType = controlType;
            Range = range;
            Width = width;
            Height = height;
        }

        public string Name { get; private set; }

        public WinformControlType ControlType { get; private set; }

        public Range Range { get; private set; }

        public double Width { get; private set; }

        public double Height { get; private set; }

    }
}
