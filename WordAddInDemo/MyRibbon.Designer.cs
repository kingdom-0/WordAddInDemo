namespace WordAddInDemo
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Group1 = this.Factory.CreateRibbonGroup();
            this.MenuAddBasicInfo = this.Factory.CreateRibbonMenu();
            this.BtnAddText = this.Factory.CreateRibbonButton();
            this.BtnAddTable = this.Factory.CreateRibbonButton();
            this.BtnAddImage = this.Factory.CreateRibbonButton();
            this.MenuAddWinFormControl = this.Factory.CreateRibbonMenu();
            this.BtnAddDatePicker = this.Factory.CreateRibbonButton();
            this.BtnAddRichText = this.Factory.CreateRibbonButton();
            this.BtnAddDropDownList = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.box1 = this.Factory.CreateRibbonBox();
            this.TextInputField = this.Factory.CreateRibbonEditBox();
            this.BtnAddInputText = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.MyComboBox = this.Factory.CreateRibbonComboBox();
            this.MyDropDown = this.Factory.CreateRibbonDropDown();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.MySplitButton = this.Factory.CreateRibbonSplitButton();
            this.MyGallery = this.Factory.CreateRibbonGallery();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.GroupBtn1 = this.Factory.CreateRibbonButton();
            this.GroupBtn2 = this.Factory.CreateRibbonButton();
            this.GroupBtn3 = this.Factory.CreateRibbonButton();
            this.box2 = this.Factory.CreateRibbonBox();
            this.ToggleButtonSelectAll = this.Factory.CreateRibbonToggleButton();
            this.ToggleBtnSwitchFont = this.Factory.CreateRibbonToggleButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.MyCheckBox = this.Factory.CreateRibbonCheckBox();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.tab1.SuspendLayout();
            this.Group1.SuspendLayout();
            this.box1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group5.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.box2.SuspendLayout();
            this.group6.SuspendLayout();
            this.tab2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Label = "测试Tab";
            this.tab1.Name = "tab1";
            // 
            // Group1
            // 
            this.Group1.Items.Add(this.MenuAddBasicInfo);
            this.Group1.Items.Add(this.MenuAddWinFormControl);
            this.Group1.Items.Add(this.separator2);
            this.Group1.Items.Add(this.box1);
            this.Group1.Label = "添加基本数据";
            this.Group1.Name = "Group1";
            // 
            // MenuAddBasicInfo
            // 
            this.MenuAddBasicInfo.Image = global::WordAddInDemo.Properties.Resources.Menu;
            this.MenuAddBasicInfo.Items.Add(this.BtnAddText);
            this.MenuAddBasicInfo.Items.Add(this.BtnAddTable);
            this.MenuAddBasicInfo.Items.Add(this.BtnAddImage);
            this.MenuAddBasicInfo.Label = "添加基本数据";
            this.MenuAddBasicInfo.Name = "MenuAddBasicInfo";
            this.MenuAddBasicInfo.ScreenTip = "添加基本数据";
            this.MenuAddBasicInfo.ShowImage = true;
            // 
            // BtnAddText
            // 
            this.BtnAddText.Image = global::WordAddInDemo.Properties.Resources.Text;
            this.BtnAddText.Label = "添加文本";
            this.BtnAddText.Name = "BtnAddText";
            this.BtnAddText.ShowImage = true;
            // 
            // BtnAddTable
            // 
            this.BtnAddTable.Image = global::WordAddInDemo.Properties.Resources.Table;
            this.BtnAddTable.Label = "添加表格";
            this.BtnAddTable.Name = "BtnAddTable";
            this.BtnAddTable.ScreenTip = "在当前鼠标位置添加表格";
            this.BtnAddTable.ShowImage = true;
            // 
            // BtnAddImage
            // 
            this.BtnAddImage.Image = global::WordAddInDemo.Properties.Resources.Image;
            this.BtnAddImage.Label = "添加图片";
            this.BtnAddImage.Name = "BtnAddImage";
            this.BtnAddImage.ShowImage = true;
            // 
            // MenuAddWinFormControl
            // 
            this.MenuAddWinFormControl.Image = global::WordAddInDemo.Properties.Resources.Menu;
            this.MenuAddWinFormControl.Items.Add(this.BtnAddDatePicker);
            this.MenuAddWinFormControl.Items.Add(this.BtnAddRichText);
            this.MenuAddWinFormControl.Items.Add(this.BtnAddDropDownList);
            this.MenuAddWinFormControl.Label = "添加WinForm控件";
            this.MenuAddWinFormControl.Name = "MenuAddWinFormControl";
            this.MenuAddWinFormControl.ScreenTip = "添加WinForm控件";
            this.MenuAddWinFormControl.ShowImage = true;
            // 
            // BtnAddDatePicker
            // 
            this.BtnAddDatePicker.Image = global::WordAddInDemo.Properties.Resources.Add;
            this.BtnAddDatePicker.Label = "添加DatePicker";
            this.BtnAddDatePicker.Name = "BtnAddDatePicker";
            this.BtnAddDatePicker.ShowImage = true;
            // 
            // BtnAddRichText
            // 
            this.BtnAddRichText.Image = global::WordAddInDemo.Properties.Resources.Add;
            this.BtnAddRichText.Label = "添加RichText";
            this.BtnAddRichText.Name = "BtnAddRichText";
            this.BtnAddRichText.ShowImage = true;
            // 
            // BtnAddDropDownList
            // 
            this.BtnAddDropDownList.Image = global::WordAddInDemo.Properties.Resources.Add;
            this.BtnAddDropDownList.Label = "添加DropDownList";
            this.BtnAddDropDownList.Name = "BtnAddDropDownList";
            this.BtnAddDropDownList.ShowImage = true;
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // box1
            // 
            this.box1.Items.Add(this.TextInputField);
            this.box1.Items.Add(this.BtnAddInputText);
            this.box1.Name = "box1";
            // 
            // TextInputField
            // 
            this.TextInputField.Label = "文本:";
            this.TextInputField.Name = "TextInputField";
            this.TextInputField.ScreenTip = "输入文本";
            this.TextInputField.Text = null;
            // 
            // BtnAddInputText
            // 
            this.BtnAddInputText.Image = global::WordAddInDemo.Properties.Resources.Add;
            this.BtnAddInputText.Label = "添加";
            this.BtnAddInputText.Name = "BtnAddInputText";
            this.BtnAddInputText.ScreenTip = "点击添加";
            this.BtnAddInputText.ShowImage = true;
            this.BtnAddInputText.ShowLabel = false;
            // 
            // group2
            // 
            this.group2.Items.Add(this.MyComboBox);
            this.group2.Items.Add(this.MyDropDown);
            this.group2.Items.Add(this.separator1);
            this.group2.Items.Add(this.MySplitButton);
            this.group2.Items.Add(this.MyGallery);
            this.group2.Label = "下拉列表";
            this.group2.Name = "group2";
            // 
            // MyComboBox
            // 
            this.MyComboBox.Image = global::WordAddInDemo.Properties.Resources.List0;
            ribbonDropDownItemImpl1.Image = global::WordAddInDemo.Properties.Resources.Map;
            ribbonDropDownItemImpl1.Label = "Item0";
            ribbonDropDownItemImpl2.Image = global::WordAddInDemo.Properties.Resources.Map;
            ribbonDropDownItemImpl2.Label = "Item1";
            ribbonDropDownItemImpl3.Image = global::WordAddInDemo.Properties.Resources.Map;
            ribbonDropDownItemImpl3.Label = "Item2";
            ribbonDropDownItemImpl4.Image = global::WordAddInDemo.Properties.Resources.Map;
            ribbonDropDownItemImpl4.Label = "Item3";
            ribbonDropDownItemImpl5.Image = global::WordAddInDemo.Properties.Resources.Map;
            ribbonDropDownItemImpl5.Label = "Item4";
            this.MyComboBox.Items.Add(ribbonDropDownItemImpl1);
            this.MyComboBox.Items.Add(ribbonDropDownItemImpl2);
            this.MyComboBox.Items.Add(ribbonDropDownItemImpl3);
            this.MyComboBox.Items.Add(ribbonDropDownItemImpl4);
            this.MyComboBox.Items.Add(ribbonDropDownItemImpl5);
            this.MyComboBox.Label = "测试ComboBox";
            this.MyComboBox.Name = "MyComboBox";
            this.MyComboBox.ShowImage = true;
            this.MyComboBox.Text = null;
            // 
            // MyDropDown
            // 
            this.MyDropDown.Buttons.Add(this.button4);
            this.MyDropDown.Buttons.Add(this.button5);
            this.MyDropDown.Buttons.Add(this.button6);
            this.MyDropDown.Buttons.Add(this.button7);
            this.MyDropDown.Image = global::WordAddInDemo.Properties.Resources.List1;
            this.MyDropDown.Label = "测试DropDown";
            this.MyDropDown.Name = "MyDropDown";
            this.MyDropDown.ShowImage = true;
            // 
            // button4
            // 
            this.button4.Image = global::WordAddInDemo.Properties.Resources.Map;
            this.button4.Label = "button4";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            // 
            // button5
            // 
            this.button5.Image = global::WordAddInDemo.Properties.Resources.Map;
            this.button5.Label = "button5";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            // 
            // button6
            // 
            this.button6.Image = global::WordAddInDemo.Properties.Resources.Map;
            this.button6.Label = "button6";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            // 
            // button7
            // 
            this.button7.Image = global::WordAddInDemo.Properties.Resources.Map;
            this.button7.Label = "button7";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // MySplitButton
            // 
            this.MySplitButton.Image = global::WordAddInDemo.Properties.Resources.List3;
            this.MySplitButton.Label = "测试SplitButton";
            this.MySplitButton.Name = "MySplitButton";
            // 
            // MyGallery
            // 
            this.MyGallery.Buttons.Add(this.button8);
            this.MyGallery.Buttons.Add(this.button9);
            this.MyGallery.Buttons.Add(this.button10);
            this.MyGallery.Image = global::WordAddInDemo.Properties.Resources.List2;
            ribbonDropDownItemImpl6.Image = global::WordAddInDemo.Properties.Resources.MenuItem;
            ribbonDropDownItemImpl6.Label = "Item0";
            ribbonDropDownItemImpl7.Image = global::WordAddInDemo.Properties.Resources.MenuItem;
            ribbonDropDownItemImpl7.Label = "Item1";
            ribbonDropDownItemImpl8.Image = global::WordAddInDemo.Properties.Resources.MenuItem;
            ribbonDropDownItemImpl8.Label = "Item2";
            this.MyGallery.Items.Add(ribbonDropDownItemImpl6);
            this.MyGallery.Items.Add(ribbonDropDownItemImpl7);
            this.MyGallery.Items.Add(ribbonDropDownItemImpl8);
            this.MyGallery.Label = "测试Gallery";
            this.MyGallery.Name = "MyGallery";
            this.MyGallery.ShowImage = true;
            // 
            // button8
            // 
            this.button8.Image = global::WordAddInDemo.Properties.Resources.Map;
            this.button8.Label = "button8";
            this.button8.Name = "button8";
            this.button8.ShowImage = true;
            // 
            // button9
            // 
            this.button9.Image = global::WordAddInDemo.Properties.Resources.Map;
            this.button9.Label = "button9";
            this.button9.Name = "button9";
            this.button9.ShowImage = true;
            // 
            // button10
            // 
            this.button10.Image = global::WordAddInDemo.Properties.Resources.Map;
            this.button10.Label = "button10";
            this.button10.Name = "button10";
            this.button10.ShowImage = true;
            // 
            // group5
            // 
            this.group5.Items.Add(this.buttonGroup1);
            this.group5.Items.Add(this.box2);
            this.group5.Label = "自定义组";
            this.group5.Name = "group5";
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.GroupBtn1);
            this.buttonGroup1.Items.Add(this.GroupBtn2);
            this.buttonGroup1.Items.Add(this.GroupBtn3);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // GroupBtn1
            // 
            this.GroupBtn1.Image = global::WordAddInDemo.Properties.Resources.MenuItem;
            this.GroupBtn1.Label = "设置项1";
            this.GroupBtn1.Name = "GroupBtn1";
            this.GroupBtn1.ShowImage = true;
            // 
            // GroupBtn2
            // 
            this.GroupBtn2.Image = global::WordAddInDemo.Properties.Resources.MenuItem;
            this.GroupBtn2.Label = "设置项2";
            this.GroupBtn2.Name = "GroupBtn2";
            this.GroupBtn2.ShowImage = true;
            // 
            // GroupBtn3
            // 
            this.GroupBtn3.Image = global::WordAddInDemo.Properties.Resources.MenuItem;
            this.GroupBtn3.Label = "设置项3";
            this.GroupBtn3.Name = "GroupBtn3";
            this.GroupBtn3.ShowImage = true;
            // 
            // box2
            // 
            this.box2.Items.Add(this.ToggleButtonSelectAll);
            this.box2.Items.Add(this.ToggleBtnSwitchFont);
            this.box2.Name = "box2";
            // 
            // ToggleButtonSelectAll
            // 
            this.ToggleButtonSelectAll.Image = global::WordAddInDemo.Properties.Resources.Select;
            this.ToggleButtonSelectAll.Label = "选择全部";
            this.ToggleButtonSelectAll.Name = "ToggleButtonSelectAll";
            this.ToggleButtonSelectAll.ShowImage = true;
            // 
            // ToggleBtnSwitchFont
            // 
            this.ToggleBtnSwitchFont.Image = global::WordAddInDemo.Properties.Resources.Pen;
            this.ToggleBtnSwitchFont.Label = "切换文本样式";
            this.ToggleBtnSwitchFont.Name = "ToggleBtnSwitchFont";
            this.ToggleBtnSwitchFont.ShowImage = true;
            // 
            // group6
            // 
            this.group6.Items.Add(this.MyCheckBox);
            this.group6.Items.Add(this.label2);
            this.group6.Label = "其它";
            this.group6.Name = "group6";
            // 
            // MyCheckBox
            // 
            this.MyCheckBox.Label = "测试CheckBox";
            this.MyCheckBox.Name = "MyCheckBox";
            // 
            // label2
            // 
            this.label2.Label = "文本信息";
            this.label2.Name = "label2";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group3);
            this.tab2.Label = "测试Tab2";
            this.tab2.Name = "tab2";
            this.tab2.Position = this.Factory.RibbonPosition.AfterOfficeId("TabAddIns");
            // 
            // group3
            // 
            this.group3.Label = "group3";
            this.group3.Name = "group3";
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Group1.ResumeLayout(false);
            this.Group1.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAddTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox MyComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAddImage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAddText;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu MenuAddBasicInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu MenuAddWinFormControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAddDatePicker;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAddRichText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAddDropDownList;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox TextInputField;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAddInputText;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GroupBtn1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GroupBtn2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GroupBtn3;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox MyCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton MySplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown MyDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery MyGallery;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton ToggleButtonSelectAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton ToggleBtnSwitchFont;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        private Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        private Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        private Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        private Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        private Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        private Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        private Microsoft.Office.Tools.Ribbon.RibbonButton button10;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon MyRibbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
