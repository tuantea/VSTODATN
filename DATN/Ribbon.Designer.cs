namespace DATN
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Algorithm = this.Factory.CreateRibbonGroup();
            this.buttonColorize = this.Factory.CreateRibbonButton();
            this.dropDownColorRGB = this.Factory.CreateRibbonButton();
            this.editSaturationPeak = this.Factory.CreateRibbonButton();
            this.checkBoxInvert = this.Factory.CreateRibbonButton();
            this.Speed = this.Factory.CreateRibbonGroup();
            this.SpeedText = this.Factory.CreateRibbonButton();
            this.Finance = this.Factory.CreateRibbonGroup();
            this.Exchange = this.Factory.CreateRibbonButton();
            this.ReadNumber = this.Factory.CreateRibbonButton();
            this.Json = this.Factory.CreateRibbonGroup();
            this.Export = this.Factory.CreateRibbonButton();
            this.Import = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Algorithm.SuspendLayout();
            this.Speed.SuspendLayout();
            this.Finance.SuspendLayout();
            this.Json.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Algorithm);
            this.tab1.Groups.Add(this.Speed);
            this.tab1.Groups.Add(this.Finance);
            this.tab1.Groups.Add(this.Json);
            this.tab1.Label = "new tab";
            this.tab1.Name = "tab1";
            // 
            // Algorithm
            // 
            this.Algorithm.Items.Add(this.buttonColorize);
            this.Algorithm.Items.Add(this.dropDownColorRGB);
            this.Algorithm.Items.Add(this.editSaturationPeak);
            this.Algorithm.Items.Add(this.checkBoxInvert);
            this.Algorithm.Label = "Algorithm";
            this.Algorithm.Name = "Algorithm";
            // 
            // buttonColorize
            // 
            this.buttonColorize.Label = "buttonColorize";
            this.buttonColorize.Name = "buttonColorize";
            this.buttonColorize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // dropDownColorRGB
            // 
            this.dropDownColorRGB.Label = "dropDownColorRGB";
            this.dropDownColorRGB.Name = "dropDownColorRGB";
            this.dropDownColorRGB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonImage2Cells_Click);
            // 
            // editSaturationPeak
            // 
            this.editSaturationPeak.Label = "editSaturationPeak";
            this.editSaturationPeak.Name = "editSaturationPeak";
            this.editSaturationPeak.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_QRCoder);
            // 
            // checkBoxInvert
            // 
            this.checkBoxInvert.Label = "checkBoxInvert";
            this.checkBoxInvert.Name = "checkBoxInvert";
            this.checkBoxInvert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_ExportExcelToJson1);
            // 
            // Speed
            // 
            this.Speed.Items.Add(this.SpeedText);
            this.Speed.Label = "Speed";
            this.Speed.Name = "Speed";
            // 
            // SpeedText
            // 
            this.SpeedText.Label = "SpeedText(Eng)";
            this.SpeedText.Name = "SpeedText";
            this.SpeedText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSpeechVN_Click);
            // 
            // Finance
            // 
            this.Finance.Items.Add(this.Exchange);
            this.Finance.Items.Add(this.ReadNumber);
            this.Finance.Label = "Finance";
            this.Finance.Name = "Finance";
            // 
            // Exchange
            // 
            this.Exchange.Label = "Exchange";
            this.Exchange.Name = "Exchange";
            this.Exchange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_ExChange);
            // 
            // ReadNumber
            // 
            this.ReadNumber.Label = "ReadNumber";
            this.ReadNumber.Name = "ReadNumber";
            this.ReadNumber.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_ReadNumber);
            // 
            // Json
            // 
            this.Json.Items.Add(this.Export);
            this.Json.Items.Add(this.Import);
            this.Json.Label = "Json";
            this.Json.Name = "Json";
            // 
            // Export
            // 
            this.Export.Label = "Export";
            this.Export.Name = "Export";
            this.Export.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_ExportExcelToJson);
            // 
            // Import
            // 
            this.Import.Label = "Import";
            this.Import.Name = "Import";
            this.Import.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_ImportJsonToExcel);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Algorithm.ResumeLayout(false);
            this.Algorithm.PerformLayout();
            this.Speed.ResumeLayout(false);
            this.Speed.PerformLayout();
            this.Finance.ResumeLayout(false);
            this.Finance.PerformLayout();
            this.Json.ResumeLayout(false);
            this.Json.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Algorithm;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonColorize;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton dropDownColorRGB;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton editSaturationPeak;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton checkBoxInvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Speed;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SpeedText;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Finance;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ReadNumber;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Json;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Export;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Import;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Exchange;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
