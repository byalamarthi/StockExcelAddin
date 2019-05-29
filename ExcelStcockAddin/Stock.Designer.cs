namespace ExcelStcockAddin
{
    partial class Stock : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Stock()
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
            this.GetStock = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnGetStock = this.Factory.CreateRibbonButton();
            this.btnTemplate = this.Factory.CreateRibbonButton();
            this.GetStock.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // GetStock
            // 
            this.GetStock.Groups.Add(this.group1);
            this.GetStock.Label = "Stock";
            this.GetStock.Name = "GetStock";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnTemplate);
            this.group1.Items.Add(this.btnGetStock);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btnGetStock
            // 
            this.btnGetStock.Label = "GetStockValue";
            this.btnGetStock.Name = "btnGetStock";
            this.btnGetStock.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGetStock_Click);
            // 
            // btnTemplate
            // 
            this.btnTemplate.Label = "Load Template";
            this.btnTemplate.Name = "btnTemplate";
            this.btnTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnTemplate_Click);
            // 
            // Stock
            // 
            this.Name = "Stock";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.GetStock);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Stock_Load);
            this.GetStock.ResumeLayout(false);
            this.GetStock.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab GetStock;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetStock;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTemplate;
    }

    partial class ThisRibbonCollection
    {
        internal Stock Stock
        {
            get { return this.GetRibbon<Stock>(); }
        }
    }
}
