namespace OutlookAddIn
{
    partial class RibbonMail : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMail()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">"true", wenn verwaltete Ressourcen gelöscht werden sollen, andernfalls "false".</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.dropDownOrg = this.Factory.CreateRibbonDropDown();
            this.dropDownProj = this.Factory.CreateRibbonDropDown();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.dropDownType = this.Factory.CreateRibbonDropDown();
            this.dropDownCol = this.Factory.CreateRibbonDropDown();
            this.addItemBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Azure Devops";
            this.tab1.Name = "tab1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.dropDownOrg);
            this.group2.Items.Add(this.dropDownProj);
            this.group2.Items.Add(this.button1);
            this.group2.Name = "group2";
            // 
            // dropDownOrg
            // 
            this.dropDownOrg.Label = "Organisation";
            this.dropDownOrg.Name = "dropDownOrg";
            this.dropDownOrg.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDownOrg_SelectionChanged);
            // 
            // dropDownProj
            // 
            this.dropDownProj.Label = "Project";
            this.dropDownProj.Name = "dropDownProj";
            // 
            // button1
            // 
            this.button1.Label = "Edit Organisation";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.dropDownType);
            this.group1.Items.Add(this.dropDownCol);
            this.group1.Items.Add(this.addItemBtn);
            this.group1.Name = "group1";
            // 
            // dropDownType
            // 
            this.dropDownType.Label = "Typ";
            this.dropDownType.Name = "dropDownType";
            // 
            // dropDownCol
            // 
            this.dropDownCol.Label = "Board Column";
            this.dropDownCol.Name = "dropDownCol";
            // 
            // addItemBtn
            // 
            this.addItemBtn.Label = "Add Item";
            this.addItemBtn.Name = "addItemBtn";
            this.addItemBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // RibbonMail
            // 
            this.Name = "RibbonMail";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonMail_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownOrg;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownProj;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownType;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownCol;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addItemBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonMail RibbonMail
        {
            get { return this.GetRibbon<RibbonMail>(); }
        }
    }
}
