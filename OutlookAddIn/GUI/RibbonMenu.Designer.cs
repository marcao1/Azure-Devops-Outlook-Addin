﻿using System;
using Microsoft.Office.Tools.Ribbon;

namespace OutlookAddIn
{
    partial class RibbonMenu : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMenu()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonMenu));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.dropDownOrg = this.Factory.CreateRibbonDropDown();
            this.dropDownProj = this.Factory.CreateRibbonDropDown();
            this.editOrgBtn = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.dropDownType = this.Factory.CreateRibbonDropDown();
            this.dropDownCol = this.Factory.CreateRibbonDropDown();
            this.AddBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "Azure Devops";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Azure Devops Login";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.dropDownOrg);
            this.group2.Items.Add(this.dropDownProj);
            this.group2.Items.Add(this.editOrgBtn);
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
            // editOrgBtn
            // 
            this.editOrgBtn.Label = "Edit Organisations";
            this.editOrgBtn.Name = "editOrgBtn";
            this.editOrgBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editOrgBtn_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.dropDownType);
            this.group3.Items.Add(this.dropDownCol);
            this.group3.Items.Add(this.AddBtn);
            this.group3.Name = "group3";
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
            // AddBtn
            // 
            this.AddBtn.Label = "Add Item";
            this.AddBtn.Name = "AddBtn";
            this.AddBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddBtn_Click);
            // 
            // RibbonMenu
            // 
            this.Name = "RibbonMenu";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonMenu_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

     
        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownOrg;
        public Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownProj;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton editOrgBtn;
        internal RibbonGroup group3;
        internal RibbonDropDown dropDownType;
        internal RibbonDropDown dropDownCol;
        internal RibbonButton AddBtn;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonMenu RibbonMenu
        {
            get { return this.GetRibbon<RibbonMenu>(); }
        }
    }
}
