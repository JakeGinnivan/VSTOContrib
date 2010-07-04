namespace FacebookToOutlookAddin
{
    partial class AppointmentInspectorRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AppointmentInspectorRibbon()
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
            this.facebookGroup = this.Factory.CreateRibbonGroup();
            this.showFacebookPaneButton = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.facebookGroup.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabAppointment";
            this.tab1.Groups.Add(this.facebookGroup);
            this.tab1.Label = "TabAppointment";
            this.tab1.Name = "tab1";
            // 
            // facebookGroup
            // 
            this.facebookGroup.Items.Add(this.showFacebookPaneButton);
            this.facebookGroup.Label = "Facebook";
            this.facebookGroup.Name = "facebookGroup";
            // 
            // showFacebookPaneButton
            // 
            this.showFacebookPaneButton.Label = "Show Facebook Panel";
            this.showFacebookPaneButton.Name = "showFacebookPaneButton";
            this.showFacebookPaneButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TogglePanelVisibility);
            // 
            // AppointmentInspectorRibbon
            // 
            this.Name = "AppointmentInspectorRibbon";
            this.RibbonType = "Microsoft.Outlook.Appointment";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AppointmentRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.facebookGroup.ResumeLayout(false);
            this.facebookGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup facebookGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton showFacebookPaneButton;


    }

    partial class ThisRibbonCollection
    {
        internal AppointmentInspectorRibbon AppointmentInspectorRibbon
        {
            get { return this.GetRibbon<AppointmentInspectorRibbon>(); }
        }
    }
}
