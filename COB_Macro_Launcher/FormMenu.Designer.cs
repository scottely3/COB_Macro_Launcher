namespace LibEndUserToolLauncher
{
    partial class FormMenu
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMenu));
            this.NotifyIconSystemTray = new System.Windows.Forms.NotifyIcon(this.components);
            this.ContextMenuStripSystemTray = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.ToolStripMenuItemRestore = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolStripMenuItemClose = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolStripMenuItemConsole = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolStripMenuItemConsoleShow = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolStripMenuItemConsoleHide = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolStripMenuItemVersion = new System.Windows.Forms.ToolStripMenuItem();
            this.DataGridViewTools = new System.Windows.Forms.DataGridView();
            this.ColumnRun = new System.Windows.Forms.DataGridViewButtonColumn();
            this.ColumnName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnActivation = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnDescription = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnEnabled = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.ToolStripMenuItemTools = new System.Windows.Forms.ToolStripMenuItem();
            this.ContextMenuStripSystemTray.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DataGridViewTools)).BeginInit();
            this.SuspendLayout();
            // 
            // NotifyIconSystemTray
            // 
            this.NotifyIconSystemTray.ContextMenuStrip = this.ContextMenuStripSystemTray;
            this.NotifyIconSystemTray.Icon = ((System.Drawing.Icon)(resources.GetObject("NotifyIconSystemTray.Icon")));
            this.NotifyIconSystemTray.Text = "End User Tool Activator";
            this.NotifyIconSystemTray.Visible = true;
            this.NotifyIconSystemTray.DoubleClick += new System.EventHandler(this.NotifyIconSystemTray_DoubleClick);
            // 
            // ContextMenuStripSystemTray
            // 
            this.ContextMenuStripSystemTray.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolStripMenuItemRestore,
            this.ToolStripMenuItemTools,
            this.ToolStripMenuItemClose,
            this.ToolStripMenuItemConsole,
            this.ToolStripMenuItemVersion});
            this.ContextMenuStripSystemTray.Name = "ContextMenuStripSystemTray";
            this.ContextMenuStripSystemTray.Size = new System.Drawing.Size(198, 136);
            // 
            // ToolStripMenuItemRestore
            // 
            this.ToolStripMenuItemRestore.Name = "ToolStripMenuItemRestore";
            this.ToolStripMenuItemRestore.Size = new System.Drawing.Size(197, 22);
            this.ToolStripMenuItemRestore.Text = "Show launcher window";
            this.ToolStripMenuItemRestore.Click += new System.EventHandler(this.ToolStripMenuItemRestore_Click);
            // 
            // ToolStripMenuItemClose
            // 
            this.ToolStripMenuItemClose.Name = "ToolStripMenuItemClose";
            this.ToolStripMenuItemClose.Size = new System.Drawing.Size(197, 22);
            this.ToolStripMenuItemClose.Text = "Close";
            this.ToolStripMenuItemClose.Click += new System.EventHandler(this.ToolStripMenuItemClose_Click);
            // 
            // ToolStripMenuItemConsole
            // 
            this.ToolStripMenuItemConsole.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolStripMenuItemConsoleShow,
            this.ToolStripMenuItemConsoleHide});
            this.ToolStripMenuItemConsole.Name = "ToolStripMenuItemConsole";
            this.ToolStripMenuItemConsole.Size = new System.Drawing.Size(197, 22);
            this.ToolStripMenuItemConsole.Text = "Console";
            this.ToolStripMenuItemConsole.Click += new System.EventHandler(this.ToolStripMenuItemConsoleShow_Click);
            // 
            // ToolStripMenuItemConsoleShow
            // 
            this.ToolStripMenuItemConsoleShow.Name = "ToolStripMenuItemConsoleShow";
            this.ToolStripMenuItemConsoleShow.Size = new System.Drawing.Size(180, 22);
            this.ToolStripMenuItemConsoleShow.Text = "Show";
            this.ToolStripMenuItemConsoleShow.Click += new System.EventHandler(this.ToolStripMenuItemConsoleShow_Click);
            // 
            // ToolStripMenuItemConsoleHide
            // 
            this.ToolStripMenuItemConsoleHide.Name = "ToolStripMenuItemConsoleHide";
            this.ToolStripMenuItemConsoleHide.Size = new System.Drawing.Size(180, 22);
            this.ToolStripMenuItemConsoleHide.Text = "Hide";
            this.ToolStripMenuItemConsoleHide.Click += new System.EventHandler(this.ToolStripMenuItemConsoleHide_Click);
            // 
            // ToolStripMenuItemVersion
            // 
            this.ToolStripMenuItemVersion.Name = "ToolStripMenuItemVersion";
            this.ToolStripMenuItemVersion.Size = new System.Drawing.Size(197, 22);
            this.ToolStripMenuItemVersion.Text = "Version";
            // 
            // DataGridViewTools
            // 
            this.DataGridViewTools.AllowUserToAddRows = false;
            this.DataGridViewTools.AllowUserToDeleteRows = false;
            this.DataGridViewTools.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DataGridViewTools.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            this.DataGridViewTools.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.DataGridViewTools.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DataGridViewTools.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnRun,
            this.ColumnName,
            this.ColumnActivation,
            this.ColumnDescription,
            this.ColumnEnabled});
            this.DataGridViewTools.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.DataGridViewTools.Location = new System.Drawing.Point(13, 13);
            this.DataGridViewTools.MultiSelect = false;
            this.DataGridViewTools.Name = "DataGridViewTools";
            this.DataGridViewTools.RowHeadersVisible = false;
            this.DataGridViewTools.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.DataGridViewTools.Size = new System.Drawing.Size(755, 436);
            this.DataGridViewTools.TabIndex = 1;
            this.DataGridViewTools.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridViewTools_CellContentClick);
            this.DataGridViewTools.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridViewTools_CellValueChanged);
            this.DataGridViewTools.SelectionChanged += new System.EventHandler(this.DataGridViewTools_SelectionChanged);
            // 
            // ColumnRun
            // 
            this.ColumnRun.HeaderText = "Run";
            this.ColumnRun.Name = "ColumnRun";
            this.ColumnRun.Text = "Run";
            this.ColumnRun.ToolTipText = "Run this end user tool now";
            this.ColumnRun.UseColumnTextForButtonValue = true;
            this.ColumnRun.Width = 35;
            // 
            // ColumnName
            // 
            this.ColumnName.HeaderText = "Name";
            this.ColumnName.Name = "ColumnName";
            this.ColumnName.ReadOnly = true;
            this.ColumnName.ToolTipText = "Name of this end user tool";
            this.ColumnName.Width = 200;
            // 
            // ColumnActivation
            // 
            this.ColumnActivation.HeaderText = "Activation";
            this.ColumnActivation.Name = "ColumnActivation";
            this.ColumnActivation.ToolTipText = "Key combo used to activate/run this tool";
            this.ColumnActivation.Width = 150;
            // 
            // ColumnDescription
            // 
            this.ColumnDescription.HeaderText = "Description";
            this.ColumnDescription.Name = "ColumnDescription";
            this.ColumnDescription.ReadOnly = true;
            this.ColumnDescription.ToolTipText = "Quick description of what the end user tool does";
            this.ColumnDescription.Width = 300;
            // 
            // ColumnEnabled
            // 
            this.ColumnEnabled.FalseValue = "0";
            this.ColumnEnabled.HeaderText = "Enabled";
            this.ColumnEnabled.Name = "ColumnEnabled";
            this.ColumnEnabled.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.ColumnEnabled.ToolTipText = "Whether or not this tool is currently enabled";
            this.ColumnEnabled.TrueValue = "1";
            this.ColumnEnabled.Width = 50;
            // 
            // ToolStripMenuItemTools
            // 
            this.ToolStripMenuItemTools.Name = "ToolStripMenuItemTools";
            this.ToolStripMenuItemTools.Size = new System.Drawing.Size(197, 22);
            this.ToolStripMenuItemTools.Text = "Tools";
            // 
            // FormMenu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(780, 461);
            this.Controls.Add(this.DataGridViewTools);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimizeBox = false;
            this.Name = "FormMenu";
            this.Text = "End User Tool Activator";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormMenu_FormClosing);
            this.Load += new System.EventHandler(this.FormMenu_Load);
            this.Shown += new System.EventHandler(this.FormMenu_Shown);
            this.Resize += new System.EventHandler(this.FormMenu_Resize);
            this.ContextMenuStripSystemTray.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DataGridViewTools)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.NotifyIcon NotifyIconSystemTray;
        private System.Windows.Forms.ContextMenuStrip ContextMenuStripSystemTray;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemRestore;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemClose;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemConsole;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemVersion;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemConsoleShow;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemConsoleHide;
        private System.Windows.Forms.DataGridView DataGridViewTools;
        private System.Windows.Forms.DataGridViewButtonColumn ColumnRun;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnActivation;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnDescription;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColumnEnabled;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemTools;
    }
}