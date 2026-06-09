using System;
using System.Collections.Generic;
using System.Deployment.Application;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace LibEndUserToolLauncher
{
    /// <summary>
    /// Main Form used to list all available macros to user
    /// </summary>
    public partial class FormMenu : Form
    {

        #region DLL imports and related variables
        // DLLs used to manage hotkeys
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);
        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")]
        public static extern IntPtr GetActiveWindow();

        // DLLs used to show/hide console window
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        // values related to toggling console visibility
        private const int SW_HIDE = 0;
        private const int SW_SHOWMINIMIZED = 2;
        private const int SW_SHOW = 5;
        readonly IntPtr ConsoleHandle = GetConsoleWindow();
        /// <summary>
        /// Handle keystrokes, activate corresponding end user tool
        /// </summary>
        /// <param name="m"></param>
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312)
            {
                //loop through array of registered tools
                //searching a match to the registered hotkey
                int KeyPressed = m.WParam.ToInt32();
                for (int i = 0; i < Tools.Count; i++)
                {
                    //if the key that was pressed matches one of the 
                    //registered hotkeys
                    if (KeyPressed == Tools[i].ActivationHotKeyId)
                    {
                        //if Action is not null, run that method
                        Tools[i].Action?.Invoke();
                        break;
                    }
                }
            }
            base.WndProc(ref m);
        }
        #endregion
        #region Properties
        private bool showConsole { get; set; }
        /// <summary>
        /// Change visibility of console for debugging purposes
        /// </summary>
        public bool ShowConsole 
        {
            get { return showConsole; }
            set
            {
                if (value) ConsoleShow();
                else ConsoleHide();
                showConsole = value;
            }
        }
        private bool showForm { get; set; }
        /// <summary>
        /// Change visibility of console for debugging purposes
        /// </summary>
        public bool ShowForm
        {
            get { return showForm; }
            set
            {
                if (value) Show();
                else Hide();
                showForm = value;
            }
        }
        private bool allowShorcutChanges { get; set; }
        /// <summary>
        /// Allow users to change the shortcut keys for the tools in the list from their default predefined values to whatever they want
        /// </summary>
        public bool AllowShorcutChanges
        {
            get { return allowShorcutChanges; }
            set
            {
                allowShorcutChanges = value;
            }
        }
        /// <summary>
        /// Title of the launcher window. Also the system tray 
        /// </summary>
        public string Title {  get; set; }
        /// <summary>
        /// List of all macros registered to the tool
        /// </summary>
        private List<EutTool> tools;
        /// <summary>
        /// List of all end user tools registered to the launcher. Setting this property will populate the launcher Form with the provided tools and register their hotkeys.
        /// </summary>
        public List<EutTool> Tools
        {
            get 
            {   
                if (tools == null)
                    tools = new List<EutTool>();
                return tools; 
            }
            set { tools = value; }
        }
        #endregion
        #region Event handlers
        private void FormMenu_Shown(object sender, EventArgs e)
        {
            CreateForm();
        }
        private void FormMenu_Load(object sender, EventArgs e)
        {
            // at one time, it worked to have the logic to create the form in this event handler, but for some reason it stopped working and I had to move it to the Shown event handler. Leaving this note here in case I need to move it back or anything in the future
        }
        /// <summary>
        /// Minimize window to system tray
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMenu_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == WindowState)
            {
                Hide();
            }
        }
        /// <summary>
        /// Close window to system tray
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMenu_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                Hide();
                e.Cancel = true;
            }
        }
        /// <summary>
        /// Show Form when system tray icon is double clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NotifyIconSystemTray_DoubleClick(object sender, EventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
        }
        /// <summary>
        /// Show Form when system tray option is selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolStripMenuItemRestore_Click(object sender, EventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
        }
        /// <summary>
        /// Close application when system tray option is selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolStripMenuItemClose_Click(object sender, EventArgs e)
        {
            ConsoleHide();
            NotifyIconSystemTray.Visible = false;
            Close();
            System.Windows.Forms.Application.Exit();
        }
        /// <summary>
        /// Show console when system tray option is selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolStripMenuItemConsoleShow_Click(object sender, EventArgs e)
        {
            ConsoleShow();
        }
        /// <summary>
        /// Hide console when system tray option is selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolStripMenuItemConsoleHide_Click(object sender, EventArgs e)
        {
            ConsoleHide();
        }
        private void DataGridViewTools_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridView clickedGrid = (DataGridView)sender;
                DataGridViewColumn clickedColumn = clickedGrid.Columns[e.ColumnIndex];
                DataGridViewRow clickedRow = clickedGrid.Rows[e.RowIndex];
                EutTool clickedTool = (EutTool)clickedRow.Tag;
                switch (e.ColumnIndex)
                {
                    //if button in Run column was clicked
                    case 0:
                        //run the EutTool's Action method if Enabled
                        if ((bool)clickedRow.Cells[4].Value)
                            clickedTool.Action?.Invoke();
                        break;
                    case 4:
                        //force the CellValueChanged event when a CheckBox value is changed
                        clickedGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                        break;
                }
            }
        }

        private void DataGridViewTools_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridView clickedGrid = (DataGridView)sender;
                DataGridViewColumn clickedColumn = clickedGrid.Columns[e.ColumnIndex];
                DataGridViewRow clickedRow = clickedGrid.Rows[e.RowIndex];
                DataGridViewCell clickedCell = clickedGrid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                EutTool clickedTool = (EutTool)clickedGrid.Rows[e.RowIndex].Tag;
                //handle this event differently depending on which column caused the event
                switch (e.ColumnIndex)
                {
                    //if Activation TextBox changes
                    case 2:
                        string ActivationStringNew = clickedCell.Value.ToString();
                        //validate new ActivationString
                        if (!clickedTool.ActivationStringIsValid(ActivationStringNew))
                        {
                            //supplied ActivationString does not match acceptable regex pattern
                            //reset the TextBox to old ActivationString (without triggering event)
                            DataGridViewTools.CellValueChanged -= DataGridViewTools_CellValueChanged;
                            clickedCell.Value = clickedTool.ActivationString;
                            DataGridViewTools.CellValueChanged += DataGridViewTools_CellValueChanged;
                        }
                        else
                        {
                            ActivationStringNew = clickedTool.ActivationStringReformat(ActivationStringNew);
                            //check against all other ActivationStrings in DataGridView for duplicate
                            bool DuplicateFound = false;
                            foreach (DataGridViewRow CurrentRow in DataGridViewTools.Rows)
                            {
                                DataGridViewCell CurrentCell = CurrentRow.Cells[2];
                                //if duplicate found
                                if (ActivationStringNew == CurrentCell.Value.ToString() && !CurrentCell.Equals(clickedCell))
                                {
                                    clickedCell.Style.BackColor = Color.Red;
                                    CurrentCell.Style.BackColor = Color.Red;
                                    DuplicateFound = true;
                                }
                                else
                                {
                                    clickedCell.Style.BackColor = clickedColumn.DefaultCellStyle.BackColor;
                                    CurrentCell.Style.BackColor = clickedColumn.DefaultCellStyle.BackColor;
                                }
                            }
                            if (!DuplicateFound)
                            {
                                //set EutTool.ActivationString, which calculates and
                                //sets EutTool.Key and EutTool.Modifier properties
                                clickedTool.ActivationString = ActivationStringNew;
                                DataGridViewTools.CellValueChanged -= DataGridViewTools_CellValueChanged;
                                clickedCell.Value = clickedTool.ActivationString;
                                DataGridViewTools.CellValueChanged += DataGridViewTools_CellValueChanged;

                                //re-register new hotkey for this tool
                                UnregisterHotKey(Handle, clickedTool.ActivationHotKeyId);
                                RegisterHotKey
                                (
                                    Handle,
                                    clickedTool.ActivationHotKeyId,
                                    clickedTool.Modifier,
                                    (int)clickedTool.Key
                                );
                            }
                        }

                        break;
                    //if CheckBox in the Enabled column has been altered
                    case 4:
                        //update the Enabled property in the associated EutTool
                        clickedTool.Enabled = (bool)clickedCell.Value;
                        break;
                }
            }
        }

        private void DataGridViewTools_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewTools.ClearSelection();
        }
        #endregion
        #region Private methods
        /// <summary>
        /// Build the Form, including creating a keystroke listener for each EutTool 
        /// from the FormMenu.Tools property and adding it to FormMenu.ListViewTools
        /// </summary>
        private void CreateForm()
        {
            //set the tile bar text and system tray text to the value of the Title property if it is not null
            if (Title != null)
            {
                Text = Title;
                NotifyIconSystemTray.Text = Title;
            }
            //set version number in system tray menu
            if (ApplicationDeployment.IsNetworkDeployed)
                ToolStripMenuItemVersion.Text = $"Version {ApplicationDeployment.CurrentDeployment.CurrentVersion}";
            else
                ToolStripMenuItemVersion.Text = $"Version {FileVersionInfo.GetVersionInfo(new StackFrame(0).GetMethod().Module.Assembly.Location).FileVersion}";
            //set Activation column to a monospace font to make it easier to distinguish
            //between things like a lowercase l and capital I
            DataGridViewTools.Columns[2].DefaultCellStyle.Font = new Font(FontFamily.GenericMonospace, DataGridViewTools.DefaultCellStyle.Font.Size);

            //fill DataGridViewTools and system tray with available macros and register hotkey for each
            List<string> AllKeyCombos = new List<string>();
            if (Tools != null)
            {
                int HotkeyId = 1;
                foreach (EutTool CurrentTool in Tools)
                {
                    if (AllKeyCombos.Contains(CurrentTool.ActivationString))
                    {
                        throw new Exception("Duplicate activation key combos are not allowed");
                    }
                    else
                    {
                        //register hotkey activation for this tool
                        if (CurrentTool.Key != Keys.None)
                        {
                            CurrentTool.ActivationHotKeyId = HotkeyId++;
                            AllKeyCombos.Add(CurrentTool.ActivationString);
                            RegisterHotKey
                            (
                                Handle,
                                CurrentTool.ActivationHotKeyId,
                                CurrentTool.Modifier,
                                (int)CurrentTool.Key
                            );
                        }

                        //create new menu item in system tray for this tool, with click event to run the tool's Action method
                        ToolStripMenuItem CurrentToolMenuItem = new ToolStripMenuItem(CurrentTool.Name);
                        CurrentToolMenuItem.Click += (sender, e) =>
                        {
                            if (CurrentTool.Enabled)
                                CurrentTool.Action?.Invoke();
                        };
                        ToolStripMenuItemTools.DropDownItems.Add(CurrentToolMenuItem);

                        //create new row for DataGridViewTools
                        DataGridViewRow CurrentRow = new DataGridViewRow
                        {
                            //set this row's Tag property to the EutTool object that this row represents
                            Tag = CurrentTool
                        };
                        CurrentRow.Cells.AddRange(new DataGridViewCell[]
                        {
                            new DataGridViewButtonCell()
                            {
                                Value = "▶"
                            },
                            new DataGridViewTextBoxCell()
                            {
                                Value = CurrentTool.Name
                            },
                            new DataGridViewTextBoxCell()
                            {
                                Value = CurrentTool.ActivationStringFriendly
                            },
                            new DataGridViewTextBoxCell()
                            {
                                Value = CurrentTool.Description
                            },
                            new DataGridViewCheckBoxCell()
                            {
                                Value = true
                            }
                        });
                        //for whatever reason, the ReadOnly property for a cell has to be set AFTER the cell is added to the row, not while it is being created
                        CurrentRow.Cells[2].ReadOnly = !allowShorcutChanges;
                        DataGridViewTools.Rows.Add(CurrentRow);
                    }
                }
            }
            //set to center of display
            Left = (Screen.FromControl(Controls[0]).WorkingArea.Width / 2) - (Width / 2);
            Top = (Screen.FromControl(Controls[0]).WorkingArea.Height / 2) - (Height / 2);
            if (!showConsole) ConsoleHide();
            if (!showForm) Hide();
            BringToFront();
        }

        private void ConsoleShow()
        {
            ShowWindow(ConsoleHandle, SW_SHOW);
        }
        private void ConsoleHide()
        {
            ShowWindow(ConsoleHandle, SW_SHOWMINIMIZED);
            ShowWindow(ConsoleHandle, SW_HIDE);
        }
        #endregion
        #region Public methods
        /// <summary>
        /// Begin running the Form
        /// </summary>
        public void Run()
        {
            System.Windows.Forms.Application.Run(this);
        }
        /// <summary>
        /// Add a tool to the list of tools in the launcher at a specified position. If position is greater than the number of tools in the list, the tool will be added to the end of the list.
        /// </summary>
        /// <param name="ToolToAdd">The tool to add to the launcher.</param>
        /// <param name="Position">The position at which to insert the tool.</param>
        public void AddTool(EutTool ToolToAdd, int Position)
        {
            tools.Insert(Position, ToolToAdd);
        }
        /// <summary>
        /// Add a tool to the end of the list of tools in the launcher.
        /// </summary>
        /// <param name="ToolToAdd">The tool to add to the launcher.</param>
        public void AddTool(EutTool ToolToAdd)
        {
            tools.Add(ToolToAdd);
        }
        /// <summary>
        /// Remove a tool from the list of tools in the launcher by specifying the tool to remove. If the specified tool is not in the list, no changes will be made.
        /// </summary>
        /// <param name="ToolToRemove">The tool to remove from the launcher.</param>
        public void RemoveTool(EutTool ToolToRemove)
        {
            if(tools.Contains(ToolToRemove))
                tools.Remove(ToolToRemove);
        }
        /// <summary>
        /// Remove a tool from the list of tools in the launcher by specifying the position of the tool to remove. If the specified position is out of range, no changes will be made.
        /// </summary>
        /// <param name="Position">The 0-based position of the tool to remove from the launcher.</param>
        public void RemoveTool(int Position)
        {
            if(Position >= 0 && Position < tools.Count)
                tools.RemoveAt(Position);
        }
        /// <summary>
        /// Clear all tools from the list
        /// </summary>
        public void ClearTools()
        {
            tools.Clear();
        }
        #endregion
        /// <summary>
        /// Constructor for the menu window that displays to the user all available tools and their corresponding shortcut keys, descriptions, and options to enable/disable.
        /// </summary>
        public FormMenu()
        {
            //default values for properties
            showConsole = false;
            showForm = true;
            allowShorcutChanges = false;

            //todo: add new "Help" column with a button that navigates to the proper ServiceNow page to submit a help ticket for that tool, inputting all the relevant information and assigning it to the appropriate group

            InitializeComponent();
        }
    }
}
