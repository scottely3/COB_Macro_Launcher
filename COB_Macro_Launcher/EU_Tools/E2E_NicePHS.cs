using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;
using COB_Macro_Launcher.UI;

namespace COB_Macro_Launcher.EU_Tools
{
    internal class E2E_NicePHS
    {
        public static void Start()
        {
            string COMP_CODE = "A"; //T1
            string MEMBERID = "B"; //T2
            string COMMENT = "C"; //T3

            string Data_Row = "";
            Data_Row = TopMostDialogs.InputBox("What row does your data begin on?.", "Data_Row");

            int.TryParse(Data_Row, out int DR);

            // Excel interop setup
            Excel.Application excelApp = null;
            Excel.Range activeCell = null;
            object cellValue = null;
            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                activeCell = (Excel.Range)excelApp.ActiveCell;
                cellValue = activeCell.Value;
            }
            catch (Exception ex)
            {
                TopMostDialogs.MessageBox(Program.LauncherOwner, "Could not access Excel active cell.\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
        }
    }
}
