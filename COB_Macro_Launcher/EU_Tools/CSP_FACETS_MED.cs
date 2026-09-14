using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using COB_Macro_Launcher.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace COB_Macro_Launcher.EU_Tools
{
    internal class CSP_FACETS_MED
    {
        public static void Start()
        {
            string MEMBERID = "A"; // T1
            string LNAME = "B"; // T2
            string FNAME = "C"; // T3
            string VERF_METH = "D"; // T4
            string POLICY_NO = "E"; // T5
            string IMP_DT = "F"; // T6
            string CARR_ID = "G"; // T7
            string TERM_DT = "H"; // T8
            string VERF_NAM = "I"; // T9

            string Data_Row = TopMostDialogs.InputBox("What row does your data begin on?.", "Data_Row");

            int.TryParse(Data_Row, out int DR);

            if (DR <= 0)
            {
                TopMostDialogs.MessageBox(Program.LauncherOwner, "Please enter a valid Excel row number.", "Invalid Row", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                return;
            }

            Excel.Application excelApp = null;
            Excel.Worksheet activeSheet = null;
            Excel.Range usedRange = null;

            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                activeSheet = (Excel.Worksheet)excelApp.ActiveSheet;
                usedRange = activeSheet.UsedRange;

                object[,] values = usedRange == null ? null : usedRange.Value2 as object[,];
                if (values == null)
                {
                    return;
                }

                int startIndex = Math.Max(DR, usedRange.Row) - usedRange.Row + 1;
                if (startIndex < 1 || startIndex > values.GetLength(0))
                {
                    TopMostDialogs.MessageBox(Program.LauncherOwner, "The starting row is outside the worksheet's used range.", "Invalid Row", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    return;
                }

                for (int row = startIndex; row <= values.GetLength(0); row++)
                {
                    string memberId = GetCellText(values, row, MEMBERID);
                    string lName = GetCellText(values, row, LNAME);
                    string fName = GetCellText(values, row, FNAME);
                    string verfMeth = GetCellText(values, row, VERF_METH);
                    string policyNo = GetCellText(values, row, POLICY_NO);
                    string impDt = GetCellText(values, row, IMP_DT);
                    string carrId = GetCellText(values, row, CARR_ID);
                    string termDt = GetCellText(values, row, TERM_DT);
                    string verfNam = GetCellText(values, row, VERF_NAM);

                    if (string.IsNullOrWhiteSpace(memberId) &&
                        string.IsNullOrWhiteSpace(lName) &&
                        string.IsNullOrWhiteSpace(fName) &&
                        string.IsNullOrWhiteSpace(verfMeth) &&
                        string.IsNullOrWhiteSpace(policyNo) &&
                        string.IsNullOrWhiteSpace(impDt) &&
                        string.IsNullOrWhiteSpace(carrId) &&
                        string.IsNullOrWhiteSpace(termDt) &&
                        string.IsNullOrWhiteSpace(verfNam))
                    {
                        continue;
                    }

                    // Existing FACETS_MED row processing should continue here.
                }
            }
            catch (Exception ex)
            {
                TopMostDialogs.MessageBox(Program.LauncherOwner, "Could not access Excel active sheet.\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        private static string GetCellText(object[,] values, int row, string columnLetter)
        {
            int column = GetColumnNumber(columnLetter);
            if (column < 1 || column > values.GetLength(1))
            {
                return string.Empty;
            }

            object value = values[row, column];
            return value == null ? string.Empty : value.ToString().Trim();
        }

        private static int GetColumnNumber(string columnLetter)
        {
            if (string.IsNullOrWhiteSpace(columnLetter))
            {
                return 0;
            }

            int columnNumber = 0;

            foreach (char character in columnLetter.Trim().ToUpperInvariant())
            {
                if (character < 'A' || character > 'Z')
                {
                    return 0;
                }

                columnNumber = (columnNumber * 26) + (character - 'A' + 1);
            }

            return columnNumber;
        }
    }
}