using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using COB_Macro_Launcher.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace COB_Macro_Launcher.EU_Tools
{
    internal class COSMOS_Update_for_CAQH
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const uint MouseEventLeftDown = 0x0002;
        private const uint MouseEventLeftUp = 0x0004;

        private const string CarrierNumberColumn = "A";
        private const string CobRowEffDateColumn = "B";
        private const string CobRowEndDateColumn = "C";
        private const string SubscriberNumberColumn = "D";
        private const string SiteCodeColumn = "E";
        private const string CommentColumn = "F";
        private const string SourceSystemGroupCodeColumn = "G";
        private const string CobIndicatorColumn = "H";
        private const string StatusColumn = "I";

        public static void Start()
        {
            TopMostDialogs.MessageBox(
                Program.LauncherOwner,
                "WARNING: This launcher uses Excel data and fixed screen coordinates to update COSMOS CAQH.",
                "COSMOS Update for CAQH",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);

            string dataRow = TopMostDialogs.InputBox("What row does your data begin on?", "Data_Row");

            int.TryParse(dataRow, out int startRow);
            if (startRow <= 0)
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

                int firstUsedRow = usedRange.Row;
                int firstUsedColumn = usedRange.Column;
                int rowCount = values.GetLength(0);
                int colCount = values.GetLength(1);
                int startIndex = Math.Max(startRow, firstUsedRow) - firstUsedRow + 1;

                if (startIndex < 1 || startIndex > rowCount)
                {
                    TopMostDialogs.MessageBox(Program.LauncherOwner, "The starting row is outside the worksheet's used range.", "Invalid Row", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    return;
                }

                for (int row = startIndex; row <= rowCount; row++)
                {
                    int worksheetRow = firstUsedRow + row - 1;
                    CaqhRecord record = BuildRecord(values, row, firstUsedColumn, colCount);
                    if (record.IsEmpty)
                    {
                        continue;
                    }

                    string status = ProcessRecord(record);
                    WriteStatus(activeSheet, worksheetRow, status);
                    ((Excel.Workbook)activeSheet.Parent).Save();
                }
            }
            catch (Exception ex)
            {
                TopMostDialogs.MessageBox(Program.LauncherOwner, "Could not access Excel active sheet.\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        private static CaqhRecord BuildRecord(object[,] values, int row, int firstUsedColumn, int colCount)
        {
            return new CaqhRecord
            {
                CarrierNumber = GetCellText(values, row, CarrierNumberColumn, firstUsedColumn, colCount),
                CobRowEffDate = GetCellText(values, row, CobRowEffDateColumn, firstUsedColumn, colCount),
                CobRowEndDate = GetCellText(values, row, CobRowEndDateColumn, firstUsedColumn, colCount),
                SubscriberNumber = GetCellText(values, row, SubscriberNumberColumn, firstUsedColumn, colCount),
                SiteCode = GetCellText(values, row, SiteCodeColumn, firstUsedColumn, colCount),
                Comment = GetCellText(values, row, CommentColumn, firstUsedColumn, colCount),
                SourceSystemGroupCode = GetCellText(values, row, SourceSystemGroupCodeColumn, firstUsedColumn, colCount),
                CobIndicator = GetCellText(values, row, CobIndicatorColumn, firstUsedColumn, colCount)
            };
        }

        private static string ProcessRecord(CaqhRecord record)
        {
            if (string.IsNullOrWhiteSpace(record.CarrierNumber)
                || string.IsNullOrWhiteSpace(record.CobRowEffDate)
                || string.IsNullOrWhiteSpace(record.CobRowEndDate)
                || string.IsNullOrWhiteSpace(record.SubscriberNumber)
                || string.IsNullOrWhiteSpace(record.SiteCode)
                || string.IsNullOrWhiteSpace(record.SourceSystemGroupCode))
            {
                return "Missing required Excel data";
            }

            if (!ActivateWindowByTitleContains("A - Emulator - \\\\Remote"))
            {
                return "Emulator window not found";
            }

            try
            {
                RunCosmosFlow(record);
                return "successful";
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }

        private static void RunCosmosFlow(CaqhRecord record)
        {
            string today = DateTime.Now.ToString("MMddyy");

            Pause(300);
            SendKeys.SendWait("{TAB}");
            Pause(200);
            TypeText(record.SiteCode);
            Pause(300);
            SendKeys.SendWait("{ENTER}");
            Pause(200);
            TypeText("CL300");
            Pause(300);
            SendKeys.SendWait("{ENTER}");
            Pause(200);
            TypeText("I");
            Pause(100);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            TypeText(today);
            Pause(100);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            TypeText(record.SourceSystemGroupCode);
            Pause(210);
            TypeText(record.SubscriberNumber);
            Pause(210);
            TypeText("00");
            Pause(210);
            SendKeys.SendWait("{ENTER}");
            Pause(210);
            TypeText("C");
            Pause(210);
            SendKeys.SendWait("{TAB}");
            Pause(110);
            TypeText(today);
            Pause(310);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            TypeText(record.SourceSystemGroupCode);
            Pause(210);
            TypeText(record.SubscriberNumber);
            Pause(210);
            TypeText("00");
            Pause(210);
            SendKeys.SendWait("{TAB}");
            Pause(210);
            TypeText("C");
            Pause(100);
            TypeText(record.CobRowEffDate);
            Pause(310);
            TypeText(record.CobRowEndDate);
            Pause(310);
            TypeText("0");
            Pause(100);
            TypeText(record.CarrierNumber);
            Pause(510);

            Click(320, 605, 1, 210);
            TypeText(record.CobIndicator ?? string.Empty);
            Pause(310);

            SendKeys.SendWait("{ENTER}");
            Pause(210);
            Click(46, 33, 1, 210);
            Click(116, 359, 1, 310);
            TypeText("{ENTER}{ENTER}**************************************{ENTER}", true);
            Pause(310);
            Click(585, 481, 1, 310);
            SendKeys.SendWait("{HOME}");
            Pause(510);
            TypeText("SITES");
            Pause(410);
            SendKeys.SendWait("{ENTER}");
            Pause(310);
        }

        private static void WriteStatus(Excel.Worksheet worksheet, int row, string status)
        {
            worksheet.Cells[row, GetColumnNumber(StatusColumn)] = status;
        }

        private static string GetCellText(object[,] values, int row, string columnLetter, int firstUsedColumn, int colCount)
        {
            int sheetColumn = GetColumnNumber(columnLetter);
            int relativeColumn = sheetColumn - firstUsedColumn + 1;

            if (relativeColumn < 1 || relativeColumn > colCount)
            {
                return string.Empty;
            }

            object value = values[row, relativeColumn];
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

        private static void TypeText(string text, bool containsSendKeysTokens = false)
        {
            if (containsSendKeysTokens)
            {
                SendKeys.SendWait(text ?? string.Empty);
                return;
            }

            SendKeys.SendWait(EscapeSendKeys(text ?? string.Empty));
        }

        private static string EscapeSendKeys(string text)
        {
            StringBuilder builder = new StringBuilder(text.Length);

            foreach (char character in text)
            {
                switch (character)
                {
                    case '+':
                    case '^':
                    case '%':
                    case '~':
                    case '(':
                    case ')':
                    case '{':
                    case '}':
                    case '[':
                    case ']':
                        builder.Append('{').Append(character).Append('}');
                        break;
                    default:
                        builder.Append(character);
                        break;
                }
            }

            return builder.ToString();
        }

        private static void Click(int x, int y, int clickCount, int delayAfter)
        {
            SetCursorPos(x, y);
            Pause(110);

            for (int i = 0; i < clickCount; i++)
            {
                mouse_event(MouseEventLeftDown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseEventLeftUp, 0, 0, 0, UIntPtr.Zero);
                Pause(60);
            }

            Pause(delayAfter);
        }

        private static void Pause(int milliseconds)
        {
            Thread.Sleep(milliseconds);
        }

        private static bool ActivateWindowByTitleContains(string titlePart)
        {
            IntPtr foundWindow = IntPtr.Zero;

            EnumWindows((hWnd, lParam) =>
            {
                StringBuilder sb = new StringBuilder(256);
                GetWindowText(hWnd, sb, sb.Capacity);
                if (sb.ToString().IndexOf(titlePart, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    foundWindow = hWnd;
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            if (foundWindow == IntPtr.Zero)
            {
                return false;
            }

            SetForegroundWindow(foundWindow);
            Pause(700);
            return true;
        }

        private sealed class CaqhRecord
        {
            public string CarrierNumber { get; set; }
            public string CobRowEffDate { get; set; }
            public string CobRowEndDate { get; set; }
            public string SubscriberNumber { get; set; }
            public string SiteCode { get; set; }
            public string Comment { get; set; }
            public string SourceSystemGroupCode { get; set; }
            public string CobIndicator { get; set; }

            public bool IsEmpty
            {
                get
                {
                    return string.IsNullOrWhiteSpace(CarrierNumber)
                        && string.IsNullOrWhiteSpace(CobRowEffDate)
                        && string.IsNullOrWhiteSpace(CobRowEndDate)
                        && string.IsNullOrWhiteSpace(SubscriberNumber)
                        && string.IsNullOrWhiteSpace(SiteCode)
                        && string.IsNullOrWhiteSpace(Comment)
                        && string.IsNullOrWhiteSpace(SourceSystemGroupCode)
                        && string.IsNullOrWhiteSpace(CobIndicator);
                }
            }
        }
    }
}

