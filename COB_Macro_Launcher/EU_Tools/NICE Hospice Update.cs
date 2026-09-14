using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using COB_Macro_Launcher.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace COB_Macro_Launcher.EU_Tools
{
    internal class NICE_Hospice_Update
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

        private const string CompanyIdColumn = "A";
        private const string MemberIdColumn = "B";
        private const string LastNameColumn = "C";
        private const string FirstNameColumn = "D";
        private const string SubscriberLastColumn = "E";
        private const string SubscriberFirstColumn = "F";
        private const string DobColumn = "G";
        private const string RelationshipColumn = "H";
        private const string InfoSourceColumn = "I";
        private const string TypeOfPlanColumn = "J";
        private const string OrderOfBenefitsColumn = "K";
        private const string CobLastResearchedColumn = "L";
        private const string CobIndicatorColumn = "M";
        private const string OicNameColumn = "N";
        private const string OicPolicyColumn = "O";
        private const string OicEffDateColumn = "P";
        private const string OicTermDateColumn = "Q";
        private const string CommentsColumn = "R";
        private const string StatusColumn = "S";

        public static void Start()
        {
            TopMostDialogs.MessageBox(
                Program.LauncherOwner,
                "WARNING: This launcher uses Excel data and screen coordinates to load Hospice OI information into NICE COB Portal.",
                "NICE Hospice Update",
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
                    HospiceRecord record = BuildRecord(values, row, firstUsedColumn, colCount);

                    if (record.IsEmpty)
                    {
                        continue;
                    }

                    string status = ProcessRecord(record);
                    WriteStatus(activeSheet, worksheetRow, status);
                }

                ((Excel.Workbook)activeSheet.Parent).Save();
            }
            catch (Exception ex)
            {
                TopMostDialogs.MessageBox(Program.LauncherOwner, "Could not access Excel active sheet.\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        private static HospiceRecord BuildRecord(object[,] values, int row, int firstUsedColumn, int colCount)
        {
            return new HospiceRecord
            {
                CompanyId = GetCellText(values, row, CompanyIdColumn, firstUsedColumn, colCount),
                MemberId = GetCellText(values, row, MemberIdColumn, firstUsedColumn, colCount),
                LastName = GetCellText(values, row, LastNameColumn, firstUsedColumn, colCount),
                FirstName = GetCellText(values, row, FirstNameColumn, firstUsedColumn, colCount),
                SubscriberLast = GetCellText(values, row, SubscriberLastColumn, firstUsedColumn, colCount),
                SubscriberFirst = GetCellText(values, row, SubscriberFirstColumn, firstUsedColumn, colCount),
                Dob = GetCellText(values, row, DobColumn, firstUsedColumn, colCount),
                Relationship = GetCellText(values, row, RelationshipColumn, firstUsedColumn, colCount),
                InfoSource = GetCellText(values, row, InfoSourceColumn, firstUsedColumn, colCount),
                TypeOfPlan = GetCellText(values, row, TypeOfPlanColumn, firstUsedColumn, colCount),
                OrderOfBenefits = GetCellText(values, row, OrderOfBenefitsColumn, firstUsedColumn, colCount),
                CobLastResearched = GetCellText(values, row, CobLastResearchedColumn, firstUsedColumn, colCount),
                CobIndicator = GetCellText(values, row, CobIndicatorColumn, firstUsedColumn, colCount),
                OicName = GetCellText(values, row, OicNameColumn, firstUsedColumn, colCount),
                OicPolicy = GetCellText(values, row, OicPolicyColumn, firstUsedColumn, colCount),
                OicEffDate = GetCellText(values, row, OicEffDateColumn, firstUsedColumn, colCount),
                OicTermDate = GetCellText(values, row, OicTermDateColumn, firstUsedColumn, colCount),
                Comments = GetCellText(values, row, CommentsColumn, firstUsedColumn, colCount)
            };
        }

        private static string ProcessRecord(HospiceRecord record)
        {
            if (string.IsNullOrWhiteSpace(record.CompanyId) || string.IsNullOrWhiteSpace(record.MemberId))
            {
                return "Missing CompanyID or MemberID";
            }

            if (!ActivateWindowByTitleContains("COB Portal"))
            {
                return "COB Portal window not found";
            }

            if (!SearchMember(record))
            {
                return "Unable to load record";
            }

            if (!TryOpenRelationship(record.Relationship))
            {
                TryCloseCurrentDialog();
                return "Unable to load record";
            }

            if (!OpenOicSection())
            {
                return "Unable to load record";
            }

            string pageTitle = GetPageTitle();
            if (pageTitle.IndexOf("OIC", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return CreateNewOic(record) ? VerifyMainOrComplete() : "Unable to load record";
            }

            string existingStatus = TryUpdateExistingOic(record);
            if (!string.IsNullOrWhiteSpace(existingStatus))
            {
                return existingStatus;
            }

            return VerifyMainOrComplete();
        }

        private static bool SearchMember(HospiceRecord record)
        {
            Click(582, 329, 2, 700);
            Click(457, 259, 2, 2000);
            SendKeys.SendWait("{TAB}");
            Pause(110);
            SendKeys.SendWait("{TAB}");
            Pause(110);
            PasteText(record.CompanyId);
            Pause(410);

            Click(279, 366, 3, 1000);
            SendKeys.SendWait("{DELETE}");
            Pause(310);
            PasteText("00" + record.MemberId);
            Pause(210);

            Click(1067, 333, 1, 3210);
            return true;
        }

        private static bool TryOpenRelationship(string relationship)
        {
            int[] rowY = { 436, 456, 476, 496, 516, 536, 556, 576, 596 };

            foreach (int y in rowY)
            {
                string value = SelectAndCopy(320, y);
                if (string.Equals(Normalize(value), Normalize(relationship), StringComparison.OrdinalIgnoreCase))
                {
                    Pause(310);
                    SendKeys.SendWait("{TAB}");
                    Pause(210);
                    SendKeys.SendWait("{ENTER}");
                    Pause(710);
                    return true;
                }
            }

            return false;
        }

        private static bool OpenOicSection()
        {
            Click(1009, 440, 1, 1000);
            return true;
        }

        private static string GetPageTitle()
        {
            return SelectAndCopy(457, 159);
        }

        private static bool CreateNewOic(HospiceRecord record)
        {
            Click(809, 120, 1, 310);
            SendKeys.SendWait("{TAB}{TAB}");
            Pause(100);
            SendKeys.SendWait("{TAB}{TAB}");
            Pause(110);
            SendKeys.SendWait("{TAB}");
            Pause(110);
            SendKeys.SendWait("{TAB}{TAB}{TAB}");
            Pause(110);
            SendKeys.SendWait("{TAB}{TAB}");
            Pause(100);
            PasteText(record.SubscriberFirst);
            Pause(210);
            SendKeys.SendWait("{TAB}");
            Pause(110);
            SendKeys.SendWait("{TAB}");
            Pause(110);
            PasteText(record.SubscriberLast);
            Pause(310);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            PasteText(record.Dob);
            Pause(1000);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            PasteText("Ca");
            Pause(310);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            PasteText("M");
            Pause(810);
            SendKeys.SendWait("{ENTER}");
            Pause(1500);

            Click(1066, 598, 3, 210);
            PasteText(record.CobLastResearched);
            Pause(310);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            PasteText(record.OrderOfBenefits);
            Pause(510);
            SendKeys.SendWait("{TAB}");
            Pause(110);
            PasteText(record.CobIndicator);
            Pause(110);
            SendKeys.SendWait("{TAB}");

            Click(192, 445, 1, 1500);
            Click(215, 514, 1, 310);
            PasteText(record.OicName);
            Pause(710);
            Click(404, 513, 1, 3000);

            Click(157, 232, 1, 610);
            Click(598, 264, 1, 310);
            Click(597, 284, 1, 210);
            Click(598, 304, 1, 210);
            Click(598, 324, 1, 210);
            Click(598, 344, 1, 710);

            Click(215, 661, 1, 1510);
            PasteText(record.OicPolicy);
            Pause(710);
            Click(296, 683, 3, 1510);
            PasteText(record.OicEffDate);
            Pause(310);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            PasteText(record.OicTermDate);
            Pause(510);

            Click(298, 437, 1, 1510);
            Click(414, 464, 1, 1510);
            PasteText(record.SubscriberFirst);
            Pause(310);
            SendKeys.SendWait("{TAB}");
            Pause(110);
            SendKeys.SendWait("{TAB}");
            Pause(110);
            PasteText(record.SubscriberLast);
            Pause(310);
            SendKeys.SendWait("{TAB}");
            Pause(110);
            PasteText(record.OicPolicy);
            Pause(310);
            Click(409, 438, 1, 1000);
            Click(414, 443, 1, 310);

            EnterComments(record.Comments);
            SaveAndCloseRecord();
            return true;
        }

        private static string TryUpdateExistingOic(HospiceRecord record)
        {
            int[] carrierY = { 456, 476, 496, 516, 536 };

            for (int index = 0; index < carrierY.Length; index++)
            {
                string carrier = SelectAndCopy(index == 4 ? 488 : 459, carrierY[index]);
                if (!string.Equals(Normalize(carrier), Normalize(record.OicName), StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                OpenExistingOicRow(carrierY[index]);

                if (!PrepareExistingOic(record))
                {
                    TryCloseCurrentDialog();
                    return "Unable to load record";
                }

                string currentEffDate = SelectAndCopy(295, 682);
                string currentTermDate = SelectAndCopy(849, 682);

                if (string.Equals(Normalize(currentEffDate), Normalize(record.OicEffDate), StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(Normalize(currentTermDate), Normalize(record.OicTermDate), StringComparison.OrdinalIgnoreCase))
                {
                    TryCloseCurrentDialog();
                    return "Already loaded to Nice";
                }

                UpdateExistingDatesAndComments(record);
                return VerifyMainOrComplete();
            }

            TryCloseCurrentDialog();
            return "No Match on OI Carrier";
        }

        private static void OpenExistingOicRow(int y)
        {
            Click(65, y, 1, 250);
            SendKeys.SendWait("{TAB}");
            Pause(250);
            SendKeys.SendWait("{ENTER}");
            Pause(3000);
            Click(1000, 121, 1, 2000);
        }

        private static bool PrepareExistingOic(HospiceRecord record)
        {
            Click(538, 465, 1, 1000);
            SendKeys.SendWait("{TAB}{TAB}");
            Pause(310);
            SendKeys.SendWait("{TAB}{TAB}");
            Pause(310);
            SendKeys.SendWait("{TAB}{TAB}");
            Pause(310);
            PasteText(record.InfoSource);
            Pause(1000);
            SendKeys.SendWait("{TAB}{TAB}");
            Pause(310);
            PasteText(record.CobLastResearched);
            Pause(250);
            SendKeys.SendWait("{TAB}");
            Pause(250);
            PasteText(record.OrderOfBenefits);
            Pause(250);
            SendKeys.SendWait("{TAB}");
            Pause(150);
            PasteText(record.CobIndicator);
            Pause(250);
            Click(192, 445, 1, 500);
            return true;
        }

        private static void UpdateExistingDatesAndComments(HospiceRecord record)
        {
            Click(215, 661, 3, 1510);
            PasteText(record.OicPolicy);
            Pause(710);

            Click(295, 682, 3, 250);
            SendKeys.SendWait("{DELETE}");
            Pause(1000);
            PasteText(record.OicEffDate);
            Pause(250);
            SendKeys.SendWait("{TAB}");
            Pause(250);
            SendKeys.SendWait("{DELETE}");
            Pause(510);
            PasteText(record.OicTermDate);
            Pause(250);

            Click(414, 443, 1, 810);
            EnterComments(record.Comments);
            SaveAndCloseRecord();
        }

        private static void EnterComments(string comments)
        {
            Click(584, 121, 1, 310);
            SendKeys.SendWait("{TAB}{TAB}{TAB}{TAB}{TAB}");
            Pause(310);
            SendKeys.SendWait("{TAB}{TAB}{TAB}{TAB}{TAB}");
            Pause(310);
            PasteText(comments);
            Pause(410);
        }

        private static void SaveAndCloseRecord()
        {
            Click(1045, 150, 1, 310);
            SendKeys.SendWait("{PGUP}");
            Pause(310);
            Click(1078, 121, 1, 510);
            SendKeys.SendWait("{ENTER}");
            Pause(1000);
            Click(1226, 121, 1, 510);
            SendKeys.SendWait("{ENTER}");
            Pause(1000);
        }

        private static string VerifyMainOrComplete()
        {
            string mainIndicator = SelectAndCopy(581, 221);
            if (mainIndicator.IndexOf("Main", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "COMPLETE";
            }

            TryCloseCurrentDialog();
            return "Unable to load record";
        }

        private static void TryCloseCurrentDialog()
        {
            Click(1226, 121, 1, 410);
            SendKeys.SendWait("{ENTER}");
            Pause(500);
        }

        private static void WriteStatus(Excel.Worksheet worksheet, int row, string status)
        {
            int column = GetColumnNumber(StatusColumn);
            worksheet.Cells[row, column] = status;
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

        private static string SelectAndCopy(int x, int y)
        {
            Click(x, y, 3, 250);
            Clipboard.Clear();
            SendKeys.SendWait("^c");
            Pause(400);
            return Normalize(Clipboard.ContainsText() ? Clipboard.GetText() : string.Empty);
        }

        private static void PasteText(string text)
        {
            Clipboard.SetText(text ?? string.Empty);
            SendKeys.SendWait("^v");
        }

        private static string Normalize(string value)
        {
            return (value ?? string.Empty).Trim();
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

        private sealed class HospiceRecord
        {
            public string CompanyId { get; set; }
            public string MemberId { get; set; }
            public string LastName { get; set; }
            public string FirstName { get; set; }
            public string SubscriberLast { get; set; }
            public string SubscriberFirst { get; set; }
            public string Dob { get; set; }
            public string Relationship { get; set; }
            public string InfoSource { get; set; }
            public string TypeOfPlan { get; set; }
            public string OrderOfBenefits { get; set; }
            public string CobLastResearched { get; set; }
            public string CobIndicator { get; set; }
            public string OicName { get; set; }
            public string OicPolicy { get; set; }
            public string OicEffDate { get; set; }
            public string OicTermDate { get; set; }
            public string Comments { get; set; }

            public bool IsEmpty
            {
                get
                {
                    return string.IsNullOrWhiteSpace(CompanyId)
                        && string.IsNullOrWhiteSpace(MemberId)
                        && string.IsNullOrWhiteSpace(LastName)
                        && string.IsNullOrWhiteSpace(FirstName)
                        && string.IsNullOrWhiteSpace(SubscriberLast)
                        && string.IsNullOrWhiteSpace(SubscriberFirst)
                        && string.IsNullOrWhiteSpace(Dob)
                        && string.IsNullOrWhiteSpace(Relationship)
                        && string.IsNullOrWhiteSpace(InfoSource)
                        && string.IsNullOrWhiteSpace(TypeOfPlan)
                        && string.IsNullOrWhiteSpace(OrderOfBenefits)
                        && string.IsNullOrWhiteSpace(CobLastResearched)
                        && string.IsNullOrWhiteSpace(CobIndicator)
                        && string.IsNullOrWhiteSpace(OicName)
                        && string.IsNullOrWhiteSpace(OicPolicy)
                        && string.IsNullOrWhiteSpace(OicEffDate)
                        && string.IsNullOrWhiteSpace(OicTermDate)
                        && string.IsNullOrWhiteSpace(Comments);
                }
            }
        }
    }
}
