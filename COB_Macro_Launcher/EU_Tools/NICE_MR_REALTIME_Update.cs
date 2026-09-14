using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using COB_Macro_Launcher.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace COB_Macro_Launcher.EU_Tools
{
    internal class NICE_MR_REALTIME_Update
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
                "WARNING: This launcher uses the active Excel sheet and fixed NICE screen coordinates to run the MR Realtime update flow.",
                "NICE MR REALTIME Update",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);

            string dataRow = TopMostDialogs.InputBox("What row does your data begin on?", "Data_Row");

            int startRow;
            if (!int.TryParse(dataRow, out startRow) || startRow <= 0)
            {
                TopMostDialogs.MessageBox(Program.LauncherOwner, "Please enter a valid Excel row number.", "Invalid Row", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                return;
            }

            Excel.Application excelApp = null;
            Excel.Worksheet activeSheet = null;
            Excel.Range usedRange = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                activeSheet = (Excel.Worksheet)excelApp.ActiveSheet;
                workbook = (Excel.Workbook)activeSheet.Parent;
                usedRange = activeSheet.UsedRange;

                if (usedRange == null || usedRange.Rows == null || usedRange.Rows.Count == 0)
                {
                    return;
                }

                int firstUsedRow = usedRange.Row;
                int lastUsedRow = firstUsedRow + usedRange.Rows.Count - 1;
                int firstRowToProcess = Math.Max(startRow, firstUsedRow);

                if (firstRowToProcess > lastUsedRow)
                {
                    TopMostDialogs.MessageBox(Program.LauncherOwner, "The starting row is outside the worksheet's used range.", "Invalid Row", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    return;
                }

                for (int worksheetRow = firstRowToProcess; worksheetRow <= lastUsedRow; worksheetRow++)
                {
                    MrRealtimeRecord record = BuildRecord(activeSheet, worksheetRow);
                    if (record.IsEmpty)
                    {
                        continue;
                    }

                    string status = ProcessRecord(record);
                    WriteStatus(activeSheet, worksheetRow, status);
                    workbook.Save();
                }
            }
            catch (Exception ex)
            {
                TopMostDialogs.MessageBox(Program.LauncherOwner, "Could not access Excel active sheet.\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        private static MrRealtimeRecord BuildRecord(Excel.Worksheet worksheet, int row)
        {
            return new MrRealtimeRecord
            {
                CompanyId = GetCellText(worksheet, row, CompanyIdColumn),
                MemberId = GetCellText(worksheet, row, MemberIdColumn),
                LastName = GetCellText(worksheet, row, LastNameColumn),
                FirstName = GetCellText(worksheet, row, FirstNameColumn),
                SubscriberLast = GetCellText(worksheet, row, SubscriberLastColumn),
                SubscriberFirst = GetCellText(worksheet, row, SubscriberFirstColumn),
                Dob = GetCellText(worksheet, row, DobColumn),
                Relationship = GetCellText(worksheet, row, RelationshipColumn),
                InfoSource = GetCellText(worksheet, row, InfoSourceColumn),
                TypeOfPlan = GetCellText(worksheet, row, TypeOfPlanColumn),
                OrderOfBenefits = GetCellText(worksheet, row, OrderOfBenefitsColumn),
                CobLastResearched = GetCellText(worksheet, row, CobLastResearchedColumn),
                CobIndicator = GetCellText(worksheet, row, CobIndicatorColumn),
                OicName = GetCellText(worksheet, row, OicNameColumn),
                OicPolicy = GetCellText(worksheet, row, OicPolicyColumn),
                OicEffDate = GetCellText(worksheet, row, OicEffDateColumn),
                OicTermDate = GetCellText(worksheet, row, OicTermDateColumn),
                Comments = GetCellText(worksheet, row, CommentsColumn),
            };
        }

        private static string ProcessRecord(MrRealtimeRecord record)
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

            return TryUpdateExistingOic(record);
        }

        private static bool SearchMember(MrRealtimeRecord record)
        {
            Click(582, 329, 2, 2000);
            Click(568, 254, 2, 1000);
            SendKeys.SendWait("{TAB}");
            Pause(110);
            SendKeys.SendWait("{TAB}");
            //SendKeys.SendWait("{DELETE}");
            Click(764, 333, 1, 210);
            Pause(510);
            Click(743, 377, 1, 210);
            //PasteText(record.CompanyId);
            //PasteText(record.CompanyId);
            Pause(2000);

            Click(278, 358, 3, 710);
            SendKeys.SendWait("{DELETE}");
            Pause(510);
            PasteText("00" + record.MemberId);
            //PasteText(record.MemberId);
            Pause(410);

            Click(1067, 333, 1, 1500);
            return true;
        }

        private static bool TryOpenRelationship(string relationship)
        {
            int[] rowY = { 441, 461, 481, 501, 521, 541, 561, 581, 601 };

            foreach (int y in rowY)
            {
                string value = SelectAndCopy(282, y);
                if (string.Equals(Normalize(value), Normalize(relationship), StringComparison.OrdinalIgnoreCase))
                {
                    Pause(500);
                    SendKeys.SendWait("{TAB}");
                    Pause(310);
                    SendKeys.SendWait("{ENTER}");
                    Pause(1000);
                    Clipboard.Clear();
                    return true;
                }
            }

            return false;
        }

        private static bool OpenOicSection()
        {
            Click(1009, 440, 1, 2000);
            return true;
        }

        private static string GetPageTitle()
        {
            return SelectAndCopy(457, 159);
        }

        private static bool CreateNewOic(MrRealtimeRecord record)
        {
            Click(847, 123, 1, 1300);
            //Click(997, 123, 2, 2300);
            SendKeys.SendWait("{TAB}{TAB}");
            Pause(210);
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
            Pause(210);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            PasteText(record.Dob);
            Pause(410);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            Click(1222, 557, 2, 2210);
            Click(1080, 379, 1, 2210);
            //PasteText(record.InfoSource);
            Pause(310);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            PasteText(record.TypeOfPlan);
            Pause(310);
            SendKeys.SendWait("{ENTER}");
            Pause(1000);

            Click(1066, 598, 3, 210);
            PasteText(record.CobLastResearched);
            Pause(710);
            Click(398, 628, 1, 2000);
            //click(1022, 561, 1, 210);
            Click(287, 552, 1, 2210);

            //PasteText(record.OrderOfBenefits);
            Pause(2410);
            SendKeys.SendWait("{TAB}");
            Pause(110);
            PasteText(record.CobIndicator);
            Pause(110);
            SendKeys.SendWait("{TAB}");
            Pause(110);

            Click(192, 445, 1, 1000);
            Click(217, 509, 1, 310);
            PasteText(record.OicName);
            Pause(1000);
            Click(403, 509, 1, 2000);

            Click(157, 232, 1, 1000);
            Click(132, 232, 1, 610);
            Click(598, 264, 1, 310);
            Click(598, 284, 1, 210);
            Click(598, 304, 1, 210);
            Click(598, 324, 1, 210);
            Click(598, 344, 1, 1000);

            Click(215, 661, 1, 1000);
            PasteText(record.OicPolicy);
            Pause(1000);
            Click(296, 693, 1, 1000);
            PasteText(record.OicEffDate);
            Pause(310);
            SendKeys.SendWait("{TAB}");
            Pause(310);
            PasteText(record.OicTermDate);
            Pause(1000);

            Click(412, 440, 3, 1000);
            EnterComments(record.Comments);
            SaveAndCloseRecord();
            return true;
        }

        private static string TryUpdateExistingOic(MrRealtimeRecord record)
        {
            int[] carrierRowY = { 453, 473, 493, 513, 533 };

            foreach (int y in carrierRowY)
            {
                string carrier = SelectAndCopy(443, y);
                if (!string.Equals(Normalize(carrier), Normalize(record.OicName), StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string rowMarker = SelectAndCopy(778, y);
                if (string.Equals(Normalize(rowMarker), "/", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                OpenExistingOicRow(y);

                if (!PrepareExistingOic(record))
                {
                    TryCloseCurrentDialog();
                    return "Unable to load record";
                }

                string currentEffDate = SelectAndCopy(295, 682);
                string currentTermDate = SelectAndCopy(849, 682);

                if (string.Equals(Normalize(currentEffDate), Normalize(record.OicEffDate), StringComparison.OrdinalIgnoreCase)
                    && string.Equals(Normalize(currentTermDate), Normalize(record.OicTermDate), StringComparison.OrdinalIgnoreCase))
                {
                    TryCloseCurrentDialog();
                    return "Already loaded to Nice";
                }

                UpdateExistingDatesAndComments(record);
                return VerifyMainOrComplete();
            }

            return "No Match on OI Carrier";
        }

        private static void OpenExistingOicRow(int y)
        {
            Click(65, y, 1, 1500);
            SendKeys.SendWait("{TAB}");
            Pause(1250);
            SendKeys.SendWait("{ENTER}");
            Pause(3500);
            Click(997, 121, 3, 3000);
        }

        private static bool PrepareExistingOic(MrRealtimeRecord record)
        {
            Click(538, 465, 1, 1000);
            SendKeys.SendWait("{TAB}{TAB}");
            Pause(100);
            SendKeys.SendWait("{TAB}{TAB}");
            Pause(100);
            SendKeys.SendWait("{TAB}{TAB}");
            Pause(100);
            Click(1222, 557, 1, 2210);
            Click(1080, 379, 1, 2210);
            //PasteText(record.InfoSource);
            Pause(810);
            SendKeys.SendWait("{TAB}");
            Pause(100);
            SendKeys.SendWait("{TAB}");
            Pause(250);
            PasteText(record.CobLastResearched);
            Pause(1000);
            //Click(387, 620, 1, 1000);
            Click(398, 628, 1, 2000);
            //click(1022, 561, 1, 210);
            Click(287, 552, 1, 2210);
            //PasteText(record.OrderOfBenefits);
            Pause(1310);
            SendKeys.SendWait("{TAB}");
            Pause(110);
            PasteText(record.CobIndicator);
            Pause(250);
            Click(181, 435, 1, 500);
            return true;
        }

        private static void UpdateExistingDatesAndComments(MrRealtimeRecord record)
        {
            Click(215, 661, 2, 1000);
            PasteText(record.OicPolicy);
            Pause(310);

            Click(295, 682, 3, 250);
            SendKeys.SendWait("{DELETE}");
            Pause(1000);
            PasteText(record.OicEffDate);
            Pause(1250);
            SendKeys.SendWait("{TAB}");
            Pause(250);
            SendKeys.SendWait("{DELETE}");
            Pause(510);
            PasteText(record.OicTermDate);
            Pause(1000);

            Click(412, 440, 3, 1500);
            EnterComments(record.Comments);
            SaveAndCloseRecord();
        }

        private static void EnterComments(string comments)
        {
            Click(584, 116, 1, 1310);
            SendKeys.SendWait("{TAB}{TAB}{TAB}{TAB}{TAB}");
            Pause(310);
            SendKeys.SendWait("{TAB}{TAB}{TAB}{TAB}{TAB}");
            Pause(310);
            PasteText(comments);
            Pause(410);
        }

        private static void SaveAndCloseRecord()
        {
            Click(1032, 224, 2, 500);
            SendKeys.SendWait("{PGUP}");
            Pause(1000);
            Click(1077, 121, 1, 1510);
            SendKeys.SendWait("{ENTER}");
            Pause(1000);
            Click(1224, 121, 1, 1510);
            Pause(1000);
        }

        private static string VerifyMainOrComplete()
        {
            string mainIndicator = SelectAndCopy(581, 221);
            if (mainIndicator.IndexOf("Main", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "SUCCESSFUL";
            }

            TryCloseCurrentDialog();
            return "Unable to load record";
        }

        private static void TryCloseCurrentDialog()
        {
            Click(1224, 121, 1, 410);
            SendKeys.SendWait("{ENTER}");
            Pause(500);
        }

        private static void WriteStatus(Excel.Worksheet worksheet, int row, string status)
        {
            worksheet.Cells[row, GetColumnNumber(StatusColumn)] = status;
        }

        private static string GetCellText(Excel.Worksheet worksheet, int row, string columnLetter)
        {
            Excel.Range cell = null;

            try
            {
                cell = (Excel.Range)worksheet.Cells[row, GetColumnNumber(columnLetter)];
                object value = cell == null ? null : cell.Text;
                return value == null ? string.Empty : value.ToString().Trim();
            }
            finally
            {
                if (cell != null)
                {
                    Marshal.ReleaseComObject(cell);
                }
            }
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

        private sealed class MrRealtimeRecord
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
