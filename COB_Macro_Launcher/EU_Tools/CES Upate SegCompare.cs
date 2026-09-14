using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using COB_Macro_Launcher.UI;

namespace COB_Macro_Launcher.EU_Tools
{
    internal class CES_UNET_SegCompare
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
        private const string CsvPath = "\\\\nas01592pn\\Fester_Common\\COB\\COB Misc\\New CES Commercial COB - KUCopy.csv";

        public static void Start()
        {
            TopMostDialogs.MessageBox(
                Program.LauncherOwner,
                "WARNING: This launcher reads the CES UNET SegCompare CSV and uses fixed emulator window coordinates.",
                "CES UNET SegCompare",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);

            if (!File.Exists(CsvPath))
            {
                TopMostDialogs.MessageBox(Program.LauncherOwner, "CSV file not found:\n" + CsvPath, "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                return;
            }

            List<SegCompareRecord> records = LoadRecords(CsvPath);
            if (records.Count == 0)
            {
                TopMostDialogs.MessageBox(Program.LauncherOwner, "No CSV records were found to process.", "CES UNET SegCompare", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                return;
            }

            int processed = 0;

            try
            {
                foreach (SegCompareRecord record in records)
                {
                    if (record.IsEmpty)
                    {
                        continue;
                    }

                    if (!ActivateWindowByTitleContains("A - E"))
                    {
                        throw new InvalidOperationException("Emulator window not found.");
                    }

                    ProcessRecord(record);
                    processed++;
                }

                TopMostDialogs.MessageBox(Program.LauncherOwner, "Processed " + processed + " record(s).", "CES UNET SegCompare", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
            catch (Exception ex)
            {
                TopMostDialogs.MessageBox(Program.LauncherOwner, "CES UNET SegCompare failed.\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        private static List<SegCompareRecord> LoadRecords(string csvPath)
        {
            List<SegCompareRecord> records = new List<SegCompareRecord>();

            foreach (string line in File.ReadLines(csvPath))
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                List<string> fields = ParseCsvLine(line);
                SegCompareRecord record = BuildRecord(fields);
                NormalizeRecord(record);
                records.Add(record);
            }

            return records;
        }

        private static SegCompareRecord BuildRecord(List<string> fields)
        {
            return new SegCompareRecord
            {
                GroupId = GetField(fields, 0),
                SubscriberNumber = GetField(fields, 1),
                MemberType = GetField(fields, 2),
                DependentValue = GetField(fields, 3),
                Field5 = GetField(fields, 4),
                Date1 = GetField(fields, 5),
                Date2 = GetField(fields, 6),
                Field8 = GetField(fields, 7),
                Field9 = GetField(fields, 8),
                Field10 = GetField(fields, 9)
            };
        }

        private static void NormalizeRecord(SegCompareRecord record)
        {
            record.GroupId = PadLeft(record.GroupId, 7, '0');
            record.SubscriberNumber = PadLeft(record.SubscriberNumber, 11, '0');
            record.Date1 = NormalizeDateField(record.Date1);
            record.Date2 = NormalizeDateField(record.Date2);
        }

        private static void ProcessRecord(SegCompareRecord record)
        {
            if (string.Equals(record.MemberType, "E", StringComparison.OrdinalIgnoreCase))
            {
                ProcessEmployee(record);
            }
            else
            {
                ProcessSpouseDependent(record);
            }

            RunClipboardValidation(record);
        }

        private static void ProcessEmployee(SegCompareRecord record)
        {
            Pause(1000);
            TypeText("T");
            Pause(100);
            SendKeys.SendWait("{ENTER}");
            TypeText("c");
            Pause(110);
            TypeText("h");
            SendKeys.SendWait("{TAB}");
            Pause(110);
            TypeText(record.GroupId);
            SendKeys.SendWait("{UP}{TAB}{TAB}");
            Pause(110);
            TypeText(record.SubscriberNumber);
            SendKeys.SendWait("{ENTER}");
            Pause(400);
            TypeText("U");
            Pause(310);
            Click(202, 570, 1, 310);
            SendKeys.SendWait("{TAB}{TAB}");
            Pause(310);
            TypeText(record.Date1);
            Pause(310);
            TypeText(record.Date2);
            Pause(310);
            SendKeys.SendWait("{ENTER}");
            Pause(310);
        }

        private static void ProcessSpouseDependent(SegCompareRecord record)
        {
            Pause(1000);
            TypeText("T");
            Pause(100);
            SendKeys.SendWait("{ENTER}");
            TypeText("c");
            Pause(10);
            TypeText("h");
            Pause(20);
            TypeText(record.MemberType);
            Pause(20);
            TypeText(record.GroupId);
            SendKeys.SendWait("{UP}{TAB}{TAB}");
            Pause(110);
            TypeText(record.SubscriberNumber);
            Pause(110);
            SendKeys.SendWait("{TAB}{TAB}{TAB}{TAB}");
            Pause(110);
            TypeText(record.DependentValue);
            Pause(110);
            SendKeys.SendWait("{ENTER}");
            Pause(400);
            TypeText("U");
            Pause(310);
            Click(202, 570, 1, 310);
            SendKeys.SendWait("{TAB}{TAB}");
            Pause(310);
            TypeText(record.Date1);
            Pause(310);
            TypeText(record.Date2);
            Pause(310);
        }

        private static void RunClipboardValidation(SegCompareRecord record)
        {
            string clipboardText = CaptureStableClipboardText();
            string clipValue20 = SafeSubstring(clipboardText, 1200, 10);
            string clipValue21 = SafeSubstring(clipboardText, 160, 40);

            SendKeys.SendWait("{F3}");
            if (!ActivateWindowByTitleContains("cobT Validation"))
            {
                throw new InvalidOperationException("cobT Validation window not found.");
            }

            Pause(150);
            TypeText(record.GroupId + "  ");
            Pause(75);
            TypeText(record.SubscriberNumber + " ");
            Pause(100);
            TypeText(clipValue20 + " ");
            Pause(200);
            TypeText(clipValue21);
            SendKeys.SendWait("{ENTER}{ENTER}?{ENTER}{ENTER}");
            Pause(410);
        }

        private static string CaptureStableClipboardText()
        {
            string lastValue = string.Empty;

            for (int i = 0; i < 50; i++)
            {
                ActivateWindowByTitleContains("A - E");
                SendKeys.SendWait("^(%{INSERT})");
                Pause(110);

                string currentValue = ReadClipboardWithRetry(50, 50);
                if (!string.IsNullOrEmpty(currentValue))
                {
                    lastValue = currentValue;
                    break;
                }
            }

            for (int i = 0; i < 100; i++)
            {
                string currentValue = ReadClipboardWithRetry(1, 10);
                if (string.Equals(currentValue, lastValue, StringComparison.OrdinalIgnoreCase))
                {
                    return currentValue;
                }

                if (!string.IsNullOrEmpty(currentValue))
                {
                    lastValue = currentValue;
                }
            }

            return lastValue;
        }

        private static string ReadClipboardWithRetry(int attempts, int delayMilliseconds)
        {
            for (int i = 0; i < attempts; i++)
            {
                try
                {
                    if (Clipboard.ContainsText())
                    {
                        return Clipboard.GetText();
                    }
                }
                catch
                {
                }

                Pause(delayMilliseconds);
            }

            return string.Empty;
        }

        private static List<string> ParseCsvLine(string line)
        {
            List<string> fields = new List<string>();
            StringBuilder current = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }

                    continue;
                }

                if (c == ',' && !inQuotes)
                {
                    fields.Add(current.ToString());
                    current.Clear();
                    continue;
                }

                current.Append(c);
            }

            fields.Add(current.ToString());
            return fields;
        }

        private static string GetField(List<string> fields, int index)
        {
            if (fields == null || index < 0 || index >= fields.Count)
            {
                return string.Empty;
            }

            return (fields[index] ?? string.Empty).Trim();
        }

        private static string PadLeft(string value, int totalLength, char paddingCharacter)
        {
            value = value ?? string.Empty;
            return value.Length >= totalLength ? value : value.PadLeft(totalLength, paddingCharacter);
        }

        private static string NormalizeDateField(string value)
        {
            value = (value ?? string.Empty).Trim();
            if (value.Length == 0)
            {
                return new string(' ', 8);
            }

            return value.Length >= 8 ? value : value.PadLeft(8, '0');
        }

        private static string SafeSubstring(string value, int oneBasedStart, int count)
        {
            value = value ?? string.Empty;
            if (oneBasedStart <= 0 || count <= 0)
            {
                return string.Empty;
            }

            int zeroBasedStart = oneBasedStart - 1;
            if (zeroBasedStart >= value.Length)
            {
                return string.Empty;
            }

            int length = Math.Min(count, value.Length - zeroBasedStart);
            return value.Substring(zeroBasedStart, length).TrimEnd();
        }

        private static void TypeText(string text)
        {
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

        private sealed class SegCompareRecord
        {
            public string GroupId { get; set; }
            public string SubscriberNumber { get; set; }
            public string MemberType { get; set; }
            public string DependentValue { get; set; }
            public string Field5 { get; set; }
            public string Date1 { get; set; }
            public string Date2 { get; set; }
            public string Field8 { get; set; }
            public string Field9 { get; set; }
            public string Field10 { get; set; }

            public bool IsEmpty
            {
                get
                {
                    return string.IsNullOrWhiteSpace(GroupId)
                        && string.IsNullOrWhiteSpace(SubscriberNumber)
                        && string.IsNullOrWhiteSpace(MemberType)
                        && string.IsNullOrWhiteSpace(DependentValue)
                        && string.IsNullOrWhiteSpace(Field5)
                        && string.IsNullOrWhiteSpace(Date1)
                        && string.IsNullOrWhiteSpace(Date2)
                        && string.IsNullOrWhiteSpace(Field8)
                        && string.IsNullOrWhiteSpace(Field9)
                        && string.IsNullOrWhiteSpace(Field10);
                }
            }
        }
    }
}
