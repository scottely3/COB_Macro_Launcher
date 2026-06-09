using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace LibEndUserToolLauncher
{
    /// <summary>
    /// Information about an end user tool 
    /// </summary>
    public class EutTool
    {
        #region Properties
        /// <summary>
        /// Friendly name for tool
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Description of what tool does
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Version string of tool
        /// </summary>
        public string Version { get; set; }
        /// <summary>
        /// Whether or not the tool is disabled in the menu
        /// </summary>
        public bool Enabled { get; set; }
        /// <summary>
        /// Share drive location of executable
        /// </summary>
        private string location;
        public string Location 
        {
            get
            {
                return location;
            }
            set
            {
                location = value;
                if (Action == null) Action = () => { Process.Start(value); };
            } 
        }
        /// <summary>
        /// Key code (ex. Keys.Escape, Keys.A, Keys.Enter) used to activate tool
        /// </summary>
        public Keys Key { get; set; }
        /// <summary>
        /// Modifier key(s) used to activate tool.
        /// </summary>
        public int Modifier { get; set; }
        /// <summary>
        /// Unique number programmatically generated to denote the hotkey used to activate this tool
        /// </summary>
        public int ActivationHotKeyId { get; set; }
        private string activationString;
        /// <summary>
        /// Human-readable string representing key combo used to activate tool. If an ActivationString is explicitly set, it will be used to determine the Key and Modifier properties to set for this object.
        /// </summary>
        public string ActivationString
        {
            get
            {
                //if a raw ActivationString has not been set,
                //create one based on the provided Key and Modifier Properties
                if (activationString == null)
                {
                    string[] ActivationStrings = new string[5];
                    int ActivationCountdown = Modifier;
                    if (ActivationCountdown >= 8)
                    {
                        ActivationCountdown -= 8;
                        ActivationStrings[0] = "Win + ";
                    }
                    if (ActivationCountdown >= 4)
                    {
                        ActivationCountdown -= 4;
                        ActivationStrings[3] = "Shift + ";
                    }
                    if (ActivationCountdown >= 2)
                    {
                        ActivationCountdown -= 2;
                        ActivationStrings[1] = "Ctrl + ";
                    }
                    if (ActivationCountdown >= 1)
                    {
                        ActivationCountdown -= 1;
                        ActivationStrings[2] = "Alt + ";
                    }
                    ActivationStrings[4] = Key.ToString();
                    activationString = String.Join("", ActivationStrings);
                }
                else
                    ActivationString = activationString;
                return activationString;
            }
            set
            {
                //reformat the passed ActivationString into friendly visual format
                //and use it to determine the Modifier and Key Properties
                string NewActivationString = ActivationStringReformat(value);

                //find the main key used to activate
                Enum.TryParse(NewActivationString.Split('+').Last(), out Keys MyKey);

                //find any modifier keys used to activate
                int MyModifier = 0;
                if (NewActivationString.Contains("Alt +"))
                    MyModifier += 1;
                if (NewActivationString.Contains("Ctrl + "))
                    MyModifier += 2;
                if (NewActivationString.Contains("Shift + "))
                    MyModifier += 4;
                if (NewActivationString.Contains("Win + "))
                    MyModifier += 8;

                //if both main key and modifiers found, set the appropriate properties
                if(MyModifier >= 0 && MyKey != 0)
                {
                    activationString = NewActivationString;
                    Modifier = MyModifier;
                    Key = MyKey;
                }
            }
        }
        private string activationStringFriendly;
        /// <summary>
        /// Display version of ActivationString that may be more user-friendly than what is shown by default (ex. "Win + Keypad /" instead of "Win + Divide" or "CTRL + Keypad 1" instead of "CTRL + D1"). If an ActivationStringFriendly is explicitly set, that will be what is displayed to users instead of whatever the ActivationString, Key, or Modifier properties are, but it will not be used to determine the Key and Modifier properties to set for this object (only the ActivationString will be used for that). If an ActivationStringFriendly is not explicitly set, the raw ActivationString will be used as the display version in the menu.
        /// </summary>
        public string ActivationStringFriendly
        {
            get
            {
                if (activationStringFriendly != null)
                {
                    return activationStringFriendly;
                }
                else
                {
                    return ActivationString;
                }
            }
            set { activationStringFriendly = value; }
        }
        public delegate void ActionDelegate();
        private ActionDelegate action;
        /// <summary>
        /// Method to be called when tool is activated
        /// </summary>
        public ActionDelegate Action
        {
            get
            {
                if (Enabled) return action;
                else return null;
            }
            set
            {
                action = value;
            }
        }
        #endregion
        #region Static members
        /// <summary>
        /// Modifier keys used in activation key combos
        /// </summary>
        public static class Modifiers
        {
            public static readonly int None = 0;
            public static readonly int Alt = 1;
            public static readonly int Ctrl = 2;
            public static readonly int Shift = 4;
            public static readonly int Win = 8;
        }
        #endregion
        #region Methods
        /// <summary>
        ///reformat ActivationString into friendly visual format
        ///and use it to determine the Modifier and Key Properties
        /// </summary>
        /// <param name="ActivationString"></param>
        /// <returns></returns>
        public string ActivationStringReformat(string ActivationString)
        {
            string ReturnString = string.Empty;
            if (ActivationStringIsValid(ActivationString))
            {
                string[] ActivationStrings = new string[5];

                //create list of all keys in ActivationString
                List<string> AllKeys = ActivationString.ToLower().Split('+').ToList();

                //trim whitespace
                for (int i = 0; i < AllKeys.Count; i++)
                    AllKeys[i] = AllKeys[i].Trim();

                //find the main key used to activate
                Enum.TryParse(CultureInfo.CurrentCulture.TextInfo.ToTitleCase(AllKeys.Last()), out Keys MyKey);
                ActivationStrings[4] = MyKey.ToString();

                //find any modifier keys used to activate
                if (AllKeys.Contains("alt"))
                    ActivationStrings[2] = "Alt + ";
                if (AllKeys.Contains("ctrl"))
                    ActivationStrings[1] = "Ctrl + ";
                if (AllKeys.Contains("shift"))
                    ActivationStrings[3] = "Shift + ";
                if (AllKeys.Contains("win"))
                    ActivationStrings[0] = "Win + ";

                //set activationString to string that conforms to format
                ReturnString = String.Join("", ActivationStrings);
            }

            return ReturnString;
        }
        public bool ActivationStringIsValid(string TestString = null)
        {
            /// Regex pattern used to validate ActivationStrings
            /// Must be separated by "+", with or witout spaces, can include alt/shift/win/ctrl in any order, 
            /// the last item has to be a alphanumeric string of some sort other than alt/shift/win/ctrl
            string ActivationStringPattern = @"^(?:(?:alt|ctrl|shift|win) *\+ *){0,3}(?:(?!.*(?:\balt\b|\bctrl\b|\bshift\b|\bwin\b))[a-z0-9])+$";

            TestString = TestString ?? activationString;
            return new Regex(ActivationStringPattern, RegexOptions.IgnoreCase).IsMatch(TestString);
        }
        #endregion
        public EutTool() 
        {
            Modifier = Modifiers.None;
            Enabled = true;
            ActivationStringFriendly = null;
        }
    }
}
