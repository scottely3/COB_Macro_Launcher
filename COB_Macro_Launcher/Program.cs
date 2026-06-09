using COB_Macro_Launcher.EU_Tools;
using LibEndUserToolLauncher;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace COB_Macro_Launcher
{
    public static class Program
    {
        public static IWin32Window LauncherOwner { get; private set; }
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main()
        {
            LibEndUserToolLauncher.FormMenu ToolMenu = new LibEndUserToolLauncher.FormMenu()
            {
                 //whether or not the Form itself is actually displayed when
                 //the program executes and begins listening to keystrokes
                 ShowForm = true,
                 //whether or not the console is displayed when this Form is run
                 ShowConsole = false,
                 //title text that appears on the window and the system tray icon/notification
                 //Title = Macro_FileName,
                 //add some tools whose Action method resides in other Class files in this same project
                 Tools = new List<EutTool>()
                 {
                    // new EutTool()
                    //{
                    //    Name = "E2E_NicePHS",
                    //    Description = "member update to nice phs",
                    //    ActivationString = "CTRL+ALT+Z",
                    //    Modifier = EutTool.Modifiers.Ctrl|EutTool.Modifiers.Alt|EutTool.Modifiers.Shift,
                    //    Key = Keys.D3,
                    //    Action = E2E_NicePHS.Start
                    //},
                    new EutTool()
                    {
                        Name = "E2E_NicePHS",
                        Description = "member update to nice phs",
                        ActivationString = "CTRL+ALT+Z",
                        //Modifier = EutTool.Modifiers.Ctrl|EutTool.Modifiers.Alt|EutTool.Modifiers.Shift,
                        //Key = Keys.D3,
                        Action = E2E_NicePHS.Start
                    }
                 }
            };
            LauncherOwner = ToolMenu;
            ToolMenu.Run();
        }
    }
    internal sealed class Win32WindowWrapper : IWin32Window
    {
        public Win32WindowWrapper(IntPtr handle) { Handle = handle; }
        public IntPtr Handle { get; }
    }
}
