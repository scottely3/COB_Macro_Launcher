using System;
using System.Windows.Forms;
using COB_Macro_Launcher.EU_Tools;
using LibEndUserToolLauncher;

namespace COB_Macro_Launcher
{
    internal static class Program
    {
        public static IWin32Window LauncherOwner { get; private set; }

        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            FormMenu launcher = new FormMenu
            {
                Title = "COB Macro Launcher",
                ShowForm = true,
                ShowConsole = false,
                AllowShorcutChanges = true
            };

            LauncherOwner = launcher;

            launcher.AddTool(new EutTool
            {
                Name = "CES UNET SegCompare",
                ActivationString = "CTRL+ALT+1",
                Description = "Run CES UNET SegCompare flow.",
                Action = CES_UNET_SegCompare.Start,
                Enabled = true
            });

            launcher.AddTool(new EutTool
            {
                Name = "COSMOS Update for CAQH",
                ActivationString = "CTRL+ALT+2",
                Description = "Run COSMOS update for CAQH data.",
                Action = COSMOS_Update_for_CAQH.Start,
                Enabled = true
            });

            launcher.AddTool(new EutTool
            {
                Name = "COSMOS Update for SegComp",
                ActivationString = "CTRL+ALT+3",
                Description = "Run COSMOS update for SegComp data.",
                Action = COSMOS_Update_for_SegComp.Start,
                Enabled = true
            });

            launcher.AddTool(new EutTool
            {
                Name = "CSP FACETS MED",
                ActivationString = "CTRL+ALT+4",
                Description = "Run CSP FACETS MED flow.",
                Action = CSP_FACETS_MED.Start,
                Enabled = true
            });

            launcher.AddTool(new EutTool
            {
                Name = "E2E NicePHS",
                ActivationString = "CTRL+ALT+5",
                Description = "Run E2E NicePHS flow.",
                Action = E2E_NicePHS.Start,
                Enabled = true
            });

            launcher.AddTool(new EutTool
            {
                Name = "NICE Hospice Update",
                ActivationString = "CTRL+ALT+6",
                Description = "Run NICE Hospice update flow.",
                Action = NICE_Hospice_Update.Start,
                Enabled = true
            });

            launcher.AddTool(new EutTool
            {
                Name = "NICE MR REALTIME Update",
                ActivationString = "CTRL+ALT+7",
                Description = "Run NICE MR REALTIME update flow.",
                Action = NICE_MR_REALTIME_Update.Start,
                Enabled = true
            });

            launcher.Run();
        }
    }
}

