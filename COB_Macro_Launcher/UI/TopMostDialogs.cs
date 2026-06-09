using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using System.Windows.Forms;

namespace COB_Macro_Launcher.UI
{
    internal static class TopMostDialogs
    {
        internal static DialogResult MessageBox(IWin32Window owner, string text)
        {
            return MessageBox(owner, text, "", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1);
        }

        internal static DialogResult MessageBox(
            IWin32Window owner,
            string text,
            string caption,
            MessageBoxButtons buttons,
            MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton,
            MessageBoxOptions options = 0)
        {
            // Even when we have an owner, Edge/CPA can be TopMost/foreground and Windows will place the
            // MessageBox behind it. Use a temporary TopMost form as the dialog owner to reliably bring
            // the dialog to the foreground.
            using (var top = CreateTopMostOwner())
            {
                try { top.Show(); top.Activate(); } catch { /* ignore */ }

                // Attempt to activate the provided owner as well; harmless if it fails.
                if (owner is Form f)
                {
                    try { f.Activate(); } catch { /* ignore */ }
                }

                return System.Windows.Forms.MessageBox.Show(top, text, caption, buttons, icon, defaultButton, options);
            }
        }

        internal static string InputBox(string prompt, string title, string defaultResponse = "")
        {
            return InputBox(null, prompt, title, defaultResponse);
        }

        internal static string InputBox(IWin32Window owner, string prompt, string title, string defaultResponse = "")
        {
            // Microsoft.VisualBasic.Interaction.InputBox has no owner parameter.
            // We can still force topmost behavior by creating a temporary TopMost form
            // and showing it before invoking InputBox.
            using (var top = CreateTopMostOwner())
            {
                try { top.Show(); top.Activate(); } catch { /* ignore */ }

                if (owner is Form f)
                {
                    try { f.Activate(); } catch { /* ignore */ }
                }

                return Interaction.InputBox(prompt, title, defaultResponse, -1, -1);
            }
        }

        private static Form CreateTopMostOwner()
        {
            var f = new Form
            {
                TopMost = true,
                ShowInTaskbar = false,
                StartPosition = FormStartPosition.Manual,
                Size = new System.Drawing.Size(1, 1),
                Location = new System.Drawing.Point(-32000, -32000)
            };

            // Ensure handle exists for ownership/z-order.
            f.Show();
            f.Hide();
            return f;
        }
    }
}
