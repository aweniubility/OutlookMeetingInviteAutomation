using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookCLI;
using System.Diagnostics;

namespace OutlookGUI
{
    static class OutlookGUI
    {

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new GUI());
            }
            catch
            {
                MessageBox.Show("Error has occurred, check logs in C:\\Users\\Public", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
    }
}
