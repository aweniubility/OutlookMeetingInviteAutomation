using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

/*
 * Prerequisites:
 * Outlook 2013 or 2016 installed
 * Visual studios installed
 * OutlookAutomationSuite solution built. This should install CalendarAutomationAddIn and complie required executables
 * 
 * 
 * Resources:
 * Outlook Aplication Interface: https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.application.aspx
 * Outlook Object Model Overview: https://msdn.microsoft.com/en-us/library/ms268893.aspx
 * Working with Calendar Items: https://msdn.microsoft.com/en-us/library/bb386291.aspx
 * Exposing VSTO to other solutions: https://msdn.microsoft.com/en-us/library/bb608621.aspx
 */
namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private AddInUtilities utilities;

        /// <summary>
        /// Creates an onject of AddInUtilities and returns it. Allowing outside
        /// programs to access its methods and attributes.
        /// </summary>
        /// <returns>utilities - AddInUtilities object</returns>
        protected override object RequestComAddInAutomationService()
        {
            if (utilities == null)
                utilities = new AddInUtilities();

            return utilities;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
