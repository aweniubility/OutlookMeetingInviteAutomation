using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using OutlookCLI;
using System.Diagnostics;

namespace OutlookGUI
{
    public partial class GUI : Form
    {

        static string TAG = "MainOutlookGUI";
        static Office.COMAddIn addIn = null;

        public GUI()
        {
            InitializeComponent();

            OpenOutlookIfNotRunning();

            string methodTag = "GUI";

            try
            {
                object addInName = "OutlookAddIn";
                Outlook.Application outlookApp = new Outlook.Application();
                addIn = outlookApp.COMAddIns.Item(ref addInName);
                outlookApp = null;
            }
            catch (COMException ex)
            {

                LogWriter.WriteException(TAG, methodTag, ex);
                Environment.Exit(-1);
            }


        }

        public static void OpenOutlookIfNotRunning()
        {
            string methodTag = "OpenOutlookIfNotRunning";

            Outlook.Application outlookObj = null;

            try
            {
                outlookObj = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
            }
            catch
            {
                LogWriter.WriteInfo(TAG, methodTag, "Outlook was not running, starting outlook.exe");
                Process.Start("outlook.exe");

                System.Threading.Thread.Sleep(10000);
                LogWriter.WriteInfo(TAG, methodTag, "Outlook started");
            }

            return;
        }

        private void otherRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (otherRadio.Checked)
            {
                otherTextBox.Enabled = true;
            }
            else
            {
                otherTextBox.Text = "";
                otherTextBox.Enabled = false;
            }
        }

        private void sendButton_Click(object sender, EventArgs e)
        {

            string methodTag = "sendButton_Click";

            string recurrenceType = null;
            string subject = null;
            string otherCalendar = null;
            string endDate = null;
            string startDate = null;
            string duration = null;
            string[] recipients;


            if (subjectTextBox.Text == "")
            {
                LogWriter.WriteWarning(TAG, methodTag, "Subject not set, may cause issues. Using default of blank title");
                subject = null;
            }
            else
            {
                subject = subjectTextBox.Text;
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
            }

            if (startTextBox.Text == "")
            {
                LogWriter.WriteWarning(TAG, methodTag, "Start date not specified, using default of the next hour");
                startDate = null;
            }
            else
            {
                startDate = startTextBox.Text;
                LogWriter.WriteInfo(TAG, methodTag, "Start date is set to: " + startDate);
            }

            if (durationTextBox.Text == "")
            {
                LogWriter.WriteInfo(TAG, methodTag, "Duration not set, using default of 30 minutes");
                duration = null;
            }
            else
            {
                duration = durationTextBox.Text;
                LogWriter.WriteInfo(TAG, methodTag, "Duration set to: " + duration);
            }




            if (otherRadio.Checked)
            {
                otherCalendar = otherTextBox.Text;
                LogWriter.WriteInfo(TAG, methodTag, "Other calendar is set to: " + otherCalendar);
            }

            if (dailyRadio.Checked)
            {
                LogWriter.WriteInfo(TAG, methodTag, "Recurrence type set to Daily");
                recurrenceType = "Daily";
            }
            else if (weeklyRadio.Checked)
            {
                LogWriter.WriteInfo(TAG, methodTag, "Recurrence type set to Weekly");
                recurrenceType = "Weekly";
            }
            else if (monthlyRadio.Checked)
            {
                LogWriter.WriteInfo(TAG, methodTag, "Recurrence type set to Monthly");
                recurrenceType = "Monthly";
            }
            else if (yearlyRadio.Checked)
            {
                LogWriter.WriteInfo(TAG, methodTag, "Recurrence type set to Yearly");
                recurrenceType = "Yearly";
            }

            if (!noneRadio.Checked)
            {
                LogWriter.WriteInfo(TAG, methodTag, "End date set to: " + endDate);
                endDate = endDateTextBox.Text;
            }


            if(recipientTextBox.Text == "")
            {
                recipients = new string[] { };
            }
            else
            {
                recipients = recipientTextBox.Text.Split('\n');
                LogWriter.WriteInfo(TAG, methodTag, "Recipeints added: " + string.Join(" ", recipients));
            }
            



            if (otherRadio.Checked)
            {
                int iterate = Convert.ToInt32(iterationTextBox.Text);
                for (int i = 0; i < iterate; i++)
                {
                    addIn.Object.AddEventToOtherCalendar(subject, startDate, recurrenceType, endDate, duration, otherCalendar, recipients);
                }
            }
            else
            {
                int iterate = Convert.ToInt32(iterationTextBox.Text);
                for (int i = 0; i < iterate; i++)
                {
                    addIn.Object.AddEventToMainCalendar(subject, startDate, recurrenceType, endDate, duration, recipients);
                }
            }
        }

        private void updateButton_Click(object sender, EventArgs e)
        {

            string methodTag = "updateButton_Click";

            string subject = null;
            string eventDate = null;
            string updatedTitle = null;
            string updatedStart = null;
            string updatedDuration = null;
            string otherCalendar = null;
            string[] recipients;

            if (otherRadio.Checked)
            {
                otherCalendar = otherTextBox.Text;
                LogWriter.WriteInfo(TAG, methodTag, "Other calendar set to: " + otherCalendar);
            }

            if (subjectEventTextBox.Text == "")
            {
                LogWriter.WriteWarning(TAG, methodTag, "Subject required to update an event");
                MessageBox.Show("Subject required to update an event", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                subject = subjectEventTextBox.Text;
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
            }

            if(dateUpdateTextBox.Text == "")
            {
                eventDate = null;
                LogWriter.WriteInfo(TAG, methodTag, "No date set, finding single instance");
            }
            else
            {
                eventDate = dateUpdateTextBox.Text;
                LogWriter.WriteInfo(TAG, methodTag, "Looking for event on date: " + eventDate);
            }

            if(subjectUpdateTextBox.Text == "")
            {
                updatedTitle = null;
                LogWriter.WriteInfo(TAG, methodTag, "Keeping original title");
            }
            else
            {
                updatedTitle = subjectUpdateTextBox.Text;
                LogWriter.WriteInfo(TAG, methodTag, "Updated title set to: " + updatedTitle);
            }


            if(startDateTextBox.Text == "")
            {
                updatedStart = null;
                LogWriter.WriteInfo(TAG, methodTag, "Start time remaining the same");
            }
            else
            {
                updatedStart = startDateTextBox.Text;
                LogWriter.WriteInfo(TAG, methodTag, "Setting start date to: " + updatedStart);
            }

            if(durationUpdateTextBox.Text == "")
            {
                updatedDuration = null;
                LogWriter.WriteInfo(TAG, methodTag, "Duration remaining the same");
            }
            else
            {
                updatedDuration = durationUpdateTextBox.Text;
                LogWriter.WriteInfo(TAG, methodTag, "Duration set to: " + updatedDuration);
            }

            if (recipientTextBox.Text == "")
            {
                recipients = new string[] { };
            }
            else
            {
                recipients = recipientTextBox.Text.Split('\n');
                LogWriter.WriteInfo(TAG, methodTag, "Recipeints added: " + string.Join(" ", recipients));
            }



            if (otherRadio.Checked)
            {
                addIn.Object.UpdateOtherCalendarEvent(subject, eventDate, updatedTitle, updatedStart, updatedDuration, otherCalendar, recipients);
            }
            else
            {
                addIn.Object.UpdateMainCalendarEvent(subject, eventDate, updatedTitle, updatedStart, updatedDuration, recipients);
            }


        }

        private void noneRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (noneRadio.Checked)
            {
                endDateTextBox.Enabled = false;
            }
        }

        private void dailyRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (dailyRadio.Checked)
            {
                endDateTextBox.Enabled = true;
            }
        }

        private void weeklyRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (weeklyRadio.Checked)
            {
                endDateTextBox.Enabled = true;
            }
        }

        private void monthlyRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (monthlyRadio.Checked)
            {
                endDateTextBox.Enabled = true;
            }
        }

        private void yearlyRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (yearlyRadio.Checked)
            {
                endDateTextBox.Enabled = true;
            }
        }

        private void otherUpdateRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (otherUpdateRadio.Checked)
            {
                otherUpdateTextBox.Enabled = true;
            }
            else
            {
                otherUpdateTextBox.Text = "";
                otherUpdateTextBox.Enabled = false;
            }
        }

        private void otherRadioRemove_CheckedChanged(object sender, EventArgs e)
        {
            if (otherRadioRemove.Checked)
            {
                otherTextBoxRemove.Enabled = true;
            }
            else
            {
                otherTextBoxRemove.Text = "";
                otherTextBoxRemove.Enabled = false;
            }

        }

        private void removeButton_Click(object sender, EventArgs e)
        {
            string methodTag = "removeButton_Click";

            string otherCalendar = null;
            string subject = null;
            string eventDate = null;

            if(otherRadioRemove.Checked == true)
            {
                otherCalendar = otherTextBoxRemove.Text;
                LogWriter.WriteInfo(TAG, methodTag, "Other calendar set to: " + otherCalendar);
            }

            if(subjectTextBoxRemove.Text == "")
            {
                LogWriter.WriteWarning(TAG, methodTag, "Subject required to remove an event");
                MessageBox.Show("Subject required to remove an event", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                subject = subjectTextBoxRemove.Text;
                LogWriter.WriteInfo(TAG, methodTag, "Removing the event with subject: " + subject);
            }
            

            if(dateTextBoxRemove.Text != "")
            {
                eventDate = dateTextBoxRemove.Text;
                LogWriter.WriteInfo(TAG, methodTag, "Event date set to: " + eventDate);
            }

            if (otherRadioRemove.Checked)
            {
                addIn.Object.RemoveEventOtherCalendar(subject, eventDate, otherCalendar);
            }
            else
            {
                addIn.Object.RemoveEventMainCalendar(subject, eventDate);
            }



        }

    }
}
