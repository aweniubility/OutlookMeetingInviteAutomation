using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookCLI;

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
    /// <summary>
    /// Class containing common functions for manipulating outlook
    /// </summary>
    class CommonLibrary
    {

        static string TAG = "CommonLibrary";

        /// <summary>
        /// Adds the attributes to an AppointmentItem
        /// </summary>
        /// <param name="calEvent">The event of a specific calendar</param>
        /// <param name="subject">Title of the event</param>
        /// <param name="startDate">Start date of the event</param>
        /// <param name="recurrenceType">Recurrence type</param>
        /// <param name="endDate">End date for the recurrence</param>
        /// <param name="duration">Duration of the event</param>
        /// <param name="recipients">Recipients for the event</param>
        /// <returns>returns the filled in event so that it can be saved and sent from outlook</returns>
        internal static Outlook.AppointmentItem CreateEvent(Outlook.AppointmentItem calEvent, string subject, string startDate, string recurrenceType, string endDate,
            string duration, string[] recipients)
        {

            string methodTag = "CreateEvent";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);

            Outlook.RecurrencePattern recurrPatt = null;

            if (calEvent != null)
            {

                calEvent.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                calEvent.Sensitivity = Outlook.OlSensitivity.olNormal;

                if (subject != null)
                {
                    calEvent.Subject = subject;
                    LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
                }
                else
                {
                    LogWriter.WriteWarning(TAG, methodTag, "Subject is null, may cause failures when searching for it");
                }
                
                if(startDate != null)
                {
                    calEvent.Start = DateTime.Parse(startDate);
                    LogWriter.WriteInfo(TAG, methodTag, "Start date set to: " + startDate);
                }
                else
                {
                    LogWriter.WriteInfo(TAG, methodTag, "No start date provided, using default of the next closest hour");
                    LogWriter.WriteWarning(TAG, methodTag, "No start date provided, may cause failures when searching for it");
                }

                if (duration != null)
                {
                    calEvent.Duration = Convert.ToInt32(duration);
                    LogWriter.WriteInfo(TAG, methodTag, "Duration set to: " + duration);
                }
                else
                {
                    LogWriter.WriteInfo(TAG, methodTag, "No duration provided, using default of 30 minutes");
                }


                if (recurrenceType != null)
                {
                    recurrPatt = calEvent.GetRecurrencePattern();
                    if (endDate != null)
                    {
                        recurrPatt.PatternEndDate = DateTime.ParseExact(endDate, "M/d/yyyy", null);
                        LogWriter.WriteInfo(TAG, methodTag, "End date of recurrence set to: " + endDate);
                    }
                    else
                    {
                        LogWriter.WriteWarning(TAG, methodTag, "Recurrence date not set, infinite recurrence may cause problems when searching");
                    }

                    if (recurrenceType == "Daily")
                    {
                        recurrPatt.RecurrenceType = Outlook.OlRecurrenceType.olRecursDaily;
                    }
                    else if (recurrenceType == "Weekly")
                    {
                        recurrPatt.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly;

                    }
                    else if (recurrenceType == "Monthly")
                    {
                        recurrPatt.RecurrenceType = Outlook.OlRecurrenceType.olRecursMonthly;
                    }
                    else if (recurrenceType == "Yearly")
                    {
                        recurrPatt.RecurrenceType = Outlook.OlRecurrenceType.olRecursYearly;
                    }

                    LogWriter.WriteInfo(TAG, methodTag, "Recurrence type set to: " + recurrenceType);

                }

                foreach (String contact in recipients)
                {
                    calEvent.Recipients.Add(contact);
                    LogWriter.WriteInfo(TAG, methodTag, "Adding recipient: " + contact + " to meeting invite");
                }
            }
            else
            {
                LogWriter.WriteWarning(TAG, methodTag, "Calendar event is null, unable to continue with creating event");
                throw new NullReferenceException();
            }

            return calEvent;
        }

        /// <summary>
        /// Updates the attributes of a specific event
        /// </summary>
        /// <param name="calEvent">Event to be updated</param>
        /// <param name="eventDate">Date of the specific event to be updated M/d/yyyy hh:mm:ss tt</param>
        /// <param name="updatedTitle">Updated title if desired</param>
        /// <param name="updatedStartDate">Updated start date if desired</param>
        /// <param name="updatedDuration">Updated duration if desired</param>
        /// <param name="recipients">Additonal recipients if desired</param>
        /// <returns></returns>
        internal static Outlook.AppointmentItem UpdateEvent(Outlook.AppointmentItem calEvent, string eventDate, string updatedTitle, string updatedStartDate,
            string updatedDuration, string[] recipients)
        {

            string methodTag = "UpdateEvent";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);

            Outlook.AppointmentItem agendaMeeting = null;
            Outlook.RecurrencePattern recurrPatt = null;
            

            if (calEvent != null)
            {

                if (eventDate != null)
                {
                    LogWriter.WriteInfo(TAG, methodTag, "Obtaining specific event on date: " + eventDate + " with title: " + calEvent.Subject);
                    recurrPatt = calEvent.GetRecurrencePattern();
                    agendaMeeting = recurrPatt.GetOccurrence(DateTime.Parse(eventDate));
                }
                else
                {
                    LogWriter.WriteWarning(TAG, methodTag, "No event date specified, may cause issues finding event if it is a recurring event");
                    agendaMeeting = calEvent;
                }

                if (updatedTitle != null)
                {
                    LogWriter.WriteInfo(TAG, methodTag, "Updating the events title to: " + updatedTitle);
                    agendaMeeting.Subject = updatedTitle;
                }
                else
                {
                    LogWriter.WriteInfo(TAG, methodTag, "No updated subject specified, keeping subject the same");
                }

                if (updatedDuration != null)
                {
                    LogWriter.WriteInfo(TAG, methodTag, "Updating the events duration to: " + updatedDuration);
                    agendaMeeting.Duration = Convert.ToInt32(updatedDuration);
                }
                else
                {
                    LogWriter.WriteInfo(TAG, methodTag, "Updated duration not specified, keepign duration the same");
                }

                foreach (String contact in recipients)
                {
                    agendaMeeting.Recipients.Add(contact);
                    LogWriter.WriteInfo(TAG, methodTag, "Adding recipient: " + contact + " to meeting invite");
                }

                if (updatedStartDate != null)
                {
                    LogWriter.WriteInfo(TAG, methodTag, "Updating the events start date to: " + updatedStartDate);
                    agendaMeeting.Start = DateTime.Parse(updatedStartDate);
                }
                else
                {
                    LogWriter.WriteInfo(TAG, methodTag, "Updated start date not specified, keeping start date the same");
                }


                agendaMeeting.Save();
                LogWriter.WriteInfo(TAG, methodTag, "Event updated sucessfully");

            }
            else
            {
                LogWriter.WriteWarning(TAG, methodTag, "Calendar event is null, unable to continue with update event");
                throw new NullReferenceException();
            }

            return agendaMeeting;

        }

        /// <summary>
        /// Removes a specific event or a series of events
        /// </summary>
        /// <param name="objItems">Contains all the event items in a specific calendar/folder</param>
        /// <param name="subject">Subject of the event to delete</param>
        /// <param name="eventDate">If specified, this specific event is deleted instead of all events with subject</param>
        internal static void RemoveEvent(Outlook.Items objItems, string subject, string eventDate)
        {

            string methodTag = "RemoveEvent";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);

            Outlook.AppointmentItem agendaMeeting = null;

            if (eventDate == null)
            {

                objItems.Sort("[Subject]");
                objItems.IncludeRecurrences = true;

                LogWriter.WriteInfo(TAG, methodTag, "Attempting to find event with subject: " + subject);
                agendaMeeting = objItems.Find("[Subject]=" + subject);

                if (agendaMeeting == null)
                {
                    LogWriter.WriteWarning(TAG, methodTag, "No event found with subject: " + subject);
                }
                else
                {
                    LogWriter.WriteInfo(TAG, methodTag, "Removing all events with subject: " + subject);
                    do
                    {

                        agendaMeeting.Delete();

                        agendaMeeting = objItems.FindNext();

                        LogWriter.WriteInfo(TAG, methodTag, "Event found and deleted, finding next");


                    } while (agendaMeeting != null);

                    LogWriter.WriteInfo(TAG, methodTag, "All events with subject: " + subject + " found and deleted");
                }
            }
            else
            {
                LogWriter.WriteInfo(TAG, methodTag, "Finding event with subject" + subject + " and date: " + eventDate);
                agendaMeeting = objItems.Find("[Subject]=" + subject);

                if(agendaMeeting == null)
                {
                    LogWriter.WriteWarning(TAG, methodTag, "No event found with subject: " + subject);
                }
                else
                {
                    Outlook.RecurrencePattern recurrPatt = agendaMeeting.GetRecurrencePattern();
                    agendaMeeting = recurrPatt.GetOccurrence(DateTime.Parse(eventDate));

                    agendaMeeting.Delete();

                    LogWriter.WriteInfo(TAG, methodTag, "Event found and deleted");
                }

                
            }
        }
    }
}
