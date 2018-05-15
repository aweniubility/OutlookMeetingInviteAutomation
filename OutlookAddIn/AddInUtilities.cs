using System;
using System.Runtime.InteropServices;
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
    /*
         * Exposing VSTO to other solutions: https://msdn.microsoft.com/en-us/library/bb608621.aspx
         * Interface that is exposed to the outside solution. For more information see above link.
         * Interface includes operations that can be done to manipluate outlook calendar objects
         */
    [ComVisible(true)]
    public interface IAddInUtilities
    {
        void AddEventToMainCalendar(string subject, string startDate, string recurrenceType, string endDate, string duration, string[] recipients);
        void AddEventToOtherCalendar(string subject, string startDate, string recurrenceType, string endDate, string duration, string otherCalendar, string[] recipients);
        void UpdateMainCalendarEvent(string subject, string eventDate, string updatedTitle, string updatedStartDate, string updatedDuration, string[] recipients);
        void UpdateOtherCalendarEvent(string subject, string eventDate, string updatedTitle, string updatedStartDate, string updatedDuration,
            string otherCalendar, string[] recipients);
        void RemoveEventMainCalendar(string subject, string eventDate);
        void RemoveEventOtherCalendar(string subject, string eventDate, string otherCalendar);
    }

    /*
     * Exposing VSTO to other solutions: https://msdn.microsoft.com/en-us/library/bb608621.aspx
     * Implements the above exposed interface and is derived from the StandardOleMarshalObject to ensure
     * out-of-process clients are able to interact with the exposed interface   
     */
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddInUtilities : StandardOleMarshalObject, IAddInUtilities
    {

        static string TAG = "AddInUtilities";

        /// <summary>
        /// Adds an event to the main calendar of the open Outlook Application.
        /// Utilizes the common methods found in the common library class
        /// </summary>
        /// <param name="subject">The subject of the event</param>
        /// <param name="startDate">The start date of the event in M/d/yyyy hh:mm:ss tt format</param>
        /// <param name="recurrenceType">What kind of recurrence it is, if one. None, daily, weekly, etc</param>
        /// <param name="endDate">If there is a recurrence when does the recurrence end.</param>
        /// <param name="duration">Duration of the event</param>
        /// <param name="recipients">Who to invite to the event.</param>
        public void AddEventToMainCalendar(string subject, string startDate, string recurrenceType, string endDate,
            string duration, string[] recipients)
        {
            string methodTag = "AddEventToMainCalendar";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);

            Outlook.Application app = null;
            Outlook.AppointmentItem agendaMeeting = null;

            app = new Outlook.Application();
            agendaMeeting = (Outlook.AppointmentItem)app.CreateItem(Outlook.OlItemType.olAppointmentItem);

            LogWriter.WriteInfo(TAG, methodTag, "Adding details to event item");
            try
            {
                agendaMeeting = CommonLibrary.CreateEvent(agendaMeeting, subject, startDate, recurrenceType, endDate, duration, recipients);
            }
            catch(Exception e)
            {
                throw e;
            }
            

            agendaMeeting.Send();
            LogWriter.WriteInfo(TAG, methodTag, subject + " created and sent");

            return;
        }

        /// <summary>
        /// Updates a specific event (a single event or one instance of a recurrence) 
        /// that is currently in the main calendar of the open Outlook Application
        /// Utilizes the common methods found in the common library class
        /// </summary>
        /// <param name="subject">The subject of the event to edit</param>
        /// <param name="eventDate">The date of the specific event M/d/yyyy hh:mm:ss tt</param>
        /// <param name="updatedTitle">An updated title if desired</param>
        /// <param name="updatedStartDate">An updated start date if desired</param>
        /// <param name="updatedDuration">An updated duration if desired</param>
        /// <param name="recipients">New recipients to add</param>
        public void UpdateMainCalendarEvent(string subject, string eventDate, string updatedTitle, string updatedStartDate, string updatedDuration, string[] recipients)
        {
            string methodTag = "UpdateMainCalendarEvent";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);

            Outlook.Application app = null;
            Outlook.AppointmentItem agendaMeeting = null;
            Outlook.Items objItems = null;
            Outlook.MAPIFolder objFolder = null;

            app = new Outlook.Application();
            Outlook.NameSpace NS = app.GetNamespace("MAPI");
            agendaMeeting = (Outlook.AppointmentItem)app.CreateItem(Outlook.OlItemType.olAppointmentItem);

            objFolder = NS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            objItems = objFolder.Items;

            LogWriter.WriteInfo(TAG, methodTag, "Attempting to find event with subject: " + subject);

            agendaMeeting = objItems.Find("[Subject]=" + subject);

            if(agendaMeeting == null)
            {
                LogWriter.WriteWarning(TAG, methodTag, "Failed to find event with subject: " + subject);
                throw new NullReferenceException();
            }

            LogWriter.WriteInfo(TAG, methodTag, "Found event with subject: " + subject + " and updating details");

            agendaMeeting = CommonLibrary.UpdateEvent(agendaMeeting, eventDate, updatedTitle, updatedStartDate, updatedDuration, recipients);

            LogWriter.WriteInfo(TAG, methodTag, "Sending event");
            agendaMeeting.Send();
            LogWriter.WriteInfo(TAG, methodTag, subject + " sucessfully sent");

            return;
        }

        /// <summary>
        /// Removes an event or a series of events based on the title 
        /// that is currently in the main calendar of the open Outlook Application
        /// If the eventDate is specified it will only remove that specific instance
        /// If it is not specified will remove all events with that title
        /// Utilizes the common methods found in the common library class
        /// </summary>
        /// <param name="subject">Title of the event to remove</param>
        /// <param name="eventDate">The date of the specific event M/d/yyyy hh:mm:ss tt. Not required</param>
        public void RemoveEventMainCalendar(string subject, string eventDate)
        {

            string methodTag = "RemoveEventMainCalendar";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);

            Outlook.Application app = null;
            Outlook.Items objItems = null;
            Outlook.MAPIFolder objFolder = null;

            app = new Outlook.Application();
            Outlook.NameSpace NS = app.GetNamespace("MAPI");

            objFolder = NS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            objItems = objFolder.Items;

            LogWriter.WriteInfo(TAG, methodTag, "Attempting to remove: " + subject);
            CommonLibrary.RemoveEvent(objItems, subject, eventDate);
            LogWriter.WriteInfo(TAG, methodTag, subject + " successfully removed");

            return;


        }

        /// <summary>
        /// Adds an event to another calendar in the outlook application. Calendar must be present in
        /// the shared folder of outlook. Creates an event object in that folder and then utilizes
        /// common library to add the attributes
        /// </summary>
        /// <param name="subject">Title of the event to add</param>
        /// <param name="startDate">The date of the specific event M/d/yyyy hh:mm:ss tt</param>
        /// <param name="recurrenceType">Recurrence type: daily, weekly, monthly, yearly</param>
        /// <param name="endDate">End date for the recurrence</param>
        /// <param name="duration">Duration of the event</param>
        /// <param name="otherCalendar">Name of the other calendar</param>
        /// <param name="recipients">Array of recipeints to add to the event</param>
        public void AddEventToOtherCalendar(string subject, string startDate, string recurrenceType, string endDate, string duration, string otherCalendar, string[] recipients)
        {
            string methodTag = "AddEventToOtherCalendar";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);

            Outlook.Application app = null;
            Outlook.AppointmentItem agendaMeeting = null;
            Outlook.NameSpace NS = null;
            Outlook.MAPIFolder objFolder = null;
            Outlook.MailItem objTemp = null;
            Outlook.Recipient objRecip = null;
            Outlook.Items objItems = null;

            app = new Outlook.Application();
            NS = app.GetNamespace("MAPI");
            objTemp = app.CreateItem(Outlook.OlItemType.olMailItem);
            objRecip = objTemp.Recipients.Add(otherCalendar);
            objTemp = null;
            objRecip.Resolve();

            if (objRecip.Resolved)
            {
                objFolder = NS.GetSharedDefaultFolder(objRecip, Outlook.OlDefaultFolders.olFolderCalendar);
                objItems = objFolder.Items;
                agendaMeeting = objItems.Add();

                LogWriter.WriteInfo(TAG, methodTag, "Adding details to event item");
                agendaMeeting = CommonLibrary.CreateEvent(agendaMeeting, subject, startDate, recurrenceType, endDate, duration, recipients);

                LogWriter.WriteInfo(TAG, methodTag, "Sending event");
                agendaMeeting.Send();
                LogWriter.WriteInfo(TAG, methodTag, subject + " sucessfully sent");

            }
            else
            {
                LogWriter.WriteWarning(TAG, methodTag, "Recipient object was not sucessfully resolved");
                throw new NullReferenceException();
            }

            return;

        }

        /// <summary>
        /// Updates a specific instance of an event from another calendar in Outlook
        /// The other shared calendar must be opened in outlook before running.
        /// Utilizes the updateEvent function in the common library.
        /// </summary>
        /// <param name="subject">Title of the Event to update</param>
        /// <param name="eventDate">The date of the specific event M/d/yyyy hh:mm:ss tt</param>
        /// <param name="updatedTitle">Updated title. not required</param>
        /// <param name="updatedStartDate">Updated start date. not required</param>
        /// <param name="updatedDuration">updated duration. not required</param>
        /// <param name="otherCalendar">Name of the calendar the event is located in</param>
        /// <param name="recipients">Recipients to add to the updated event</param>
        public void UpdateOtherCalendarEvent(string subject, string eventDate, string updatedTitle, string updatedStartDate, string updatedDuration,
            string otherCalendar, string[] recipients)
        {

            string methodTag = "UpdateOtherCalendarEvent";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);

            Outlook.Application app = null;
            Outlook.AppointmentItem agendaMeeting = null;
            Outlook.NameSpace NS = null;
            Outlook.MAPIFolder objFolder = null;
            Outlook.MailItem objTemp = null;
            Outlook.Recipient objRecip = null;
            Outlook.Items objItems = null;

            app = new Outlook.Application();
            NS = app.GetNamespace("MAPI");
            objTemp = app.CreateItem(Outlook.OlItemType.olMailItem);
            objRecip = objTemp.Recipients.Add(otherCalendar);
            objTemp = null;

            LogWriter.WriteInfo(TAG, methodTag, "Attempting to resolve recipient object for: " + otherCalendar);
            objRecip.Resolve();

            if (objRecip.Resolved)
            {
                objFolder = NS.GetSharedDefaultFolder(objRecip, Outlook.OlDefaultFolders.olFolderCalendar);
                objItems = objFolder.Items;
                agendaMeeting = objItems.Find("[Subject]=" + subject);

                agendaMeeting = CommonLibrary.UpdateEvent(agendaMeeting, eventDate, updatedTitle, updatedStartDate, updatedDuration, recipients);

                LogWriter.WriteInfo(TAG, methodTag, "Sending event");
                agendaMeeting.Send();
                LogWriter.WriteInfo(TAG, methodTag, subject + " sucessfully sent");

            }
            else
            {
                LogWriter.WriteWarning(TAG, methodTag, "Recipient object was not sucessfully resolved");
                throw new NullReferenceException();
            }


            return;

        }

        /// <summary>
        /// Remove a specific event or a series of event from a shared calendar in outlook
        /// Requires the shared calendar to be opened in outlook.
        /// If the eventDate is specificed removes a specific instance. If not removes all events with that subject
        /// </summary>
        /// <param name="subject">Subject of the event to remove</param>
        /// <param name="eventDate">Date of the specific event to remove. If null, deletes all event with the title. M/d/yyyy hh:mm:ss tt</param>
        /// <param name="otherCalendar">Name of the calendar the event exists in</param>
        public void RemoveEventOtherCalendar(string subject, string eventDate, string otherCalendar)
        {
            string methodTag = "RemoveEventOtherCalendar";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);

            Outlook.Application app = null;
            Outlook.NameSpace NS = null;
            Outlook.MAPIFolder objFolder = null;
            Outlook.MailItem objTemp = null;
            Outlook.Recipient objRecip = null;
            Outlook.Items objItems = null;

            app = new Outlook.Application();
            NS = app.GetNamespace("MAPI");
            objTemp = app.CreateItem(Outlook.OlItemType.olMailItem);
            objRecip = objTemp.Recipients.Add(otherCalendar);
            objTemp = null;

            LogWriter.WriteInfo(TAG, methodTag, "Attempting to resolve recipient object for: " + otherCalendar);
            objRecip.Resolve();

            if (objRecip.Resolved)
            {
                objFolder = NS.GetSharedDefaultFolder(objRecip, Outlook.OlDefaultFolders.olFolderCalendar);
                objItems = objFolder.Items;

                LogWriter.WriteInfo(TAG, methodTag, "Attempting to remove: " + subject);
                CommonLibrary.RemoveEvent(objItems, subject, eventDate);
                LogWriter.WriteInfo(TAG, methodTag, subject + " successfully removed");
            }
            else
            {
                LogWriter.WriteWarning(TAG, methodTag, "Recipient object was not sucessfully resolved");
                throw new NullReferenceException();
            }

            return;
            
        }
    }
}
