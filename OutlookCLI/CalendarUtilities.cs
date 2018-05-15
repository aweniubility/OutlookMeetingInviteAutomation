using System;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using Newtonsoft.Json;
using System.IO;


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
 * Generating C# Classes from JSON https://jsonutils.com/
 */

namespace OutlookCLI
{
    /// <summary>
    /// Command line application that hooks into a Outlook VSTO add in.
    /// Utilizies command line arguments to call functions and pass method arguments
    /// to the VSTO plugin to use in creating and manipulating Outlook Events.
    /// </summary>
    public class CalendarUtilities
    {

        static string TAG = "CalendarUtilities";

        public class SingleEventNoRecipients
        {

            [JsonProperty("Subject")]
            public string Subject { get; set; }

            [JsonProperty("StartDate")]
            public string StartDate { get; set; }

            [JsonProperty("Duration")]
            public string Duration { get; set; }

            [JsonProperty("Recipients")]
            public string Recipients { get; set; }

            [JsonProperty("RecurrenceType")]
            public string RecurrenceType { get; set; }

            [JsonProperty("EndDate")]
            public string EndDate { get; set; }

            [JsonProperty("OtherCalendar")]
            public string OtherCalendar { get; set; }
        }

        public class SingleEventWithRecipients
        {

            [JsonProperty("Subject")]
            public string Subject { get; set; }

            [JsonProperty("StartDate")]
            public string StartDate { get; set; }

            [JsonProperty("Duration")]
            public string Duration { get; set; }

            [JsonProperty("Recipients")]
            public string Recipients { get; set; }

            [JsonProperty("RecurrenceType")]
            public string RecurrenceType { get; set; }

            [JsonProperty("EndDate")]
            public string EndDate { get; set; }

            [JsonProperty("OtherCalendar")]
            public string OtherCalendar { get; set; }
        }

        public class RecurrenceNoRecipients
        {

            [JsonProperty("Subject")]
            public string Subject { get; set; }

            [JsonProperty("StartDate")]
            public string StartDate { get; set; }

            [JsonProperty("Duration")]
            public string Duration { get; set; }

            [JsonProperty("RecurrenceType")]
            public string RecurrenceType { get; set; }

            [JsonProperty("EndDate")]
            public string EndDate { get; set; }

            [JsonProperty("Recipients")]
            public string Recipients { get; set; }

            [JsonProperty("OtherCalendar")]
            public string OtherCalendar { get; set; }
        }

        public class RecurrenceWithRecipients
        {

            [JsonProperty("Subject")]
            public string Subject { get; set; }

            [JsonProperty("StartDate")]
            public string StartDate { get; set; }

            [JsonProperty("Duration")]
            public string Duration { get; set; }

            [JsonProperty("RecurrenceType")]
            public string RecurrenceType { get; set; }

            [JsonProperty("EndDate")]
            public string EndDate { get; set; }

            [JsonProperty("Recipients")]
            public string Recipients { get; set; }

            [JsonProperty("OtherCalendar")]
            public string OtherCalendar { get; set; }
        }

        public class UpdateSingleEventNoRecipients
        {

            [JsonProperty("Subject")]
            public string Subject { get; set; }

            [JsonProperty("UpdatedTitle")]
            public string UpdatedTitle { get; set; }

            [JsonProperty("EventDate")]
            public string EventDate { get; set; }

            [JsonProperty("UpdatedStartDate")]
            public string UpdatedStartDate { get; set; }

            [JsonProperty("UpdatedDuration")]
            public string UpdatedDuration { get; set; }

            [JsonProperty("Recipients")]
            public string Recipients { get; set; }

            [JsonProperty("OtherCalendar")]
            public string OtherCalendar { get; set; }
        }

        public class UpdateSingleEventWithRecipients
        {

            [JsonProperty("Subject")]
            public string Subject { get; set; }

            [JsonProperty("UpdatedTitle")]
            public string UpdatedTitle { get; set; }

            [JsonProperty("EventDate")]
            public string EventDate { get; set; }

            [JsonProperty("UpdatedStartDate")]
            public string UpdatedStartDate { get; set; }

            [JsonProperty("UpdatedDuration")]
            public string UpdatedDuration { get; set; }

            [JsonProperty("Recipients")]
            public string Recipients { get; set; }

            [JsonProperty("OtherCalendar")]
            public string OtherCalendar { get; set; }
        }

        public class UpdateRecurrenceNoRecipients
        {

            [JsonProperty("Subject")]
            public string Subject { get; set; }

            [JsonProperty("EventDate")]
            public string EventDate { get; set; }

            [JsonProperty("UpdatedTitle")]
            public string UpdatedTitle { get; set; }

            [JsonProperty("UpdatedStartDate")]
            public string UpdatedStartDate { get; set; }

            [JsonProperty("UpdatedDuration")]
            public string UpdatedDuration { get; set; }

            [JsonProperty("Recipients")]
            public string Recipients { get; set; }

            [JsonProperty("OtherCalendar")]
            public string OtherCalendar { get; set; }
        }

        public class UpdateRecurrenceWithRecipients
        {

            [JsonProperty("Subject")]
            public string Subject { get; set; }

            [JsonProperty("UpdatedTitle")]
            public string UpdatedTitle { get; set; }

            [JsonProperty("EventDate")]
            public string EventDate { get; set; }

            [JsonProperty("UpdatedStartDate")]
            public string UpdatedStartDate { get; set; }

            [JsonProperty("UpdatedDuration")]
            public string UpdatedDuration { get; set; }

            [JsonProperty("Recipients")]
            public string Recipients { get; set; }

            [JsonProperty("OtherCalendar")]
            public string OtherCalendar { get; set; }
        }

        public class RemoveSingleInstance
        {
            [JsonProperty("Subject")]
            public string Subject { get; set; }

            [JsonProperty("EventDate")]
            public string EventDate { get; set; }

            [JsonProperty("OtherCalendar")]
            public string OtherCalendar { get; set; }
        }

        public class RemoveAllInstances
        {
            [JsonProperty("Subject")]
            public string Subject { get; set; }

            [JsonProperty("EventDate")]
            public string EventDate { get; set; }

            [JsonProperty("OtherCalendar")]
            public string OtherCalendar { get; set; }
        }

        public class EventTypes
        {

            [JsonProperty("SingleEventNoRecipients")]
            public SingleEventNoRecipients SingleEventNoRecipients { get; set; }

            [JsonProperty("SingleEventWithRecipients")]
            public SingleEventWithRecipients SingleEventWithRecipients { get; set; }

            [JsonProperty("RecurrenceNoRecipients")]
            public RecurrenceNoRecipients RecurrenceNoRecipients { get; set; }

            [JsonProperty("RecurrenceWithRecipients")]
            public RecurrenceWithRecipients RecurrenceWithRecipients { get; set; }

            [JsonProperty("UpdateSingleEventNoRecipients")]
            public UpdateSingleEventNoRecipients UpdateSingleEventNoRecipients { get; set; }

            [JsonProperty("UpdateSingleEventWithRecipients")]
            public UpdateSingleEventWithRecipients UpdateSingleEventWithRecipients { get; set; }

            [JsonProperty("UpdateRecurrenceNoRecipients")]
            public UpdateRecurrenceNoRecipients UpdateRecurrenceNoRecipients { get; set; }

            [JsonProperty("UpdateRecurrenceWithRecipients")]
            public UpdateRecurrenceWithRecipients UpdateRecurrenceWithRecipients { get; set; }

            [JsonProperty("RemoveSingleInstance")]
            public RemoveSingleInstance RemoveSingleInstance { get; set; }

            [JsonProperty("RemoveAllInstances")]
            public RemoveAllInstances RemoveAllInstances { get; set; }

        }

        public class RootObject
        {

            [JsonProperty("Events")]
            public EventTypes EventTypes { get; set; }
        }



        public static RootObject LoadJSON()
        {
            RootObject events = null;

            try
            {
                string location = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

                using (StreamReader file = File.OpenText(location + "\\config.json"))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    events = (RootObject)serializer.Deserialize(file, typeof(RootObject));
                }
            }catch(Exception e)
            {
                Console.WriteLine("Error loading JSON File");
                throw e;
            }



            return events;

        }

        
        /// <summary>
        /// Main method of the command line application. Is used to call appropriate
        /// functions and hook into the Outlook VSTO. Obtains function inputs from user
        /// from the command line
        /// </summary>
        /// <param name="args">Command line arguments from users command line call</param>
        static void Main(string[] args)
        {

            string methodTag = "OutlookCLIMain";

            //Setup(); //Used for when testing since the plugin needs to be reloaded on startup.
            OpenOutlookIfNotRunning();
            Office.COMAddIn addIn = null;

            //finds the addin in Outlook and creates an COMAddIn Object
            //so that the functions of the add in can be called
            try
            {
                object addInName = "OutlookAddIn";
                Outlook.Application outlookApp = new Outlook.Application();
                addIn = outlookApp.COMAddIns.Item(ref addInName);
                outlookApp = null;
            }
            catch (COMException e)
            {
                Console.WriteLine("Error has occurred, check log files for details");
                LogWriter.WriteException(TAG, methodTag, e);
                Environment.Exit(-1);
            }


            /* Based on the first argument of the command line call depends which
             * function will be called. The rest of the args are passed to the function
             * along with the the outlook VSTO add in
             */

            try
            {

                switch (args[0])
                {
                    case "AddEventToMainCalendar":
                        AddEventToMainCalendar(args, addIn);
                        break;
                    case "AddEventToOtherCalendar":
                        AddEventToOtherCalendar(args, addIn);
                        break;
                    case "UpdateMainCalendarEvent":
                        UpdateMainCalendarEvent(args, addIn);
                        break;
                    case "UpdateOtherCalendarEvent":
                        UpdateOtherCalendarEvent(args, addIn);
                        break;
                    case "RemoveEventMainCalendar":
                        RemoveEventMainCalendar(args, addIn);
                        break;
                    case "RemoveEventOtherCalendar":
                        RemoveEventOtherCalendar(args, addIn);
                        break;
                }
            }
            catch(Exception e)
            {
                Console.WriteLine("Error has occurred, check log file");
                LogWriter.WriteException(TAG, methodTag, e);
                Environment.Exit(-1);
            }



        }
        /// <summary>
        /// Currently closes outlook if the application is running and restarts it.
        /// This is to ensure that any changes to the plugin are loaded in.
        /// Once the plugin is stable with little changes, it would be more efficient
        /// for tests if outlook remained open.
        /// </summary>
        public static void Setup()
        {

            string methodTag = "SetupOutlookCLI";

            try
            {
                CloseOutlookIfRunning();
                LogWriter.WriteInfo(TAG, methodTag, "Outlook was sucessfully closed");
            }
            catch { }
            finally
            {
                Process.Start("outlook.exe");

                System.Threading.Thread.Sleep(10000);
            }

            return;

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

        /// <summary>
        /// Trys to set OutlookApp to an active Outlook object. If the object is
        /// found then outlook is closed.
        /// </summary>
        public static void CloseOutlookIfRunning()
        {

            Outlook.Application OutlookApp;
            OutlookApp = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
            OutlookApp.Quit();

            System.Threading.Thread.Sleep(2000);

            return;

        }

        /// <summary>
        /// Sets any of the parameteres that are a string containing null to the value null.
        /// Handles having recipients and not having recipients. Depending on the number of parameters
        /// If the recipients parameter is not included in the command line call, sends an empty string
        /// array in its place. Calls the VSTO plugin to add an event to the main calendar in outlook.
        /// </summary>
        /// <param name="param">Array of the values given in the command line call</param>
        /// <param name="addIn">The hook into the VSTO plug in functions</param>
        public static void AddEventToMainCalendar(string[] param, Office.COMAddIn addIn)
        {
            string subject, startDate, recurrenceType, endDate, duration, recipientString;
            string[] recipients = null;
            string methodTag = "AddEventToMainCalendar";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);

            if(param[1].ToLower() == "json")
            {

                LogWriter.WriteInfo(TAG, methodTag, "Using JSON file instead of parameters");

                RootObject ro = new RootObject();
                ro = LoadJSON();
                dynamic et = null;

                if(param.Length < 2)
                {
                    Console.WriteLine("Missing parameter, Please add what type of event to add");
                    LogWriter.WriteWarning(TAG, methodTag, "Missing parameter, Please add what type of event to add");
                }
                else
                {
                    if (param[2].ToLower() == "singleeventnorecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Single Event No Recipients loaded from JSON");
                       et = ro.EventTypes.SingleEventNoRecipients;
                    }
                    else if(param[2].ToLower() == "singleeventwithrecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Single Event With Recipients loaded from JSON");
                        et = ro.EventTypes.SingleEventWithRecipients;
                    }
                    else if (param[2].ToLower() == "recurrencenorecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Recurrence No Recipients loaded from JSON");
                        et = ro.EventTypes.RecurrenceNoRecipients;
                    }
                    else if (param[2].ToLower() == "recurrencewithrecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Recurrence With Recipients loaded from JSON");
                        et = ro.EventTypes.RecurrenceWithRecipients;
                    }
                    else
                    {
                        Console.WriteLine("Invalid JSON function, see instructions or code for options");
                        LogWriter.WriteInfo(TAG, methodTag, "Invalid JSON function, see instructions or code for options");
                    }
                }

                subject = et.Subject;
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
                startDate = et.StartDate;
                LogWriter.WriteInfo(TAG, methodTag, "startDate set to: " + startDate);
                recurrenceType = et.RecurrenceType;
                LogWriter.WriteInfo(TAG, methodTag, "recurrenceType set to: " + recurrenceType);
                endDate = et.EndDate;
                LogWriter.WriteInfo(TAG, methodTag, "endDate set to: " + endDate);
                duration = et.Duration;
                LogWriter.WriteInfo(TAG, methodTag, "Duration set to: " + duration);
                recipientString = et.Recipients;
                LogWriter.WriteInfo(TAG, methodTag, "Recipients set to: " + recipientString);

                if (recipientString == null)
                {
                    recipients = new string[] { };
                }
                else
                {
                    recipients = recipientString.Split(',').Select(x => x.Trim()).ToArray();
                }

                LogWriter.WriteInfo(TAG, param[0], "Adding event in main calendar with params: "
                + subject + " " + startDate + " " + recurrenceType + " " + endDate + " " + duration + " "
                + string.Join(",", recipients));

            }
            else
            {

                LogWriter.WriteInfo(TAG, methodTag, "Using parameters instead of JSON file");

                if (param.Length < 6)
                {
                    LogWriter.WriteWarning(TAG, methodTag, "Parameters entered should be atleast 6 even if null. Entered: " + param.Length);
                    throw new IndexOutOfRangeException();
                }

                for (int i = 1; i < param.Length; i++)
                {
                    if (param[i].ToLower() == "null")
                    {
                        param[i] = null;
                    }
                }

                subject = param[1];
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
                startDate = param[2];
                LogWriter.WriteInfo(TAG, methodTag, "startDate set to: " + startDate);
                recurrenceType = param[3];
                LogWriter.WriteInfo(TAG, methodTag, "recurrenceType set to: " + recurrenceType);
                endDate = param[4];
                LogWriter.WriteInfo(TAG, methodTag, "endDate set to: " + endDate);
                duration = param[5];
                LogWriter.WriteInfo(TAG, methodTag, "duration set to: " + duration);

                if (param.Length == 7 && param[6] != null)
                {
                    recipients = param[6].Split(',').Select(x => x.Trim()).ToArray();
                    LogWriter.WriteInfo(TAG, param[0], "Adding event in main calendar with params: "
                        + subject + " " + startDate + " " + recurrenceType + " " + endDate + " " + duration + " "
                        + string.Join(",", recipients));

                }
                else
                {
                    //passing null in place of the array didn't work, so instead pass an empty string array
                    recipients = new string[] { };
                    LogWriter.WriteInfo(TAG, param[0], "Adding event in main calendar with params: "
                        + param[1] + " " + param[2] + " " + param[3] + " " + param[4] + " " + param[5]);
                }
            }

            try
            {
                addIn.Object.AddEventToMainCalendar(subject, startDate, recurrenceType, endDate, duration, recipients);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        /// <summary>
        /// Sets any of the parameteres that are a string containing null to the value null.
        /// Handles having recipients and not having recipients. Depending on the number of parameters
        /// If the recipients parameter is not included in the command line call, sends an empty string
        /// array in its place. Calls the VSTO plugin to update an event to the main calendar in outlook.
        /// </summary>
        /// <param name="param">Array of the values given in the command line call</param>
        /// <param name="addIn">The hook into the VSTO plug in functions</param>
        public static void UpdateMainCalendarEvent(string[] param, Office.COMAddIn addIn)
        {
            string subject, eventDate, updatedTitle, updatedStartDate, updatedDuration, recipientString; 
            string[] recipients = null;
            string methodTag = "UpdateMainCalendarEvent";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);

            if (param[1].ToLower() == "json")
            {

                LogWriter.WriteInfo(TAG, methodTag, "Using JSON file instead of parameters");

                RootObject ro = new RootObject();
                ro = LoadJSON();
                dynamic et = null;

                if (param.Length < 2)
                {
                    Console.WriteLine("Missing parameter, Please add what type of event to add");
                    LogWriter.WriteWarning(TAG, methodTag, "Missing parameter, Please add what type of event to update");
                }
                else
                {
                    if (param[2].ToLower() == "updatesingleeventnorecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Update Single Event No Recipients loaded from JSON");
                        et = ro.EventTypes.UpdateSingleEventNoRecipients;
                    }
                    else if (param[2].ToLower() == "updatesingleeventwithrecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Update Single Event with Recipients loaded from JSON");
                        et = ro.EventTypes.UpdateSingleEventWithRecipients;
                    }
                    else if (param[2].ToLower() == "updaterecurrencenorecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Update Recurrence No Recipients loaded from JSON");
                        et = ro.EventTypes.UpdateRecurrenceNoRecipients;
                    }
                    else if (param[2].ToLower() == "updaterecurrencewithrecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Update Recurrence With Recipients loaded from JSON");
                        et = ro.EventTypes.UpdateRecurrenceWithRecipients;
                    }
                    else
                    {
                        Console.WriteLine("Invalid JSON function, see instructions or code for options");
                        LogWriter.WriteInfo(TAG, methodTag, "Invalid JSON function, see instructions or code for options");
                    }                    
                }

                subject = et.Subject;
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
                eventDate = et.EventDate;
                LogWriter.WriteInfo(TAG, methodTag, "eventDate set to: " + eventDate);
                updatedTitle = et.UpdatedTitle;
                LogWriter.WriteInfo(TAG, methodTag, "updatedTitlte set to: " + updatedTitle);
                updatedStartDate = et.UpdatedStartDate;
                LogWriter.WriteInfo(TAG, methodTag, "updatedStartDate set to: " + updatedStartDate);
                updatedDuration = et.UpdatedDuration;
                LogWriter.WriteInfo(TAG, methodTag, "updatedDuration set to: " + updatedDuration);
                recipientString = et.Recipients;
                LogWriter.WriteInfo(TAG, methodTag, "Recipients set to: " + recipientString);

                if (recipientString == null)
                {
                    recipients = new string[] { };
                }
                else
                {
                    recipients = recipientString.Split(',').Select(x => x.Trim()).ToArray();
                }

            }
            else
            {

                LogWriter.WriteInfo(TAG, methodTag, "Using parameters instead of JSON file");

                if (param.Length < 6)
                {
                    LogWriter.WriteWarning(TAG, methodTag, "Parameters entered should be atleast 6 even if null. Entered: " + param.Length);
                    throw new IndexOutOfRangeException();
                }


                for (int i = 1; i < param.Length; i++)
                {
                    if (param[i].ToLower() == "null")
                    {
                        param[i] = null;
                    }
                }

                subject = param[1];
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
                eventDate = param[2];
                LogWriter.WriteInfo(TAG, methodTag, "eventDate set to: " + eventDate);
                updatedTitle = param[3];
                LogWriter.WriteInfo(TAG, methodTag, "updatedTitlte set to: " + updatedTitle);
                updatedStartDate = param[4];
                LogWriter.WriteInfo(TAG, methodTag, "updatedStartDate set to: " + updatedStartDate);
                updatedDuration = param[5];
                LogWriter.WriteInfo(TAG, methodTag, "updatedDuration set to: " + updatedDuration);

                if (param.Length == 7 && param[6] != null)
                {
                    recipients = param[6].Split(',').Select(x => x.Trim()).ToArray();
                    LogWriter.WriteInfo(TAG, param[0], "Updating event in main calendar with params: "
                     + subject + " " + eventDate + " " + updatedTitle + " " + updatedStartDate + " " + updatedDuration + " "
                     + string.Join(",", recipients));

                }
                else
                {
                    //passing null in place of the array didn't work, so instead pass an empty string array
                    recipients = new string[] { };
                    LogWriter.WriteInfo(TAG, param[0], "Updating event in main calendar with params: "
                        + subject + " " + eventDate + " " + updatedTitle + " "
                        + updatedStartDate + " " + updatedDuration);
                }
            }

            try
            {
                addIn.Object.UpdateMainCalendarEvent(subject, eventDate, updatedTitle, updatedStartDate,
                updatedDuration, recipients);
            }
            catch (Exception e)
            {
                throw e;
            }


        }

        /// <summary>
        /// Sets any of the parameteres that are a string containing null to the value null.
        /// Handles having recipients and not having recipients. Depending on the number of parameters
        /// If the recipients parameter is not included in the command line call, sends an empty string
        /// array in its place. Calls the VSTO plugin to update an event to other calendar in outlook.
        /// </summary>
        /// <param name="param">Array of the values given in the command line call</param>
        /// <param name="addIn">The hook into the VSTO plug in functions</param>
        public static void UpdateOtherCalendarEvent(string[] param, Office.COMAddIn addIn)
        {

            string subject, eventDate, updatedTitle, updatedStartDate, updatedDuration, recipientString, otherCalendar;
            string[] recipients;
            string methodTag = "UpdateOtherCalendarEvent";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);

            if (param[1].ToLower() == "json")
            {

                LogWriter.WriteInfo(TAG, methodTag, "Using JSON file instead of parameters");

                RootObject ro = new RootObject();
                ro = LoadJSON();
                dynamic et = null;

                if (param.Length < 2)
                {
                    Console.WriteLine("Missing parameter, Please add what type of event to add");
                    LogWriter.WriteWarning(TAG, methodTag, "Missing parameter, Please add what type of event to update");
                }
                else
                {
                    if (param[2].ToLower() == "updatesingleeventnorecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Update Single Event No Recipients loaded from JSON");
                        et = ro.EventTypes.UpdateSingleEventNoRecipients;
                    }
                    else if (param[2].ToLower() == "updatesingleeventwithrecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Update Single Event with Recipients loaded from JSON");
                        et = ro.EventTypes.UpdateSingleEventWithRecipients;
                    }
                    else if (param[2].ToLower() == "updaterecurrencenorecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Update Recurrence No Recipients loaded from JSON");
                        et = ro.EventTypes.UpdateRecurrenceNoRecipients;
                    }
                    else if (param[2].ToLower() == "updaterecurrencewithrecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Update Recurrence With Recipients loaded from JSON");
                        et = ro.EventTypes.UpdateRecurrenceWithRecipients;
                    }
                    else
                    {
                        Console.WriteLine("Invalid JSON function, see instructions or code for options");
                        LogWriter.WriteInfo(TAG, methodTag, "Invalid JSON function, see instructions or code for options");
                    }
                }

                subject = et.Subject;
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
                eventDate = et.EventDate;
                LogWriter.WriteInfo(TAG, methodTag, "eventDate set to: " + eventDate);
                updatedTitle = et.UpdatedTitle;
                LogWriter.WriteInfo(TAG, methodTag, "updatedTitlte set to: " + updatedTitle);
                updatedStartDate = et.UpdatedStartDate;
                LogWriter.WriteInfo(TAG, methodTag, "updatedStartDate set to: " + updatedStartDate);
                updatedDuration = et.UpdatedDuration;
                LogWriter.WriteInfo(TAG, methodTag, "updatedDuration set to: " + updatedDuration);
                recipientString = et.Recipients;
                LogWriter.WriteInfo(TAG, methodTag, "Recipients set to: " + recipientString);
                otherCalendar = et.OtherCalendar;
                LogWriter.WriteInfo(TAG, methodTag, "otherCalendar set to: " + otherCalendar);

                if (recipientString == null)
                {
                    recipients = new string[] { };
                }
                else
                {
                    recipients = recipientString.Split(',').Select(x => x.Trim()).ToArray();
                }

            }
            else
            {

                LogWriter.WriteInfo(TAG, methodTag, "Using parameters instead of JSON file");

                if (param.Length < 7)
                {
                    LogWriter.WriteWarning(TAG, methodTag, "Parameters entered should be atleast 2 even if null. Entered: " + param.Length);
                    throw new IndexOutOfRangeException();
                }

                for (int i = 1; i < param.Length; i++)
                {
                    if (param[i].ToLower() == "null")
                    {
                        param[i] = null;
                    }
                }

                subject = param[1];
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
                eventDate = param[2];
                LogWriter.WriteInfo(TAG, methodTag, "eventDate set to: " + eventDate);
                updatedTitle = param[3];
                LogWriter.WriteInfo(TAG, methodTag, "updatedTitlte set to: " + updatedTitle);
                updatedStartDate = param[4];
                LogWriter.WriteInfo(TAG, methodTag, "updatedStartDate set to: " + updatedStartDate);
                updatedDuration = param[5];
                LogWriter.WriteInfo(TAG, methodTag, "updatedDuration set to: " + updatedDuration);
                otherCalendar = param[6];
                LogWriter.WriteInfo(TAG, methodTag, "otherCalendar set to: " + otherCalendar);

                if (param.Length == 8 && param[7] != null)
                {
                    recipients = param[7].Split(',').Select(x => x.Trim()).ToArray();
                    LogWriter.WriteInfo(TAG, param[0], "Updating event in other calendar with params: "
                     + subject + " " + eventDate + " " + updatedTitle + " " + updatedStartDate + " " + updatedDuration + " "
                     + otherCalendar + " " + string.Join(",", recipients));

                }
                else
                {
                    //passing null in place of the array didn't work, so instead pass an empty string array
                    recipients = new string[] { };
                    LogWriter.WriteInfo(TAG, param[0], "Updating event in other calendar with params: "
                     + subject + " " + eventDate + " " + updatedTitle + " " + updatedStartDate + " " + updatedDuration + " "
                     + otherCalendar);
                }
            }

                try
                {
                    addIn.Object.UpdateOtherCalendarEvent(subject, eventDate, updatedTitle, updatedStartDate,
                    updatedDuration, otherCalendar, recipients);
                }
                catch (Exception e)
                {
                    throw e;
                }
            

        }

        /// <summary>
        /// Sets any of the parameteres that are a string containing null to the value null.
        /// Handles having recipients and not having recipients. Depending on the number of parameters
        /// If the recipients parameter is not included in the command line call, sends an empty string
        /// array in its place. Calls the VSTO plugin to add an event to other calendar in outlook.
        /// </summary>
        /// <param name="param">Array of the values given in the command line call</param>
        /// <param name="addIn">The hook into the VSTO plug in functions</param>
        public static void AddEventToOtherCalendar(string[] param, Office.COMAddIn addIn)
        {
            string subject, startDate, recurrenceType, endDate, duration, recipientString, otherCalendar;
            string[] recipients;
            string methodTag = "AddEventToOtherCalendar";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);


            if (param[1].ToLower() == "json")
            {

                LogWriter.WriteInfo(TAG, methodTag, "Using JSON file instead of parameters");

                RootObject ro = new RootObject();
                ro = LoadJSON();
                dynamic et = null;

                if (param.Length < 2)
                {
                    Console.WriteLine("Missing parameter, Please add what type of event to add");
                    LogWriter.WriteInfo(TAG, methodTag, "Missing parameter, Please add what type of event to add");
                }
                else
                {
                    if (param[2].ToLower() == "updatesingleeventnorecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Update Single Event No Recipients loaded from JSON");
                        et = ro.EventTypes.UpdateSingleEventNoRecipients;
                    }
                    else if (param[2].ToLower() == "updatesingleeventwithrecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Update Single Event with Recipients loaded from JSON");
                        et = ro.EventTypes.UpdateSingleEventWithRecipients;
                    }
                    else if (param[2].ToLower() == "updaterecurrencenorecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Update Recurrence No Recipients loaded from JSON");
                        et = ro.EventTypes.UpdateRecurrenceNoRecipients;
                    }
                    else if (param[2].ToLower() == "updaterecurrencewithrecipients")
                    {
                        LogWriter.WriteInfo(TAG, methodTag, "Update Recurrence With Recipients loaded from JSON");
                        et = ro.EventTypes.UpdateRecurrenceWithRecipients;
                    }
                    else
                    {
                        Console.WriteLine("Invalid JSON function, see instructions or code for options");
                        LogWriter.WriteInfo(TAG, methodTag, "Invalid JSON function, see instructions or code for options");
                    }   
                }

                subject = et.Subject;
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
                startDate = et.StartDate;
                LogWriter.WriteInfo(TAG, methodTag, "startDate set to: " + startDate);
                recurrenceType = et.RecurrenceType;
                LogWriter.WriteInfo(TAG, methodTag, "recurrenceType set to: " + recurrenceType);
                endDate = et.EndDate;
                LogWriter.WriteInfo(TAG, methodTag, "endDate set to: " + endDate);
                duration = et.Duration;
                LogWriter.WriteInfo(TAG, methodTag, "Duration set to: " + duration);
                recipientString = et.Recipients;
                LogWriter.WriteInfo(TAG, methodTag, "Recipients set to: " + recipientString);
                otherCalendar = et.OtherCalendar;
                LogWriter.WriteInfo(TAG, methodTag, "otherCalendar set to: " + otherCalendar);

                if (recipientString == null)
                {
                    recipients = new string[] { };
                }
                else
                {
                    recipients = recipientString.Split(',').Select(x => x.Trim()).ToArray();
                }

                LogWriter.WriteInfo(TAG, param[0], "Adding event in other calendar with params: "
                + subject + " " + startDate + " " + recurrenceType + " " + endDate + " " + duration + " "
                + string.Join(",", recipients));

            }
            else
            {

                LogWriter.WriteInfo(TAG, methodTag, "Using parameters instead of JSON file");

                if (param.Length < 7)
                {
                    LogWriter.WriteWarning(TAG, methodTag, "Parameters entered should be atleast 7 even if null. Entered: " + param.Length);
                    throw new IndexOutOfRangeException();
                }

                for (int i = 1; i < param.Length; i++)
                {
                    if (param[i].ToLower() == "null")
                    {
                        param[i] = null;
                    }
                }

                subject = param[1];
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
                startDate = param[2];
                LogWriter.WriteInfo(TAG, methodTag, "startDate set to: " + startDate);
                recurrenceType = param[3];
                LogWriter.WriteInfo(TAG, methodTag, "recurrenceType set to: " + recurrenceType);
                endDate = param[4];
                LogWriter.WriteInfo(TAG, methodTag, "endDate set to: " + endDate);
                duration = param[5];
                LogWriter.WriteInfo(TAG, methodTag, "duration set to: " + duration);
                otherCalendar = param[6];
                LogWriter.WriteInfo(TAG, methodTag, "otherCalendar set to: " + otherCalendar);

                if (param.Length == 8 && param[7] != null)
                {
                    recipients = param[7].Split(',').Select(x => x.Trim()).ToArray();
                    LogWriter.WriteInfo(TAG, param[0], "Adding event in other calendar with params: "
                     + subject + " " + startDate + " " + recurrenceType + " " + endDate + " " + duration + " "
                     + " " + string.Join(",", recipients));
                }
                else
                {
                    //passing null in place of the array didn't work, so instead pass an empty string array
                    recipients = new string[] { };
                    LogWriter.WriteInfo(TAG, param[0], "Adding event in other calendar with params: "
                     + subject + " " + startDate + " " + recurrenceType + " " + endDate + " " + duration);
                }
            }
            try
            {
                addIn.Object.AddEventToOtherCalendar(subject, startDate, recurrenceType, endDate, duration, otherCalendar, recipients);
            }
            catch(Exception e)
            {
                throw e;
            }
            
        }

        /// <summary>
        /// Sets any of the parameteres that are a string containing null to the value null.
        /// Handles having recipients and not having recipients. Depending on the number of parameters
        /// If the recipients parameter is not included in the command line call, sends an empty string
        /// array in its place. Calls the VSTO plugin to remove an event from the main  calendar in outlook.
        /// </summary>
        /// <param name="param">Array of the values given in the command line call</param>
        /// <param name="addIn">The hook into the VSTO plug in functions</param>
        public static void RemoveEventMainCalendar(string[] param, Office.COMAddIn addIn)
        {
            string subject, eventDate;
            string methodTag = "RemoveEventMainCalendar";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);
            if (param[1] == "json")
            {

                LogWriter.WriteInfo(TAG, methodTag, "Using JSON file instead of parameters");

                RootObject ro = new RootObject();
                ro = LoadJSON();
                dynamic et = null;

                if (param[2].ToLower() == "removesingleinstance")
                {
                    LogWriter.WriteInfo(TAG, methodTag, "Remove Single Instance loaded from JSON file");
                    et = ro.EventTypes.RemoveSingleInstance;
                }
                else if (param[2].ToLower() == "removeallinstances")
                {
                    LogWriter.WriteInfo(TAG, methodTag, "Remove All Instances loaded from JSON file");
                    et = ro.EventTypes.RemoveAllInstances;
                }
                else
                {
                    Console.WriteLine("Invalid JSON function, see instructions or code for options");
                    LogWriter.WriteInfo(TAG, methodTag, "Invalid JSON function, see instructions or code for options");
                }

                subject = et.Subject;
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
                eventDate = et.EventDate;
                LogWriter.WriteInfo(TAG, methodTag, "eventDate set to: " + eventDate);

            }
            else
            {
                LogWriter.WriteInfo(TAG, methodTag, "Using parameters instead of JSON file");

                if (param.Length < 2)
                {
                    LogWriter.WriteWarning(TAG, methodTag, "Parameters entered should be atleast 2 even if null. Entered: " + param.Length);
                    throw new IndexOutOfRangeException();
                }

                for (int i = 1; i < param.Length; i++)
                {
                    if (param[i].ToLower() == "null")
                    {
                        param[i] = null;
                    }
                }

                subject = param[1];
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
                eventDate = param[2];
                LogWriter.WriteInfo(TAG, methodTag, "eventDate set to: " + eventDate);

            }

            LogWriter.WriteInfo(TAG, param[0], "Removing event from main calendar with params: " + subject + " " + eventDate);

            try
            {
                addIn.Object.RemoveEventMainCalendar(subject, eventDate);
            }
            catch(Exception e)
            {
                throw e;
            }
            
        }

        /// <summary>
        /// Sets any of the parameteres that are a string containing null to the value null.
        /// Handles having recipients and not having recipients. Depending on the number of parameters
        /// If the recipients parameter is not included in the command line call, sends an empty string
        /// array in its place. Calls the VSTO plugin to remove an event from other calendar in outlook.
        /// </summary>
        /// <param name="param">Array of the values given in the command line call</param>
        /// <param name="addIn">The hook into the VSTO plug in functions</param>
        public static void RemoveEventOtherCalendar(string[] param, Office.COMAddIn addIn)
        {
            string subject, eventDate, otherCalendar;
            string methodTag = "RemoveEventOtherCalendar";

            LogWriter.WriteInfo(TAG, methodTag, "Starting: " + methodTag + " in " + TAG);

            if(param[1] == "json")
            {

                LogWriter.WriteInfo(TAG, methodTag, "Using JSON file instead of parameters");

                RootObject ro = new RootObject();
                ro = LoadJSON();
                dynamic et = null;

                if (param[2].ToLower() == "removesingleinstance")
                {
                    LogWriter.WriteInfo(TAG, methodTag, "Remove Single Instance loaded from JSON file");
                    et = ro.EventTypes.RemoveSingleInstance;
                }
                else if(param[2].ToLower() == "removeallinstances")
                {
                    LogWriter.WriteInfo(TAG, methodTag, "Remove All Instances loaded from JSON file");
                    et = ro.EventTypes.RemoveAllInstances;
                }
                else
                {
                    Console.WriteLine("Invalid JSON function, see instructions or code for options");
                    LogWriter.WriteInfo(TAG, methodTag, "Invalid JSON function, see instructions or code for options");
                }

                subject = et.Subject;
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
                eventDate = et.EventDate;
                LogWriter.WriteInfo(TAG, methodTag, "eventDate set to: " + eventDate);
                otherCalendar = et.OtherCalendar;
                LogWriter.WriteInfo(TAG, methodTag, "otherCalendar set to: " + otherCalendar);


            }
            else{
                if (param.Length < 3)
                {
                    LogWriter.WriteWarning(TAG, methodTag, "Parameters entered should be atleast 2 even if null. Entered: " + param.Length);
                    throw new IndexOutOfRangeException();
                }

                for (int i = 1; i < param.Length; i++)
                {
                    if (param[i].ToLower() == "null")
                    {
                        param[i] = null;
                    }
                }

                subject = param[1];
                LogWriter.WriteInfo(TAG, methodTag, "Subject set to: " + subject);
                eventDate = param[2];
                LogWriter.WriteInfo(TAG, methodTag, "eventDate set to: " + eventDate);
                otherCalendar = param[3];


            }

            LogWriter.WriteInfo(TAG, param[0], "Removing event from other calendar with params: " + subject + " " + eventDate + " " + otherCalendar);

            try
            {
                addIn.Object.RemoveEventOtherCalendar(subject, eventDate, otherCalendar);
            }
            catch(Exception e)
            {
                throw e;
            }
            
        }

    }
}
