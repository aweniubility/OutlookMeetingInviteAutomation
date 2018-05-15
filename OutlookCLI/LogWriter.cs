using System;
using System.IO;


namespace OutlookCLI
{
    public static class LogWriter
    {
        public static string logFile = "OutlookTool-Log.txt";
        public static string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\" + "Desktop\\Calendar_LOGS\\";
        public static string logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\" + "Desktop\\Calendar_LOGS\\" + logFile; //do to permissions issues in debug folder

        public static void WriteInfo(string compontentTag, string methodTag, string info)
        {
            StreamWriter log;
            FileStream fileStream = null;
            FileInfo logFileInfo;

            string timeStamp = DateTime.Now.ToString("yyyy-MM-dd-hh:mm:ss");

            Directory.CreateDirectory(folderPath);

            logFileInfo = new FileInfo(logFilePath);

            if (!logFileInfo.Exists)
            {
                fileStream = logFileInfo.Create();
            }
            else
            {
                fileStream = new FileStream(logFilePath, FileMode.Append);

            }
            log = new StreamWriter(fileStream);
            log.WriteLine("{0,-22}{1,-12}{2,-18} - {3,-24} -> {4}", timeStamp, "INFO", compontentTag, methodTag, info);
            log.Close();

        }

        public static void WriteWarning(string compontentTag, string methodTag, string info)
        {
            StreamWriter log;
            FileStream fileStream = null;
            FileInfo logFileInfo;

            Directory.CreateDirectory(folderPath);

            logFileInfo = new FileInfo(logFilePath);

            string timeStamp = DateTime.Now.ToString("yyyy-MM-dd-hh:mm:ss");

            if (!logFileInfo.Exists)
            {
                fileStream = logFileInfo.Create();
            }
            else
            {
                fileStream = new FileStream(logFilePath, FileMode.Append);

            }

            log = new StreamWriter(fileStream);
            log.WriteLine("{0,-22}{1,-12}{2,-18} - {3,-24} -> {4}", timeStamp, "WARNING", compontentTag, methodTag, info);
            log.Close();

        }

        public static void WriteException(string compontentTag, string methodTag, Exception e)
        {
            StreamWriter log;
            FileStream fileStream = null;
            FileInfo logFileInfo;

            string timeStamp = DateTime.Now.ToString("yyyy-MM-dd-hh:mm:ss");

            Directory.CreateDirectory(folderPath);

            logFileInfo = new FileInfo(logFilePath);

            if (!logFileInfo.Exists)
            {
                fileStream = logFileInfo.Create();
            }
            else
            {
                fileStream = new FileStream(logFilePath, FileMode.Append);

            }
            log = new StreamWriter(fileStream);
            log.WriteLine("{0,-22}{1,-12}{2,-18} - {3,-24} -> {4}",timeStamp, "EXCEPTION", compontentTag, methodTag, e.Message.ToString());
            log.WriteLine("{0}", e.ToString());
            log.Close();

        }



    }
}
