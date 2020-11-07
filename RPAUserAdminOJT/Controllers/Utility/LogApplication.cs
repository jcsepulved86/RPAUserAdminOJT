#region USINGS
using System;
using System.IO;
using System.Reflection;
using System.Configuration;
using System.Globalization;
using System.Threading;
#endregion

namespace RPAUserAdminOJT.Controllers.Utility
{
    public class LogApplication
    {
        #region  STATIC VARIABLE
        public static string resultLogWr = string.Empty;
        public static string resultLog = string.Empty;
        public static string m_exePath = string.Empty;
        public static string NameRPA = ConfigurationManager.AppSettings["NameRPA"];
        public static string DirectoryLogs = ConfigurationManager.AppSettings["DirectoryLogs"];
        #endregion

        #region FUNCTION FOR WRITER IN LOG FILES
        public static void LogWrite(string logMessage)
        {
            try
            {
                CultureInfo ci = CultureInfo.CreateSpecificCulture(CultureInfo.CurrentCulture.Name); ci.DateTimeFormat.ShortDatePattern = "ddMMyyyy";
                Thread.CurrentThread.CurrentCulture = ci;

                CreateDirectory();

                string timeLogs = Models.GlobalVar.TimeLogs.ToString("ddMMyyyy_Hmm");

                using (StreamWriter w = File.AppendText(DirectoryLogs + "\\" + NameRPA + "_" + timeLogs + "_" + Environment.UserName + ".log"))
                {
                    Log(logMessage, w);
                    resultLogWr = "true";
                }
            }
            catch (Exception)
            {
                resultLogWr = "false";
            }

        }
        #endregion

        #region FUNCTION AND LOGIC OF LOG MESSAGE
        private static void Log(string logMessage, TextWriter txtWriter)
        {
            try
            {
                txtWriter.WriteLine("[" + DateTime.Now.ToLocalTime() + " " + Environment.MachineName + "\\" + Environment.UserName + "] " + NameRPA + " -> " + logMessage);
                resultLog = "true";
            }
            catch (Exception)
            {
                resultLog = "false";
            }

        }
        #endregion

        #region FUNCTION FOR CREATE DIRECTORY
        private static void CreateDirectory()
        {
            try
            {
                if (!Directory.Exists(DirectoryLogs))
                    Directory.CreateDirectory(DirectoryLogs);


            }
            catch (DirectoryNotFoundException ex)
            {
                throw new Exception(ex.Message);

            }
        }
        #endregion

    }
}