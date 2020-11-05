using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using CsvHelper;
using System.Text;
using System.Globalization;
using System.Threading;

namespace RPAUserAdminOJT.Controllers.Utility
{
    public class FillExcel
    {
        #region Global Variable
        public static string NameRPA = ConfigurationManager.AppSettings["NameRPA"];
        public static string DirectoryCSV = ConfigurationManager.AppSettings["directoryCSV"];
        #endregion

        #region Read Each Line of Writing an Excel
        public static void WriteExcel(string Action, string Operation, string Result)
        {
            try
            {
                CultureInfo ci = CultureInfo.CreateSpecificCulture(CultureInfo.CurrentCulture.Name); ci.DateTimeFormat.ShortDatePattern = "ddMMyyyy";
                Thread.CurrentThread.CurrentCulture = ci;

                CreateDirectory();

                DateTime time = DateTime.Now;
                String today;
                today = time.ToString("ddMMyyyy");

                StringBuilder sbOutput = new StringBuilder();

                string[][] inaOutput = new string[][]{
                new string[]{ DateTime.Now.ToLocalTime().ToString(), Environment.MachineName, Environment.UserName, Action, Operation, Result }
                };

                sbOutput.AppendLine(string.Join(",", inaOutput[0]));

                File.AppendAllText(DirectoryCSV + "\\" + NameRPA + "_" + today + "_" + Environment.UserName + ".csv", sbOutput.ToString());
            }
            catch (Exception ex)
            {
                LOGRobotica.Controllers.LogApplication.LogWrite("WriteExcel ==> " + "Exception: " + ex.Message.ToString());
            }

            

        }
        #endregion

        #region FUNCTION FOR CREATE DIRECTORY
        private static void CreateDirectory()
        {
            try
            {
                if (!Directory.Exists(DirectoryCSV))
                    Directory.CreateDirectory(DirectoryCSV);


            }
            catch (DirectoryNotFoundException ex)
            {
                throw new Exception(ex.Message);

            }
        }
        #endregion

    }
}
