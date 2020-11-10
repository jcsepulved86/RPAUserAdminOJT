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
        public static void WriteExcel(string Action, string DateModify = "", string NameBox = "", string EmailBox = "", string IdBox = "", string FullName = "", string client = "", string CodPCR = "", string Operation = "", string Result = "", string Bassement = "")
        {
            try
            {
                CultureInfo ci = CultureInfo.CreateSpecificCulture(CultureInfo.CurrentCulture.Name); ci.DateTimeFormat.ShortDatePattern = "ddMMyyyy";
                Thread.CurrentThread.CurrentCulture = ci;

                CreateDirectory();

                string timeLogs = Models.GlobalVar.TimeLogs.ToString("ddMMyyyy_Hmm");

                StringBuilder sbOutput = new StringBuilder();

                string[][] inaOutput = new string[][]{
                new string[]{ DateTime.Now.ToLocalTime().ToString(), Environment.MachineName, Environment.UserName, Action, DateModify, NameBox, EmailBox, IdBox, FullName, client, CodPCR, Operation, Result }
                };

                sbOutput.AppendLine(string.Join(",", inaOutput[0]));

                File.AppendAllText(DirectoryCSV + "\\" + NameRPA + "_" + timeLogs + "_" + Environment.UserName + ".csv", sbOutput.ToString());
            }
            catch (Exception ex)
            {
                Utility.LogApplication.LogWrite("WriteCsv ==> " + "Exception: " + ex.Message.ToString());
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
