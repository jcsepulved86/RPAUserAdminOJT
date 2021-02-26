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
        public static int countCSV = 0;
        public static StringBuilder sbOutput = new StringBuilder();
        public static string NameRPA = ConfigurationManager.AppSettings["NameRPA"];
        public static string DirectoryCSV = ConfigurationManager.AppSettings["directoryCSV"];
        #endregion

        #region Read Each Line of Writing an Excel
        public static void WriteExcel(string Action, string DateModify = "", string NumLote = "", string NameBox = "", string EmailBox = "", string IdBox = "", string FullName = "", string Cedula = "", string Client = "", string CodPCR = "", string Bassement = "", string UserRed = "", string Password = "")
        {
            
            try
            {

                CultureInfo ci = CultureInfo.CreateSpecificCulture(CultureInfo.CurrentCulture.Name); ci.DateTimeFormat.ShortDatePattern = "ddMMyyyy";
                Thread.CurrentThread.CurrentCulture = ci;

                string timeLogs = Models.GlobalVar.TimeLogs.ToString("ddMMyyyy_Hmm");

                CreateDirectory();


                string[][] inTittle = new string[][]{
                    new string[]{ "Fecha y Hora","Maquina","Usuario Maquina","Accion","Fecha Modificacion", "Numero Lote", "Nombre Jefe","Email Jefe","Cedula Jefe","Nombre Usuario","Cedula Usuario", "Cliente","Codigo PCR", "Usuario Base", "Usuario Red", "Clave"}
                };

                string[][] inOutput = new string[][]{
                    new string[]{ DateTime.Now.ToLocalTime().ToString(), Environment.MachineName, Environment.UserName, Action, DateModify, NumLote, NameBox, EmailBox, IdBox, FullName, Cedula, Client, CodPCR, Bassement, UserRed, Password }
                };


                if (countCSV == 0)
                {
                    sbOutput.AppendLine(string.Join(",", inTittle[0]));
                    sbOutput.AppendLine(string.Join(",", inOutput[0]));
                }
                else
                {
                    sbOutput.AppendLine(string.Join(",", inOutput[0]));
                }

                File.AppendAllText(DirectoryCSV + "\\" + NameRPA + "_" + timeLogs + "_" + Environment.UserName + ".csv", sbOutput.ToString());

                Models.GlobalVar.rootMain = DirectoryCSV + "\\" + NameRPA + "_" + timeLogs + "_" + Environment.UserName + ".csv";
                sbOutput.Clear();

                countCSV++;

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
