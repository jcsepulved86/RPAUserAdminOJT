﻿using System;
using System.Configuration;

namespace RPAUserAdminOJT.Controllers.Models
{
    public class GlobalVar
    {
        public static string Domain = "multienlace.com.co";
        public static string Move = "Retiros";
        public static string UserAdm = ConfigurationManager.AppSettings["BOTUser"];
        public static string Passwrd = ConfigurationManager.AppSettings["BOTPassword"];
        public static string passwordUser = string.Empty;
        public static bool addUserGroup;
        public static bool NeverPswd;
        public static bool desactiveAccount;
        public static bool accessEntires = false;
        public static bool deleteEntires = false;
        public static int countYESProcess = 0;
        public static int countNOProcess = 0;
        public static int CountDeshabilitadoYESProcess = 0;
        public static int CountDeshabilitadoNOProcess = 0;
        public static string filePath = string.Empty;
        public static string RedUser = string.Empty;
        public static bool outLoadFile = false;
        public static bool existUser = false;
        public static DateTime TimeLogs;
        public static bool validTittle = true;

        public static int countCSV = 0;

        public static string rootMain { get; set; }
        public static string rootLote { get; set; }
        public static string rootLogs { get; set; }


    }
}
