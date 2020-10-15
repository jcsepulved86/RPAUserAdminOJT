using System.DirectoryServices.AccountManagement;

namespace RPAUserAdminOJT.Controllers.Models
{
    public class GlobalVar
    {

        public static string Domain = "multienlace.com.co";
        public static string UserAdm = "user.gestusuarios";
        public static string Passwrd = "6r](71w67vy9*";
        public static string passwordUser = string.Empty;

        public static bool addUserGroup;
        public static bool NeverPswd;
        public static bool desactiveAccount;
        public static bool accessEntires = false;
        public static bool deleteEntires = false;

        public static int countYESProcess = 0;
        public static int countNOProcess = 0;

        public static string filePath = string.Empty;

        public static bool outLoadFile = false;

    }
}
