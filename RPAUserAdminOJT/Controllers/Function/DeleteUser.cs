using System;
using System.DirectoryServices.AccountManagement;
using System.Windows.Forms;

namespace RPAUserAdminOJT.Controllers.Function
{
    public class Delete
    {

        public static PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, Models.GlobalVar.Domain, Models.GlobalVar.UserAdm, Models.GlobalVar.Passwrd);

        public static void User(string users)
        {
            string CrtMessage = string.Empty;

            try
            {
                ActiveDirectory.Query.QueryProvider queryProvider = new ActiveDirectory.Query.QueryProvider(principalContext);
                UserPrincipal user = ActiveDirectory.Query.QueryProvider.Get(users);

                if (user != null)
                {
                    user.Delete();
                    //Controllers.LogFiles.LogApplication.LogWrite("Delete User ==> " + "Usuario: " + users + ", eliminado");
                    Models.GlobalVar.countYESProcess++;
                }
                else
                {
                    //Controllers.LogFiles.LogApplication.LogWrite("Delete User ==> " + "Usuario: " + users + ", No existe");
                    Models.GlobalVar.countNOProcess++;
                }

                principalContext.Dispose();

            }
            catch (Exception e)
            {

                CrtMessage = e.Message.ToString();

                if (CrtMessage.Contains("Acceso denegado"))
                {
                    MessageBox.Show("El usuario: " + Models.GlobalVar.UserAdm + ", no tiene permisos para modificar el Directorio Activo", "RPAKonecta - Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Models.GlobalVar.deleteEntires = true;
                }
                else if (CrtMessage.Contains("Access denied."))
                {
                    MessageBox.Show("El usuario: " + Models.GlobalVar.UserAdm + ", no tiene permisos para modificar el Directorio Activo", "RPAKonecta - Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Models.GlobalVar.deleteEntires = true;
                }
                else
                {
                    //Controllers.LogFiles.LogApplication.LogWrite("Delete User ==> " + "Error: No se pudo encontrar el usuario");
                    Models.GlobalVar.countNOProcess++;
                }

                principalContext.Dispose();

            }

        }
    }
}
