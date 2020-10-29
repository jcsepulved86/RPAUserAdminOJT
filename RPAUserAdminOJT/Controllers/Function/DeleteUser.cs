using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Windows.Forms;

namespace RPAUserAdminOJT.Controllers.Function
{
    public class Delete
    {
        

        public static PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, Models.GlobalVar.Domain, Models.GlobalVar.UserAdm, Models.GlobalVar.Passwrd);



        public static string Cedula(string Cedula)
        {
            string userRED = string.Empty;

            try
            {
                string connectionPrefix = $"LDAP://{Models.GlobalVar.Domain}";

                DirectoryEntry ldapConnection = new DirectoryEntry(connectionPrefix, Models.GlobalVar.UserAdm, Models.GlobalVar.Passwrd);
                
                ldapConnection.Path = connectionPrefix;
                ldapConnection.AuthenticationType = AuthenticationTypes.Secure;

                DirectoryEntry myLdapConnection = ldapConnection;

                DirectorySearcher ds = new DirectorySearcher(myLdapConnection);

                ds.Filter = "(&((&(objectCategory=Person)(objectClass=User)))(postOfficeBox=" + Cedula + "))";

                ds.SearchScope = SearchScope.Subtree;

                SearchResult rs = ds.FindOne();

                if (rs.GetDirectoryEntry().Properties["samaccountname"].Value != null)
                {
                    userRED =  rs.GetDirectoryEntry().Properties["samaccountname"].Value.ToString();
                }

                return userRED;

            }

            catch (Exception e)
            {
                Console.WriteLine("Exception caught:\n\n" + e.ToString());
                return userRED;
            }
        }



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
