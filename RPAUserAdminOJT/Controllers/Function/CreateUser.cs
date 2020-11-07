using System;
using System.DirectoryServices.AccountManagement;
using System.Windows.Forms;

namespace RPAUserAdminOJT.Controllers.Function
{
    public class Create
    {
        public static PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, Models.GlobalVar.Domain, Models.GlobalVar.UserAdm, Models.GlobalVar.Passwrd);

        public static void User(string FirstLastName, string SecondLastName, string GivenName, string MiddleName, string Domain, string HomePage, string Country, string State, string City, string PostOfficeBox, string SamAccountName, string Description, string Department, string Company, string UserBasement, string DateModify, string NameBox, string EmailBox, string IdBox, string CodPRC, string client)
        {
            
            string result;
            try
            {

                ActiveDirectory.Query.QueryProvider queryProvider = new ActiveDirectory.Query.QueryProvider(principalContext);
                string ldapPath = queryProvider.GetDistinguishedName(UserBasement);

                Models.AdUser user = new Models.AdUser
                {
                    FirsLastName = FirstLastName,
                    SecondLastName = SecondLastName,
                    GivenName = GivenName,
                    MiddleName = MiddleName,
                    Domain = Domain,
                    HomePage = HomePage,
                    Country = Country,
                    State = State,
                    City = City,
                    PostOfficeBox = PostOfficeBox,
                    SamAccountName = SamAccountName,
                    Description = Description,
                    Department = Department,
                    Company = Company,
                    DistinguishedName = ldapPath,
                    Password = Utility.PasswordGenerator.GetSecurePassword()
                };

                Models.GlobalVar.passwordUser = user.Password.ToString();

                ActiveDirectory.Services.ServicesProvider servicesProvider = new ActiveDirectory.Services.ServicesProvider(principalContext, Models.GlobalVar.Passwrd);
                bool validUser = servicesProvider.Create(user, true);

                if (validUser == false)
                {

                    if (Models.GlobalVar.accessEntires == false)
                    {
                        //MessageBox.Show("El usuario: " + Models.GlobalVar.UserAdm + ", no tiene permisos para modificar el Directorio Activo", "RPAKonecta - Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Utility.LogApplication.LogWrite("User Create ==> " + "El usuario: " + Models.GlobalVar.UserAdm + ", no tiene permisos para modificar el Directorio Activo");
                        //Utility.FillExcel.WriteExcel("User Create", Models.GlobalVar.UserAdm, "no tiene permisos para modificar el Directorio Activo");
                        Models.GlobalVar.outLoadFile = true;
                        Models.GlobalVar.existUser = false;

                    }
                    else
                    {
                        Utility.LogApplication.LogWrite("User Create ==> " + "Usuario: " + user.SamAccountName + ", ya existe");
                        //Utility.FillExcel.WriteExcel("User Create", user.SamAccountName, "ya existe");
                        Models.GlobalVar.countNOProcess++;
                        Models.GlobalVar.existUser = false;
                    }

                }
                else
                {
                    Utility.LogApplication.LogWrite("User Create ==> " + "Usuario: " + user.SamAccountName);
                    Utility.FillExcel.WriteExcel("User Create", DateModify, NameBox, EmailBox, IdBox, GivenName + MiddleName + FirstLastName + SecondLastName, client,  CodPRC, user.SamAccountName, Models.GlobalVar.passwordUser);
                    Models.GlobalVar.countYESProcess++;
                    Models.GlobalVar.existUser = true;
                }

                principalContext.Dispose();

            }
            catch (Exception ex)
            {
                result = ex.Message;
                principalContext.Dispose();
            }

        }
    }
}
