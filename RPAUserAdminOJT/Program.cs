using System;
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RPAUserAdminOJT.Controllers
{
    public class Program
    {


        public static void Main(string[] args)
        {
            Models.GlobalVar.filePath = @"C:\Users\juan.sepulveda.m\Downloads\Respaldo_Download\usersAdmir.csv";

            using (PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, Models.GlobalVar.Domain))
            {
                bool isValid = principalContext.ValidateCredentials(Models.GlobalVar.UserAdm, Models.GlobalVar.Passwrd);

                if (isValid != false)
                {
                    //Create();
                    //Deleted();

                    string NToken = Function.ConsultAPI.getToken();
                    Function.ConsultAPI.getExpedientes(NToken);

                }
                else
                {
                    MessageBox.Show("Las credenciales ingresadas no son validas, por favor verifique e intente nuevamente", "RPAKonecta - Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                principalContext.Dispose();
            }

        }


        public static void Create()
        {

            string result = string.Empty;
            string[] filelines = File.ReadAllLines(Models.GlobalVar.filePath);
            ArrayList employeeList = new ArrayList();
            Models.Employee employee = new Models.Employee();

            Models.GlobalVar.outLoadFile = false;

            for (int a = 1; a < filelines.Length; a++)
            {
                string[] fillines = filelines[a].Split(',');
                int yillFiles = fillines.Length;

                if (yillFiles == 16)
                {
                    for (int i = 0; i < fillines.Length; i++)
                    {
                        result = fillines[i].ToString();
                        employeeList.Add(result);
                    }

                    employee.SamAccountName = employeeList[0].ToString();
                    employee.GivenName = employeeList[1].ToString();
                    employee.MiddleName = employeeList[2].ToString();
                    employee.FirstLastName = employeeList[3].ToString();
                    employee.SecondLastName = employeeList[4].ToString();
                    employee.PostOfficeBox = employeeList[5].ToString();
                    employee.UserBasement = employeeList[6].ToString();
                    employee.TicketID = employeeList[7].ToString();
                    employee.HomePage = employeeList[8].ToString();
                    employee.City = employeeList[9].ToString();
                    employee.State = employeeList[10].ToString();
                    employee.Country = employeeList[11].ToString();
                    employee.Description = employeeList[12].ToString();
                    employee.Domain = employeeList[13].ToString();
                    employee.Department = employeeList[14].ToString();
                    employee.Company = employeeList[15].ToString();

                    if (Models.GlobalVar.desactiveAccount == true)
                    {
                        ActiveDirectory.Services.ServicesProvider.DisableAccount(employee.SamAccountName);
                        Models.GlobalVar.countYESProcess++;
                    }
                    else
                    {
                        Function.Create.User(employee.FirstLastName, employee.SecondLastName,
                                    employee.GivenName, employee.MiddleName, employee.Domain,
                                    employee.HomePage, employee.Country, employee.State, employee.City, employee.PostOfficeBox, employee.SamAccountName,
                                    employee.Description, employee.Department, employee.Company, employee.UserBasement);

                        if (Models.GlobalVar.outLoadFile == true)
                        {
                            break;
                        }

                        if (Models.GlobalVar.NeverPswd == true)
                        {
                            ActiveDirectory.Services.ServicesProvider.PasswordNeverExpires(employee.SamAccountName);
                        }

                        if (Models.GlobalVar.addUserGroup == true)
                        {
                            ActiveDirectory.Services.ServicesProvider.User_Add_Group(employee.SamAccountName, employee.UserBasement);
                        }
                    }

                    employeeList.Clear();
                }
                else
                {
                    break;
                }

            }

            MessageBox.Show("Datos Creados: " + Models.GlobalVar.countYESProcess.ToString() + "\n" + "Datos No Creados: " + Models.GlobalVar.countNOProcess.ToString());
        }


        public static void Deleted()
        {
            int qtyFile = 0;
            Models.GlobalVar.countYESProcess = 0;
            Models.GlobalVar.countNOProcess = 0;

            string result = string.Empty;
            string[] filelines = File.ReadAllLines(Models.GlobalVar.filePath);

            ArrayList employeeList = new ArrayList();
            Controllers.Models.Employee employee = new Controllers.Models.Employee();

            qtyFile = filelines.Count();

            Models.GlobalVar.deleteEntires = false;

            for (int a = 1; a < filelines.Length; a++)
            {
                string[] fillines = filelines[a].Split(',');
                int yillFiles = fillines.Length;

                if (yillFiles == 16)
                {
                    for (int i = 0; i < fillines.Length; i++)
                    {
                        result = fillines[i].ToString();
                        employeeList.Add(result);
                    }

                    employee.SamAccountName = employeeList[0].ToString();
                    employee.GivenName = employeeList[1].ToString();
                    employee.MiddleName = employeeList[2].ToString();
                    employee.FirstLastName = employeeList[3].ToString();
                    employee.SecondLastName = employeeList[4].ToString();
                    employee.PostOfficeBox = employeeList[5].ToString();
                    employee.UserBasement = employeeList[6].ToString();
                    employee.TicketID = employeeList[7].ToString();
                    employee.HomePage = employeeList[8].ToString();
                    employee.City = employeeList[9].ToString();
                    employee.State = employeeList[10].ToString();
                    employee.Country = employeeList[11].ToString();
                    employee.Description = employeeList[12].ToString();
                    employee.Domain = employeeList[13].ToString();
                    employee.Department = employeeList[14].ToString();
                    employee.Company = employeeList[15].ToString();

                    Function.Delete.User(employee.SamAccountName);

                    if (Models.GlobalVar.deleteEntires == true)
                    {
                        //Controllers.LogFiles.LogApplication.LogWrite("Delete User ==> " + "El usuario: " + Controllers.Models.globalvar.useradm + ", no tiene permisos para modificar el Directorio Activo");
                        break;
                    }

                    employeeList.Clear();
                }
                else
                {
                    break;
                }

            }

            MessageBox.Show("Datos Eliminados: " + Models.GlobalVar.countYESProcess.ToString() + "\n" + "Datos No Eliminados: " + Models.GlobalVar.countNOProcess.ToString());
        }


    }
}