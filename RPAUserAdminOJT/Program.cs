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

            Initialize();

        }


        public static void Initialize()
        {
            string ceco = string.Empty;
            string cod_pcrc = string.Empty;
            string documento = string.Empty;
            string id_dp_cargos = string.Empty;
            string id_dp_estados = string.Empty;
            string nombre_cargo = string.Empty;
            string nombre_ceco = string.Empty;
            string nombre_completo = string.Empty;
            string nombre_pcrc = string.Empty;
            string tipo_estado = string.Empty;

            Models.GlobalVar.filePath = @"C:\Users\juan.sepulveda.m\Downloads\Respaldo_Download\usersAdmir.csv";

            using (PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, Models.GlobalVar.Domain))
            {
                bool isValid = principalContext.ValidateCredentials(Models.GlobalVar.UserAdm, Models.GlobalVar.Passwrd);

                if (isValid != false)
                {

                    string NToken = Function.ConsultAPI.getToken();
                    IList<Models.ModelExpedientes> exped = Function.ConsultAPI.getExpedientes(NToken);

                    for (int i = 0; i < exped.Count; i++)
                    {

                        try
                        {
                            ceco = string.Empty;
                            cod_pcrc = string.Empty;
                            documento = string.Empty;
                            id_dp_cargos = string.Empty;
                            id_dp_estados = string.Empty;
                            nombre_cargo = string.Empty;
                            nombre_ceco = string.Empty;
                            nombre_completo = string.Empty;
                            nombre_pcrc = string.Empty;
                            tipo_estado = string.Empty;

                            ceco = exped[i].ceco.ToString();
                            cod_pcrc = exped[i].cod_pcrc.ToString();
                            documento = exped[i].documento.ToString();
                            id_dp_cargos = exped[i].id_dp_cargos.ToString();
                            id_dp_estados = exped[i].id_dp_estados.ToString();
                            nombre_cargo = exped[i].nombre_cargo.ToString();
                            nombre_ceco = exped[i].nombre_ceco.ToString();
                            nombre_completo = exped[i].nombre_completo.ToString();
                            nombre_pcrc = exped[i].nombre_pcrc.ToString();
                            tipo_estado = exped[i].tipo_estado.ToString();

                            if (id_dp_estados == "309")
                            {
                                Models.ModelExpedientes.candidatForma.Add(ceco + ";" + cod_pcrc + ";" + documento + ";" + id_dp_cargos + ";" + id_dp_estados + ";" + nombre_cargo + ";" + nombre_ceco + ";" + nombre_completo + ";" + nombre_pcrc + ";" + tipo_estado);
                            }
                            else if (id_dp_estados == "301")
                            {
                                Models.ModelExpedientes.candidatOJT.Add(ceco + ";" + cod_pcrc + ";" + documento + ";" + id_dp_cargos + ";" + id_dp_estados + ";" + nombre_cargo + ";" + nombre_ceco + ";" + nombre_completo + ";" + nombre_pcrc + ";" + tipo_estado);
                            }
                            else if (id_dp_estados == "317")
                            {
                                Models.ModelExpedientes.candidatRechz.Add(ceco + ";" + cod_pcrc + ";" + documento + ";" + id_dp_cargos + ";" + id_dp_estados + ";" + nombre_cargo + ";" + nombre_ceco + ";" + nombre_completo + ";" + nombre_pcrc + ";" + tipo_estado);
                            }

                        }
                        catch
                        {

                        }

                    }

                    //Models.ModelExpedientes.candidatForma.ToString();
                    //Models.ModelExpedientes.candidatOJT.ToString();
                    //Models.ModelExpedientes.candidatRechz.ToString();

                    //ValidateFormaUsers();
                    ValidateOJTUsers();
                    //ValidateRECHZUsers();

                }
                else
                {
                    MessageBox.Show("Las credenciales ingresadas no son validas, por favor verifique e intente nuevamente", "RPAKonecta - Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                principalContext.Dispose();
            }
        }


        public static void ValidateFormaUsers()
        {

            ArrayList formaList = new ArrayList();
            string resultForma = string.Empty;

            for (int i = 0; i < Models.ModelExpedientes.candidatForma.Count; i++)
            {
                string forma = Models.ModelExpedientes.candidatForma[i].ToString();
                string[] SPLTforma = forma.Split(';');

                int yillFiles = SPLTforma.Length;

                if (yillFiles == 10)
                {
                    for (int j = 0; j < SPLTforma.Count(); j++)
                    {
                        resultForma = SPLTforma[j].ToString();
                        formaList.Add(resultForma);
                    }
                    formaList.ToString();
                }

                formaList.Clear();
            }

        }


        public static void ValidateOJTUsers()
        {

            ArrayList OJTList = new ArrayList();
            string resultOJT = string.Empty;

            for (int i = 0; i < Models.ModelExpedientes.candidatOJT.Count; i++)
            {
                string ojt = Models.ModelExpedientes.candidatOJT[i].ToString();
                string[] SPLTOjt = ojt.Split(';');

                int yillFiles = SPLTOjt.Length;

                if (yillFiles == 10)
                {
                    for (int j = 0; j < SPLTOjt.Count(); j++)
                    {
                        resultOJT = SPLTOjt[j].ToString();
                        OJTList.Add(resultOJT);
                    }
                    OJTList.ToString();
                }

                OJTList.Clear();
            }

        }


        public static void ValidateRECHZUsers()
        {

            ArrayList RECHZList = new ArrayList();
            string resultRECHZ = string.Empty;

            for (int i = 0; i < Models.ModelExpedientes.candidatRechz.Count; i++)
            {
                string rechz = Models.ModelExpedientes.candidatRechz[i].ToString();
                string[] SPLTRechz = rechz.Split(';');

                int yillFiles = SPLTRechz.Length;

                if (yillFiles == 10)
                {
                    for (int j = 0; j < SPLTRechz.Count(); j++)
                    {
                        resultRECHZ = SPLTRechz[j].ToString();
                        RECHZList.Add(resultRECHZ);
                    }
                    RECHZList.ToString();
                }

                RECHZList.Clear();
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