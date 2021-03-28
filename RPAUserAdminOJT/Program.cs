#region USINGS
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
#endregion

namespace RPAUserAdminOJT.Controllers
{
    public class Program
    {
        #region STATIC VARIABLE
        public static string firstNAME = string.Empty;
        public static string secondNAME = string.Empty;
        public static string lastNAME = string.Empty;
        public static string secondlastNAME = string.Empty;
        public static Dictionary<string, string> DCTbassement = new Dictionary<string, string>();
        public static string rootUBassemet = ConfigurationManager.AppSettings["rootUBassemet"];
        public static string GUID = Guid.NewGuid().ToString("N");
        public static DateTime TiempoInicial = DateTime.Now;
        #endregion

        #region MAIN ACTIVITIES
        public static void Main(string[] args)
        {

            int procCount = 0;
            foreach (Process pp in Process.GetProcesses())
            {
                try
                {
                    if (String.Compare(pp.MainModule.FileName, Application.ExecutablePath, true) == 0)
                    {
                        procCount++;
                        if (procCount > 1)
                        {
                            Application.Exit();
                            return;
                        }
                    }
                }
                catch (Exception)
                {
                    //MessageBox.Show("1Exception: " + ex.Message.ToString(), "RPA RobotCopy", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            #region HIDDEN
            //Function.Delete.User("robotica.testa");
            #endregion

            LOGRobotica.Controllers.LogWebServices.logsWS(TiempoInicial, GUID, "Inicia proceso de Creacion de Usuario OJT", "Consulta Exitosa");
            Utility.LogApplication.LogWrite("Main ==> " + "Inicia el proceso de automatizado para la creación de cuentas de usuario");

            infintyWhile();

        }
        #endregion

        #region INFINITY WHILE
        public static void infintyWhile()
        {

            while (true)
            {
                Models.GlobalVar.TimeLogs = DateTime.Now;

                CultureInfo ci = CultureInfo.CreateSpecificCulture(CultureInfo.CurrentCulture.Name); ci.DateTimeFormat.ShortDatePattern = "dd-MM-yyyy";
                Thread.CurrentThread.CurrentCulture = ci;

                string hours = DateTime.Now.ToString("HH");

                DateTime input = DateTime.ParseExact(hours, "HH", null);
                DateTime date1 = DateTime.ParseExact("05", "HH", null);
                DateTime date2 = DateTime.ParseExact("23", "HH", null);

                bool snooze = IsBetween(input, date1, date2);

                if (snooze != false)
                {
                    Initialize();

                    LOGRobotica.Controllers.LogWebServices.logsWS(TiempoInicial, GUID, "Finaliza proceso de Creacion de Usuario OJT", "Consulta Exitosa", "Creados", Models.GlobalVar.countYESProcess.ToString(), "NoCreados", Models.GlobalVar.countNOProcess.ToString(), "", "Deshabilitado", Models.GlobalVar.CountDeshabilitadoYESProcess.ToString(), "NoDeshabilitado", Models.GlobalVar.CountDeshabilitadoNOProcess.ToString());
                    Utility.LogApplication.LogWrite("Main ==> " + "Finaliza el proceso de automatizado para la creación de cuentas de usuario");

                    Utility.SendMail.SendingMessage("gestion_usuarios@grupokonecta.com", 3);
                    //Utility.SendMail.SendingMessage("jcsepulveda@grupokonecta.com", 3);
                }

                Thread.Sleep(3600000);
                //Thread.Sleep(60000);

            }

        }
        #endregion

        #region COMPARE BETWEEN A RANGE HOURS
        public static bool IsBetween(DateTime input, DateTime date1, DateTime date2)
        {
            try
            {
                return (input > date1 && input < date2);
            }
            catch (Exception ex)
            {
                LOGRobotica.Controllers.LogApplication.LogWrite("IsBetween" + " : " + "Error: " + ex.Message);
                return false;
            }
        }
        #endregion

        #region INITIALIZE
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
            string documento_jefe = string.Empty;
            string nombre_jefe = string.Empty;
            string email_jefe = string.Empty;
            string fecha_modificacion = string.Empty;
            string numero_lote = string.Empty;

            Models.GlobalVar.filePath = rootUBassemet;

            string result = string.Empty;
            string[] filelines = File.ReadAllLines(Models.GlobalVar.filePath);
            ArrayList UBassement = new ArrayList();
            Models.Employee employee = new Models.Employee();

            for (int a = 1; a < filelines.Length; a++)
            {
                string[] fillines = filelines[a].Split(',');
                int yillFiles = fillines.Length;

                if (yillFiles == 12)
                {
                    try
                    {
                        DCTbassement.Add(fillines[3].ToString().Trim(), fillines[1].ToString() +";"+ fillines[6].ToString() +";"+ fillines[7].ToString() + ";" + fillines[8].ToString() + ";" + fillines[9].ToString() + ";" + fillines[10].ToString() + ";" + fillines[11].ToString());
                    }
                    catch { }

                }
            }

            using (PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, Models.GlobalVar.Domain))
            {
                bool isValid = principalContext.ValidateCredentials(Models.GlobalVar.UserAdm, Models.GlobalVar.Passwrd);

                if (isValid != false)
                {

                    string NToken = Utility.ConsultAPI.getToken();
                    IList<Models.ModelExpedientes> exped = Utility.ConsultAPI.getExpedientes(NToken);

                    for (int i = 0; i < exped.Count; i++)
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
                        documento_jefe = string.Empty;
                        nombre_jefe = string.Empty;
                        email_jefe = string.Empty;
                        fecha_modificacion = string.Empty;
                        numero_lote = string.Empty;


                        if (exped[i].ceco == null)
                            ceco = "";
                        else
                            ceco = exped[i].ceco.ToString();

                        if (exped[i].cod_pcrc == null)
                            cod_pcrc = "";
                        else
                            cod_pcrc = exped[i].cod_pcrc.ToString();

                        if (exped[i].documento == null)
                            documento = "";
                        else
                            documento = exped[i].documento.ToString();

                        if (exped[i].id_dp_cargos == null)
                            id_dp_cargos = "";
                        else
                            id_dp_cargos = exped[i].id_dp_cargos.ToString();

                        if (exped[i].id_dp_estados == null)
                            id_dp_estados = "";
                        else
                            id_dp_estados = exped[i].id_dp_estados.ToString();

                        if (exped[i].nombre_cargo == null)
                            nombre_cargo = "";
                        else
                            nombre_cargo = exped[i].nombre_cargo.ToString();

                        if (exped[i].nombre_ceco == null)
                            nombre_ceco = "";
                        else
                            nombre_ceco = exped[i].nombre_ceco.ToString();

                        if (exped[i].nombre_completo == null)
                            nombre_completo = "";
                        else
                            nombre_completo = exped[i].nombre_completo.ToString();

                        if (exped[i].nombre_pcrc == null)
                            nombre_pcrc = "";
                        else
                            nombre_pcrc = exped[i].nombre_pcrc.ToString();

                        if (exped[i].tipo_estado == null)
                            tipo_estado = "";
                        else
                            tipo_estado = exped[i].tipo_estado.ToString();

                        if (exped[i].documento_jefe == null)
                            documento_jefe = "";
                        else
                            documento_jefe = exped[i].documento_jefe.ToString();

                        if (exped[i].nombre_jefe == null)
                            nombre_jefe = "";
                        else
                            nombre_jefe = exped[i].nombre_jefe.ToString();

                        if (exped[i].email_jefe == null)
                            email_jefe = "";
                        else
                            email_jefe = exped[i].email_jefe.ToString();

                        if (exped[i].fecha_modificacion == null)
                            fecha_modificacion = "";
                        else
                            fecha_modificacion = exped[i].fecha_modificacion.ToString();

                        if (exped[i].id_dp_solicitudes == null)
                            numero_lote = "";
                        else
                            numero_lote = exped[i].id_dp_solicitudes.ToString();


                        if (id_dp_estados == "309" || id_dp_estados == "310" || id_dp_estados == "311")
                        {
                            Models.ModelExpedientes.candidatBeginner.Add(ceco + ";" + cod_pcrc + ";" + documento + ";" + id_dp_cargos + ";" + id_dp_estados + ";" + nombre_cargo + ";" + nombre_ceco + ";" + nombre_completo + ";" + nombre_pcrc + ";" + tipo_estado + ";" + documento_jefe + ";" + nombre_jefe + ";" + email_jefe + ";" + fecha_modificacion + ";" + numero_lote);
                        }
                        else if (id_dp_estados == "305" || id_dp_estados == "317" || id_dp_estados == "327")
                        {
                            Models.ModelExpedientes.candidatMovement.Add(ceco + ";" + cod_pcrc + ";" + documento + ";" + id_dp_cargos + ";" + id_dp_estados + ";" + nombre_cargo + ";" + nombre_ceco + ";" + nombre_completo + ";" + nombre_pcrc + ";" + tipo_estado + ";" + documento_jefe + ";" + nombre_jefe + ";" + email_jefe + ";" + fecha_modificacion + ";" + numero_lote);
                        }
                        else if (id_dp_estados == "300")
                        {
                            Models.ModelExpedientes.candidatDelete.Add(ceco + ";" + cod_pcrc + ";" + documento + ";" + id_dp_cargos + ";" + id_dp_estados + ";" + nombre_cargo + ";" + nombre_ceco + ";" + nombre_completo + ";" + nombre_pcrc + ";" + tipo_estado + ";" + documento_jefe + ";" + nombre_jefe + ";" + email_jefe + ";" + fecha_modificacion + ";" + numero_lote);
                        }

                    }

                    Models.ModelExpedientes.candidatBeginner.ToString();
                    Models.ModelExpedientes.candidatMovement.ToString();
                    Models.ModelExpedientes.candidatDelete.ToString();

                    if (Models.ModelExpedientes.candidatBeginner.Count != 0)
                    {
                        ValidateBeginner();
                    }

                    if (Models.ModelExpedientes.candidatMovement.Count != 0)
                    {
                        ValidateMovement();
                    }

                    if (Models.ModelExpedientes.candidatDelete.Count != 0)
                    {
                        ValidateDeleting();
                    }


                    if (Models.ModelExpedientes.candidatBeginner.Count != 0 || Models.ModelExpedientes.candidatMovement.Count != 0 || Models.ModelExpedientes.candidatDelete.Count != 0)
                    {
                        Utility.SendMail.SendingMessage("gestion_usuarios@grupokonecta.com", 1);
                        //Utility.SendMail.SendingMessage("jcsepulveda@grupokonecta.com", 1);
                        Utility.SendMail.DataAnalytics();
                    }

                    Models.ModelExpedientes.candidatBeginner.Clear();
                    Models.ModelExpedientes.candidatMovement.Clear();
                    Models.ModelExpedientes.candidatDelete.Clear();

                }
                else
                {
                    Utility.LogApplication.LogWrite("Initialize ==> " + "El usuario: " + Models.GlobalVar.UserAdm + ", las credenciales ingresadas no son validas, por favor verifique e intente nuevamente");
                }
                principalContext.Dispose();
            }
        }
        #endregion

        #region VALIDATE FORMATION USERS
        public static void ValidateBeginner()
        {
            try
            {
                Models.Employee employee = new Models.Employee();

                ArrayList beginnerList = new ArrayList();
                string resultBeginner = string.Empty;

                for (int i = 0; i < Models.ModelExpedientes.candidatBeginner.Count; i++)
                {
                    string beginner = Models.ModelExpedientes.candidatBeginner[i].ToString();
                    string[] SPLTbeginner = beginner.Split(';');

                    int yillFiles = SPLTbeginner.Length;

                    if (yillFiles == 15)
                    {
                        for (int j = 0; j < SPLTbeginner.Count(); j++)
                        {
                            resultBeginner = SPLTbeginner[j].ToString();
                            beginnerList.Add(resultBeginner);
                        }

                        beginnerList.ToString();

                        string nombre = beginnerList[7].ToString();
                        string usuarioRed = string.Empty;
                        string dictionaryArray = string.Empty;
                        string[] splBassment;

                        if (beginnerList[1].ToString() != "0")
                        {
                            try
                            {
                                dictionaryArray = DCTbassement[beginnerList[1].ToString()];
                                splBassment = dictionaryArray.Split(';');
                            }
                            catch(Exception)
                            {
                                Utility.LogApplication.LogWrite("ValidateBeginner ==> " + "El nombre: " + beginnerList[7].ToString() + ", no se encuentra codigo PCRC: " + beginnerList[1].ToString() + ", en la lista de usuarios base");
                                beginnerList.Clear();
                                continue;
                            }
                            
                            usuarioRed = NetworkAccountName(nombre, 0);

                            employee.SamAccountName = usuarioRed;
                            employee.GivenName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(firstNAME);
                            employee.MiddleName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(secondNAME);
                            employee.FirstLastName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lastNAME);
                            employee.SecondLastName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(secondlastNAME);
                            employee.PostOfficeBox = beginnerList[2].ToString();

                            if (beginnerList[1].ToString() == "0")
                            {
                                //employee.UserBasement = "pruebas.bot.1";
                                Utility.LogApplication.LogWrite("Usuario Base ==> " + "El usuario: " + usuarioRed + ", no se encontro usuario base");

                                continue;

                            }
                            else
                            {
                                employee.UserBasement = splBassment[0].ToString();
                            }

                            employee.numero_lote = beginnerList[14].ToString();
                            employee.HomePage = splBassment[4].ToString();
                            employee.City = splBassment[1].ToString();
                            employee.State = splBassment[2].ToString();
                            employee.Country = splBassment[3].ToString();
                            employee.Description = beginnerList[6].ToString();
                            employee.Domain = splBassment[5].ToString();
                            employee.Department = splBassment[6].ToString();
                            employee.Company = "Konecta";
                            employee.fecha_modificacion = beginnerList[13].ToString();
                            employee.nombre_jefe = beginnerList[11].ToString();
                            employee.email_jefe = beginnerList[12].ToString();
                            employee.documento_jefe = beginnerList[10].ToString();
                            employee.codPCR = beginnerList[1].ToString();
                            employee.cliente = beginnerList[8].ToString();


                            string usred = Controllers.ActiveDirectory.Services.ServicesProvider.UserNetwork(beginnerList[2].ToString());

                            if (usred == "")
                            {
                                Function.Create.User(employee.FirstLastName, employee.SecondLastName,
                                        employee.GivenName, employee.MiddleName, employee.Domain,
                                        employee.HomePage, employee.Country, employee.State, employee.City, employee.PostOfficeBox, employee.SamAccountName,
                                        employee.Description, employee.Department, employee.Company, employee.UserBasement, employee.fecha_modificacion, employee.nombre_jefe, employee.email_jefe, employee.documento_jefe, employee.codPCR, employee.cliente, employee.numero_lote);

                                if (Models.GlobalVar.existUser != false)
                                {

                                    ActiveDirectory.Services.ServicesProvider.User_Add_Group(employee.SamAccountName, employee.UserBasement);

                                    //if (employee.UserBasement != "pruebas.bot.1")
                                    //{
                                    //    ActiveDirectory.Services.ServicesProvider.User_Add_Group(employee.SamAccountName, employee.UserBasement);
                                    //}
                                    //else
                                    //{
                                    //    Utility.LogApplication.LogWrite("Sin Usuario Base ==> " + "El usuario: " + usuarioRed + ", imposible asignarle grupo");
                                    //}

                                    if (usuarioRed != "")
                                    {
                                        ActiveDirectory.Services.ServicesProvider.EnableAccount(usuarioRed);
                                    }


                                }
                            }
                            else
                            {
                                Utility.LogApplication.LogWrite("ValidateBeginner ==> " + "El nombre: " + usred + ", ya existe en la organizacion");
                                Utility.FillExcel.WriteExcel("User No Create", employee.fecha_modificacion, employee.numero_lote, employee.nombre_jefe, employee.email_jefe, employee.documento_jefe, beginnerList[7].ToString(), beginnerList[2].ToString(), employee.cliente, employee.codPCR, employee.UserBasement );
                            }
                        }
                        else
                        {
                            Utility.LogApplication.LogWrite("ValidateBeginner ==> " + "El nombre: " + beginnerList[7].ToString() + ", no posee codigo PCRC");
                        }
                    }

                    beginnerList.Clear();
                }

                Utility.LogApplication.LogWrite("ValidateBeginner ==> " + "Finaliza proceso de creación de cuentas de usuarios");

            }
            catch(Exception ex)
            {
                Utility.LogApplication.LogWrite("ValidateBeginner ==> " + "Exception: " + ex.Message.ToString());
            }
        }
        #endregion

        #region VALIDATE MOVEMENT USERS
        public static void ValidateMovement()
        {
            try
            {
                ArrayList MovementList = new ArrayList();
                string resultMovement = string.Empty;

                for (int i = 0; i < Models.ModelExpedientes.candidatMovement.Count; i++)
                {
                    string movement = Models.ModelExpedientes.candidatMovement[i].ToString();
                    string[] SPLTMovement = movement.Split(';');

                    int yillFiles = SPLTMovement.Length;

                    if (yillFiles == 15)
                    {
                        for (int j = 0; j < SPLTMovement.Count(); j++)
                        {
                            resultMovement = SPLTMovement[j].ToString();
                            MovementList.Add(resultMovement);
                        }

                        MovementList.ToString();

                        string usred = Controllers.ActiveDirectory.Services.ServicesProvider.UserNetwork(MovementList[2].ToString());

                        if (usred != "")
                        {
                            string distinguishedname = Controllers.ActiveDirectory.Services.ServicesProvider.DistGuiName(MovementList[2].ToString());
                            ActiveDirectory.Services.ServicesProvider.DisableAccount(usred, distinguishedname);
                            //Models.GlobalVar.CountDeshabilitadoYESProcess++;
                        }
                        else
                        {
                            Utility.LogApplication.LogWrite("ValidateMovement - Type: " + MovementList[9].ToString() + " ==> " + "State: " + MovementList[4].ToString() + ", FullName: " + MovementList[7].ToString() + ", no posee codigo PCRC");
                            //Models.GlobalVar.CountDeshabilitadoNOProcess++;
                        }
                    }
                    MovementList.Clear();
                }

                Utility.LogApplication.LogWrite("ValidateMovement ==> " + "Finaliza proceso de movimiento de unidad organizacional para cuentas de usuarios");
            }
            catch (Exception ex)
            {
                Utility.LogApplication.LogWrite("ValidateMovement ==> " + "Exception: " + ex.Message.ToString());
            }
        }
        #endregion

        #region VALIDATE DELETING USERS
        public static void ValidateDeleting()
        {
            try
            {

                ArrayList deleteList = new ArrayList();
                string resultDelete = string.Empty;

                for (int i = 0; i < Models.ModelExpedientes.candidatDelete.Count; i++)
                {
                    string delete = Models.ModelExpedientes.candidatDelete[i].ToString();
                    string[] SPLTDelete = delete.Split(';');

                    int yillFiles = SPLTDelete.Length;

                    if (yillFiles == 15)
                    {
                        for (int j = 0; j < SPLTDelete.Count(); j++)
                        {
                            resultDelete = SPLTDelete[j].ToString();
                            deleteList.Add(resultDelete);
                        }

                        deleteList.ToString();

                        string usred = Controllers.ActiveDirectory.Services.ServicesProvider.UserNetwork(deleteList[2].ToString());

                        if (usred != "")
                        {
                            Function.Delete.User(usred);
                        }
                        else
                        {
                            Utility.LogApplication.LogWrite("ValidateDeleting ==> " + "No se encontro el usuario de red debido a que no existe en la organización");
                        }
                    }
                    deleteList.Clear();
                }

                Utility.LogApplication.LogWrite("ValidateDeleting ==> " + "Finaliza proceso de elimanción de cuentas de usuarios");

            }
            catch (Exception ex)
            {
                Utility.LogApplication.LogWrite("ValidateDeleting ==> " + "Exception: " + ex.Message.ToString());
            }
        }
        #endregion

        #region CREATE NETWORK ACCOUNT NAME
        public static string NetworkAccountName(string REDUser, int vState)
        {
            string validAccountName = string.Empty;
            Regex reg = new Regex("[^a-zA-Z0-9 ]");

            try
            {
                string[] SPLName = REDUser.Split(' ');


                if (vState == 0)
                {
                    firstNAME = SPLName[0].ToString().ToLower();
                    string firstNormalizado = firstNAME.Normalize(NormalizationForm.FormD);
                    firstNAME = reg.Replace(firstNormalizado, "");

                    secondNAME = SPLName[1].ToString().ToLower();
                    string secondNormalizado = secondNAME.Normalize(NormalizationForm.FormD);
                    secondNAME = reg.Replace(secondNormalizado, "");

                    lastNAME = SPLName[2].ToString().ToLower();
                    string thirdNormalizado = lastNAME.Normalize(NormalizationForm.FormD);
                    lastNAME = reg.Replace(thirdNormalizado, "");

                    secondlastNAME = SPLName[3].ToString().ToLower();
                    string quarterNormalizado = secondlastNAME.Normalize(NormalizationForm.FormD);
                    secondlastNAME = reg.Replace(quarterNormalizado, "");

                }
                else
                {

                    firstNAME = SPLName[2].ToString().ToLower();
                    string firstNormalizado = firstNAME.Normalize(NormalizationForm.FormD);
                    firstNAME = reg.Replace(firstNormalizado, "");

                    secondNAME = SPLName[3].ToString().ToLower();
                    string secondNormalizado = secondNAME.Normalize(NormalizationForm.FormD);
                    secondNAME = reg.Replace(secondNormalizado, "");

                    lastNAME = SPLName[0].ToString().ToLower();
                    string thirdNormalizado = lastNAME.Normalize(NormalizationForm.FormD);
                    lastNAME = reg.Replace(thirdNormalizado, "");

                    secondlastNAME = SPLName[1].ToString().ToLower();
                    string quarterNormalizado = secondlastNAME.Normalize(NormalizationForm.FormD);
                    secondlastNAME = reg.Replace(quarterNormalizado, "");

                }
                
                bool flag = false;

                validAccountName = firstNAME + "." + lastNAME;

                using (PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, "multienlace.com.co"))
                {
                    UserPrincipal user = UserPrincipal.FindByIdentity(principalContext, validAccountName);

                    if (user == null)
                    {
                        return validAccountName;
                    }
                    else
                    {
                        if (secondlastNAME != "")
                        {
                            for (int i = 1; i <= 3; i++)
                            {
                                validAccountName = firstNAME + "." + lastNAME + "." + secondlastNAME.Substring(0, i);

                                user = UserPrincipal.FindByIdentity(principalContext, validAccountName);

                                if (user == null)
                                {
                                    flag = true;
                                    break;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            if (!flag)
                            {
                                validAccountName = firstNAME + "." + lastNAME.Substring(0, lastNAME.Length - 1);
                                user = UserPrincipal.FindByIdentity(principalContext, validAccountName);

                                if (user != null)
                                {
                                    Utility.LogApplication.LogWrite("NetworkAccountName ==> " + "No existe combinacion valida para crear el usuario de red de: " + user);
                                }
                            }
                        }
                        else
                        {
                            if (secondNAME != "")
                            {
                                validAccountName = firstNAME + "." + lastNAME + "." + secondNAME.Substring(0, 1);
                            }
                            else
                            {
                                validAccountName = firstNAME + "." + lastNAME.Substring(0, lastNAME.Length - 1);
                            }


                            user = UserPrincipal.FindByIdentity(principalContext, validAccountName);

                            if (user != null)
                            {
                                Utility.LogApplication.LogWrite("NetworkAccountName ==> " + "No existe combinacion valida para crear el usuario de red de: " + user);
                            }
                        }
                    }

                    principalContext.Dispose();
                    return validAccountName;
                }
            }
            catch (Exception ex)
            {
                Utility.LogApplication.LogWrite("NetworkAccountName ==> " + "Exception: " + ex.Message.ToString());
                return validAccountName;
            }

        }
        #endregion

    }
}