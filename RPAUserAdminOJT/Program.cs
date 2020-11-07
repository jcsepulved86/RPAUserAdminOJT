#region USINGS
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices.AccountManagement;
using System.Globalization;
using System.IO;
using System.Linq;
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
            LOGRobotica.Controllers.LogWebServices.logsWS(TiempoInicial, GUID, "Inicia proceso de Creacion de Usuario OJT", "Consulta Exitosa");
            Initialize();
            LOGRobotica.Controllers.LogWebServices.logsWS(TiempoInicial, GUID, "Finaliza proceso de Creacion de Usuario OJT", "Consulta Exitosa", "Creados", Models.GlobalVar.countYESProcess.ToString(), "NoCreados", Models.GlobalVar.countNOProcess.ToString(), "", "Deshabilitado", Models.GlobalVar.CountDeshabilitadoYESProcess.ToString(), "NoDeshabilitado", Models.GlobalVar.CountDeshabilitadoNOProcess.ToString());
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


                        if (id_dp_estados == "309")
                        {
                            Models.ModelExpedientes.candidatForma.Add(ceco + ";" + cod_pcrc + ";" + documento + ";" + id_dp_cargos + ";" + id_dp_estados + ";" + nombre_cargo + ";" + nombre_ceco + ";" + nombre_completo + ";" + nombre_pcrc + ";" + tipo_estado + ";" + documento_jefe + ";" + nombre_jefe + ";" + email_jefe + ";" + fecha_modificacion);
                        }
                        else if (id_dp_estados == "305" || id_dp_estados == "317" || id_dp_estados == "327")
                        {
                            Models.ModelExpedientes.candidatRechz.Add(ceco + ";" + cod_pcrc + ";" + documento + ";" + id_dp_cargos + ";" + id_dp_estados + ";" + nombre_cargo + ";" + nombre_ceco + ";" + nombre_completo + ";" + nombre_pcrc + ";" + tipo_estado + ";" + documento_jefe + ";" + nombre_jefe + ";" + email_jefe + ";" + fecha_modificacion);
                        }

                    }

                    Models.ModelExpedientes.candidatForma.ToString();
                    Models.ModelExpedientes.candidatRechz.ToString();

                    ValidateFormaUsers();
                    ValidateRECHZUsers();

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
        public static void ValidateFormaUsers()
        {
            try
            {
                Models.Employee employee = new Models.Employee();

                ArrayList formaList = new ArrayList();
                string resultForma = string.Empty;

                for (int i = 0; i < Models.ModelExpedientes.candidatForma.Count; i++)
                {
                    string forma = Models.ModelExpedientes.candidatForma[i].ToString();
                    string[] SPLTforma = forma.Split(';');

                    int yillFiles = SPLTforma.Length;

                    if (yillFiles == 14)
                    {
                        for (int j = 0; j < SPLTforma.Count(); j++)
                        {
                            resultForma = SPLTforma[j].ToString();
                            formaList.Add(resultForma);
                        }

                        formaList.ToString();

                        string nombre = formaList[7].ToString();
                        string usuarioRed = string.Empty;
                        if (formaList[1].ToString() != "0")
                        {
                            string dictionaryArray = DCTbassement[formaList[1].ToString()];
                            string[] splBassment = dictionaryArray.Split(';');
                            usuarioRed = NetworkAccountName(nombre, 0);

                            employee.SamAccountName = usuarioRed;
                            employee.GivenName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(firstNAME);
                            employee.MiddleName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(secondNAME);
                            employee.FirstLastName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lastNAME);
                            employee.SecondLastName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(secondlastNAME);
                            employee.PostOfficeBox = formaList[2].ToString();

                            if (formaList[1].ToString() == "0")
                            {
                                employee.UserBasement = "pruebas.bot.1";
                                Utility.LogApplication.LogWrite("Usuario Base ==> " + "El usuario: " + usuarioRed + ", no se encontro usuario base se mueve a Formacion");
                            }
                            else
                            {
                                employee.UserBasement = splBassment[0].ToString();
                            }

                            employee.TicketID = "";
                            employee.HomePage = splBassment[4].ToString();
                            employee.City = splBassment[1].ToString();
                            employee.State = splBassment[2].ToString();
                            employee.Country = splBassment[3].ToString();
                            employee.Description = formaList[6].ToString();
                            employee.Domain = splBassment[5].ToString();
                            employee.Department = splBassment[6].ToString();
                            employee.Company = "Konecta";
                            employee.fecha_modificacion = formaList[13].ToString();
                            employee.nombre_jefe = formaList[11].ToString();
                            employee.email_jefe = formaList[12].ToString();
                            employee.documento_jefe = formaList[10].ToString();
                            employee.codPCR = formaList[1].ToString();
                            employee.cliente = formaList[8].ToString();

                            Function.Create.User(employee.FirstLastName, employee.SecondLastName,
                                            employee.GivenName, employee.MiddleName, employee.Domain,
                                            employee.HomePage, employee.Country, employee.State, employee.City, employee.PostOfficeBox, employee.SamAccountName,
                                            employee.Description, employee.Department, employee.Company, employee.UserBasement, employee.fecha_modificacion, employee.nombre_jefe, employee.email_jefe, employee.documento_jefe, employee.codPCR, employee.cliente);

                            if (Models.GlobalVar.existUser != false)
                            {
                                if (employee.UserBasement != "pruebas.bot.1")
                                {
                                    ActiveDirectory.Services.ServicesProvider.User_Add_Group(employee.SamAccountName, employee.UserBasement);
                                }
                                else
                                {
                                    Utility.LogApplication.LogWrite("Sin Usuario Base ==> " + "El usuario: " + usuarioRed + ", imposible asignarle grupo");
                                }

                                if (usuarioRed != "")
                                {
                                    ActiveDirectory.Services.ServicesProvider.EnableAccount(usuarioRed);
                                }
                            }
                        }
                        else
                        {
                            Utility.LogApplication.LogWrite("ValidateFormaUsers ==> " + "El nombre: " + formaList[7].ToString() + ", no posee codigo PCRC");
                        }
                    }

                    formaList.Clear();
                }
            }
            catch(Exception ex)
            {
                Utility.LogApplication.LogWrite("ValidateFormaUsers ==> " + "Exception: " + ex.Message.ToString());
            }
        }
        #endregion

        #region VALIDATE USER FOR TYPE RETREAT
        public static void ValidateRECHZUsers()
        {
            ArrayList RECHZList = new ArrayList();
            string resultRECHZ = string.Empty;

            for (int i = 0; i < Models.ModelExpedientes.candidatRechz.Count; i++)
            {
                string rechz = Models.ModelExpedientes.candidatRechz[i].ToString();
                string[] SPLTRechz = rechz.Split(';');

                int yillFiles = SPLTRechz.Length;

                if (yillFiles == 14)
                {
                    for (int j = 0; j < SPLTRechz.Count(); j++)
                    {
                        resultRECHZ = SPLTRechz[j].ToString();
                        RECHZList.Add(resultRECHZ);
                    }

                    RECHZList.ToString();

                    string usred = Controllers.ActiveDirectory.Services.ServicesProvider.UserNetwork(RECHZList[2].ToString());

                    if (usred != "")
                    {
                        string distinguishedname = Controllers.ActiveDirectory.Services.ServicesProvider.DistGuiName(RECHZList[2].ToString());
                        ActiveDirectory.Services.ServicesProvider.DisableAccount(usred, distinguishedname);
                        Models.GlobalVar.CountDeshabilitadoYESProcess++;
                    }
                    else
                    {
                        Utility.LogApplication.LogWrite("Validate Users - Type: " + RECHZList[9].ToString() + " ==> " + "State: " + RECHZList[4].ToString() + ", FullName: " + RECHZList[7].ToString() + ", no posee codigo PCRC");
                        Models.GlobalVar.CountDeshabilitadoNOProcess++;
                    }
                }
                RECHZList.Clear();
            }
        }
        #endregion

        #region CREATE NETWORK ACCOUNT NAME
        public static string NetworkAccountName(string REDUser, int vState)
        {
            string validAccountName = string.Empty;

            try
            {
                string[] SPLName = REDUser.Split(' ');

                if (vState == 0)
                {
                    firstNAME = SPLName[0].ToString().ToLower();
                    secondNAME = SPLName[1].ToString().ToLower();
                    lastNAME = SPLName[2].ToString().ToLower();
                    secondlastNAME = SPLName[3].ToString().ToLower();
                }
                else
                {
                    firstNAME = SPLName[2].ToString().ToLower();
                    secondNAME = SPLName[3].ToString().ToLower();
                    lastNAME = SPLName[0].ToString().ToLower();
                    secondlastNAME = SPLName[1].ToString().ToLower();
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