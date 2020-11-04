using System;
using System.Collections;
#region USING
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
        #region Global Variable
        public static string firstNAME = string.Empty;
        public static string secondNAME = string.Empty;
        public static string lastNAME = string.Empty;
        public static string secondlastNAME = string.Empty;
        public static Dictionary<string, string> DCTbassement = new Dictionary<string, string>();
        public static string PathRoot = ConfigurationManager.AppSettings["pathRoot"];
        #endregion

        #region MAIN ACTIVITIES
        public static void Main(string[] args)
        {
            #region Hidden
            //Function.Delete.User("harold.rodriguez.d");
            //Function.Delete.User("manuela.ruda");
            //Function.Delete.User("guillermo.poveda.a");
            //Function.Delete.User("jaime.martinez.a");
            //Function.Delete.User("daniel.bustamante.v");
            //Function.Delete.User("camilo.perez.j");
            //Function.Delete.User("valentina.arias.v");
            //Function.Delete.User("leon.posada.l");
            //Function.Delete.User("maria.viloria.v");
            //Function.Delete.User("estefanny.santamaria.v");
            //Function.Delete.User("leidy.martínez.go");
            //Function.Delete.User("maria.andrade.t");
            //Function.Delete.User("cristian.calle.u");
            //Function.Delete.User("mayra.serna.ur");
            //Function.Delete.User("juliana.gonzalez.a");
            //Function.Delete.User("juan.rio");
            //Function.Delete.User("elizabeth.rodriguez.t");
            //Function.Delete.User("daniela.olarte.m");
            //Function.Delete.User("cristian.ocampo.c");
            //Function.Delete.User("arley.montoya.m");
            //Function.Delete.User("blyangil.zuleta.c");

            //ActiveDirectory.Services.ServicesProvider.DisableAccount("prueba.prueba");
            //ActiveDirectory.Services.ServicesProvider.EnableAccount(usrd);
            #endregion

            Initialize();

        }
        #endregion

        #region Initialize
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

            Models.GlobalVar.filePath = PathRoot;

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
                            else if (id_dp_estados == "305" || id_dp_estados == "317" || id_dp_estados == "327")
                            {
                                Models.ModelExpedientes.candidatRechz.Add(ceco + ";" + cod_pcrc + ";" + documento + ";" + id_dp_cargos + ";" + id_dp_estados + ";" + nombre_cargo + ";" + nombre_ceco + ";" + nombre_completo + ";" + nombre_pcrc + ";" + tipo_estado);
                            }

                        }
                        catch(Exception ex)
                        {
                            LOGRobotica.Controllers.LogApplication.LogWrite("Initialize ==> " + "Error: " + ex.Message);
                        }

                    }

                    Models.ModelExpedientes.candidatForma.ToString();
                    Models.ModelExpedientes.candidatRechz.ToString();

                    //ValidateFormaUsers();
                    ValidateRECHZUsers();

                }
                else
                {
                    MessageBox.Show("Las credenciales ingresadas no son validas, por favor verifique e intente nuevamente", "RPAKonecta - Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                principalContext.Dispose();
            }
        }
        #endregion

        #region Validate Formation Users
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

                    if (yillFiles == 10)
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
                                LOGRobotica.Controllers.LogApplication.LogWrite("Usuario Base ==> " + "El usuario: " + usuarioRed + ", no se encontro usuario base se mueve a Formacion");
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

                            Function.Create.User(employee.FirstLastName, employee.SecondLastName,
                                            employee.GivenName, employee.MiddleName, employee.Domain,
                                            employee.HomePage, employee.Country, employee.State, employee.City, employee.PostOfficeBox, employee.SamAccountName,
                                            employee.Description, employee.Department, employee.Company, employee.UserBasement);


                            if (employee.UserBasement != "pruebas.bot.1")
                            {
                                ActiveDirectory.Services.ServicesProvider.User_Add_Group(employee.SamAccountName, employee.UserBasement);
                            }
                            else
                            {
                                LOGRobotica.Controllers.LogApplication.LogWrite("Sin Usuario Base ==> " + "El usuario: " + usuarioRed + ", imposible asignarle grupo");
                            }

                            if (usuarioRed != "")
                            {
                                ActiveDirectory.Services.ServicesProvider.EnableAccount(usuarioRed);
                            }
                        }
                        else
                        {
                            LOGRobotica.Controllers.LogApplication.LogWrite("ValidateFormaUsers ==> " + "El nombre: " + formaList[7].ToString() + ", no posee codigo PCRC");
                        }


                    }

                    formaList.Clear();
                }
            }
            catch(Exception ex)
            {
                LOGRobotica.Controllers.LogApplication.LogWrite("ValidateFormaUsers ==> " + "Exception: " + ex.Message.ToString());
            }
        }
        #endregion

        #region Validate Users for type retreat
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

                    string usred = Controllers.ActiveDirectory.Services.ServicesProvider.UserNetwork(RECHZList[2].ToString());

                    if (usred != "")
                    {
                        string distinguishedname = Controllers.ActiveDirectory.Services.ServicesProvider.DistGuiName(RECHZList[2].ToString());
                        ActiveDirectory.Services.ServicesProvider.DisableAccount(usred, distinguishedname);
                    }
                    else
                    {
                        LOGRobotica.Controllers.LogApplication.LogWrite("Validate Users - Type: " + RECHZList[9].ToString() +" ==> " + "State: " + RECHZList[4].ToString() + ", FullName: " + RECHZList[7].ToString());
                    }

                }

                RECHZList.Clear();
            }

        }
        #endregion

        #region Create Network Account Name
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
                            for (int i = 1; i < 3; i++)
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
                                    //Log error
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
                                //Log error
                            }

                        }


                    }

                    principalContext.Dispose();
                    return validAccountName;
                }
            }
            catch (Exception ex)
            {
                LOGRobotica.Controllers.LogApplication.LogWrite("NetworkAccountName ==> " + "Exception: " + ex.Message.ToString());
                return validAccountName;
            }

        }
        #endregion

        #region CODEBEHIND
        //else if (id_dp_estados == "301")
        //{
        //    Models.ModelExpedientes.candidatOJT.Add(ceco + ";" + cod_pcrc + ";" + documento + ";" + id_dp_cargos + ";" + id_dp_estados + ";" + nombre_cargo + ";" + nombre_ceco + ";" + nombre_completo + ";" + nombre_pcrc + ";" + tipo_estado);
        //}

        //Models.ModelExpedientes.candidatOJT.ToString();
        //ValidateOJTUsers();
        #endregion

        #region HIDDEN LOGIC OJT
        public static void ValidateOJTUsers()
        {

            try
            {
                Models.Employee employee = new Models.Employee();

                ArrayList OJTList = new ArrayList();
                string resultOJT = string.Empty;

                for (int i = 2; i < Models.ModelExpedientes.candidatOJT.Count; i++)
                {
                    string forma = Models.ModelExpedientes.candidatOJT[i].ToString();
                    string[] SPLTforma = forma.Split(';');

                    int yillFiles = SPLTforma.Length;

                    if (yillFiles == 10)
                    {
                        for (int j = 0; j < SPLTforma.Count(); j++)
                        {
                            resultOJT = SPLTforma[j].ToString();
                            OJTList.Add(resultOJT);
                        }
                        OJTList.ToString();

                        string usrd = Function.Delete.Cedula(OJTList[2].ToString());

                        if (usrd != "")
                        {
                            //Function.Delete.User(usrd);
                        }

                        string nombre = OJTList[7].ToString();

                        string usuarioRed = NetworkAccountName(nombre, 1);

                        employee.SamAccountName = usuarioRed;
                        employee.GivenName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(firstNAME);
                        employee.MiddleName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(secondNAME);
                        employee.FirstLastName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lastNAME);
                        employee.SecondLastName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(secondlastNAME);
                        employee.PostOfficeBox = OJTList[2].ToString();
                        if (OJTList[1].ToString() == "0")
                        {
                            //employee.UserBasement = "pruebas.bot.1";
                            //LOG DE NO SE ENCONTRO USUARIO BASE
                        }
                        else
                        {
                            employee.UserBasement = DCTbassement[OJTList[1].ToString()];
                        }
                        employee.TicketID = "";
                        employee.HomePage = "allus.com.co";
                        employee.City = "Medellin";
                        employee.State = "Antioquia";
                        employee.Country = "Colombia";
                        employee.Description = OJTList[6].ToString();
                        employee.Domain = "multienlace.com.co";
                        employee.Department = "Interno";
                        employee.Company = "Konecta";

                        Function.Create.User(employee.FirstLastName, employee.SecondLastName,
                                        employee.GivenName, employee.MiddleName, employee.Domain,
                                        employee.HomePage, employee.Country, employee.State, employee.City, employee.PostOfficeBox, employee.SamAccountName,
                                        employee.Description, employee.Department, employee.Company, employee.UserBasement);


                        if (employee.UserBasement != "pruebas.bot.1")
                        {
                            ActiveDirectory.Services.ServicesProvider.User_Add_Group(employee.SamAccountName, employee.UserBasement);
                        }
                        else
                        {
                            ///LOG ERROR sin usuario base
                        }


                    }

                    OJTList.Clear();
                }
            }
            catch (Exception ex)
            {
                LOGRobotica.Controllers.LogApplication.LogWrite("ValidateOJTUsers ==> " + "Exception: " + ex.Message.ToString());
            }

        }
        #endregion

    }
}