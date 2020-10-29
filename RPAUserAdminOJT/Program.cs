using System;
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RPAUserAdminOJT.Controllers
{
    public class Program
    {
        public static string firstNAME = string.Empty;
        public static string secondNAME = string.Empty;
        public static string lastNAME = string.Empty;
        public static string secondlastNAME = string.Empty;
        public static Dictionary<string, string> DCTbassement = new Dictionary<string, string>();

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

            Models.GlobalVar.filePath = @"C:\AllGithub\RPAUserAdminOJT\RPAUserAdminOJT\bin\Debug\UserBassemet.csv";

            string result = string.Empty;
            string[] filelines = File.ReadAllLines(Models.GlobalVar.filePath);
            ArrayList UBassement = new ArrayList();
            Models.Employee employee = new Models.Employee();

            for (int a = 1; a < filelines.Length; a++)
            {
                string[] fillines = filelines[a].Split(';');
                int yillFiles = fillines.Length;

                if (yillFiles == 7)
                {
                    try
                    {
                        DCTbassement.Add(fillines[3].ToString().Trim(), fillines[0].ToString());
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


        public static void ValidateOJTUsers()
        {

            try
            {
                Models.Employee employee = new Models.Employee();

                ArrayList OJTList = new ArrayList();
                string resultOJT = string.Empty;

                for (int i = 25; i < Models.ModelExpedientes.candidatOJT.Count; i++)
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

                        if (usrd!= "")
                        {
                            //Function.Delete.User(usrd);
                        }

                        string nombre = OJTList[7].ToString();

                        string usuarioRed = NetworkAccountName(nombre,1);

                        employee.SamAccountName = usuarioRed;
                        employee.GivenName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(firstNAME);
                        employee.MiddleName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(secondNAME);
                        employee.FirstLastName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lastNAME);
                        employee.SecondLastName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(secondlastNAME);
                        employee.PostOfficeBox = OJTList[2].ToString();
                        if (OJTList[1].ToString() == "0")
                        {
                            employee.UserBasement = "pruebas.bot.1";
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
            catch(Exception ex)
            {

            }

        }


        public static void ValidateFormaUsers()
        {

             try
            {
                Models.Employee employee = new Models.Employee();

                ArrayList formaList = new ArrayList();
                string resultForma = string.Empty;

                for (int i = 1; i < Models.ModelExpedientes.candidatForma.Count; i++)
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

                        string usuarioRed = NetworkAccountName(nombre,0);

                        employee.SamAccountName = usuarioRed;
                        employee.GivenName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(firstNAME);
                        employee.MiddleName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(secondNAME);
                        employee.FirstLastName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lastNAME);
                        employee.SecondLastName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(secondlastNAME);
                        employee.PostOfficeBox = formaList[2].ToString();
                        employee.UserBasement = "pruebas.bot.1";
                        employee.TicketID = "";
                        employee.HomePage = "allus.com.co";
                        employee.City = "Medellin";
                        employee.State = "Antioquia";
                        employee.Country = "Colombia";
                        employee.Description = "Usuarios en formacion BOT";
                        employee.Domain = "multienlace.com.co";
                        employee.Department = "Interno";
                        employee.Company = "Konecta";

                        Function.Create.User(employee.FirstLastName, employee.SecondLastName,
                                        employee.GivenName, employee.MiddleName, employee.Domain,
                                        employee.HomePage, employee.Country, employee.State, employee.City, employee.PostOfficeBox, employee.SamAccountName,
                                        employee.Description, employee.Department, employee.Company, employee.UserBasement);


                    }

                    formaList.Clear();
                }
            }
            catch(Exception ex)
            {

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

                    string usrd = Function.Delete.Cedula(RECHZList[2].ToString());

                    if (usrd != "")
                    {
                        //Function.Delete.User(usrd);
                    }
                }

                RECHZList.Clear();
            }

        }


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
                return validAccountName;
            }

        }



    }
}