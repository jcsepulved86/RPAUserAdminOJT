using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;

namespace RPAUserAdminOJT.Controllers.Utility
{
    public class SendMail
    {

        public static Dictionary<string, string> DCTmail = new Dictionary<string, string>();

        public static string result = string.Empty;


        public static string SendingMessage(string NameBox, int PYG)
        {
            try
            {

                string file = Models.GlobalVar.rootMain;
                string fileLote = Models.GlobalVar.rootLote;
                string fileLog = Models.GlobalVar.rootLogs;


                if (PYG == 1 && file != "")
                {
                    string body = "<head>" + "[Su caso ha sido solucionado]" + "</head>" +
                              "<body>" +
                                "<h1>Servicio Automatizado para la Creación de Usuarios entrega de archivo principal</h1>" + "\n" +
                                "<a>Recuerde que NO debe responder este correo ya que es una notificación de información.</a>" + "\n" +
                                "<h1> </h1>" + "\n" +
                                "<a>A continuación se encuentra en archivo adjunto la informacion procesada.</a>" + "\n" +
                                "<h1> </h1>" + "\n" +
                                "<a>Servicios TI / Tecnología</a>" + "\n" +
                                "<h1> </h1>" + "\n" +
                                "<a>Konecta</a>" + "\n" +
                              "</body>";

                    Attachment data = new Attachment(file, MediaTypeNames.Application.Octet);

                    MailMessage message = new MailMessage();
                    message.From = new MailAddress(ConfigurationManager.AppSettings["AdminUser"]);
                    message.To.Add(new MailAddress(NameBox));
                    message.Subject = "Creación de Usuarios: Archivo Principal";
                    message.IsBodyHtml = true;
                    message.Body = body;
                    message.Attachments.Add(data);

                    SmtpClient client = new SmtpClient();

                    client.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["AdminUser"], ConfigurationManager.AppSettings["AdminPassword"]);
                    client.Port = 587;
                    client.EnableSsl = true;
                    client.Host = "smtp.gmail.com";

                    client.Send(message);

                    result = "true";
                }
                else if (PYG == 2 && fileLote != "")
                {
                    string bodyLog = "<head>" + "[Su caso ha sido solucionado]" + "</head>" +
                              "<body>" +
                                "<h1>Servicio Automatizado para la Creación de Usuarios entrega a tutores</h1>" + "\n" +
                                "<a>Recuerde que NO debe responder este correo ya que es una notificación de información.</a>" + "\n" +
                                "<h1> </h1>" + "\n" +
                                "<a>A continuación se encuentra en archivo adjunto la informacion procesada.</a>" + "\n" +
                                "<h1> </h1>" + "\n" +
                                "<a>Servicios TI / Tecnología</a>" + "\n" +
                                "<h1> </h1>" + "\n" +
                                "<a>Konecta</a>" + "\n" +
                              "</body>";

                    Attachment dataLog = new Attachment(fileLote, MediaTypeNames.Application.Octet);

                    MailMessage messageLog = new MailMessage();
                    messageLog.From = new MailAddress(ConfigurationManager.AppSettings["AdminUser"]);
                    messageLog.To.Add(new MailAddress(NameBox));
                    messageLog.Subject = "Creación de Usuarios";
                    messageLog.IsBodyHtml = true;
                    messageLog.Body = bodyLog;
                    messageLog.Attachments.Add(dataLog);

                    SmtpClient clientLog = new SmtpClient();

                    clientLog.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["AdminUser"], ConfigurationManager.AppSettings["AdminPassword"]);
                    clientLog.Port = 587;
                    clientLog.EnableSsl = true;
                    clientLog.Host = "smtp.gmail.com";

                    clientLog.Send(messageLog);

                    result = "true";
                }
                else if (PYG == 3 && fileLog != "")
                {
                    string bodyLog = "<head>" + "[Su caso ha sido solucionado]" + "</head>" +
                              "<body>" +
                                "<h1>Servicio Automatizado para la Creación de Usuarios entrega de trazabilidad</h1>" + "\n" +
                                "<a>Recuerde que NO debe responder este correo ya que es una notificación de información.</a>" + "\n" +
                                "<h1> </h1>" + "\n" +
                                "<a>A continuación se encuentra en archivo adjunto la informacion procesada.</a>" + "\n" +
                                "<h1> </h1>" + "\n" +
                                "<a>Servicios TI / Tecnología</a>" + "\n" +
                                "<h1> </h1>" + "\n" +
                                "<a>Konecta</a>" + "\n" +
                              "</body>";

                    Attachment dataLog = new Attachment(fileLog, MediaTypeNames.Application.Octet);

                    MailMessage messageLog = new MailMessage();
                    messageLog.From = new MailAddress(ConfigurationManager.AppSettings["AdminUser"]);
                    messageLog.To.Add(new MailAddress(NameBox));
                    messageLog.Subject = "Creación de Usuarios: Trazabilidad";
                    messageLog.IsBodyHtml = true;
                    messageLog.Body = bodyLog;
                    messageLog.Attachments.Add(dataLog);

                    SmtpClient clientLog = new SmtpClient();

                    clientLog.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["AdminUser"], ConfigurationManager.AppSettings["AdminPassword"]);
                    clientLog.Port = 587;
                    clientLog.EnableSsl = true;
                    clientLog.Host = "smtp.gmail.com";

                    clientLog.Send(messageLog);

                    result = "true";
                }
                else
                {
                    result = "false";
                }

                return result;
            }
            catch (SmtpException e)
            {
                string result = e.Message;
                return result;

            }
        }


        public static void DataAnalytics()
        {
            Models.GlobalVar.filePath = @"C:\AllGithub\RPAUserAdminOJT\RPAUserAdminOJT\bin\Debug\OJTMainProcess\RPAUserAdminOJT_01022021_838_jonathan.arias.o.csv";

            string result = string.Empty;
            string[] filelines = File.ReadAllLines(Models.GlobalVar.filePath);
            ArrayList arrayEmail = new ArrayList();
            Models.Employee employee = new Models.Employee();

            for (int a = 1; a < filelines.Length; a++)
            {
                string[] fillines = filelines[a].Split(',');
                int yillFiles = fillines.Length;

                string correo = fillines[7].ToString();
                string accion = fillines[3].ToString();

                string dictionaryArray = string.Empty;

                if (accion == "User Create")
                {
                    try
                    {
                        dictionaryArray = DCTmail[correo];
                        dictionaryArray = dictionaryArray + ";" + filelines[a];
                        DCTmail[correo] = dictionaryArray;

                    }
                    catch (Exception)
                    {
                        DCTmail.Add(correo, filelines[a]);
                        arrayEmail.Add(correo);
                    }
                }
                else
                {
                    continue;
                }
            }


            for (int i = 0; i < arrayEmail.Count; i++)
            {
                string info = string.Empty;

                info = DCTmail[arrayEmail[i].ToString()];
                string[] mailSPLT = info.Split(';');

                for (int j = 0; j < mailSPLT.Length; j++)
                {
                    string most = mailSPLT[j];

                    string[] mostSPLT = most.Split(',');

                    string accion = mostSPLT[3].ToString();
                    string DateModify = mostSPLT[4].ToString();
                    string NumLote = mostSPLT[5].ToString();
                    string NameBox = mostSPLT[6].ToString();
                    string EmailBox = mostSPLT[7].ToString();
                    string IdBox = mostSPLT[8].ToString();
                    string GivenName = mostSPLT[9].ToString();
                    string PostOfficeBox = mostSPLT[10].ToString();
                    string client = mostSPLT[11].ToString();
                    string CodPRC = mostSPLT[12].ToString();
                    string UserBasement = mostSPLT[13].ToString();
                    string SamAccountName = mostSPLT[14].ToString();
                    string PasswordUser = mostSPLT[15].ToString();

                    Utility.FillExcel.WriteExcelLote(accion, DateModify, NumLote, NameBox, EmailBox, IdBox, GivenName, PostOfficeBox, client, CodPRC, UserBasement, SamAccountName, PasswordUser);

                }

                Models.GlobalVar.countCSV = 0;

                SendingMessage("gestion_usuarios@grupokonecta.com", 2);
            }

        }


    }
}
