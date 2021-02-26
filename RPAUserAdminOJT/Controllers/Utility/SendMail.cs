using System;
using System.Collections.Generic;
using System.Configuration;
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
        public static string result = string.Empty;

        public static string Message(string tutor)
        {
            try
            {

                string file = Models.GlobalVar.rootMain;
                string fileLog = Models.GlobalVar.rootLogs;

                if (file != null)
                {
                    string body = "<head>" + "[LOTE #202020773782] ha sido completado" + "</head>" +
                              "<body>" +
                                "<h1>Servicio Automatizado para la Creación de Cuentas de Usuarios</h1>" + "\n" +
                                "<a>Por favor, no responda a este correo, ya que es meramente informativo</a>" + "\n" +
                                "<a>A continuación se encuentra en archivo adjunto la informacion procesada.</a>" + "\n" +
                                "<a>Servicios TI / Tecnología</a>" + "\n" +
                                "<a>Konecta</a>" + "\n" +
                              "</body>";



                    Attachment data = new Attachment(file, MediaTypeNames.Application.Octet);

                    MailMessage message = new MailMessage();
                    message.From = new MailAddress(ConfigurationManager.AppSettings["AdminUser"]);
                    message.To.Add(new MailAddress(tutor));
                    message.Subject = "Creación de Cuentas de Usuarios";
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
                else if (fileLog != null)
                {
                    string bodyLog = "<head>" + "[LOTE #202020773782] ha sido completado" + "</head>" +
                              "<body>" +
                                "<h1>Servicio Automatizado para la trazabilidad de la Creación de Cuentas de Usuarios</h1>" + "\n" +
                                "<a>Por favor, no responda a este correo, ya que es meramente informativo</a>" + "\n" +
                                "<a>A continuación se encuentra en archivo adjunto la informacion procesada.</a>" + "\n" +
                                "<a>Servicios TI / Tecnología</a>" + "\n" +
                                "<a>Konecta</a>" + "\n" +
                              "</body>";

                    Attachment dataLog = new Attachment(fileLog, MediaTypeNames.Application.Octet);

                    MailMessage messageLog = new MailMessage();
                    messageLog.From = new MailAddress(ConfigurationManager.AppSettings["AdminUser"]);
                    messageLog.To.Add(new MailAddress(tutor));
                    messageLog.Subject = "Trazabilidad de la Creación de Cuentas de Usuarios";
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

    }
}
