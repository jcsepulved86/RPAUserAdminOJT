using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Web.Script.Serialization;

namespace RPAUserAdminOJT.Controllers.Function
{
    public class ConsultAPI
    {

        #region GETTOKEN
        public static string getToken()
        {
            string token = string.Empty;

            try
            {
                var httpWebRequest = (HttpWebRequest)WebRequest.Create(ConfigurationManager.AppSettings["authLogin"].ToString());
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "POST";

                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    var input = "{\"userName\":\"" + ConfigurationManager.AppSettings["UserName"].ToString() + "\"," +
                                "\"password\":\"" + ConfigurationManager.AppSettings["Password"].ToString() + "\"}";
                    streamWriter.Write(input);
                    streamWriter.Flush();
                    streamWriter.Close();
                }

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                string response;

                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    response = streamReader.ReadToEnd();
                }

                var entries = response.TrimStart('{').TrimEnd('}').Replace("\"", String.Empty).Split(',');
                token = entries[1].Split(':')[1];
            }
            catch (Exception e)
            {
                string resultEx = e.Message.ToString();
            }
            return token;
        }
        #endregion

        #region GETRESPONSE
        public static List<Models.ModelExpedientes> getExpedientes(string tokenAuth)
        {

            IList<Models.ModelExpedientes> expedientes = new List<Models.ModelExpedientes>();

            try
            {
                string url = ConfigurationManager.AppSettings["UrlExpedientes"].ToString();

                var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                httpWebRequest.ContentType = "application/json; charset=utf-8";
                httpWebRequest.Method = "POST";
                httpWebRequest.Accept = "application/json; charset=utf-8";

                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    string loginjson = new JavaScriptSerializer().Serialize(new
                    {
                        token = tokenAuth,
                    });

                    streamWriter.Write(loginjson);
                    streamWriter.Flush();
                    streamWriter.Close();

                    var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        var result = streamReader.ReadToEnd();

                        JObject resultado = JObject.Parse(result);
                        IList<JToken> results = resultado["data"].Children().ToList();
                        IList<Models.ModelExpedientes> searchResults = new List<Models.ModelExpedientes>();

                        foreach (JToken data in results)
                        {
                            Models.ModelExpedientes searchResult = JsonConvert.DeserializeObject<Models.ModelExpedientes>(data.ToString());
                            searchResults.Add(searchResult);
                        }

                        expedientes = searchResults;
                    }
                }
            }
            catch (Exception e)
            {
                string resultEx = e.Message.ToString();
            }
            return expedientes.ToList();
        }

        #endregion

    }
}