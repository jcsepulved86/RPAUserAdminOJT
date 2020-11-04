using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
#region USING
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
#endregion

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
                var httpWebRequest = (HttpWebRequest)WebRequest.Create(ConfigurationManager.AppSettings["UrlGenToken"].ToString());
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "POST";

                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    var input = "{\"user\":\"" + ConfigurationManager.AppSettings["UserName"].ToString() + "\"," +
                                "\"pw\":\"" + ConfigurationManager.AppSettings["Password"].ToString() + "\"}";
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
                token = entries[1].Split(':')[2];
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
                string url = ConfigurationManager.AppSettings["UrlDocuments"].ToString();

                var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                httpWebRequest.ContentType = "application/json; charset=utf-8";
                httpWebRequest.Method = "GET";
                httpWebRequest.Accept = "application/json";
                httpWebRequest.Headers.Add("Authorization", tokenAuth);

                WebResponse response = httpWebRequest.GetResponse();

                using (var streamReader = new StreamReader(response.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();

                    JObject resultado = JObject.Parse(result);
                    IList<JToken> results = resultado["payload"].Children().ToList();
                    IList<Models.ModelExpedientes> searchResults = new List<Models.ModelExpedientes>();

                    foreach (JToken data in results)
                    {
                        Models.ModelExpedientes searchResult = JsonConvert.DeserializeObject<Models.ModelExpedientes>(data.ToString());
                        searchResults.Add(searchResult);
                    }

                    expedientes = searchResults;
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