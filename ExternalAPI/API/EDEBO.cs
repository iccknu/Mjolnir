using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExternalAPI
{
    public class EDEBO
    {
        private string mainUrl = "";
        private string authUrl = "";

        private string token = "";

        private string login = "";
        private string password = "";
  
        /// <summary>
        /// Init EDEBO
        /// </summary>
        /// <param name="configuration">Configuration</param>
        public EDEBO(APIConfiguration configuration)
        {
            this.login = configuration.login;
            this.password = configuration.password;
            this.mainUrl = configuration.mainUrl;
            this.authUrl = configuration.authUrl;
        }
        /// <summary>
        /// Auth
        /// </summary>
        /// <returns></returns>
        public bool Auth()
        {
            string grant_type = "password";

            FormUrlEncodedContent formContent = new FormUrlEncodedContent(
                new[] {
                    new KeyValuePair<string, string>("grant_type", grant_type),
                    new KeyValuePair<string, string>("username", login),
                    new KeyValuePair<string, string>("password",password)
                });

            HttpClient client = new HttpClient();

            var response = client.PostAsync(this.authUrl, formContent);

            return true;
        }
    }
}
