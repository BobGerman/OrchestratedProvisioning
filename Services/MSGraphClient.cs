using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using OrchestratedProvisioning.Model;

namespace OrchestratedProvisioning.Services
{
    class MSGraphClient
    {
        public async Task<JObject> Get(string token, string url)
        {
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", token);

                var response = await client.GetAsync(url);

                await EnsureSuccess(response);

                var responseContent = await response.Content.ReadAsStringAsync();
                var result = JObject.Parse(responseContent);
                return result;
            }
        }

        public async Task<(JObject, QueueMessage.ResultCode)> PostTeamsAsyncOperation(string token, string url, string stringContent)
        {
            JObject result = null;
            var done = false;

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", token);

                var body = new StringContent(stringContent);
                body.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                var response = await client.PostAsync(url, body);

                await EnsureSuccess(response);

                var operationUrl = "https://graph.microsoft.com/beta" + response.Headers.Location;
                var retriesRemaining = Constants.RetryMax;
                while (!done && retriesRemaining-- > 0)
                {
                    await Task.Delay(Constants.RetryInterval);
                    result = await Get(token, operationUrl);
                    done = result["status"]?.ToString() == "succeeded";
                }
            }

            return (result, done ? QueueMessage.ResultCode.succeeded : QueueMessage.ResultCode.incomplete);
        }

        // If the response isn't successful, get the message from Graph and throw an excaption
        private async Task EnsureSuccess(HttpResponseMessage response)
        {
            if (!response.IsSuccessStatusCode)
            {
                var responseContent = await response.Content.ReadAsStringAsync();
                var responseJson = JObject.Parse(responseContent);
                var errorMessage = responseJson["error"]["message"].ToString();

                throw new Exception(errorMessage);
            }
        }
    }
}
