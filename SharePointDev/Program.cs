using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;

namespace SharePointDev
{
    class Program
    {
        // https://developer.microsoft.com/en-us/graph/graph-explorer
        // https://docs.microsoft.com/en-us/graph/auth-overview
        // https://docs.microsoft.com/en-us/onedrive/developer/rest-api/api/driveitem_get_content?view=odsp-graph-online

        
        public static class SharepointUrlConst
        {
            public const string AccessToken = "https://accounts.accesscontrol.windows.net/{0}/tokens/OAuth/2";
            public const string Token = "https://login.microsoftonline.com/{0}/oauth2/token";

            public const string ListDrive= "https://graph.microsoft.com/v1.0/drives/{0}/root/children"; // 0 = DriveId
            public const string SearchFromDrive = "https://graph.microsoft.com/v1.0/drives/{0}/root/search(q='{{1}}')"; // 0 = DriveId && 1 = Text to search

            public const string ListFolder = "https://graph.microsoft.com/v1.0/drives/{0}/items/{1}/children"; // 0 = DriveId && 1 = FolderId
            public const string ListFolderByPath = "https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}:/children"; // 0 = DriveId && 1 = Folder Path
            public const string GetInfoFile = "https://graph.microsoft.com/v1.0/drives/{0}/items/{1}"; // 0 = DriveId && 1 = FileId
            public const string GetInfoFileByPath = "https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}:/"; // 0 = DriveId && 1 = File Path
            public const string DownloadFile = "https://graph.microsoft.com/v1.0/drives/{0}/items/{1}/content"; // 0 = DriveId && 1 = FileId
            public const string DownloadFileByPath = "https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}:/content"; // 0 = DriveId && 1 = File Path
            public const string DownloadFileSH = "https://{0}/_layouts/15/download.aspx?UniqueId={1}&Translate=false&tempauth={2}&ApiVersion=2.0"; // 0 = SharepointUrl && 1 = FileId && 2 = TokenBearer
        }


        /* NEED AN APPLICATION WITH ADMIN RIGHTS. LIST READ DOWNLAOD */
        public static class ClientConst
        {
            /* Needed to connect to Sharepoint Online */
            public const string Authority = "siteweb";
            public const string Sharepoint = "company.sharepoint.com/sites/drive";
            public const string UserName = "username";
            public const string UserPassword = "password";

            /* Needed to connect to Graph */
            public const string TenantId = "4f2ee46a-1234-4ca1-1234-f99645988ba3";
            public const string Instance = "https://login.microsoftonline.com/";
            public const string ClientId = "12345678-f3fa-44d6-a660-ec40ec39a11b";
            public const string ClientSecret = "wdcrSOTP0qsqsqcqegioijgkll";
            public const string ApiEndPoint = "https://graph.microsoft.com/"; // Office 365
            // public const string ApiEndPoint = "https://graph.windows.net/"; // Azure AD

            public const string DriveId = "U9U3iYvQy2Xw6BOgEb9QTxLSd3UZri442mBSYnq5t6KD16z";
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Settings: {0}{1} \r\n", ClientConst.Instance, ClientConst.Authority);

            // [GOOD] Fetch TenantId + RessouceId + Token 
            //Console.WriteLine("Fetch TenantId + RessouceId + Token");
            //var tokenNull = GetTenantId();
            //Console.WriteLine("Token: {0} \r\n", tokenNull);


            // [GOOD] Fetching Token
            Console.WriteLine("Fetch Token");
            var token = GetGraphToken();
            Console.WriteLine("Token: {0} \r\n", token);


            // [GOOD] List Drive
            Console.WriteLine("List Drive");
            var filesInDrive = ListDrive(token);
            Console.WriteLine("List Drive: {0} \r\n", filesInDrive.Count);


            // [GOOD] Search Drive
            Console.WriteLine("Search Drive");
            var filesSearch = SearchDrive(token, "test");
            Console.WriteLine("Search Drive: {0} \r\n", filesSearch.Count);


            // [GOOD] List Files
            Console.WriteLine("List Files");
            var files = ListFolder(token, filesInDrive.First().Name);
            Console.WriteLine("List Files: {0} \r\n", files.Count);


            // [GOOD] Download File
            Console.WriteLine("Download File");
            var fileInfo = FileInfo(token, filesInDrive.Last().Name);
            var file = DownloadFile(token, filesInDrive.Last().Name);
            Console.WriteLine("Downloaded File: {0} {2}/{1} \r\n", fileInfo.Name, fileInfo.Size, file.Length);


            Console.ReadLine();
        }

        public static Stream DownloadFile(string token, string filePath)
        {
            var result = new MemoryStream();

            try
            {
                string requestUrl = $"{string.Format(SharepointUrlConst.DownloadFileByPath, ClientConst.DriveId, filePath)}";
                var client = new HttpClient();
                var httpRequestMessage = new HttpRequestMessage(HttpMethod.Get, new Uri(requestUrl));
                client.BaseAddress = new Uri(requestUrl);
                client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", $"Bearer {token}");
                client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "charset=utf-8");
                var httpResponseMessage = client.SendAsync(httpRequestMessage).Result;

                Console.WriteLine("Response: {0}", httpResponseMessage.ReasonPhrase);

                result = new MemoryStream(httpResponseMessage.Content.ReadAsByteArrayAsync().Result);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}", ex.Message);
            }

            return result;
        }

        public static FileModel FileInfo(string token, string filePath)
        {
            var result = new FileModel();

            try
            {
                string requestUrl = $"{string.Format(SharepointUrlConst.GetInfoFileByPath, ClientConst.DriveId, filePath)}";
                var client = new HttpClient();
                var httpRequestMessage = new HttpRequestMessage(HttpMethod.Get, new Uri(requestUrl));
                client.BaseAddress = new Uri(requestUrl);
                client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", $"Bearer {token}");
                client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "charset=utf-8");
                var httpResponseMessage = client.SendAsync(httpRequestMessage).Result;

                Console.WriteLine("Response: {0}", httpResponseMessage.ReasonPhrase);

                var responseResult = httpResponseMessage.Content.ReadAsStringAsync().Result;
                result = JsonConvert.DeserializeObject<FileModel>(responseResult);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}", ex.Message);
            }

            return result;
        }

        public static List<FileModel> ListFolder(string token, string folderPath)
        {
            var result = new List<FileModel>();

            try
            {
                string requestUrl = $"{string.Format(SharepointUrlConst.ListFolderByPath, ClientConst.DriveId, folderPath)}";
                var client = new HttpClient();
                var httpRequestMessage = new HttpRequestMessage(HttpMethod.Get, new Uri(requestUrl));
                client.BaseAddress = new Uri(requestUrl);
                client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", $"Bearer {token}");
                client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json; charset=utf-8");
                var httpResponseMessage = client.SendAsync(httpRequestMessage).Result;

                Console.WriteLine("Response: {0}", httpResponseMessage.ReasonPhrase);

                var responseResult = httpResponseMessage.Content.ReadAsStringAsync().Result;
                result = JsonConvert.DeserializeObject<ResponseModel>(responseResult).Value;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}", ex.Message);
            }

            return result;
        }

        public static List<FileModel> SearchDrive(string token, string text)
        {
            var result = new List<FileModel>();

            try
            {
                string requestUrl = $"{string.Format(SharepointUrlConst.SearchFromDrive, ClientConst.DriveId, text)}";
                var client = new HttpClient();
                var httpRequestMessage = new HttpRequestMessage(HttpMethod.Get, new Uri(requestUrl));
                client.BaseAddress = new Uri(requestUrl);
                client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", $"Bearer {token}");
                client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json; charset=utf-8");
                var httpResponseMessage = client.SendAsync(httpRequestMessage).Result;
                Console.WriteLine("Response: {0}", httpResponseMessage.ReasonPhrase);

                var responseResult = httpResponseMessage.Content.ReadAsStringAsync().Result;
                result = JsonConvert.DeserializeObject<ResponseModel>(responseResult).Value;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}", ex.Message);
            }


            return result;
        }

        public static List<FileModel> ListDrive(string token)
        {
            var result = new List<FileModel>();

            try
            {
                string requestUrl = $"{string.Format(SharepointUrlConst.ListDrive, ClientConst.DriveId)}";
                var client = new HttpClient();
                var httpRequestMessage = new HttpRequestMessage(HttpMethod.Get, new Uri(requestUrl));
                client.BaseAddress = new Uri(requestUrl);
                client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", $"Bearer {token}");
                client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json; charset=utf-8");
                var httpResponseMessage = client.SendAsync(httpRequestMessage).Result;
                Console.WriteLine("Response: {0}", httpResponseMessage.ReasonPhrase);

                var responseResult = httpResponseMessage.Content.ReadAsStringAsync().Result;
                result = JsonConvert.DeserializeObject<ResponseModel>(responseResult).Value;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}", ex.Message);
            }


            return result;
        }

        public static string GetGraphToken()
        {
            try
            {
                string requestUrl = $"{string.Format(SharepointUrlConst.Token, ClientConst.TenantId)}";
                var values = new Dictionary<string, string> {
                    { "grant_type", "client_credentials" },
                    { "client_id", ClientConst.ClientId},
                    { "client_secret", ClientConst.ClientSecret},
                    { "resource", ClientConst.ApiEndPoint }
                };


                var client = new HttpClient();
                var httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, new Uri(requestUrl));
                client.BaseAddress = new Uri(requestUrl);
                client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/x-www-form-urlencoded");
                var content = new FormUrlEncodedContent(values);
                httpRequestMessage.Content = content;
                var httpResponseMessage = client.SendAsync(httpRequestMessage).Result;
                var result = httpResponseMessage.Content.ReadAsStringAsync().Result;


                Console.WriteLine("Response: {0}", httpResponseMessage.ReasonPhrase);


                var accessToken = "";
                var resultObj = JsonConvert.DeserializeObject<dynamic>(result);
                if (resultObj != null)
                    accessToken = Convert.ToString(resultObj.access_token);

                return accessToken;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "";
            }
        }

        public static string GetTenantId()
        {
            string requesrUrl = $"{string.Format("https://{0}", ClientConst.Sharepoint)}";
            var myWebRequest = WebRequest.Create(requesrUrl);
            myWebRequest.Method = "GET";
            myWebRequest.Headers.Add("Authorization", "Bearer");

            var tenantID = "";
            var resourceID = "";
            WebResponse myWebResponse = null;
            try
            {
                myWebResponse = myWebRequest.GetResponse();
                return myWebResponse.ToString();
            }
            catch (WebException ex)
            {
                //Console.Write(ex.Message);
                //get the Web exception and read the headers

                string[] headerAuthenticateValue = ex.Response.Headers.GetValues("WWW-Authenticate");
                if (headerAuthenticateValue != null)
                {
                    //get the array separated by comma
                    //Console.WriteLine(" Value => " + headerAuthenticateValue.Length);

                    foreach (string stHeader in headerAuthenticateValue)
                    {
                        string[] stArrHeaders = stHeader.Split(',');
                        //loop all the key value pair of WWW-Authenticate

                        foreach (string stValues in stArrHeaders)
                        {
                            // Console.WriteLine(" Value =>" + stValues);
                            if (stValues.StartsWith("Bearer realm="))
                            {
                                tenantID = stValues.Substring(14);
                                tenantID = tenantID.Substring(0, tenantID.Length - 1);
                            }

                            if (stValues.StartsWith("client_id="))
                            {
                                //this value is consider as resourceid which is required for getting the access token
                                resourceID = stValues.Substring(11);
                                resourceID = resourceID.Substring(0, resourceID.Length - 1);
                            }
                        }
                    }
                }

                Console.WriteLine("TenantId: " + tenantID);
                Console.WriteLine("ResourceID: " + resourceID);


                var accessTokenUrl = string.Format(SharepointUrlConst.AccessToken, tenantID);
                myWebRequest = WebRequest.Create(accessTokenUrl);
                myWebRequest.ContentType = "application/x-www-form-urlencoded";
                myWebRequest.Method = "POST";


                // Add the below body attributes to the request
                /*
                 *  grant_type  client_credentials  client_credentials
                 client_id  ClientID@TenantID 
                 client_secret  ClientSecret 
                 resource  resource/SiteDomain@TenantID  resourceid/abc.sharepoint.com@tenantID
                 */


                var postData = "grant_type=client_credentials";
                postData += "&client_id=" + ClientConst.ClientId + "@" + tenantID;
                postData += "&client_secret=" + ClientConst.ClientSecret;
                postData += "&resource=" + resourceID + "/" + ClientConst.Authority + "@" + tenantID;
                var data = Encoding.ASCII.GetBytes(postData);

                using (var stream = myWebRequest.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }
                var response = (HttpWebResponse)myWebRequest.GetResponse();

                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

                string[] stArrResponse = responseString.Split(',');

                //get the access token and expiry time ,etc

                var accessToken = "";
                foreach (var stValues in stArrResponse)
                {

                    if (stValues.StartsWith("\"access_token\":"))
                    {
                        //Console.WriteLine(" Result => " + stValues);
                        accessToken = stValues.Substring(16);
                        //Console.WriteLine(" Result => " + accessToken);
                        accessToken = accessToken.Substring(0, accessToken.Length - 2);
                        // Console.WriteLine(" Result => " + accessToken);
                    }
                }

                return accessToken;

            }
        }

        
    }
}
