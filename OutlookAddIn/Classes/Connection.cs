using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using OutlookAddIn.Shared;
using OutlookAddIn.Shared.Objects;
using OutlookAddIn.Shared.VM;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

namespace OutlookAddIn
{
    public class Connection
    {
        //============= Config [Edit these with your settings] =====================
        internal const string azureDevOpsOrganizationUrl = "https://dev.azure.com/wi17b041"; //change to the URL of your Azure DevOps account; NOTE: This must use HTTPS
        internal const string clientId = "872cd9fa-d31f-45e0-9eab-6e460a02d1f1";          //change to your app registration's Application ID, unless you are an MSA backed account
        internal const string replyUri = "urn:ietf:wg:oauth:2.0:oob";                     //change to your app registration's reply URI, unless you are an MSA backed account
        //==========================================================================

        internal static AuthenticationHeaderValue bearerAuthHeader = new AuthenticationHeaderValue("Bearer"); //Authenticationheader for querys
        internal const string azureDevOpsResourceId = "499b84ac-1321-427f-aa17-267ca6975798"; //Constant value to target Azure DevOps. Do not change  

        public AuthenticationHeaderValue ConnectMethod()
        {
            AuthenticationContext ctx = GetAuthenticationContext(null);
            AuthenticationResult result = null;


            IPlatformParameters promptBehavior = new PlatformParameters(PromptBehavior.Always);

            try
            {
                //PromptBehavior.RefreshSession will enforce an authn prompt every time. NOTE: Auto will take your windows login state if possible
                result = ctx.AcquireTokenAsync(azureDevOpsResourceId, clientId, new Uri(replyUri), promptBehavior).Result;
                Console.WriteLine("Token expires on: " + result.ExpiresOn);

                bearerAuthHeader = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                //GetProjects(bearerAuthHeader);
            }
            catch (UnauthorizedAccessException)
            {
                // If the token has expired, prompt the user with a login prompt
                result = ctx.AcquireTokenAsync(azureDevOpsResourceId, clientId, new Uri(replyUri), promptBehavior).Result;
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0}: {1}", ex.GetType(), ex.Message);
            }
            return bearerAuthHeader;
        }

        private static AuthenticationContext GetAuthenticationContext(string tenant)
        {
            AuthenticationContext ctx = null;
            if (tenant != null)
                ctx = new AuthenticationContext("https://login.microsoftonline.com/" + tenant);
            else
            {
                ctx = new AuthenticationContext("https://login.windows.net/common");
                if (ctx.TokenCache.Count > 0)
                {
                    string homeTenant = ctx.TokenCache.ReadItems().First().TenantId;
                    ctx = new AuthenticationContext("https://login.microsoftonline.com/" + homeTenant);
                }
            }

            return ctx;
        }

        public ObservableCollection<ProjectVM> GetProjects(string org, AuthenticationHeaderValue authHeader)
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(azureDevOpsOrganizationUrl);
                //client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = authHeader;


                ObservableCollection<ProjectVM> ProjectList = new ObservableCollection<ProjectVM>();
                HttpResponseMessage response = client.GetAsync(org + "/_apis/projects?stateFilter=All&api-version=2.2").Result;


                //check to see if we have a successful response
                if (response.IsSuccessStatusCode)
                {

                    var y = response.Content.ReadAsStringAsync().Result;
                    var x = Newtonsoft.Json.JsonConvert.DeserializeObject<Projects>(y);

                    //Add each Project Item to the Observable Collection as ProjectVM Item
                    foreach (var item in x.value)
                    {
                        ProjectList.Add(new ProjectVM(item.id, item.name, item.description, item.url, item.state));
                    }

                } 
                return ProjectList;
            }
        }

        public ObservableCollection<BoardColumnVM> GetBoardColumns(string org, string projekt) //Connects to DevOps and recieves List of Board Columns of given Project
        {

            //use the httpclient
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(azureDevOpsOrganizationUrl);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = bearerAuthHeader;

                //connect to the REST endpoint            
                HttpResponseMessage response = client.GetAsync(org + "/" + projekt + "/_apis/work/boardcolumns?api-version=5.0").Result;

                ObservableCollection<BoardColumnVM> BoardColumnList = new ObservableCollection<BoardColumnVM>();

                //check to see if we have a successful response
                if (response.IsSuccessStatusCode)
                {

                    var y = response.Content.ReadAsStringAsync().Result;
                    var x = Newtonsoft.Json.JsonConvert.DeserializeObject<BoardColumn>(y);

                    //Add each Project Item to the Observable Collection as BoardColumnVM Item
                    foreach (var item in x.value)
                    {
                        BoardColumnList.Add(new BoardColumnVM(item.name));
                    }

                }
                return BoardColumnList;
            }
        }

        public ObservableCollection<WorkItemVM> GetWorkItems(AuthenticationHeaderValue authHeader, string org, string project)
        {
            //use the httpclient
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(azureDevOpsOrganizationUrl);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = authHeader;


                //connect to the REST endpoint            
                HttpResponseMessage response = client.GetAsync(org + "/" + project + "/_apis/wit/workitemtypes?api-version=5.0").Result;
                ObservableCollection<WorkItemVM> WorkItemList = new ObservableCollection<WorkItemVM>();

                //check to see if we have a successful response
                if (response.IsSuccessStatusCode)
                {

                    var y = response.Content.ReadAsStringAsync().Result;
                    var x = Newtonsoft.Json.JsonConvert.DeserializeObject<WorkItem>(y);

                    //Add each Project Item to the Observable Collection as BoardColumnVM Item
                    foreach (var item in x.value)
                    {
                        WorkItemList.Add(new WorkItemVM(item.name));
                    }

                }
                return WorkItemList;
            }
        }

        public void CreateWorkItem(AuthenticationHeaderValue authHeader, string type, string organization, string project, string title)
        {
            
            string _UrlServiceCreate = $"https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/${type}?api-version=5.0";
            dynamic WorkItem = new List<dynamic>() {
                new
                {
                    op = "add",
                    path = "/fields/System.Title",
                    value = title
                }
            };

            var WorkItemValue = new StringContent(JsonConvert.SerializeObject(WorkItem), Encoding.UTF8, "application/json-patch+json");
            var JsonResultWorkItemCreated = HttpPost(_UrlServiceCreate, bearerAuthHeader, WorkItemValue);
        }


        public static string HttpPost(string urlService, AuthenticationHeaderValue auth, StringContent postValue)
        {
            try
            {
                string request = string.Empty;
                using (HttpClient httpClient = new HttpClient())
                {
                    httpClient.DefaultRequestHeaders.Accept.Clear();
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    httpClient.DefaultRequestHeaders.Authorization = auth;
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(new HttpMethod("POST"), urlService) { Content = postValue })
                    {
                        var httpResponseMessage = httpClient.SendAsync(httpRequestMessage).Result;
                        if (httpResponseMessage.IsSuccessStatusCode)
                            request = httpResponseMessage.Content.ReadAsStringAsync().Result;
                    }
                }
                return request;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }


    }



}
    
