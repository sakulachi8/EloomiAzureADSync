using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using System.Net.Http;
using System.Collections.Generic;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using RestSharp;
using Microsoft.Azure.WebJobs.Extensions.Http;
using System.Net;

namespace EloomiUsersSync
{
    public static class EloomiFunction
    {
        static readonly HttpClient client = new HttpClient();
        private static string activeDirectoryGroupId = "";
        private static string activeDirectoryTenantId = "";
        private static string activeDirectoryClientId = "";
        private static string activeDirectoryClientSecretId = "";
        private static string eloomiClientId = "";
        private static string eloomiClientSecret = "";
        private static string eloomiToken = "";
        [FunctionName("EloomiFunction")]
        // public static async Task RunAsync([TimerTrigger("0 */30 * * * *")]TimerInfo myTimer, TraceWriter log)
        // {
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info($"C# Timer trigger function executed at: {DateTime.Now}");
            string activeDirectoryToken = await GetActiveDirectoryToken();
            eloomiToken = await GetEloomiTokenAsync();
            if (activeDirectoryToken != null && eloomiToken != null)
            {
                try
                {
                    List<string> usersDeprovisionList = new List<string>();
                    List<Dictionary<string, dynamic>> usersProvisionList = new List<Dictionary<string, dynamic>>();
                    List<Dictionary<string, dynamic>> provisionUsersWithInActiveStatus = new List<Dictionary<string, dynamic>>();
                    List<Dictionary<string, dynamic>> updateUsersList = new List<Dictionary<string, dynamic>>();
                    List<Dictionary<string, dynamic>> activateUsersList = new List<Dictionary<string, dynamic>>();
                    List<Dictionary<string, dynamic>> usersEloomiList = new List<Dictionary<string, dynamic>>();
                    List<Dictionary<string, dynamic>> usersInActiveEloomiList = new List<Dictionary<string, dynamic>>();
                    Dictionary<string, dynamic> adUsersDict = new Dictionary<string, dynamic>();

                    // get all users from specific group of Active Directory
                    bool allDataRetrieve = false;
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", activeDirectoryToken);
                    string link = "https://graph.microsoft.com/v1.0/groups/" + activeDirectoryGroupId + "/members?$select=accountEnabled,userPrincipalName,givenName,surname,jobTitle,mobilePhone,department,onPremisesExtensionAttributes";
                    while (!allDataRetrieve)
                    {
                        HttpResponseMessage response = await client.GetAsync(link);
                        if (response.IsSuccessStatusCode)
                        {
                            Dictionary<string, dynamic> tempValue = await response.Content.ReadAsAsync<Dictionary<string, dynamic>>();
                            if (tempValue.ContainsKey("@odata.nextLink"))
                            {
                                link = tempValue["@odata.nextLink"];
                            }
                            else
                            {
                                allDataRetrieve = true;
                            }


                            List<Dictionary<string, dynamic>> adResponse = tempValue["value"].ToObject<List<Dictionary<string, dynamic>>>();
                            foreach (Dictionary<string, dynamic> userObj in adResponse)
                            {
                                HttpResponseMessage managerData = await client.GetAsync("https://graph.microsoft.com/v1.0/users/" + userObj["userPrincipalName"] + "/manager?$select=accountEnabled,userPrincipalName,givenName,surname,jobTitle,mobilePhone,department,onPremisesExtensionAttributes");
                                if (managerData.IsSuccessStatusCode)
                                {
                                    Dictionary<string, dynamic> userData = await managerData.Content.ReadAsAsync<Dictionary<string, dynamic>>();
                                    Dictionary<string, dynamic> userTempObj = new Dictionary<string, dynamic>(userObj);
                                    userTempObj.Add("managerData", userData);
                                    adUsersDict.Add(userObj["userPrincipalName"].ToLower(), userTempObj);
                                }
                            }
                        }
                    }
                    // Get all users from Eloomi 
                    client.DefaultRequestHeaders.Clear();
                    client.DefaultRequestHeaders.Add("ClientId", eloomiClientId);
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + eloomiToken);
                    HttpResponseMessage responseEloomi = await client.GetAsync("https://api.eloomi.com/v3/users?mode=active");
                    if (responseEloomi.IsSuccessStatusCode)
                    {
                        usersEloomiList.AddRange((await responseEloomi.Content.ReadAsAsync<Dictionary<string, dynamic>>())["data"].ToObject<List<Dictionary<string, dynamic>>>());
                    }

                    responseEloomi = await client.GetAsync("https://api.eloomi.com/v3/users?mode=inactive");
                    if (responseEloomi.IsSuccessStatusCode)
                    {
                        usersInActiveEloomiList.AddRange((await responseEloomi.Content.ReadAsAsync<Dictionary<string, dynamic>>())["data"].ToObject<List<Dictionary<string, dynamic>>>());
                    }

                    // Check if user is inactive or new user is provisioned
                    foreach (KeyValuePair<string, dynamic> item in adUsersDict)
                    {
                        bool isUserAlreadyPresent = usersEloomiList.Exists(obj => obj["email"].ToLower() == item.Value["userPrincipalName"].ToLower());
                        bool isUserPresentInActive = usersInActiveEloomiList.Exists(obj => obj["email"].ToLower() == item.Value["userPrincipalName"].ToLower());


                        if (!usersEloomiList.Exists(obj => obj["email"].ToLower() == item.Value["managerData"]["userPrincipalName"].ToLower()) && !usersInActiveEloomiList.Exists(obj => obj["email"].ToLower() == item.Value["managerData"]["userPrincipalName"].ToLower())
                            && provisionUsersWithInActiveStatus.Contains(item.Value["managerData"]))
                        {
                            provisionUsersWithInActiveStatus.Add(item.Value["managerData"]);
                        }


                        if (isUserAlreadyPresent)
                        {
                            if (!(bool)item.Value["accountEnabled"] && !usersDeprovisionList.Contains(item.Key))
                            {
                                usersDeprovisionList.Add(item.Key);
                            }
                            else if (!updateUsersList.Contains(item.Value))
                            {
                                updateUsersList.Add(item.Value);
                            }
                        }
                        else if (isUserPresentInActive)
                        {
                            if ((bool)item.Value["accountEnabled"] && !activateUsersList.Contains(item.Value))
                            {
                                activateUsersList.Add(item.Value);
                            }
                        }
                        else if (!isUserAlreadyPresent && !usersProvisionList.Contains(item.Value))
                        {
                            usersProvisionList.Add(item.Value);
                        }
                    }

                    // check if eloomi user is provisioned and deleted from active directory
                    foreach (Dictionary<string, dynamic> userObj in usersEloomiList)
                    {
                        if (!adUsersDict.ContainsKey(userObj["email"].ToLower()) && !usersDeprovisionList.Contains(userObj["email"]))
                        {
                            usersDeprovisionList.Add(userObj["email"]);
                        }
                    }








                    // Now perform actions on lists


                    // User provision with inactive status
                    foreach (Dictionary<string, dynamic> userObj in provisionUsersWithInActiveStatus)
                    {
                        bool status = await ProvisionUserOnEloomiWithInactiveStatusAsync(userObj);
                        if (status)
                        {
                            log.Info("New User is provisioned with email: " + userObj["userPrincipalName"]);
                        }
                        else
                        {
                            log.Info("Issue when a new user is provisioned with email: " + userObj["userPrincipalName"]);

                        }
                    }

                    // Provision user if any exist in list 
                    foreach (Dictionary<string, dynamic> userObj in usersProvisionList)
                    {
                        bool status = await ProvisionUserOnEloomiAsync(userObj);
                        if (status)
                        {
                            log.Info("New User is provisioned with email: " + userObj["userPrincipalName"]);
                        }
                        else
                        {
                            log.Info("Issue when a new user is provisioned with email: " + userObj["userPrincipalName"]);

                        }
                    }


                    //deprovision user if any exist in list 
                    foreach (string userEmail in usersDeprovisionList)
                    {
                        bool status = DeProvisionUserOnEloomi(userEmail.ToLower());
                        if (status)
                        {
                            log.Info("New User is de-provisioned with email: " + userEmail);
                        }
                        else
                        {
                            log.Info("Issue when a new user is de-provisioned with email: " + userEmail);
                        }
                    }


                    // Activate User Already Exist
                    foreach (Dictionary<string, dynamic> userData in activateUsersList)
                    {
                        bool status = await ActivateUserOnEloomiAsync(userData);
                        if (status)
                        {
                            log.Info("User Activated with email: " + userData["userPrincipalName"]);
                        }
                        else
                        {
                            log.Info("Issue when user is activated with email: " + userData["userPrincipalName"]);
                        }
                    }


                    //Update User if any exist in list 
                    foreach (Dictionary<string, dynamic> userData in updateUsersList)
                    {

                        bool status = await UpdateUserOnEloomiAsync(userData);
                        if (status)
                        {
                            log.Info("user data is updated with email: " + userData["userPrincipalName"]);
                        }
                        else
                        {
                            log.Info("Issue when an user is updated with email: " + userData["userPrincipalName"]);
                        }
                    }



                }
                catch (HttpRequestException e)
                {
                    log.Info("\nException Caught!");
                    log.Info("Message :{0} ", e.Message);
                }
            }
            return req.CreateResponse(HttpStatusCode.OK);
        }



        public static async Task<string> GetActiveDirectoryToken()
        {
            string result = null;
            Dictionary<string, string> jsonData = new Dictionary<string, string>()
            {
                { "grant_type","client_credentials"},
                { "client_id",activeDirectoryClientId},
                { "client_secret",activeDirectoryClientSecretId},
                { "resource","https://graph.microsoft.com"}
            };
            HttpResponseMessage responseActiveDirectory = await client.PostAsync("https://login.microsoftonline.com/" + activeDirectoryTenantId + "/oauth2/token", new FormUrlEncodedContent(jsonData));
            if (responseActiveDirectory.IsSuccessStatusCode)
            {
                result = (await responseActiveDirectory.Content.ReadAsAsync<Dictionary<string, dynamic>>())["access_token"];
            }
            return result;
        }


        public static async Task<string> GetEloomiTokenAsync()
        {
            string result = null;
            Dictionary<string, dynamic> jsonData = new Dictionary<string, dynamic>()
            {
                { "grant_type","client_credentials"},
                { "client_id",eloomiClientId},
                { "client_secret",eloomiClientSecret},
            };

            HttpResponseMessage responseActiveDirectory = await client.PostAsJsonAsync("https://api.eloomi.com/oauth/token", jsonData);
            if (responseActiveDirectory.IsSuccessStatusCode)
            {
                result = (await responseActiveDirectory.Content.ReadAsAsync<Dictionary<string, dynamic>>())["access_token"];
            }
            return result;
        }


        public static async Task<bool> ProvisionUserOnEloomiAsync(Dictionary<string, dynamic> userObj)
        {
            bool result = false;
            Dictionary<string, dynamic> jsonData = new Dictionary<string, dynamic>()
            {
                { "email",userObj["userPrincipalName"].ToLower()},
                { "first_name",userObj["givenName"]},
                { "last_name",userObj["surname"]},
                { "title",userObj["jobTitle"]},
                { "phone",userObj["mobilePhone"]},
                { "activate",true}
            };

            if (userObj.ContainsKey("onPremisesExtensionAttributes") && userObj["onPremisesExtensionAttributes"].ContainsKey("extensionAttribute1") && userObj["onPremisesExtensionAttributes"]["extensionAttribute1"] != null)
            {
                Dictionary<string, dynamic> keyValuePairs = userObj["onPremisesExtensionAttributes"].ToObject<Dictionary<string, dynamic>>();
                if (keyValuePairs.ContainsKey("extensionAttribute1") && keyValuePairs["extensionAttribute1"] != null)
                {
                    jsonData.Add("department_id", "[\"" + await GetDepartIdFromEloomiAsync(keyValuePairs["extensionAttribute1"])  + "\"]" );
                }
            }

            if (userObj["managerData"]["userPrincipalName"] != null)
            {
                jsonData.Add("direct_manager_ids", "[\"" + await GetEloomiUserIdFromEmail(userObj["managerData"]["userPrincipalName"]) + "\"]");
            }

            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Add("ClientId", eloomiClientId);
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + eloomiToken);
            HttpResponseMessage responseActiveDirectory = await client.PostAsJsonAsync("https://api.eloomi.com/v3/users", jsonData);
            if (responseActiveDirectory.IsSuccessStatusCode)
            {
                result = true;
            }
            return result;
        }



        public static async Task<long> GetDepartIdFromEloomiAsync(string name)
        {
            long id = 0;
            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Add("ClientId", eloomiClientId);
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + eloomiToken);
            HttpResponseMessage response = await client.GetAsync("https://api.eloomi.com/v3/units/");
            if (response.IsSuccessStatusCode)
            {
                List<Dictionary<string, dynamic>> departmentList = (await response.Content.ReadAsAsync<Dictionary<string, dynamic>>())["data"].ToObject<List<Dictionary<string, dynamic>>>();
                bool isPresent = departmentList.Exists(obj => obj["name"].ToLower() == name.ToLower());
                if (isPresent)
                {
                    id = departmentList[departmentList.FindIndex(obj => obj["name"].ToLower() == name.ToLower())]["id"];
                }
                else
                {
                    Dictionary<string, dynamic> jsonData = new Dictionary<string, dynamic>()
                    {
                        { "name",name.ToLower()}
                    };
                    HttpResponseMessage responseMessage = await client.PostAsJsonAsync("https://api.eloomi.com/v3/units/", jsonData);
                    if (responseMessage.IsSuccessStatusCode)
                    {
                        id = (await responseMessage.Content.ReadAsAsync<Dictionary<string, dynamic>>())["data"].ToObject<Dictionary<string, dynamic>>()["id"];
                    }
                }
            }
            return id;
        }



        public static async Task<long> GetEloomiUserIdFromEmail(string email)
        {
            long id = 0;
            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Add("ClientId", eloomiClientId);
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + eloomiToken);
            HttpResponseMessage response = await client.GetAsync("https://api.eloomi.com/v3/users-email/" + email.ToLower());
            if (response.StatusCode == HttpStatusCode.OK)
            {
                id = (await response.Content.ReadAsAsync<Dictionary<string, dynamic>>())["data"].ToObject<Dictionary<string, dynamic>>()["id"];
            }
            return id;
        }


        public static async Task<bool> ProvisionUserOnEloomiWithInactiveStatusAsync(Dictionary<string, dynamic> userObj)
        {
            bool result = false;
            Dictionary<string, dynamic> jsonData = new Dictionary<string, dynamic>()
            {
                { "email",userObj["userPrincipalName"].ToLower()},
                { "first_name",userObj["givenName"]},
                { "last_name",userObj["surname"]},
                { "title",userObj["jobTitle"]},
                { "phone",userObj["mobilePhone"]},
                { "activate","deactivate"}
            };
            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Add("ClientId", eloomiClientId);
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + eloomiToken);
            HttpResponseMessage responseActiveDirectory = await client.PostAsJsonAsync("https://api.eloomi.com/v3/users", jsonData);
            if (responseActiveDirectory.IsSuccessStatusCode)
            {
                result = true;
            }
            return result;
        }


        public static async Task<bool> UpdateUserOnEloomiAsync(Dictionary<string, dynamic> userObj)
        {
            Dictionary<string, dynamic> jsonData = new Dictionary<string, dynamic>()
            {
                { "first_name",userObj["givenName"]},
                { "last_name",userObj["surname"]},
                { "title",userObj["jobTitle"]},
                { "phone",userObj["mobilePhone"]}
            };

            if (userObj.ContainsKey("onPremisesExtensionAttributes") && userObj["onPremisesExtensionAttributes"] != null)
            {
                Dictionary<string, dynamic> keyValuePairs = userObj["onPremisesExtensionAttributes"].ToObject<Dictionary<string, dynamic>>();
                if (keyValuePairs.ContainsKey("extensionAttribute1") && keyValuePairs["extensionAttribute1"] != null)
                {
                    jsonData.Add("department_id", new long[] { await GetDepartIdFromEloomiAsync(keyValuePairs["extensionAttribute1"]) });
                }
            }

            if (userObj["managerData"]["userPrincipalName"] != null)
            {
                jsonData.Add("direct_manager_ids", new long[] { await GetEloomiUserIdFromEmail(userObj["managerData"]["userPrincipalName"]) });
            }
            return PatchAsync("https://api.eloomi.com/v3/users-email/" + userObj["userPrincipalName"].ToLower(), jsonData);

        }


        public static async Task<bool> ActivateUserOnEloomiAsync(Dictionary<string, dynamic> userObj)
        {
            Dictionary<string, dynamic> jsonData = new Dictionary<string, dynamic>()
            {
                { "first_name",userObj["givenName"]},
                { "last_name",userObj["surname"]},
                { "title",userObj["jobTitle"]},
                { "phone",userObj["mobilePhone"]},
                { "activate","instant"}
            };

            if (userObj.ContainsKey("onPremisesExtensionAttributes") && userObj["onPremisesExtensionAttributes"] != null)
            {
                Dictionary<string, dynamic> keyValuePairs = userObj["onPremisesExtensionAttributes"].ToObject<Dictionary<string, dynamic>>();
                if (keyValuePairs.ContainsKey("extensionAttribute1") && keyValuePairs["extensionAttribute1"] != null)
                {
                    jsonData.Add("department_id", new long[] { await GetDepartIdFromEloomiAsync(keyValuePairs["extensionAttribute1"]) });
                }
            }

            if (userObj["managerData"]["userPrincipalName"] != null)
            {
                jsonData.Add("direct_manager_ids", new long[] { await GetEloomiUserIdFromEmail(userObj["managerData"]["userPrincipalName"]) });
            }
            return PatchAsync("https://api.eloomi.com/v3/users-email/" + userObj["userPrincipalName"].ToLower(), jsonData);

        }


        public static bool DeProvisionUserOnEloomi(string userEmail)
        {
            Dictionary<string, dynamic> jsonData = new Dictionary<string, dynamic>()
            {
                { "activate","deactivate"}
            };
            return PatchAsync("https://api.eloomi.com/v3/users-email/" + userEmail.ToLower(), jsonData);
        }


        public static bool PatchAsync(string requestUri, Dictionary<string, dynamic> jsonData)
        {
            var clientTemp = new RestClient(requestUri);
            var request = new RestRequest(Method.PATCH);
            request.AddHeader("Authorization", "Bearer " + eloomiToken);
            request.AddHeader("ClientId", eloomiClientId);
            request.AddHeader("content-type", "application/json");
            request.AddJsonBody(jsonData);
            IRestResponse response = clientTemp.Execute(request);
            if (response.IsSuccessful)
            {
                return true;
            }
            return false;
        }

    }
}
