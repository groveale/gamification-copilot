//using Microsoft.Graph;
using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Beta;
using Microsoft.Graph.Beta.Models;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Kiota.Abstractions.Authentication;

namespace groveale.Services
{
    public interface IGraphService
    {
        Task GetTodaysCopilotUsageDataAsync();
        Task<List<M365CopilotUsage>> GetM365CopilotUsageReportAsyncJSON(Microsoft.Extensions.Logging.ILogger _logger);
        Task SetReportAnonSettingsAsync(bool displayConcealedNames);
        Task<AdminReportSettings> GetReportAnonSettingsAsync();

        Task<Dictionary<string, bool>> GetM365CopilotUsersAsync();
    }

    public class GraphService : IGraphService
    {
        private readonly GraphServiceClient _graphServiceClient;
        private DefaultAzureCredential _defaultCredential;
        private ClientSecretCredential _clientSecretCredential;

        public GraphService()
        {
            _defaultCredential = new DefaultAzureCredential();

            // _clientSecretCredential = new ClientSecretCredential(
            //    System.Environment.GetEnvironmentVariable("GRAPH_TENANT_ID"),
            //    System.Environment.GetEnvironmentVariable("GRAPH_CLIENT_ID"),
            //    System.Environment.GetEnvironmentVariable("GRAPH_CLIENT_SECRET"));

            _graphServiceClient = new GraphServiceClient(_defaultCredential,
                // Use the default scope, which will request the scopes
                // configured on the app registration
                new[] { "https://graph.microsoft.com/.default" });
        }


        public async Task GetTodaysCopilotUsageDataAsync()
        {
            var usageData = await _graphServiceClient.Reports
                .GetMicrosoft365CopilotUsageUserDetailWithPeriod("D7")
                .GetAsync();

            // Process the usage data as needed
            Console.WriteLine(usageData.Length);
        }

        // No longer needed
        public async Task<List<User>> GetAllUsersAsync()
        {
            var users = new List<User>();
            var result = await _graphServiceClient.Users.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Select = new string[] { "id", "userPrincipalName" };
            });

            users.AddRange(result.Value);

            while (result.OdataNextLink != null)
            {
                result = await _graphServiceClient.Users.WithUrl(result.OdataNextLink).GetAsync();
                users.AddRange(result.Value);
            }
            return users;
        }

        public async Task<Dictionary<string, bool>> GetM365CopilotUsersAsync()
        {
            var activeUsers = new List<Office365ActiveUserDetail>();

            var urlString = _graphServiceClient.Reports.GetOffice365ActiveUserDetailWithPeriod("D7").ToGetRequestInformation().URI.OriginalString;
            urlString += "?$format=application/json";//append the query parameter

            try
            {
                var activeUsersResponse = await _graphServiceClient.Reports.GetOffice365ActiveUserDetailWithPeriod("D7").WithUrl(urlString).GetAsGetOffice365ActiveUserDetailWithPeriodGetResponseAsync();

                activeUsers.AddRange(activeUsersResponse.Value);

                while (activeUsersResponse.OdataNextLink != null)
                {
                    activeUsersResponse = await _graphServiceClient.Reports.GetOffice365ActiveUserDetailWithPeriod("D7").WithUrl(activeUsersResponse.OdataNextLink).GetAsGetOffice365ActiveUserDetailWithPeriodGetResponseAsync();
                    activeUsers.AddRange(activeUsersResponse.Value);
                }

                var copilotUsers = new Dictionary<string, bool>();
                // find all the user that have a copilot license. let's use lynq to filter the users
                activeUsers.Where(usr => usr.AssignedProducts.Contains("MICROSOFT 365 COPILOT")).ToList().ForEach(usr =>
                {
                    // Check if user has activity in Teams, Outlook or SharePoint
                    if (usr.TeamsLastActivityDate.GetValueOrDefault().DateTime == usr.ReportRefreshDate.GetValueOrDefault().DateTime ||
                        usr.ExchangeLastActivityDate.GetValueOrDefault().DateTime == usr.ReportRefreshDate.GetValueOrDefault().DateTime ||
                        usr.SharePointLastActivityDate.GetValueOrDefault().DateTime == usr.ReportRefreshDate.GetValueOrDefault().DateTime)
                    {
                        // Add the user to the list
                        copilotUsers.Add(usr.UserPrincipalName, true);
                    }

                    // if no usage they are considered offline and not processed (should reduce weekend overhead processing)
                });

                return copilotUsers;

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting copilot users: {ex.Message}");
            }

            return null;
        }



        public async Task<List<M365CopilotUsage>> GetM365CopilotUsageReportAsyncJSON(Microsoft.Extensions.Logging.ILogger _logger)
        {

            var urlString = _graphServiceClient.Reports.GetMicrosoft365CopilotUsageUserDetailWithPeriod("D7").ToGetRequestInformation().URI.OriginalString;
            urlString += "?$format=application/json";//append the query parameter
            // default is top 200 rows, we can use the below to increase this
            //urlString += "?$format=application/json&$top=600";
            var m365CopilotUsageReportResponse = await _graphServiceClient.Reports.GetMicrosoft365CopilotUsageUserDetailWithPeriod("D7").WithUrl(urlString).GetAsync();

            byte[] buffer = new byte[8192];
            int bytesRead;
            List<M365CopilotUsage> m365CopilotUsageReports = new List<M365CopilotUsage>();

            do
            {

                string usageReportsInChunk = "";

                while ((bytesRead = await m365CopilotUsageReportResponse.ReadAsync(buffer, 0, buffer.Length)) > 0)
                {
                    // Process the chunk of data here
                    string chunk = System.Text.Encoding.UTF8.GetString(buffer, 0, bytesRead);
                    usageReportsInChunk += chunk;
                }

                using (JsonDocument doc = JsonDocument.Parse(usageReportsInChunk))
                {
                    // Append the site data to the site dataString
                    if (doc.RootElement.TryGetProperty("value", out JsonElement usageReports))
                    {
                        var reports = JsonSerializer.Deserialize<List<M365CopilotUsage>>(usageReports.GetRawText());
                        m365CopilotUsageReports.AddRange(reports);
                        _logger.LogInformation($"Total User reports: {m365CopilotUsageReports.Count}");

                    }

                    if (doc.RootElement.TryGetProperty("@odata.nextLink", out JsonElement nextLinkElement))
                    {
                        urlString = nextLinkElement.GetString();
                    }
                    else
                    {
                        urlString = null; // No more pages break out of the loop
                        break;
                    }
                }

                m365CopilotUsageReportResponse = await _graphServiceClient.Reports.GetMicrosoft365CopilotUsageUserDetailWithPeriod("D7").WithUrl(urlString).GetAsync();

            } while (urlString != null);


            return m365CopilotUsageReports;
        }

        public async Task SetReportAnonSettingsAsync(bool displayConcealedNames)
        {
            var adminReportSettings = new AdminReportSettings
            {
                DisplayConcealedNames = displayConcealedNames
            };

            var result = await _graphServiceClient.Admin.ReportSettings.PatchAsync(adminReportSettings);
        }

        public async Task<AdminReportSettings> GetReportAnonSettingsAsync()
        {
            var result = await _graphServiceClient.Admin.ReportSettings.GetAsync();
            return result;
        }

    }

}