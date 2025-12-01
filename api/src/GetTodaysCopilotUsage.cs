using System;
using groveale;
using groveale.Services;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace groverale
{
    public class GetTodaysCopilotUsage
    {
        private readonly ILogger _logger;
        private readonly IGraphService _graphService;
        private readonly ISettingsService _settingsService;
        private readonly IAzureTableService _azureTableService;
        private readonly IKeyVaultService _keyVaultService;
        private readonly IExclusionEmailService _exclusionEmailService;


        public GetTodaysCopilotUsage(ILoggerFactory loggerFactory, IGraphService graphService, ISettingsService settingsService, IAzureTableService azureTableService, IKeyVaultService keyVaultService, IExclusionEmailService exclusionEmailService)
        {
            _graphService = graphService;
            _settingsService = settingsService;
            _azureTableService = azureTableService;
            _keyVaultService = keyVaultService;
            _logger = loggerFactory.CreateLogger<GetTodaysCopilotUsage>();
            _exclusionEmailService = exclusionEmailService;
        }

        [Function("GetTodaysCopilotUsage")]
        public async Task Run([TimerTrigger("0 0 3 * * *")] TimerInfo myTimer)
        {
            _logger.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            if (myTimer.ScheduleStatus is not null)
            {
                _logger.LogInformation($"Next timer schedule at: {myTimer.ScheduleStatus.Next}");
            }

            // // Get the report settings
            // var reportSettings = await _graphService.GetReportAnonSettingsAsync();

            // // Set the report settings to display names (if hidden)
            // if (reportSettings.DisplayConcealedNames.Value)
            // {
            //     _logger.LogInformation("Setting report settings to display names...");
            //     await _graphService.SetReportAnonSettingsAsync(false);
            // }
            // else
            // {
            //     _logger.LogInformation("Report settings already set to display names.");
            // }

            // Get the usage data
            var usageData = await _graphService.GetM365CopilotUsageReportAsyncJSON(_logger);
            _logger.LogInformation($"Usage data: {usageData.Count}");

            // extract the refresh date from the first or default user
            if (usageData.Count == 0)
            {
                // unlikey to ever really happen, but just in case
                _logger.LogWarning("No usage data found.");
                return;
            }

            var firstUser = usageData.FirstOrDefault();
            var reportRefreshDate = firstUser.ReportRefreshDate;
            _logger.LogInformation($"Report refresh date: {reportRefreshDate}");

            var copilotUsers = await _graphService.GetM365CopilotUsersAsync(reportRefreshDate, _logger);
            _logger.LogInformation($"copilot users: {copilotUsers.Count}");

            // If the copilot users are empty, we can fallback to the old method
            if (copilotUsers.Count == 0)
            {
                _logger.LogWarning("No copilot users found, falling back to old method.");
                copilotUsers = await _graphService.GetM365CopilotUserFallBackAsync();
                _logger.LogInformation($"copilot users (fallback): {copilotUsers.Count}");
            }

            // // Set the report settings back to hide names (if hidden)
            // if (reportSettings.DisplayConcealedNames.Value)
            // {
            //     _logger.LogInformation("Setting report settings back to hide names...");
            //     await _graphService.SetReportAnonSettingsAsync(true);
            // }

            // Just call it - caching happens automatically
            var exclusionEmails = await _exclusionEmailService.LoadEmailsFromPersonFieldAsync();
            _logger.LogInformation($"Exclusion emails loaded: {exclusionEmails.Count}");

            // TODO test: adding a user to the Copilot users
            if (firstUser.UserPrincipalName.Contains("groverale"))
            {
                _logger.LogInformation("Adding test user to copilot users.");
                copilotUsers.TryAdd("alexg@groverale.onmicrosoft.com", true);
                // We also need to add to the usage data if not already present
                if (!usageData.Any(u => u.UserPrincipalName == "alexg@groverale.onmicrosoft.com"))
                {
                    usageData.Add(new M365CopilotUsage
                    {
                        UserPrincipalName = "alexg@groverale.onmicrosoft.com",
                        ReportRefreshDate = reportRefreshDate
                    });
                }
                        
            }


            // Remove snapshot for non-copilot users and offline users
            // Offline users have no activity in either Teams or Outlook for the day - Considered Out of Office so are skipped to allow streaks to continue
            var copilotUsageData = RemoveNonCopilotUsersAndOfflineUsers(usageData, copilotUsers, exclusionEmails);
            _logger.LogInformation($"copilot users usage data: {copilotUsageData.Count}");

            try
            {
                // Process the usage data
                if (copilotUsageData.Count == 0)
                {
                    _logger.LogInformation("No usage data to process.");
                    return;
                }

                _logger.LogInformation($"Processing usage data for {copilotUsageData.Count} users...");

                _logger.LogInformation($"Report refresh date: {reportRefreshDate}");

                // Create Encyption Service
                var encryptionService = await DeterministicEncryptionService.CreateAsync(_settingsService, _keyVaultService);

                var recordsAdded = await _azureTableService.ProcessUserDailySnapshots(copilotUsageData, encryptionService);
                _logger.LogInformation($"Records added: {recordsAdded}");

                // aggregate agent usage data - 1 row per agent (weekly, monthly, all time
                // We will do this from the previous day's agent aggregatations
                // we can use today - 1 day as the date to process
                var dateToProcess = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                var agentRecordsAdded = await _azureTableService.ProcessAgentUsageAggregationsAsync(dateToProcess);
                _logger.LogInformation($"Agent records added: {agentRecordsAdded}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing usage data. {Message}", ex.Message);
            }

        }

        private List<M365CopilotUsage> RemoveNonCopilotUsersAndOfflineUsers(List<M365CopilotUsage> usageData, Dictionary<string, bool> copilotUsers, HashSet<string> exclusionEmails)
        {
            var filteredUsageData = new List<M365CopilotUsage>();

            foreach (var u in usageData)
            {
                if (copilotUsers.ContainsKey(u.UserPrincipalName))
                {
                    if (_settingsService.IsEmailListExclusive)
                    {
                        // Inclusion list - only include users in the list
                        if (exclusionEmails.Contains(u.UserPrincipalName))
                        {
                            filteredUsageData.Add(u);
                        }
                        else
                        {
                            _logger.LogInformation($"Skipping user {u.UserPrincipalName} as they are not in the inclusion list.");
                        }
                    }
                    else
                    {
                        // Exclusion list - skip users in the list
                        if (exclusionEmails.Contains(u.UserPrincipalName))
                        {
                            _logger.LogInformation($"Skipping user {u.UserPrincipalName} as they are in the exclusion list.");
                            continue;
                        }
                        filteredUsageData.Add(u);
                    }
                }
            }

            return filteredUsageData;
        }
    }
}
