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
        private readonly IQueueService _queueService;


        public GetTodaysCopilotUsage(ILoggerFactory loggerFactory, IGraphService graphService, ISettingsService settingsService, IAzureTableService azureTableService, IKeyVaultService keyVaultService, IExclusionEmailService exclusionEmailService, IQueueService queueService)
        {
            _graphService = graphService;
            _settingsService = settingsService;
            _azureTableService = azureTableService;
            _keyVaultService = keyVaultService;
            _logger = loggerFactory.CreateLogger<GetTodaysCopilotUsage>();
            _exclusionEmailService = exclusionEmailService;
            _queueService = queueService;
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

            // Inject test data if no usage data and in test mode
            if (usageData.Count == 0 && (Environment.GetEnvironmentVariable("ENABLE_TEST_DATA")?.ToLower() == "true"))
            {
                _logger.LogInformation("No usage data found, but test mode enabled. Injecting test data...");
                var testRefreshDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                usageData["alex@groverale.onmicrosoft.com"] = new M365CopilotUsage
                {
                    UserPrincipalName = "alex@groverale.onmicrosoft.com",
                    ReportRefreshDate = testRefreshDate,
                    LastActivityDate = testRefreshDate
                };
                usageData["testuser@groverale.onmicrosoft.com"] = new M365CopilotUsage
                {
                    UserPrincipalName = "testuser@groverale.onmicrosoft.com",
                    ReportRefreshDate = testRefreshDate,
                    LastActivityDate = testRefreshDate
                };
                _logger.LogInformation($"Injected {usageData.Count} test users");
            }

            // extract the refresh date from the first or default user
            if (usageData.Count == 0)
            {
                // unlikey to ever really happen, but just in case
                _logger.LogWarning("No usage data found.");
                return;
            }

            var firstUser = usageData.Values.FirstOrDefault();
            var reportRefreshDate = firstUser.ReportRefreshDate;
            _logger.LogInformation($"Report refresh date: {reportRefreshDate}");

            var copilotUsers = await _graphService.GetM365CopilotUsersAsync(reportRefreshDate, usageData, _logger);
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

            
            if (firstUser.UserPrincipalName.Contains("groverale"))
            {
                _logger.LogInformation("Adding test user to copilot users.");
                copilotUsers.TryAdd("alex@groverale.onmicrosoft.com", true);
                // We also need to add to the usage data if not already present
                if (!usageData.ContainsKey("alex@groverale.onmicrosoft.com"))
                {
                    usageData["alex@groverale.onmicrosoft.com"] = new M365CopilotUsage
                    {
                        UserPrincipalName = "alex@groverale.onmicrosoft.com",
                        ReportRefreshDate = reportRefreshDate
                    };
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

                // Process user snapshots (handles inactive user tracking and returns encrypted UPNs)
                var encryptedUPNs = await _azureTableService.ProcessUserDailySnapshots(copilotUsageData, encryptionService);
                _logger.LogInformation($"Processed {encryptedUPNs.Count} user snapshots for queuing");

                // Queue user aggregation messages for parallel processing
                await _queueService.QueueUserAggregationsAsync(encryptedUPNs, reportRefreshDate);

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

        private List<M365CopilotUsage> RemoveNonCopilotUsersAndOfflineUsers(Dictionary<string, M365CopilotUsage> usageData, Dictionary<string, bool> copilotUsers, HashSet<string> exclusionEmails)
        {
            var filteredUsageData = new List<M365CopilotUsage>();

            foreach (var u in usageData.Values)
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
