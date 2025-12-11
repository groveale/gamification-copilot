using Azure;
using Azure.Data.Tables;
using Azure.Identity;
using groveale.Models;
using Microsoft.Extensions.Logging;
using System;
using System.Linq.Expressions;
using System.Text.Json;
using System.Threading.Tasks;

namespace groveale.Services
{
    public interface IAzureTableService
    {
        Task AddListCreationRecordAsync(ListAuditObj entity);
        Task AddCopilotInteractionRAWAysnc(dynamic entity);
        Task AddCopilotInteractionDetailsAsync(AuditData entity);
        Task AddCopilotInteractionDailyAggregationForUserAsync(List<CopilotEventData> entity, string userId);
        Task AddSingleCopilotInteractionDailyAggregationForUserAsync(CopilotEventData entity, string userId);
        Task AddSpecificCopilotInteractionDailyAggregationForUserAsync(AppType appType, string userId, int count);
        Task AddAgentInteractionsDailyAggregationForUserAsync(string agentId, string userId, int count, string agentName);
        Task LogWebhookTriggerAsync(LogEvent webhookEvent);
        Task<bool> GetWebhookStateAsync();
        Task SetWebhookStateAsync(bool isPaused);
        Task RotateKeyRecentDailyTableValues(TableClient tableClient, DeterministicEncryptionService existingKeyService, DeterministicEncryptionService newKeyService);
        Task RotateKeyTimeFrameTableValues(TableClient tableClient, DeterministicEncryptionService existingKeyService, DeterministicEncryptionService newKeyService);



        TableClient GetCopilotAgentInteractionTableClient();
        TableClient GetCopilotInteractionDailyAggregationsTableClient();
        TableClient GetWeeklyCopilotInteractionTableClient();
        TableClient GetMonthlyCopilotInteractionTableClient();
        TableClient GetAllTimeCopilotInteractionTableClient();
        TableClient GetWeeklyAgentInteractionTableClient();
        TableClient GetMonthlyAgentInteractionTableClient();
        TableClient GetAllTimeAgentInteractionTableClient();

        // from other fucntion
        Task<int> ProcessUserDailySnapshots(List<M365CopilotUsage> siteSnapshots, DeterministicEncryptionService encryptionService);
        Task<int> ProcessAgentUsageAggregationsAsync(string dateToProcess);
        Task<List<string>> GetUsersWithStreakForApp(string app, int count);
        Task<string?> GetStartDate(string timeFrame);
        Task<List<string>> GetUsersWhoHaveCompletedActivityForApp(string app, string? dayCount, string? interactionCount, string timeFrame, string? startDate);
        Task<List<InactiveUser>> GetInactiveUsers(int days);

        // User Seeding
        Task SeedMonthlyTimeFrameActivitiesAsync(List<CopilotTimeFrameUsage> userActivitiesSeed, string startDate);
        Task SeedWeeklyTimeFrameActivitiesAsync(List<CopilotTimeFrameUsage> userActivitiesSeed, string startDate);
        Task SeedAllTimeActivityAsync(List<CopilotTimeFrameUsage> userActivitiesSeed);
        Task SeedInactiveUsersAsync(List<InactiveUser> inactiveUsersSeed);


    }

    public class AzureTableService : IAzureTableService
    {
        private readonly TableServiceClient _serviceClient;
        private readonly string _listCreationTable = "ListCreationEvents";
        private readonly string _copilotInteractionTable = "RAWCopilotInteractions";
        private readonly string _copilotInteractionDetailsTable = "CopilotInteractionDetails";
        private readonly string _copilotInteractionDailyAggregationForUserTable = "CopilotInteractionDailyAggregationByAppAndUser";
        private readonly string _agentInteractionDailyAggregationForUserTable = "AgentInteractionDailyAggregationByUserAndAgentId";
        private readonly string _auditWebhookFunctionTable = "WebhookFunctionState";
        private readonly string _unhandledAppHosts = "UnhandledAppHosts";
        private readonly string _reportRefreshDateTableName = "ReportRefreshRecord";
        private readonly string _userWeeklyTableName = "CopilotUsageWeeklySnapshots";
        private readonly string _agentWeeklyTableName = "AgentUsageWeeklySnapshots";
        private readonly string _agentWeeklyByUserTableName = "AgentUsageWeeklyByUserSnapshots";
        private readonly string _userMonthlyTableName = "CopilotUsageMonthlySnapshots";
        private readonly string _agentMonthlyTableName = "AgentUsageMonthlySnapshots";
        private readonly string _agentMonthlyByUserTableName = "AgentUsageMonthlyByUserSnapshots";
        private readonly string _userAllTimeTableName = "CopilotUsageAllTimeRecord";
        private readonly string _agentAllTimeTableName = "AgentUsageAllTimeRecord";
        private readonly string _agentAllTimeByUserTableName = "AgentUsageAllTimeByUserRecord";
        private readonly string _userLastUsageTableName = "UsersLastUsageTracker";
        private readonly TableClient copilotInteractionDetailsTableClient;
        private readonly TableClient copilotInteractionDailyAggregationsTableClient;
        private readonly TableClient copilotAgentInteractionTableClient;
        private readonly TableClient auditWebhookStateTableClient;
        private readonly TableClient unhandledAppHostsTableClient;
        private readonly TableClient _reportRefreshDateTableClient;
        private readonly TableClient _userWeeklyTableClient;
        private readonly TableClient _userMonthlyTableClient;
        private readonly TableClient _userAllTimeTableClient;
        private readonly TableClient _agentWeeklyTableClient;
        private readonly TableClient _agentMonthlyTableClient;
        private readonly TableClient _agentAllTimeTableClient;
        private readonly TableClient _agentWeeklyByUserTableClient;
        private readonly TableClient _agentMonthlyByUserTableClient;
        private readonly TableClient _agentAllTimeByUserTableClient;
        private readonly TableClient _userLastUsageTableClient;
        private readonly string _webhookEventsTable = "WebhookTriggerEvents";
        private readonly ILogger<AzureTableService> _logger;

        private readonly int _daysToCheck = int.TryParse(System.Environment.GetEnvironmentVariable("ReminderDays"), out var days) ? days : 0;

        public AzureTableService(ISettingsService settingsService, ILogger<AzureTableService> logger)
        {
            _serviceClient = new TableServiceClient(
                new Uri(settingsService.StorageAccountUri),
                new DefaultAzureCredential());

            _logger = logger;

            copilotInteractionDetailsTableClient = _serviceClient.GetTableClient(_copilotInteractionDetailsTable);
            copilotInteractionDetailsTableClient.CreateIfNotExists();

            copilotInteractionDailyAggregationsTableClient = _serviceClient.GetTableClient(_copilotInteractionDailyAggregationForUserTable);
            copilotInteractionDailyAggregationsTableClient.CreateIfNotExists();

            copilotAgentInteractionTableClient = _serviceClient.GetTableClient(_agentInteractionDailyAggregationForUserTable);
            copilotAgentInteractionTableClient.CreateIfNotExists();

            auditWebhookStateTableClient = _serviceClient.GetTableClient(_auditWebhookFunctionTable);
            auditWebhookStateTableClient.CreateIfNotExists();

            unhandledAppHostsTableClient = _serviceClient.GetTableClient(_unhandledAppHosts);
            unhandledAppHostsTableClient.CreateIfNotExists();

            _reportRefreshDateTableClient = _serviceClient.GetTableClient(_reportRefreshDateTableName);
            _reportRefreshDateTableClient.CreateIfNotExists();

            _userWeeklyTableClient = _serviceClient.GetTableClient(_userWeeklyTableName);
            _userWeeklyTableClient.CreateIfNotExists();
            _userMonthlyTableClient = _serviceClient.GetTableClient(_userMonthlyTableName);
            _userMonthlyTableClient.CreateIfNotExists();
            _userAllTimeTableClient = _serviceClient.GetTableClient(_userAllTimeTableName);
            _userAllTimeTableClient.CreateIfNotExists();

            _agentWeeklyTableClient = _serviceClient.GetTableClient(_agentWeeklyTableName);
            _agentWeeklyTableClient.CreateIfNotExists();
            _agentMonthlyTableClient = _serviceClient.GetTableClient(_agentMonthlyTableName);
            _agentMonthlyTableClient.CreateIfNotExists();
            _agentAllTimeTableClient = _serviceClient.GetTableClient(_agentAllTimeTableName);
            _agentAllTimeTableClient.CreateIfNotExists();

            _agentWeeklyByUserTableClient = _serviceClient.GetTableClient(_agentWeeklyByUserTableName);
            _agentWeeklyByUserTableClient.CreateIfNotExists();
            _agentMonthlyByUserTableClient = _serviceClient.GetTableClient(_agentMonthlyByUserTableName);
            _agentMonthlyByUserTableClient.CreateIfNotExists();
            _agentAllTimeByUserTableClient = _serviceClient.GetTableClient(_agentAllTimeByUserTableName);
            _agentAllTimeByUserTableClient.CreateIfNotExists();

            _userLastUsageTableClient = _serviceClient.GetTableClient(_userLastUsageTableName);
            _userLastUsageTableClient.CreateIfNotExists();


            _daysToCheck = int.TryParse(settingsService.ReminderDays, out var days) ? days : 14;
        }

        public async Task AddCopilotInteractionRAWAysnc(dynamic entity)
        {
            var tableClient = _serviceClient.GetTableClient(_copilotInteractionTable);
            tableClient.CreateIfNotExists();

            // Ensure the creationTime is specified as UTC
            DateTime eventTime = DateTime.SpecifyKind(DateTime.UtcNow, DateTimeKind.Utc);
            var tableEntity = new TableEntity(eventTime.ToString("yyyy-MM-dd"), Guid.NewGuid().ToString())
            {
                { "Data", entity.ToString() }
            };

            try
            {
                await tableClient.AddEntityAsync(tableEntity);
                _logger.LogInformation($"Added copilot interaction event at {eventTime}");
            }
            catch (Azure.RequestFailedException ex) when (ex.Status == 409) // Conflict indicates the entity already exists
            {
                // Merge the entity if it already exists
                await tableClient.UpdateEntityAsync(tableEntity, ETag.All, TableUpdateMode.Merge);
            }
            catch (RequestFailedException ex)
            {
                // Handle the exception as needed
                _logger.LogError(ex, "Error adding copilot interaction event to table storage.");
                throw;
            }
        }

        public async Task AddCopilotInteractionDetailsAsync(AuditData copilotInteraction)
        {

            // Get values with null checks
            // string userId = copilotInteraction["UserId"]?.ToString() ?? "unknown";
            // string appHostFromDict = copilotInteraction["CopilotEventData"]?["AppHost"]?.ToString() ?? "unknown";
            // string creationTimeFromDict = copilotInteraction["CreationTime"].ToString() ?? DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");

            // // Create entity with validation
            // DateTime eventTime = DateTime.SpecifyKind(DateTime.UtcNow, DateTimeKind.Utc);
            // var tableEntity = new TableEntity(eventTime.ToString("yyyy-MM-dd"), Guid.NewGuid().ToString())
            // {
            //     { "User", userId },
            //     { "AppHost", appHostFromDict },
            //     { "CreationTime", DateTime.SpecifyKind(DateTime.Parse(creationTimeFromDict), DateTimeKind.Utc) }
            // };

            // Extract types from the Contexts list and join them as a comma-separated string
            // AppChat will have a context type of rhe name of the app or service within context
            // Example: Some examples of supported apps and services include M365 Office (docx, pptx, xlsx), TeamsMeeting, TeamsChannel, and TeamsChat. If Copilot is used in Excel, then context will be the identifier of the Excel Spreadsheet and the file type.
            // Check if Contexts is non-null before accessing it
            string contextsTypes = copilotInteraction.CopilotEventData.Contexts != null
                ? string.Join(", ", copilotInteraction.CopilotEventData.Contexts.Select(context => context.Type))
                : string.Empty;

            // Check if AISystemPlugin is non-null before accessing it
            string aiPlugin = copilotInteraction.CopilotEventData.AISystemPlugin != null
                ? string.Join(", ", copilotInteraction.CopilotEventData.AISystemPlugin.Select(plugin => plugin.Id))
                : string.Empty;

            DateTime eventTime = DateTime.SpecifyKind(copilotInteraction.CreationTime, DateTimeKind.Utc);
            var tableEntity = new TableEntity(eventTime.ToString("yyyy-MM-dd"), copilotInteraction.Id.ToString())
            {
                { "User", copilotInteraction.UserId },
                { "AppHost", copilotInteraction.CopilotEventData.AppHost },
                { "CreationTime", eventTime },
                { "Contexts", contextsTypes },
                { "AISystemPlugin", aiPlugin },
                { "AgentId", copilotInteraction.CopilotEventData?.AgentId },
                { "AgentName", copilotInteraction.CopilotEventData?.AgentName }
            };


            try
            {
                await copilotInteractionDetailsTableClient.AddEntityAsync(tableEntity);
                _logger.LogInformation($"Added copilot interaction event at {eventTime}");
            }
            catch (Azure.RequestFailedException ex) when (ex.Status == 409) // Conflict indicates the entity already exists
            {
                // Merge the entity if it already exists
                await copilotInteractionDetailsTableClient.UpdateEntityAsync(tableEntity, ETag.All, TableUpdateMode.Merge);
            }
            catch (RequestFailedException ex)
            {
                // Handle the exception as needed
                _logger.LogError(ex, "Error adding copilot interaction event to table storage.");
                throw;
            }
        }

        /// <summary>
        /// Adds a copilot interaction daily aggregation for a user.
        /// NO longer used
        /// </summary>
        public async Task AddCopilotInteractionDailyAggregationForUserAsync(List<CopilotEventData> entity, string userId)
        {
            try
            {
                // group the event data via appHost
                var wordInteractions = entity.Where(e => e.AppHost == "Word").Count();
                var excelInteractions = entity.Where(e => e.AppHost == "Excel").Count();
                var powerPointInteractions = entity.Where(e => e.AppHost == "PowerPoint").Count();
                var onenoteInteractions = entity.Where(e => e.AppHost == "OneNote").Count();
                var outlookInteractions = entity.Where(e => e.AppHost == "Outlook").Count();
                var loopInteractions = entity.Where(e => e.AppHost == "Loop").Count();
                var teamsInteractions = entity
                    .Where(e => e.AppHost == "Teams" && e.Contexts != null && e.Contexts.Any(c => c.Type.StartsWith("Teams")))
                    .Count();
                var copilotChat = entity
                    .Where(e => e.AppHost == "Office"
                        || e.AppHost == "Edge"
                        || (e.AppHost == "Teams" && e.Contexts != null && e.Contexts.Any(c => string.IsNullOrEmpty(c.Type))))
                    .Count();
                var designerInteractions = entity.Where(e => e.AppHost == "Designer").Count();
                var sharePointInteractions = entity.Where(e => e.AppHost == "SharePoint").Count();
                var adminCenterInteractions = entity.Where(e => e.AppHost == "M365AdminCenter").Count();
                var webPluginInteractions = entity
                    .Where(e => e.AISystemPlugin != null && e.AISystemPlugin.Any(p => p.Id == "BingWebSearch"))
                    .Count();
                var copilotAction = entity.Where(e => e.AppHost == "OAIAutomationAgent").Count();
                var copilotStudioInteractions = entity.Where(e => e.AppHost == "Copilot Studio").Count();

                // All agent interactions
                var agentInteractions = entity.Where(e => e?.AgentId != null).Count();



                var totalInteractions = entity.Count();


                // Ensure the creationTime is specified as UTC
                DateTime eventTime = DateTime.SpecifyKind(DateTime.UtcNow, DateTimeKind.Utc);


                // Attempt to retrieve existing entity

                var retrieveOperation = await copilotInteractionDailyAggregationsTableClient.GetEntityIfExistsAsync<TableEntity>(eventTime.ToString("yyyy-MM-dd"), userId);
                if (retrieveOperation.HasValue)
                {
                    var existingEntity = retrieveOperation.Value;
                    // Update existing entity
                    existingEntity["TotalCount"] = (int)existingEntity["TotalCount"] + totalInteractions;
                    existingEntity["WordInteractions"] = (int)existingEntity["WordInteractions"] + wordInteractions;
                    existingEntity["ExcelInteractions"] = (int)existingEntity["ExcelInteractions"] + excelInteractions;
                    existingEntity["PowerPointInteractions"] = (int)existingEntity["PowerPointInteractions"] + powerPointInteractions;
                    existingEntity["OneNoteInteractions"] = (int)existingEntity["OneNoteInteractions"] + onenoteInteractions;
                    existingEntity["OutlookInteractions"] = (int)existingEntity["OutlookInteractions"] + outlookInteractions;
                    existingEntity["LoopInteractions"] = (int)existingEntity["LoopInteractions"] + loopInteractions;
                    existingEntity["TeamsInteractions"] = (int)existingEntity["TeamsInteractions"] + teamsInteractions;
                    existingEntity["CopilotChat"] = (int)existingEntity["CopilotChat"] + copilotChat;
                    existingEntity["DesignerInteractions"] = (int)existingEntity["DesignerInteractions"] + designerInteractions;
                    existingEntity["SharePointInteractions"] = (int)existingEntity["SharePointInteractions"] + sharePointInteractions;
                    existingEntity["AdminCenterInteractions"] = (int)existingEntity["AdminCenterInteractions"] + adminCenterInteractions;
                    existingEntity["WebPluginInteractions"] = (int)existingEntity["WebPluginInteractions"] + webPluginInteractions;
                    existingEntity["CopilotAction"] = (int)existingEntity["CopilotAction"] + copilotAction;
                    existingEntity["CopilotStudioInteractions"] = (int)existingEntity["CopilotStudioInteractions"] + copilotStudioInteractions;
                    existingEntity["AgentInteractions"] = (int)existingEntity["AgentInteractions"] + agentInteractions;


                    // Update the entity
                    await copilotInteractionDailyAggregationsTableClient.UpdateEntityAsync(existingEntity, ETag.All, TableUpdateMode.Merge);
                    _logger.LogInformation($"Updated copilot interaction daily aggregation for user {userId} at {eventTime}");
                }
                else
                {
                    // Entity doesn't exist, create a new one
                    var tableEntity = new TableEntity(eventTime.ToString("yyyy-MM-dd"), userId)
                    {
                        { "TotalCount", entity.Count },
                        { "WordInteractions", wordInteractions },
                        { "ExcelInteractions", excelInteractions },
                        { "PowerPointInteractions", powerPointInteractions },
                        { "OneNoteInteractions", onenoteInteractions },
                        { "OutlookInteractions", outlookInteractions },
                        { "LoopInteractions", loopInteractions },
                        { "TeamsInteractions", teamsInteractions },
                        { "CopilotChat", copilotChat },
                        { "DesignerInteractions", designerInteractions },
                        { "SharePointInteractions", sharePointInteractions },
                        { "AdminCenterInteractions", adminCenterInteractions },
                        { "WebPluginInteractions", webPluginInteractions },
                        { "CopilotAction", copilotAction },
                        { "CopilotStudioInteractions", copilotStudioInteractions },
                        { "AgentInteractions", agentInteractions }
                    };

                    // Add the new entity
                    await copilotInteractionDailyAggregationsTableClient.AddEntityAsync(tableEntity);
                }
            }
            catch (RequestFailedException ex)
            {
                // Handle the exception as needed
                _logger.LogError(ex, "Error retrieving copilot interaction daily aggregation event from table storage.");
                _logger.LogError(ex, "Message: {Message}", ex.Message);
                _logger.LogError(ex, "Status: {Status}", ex.StackTrace);

            }
        }

        public async Task AddListCreationRecordAsync(ListAuditObj listCreationEvent)
        {
            var tableClient = _serviceClient.GetTableClient(_listCreationTable);
            tableClient.CreateIfNotExists();

            // Ensure the creationTime is specified as UTC
            DateTime creationTime = DateTime.SpecifyKind(listCreationEvent.CreationTime, DateTimeKind.Utc);

            // Extract site URL from ObjectId
            var objectIdParts = listCreationEvent.ObjectId.Split('/');
            //var siteUrl = $"{objectIdParts[0]}//{objectIdParts[2]}/{objectIdParts[3]}/{objectIdParts[4]}";
            var siteUrl = $"{objectIdParts[4]}";

            var tableEntity = new TableEntity(siteUrl, listCreationEvent.AuditLogId)
            {
                { "ListUrl", listCreationEvent.ListUrl },
                { "ListName", listCreationEvent.ListName },
                { "ListBaseTemplateType", listCreationEvent.ListBaseTemplateType },
                { "ListBaseType", listCreationEvent.ListBaseType },
                { "CreationTime", creationTime }
            };

            try
            {
                await tableClient.AddEntityAsync(tableEntity);
                _logger.LogInformation($"Added list creation event for {listCreationEvent.ListName} at {listCreationEvent.ListUrl}");
            }
            catch (Azure.RequestFailedException ex) when (ex.Status == 409) // Conflict indicates the entity already exists
            {
                // Merge the entity if it already exists
                await tableClient.UpdateEntityAsync(tableEntity, ETag.All, TableUpdateMode.Merge);
            }
            catch (RequestFailedException ex)
            {
                // Handle the exception as needed
                _logger.LogError(ex, "Error adding list creation event to table storage.");
                throw;
            }
        }

        public async Task LogWebhookTriggerAsync(LogEvent webhookEvent)
        {
            var tableClient = _serviceClient.GetTableClient(_webhookEventsTable);
            tableClient.CreateIfNotExists();

            // Ensure the creationTime is specified as UTC
            DateTime eventTime = DateTime.SpecifyKind(webhookEvent.EventTime, DateTimeKind.Utc);

            var tableEntity = new TableEntity(eventTime.ToString("yyyy-MM-dd"), webhookEvent.EventId)
            {
                { "EventId", webhookEvent.EventId},
                { "EventName", webhookEvent.EventName },
                { "EventMessage", webhookEvent.EventMessage },
                { "EventDetails", webhookEvent.EventDetails },
                { "EventCategory", webhookEvent.EventCategory },
                { "EventTime", eventTime }
            };

            try
            {
                await tableClient.AddEntityAsync(tableEntity);
                _logger.LogInformation($"Added webhook trigger event at {eventTime}");
            }
            catch (Azure.RequestFailedException ex) when (ex.Status == 409) // Conflict indicates the entity already exists
            {
                // Merge the entity if it already exists
                await tableClient.UpdateEntityAsync(tableEntity, ETag.All, TableUpdateMode.Merge);
            }
            catch (RequestFailedException ex)
            {
                // Handle the exception as needed
                _logger.LogError(ex, "Error adding webhook trigger event to table storage.");
                throw;
            }
        }

        public async Task AddSingleCopilotInteractionDailyAggregationForUserAsync(CopilotEventData entity, string userId)
        {

            var copilotUsage = new CopilotTimeFrameUsage
            {
                UPN = userId,
                TotalInteractionCount = 1
            };

            // We need to set the app type based on the entity.AppHost and sometimes other logic
            // For example, if the entity.AppHost is "Teams" and the context type is "TeamsMeeting", we set it to AppType.Teams
            switch (entity.AppHost)
            {
                case "Word":
                    // Handle Word interactions
                    copilotUsage.App = AppType.Word;
                    break;
                case "Excel":
                    // Handle Excel interactions
                    copilotUsage.App = AppType.Excel;
                    break;
                case "PowerPoint":
                    // Handle PowerPoint interactions
                    copilotUsage.App = AppType.PowerPoint;
                    break;
                case "OneNote":
                    // Handle OneNote interactions
                    copilotUsage.App = AppType.OneNote;
                    break;
                case "Outlook":
                    // Handle Outlook interactions
                    copilotUsage.App = AppType.Outlook;
                    break;
                case "Loop":
                    // Handle Loop interactions
                    copilotUsage.App = AppType.Loop;
                    break;
                case "Whiteboard":
                    // Handle Whiteboard interactions
                    copilotUsage.App = AppType.Whiteboard;
                    break;
                case "Teams":
                    // Handle Teams interactions
                    if (entity.Contexts != null && entity.Contexts.Any(c => c.Type.StartsWith("Teams")))
                    {
                        copilotUsage.App = AppType.Teams;
                    }
                    else
                    {
                        copilotUsage.App = AppType.CopilotChat;
                    }
                    break;
                case "Office":
                    // Handle Office interactions
                    if (entity.AgentId != null)
                    {
                        copilotUsage.App = AppType.Agent;
                    }
                    else
                    {
                        copilotUsage.App = AppType.CopilotChat;
                    }
                    break;
                case "Edge":
                    // Handle Edge interactions
                    copilotUsage.App = AppType.CopilotChat;
                    break;
                case "Designer":
                    // Handle Designer interactions
                    copilotUsage.App = AppType.Designer;
                    break;
                case "SharePoint":
                    // Handle SharePoint interactions
                    copilotUsage.App = AppType.SharePoint;
                    break;
                case "M365AdminCenter":
                    // Handle M365AdminCenter interactions
                    copilotUsage.App = AppType.MAC;
                    break;
                case "OAIAutomationAgent":
                    // Handle OAIAutomationAgent interactions
                    copilotUsage.App = AppType.CopilotAction;
                    break;
                case "Copilot Studio":
                    // Handle Copilot Studio interactions
                    copilotUsage.App = AppType.CopilotStudio;
                    break;
                case "Forms":
                    // Handle Forms interactions
                    copilotUsage.App = AppType.Forms;
                    break;
                default:
                    // Handle other cases or log an error
                    // We have a new appHost to handle
                    _logger.LogWarning($"Unhandled AppHost: {entity.AppHost}");

                    await AddUnhandledAppHostAsync(entity.AppHost);

                    break;
            }


            var copilotUsageEntity = copilotUsage.ToDailyAggregationTableEntity(entity.EventDateString);

            // Add the entity to the table
            await CreateOrUpdateCopilotUsageEntityAsync(copilotUsageEntity, copilotUsage.TotalInteractionCount);

        }

        private async Task AddUnhandledAppHostAsync(string appHost)
        {
            try
            {
                // Check if the appHost is already logged
                var retrieveOperation = await unhandledAppHostsTableClient.GetEntityIfExistsAsync<TableEntity>(appHost, appHost);
                if (retrieveOperation.HasValue)
                {
                    // If it exists, we don't need to add it again
                    _logger.LogInformation($"AppHost '{appHost}' is already logged as unhandled.");

                    // increment the count of unhandled app hosts
                    var existingEntity = retrieveOperation.Value;
                    existingEntity["Count"] = (int)(existingEntity.ContainsKey("Count") ? existingEntity["Count"] : 0) + 1;

                    await unhandledAppHostsTableClient.UpdateEntityAsync(existingEntity, existingEntity.ETag, TableUpdateMode.Merge);
                    _logger.LogInformation($"Incremented count for unhandled AppHost '{appHost}' to {existingEntity["Count"]}");

                    return;
                }
                else
                {
                    var unhandledAppHostEntity = new TableEntity(appHost, appHost)
                    {
                        { "Count", 1 }
                    };

                    await unhandledAppHostsTableClient.AddEntityAsync(unhandledAppHostEntity);

                    _logger.LogInformation($"Added unhandled AppHost '{appHost}' to the table.");
                }
            }
            catch (RequestFailedException ex)
            {
                _logger.LogError(ex, "Error adding unhandled AppHost to table storage.");
                _logger.LogError("Message: {Message}", ex.Message);
                _logger.LogError("Status: {Status}", ex.Status);
                _logger.LogError("StackTrace: {StackTrace}", ex.StackTrace);
            }
        }

        public async Task AddSpecificCopilotInteractionDailyAggregationForUserAsync(AppType appType, string userId, int count)
        {
            var copilotUsage = new CopilotTimeFrameUsage
            {
                UPN = userId,
                TotalInteractionCount = count,
                App = appType
            };

            var copilotUsageEntity = copilotUsage.ToDailyAggregationTableEntity(DateTime.UtcNow.ToString("yyyy-MM-dd"));
            // Add the entity to the table
            await CreateOrUpdateCopilotUsageEntityAsync(copilotUsageEntity, copilotUsage.TotalInteractionCount);
        }

        private async Task CreateOrUpdateCopilotUsageEntityAsync(TableEntity copilotUsageEntity, int interactionCount)
        {

            try
            {
                var retrieveOperation = await copilotInteractionDailyAggregationsTableClient
                    .GetEntityIfExistsAsync<CopilotTimeFrameUsage>(
                        copilotUsageEntity.PartitionKey,
                        copilotUsageEntity.RowKey);

                if (retrieveOperation.HasValue)
                {
                    var existingEntity = retrieveOperation.Value;

                    existingEntity.TotalInteractionCount += interactionCount;

                    await copilotInteractionDailyAggregationsTableClient
                        .UpdateEntityAsync(existingEntity, existingEntity.ETag, TableUpdateMode.Merge);
                }
                else
                {
                    await copilotInteractionDailyAggregationsTableClient
                        .AddEntityAsync(copilotUsageEntity);
                }
            }
            catch (RequestFailedException ex) when (ex.Status == 409)
            {
                // Entity was added by someone else just after our check
                // Retry the update path
                var newlyUpdatedEntity = await copilotInteractionDailyAggregationsTableClient
                    .GetEntityAsync<CopilotTimeFrameUsage>(copilotUsageEntity.PartitionKey, copilotUsageEntity.RowKey);

                newlyUpdatedEntity.Value.TotalInteractionCount += interactionCount;

                await copilotInteractionDailyAggregationsTableClient
                    .UpdateEntityAsync(newlyUpdatedEntity.Value, newlyUpdatedEntity.Value.ETag, TableUpdateMode.Merge);
            }
            catch (RequestFailedException ex)
            {
                _logger.LogError(ex, "Error retrieving or updating copilot interaction aggregation.");
                _logger.LogError("Message: {Message}", ex.Message);
                _logger.LogError("Status: {Status}", ex.Status);
                _logger.LogError("StackTrace: {StackTrace}", ex.StackTrace);
            }
        }

        public async Task AddAgentInteractionsDailyAggregationForUserAsync(string agentId, string userId, int count, string agentName)
        {
            var agentUsage = new AgentInteraction
            {
                UPN = userId,
                TotalInteractionCount = count,
                AgentId = agentId,
                AgentName = agentName
            };

            var agentUsageEntity = agentUsage.ToDailyTableEntity(DateTime.UtcNow.ToString("yyyy-MM-dd"));

            try
            {
                var retrieveOperation = await copilotAgentInteractionTableClient
                    .GetEntityIfExistsAsync<AgentInteraction>(
                        agentUsageEntity.PartitionKey,
                        agentUsageEntity.RowKey);

                if (retrieveOperation.HasValue)
                {
                    var existingEntity = retrieveOperation.Value;

                    existingEntity.TotalInteractionCount += count;

                    await copilotAgentInteractionTableClient
                        .UpdateEntityAsync(existingEntity, existingEntity.ETag, TableUpdateMode.Merge);
                }
                else
                {
                    await copilotAgentInteractionTableClient
                        .AddEntityAsync(agentUsageEntity);
                }
            }
            catch (RequestFailedException ex) when (ex.Status == 409)
            {
                // Entity was added by someone else just after our check
                // Retry the update path
                var newlyUpdatedEntity = await copilotAgentInteractionTableClient
                    .GetEntityAsync<CopilotTimeFrameUsage>(agentUsageEntity.PartitionKey, agentUsageEntity.RowKey);

                newlyUpdatedEntity.Value.TotalInteractionCount += count;

                await copilotAgentInteractionTableClient
                    .UpdateEntityAsync(newlyUpdatedEntity.Value, newlyUpdatedEntity.Value.ETag, TableUpdateMode.Merge);
            }
            catch (RequestFailedException ex)
            {
                _logger.LogError(ex, "Error retrieving or updating agent interaction aggregation.");
                _logger.LogError("Message: {Message}", ex.Message);
                _logger.LogError("Status: {Status}", ex.Status);
                _logger.LogError("StackTrace: {StackTrace}", ex.StackTrace);
            }
        }

        #region Key Rotation Methods

        public async Task RotateKeyRecentDailyTableValuesLegacy(TableClient tableClient, DeterministicEncryptionService existingKeyService, DeterministicEncryptionService newKeyService)
        {
            // This method is a placeholder for the decryption logic
            // It should decrypt all UPNs in the table storage for the last 7 days
            // The actual implementation will depend on the encryption method used
            _logger.LogInformation("Decrypting recent daily table values...");

            // Loop for each day from today - 7 days to today (inclusive)
            // Start with today and work backwards 7 days
            for (int i = 0; i <= 7; i++)
            {
                var date = DateTime.UtcNow.Date.AddDays(-i);
                var nextDate = date.AddDays(1);

                string start = date.ToString("yyyy-MM-dd");
                string end = nextDate.ToString("yyyy-MM-dd");

                string filter = $"PartitionKey ge '{start}' and PartitionKey lt '{end}'";

                // log the filter
                _logger.LogInformation($"Querying table with filter: {filter}");


                await foreach (var entity in tableClient.QueryAsync<TableEntity>(filter: filter))
                {

                    // LOG THE PARTITION KEY
                    _logger.LogInformation($"Processing entity with PartitionKey: {entity.PartitionKey}");

                    // Step 1: Extract encrypted part from PartitionKey
                    // need to strip out the data and hyphen from the PartitionKey. AS the format is yyyy-MM-dd-<encrypted-blob> we can just count so remove first 11 chars
                    string encyrptedUPN = entity.PartitionKey.Substring(11); // Remove the date prefix
                                                                             // Decrypt the UPN using the existing key service
                    string decryptedUPN = existingKeyService.Decrypt(encyrptedUPN);


                    // log the decrypted UPN
                    _logger.LogInformation($"Decrypted UPN: {decryptedUPN}");

                    // Encrypt the UPN using the new key service
                    string newEncryptedUPN = newKeyService.Encrypt(decryptedUPN);

                    // Create a new partition key with the new encrypted UPN
                    string newPartitionKey = $"{start}-{newEncryptedUPN}";

                    // log the new partition key
                    _logger.LogInformation($"New PartitionKey: {newPartitionKey}");

                    // Step 2: Copy all properties
                    var newEntity = new TableEntity(newPartitionKey, entity.RowKey);
                    foreach (var kvp in entity)
                    {
                        if (kvp.Key != "PartitionKey" && kvp.Key != "RowKey")
                        {
                            newEntity[kvp.Key] = kvp.Value;
                        }
                    }

                    try
                    {

                        // Step 3: Insert new entity and delete old
                        await tableClient.AddEntityAsync(newEntity);
                        await tableClient.DeleteEntityAsync(entity.PartitionKey, entity.RowKey);
                    }
                    catch (RequestFailedException ex)
                    {
                        _logger.LogError(ex, "Error querying or updating table storage for date {Date}.", date.ToString("yyyy-MM-dd"));
                        _logger.LogError("Message: {Message}", ex.Message);
                        _logger.LogError("Status: {Status}", ex.Status);
                        _logger.LogError("StackTrace: {StackTrace}", ex.StackTrace);
                    }

                    _logger.LogInformation($"Rotated key for entity with PartitionKey: {entity.PartitionKey}, RowKey: {entity.RowKey}");
                }
            }

        }

        public async Task RotateKeyTimeFrameTableValuesLegacy(TableClient tableClient, DeterministicEncryptionService existingKeyService, DeterministicEncryptionService newKeyService)
        {

            _logger.LogInformation("Decrypting timeframe table values...");
            var totalProcessed = 0;
            var totalErrors = 0;

            string filter = "";

            if (tableClient.Name == _userWeeklyTableName || tableClient.Name == _agentWeeklyByUserTableName)
            {
                // Get the weekly start date
                string startDate = await GetActiveTimeFrame("weekly");

                // convert to date
                if (!DateTime.TryParse(startDate, out DateTime date))
                {
                    _logger.LogError($"Failed to parse start date: {startDate}");
                    throw new ArgumentException($"Invalid start date format: {startDate}");
                }
                string endDate = date.AddDays(7).ToString("yyyy-MM-dd");

                filter = $"PartitionKey ge '{startDate}' and PartitionKey lt '{endDate}'";
            }
            else if (tableClient.Name == _userMonthlyTableName || tableClient.Name == _agentMonthlyByUserTableName)
            {
                // get the monthly start date
                string startDate = await GetActiveTimeFrame("monthly");

                // convert to date
                if (!DateTime.TryParse(startDate, out DateTime date))
                {
                    _logger.LogError($"Failed to parse start date: {startDate}");
                    throw new ArgumentException($"Invalid start date format: {startDate}");
                }
                string endDate = date.AddMonths(1).ToString("yyyy-MM-dd");

                filter = $"PartitionKey ge '{startDate}' and PartitionKey lt '{endDate}'";
            }
            else
            {
                // all time
            }

            // log the filter
            _logger.LogInformation($"Querying table with filter: {filter}");


            await foreach (var entity in tableClient.QueryAsync<TableEntity>(filter: filter))
            {

                // LOG THE PARTITION KEY
                _logger.LogInformation($"Processing entity with PartitionKey: {entity.PartitionKey} and RowKey: {entity.RowKey}");

                // partion key stays the same

                // we need to decypt the row key
                string encyrptedUPN = entity.RowKey; // Assuming the RowKey is the encrypted UPN
                string decryptedUPN = existingKeyService.Decrypt(encyrptedUPN);


                // log the decrypted UPN
                _logger.LogInformation($"Decrypted UPN: {decryptedUPN}");

                // Encrypt the UPN using the new key service
                string newEncryptedUPN = newKeyService.Encrypt(decryptedUPN);

                // Create a new partition key with the new encrypted UPN
                string newRowKey = newEncryptedUPN;

                // log the new partition key
                _logger.LogInformation($"New RowKey: {newRowKey}");

                // Step 2: Copy all properties
                var newEntity = new TableEntity(entity.PartitionKey, newRowKey);
                foreach (var kvp in entity)
                {
                    if (kvp.Key != "PartitionKey" && kvp.Key != "RowKey")
                    {
                        newEntity[kvp.Key] = kvp.Value;
                    }
                }

                try
                {

                    // Step 3: Insert new entity and delete old
                    await tableClient.AddEntityAsync(newEntity);
                    await tableClient.DeleteEntityAsync(entity.PartitionKey, entity.RowKey);
                }
                catch (RequestFailedException ex)
                {
                    _logger.LogError($"Error updating table storage using {filter}", ex);
                    _logger.LogError("Message: {Message}", ex.Message);
                    _logger.LogError("Status: {Status}", ex.Status);
                    _logger.LogError("StackTrace: {StackTrace}", ex.StackTrace);
                }

                _logger.LogInformation($"Rotated key for entity with PartitionKey: {entity.PartitionKey}, RowKey: {entity.RowKey}");
            }


        }

        public async Task RotateKeyRecentDailyTableValues(TableClient tableClient, DeterministicEncryptionService existingKeyService, DeterministicEncryptionService newKeyService)
        {
            _logger.LogInformation("Starting key rotation for recent daily table values...");
            var totalProcessed = 0;
            var totalErrors = 0;

            // Loop for each day from today - 7 days to today (inclusive)
            for (int i = 0; i <= 7; i++)
            {
                var date = DateTime.UtcNow.Date.AddDays(-i);
                var nextDate = date.AddDays(1);

                string start = date.ToString("yyyy-MM-dd");
                string end = nextDate.ToString("yyyy-MM-dd");
                string filter = $"PartitionKey ge '{start}' and PartitionKey lt '{end}'";

                _logger.LogInformation($"Processing date {start} with filter: {filter}");

                var dayProcessed = 0;
                var dayErrors = 0;

                try
                {
                    // Process entities in pages to avoid memory issues
                    await foreach (var page in tableClient.QueryAsync<TableEntity>(filter: filter).AsPages(pageSizeHint: 1000))
                    {
                        var entitiesToProcess = new List<PartitionKeyRotationInfo>();

                        // Prepare all entities for this page
                        foreach (var entity in page.Values)
                        {
                            try
                            {
                                var processingInfo = PrepareEntityForPartitionKeyRotation(entity, existingKeyService, newKeyService, start);
                                entitiesToProcess.Add(processingInfo);
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError(ex, "Error preparing entity {PartitionKey}/{RowKey} for partition key rotation",
                                    entity.PartitionKey, entity.RowKey);
                                dayErrors++;
                            }
                        }

                        // Process entities in batches
                        var batchResults = await ProcessPartitionKeyRotationInBatches(tableClient, entitiesToProcess);
                        dayProcessed += batchResults.Processed;
                        dayErrors += batchResults.Errors;

                        _logger.LogInformation($"Processed page: {batchResults.Processed} successful, {batchResults.Errors} errors");

                        // Small delay to avoid throttling
                        if (page.Values.Count == 1000) // Full page, likely more coming
                        {
                            await Task.Delay(100);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error processing date {Date}", start);
                    dayErrors++;
                }

                totalProcessed += dayProcessed;
                totalErrors += dayErrors;

                _logger.LogInformation($"Completed date {start}: {dayProcessed} processed, {dayErrors} errors");
            }

            _logger.LogInformation($"Partition key rotation completed. Total: {totalProcessed} processed, {totalErrors} errors");
        }

        public async Task RotateKeyTimeFrameTableValues(TableClient tableClient, DeterministicEncryptionService existingKeyService, DeterministicEncryptionService newKeyService)
        {
            _logger.LogInformation("Starting key rotation for timeframe table values...");
            var totalProcessed = 0;
            var totalErrors = 0;

            string filter = "";

            // Determine filter based on table name
            if (tableClient.Name == _userWeeklyTableName || tableClient.Name == _agentWeeklyByUserTableName)
            {
                string startDate = await GetActiveTimeFrame("weekly");
                if (!DateTime.TryParse(startDate, out DateTime date))
                {
                    _logger.LogError($"Failed to parse start date: {startDate}");
                    throw new ArgumentException($"Invalid start date format: {startDate}");
                }
                string endDate = date.AddDays(7).ToString("yyyy-MM-dd");
                filter = $"PartitionKey ge '{startDate}' and PartitionKey lt '{endDate}'";
            }
            else if (tableClient.Name == _userMonthlyTableName || tableClient.Name == _agentMonthlyByUserTableName)
            {
                string startDate = await GetActiveTimeFrame("monthly");
                if (!DateTime.TryParse(startDate, out DateTime date))
                {
                    _logger.LogError($"Failed to parse start date: {startDate}");
                    throw new ArgumentException($"Invalid start date format: {startDate}");
                }
                string endDate = date.AddMonths(1).ToString("yyyy-MM-dd");
                filter = $"PartitionKey ge '{startDate}' and PartitionKey lt '{endDate}'";
            }
            // else: all time (no filter)

            _logger.LogInformation($"Querying table with filter: {filter}");

            try
            {
                // Process entities in pages
                await foreach (var page in tableClient.QueryAsync<TableEntity>(filter: filter).AsPages(pageSizeHint: 1000))
                {
                    var entitiesToProcess = new List<RowKeyRotationInfo>();

                    // Prepare all entities for this page
                    foreach (var entity in page.Values)
                    {
                        try
                        {
                            var processingInfo = PrepareEntityForRowKeyRotation(entity, existingKeyService, newKeyService);
                            entitiesToProcess.Add(processingInfo);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Error preparing entity {PartitionKey}/{RowKey} for row key rotation",
                                entity.PartitionKey, entity.RowKey);
                            totalErrors++;
                        }
                    }

                    // Process entities in batches
                    var batchResults = await ProcessRowKeyRotationInBatches(tableClient, entitiesToProcess);
                    totalProcessed += batchResults.Processed;
                    totalErrors += batchResults.Errors;

                    _logger.LogInformation($"Processed page: {totalProcessed} total processed, {totalErrors} total errors");

                    // Small delay to avoid throttling
                    if (page.Values.Count == 1000) // Full page, likely more coming
                    {
                        await Task.Delay(100);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing timeframe table values");
                totalErrors++;
            }

            _logger.LogInformation($"Row key rotation completed. Total: {totalProcessed} processed, {totalErrors} errors");
        }

        // Helper methods for PartitionKey rotation
        private PartitionKeyRotationInfo PrepareEntityForPartitionKeyRotation(TableEntity entity,
            DeterministicEncryptionService existingKeyService,
            DeterministicEncryptionService newKeyService,
            string datePrefix)
        {
            // Extract encrypted part from PartitionKey (remove first 11 chars: yyyy-MM-dd-)
            string encryptedUPN = entity.PartitionKey.Substring(11);
            string decryptedUPN = existingKeyService.Decrypt(encryptedUPN);
            string newEncryptedUPN = newKeyService.Encrypt(decryptedUPN);
            string newPartitionKey = $"{datePrefix}-{newEncryptedUPN}";

            // Create new entity with all properties
            var newEntity = new TableEntity(newPartitionKey, entity.RowKey);
            foreach (var kvp in entity)
            {
                if (kvp.Key != "PartitionKey" && kvp.Key != "RowKey")
                {
                    newEntity[kvp.Key] = kvp.Value;
                }
            }

            return new PartitionKeyRotationInfo
            {
                OriginalEntity = entity,
                NewEntity = newEntity,
                DecryptedUPN = decryptedUPN
            };
        }

        private async Task<(int Processed, int Errors)> ProcessPartitionKeyRotationInBatches(TableClient tableClient, List<PartitionKeyRotationInfo> entities)
        {
            var processed = 0;
            var errors = 0;
            const int batchSize = 50; // Conservative batch size

            // Process in chunks
            var chunks = entities.Chunk(batchSize);

            foreach (var chunk in chunks)
            {
                var chunkResults = await ProcessPartitionKeyChunk(tableClient, chunk.ToList());
                processed += chunkResults.Processed;
                errors += chunkResults.Errors;

                // Small delay between batches
                await Task.Delay(50);
            }

            return (processed, errors);
        }

        private async Task<(int Processed, int Errors)> ProcessPartitionKeyChunk(TableClient tableClient, List<PartitionKeyRotationInfo> entities)
        {
            var processed = 0;
            var errors = 0;

            // For partition key changes, we need to add new entities first, then delete old ones individually
            // since they're in different partitions and can't be batched together

            try
            {
                // Group new entities by their new partition key for batch adds
                var addGroups = entities.GroupBy(e => e.NewEntity.PartitionKey);

                foreach (var group in addGroups)
                {
                    var addOperations = group.Select(e => new TableTransactionAction(TableTransactionActionType.Add, e.NewEntity))
                                            .Take(100) // Max 100 per batch
                                            .ToList();

                    if (addOperations.Count > 0)
                    {
                        await tableClient.SubmitTransactionAsync(addOperations);
                    }
                }

                // Delete old entities individually (different partition keys)
                foreach (var entity in entities)
                {
                    try
                    {
                        await tableClient.DeleteEntityAsync(entity.OriginalEntity.PartitionKey, entity.OriginalEntity.RowKey);
                        processed++;

                        _logger.LogDebug($"Rotated partition key for UPN: {entity.DecryptedUPN}");
                    }
                    catch (RequestFailedException ex) when (ex.Status == 404)
                    {
                        // Entity already deleted, count as success
                        processed++;
                        _logger.LogWarning($"Entity {entity.OriginalEntity.PartitionKey}/{entity.OriginalEntity.RowKey} already deleted");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error deleting entity {PartitionKey}/{RowKey}",
                            entity.OriginalEntity.PartitionKey, entity.OriginalEntity.RowKey);
                        errors++;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Batch operation failed, falling back to individual operations");

                // Fallback: process individually
                foreach (var entity in entities)
                {
                    try
                    {
                        await tableClient.AddEntityAsync(entity.NewEntity);
                        await tableClient.DeleteEntityAsync(entity.OriginalEntity.PartitionKey, entity.OriginalEntity.RowKey);
                        processed++;

                        _logger.LogDebug($"Individually rotated partition key for UPN: {entity.DecryptedUPN}");
                    }
                    catch (Exception individualEx)
                    {
                        _logger.LogError(individualEx, "Error processing entity {PartitionKey}/{RowKey}",
                            entity.OriginalEntity.PartitionKey, entity.OriginalEntity.RowKey);
                        errors++;
                    }
                }
            }

            return (processed, errors);
        }

        // Helper methods for RowKey rotation
        private RowKeyRotationInfo PrepareEntityForRowKeyRotation(TableEntity entity,
            DeterministicEncryptionService existingKeyService,
            DeterministicEncryptionService newKeyService)
        {
            // Decrypt the RowKey (assuming it's the encrypted UPN)
            string encryptedUPN = entity.RowKey;
            string decryptedUPN = existingKeyService.Decrypt(encryptedUPN);
            string newEncryptedUPN = newKeyService.Encrypt(decryptedUPN);

            // Create new entity with same PartitionKey but new RowKey
            var newEntity = new TableEntity(entity.PartitionKey, newEncryptedUPN);
            foreach (var kvp in entity)
            {
                if (kvp.Key != "PartitionKey" && kvp.Key != "RowKey")
                {
                    newEntity[kvp.Key] = kvp.Value;
                }
            }

            return new RowKeyRotationInfo
            {
                OriginalEntity = entity,
                NewEntity = newEntity,
                DecryptedUPN = decryptedUPN
            };
        }

        private async Task<(int Processed, int Errors)> ProcessRowKeyRotationInBatches(TableClient tableClient, List<RowKeyRotationInfo> entities)
        {
            var processed = 0;
            var errors = 0;

            // Group by partition key for batch operations (row key changes can be batched within same partition)
            var partitionGroups = entities.GroupBy(e => e.OriginalEntity.PartitionKey);

            foreach (var partitionGroup in partitionGroups)
            {
                // Process each partition in chunks of 100 (Azure batch limit)
                var chunks = partitionGroup.Chunk(100);

                foreach (var chunk in chunks)
                {
                    var chunkResults = await ProcessRowKeyChunk(tableClient, chunk.ToList());
                    processed += chunkResults.Processed;
                    errors += chunkResults.Errors;

                    // Small delay between batches
                    await Task.Delay(50);
                }
            }

            return (processed, errors);
        }

        private async Task<(int Processed, int Errors)> ProcessRowKeyChunk(TableClient tableClient, List<RowKeyRotationInfo> entities)
        {
            var processed = 0;
            var errors = 0;

            try
            {
                // All entities in chunk have same partition key, so we can batch both adds and deletes
                var addOperations = entities.Select(e =>
                    new TableTransactionAction(TableTransactionActionType.Add, e.NewEntity)).ToList();

                var deleteOperations = entities.Select(e =>
                    new TableTransactionAction(TableTransactionActionType.Delete, e.OriginalEntity)).ToList();

                // Add new entities first
                if (addOperations.Count > 0)
                {
                    await tableClient.SubmitTransactionAsync(addOperations);
                }

                // Then delete old entities
                if (deleteOperations.Count > 0)
                {
                    await tableClient.SubmitTransactionAsync(deleteOperations);
                }

                processed = entities.Count;

                foreach (var entity in entities)
                {
                    _logger.LogDebug($"Rotated row key for UPN: {entity.DecryptedUPN} in partition {entity.OriginalEntity.PartitionKey}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Batch row key rotation failed, falling back to individual operations");

                // Fallback: process individually
                foreach (var entity in entities)
                {
                    try
                    {
                        await tableClient.AddEntityAsync(entity.NewEntity);
                        await tableClient.DeleteEntityAsync(entity.OriginalEntity.PartitionKey, entity.OriginalEntity.RowKey);
                        processed++;

                        _logger.LogDebug($"Individually rotated row key for UPN: {entity.DecryptedUPN}");
                    }
                    catch (RequestFailedException ex2) when (ex2.Status == 409)
                    {
                        _logger.LogWarning($"Entity with new RowKey already exists: {entity.OriginalEntity.PartitionKey}/{entity.NewEntity.RowKey}");
                        errors++;
                    }
                    catch (Exception individualEx)
                    {
                        _logger.LogError(individualEx, "Error rotating row key for entity {PartitionKey}/{RowKey}",
                            entity.OriginalEntity.PartitionKey, entity.OriginalEntity.RowKey);
                        errors++;
                    }
                }
            }

            return (processed, errors);
        }

        // Helper classes
        private class PartitionKeyRotationInfo
        {
            public required TableEntity OriginalEntity { get; set; }
            public required TableEntity NewEntity { get; set; }
            public required string DecryptedUPN { get; set; }
        }

        private class RowKeyRotationInfo
        {
            public required TableEntity OriginalEntity { get; set; }
            public required TableEntity NewEntity { get; set; }
            public required string DecryptedUPN { get; set; }
        }



        public async Task<bool> GetWebhookStateAsync()
        {
            try
            {
                var response = await auditWebhookStateTableClient.GetEntityIfExistsAsync<TableEntity>("Webhook", "Pause");
                if (response.HasValue && response.Value.ContainsKey("IsPaused"))
                {
                    return response.Value.GetBoolean("IsPaused") ?? false; // If property missing, treat as not paused
                }
                // If the entity does not exist, treat as not paused
                _logger.LogWarning("Webhook state entity not found. Returning not paused (false).");
                return false;
            }
            catch (RequestFailedException ex)
            {
                _logger.LogError(ex, "Failed to get webhook state. Returning not paused (false).");
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error while getting webhook state. Returning not paused (false).");
                return false;
            }
        }

        public async Task SetWebhookStateAsync(bool isPaused)
        {
            try
            {

                var entity = new TableEntity("Webhook", "Pause")
                {
                    { "IsPaused", isPaused }
                };
                await auditWebhookStateTableClient.UpsertEntityAsync(entity);
                _logger.LogInformation($"Webhook state set to paused={isPaused}.");
            }
            catch (RequestFailedException ex)
            {
                _logger.LogError(ex, $"Failed to set webhook state to paused={isPaused}.");
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error while setting webhook state.");
                throw;
            }
        }

        public TableClient GetCopilotAgentInteractionTableClient()
        {
            return copilotAgentInteractionTableClient;
        }

        public TableClient GetCopilotInteractionDailyAggregationsTableClient()
        {
            return copilotInteractionDailyAggregationsTableClient;
        }

        public TableClient GetWeeklyCopilotInteractionTableClient()
        {
            return _userWeeklyTableClient;
        }

        public TableClient GetMonthlyCopilotInteractionTableClient()
        {
            return _userMonthlyTableClient;
        }
        public TableClient GetAllTimeCopilotInteractionTableClient()
        {
            return _userAllTimeTableClient;
        }
        public TableClient GetWeeklyAgentInteractionTableClient()
        {
            return _agentWeeklyTableClient;
        }
        public TableClient GetMonthlyAgentInteractionTableClient()
        {
            return _agentMonthlyTableClient;
        }
        public TableClient GetAllTimeAgentInteractionTableClient()
        {
            return _agentAllTimeTableClient;
        }


        #endregion

        #region From ScraperFunction

        public async Task<int> ProcessUserDailySnapshots(List<M365CopilotUsage> userSnapshots, DeterministicEncryptionService encryptionService)
        {
            int DAUadded = 0;

            // Tuple to store the user's last activity date and username
            var lastActivityDates = new List<(string, string, string, string)>();

            string reportRefreshDateString = string.Empty;



            foreach (var userSnap in userSnapshots)
            {

                try
                {
                    // Log the user snapshot being processed
                    _logger.LogInformation($"Processing user snapshot for UPN: {userSnap.UserPrincipalName}, ReportRefreshDate: {userSnap.ReportRefreshDate}, LastActivityDate: {userSnap.LastActivityDate}");


                    reportRefreshDateString = userSnap.ReportRefreshDate;
                    // if last activity date is not the same as the report refresh date, we need to validate how long usage has not occured
                    // there is no daily activity if the last activity date is not the same as the report refresh date

                    if (userSnap.LastActivityDate != userSnap.ReportRefreshDate)
                    {
                        var reportRefreshDate = DateTime.ParseExact(userSnap.ReportRefreshDate, "yyyy-MM-dd", null);

                        if (string.IsNullOrEmpty(userSnap.LastActivityDate))
                        {
                            // we need to record in another table
                            // Can we set the last activity as epoch when null?
                            var epochTime = new DateTime(1970, 1, 1);
                            lastActivityDates.Add((epochTime.ToString("yyyy-MM-dd"), userSnap.UserPrincipalName, userSnap.ReportRefreshDate, userSnap.DisplayName));
                        }
                        else
                        {
                            // Convert to date time
                            var lastActivityDate = DateTime.ParseExact(userSnap.LastActivityDate, "yyyy-MM-dd", null);


                            // Check if last activity is before days ti check
                            if (lastActivityDate.AddDays(_daysToCheck) < reportRefreshDate)
                            {
                                // we need to record in another table
                                lastActivityDates.Add((userSnap.LastActivityDate, userSnap.UserPrincipalName, userSnap.ReportRefreshDate, userSnap.DisplayName));
                            }
                        }

                    }

                    // Encrypt the UPN - lookup is already encrypted
                    userSnap.UserPrincipalName = encryptionService.Encrypt(userSnap.UserPrincipalName);

                    // if getting audit data we should get the aggregation data for the user
                    // Get the aggregation entity
                    var aggregationEntity = await GetDailyAuditDataForUser(userSnap.UserPrincipalName, userSnap.ReportRefreshDate);

                    var userEntity = ConvertToUserActivity(userSnap, aggregationEntity);

                    // Get User Activity
                    // Todo, store the more precise data in the table
                    var userActivityDictionary = ConvertToUsageDictionary(userEntity);

                    // Log the user activityDictionary
                    _logger.LogInformation($"User Activity Dictionary for {userEntity.UPN}: {string.Join(", ", userActivityDictionary.Select(kvp => $"{kvp.Key}: {kvp.Value}"))}");


                    // Todo, do we really need to store this data in the table?
                    // For now we won't

                    // try
                    // {
                    //     // Try to add the entity if it doesn't exist
                    //     await _userDAUTableClient.AddEntityAsync(userEntity.ToTableEntity());
                    //     DAUadded++;

                    // }
                    // catch (Azure.RequestFailedException ex) when (ex.Status == 409) // Conflict indicates the entity already exists
                    // {
                    //     // Merge the entity if it already exists
                    //     await _userDAUTableClient.UpdateEntityAsync(userEntity.ToTableEntity(), ETag.All, TableUpdateMode.Merge);
                    // }

                    // Agent Data
                    var dailyAgentData = await GetDailyAgentDataForUser(userEntity.UPN, userEntity.ReportDate.ToString("yyyy-MM-dd"));


                    // We need to update the  weekly, monthly and alltime tables
                    await UpdateUserSnapshots(userActivityDictionary, userEntity.UPN, "alltime", userEntity.ReportDate.ToString("yyyy-MM-dd"));
                    await UpdateAgentSnapshots(dailyAgentData, "alltime", userEntity.UPN, userEntity.ReportDate.ToString("yyyy-MM-dd"));

                    // Update Monthly
                    var firstOfMonthForSnapshot = GetMonthStartDate(userEntity.ReportDate);

                    await UpdateUserSnapshots(userActivityDictionary, userEntity.UPN, "monthly", firstOfMonthForSnapshot);
                    await UpdateAgentSnapshots(dailyAgentData, "monthly", userEntity.UPN, firstOfMonthForSnapshot);

                    // Update Weekly - Get our timeFrame
                    var firstMondayOfWeeklySnapshot = GetWeekStartDate(userEntity.ReportDate);

                    await UpdateUserSnapshots(userActivityDictionary, userEntity.UPN, "weekly", firstMondayOfWeeklySnapshot);
                    await UpdateAgentSnapshots(dailyAgentData, "weekly", userEntity.UPN, firstMondayOfWeeklySnapshot);

                }
                catch (Exception logEx)
                {
                    _logger.LogError(logEx, "Error logging user snapshot for UPN: {UPN}", userSnap.UserPrincipalName);
                    // log the trace
                    _logger.LogError("Message: {Message}", logEx.Message);
                    _logger.LogError(logEx, "Error processing usage data. {StackTrace}", logEx.StackTrace);

                }
            }

            // Update the timeFrame table
            await UpdateReportRefreshDate(reportRefreshDateString, "daily");
            await UpdateReportRefreshDate(reportRefreshDateString, "weekly");
            await UpdateReportRefreshDate(reportRefreshDateString, "monthly");

            // For notifications - now handled elsewhere
            //return DAUadded;

            // Do we have any users to record in the last activity table?
            if (lastActivityDates.Count > 0)
            {

                var today = DateTime.UtcNow.ToString("yyyy-MM-dd");

                foreach (var (lastActivityDate, userPrincipalName, reportRefreshDateItem, displayName) in lastActivityDates)
                {
                    // Encrypt the UPN
                    var encryptedUPN = encryptionService.Encrypt(userPrincipalName);

                    // Check if record exists in table for user
                    var tableEntity = new TableEntity(encryptedUPN, lastActivityDate)
                    {
                        { "LastActivityDate", lastActivityDate },
                        { "ReportRefreshDate", reportRefreshDateItem },
                        { "DaysSinceLastActivity", (DateTime.ParseExact(reportRefreshDateItem, "yyyy-MM-dd", null) - DateTime.ParseExact(lastActivityDate, "yyyy-MM-dd", null)).TotalDays },
                        // { "LastNotificationDate", today },
                        // { "DaysSinceLastNotification", (double)999 },
                        // { "NotificationCount", 0 },
                    };

                    try
                    {
                        // Try to add the entity if it doesn't exist
                        await _userLastUsageTableClient.AddEntityAsync(tableEntity);
                    }
                    catch (Azure.RequestFailedException ex) when (ex.Status == 409) // Conflict indicates the entity already exists
                    {
                        await _userLastUsageTableClient.UpdateEntityAsync(tableEntity, ETag.All, TableUpdateMode.Merge);
                    }
                }
            }

            return DAUadded;

        }

        public async Task<int> ProcessAgentUsageAggregationsAsync(string dateToProcess)
        {
            int recordsAdded = 0;

            // Query the daily agent interaction table for the specified date
            // FIX: Use the starts-with pattern for partition key
            string filter = $"PartitionKey ge '{dateToProcess}-' and PartitionKey lt '{dateToProcess}-\uFFFF'";

            _logger.LogInformation($"Querying daily agent interactions with filter: {filter}");

            var queryResults = copilotAgentInteractionTableClient.QueryAsync<AgentInteraction>(filter);

            // Group the results by agentId (RowKey) and sum the interaction counts
            var agentAggregations = new Dictionary<string, (int TotalInteractions, string AgentName)>();

            await foreach (var queryResult in queryResults)
            {
                if (agentAggregations.ContainsKey(queryResult.RowKey))
                {
                    var existing = agentAggregations[queryResult.RowKey];
                    agentAggregations[queryResult.RowKey] = (
                        existing.TotalInteractions + queryResult.TotalInteractionCount,
                        existing.AgentName
                    );
                }
                else
                {
                    agentAggregations[queryResult.RowKey] = (
                        queryResult.TotalInteractionCount,
                        queryResult.AgentName
                    );
                }
                recordsAdded++;
            }

            _logger.LogInformation($"Processed {recordsAdded} agent interaction records for date: {dateToProcess}");
            _logger.LogInformation($"Found {agentAggregations.Count} unique agents with aggregated interactions");

            // Update the weekly, monthly, and all-time agent aggregation tables
            foreach (var timeFrame in new[] { "weekly", "monthly", "alltime" })
            {
                foreach (var agent in agentAggregations)
                {
                    var agentId = agent.Key;
                    var totalInteractions = agent.Value.TotalInteractions;
                    var agentName = agent.Value.AgentName;

                    TableClient targetTableClient = timeFrame switch
                    {
                        "weekly" => _agentWeeklyTableClient,
                        "monthly" => _agentMonthlyTableClient,
                        "alltime" => _agentAllTimeTableClient,
                        _ => throw new ArgumentException("Invalid time frame")
                    };

                    var partitionKey = timeFrame switch
                    {
                        "weekly" => GetPartitionKeyFromStringDate(dateToProcess, timeFrame),
                        "monthly" => GetPartitionKeyFromStringDate(dateToProcess, timeFrame),
                        "alltime" => AgentInteraction.AllTimePartitionKeyPrefix,
                        _ => throw new ArgumentException("Invalid time frame")
                    };

                    try
                    {
                        // Try to get existing entity
                        var retrieveOp = await targetTableClient.GetEntityIfExistsAsync<AgentInteraction>(partitionKey, agentId);

                        if (retrieveOp.HasValue)
                        {
                            // Entity exists - append by adding to the count
                            var existingEntity = retrieveOp.Value;
                            existingEntity.TotalInteractionCount += totalInteractions;

                            await targetTableClient.UpdateEntityAsync(existingEntity, existingEntity.ETag, TableUpdateMode.Merge);
                            _logger.LogInformation($"Updated {timeFrame} agent {agentId}: total now {existingEntity.TotalInteractionCount}");
                        }
                        else
                        {
                            // Entity doesn't exist - create new
                            var newEntity = new AgentInteraction
                            {
                                PartitionKey = partitionKey,
                                RowKey = agentId,
                                AgentName = agentName,
                                TotalInteractionCount = totalInteractions
                            };

                            await targetTableClient.AddEntityAsync(newEntity);
                            _logger.LogInformation($"Created {timeFrame} agent {agentId}: {totalInteractions}");
                        }


                    }
                    catch (RequestFailedException ex) when (ex.Status == 409)
                    {
                        // Race condition: entity was created between our check and update
                        // Retry with get + update
                        var entity = await targetTableClient.GetEntityAsync<AgentInteraction>(partitionKey, agentId);
                        entity.Value.TotalInteractionCount += totalInteractions;
                        await targetTableClient.UpdateEntityAsync(entity.Value, entity.Value.ETag, TableUpdateMode.Merge);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, $"Error updating {timeFrame} agent {agentId}");
                    }
                }
            }

            return agentAggregations.Count;
        }

        private async Task<List<CopilotTimeFrameUsage>> GetDailyAuditDataForUser(string uPN, string reportDate)
        {
            List<CopilotTimeFrameUsage> copilotInteractions = new List<CopilotTimeFrameUsage>();

            // Log the user we are getting the data for
            _logger.LogInformation($"Retrieving daily audit data for UPN: {uPN}, ReportDate: {reportDate}");

            try
            {
                // Query by partition key only
                string filter = $"PartitionKey eq '{reportDate}-{uPN}'";

                // Log the filter being used
                _logger.LogInformation($"Querying copilot interaction daily aggregations with filter: {filter}");

                var queryResults = copilotInteractionDailyAggregationsTableClient.QueryAsync<CopilotTimeFrameUsage>(filter);

                // Process the results
                await foreach (CopilotTimeFrameUsage entity in queryResults)
                {
                    // Process each entity as needed
                    copilotInteractions.Add(new CopilotTimeFrameUsage
                    {
                        UPN = uPN,
                        App = Enum.TryParse<AppType>(entity.RowKey, out var appType) ? appType : default,
                        TotalDailyActivityCount = entity.TotalDailyActivityCount,
                        TotalInteractionCount = entity.TotalInteractionCount
                    });
                }
            }
            catch (RequestFailedException ex) when (ex.Status == 404)
            {
                _logger.LogWarning($"No entity found for UPN: {uPN}, ReportDate: {reportDate}. Returning null.");
            }
            catch (RequestFailedException ex)
            {
                _logger.LogError($"Failed to retrieve entity for UPN: {uPN}, ReportDate: {reportDate}. Error: {ex.Message}");
            }

            // log the number of interactions retrieved
            _logger.LogInformation($"Retrieved {copilotInteractions.Count} interactions for UPN: {uPN}, ReportDate: {reportDate}");

            return copilotInteractions;
        }

        private UserActivity ConvertToUserActivity(M365CopilotUsage user, List<CopilotTimeFrameUsage> aggregationUsageList)
        {
            // if no aggregation entity, we have no usage
            if (aggregationUsageList.Count == 0)
            {
                return new UserActivity
                {
                    ReportDate = DateTime.SpecifyKind(DateTime.ParseExact(user.ReportRefreshDate, "yyyy-MM-dd", null), DateTimeKind.Utc),
                    UPN = user.UserPrincipalName,
                    DisplayName = user.DisplayName,
                    DailyTeamsActivity = false,
                    DailyOutlookActivity = false,
                    DailyWordActivity = false,
                    DailyExcelActivity = false,
                    DailyPowerPointActivity = false,
                    DailyOneNoteActivity = false,
                    DailyLoopActivity = false,
                    DailyCopilotChatActivity = false,
                    DailyAllUpActivity = false,
                    DailyTeamsInteractionCount = 0,
                    DailyOutlookInteractionCount = 0,
                    DailyWordInteractionCount = 0,
                    DailyExcelInteractionCount = 0,
                    DailyPowerPointInteractionCount = 0,
                    DailyOneNoteInteractionCount = 0,
                    DailyLoopInteractionCount = 0,
                    DailyCopilotChatInteractionCount = 0,
                    DailyAllInteractionCount = 0,
                    DailyMACActivity = false,
                    DailyMACInteractionCount = 0,
                    DailyDesignerActivity = false,
                    DailyDesignerInteractionCount = 0,
                    DailySharePointActivity = false,
                    DailySharePointInteractionCount = 0,
                    DailyPlannerActivity = false,
                    DailyPlannerInteractionCount = 0,
                    DailyWhiteboardActivity = false,
                    DailyWhiteboardInteractionCount = 0,
                    DailyStreamActivity = false,
                    DailyStreamInteractionCount = 0,
                    DailyFormsActivity = false,
                    DailyFormsInteractionCount = 0,
                    DailyCopilotActionActivity = false,
                    DailyCopilotActionCount = 0,
                    DailyWebPluginActivity = false,
                    DailyWebPluginInteractions = 0,
                    DailyAgentActivity = false,
                    DailyAgentInteractions = 0,
                    DailyCopilotStudioActivity = false,
                    DailyCopilotStudioInteractionCount = 0
                };
            }

            // CopilotAllUpActivity
            // if any of the values are true, add another entry for CopilotAllUp
            bool copilotAllUpActivity = false;

            // it's a mere string comparison
            bool DailyUsage(string lastActivityDate, string reportRefreshDate)
            {
                if (lastActivityDate == reportRefreshDate)
                {
                    copilotAllUpActivity = true;
                    return true;
                }
                return false;
            }

            bool DailyUsageFromCount(int count)
            {
                if (count > 0)
                {
                    copilotAllUpActivity = true;
                    return true;
                }
                return false;
            }

            // Turn the aggregation list into a dictionary with app as the key
            // and the value as interaction count
            var aggregationUsageDict = aggregationUsageList
                .GroupBy(x => x.App)
                .ToDictionary(g => g.Key, g => g.Sum(x => x.TotalInteractionCount));

            _logger.LogInformation($"Aggregated {aggregationUsageList.Count} entries into {aggregationUsageDict.Count} unique apps for UPN: {user.UserPrincipalName}");

            int DailyInteractionCountForApp(AppType appType)
            {
                if (aggregationUsageDict.TryGetValue(appType, out var interactionCount))
                {
                    return interactionCount;
                }
                return 0;
            }



            return new UserActivity
            {
                ReportDate = DateTime.SpecifyKind(DateTime.ParseExact(user.ReportRefreshDate, "yyyy-MM-dd", null), DateTimeKind.Utc),
                UPN = user.UserPrincipalName,
                DisplayName = user.DisplayName,
                DailyTeamsActivity = DailyUsage(user.MicrosoftTeamsCopilotLastActivityDate, user.ReportRefreshDate),
                DailyTeamsInteractionCount = DailyInteractionCountForApp(AppType.Teams),
                DailyOutlookActivity = DailyUsage(user.OutlookCopilotLastActivityDate, user.ReportRefreshDate),
                DailyOutlookInteractionCount = DailyInteractionCountForApp(AppType.Outlook),
                DailyWordActivity = DailyUsage(user.WordCopilotLastActivityDate, user.ReportRefreshDate),
                DailyWordInteractionCount = DailyInteractionCountForApp(AppType.Word),
                DailyExcelActivity = DailyUsage(user.ExcelCopilotLastActivityDate, user.ReportRefreshDate),
                DailyExcelInteractionCount = DailyInteractionCountForApp(AppType.Excel),
                DailyPowerPointActivity = DailyUsage(user.PowerPointCopilotLastActivityDate, user.ReportRefreshDate),
                DailyPowerPointInteractionCount = DailyInteractionCountForApp(AppType.PowerPoint),
                DailyOneNoteActivity = DailyUsage(user.OneNoteCopilotLastActivityDate, user.ReportRefreshDate),
                DailyOneNoteInteractionCount = DailyInteractionCountForApp(AppType.OneNote),
                DailyLoopActivity = DailyUsage(user.LoopCopilotLastActivityDate, user.ReportRefreshDate),
                DailyLoopInteractionCount = DailyInteractionCountForApp(AppType.Loop),
                DailyCopilotChatActivity = DailyUsage(user.CopilotChatLastActivityDate, user.ReportRefreshDate),
                DailyCopilotChatInteractionCount = DailyInteractionCountForApp(AppType.CopilotChat),
                DailyAllInteractionCount = DailyInteractionCountForApp(AppType.All),

                DailyMACActivity = DailyUsageFromCount(DailyInteractionCountForApp(AppType.MAC)),
                DailyMACInteractionCount = DailyInteractionCountForApp(AppType.MAC),
                DailyDesignerActivity = DailyUsageFromCount(DailyInteractionCountForApp(AppType.Designer)),
                DailyDesignerInteractionCount = DailyInteractionCountForApp(AppType.Designer),
                DailySharePointActivity = DailyUsageFromCount(DailyInteractionCountForApp(AppType.SharePoint)),
                DailySharePointInteractionCount = DailyInteractionCountForApp(AppType.SharePoint),
                DailyPlannerActivity = DailyUsageFromCount(DailyInteractionCountForApp(AppType.Planner)),
                DailyPlannerInteractionCount = DailyInteractionCountForApp(AppType.Planner),
                DailyWhiteboardActivity = DailyUsageFromCount(DailyInteractionCountForApp(AppType.Whiteboard)),
                DailyWhiteboardInteractionCount = DailyInteractionCountForApp(AppType.Whiteboard),
                DailyStreamActivity = DailyUsageFromCount(DailyInteractionCountForApp(AppType.Stream)),
                DailyStreamInteractionCount = DailyInteractionCountForApp(AppType.Stream),
                DailyFormsActivity = DailyUsageFromCount(DailyInteractionCountForApp(AppType.Forms)),
                DailyFormsInteractionCount = DailyInteractionCountForApp(AppType.Forms),
                DailyCopilotActionActivity = DailyUsageFromCount(DailyInteractionCountForApp(AppType.CopilotAction)),
                DailyCopilotActionCount = DailyInteractionCountForApp(AppType.CopilotAction),
                DailyWebPluginActivity = DailyUsageFromCount(DailyInteractionCountForApp(AppType.WebPlugin)),
                DailyWebPluginInteractions = DailyInteractionCountForApp(AppType.WebPlugin),
                DailyAgentActivity = DailyUsageFromCount(DailyInteractionCountForApp(AppType.Agent)),
                DailyAgentInteractions = DailyInteractionCountForApp(AppType.Agent),
                DailyCopilotStudioActivity = DailyUsageFromCount(DailyInteractionCountForApp(AppType.CopilotStudio)),
                DailyCopilotStudioInteractionCount = DailyInteractionCountForApp(AppType.CopilotStudio),

                DailyAllUpActivity = DailyUsageFromCount(DailyInteractionCountForApp(AppType.All))
            };
        }

        private Dictionary<AppType, Tuple<bool, int>> ConvertToUsageDictionary(UserActivity userActivity)
        {

            var usage = new Dictionary<AppType, Tuple<bool, int>>
            {
                { AppType.Teams, new Tuple<bool,int>(userActivity.DailyTeamsActivity, userActivity.DailyTeamsInteractionCount) },
                { AppType.Outlook, new Tuple<bool,int>(userActivity.DailyOutlookActivity, userActivity.DailyOutlookInteractionCount) },
                { AppType.Word, new Tuple<bool,int>(userActivity.DailyWordActivity, userActivity.DailyWordInteractionCount) },
                { AppType.Excel, new Tuple<bool,int>(userActivity.DailyExcelActivity, userActivity.DailyExcelInteractionCount) },
                { AppType.PowerPoint, new Tuple<bool,int>(userActivity.DailyPowerPointActivity, userActivity.DailyPowerPointInteractionCount) },
                { AppType.OneNote, new Tuple<bool,int>(userActivity.DailyOneNoteActivity, userActivity.DailyOneNoteInteractionCount) },
                { AppType.Loop, new Tuple<bool,int>(userActivity.DailyLoopActivity, userActivity.DailyLoopInteractionCount) },
                { AppType.CopilotChat, new Tuple<bool,int>(userActivity.DailyCopilotChatActivity, userActivity.DailyCopilotChatInteractionCount) },
                { AppType.All, new Tuple<bool,int>(userActivity.DailyAllUpActivity, userActivity.DailyAllInteractionCount) },
                { AppType.MAC, new Tuple<bool,int>(userActivity.DailyMACActivity, userActivity.DailyMACInteractionCount) },
                { AppType.Designer, new Tuple<bool,int>(userActivity.DailyDesignerActivity, userActivity.DailyDesignerInteractionCount) },
                { AppType.SharePoint, new Tuple<bool,int>(userActivity.DailySharePointActivity, userActivity.DailySharePointInteractionCount) },
                { AppType.Planner, new Tuple<bool,int>(userActivity.DailyPlannerActivity, userActivity.DailyPlannerInteractionCount) },
                { AppType.Whiteboard, new Tuple<bool,int>(userActivity.DailyWhiteboardActivity, userActivity.DailyWhiteboardInteractionCount) },
                { AppType.Stream, new Tuple<bool,int>(userActivity.DailyStreamActivity, userActivity.DailyStreamInteractionCount) },
                { AppType.Forms, new Tuple<bool,int>(userActivity.DailyFormsActivity, userActivity.DailyFormsInteractionCount) },
                { AppType.CopilotAction, new Tuple<bool,int>(userActivity.DailyCopilotActionActivity, userActivity.DailyCopilotActionCount) },
                { AppType.WebPlugin, new Tuple<bool,int>(userActivity.DailyWebPluginActivity, userActivity.DailyWebPluginInteractions) },
                { AppType.Agent, new Tuple<bool,int>(userActivity.DailyAgentActivity, userActivity.DailyAgentInteractions) },
                { AppType.CopilotStudio, new Tuple<bool,int>(userActivity.DailyCopilotStudioActivity, userActivity.DailyCopilotStudioInteractionCount) }
            };

            return usage;
        }

        private async Task<List<AgentInteraction>> GetDailyAgentDataForUser(string uPN, string reportDate)
        {
            List<AgentInteraction> agentInteractions = new List<AgentInteraction>();

            try
            {
                // Query by partition key only
                string filter = $"PartitionKey eq '{reportDate}-{uPN}'";
                var queryResults = copilotAgentInteractionTableClient.QueryAsync<AgentInteraction>(filter);

                // Process the results
                await foreach (AgentInteraction entity in queryResults)
                {
                    // Process each entity as needed
                    agentInteractions.Add(entity);
                }
            }
            catch (RequestFailedException ex) when (ex.Status == 404)
            {
                _logger.LogWarning($"No entity found for UPN: {uPN}, ReportDate: {reportDate}. Returning null.");
            }
            catch (RequestFailedException ex)
            {
                _logger.LogError($"Failed to retrieve entity for UPN: {uPN}, ReportDate: {reportDate}. Error: {ex.Message}");

            }

            return agentInteractions;
        }

        private async Task UpdateReportRefreshDate(string reportRefreshDate, string timeFrame)
        {
            // switch on timeFrame to determine startdate
            timeFrame = timeFrame.ToLowerInvariant();
            var startDate = timeFrame switch
            {
                "daily" => reportRefreshDate,
                "weekly" => GetWeekStartDate(DateTime.ParseExact(reportRefreshDate, "yyyy-MM-dd", null)),
                "monthly" => GetMonthStartDate(DateTime.ParseExact(reportRefreshDate, "yyyy-MM-dd", null)),
                _ => throw new ArgumentException("Invalid timeFrame")
            };

            try
            {
                // Create the daily refresh date
                var tableEntity = new TableEntity("ReportRefreshDate", timeFrame)
                {
                    { "ReportRefreshDate", reportRefreshDate },
                    { "StartDate", startDate }
                };

                // Try to add the entity if it doesn't exist
                await _reportRefreshDateTableClient.AddEntityAsync(tableEntity);

            }
            catch (Azure.RequestFailedException ex) when (ex.Status == 409) // Conflict indicates the entity already exists
            {
                // Merge the entity if it already exists
                var existingTableEntity = await _reportRefreshDateTableClient.GetEntityAsync<TableEntity>("ReportRefreshDate", timeFrame);

                existingTableEntity.Value["ReportRefreshDate"] = reportRefreshDate;
                existingTableEntity.Value["StartDate"] = startDate;

                await _reportRefreshDateTableClient.UpdateEntityAsync(existingTableEntity.Value, ETag.All, TableUpdateMode.Merge);
            }
        }

        static string GetWeekStartDate(DateTime date)
        {
            // Get the Monday of the current week
            var dayOfWeek = (int)date.DayOfWeek;
            // Monday
            //var daysToSubtract = dayOfWeek == 0 ? 6 : dayOfWeek - 1; // Adjust for Sunday
            // Sunday
            var daysToSubtract = dayOfWeek; // Adjust for Sunday
            return date.AddDays(-daysToSubtract).ToString("yyyy-MM-dd");
        }

        static string GetMonthStartDate(DateTime date)
        {
            // Get the first day of the current month
            return date
                .AddDays(-1 * date.Day + 1)
                .ToString("yyyy-MM-dd");
        }

        private async Task UpdateAgentSnapshots(List<AgentInteraction> usersDailyAgentActivity, string timeFrame, string upn, string startDate)
        {
            TableClient tableClientByUser = _agentAllTimeByUserTableClient;
            string partitionKey = string.Empty;

            // get the table
            switch (timeFrame)
            {
                case "weekly":

                    tableClientByUser = _agentWeeklyByUserTableClient;
                    partitionKey = startDate;
                    break;
                case "monthly":
                    tableClientByUser = _agentMonthlyByUserTableClient;
                    partitionKey = startDate;
                    break;
                default:
                    partitionKey = AgentInteraction.AllTimePartitionKeyPrefix;
                    break;
            }

            foreach (var agentDailyInteractions in usersDailyAgentActivity)
            {
                int interactionCount = agentDailyInteractions.TotalInteractionCount;

                // At this point row key is the agentId
                var agentPartitionKey = $"{partitionKey}-{agentDailyInteractions.RowKey}";

                try
                {
                    // first Update the users all time usage of the agent - 
                    var existingTableEntity = await tableClientByUser.GetEntityAsync<AgentInteraction>(agentPartitionKey, upn);

                    // Increment the daily all time activity count
                    existingTableEntity.Value.TotalDailyActivityCount = existingTableEntity.Value.TotalDailyActivityCount + 1;
                    existingTableEntity.Value.TotalInteractionCount = existingTableEntity.Value.TotalInteractionCount + interactionCount;

                    await tableClientByUser.UpdateEntityAsync(existingTableEntity.Value, ETag.All, TableUpdateMode.Merge);

                }
                catch (RequestFailedException ex) when (ex.Status == 404)
                {

                    // if usage create first item
                    var newAllTimeUsage = new AgentInteraction()
                    {
                        UPN = upn,
                        AgentId = agentDailyInteractions.RowKey,
                        AgentName = agentDailyInteractions.AgentName,
                        TotalDailyActivityCount = 1,
                        TotalInteractionCount = interactionCount
                    };

                    if (timeFrame == "alltime")
                    {
                        await tableClientByUser.AddEntityAsync(newAllTimeUsage.ToAllTimeTableEntity());
                    }
                    else
                    {
                        await tableClientByUser.AddEntityAsync(newAllTimeUsage.ToTimeFrameTableEntity(startDate));
                    }

                }
                catch (RequestFailedException ex)
                {
                    Console.WriteLine($"Error updating agent records: {ex.Message}");
                }
            }
        }
        private async Task UpdateUserSnapshots(Dictionary<AppType, Tuple<bool, int>> userActivityDictionary, string upn, string timeFrame, string startDate)
        {
            // get the table
            var tableClient = _userAllTimeTableClient;
            string partitionKey = $"{CopilotTimeFrameUsage.AllTimePartitionKeyPrefix}";

            if (timeFrame == "weekly")
            {
                tableClient = _userWeeklyTableClient;
                partitionKey = $"{startDate}";
            }
            else if (timeFrame == "monthly")
            {
                tableClient = _userMonthlyTableClient;
                partitionKey = $"{startDate}";
            }

            foreach (var (app, tuple) in userActivityDictionary)
            {
                bool dailyUsage = tuple.Item1;
                int interactionCount = tuple.Item2;

                string appPartitionKey = $"{partitionKey}-{app}";

                try
                {
                    // Get the existing entity
                    var existingTableEntity = await tableClient.GetEntityAsync<CopilotTimeFrameUsage>(appPartitionKey, upn);

                    // If usage 
                    if (dailyUsage)
                    {
                        // log that we are 
                        _logger.LogInformation($"Updating {app} usage for {upn} on {timeFrame} with interaction count {interactionCount}.");

                        // Increment the daily all time activity count
                        existingTableEntity.Value.TotalDailyActivityCount = existingTableEntity.Value.TotalDailyActivityCount + 1;
                        existingTableEntity.Value.CurrentDailyStreak = existingTableEntity.Value.CurrentDailyStreak + 1;
                        existingTableEntity.Value.BestDailyStreak = Math.Max(
                            existingTableEntity.Value.BestDailyStreak ?? 0,
                            existingTableEntity.Value.CurrentDailyStreak ?? 0);
                        existingTableEntity.Value.TotalInteractionCount = existingTableEntity.Value.TotalInteractionCount + interactionCount;

                        await tableClient.UpdateEntityAsync(existingTableEntity.Value, existingTableEntity.Value.ETag, TableUpdateMode.Replace);

                        // Log the update
                        _logger.LogInformation($"Updated {app} usage for {upn} on {timeFrame} with interaction count {interactionCount}. " +
                                               $"TotalDailyActivityCount: {existingTableEntity.Value.TotalDailyActivityCount}, " +
                                               $"CurrentDailyStreak: {existingTableEntity.Value.CurrentDailyStreak}, " +
                                               $"BestDailyStreak: {existingTableEntity.Value.BestDailyStreak}, " +
                                               $"TotalInteractionCount: {existingTableEntity.Value.TotalInteractionCount}");
                    }
                    else
                    {
                        // Reset the current daily streak
                        existingTableEntity.Value.CurrentDailyStreak = 0;
                        await tableClient.UpdateEntityAsync(existingTableEntity.Value, existingTableEntity.Value.ETag, TableUpdateMode.Merge);
                    }
                }
                catch (RequestFailedException ex) when (ex.Status == 404)
                {

                    // Only create if usage
                    if (dailyUsage)
                    {
                        // log that we are creating a new entity
                        _logger.LogInformation($"Creating new {app} usage for {upn} on {timeFrame} with interaction count {interactionCount}.");

                        var newAllTimeUsage = new CopilotTimeFrameUsage()
                        {
                            UPN = upn,
                            App = app,
                            TotalDailyActivityCount = 1,
                            TotalInteractionCount = interactionCount,
                            CurrentDailyStreak = 1,
                            BestDailyStreak = 1
                        };

                        if (timeFrame == "alltime")
                        {
                            await _userAllTimeTableClient.AddEntityAsync(newAllTimeUsage.ToAllTimeTableEntity());
                        }
                        else
                        {
                            await tableClient.AddEntityAsync(newAllTimeUsage.ToTimeFrameTableEntity(startDate));
                        }

                        // Log the creation
                        _logger.LogInformation($"Created new {app} usage for {upn} on {timeFrame} with interaction count {interactionCount}. " +
                                            $"TotalDailyActivityCount: 1, " +
                                            $"CurrentDailyStreak: 1, " +
                                            $"BestDailyStreak: 1, " +
                                            $"TotalInteractionCount: {interactionCount}");
                    }



                }
                catch (RequestFailedException ex)
                {
                    _logger.LogError("Error getting notifications. Message: {Message}", ex.Message);
                    _logger.LogError("Stack Trace: {StackTrace}", ex.StackTrace);
                    _logger.LogError("Inner Exception: {InnerException}", ex.InnerException?.Message);
                    _logger.LogError("Inner Exception Stack Trace: {InnerExceptionStackTrace}", ex.InnerException?.StackTrace);
                }
            }
        }

        public async Task<List<string>> GetUsersWithStreakForApp(string app, int count)
        {
            // users
            var users = new List<string>();

            // build the query filter
            string filter = $"PartitionKey eq '{CopilotTimeFrameUsage.AllTimePartitionKeyPrefix}-{app}' and CurrentDailyStreak ge {count}";

            _logger.LogInformation($"Filter: {filter}");

            try
            {
                var queryResults = _userAllTimeTableClient.QueryAsync<TableEntity>(filter);

                await foreach (TableEntity entity in queryResults)
                {
                    users.Add(entity.RowKey);
                }
            }
            catch (RequestFailedException ex)
            {
                Console.WriteLine($"Error retrieving records: {ex.Message}");

            }

            return users;
        }

        public async Task<string?> GetStartDate(string timeFrame)
        {
            // Get the report refresh date for the timeFrame
            try
            {
                var existingTableEntity = await _reportRefreshDateTableClient.GetEntityAsync<TableEntity>("ReportRefreshDate", timeFrame);
                return existingTableEntity.Value["StartDate"].ToString();
            }
            catch (RequestFailedException ex) when (ex.Status == 404)
            {
                // Entity not found - nothing to do
                return null;
            }
        }

        public async Task<List<string>> GetUsersWhoHaveCompletedActivityForApp(string app, string? dayCount, string? interactionCount, string timeFrame, string date)
        {
            // switch statement to get the correct table on timeFrame
            var tableClient = timeFrame.ToLowerInvariant() switch
            {
                "weekly" => _userWeeklyTableClient,
                "monthly" => _userMonthlyTableClient,
                "alltime" => _userAllTimeTableClient,
                _ => throw new ArgumentException("Invalid timeFrame")
            };

            if (timeFrame.ToLowerInvariant() == "alltime")
            {
                date = CopilotTimeFrameUsage.AllTimePartitionKeyPrefix;
            }

            // Define the query filter 
            string filter = $"PartitionKey eq '{date}-{app}'";

            if (!string.IsNullOrEmpty(dayCount))
            {
                filter += $" and TotalDailyActivityCount ge {dayCount}";
            }

            if (!string.IsNullOrEmpty(interactionCount))
            {
                filter += $" and TotalInteractionCount ge {interactionCount}";
            }

            // log the filter
            _logger.LogInformation($"Filter for {timeFrame}: {filter}");

            // Get the users
            var users = new List<string>();
            try
            {
                // Query all records with filter
                AsyncPageable<TableEntity> queryResults = tableClient.QueryAsync<TableEntity>(filter);

                await foreach (TableEntity entity in queryResults)
                {
                    users.Add(entity.RowKey);
                }
            }
            catch (RequestFailedException ex)
            {
                Console.WriteLine($"Error retrieving records: {ex.Message}");
            }

            return users;

        }

        // Todo, architect a better storage logic, query is not efficient
        public async Task<List<InactiveUser>> GetInactiveUsers(int days)
        {
            // query to find user with more than days inactivity
            string filter = TableClient.CreateQueryFilter($"DaysSinceLastActivity ge {days}");

            // Get the users
            var users = new List<InactiveUser>();

            _logger.LogInformation($"Filter: {filter}");

            try
            {
                var queryResults = _userLastUsageTableClient.QueryAsync<TableEntity>(filter);

                await foreach (TableEntity entity in queryResults)
                {

                    users.Add(new InactiveUser
                    {
                        UPN = entity.PartitionKey,
                        DaysSinceLastActivity = entity.GetDouble("DaysSinceLastActivity") ?? 0,
                    });
                }
            }
            catch (RequestFailedException ex)
            {
                Console.WriteLine($"Error retrieving records: {ex.Message}");

            }

            return users;
        }

        private string GetPartitionKeyFromStringDate(string dateString, string timeFrame)
        {
            DateTime date = DateTime.ParseExact(dateString, "yyyy-MM-dd", null);

            return timeFrame.ToLowerInvariant() switch
            {
                "daily" => date.ToString("yyyy-MM-dd"),
                "weekly" => GetWeekStartDate(date),
                "monthly" => GetMonthStartDate(date),
                _ => throw new ArgumentException("Invalid timeFrame")
            };
        }

        private async Task<string> GetActiveTimeFrame(string timeFrame)
        {
            try
            {
                // Use direct entity lookup instead of query for better performance
                var response = await _reportRefreshDateTableClient.GetEntityIfExistsAsync<TableEntity>("ReportRefreshDate", timeFrame);

                if (response.HasValue)
                {
                    var startDate = response.Value.GetString("StartDate");
                    if (!string.IsNullOrEmpty(startDate))
                    {
                        return startDate;
                    }

                    _logger.LogWarning($"StartDate is null or empty for timeFrame {timeFrame}");
                }
                else
                {
                    _logger.LogWarning($"No active time frame entity found for {timeFrame}");
                }

                // Return empty string if no valid data found
                return string.Empty;
            }
            catch (RequestFailedException ex) when (ex.Status == 404)
            {
                _logger.LogWarning($"No active time frame found for {timeFrame} (404). Returning empty string.");
                return string.Empty;
            }
            catch (RequestFailedException ex)
            {
                _logger.LogError(ex, $"Request failed while retrieving active time frame for {timeFrame}. Status: {ex.Status}");
                return string.Empty; // Don't throw, return empty to allow fallback logic
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Unexpected error retrieving active time frame for {timeFrame}");
                return string.Empty; // Don't throw, return empty to allow fallback logic
            }
        }

        #endregion

        #region User Seeding
        public async Task SeedMonthlyTimeFrameActivitiesAsync(List<CopilotTimeFrameUsage> userActivitiesSeed, string startDate)
        {
            // Get daily table
            foreach (var userEntity in userActivitiesSeed)
            {
                try
                {
                    // Try to add the entity if it doesn't exist
                    await _userMonthlyTableClient.AddEntityAsync(userEntity.ToTimeFrameTableEntity(startDate));
                    _logger.LogInformation($"Added monthly seed entity for {userEntity.UPN}");
                }
                catch (Azure.RequestFailedException ex) when (ex.Status == 409) // Conflict indicates the entity already exists
                {
                    // Merge the entity if it already exists
                    await _userMonthlyTableClient.UpdateEntityAsync(userEntity.ToTimeFrameTableEntity(startDate), ETag.All, TableUpdateMode.Merge);
                }
            }
        }

        public async Task SeedWeeklyTimeFrameActivitiesAsync(List<CopilotTimeFrameUsage> userActivitiesSeed, string startDate)
        {
            // Get daily table
            foreach (var userEntity in userActivitiesSeed)
            {
                try
                {
                    // Try to add the entity if it doesn't exist
                    await _userWeeklyTableClient.AddEntityAsync(userEntity.ToTimeFrameTableEntity(startDate));
                    _logger.LogInformation($"Added weekly seed entity for {userEntity.UPN}");
                }
                catch (Azure.RequestFailedException ex) when (ex.Status == 409) // Conflict indicates the entity already exists
                {
                    // Merge the entity if it already exists
                    await _userWeeklyTableClient.UpdateEntityAsync(userEntity.ToTimeFrameTableEntity(startDate), ETag.All, TableUpdateMode.Merge);
                }
            }
        }

        public async Task SeedAllTimeActivityAsync(List<CopilotTimeFrameUsage> userActivitiesSeed)
        {
            // Get daily table
            foreach (var userEntity in userActivitiesSeed)
            {
                try
                {
                    // Try to add the entity if it doesn't exist
                    await _userAllTimeTableClient.AddEntityAsync(userEntity.ToAllTimeTableEntity());
                    _logger.LogInformation($"Added alltime seed entity for {userEntity.UPN}");
                }
                catch (Azure.RequestFailedException ex) when (ex.Status == 409) // Conflict indicates the entity already exists
                {
                    // Merge the entity if it already exists
                    await _userAllTimeTableClient.UpdateEntityAsync(userEntity.ToAllTimeTableEntity(), ETag.All, TableUpdateMode.Merge);
                }
            }
        }

        public async Task SeedInactiveUsersAsync(List<InactiveUser> userActivitiesSeed)
        {
            // Get daily table
            foreach (var userEntity in userActivitiesSeed)
            {
                // Add user record
                var tableEntity = new TableEntity(userEntity.UPN, userEntity.LastActivityDate.ToString("yyyy-MM-dd"))
                {
                    { "DaysSinceLastActivity", userEntity.DaysSinceLastActivity },
                    { "LastActivityDate", userEntity.LastActivityDate },
                    { "DisplayName", userEntity.DisplayName }
                };

                try
                {
                    // Try to add the entity if it doesn't exist
                    await _userLastUsageTableClient.AddEntityAsync(tableEntity);
                    _logger.LogInformation($"Added inactive seed entity for {userEntity.UPN}");
                }
                catch (Azure.RequestFailedException ex) when (ex.Status == 409) // Conflict indicates the entity already exists
                {
                    // Merge the entity if it already exists
                    await _userLastUsageTableClient.UpdateEntityAsync(tableEntity, ETag.All, TableUpdateMode.Merge);
                }
            }
        }
        #endregion
    }

}