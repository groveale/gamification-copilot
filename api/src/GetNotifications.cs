using groveale.Models;
using groveale.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace groveale
{
    public class GetNotifications
    {
        private readonly ILogger<GetNotifications> _logger;
        private readonly IM365ActivityService _m365ActivityService;

        private readonly IAzureTableService _azureTableService;
        private readonly IKeyVaultService _keyVaultService;
        private readonly ISettingsService _settingsService;
        private readonly IExclusionEmailService _exclusionEmailService;
        private readonly IQueueService _queueService;

        public GetNotifications(ILogger<GetNotifications> logger, IM365ActivityService m365ActivityService, IAzureTableService azureTableService, IKeyVaultService keyVaultService, ISettingsService settingsService, IExclusionEmailService exclusionEmailService, IQueueService queueService)
        {
            _logger = logger;
            _m365ActivityService = m365ActivityService;
            _azureTableService = azureTableService;
            _keyVaultService = keyVaultService;
            _settingsService = settingsService;
            _exclusionEmailService = exclusionEmailService;
            _queueService = queueService;
        }

        [Function("GetNotifications")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Admin, "get", "post")] HttpRequest req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");

            // Get content type from query parameters
            string contentType = req.Query["contentType"];
            if (string.IsNullOrEmpty(contentType))
            {
                _logger.LogError("Audit content type is not provided.");
                return new BadRequestObjectResult("Audit content type is not provided.");
            }

            try
            {
                var notifications = await _m365ActivityService.GetAvailableNotificationsAsync(contentType);

                if (notifications == null || notifications.Count == 0)
                {
                    _logger.LogInformation("No notifications found.");
                    return new OkObjectResult("No notifications found.");
                }
                _logger.LogInformation($"Found {notifications.Count} notifications.");

                var exclusionEmails = await _exclusionEmailService.LoadEmailsFromPersonFieldAsync();

                // log the exclusion emails
                // TODO remove this in production
                _logger.LogInformation("Exclusion Emails: {ExclusionEmails}", string.Join(", ", exclusionEmails));

                // Create the EncryptionService
                _logger.LogInformation("Attempting to create EncryptionService");
                var encryptionService = await DeterministicEncryptionService.CreateAsync(_settingsService, _keyVaultService);
                _logger.LogInformation("Encryption Service created");

                // log an encryption test
                var test = encryptionService.Encrypt("test");
                _logger.LogInformation("Encryption Test: {Test}", test);



                // Get copilot audit records
                var copilotAuditRecords = await _m365ActivityService.GetCopilotActivityNotificationsAsync(notifications, encryptionService, exclusionEmails);

                // Store interaction details using batch writes
                await _azureTableService.AddBatchCopilotInteractionDetailsAsync(copilotAuditRecords);

                // Build pre-aggregated messages (single pass) and queue for async processing
                var aggregationMessages = EventAggregationHelper.BuildAggregationMessages(copilotAuditRecords, _logger);
                await _queueService.QueueCopilotEventAggregationsAsync(aggregationMessages);

                var uniqueUsers = copilotAuditRecords.Select(r => r.UserId).Distinct().Count();
                _logger.LogInformation(
                    "GetNotifications processed: {TotalEvents} events, {UniqueUsers} unique users, {QueueMessages} aggregation messages queued",
                    copilotAuditRecords.Count, uniqueUsers, aggregationMessages.Count);

                return new OkObjectResult(new { Events = copilotAuditRecords.Count, Users = uniqueUsers, Queued = aggregationMessages.Count });
            }
            catch (Exception ex)
            {
                _logger.LogError("Error getting notifications. Message: {Message}", ex.Message);
                _logger.LogError("Stack Trace: {StackTrace}", ex.StackTrace);
                _logger.LogError("Inner Exception: {InnerException}", ex.InnerException?.Message);
                _logger.LogError("Inner Exception Stack Trace: {InnerExceptionStackTrace}", ex.InnerException?.StackTrace);
                return new StatusCodeResult(StatusCodes.Status500InternalServerError);
            }
        }
    }
}
