using groveale.Models;
using groveale.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace groveale
{
    public class ReceiveEvents
    {
        private readonly ILogger<ReceiveEvents> _logger;
        private readonly ISettingsService _settingsService;
        private readonly IM365ActivityService _m365ActivityService;
        private readonly IAzureTableService _azureTableService; 
        private readonly IKeyVaultService _keyVaultService;
        private readonly IExclusionEmailService _exclusionEmailService;
        private readonly IQueueService _queueService;

        public ReceiveEvents(ILogger<ReceiveEvents> logger, ISettingsService settingsService, IM365ActivityService m365ActivityService, IAzureTableService azureTableService, IKeyVaultService keyVaultService, IExclusionEmailService exclusionEmailService, IQueueService queueService)
        {
            _logger = logger;
            _settingsService = settingsService;
            _m365ActivityService = m365ActivityService;
            _azureTableService = azureTableService;
            _keyVaultService = keyVaultService;
            _exclusionEmailService = exclusionEmailService;
            _queueService = queueService;

        }

        [Function("ReceiveEvents")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequest req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");

            // check if the webhook is paused
            bool webhookPaused = await _azureTableService.GetWebhookStateAsync();

            _logger.LogInformation($"Webhook state: {webhookPaused}");

            if (webhookPaused)
            {
                _logger.LogInformation("Webhook is paused, returning 503.");
                // Log the webhook trigger event
                await _azureTableService.LogWebhookTriggerAsync(new LogEvent
                {
                    EventId = Guid.NewGuid().ToString(),
                    EventName = "WebhookTrigger",
                    EventMessage = "Webhook triggered",
                    EventDetails = "Webhook is paused",
                    EventCategory = "Webhook",
                    EventTime = DateTime.UtcNow
                });
                return new StatusCodeResult(StatusCodes.Status503ServiceUnavailable);
            }

            try
            {
                // Log the webhook trigger event
                await _azureTableService.LogWebhookTriggerAsync(new LogEvent
                {
                    EventId = Guid.NewGuid().ToString(),
                    EventName = "WebhookTrigger",
                    EventMessage = "Webhook triggered",
                    EventDetails = "Webhook event",
                    EventCategory = "Webhook",
                    EventTime = DateTime.UtcNow
                });
            }
            catch
            {
                _logger.LogError("Error logging event.");
                // continue
            }

            // validate that the headers contains the correct Webhook-AuthID and matches the configured value
            var authId = req.Headers["Webhook-AuthID"];
            if (string.IsNullOrEmpty(authId) || authId != _settingsService.AuthGuid)
            {
                _logger.LogError("Invalid Webhook-AuthID header.");
                // Log the webhook trigger event
                await _azureTableService.LogWebhookTriggerAsync(new LogEvent
                {
                    EventId = Guid.NewGuid().ToString(),
                    EventName = "WebhookTrigger",
                    EventMessage = "Webhook triggered",
                    EventDetails = "Invalid Webhook-AuthID header",
                    EventCategory = "Webhook",
                    EventTime = DateTime.UtcNow
                });
                return new BadRequestObjectResult("Invalid Webhook-AuthID header.");
            }

            // Parse the request body to extract the validation code from the payload
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();

            try 
            {
                var payload = JObject.Parse(requestBody);
                var validationCodeFromPayload = payload["validationCode"]?.ToString();

                // If the validation code is present, validate it against the Webhook-ValidationCode header
                // This is M365 initial validation request
                if (!string.IsNullOrEmpty(validationCodeFromPayload))
                {
                    var validationHeaderCode = req.Headers["Webhook-ValidationCode"];
                    if (validationHeaderCode != validationCodeFromPayload)
                    {
                        _logger.LogError("Invalid Webhook-ValidationCode header.");
                        return new BadRequestObjectResult("Invalid Webhook-ValidationCode header.");
                    }
                    return new OkResult();
                }
            }
            catch (JsonReaderException)
            {
                // The payload is not a JSON object, its a list of notifications (Hopefully)
                // Could probably use a better way to determine if the payload is a list of notifications
            }
            
            try
            {
                // Deserialize the request body into a list of NotificationResponse objects
                var notifications = JsonConvert.DeserializeObject<List<NotificationResponse>>(requestBody);

                // Create the EncyptionService
                // Encrypt the upn, data added to the tables will be encrypted
                var encryptionService = await DeterministicEncryptionService.CreateAsync(_settingsService, _keyVaultService);

                // Just call it - caching happens automatically
                var exclusionEmails = await _exclusionEmailService.LoadEmailsFromPersonFieldAsync();

                // Get copilot audit records
                var copilotAuditRecords = await _m365ActivityService.GetCopilotActivityNotificationsAsync(notifications, encryptionService, exclusionEmails);

                // Store interaction details using batch writes (up to 100 per transaction)
                await _azureTableService.AddBatchCopilotInteractionDetailsAsync(copilotAuditRecords);

                // Build pre-aggregated messages (single pass over events) and queue for async processing
                var aggregationMessages = EventAggregationHelper.BuildAggregationMessages(copilotAuditRecords, _logger);
                await _queueService.QueueCopilotEventAggregationsAsync(aggregationMessages);

                // Log reconciliation data
                var uniqueUsers = copilotAuditRecords.Select(r => r.UserId).Distinct().Count();
                _logger.LogInformation(
                    "Webhook batch processed: {TotalEvents} events, {UniqueUsers} unique users, {QueueMessages} aggregation messages queued",
                    copilotAuditRecords.Count, uniqueUsers, aggregationMessages.Count);

                try
                {
                    await _azureTableService.LogWebhookTriggerAsync(new LogEvent
                    {
                        EventId = Guid.NewGuid().ToString(),
                        EventName = "WebhookBatchProcessed",
                        EventMessage = "Webhook batch processed",
                        EventDetails = $"Events: {copilotAuditRecords.Count}, Users: {uniqueUsers}, Queued: {aggregationMessages.Count}",
                        EventCategory = "Webhook",
                        EventTime = DateTime.UtcNow
                    });
                }
                catch
                {
                    _logger.LogError("Error logging batch processed event.");
                }

                return new OkResult();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing notifications.");
                return new StatusCodeResult(StatusCodes.Status500InternalServerError);
            }
        }
    }
}
