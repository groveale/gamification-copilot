using Azure.Identity;
using Azure.Storage.Queues;
using groveale.Models;
using Microsoft.Extensions.Logging;
using System;
using System.Text.Json;
using System.Threading.Tasks;

namespace groveale.Services
{
    public interface IQueueService
    {
        Task QueueUserAggregationsAsync(List<string> encryptedUPNs, string reportRefreshDate);
    }

    public class QueueService : IQueueService
    {
        private readonly QueueServiceClient _queueServiceClient;
        private readonly ILogger<QueueService> _logger;
        private readonly string _userAggregationQueueName;

        public QueueService(ISettingsService settingsService, ILogger<QueueService> logger)
        {
            _logger = logger;
            
            var storageAccountName = settingsService.StorageAccountName;
            var connectionString = Environment.GetEnvironmentVariable("AzureWebJobsStorage");

            _logger.LogInformation("Initializing QueueService...");
            _logger.LogInformation($"StorageAccountName from settings: {storageAccountName ?? "NULL"}");
            _logger.LogInformation($"AzureWebJobsStorage configured: {!string.IsNullOrEmpty(connectionString)}");

            // Use connection string for local development (Azurite), URI + Managed Identity for Azure
            if (!string.IsNullOrEmpty(connectionString) && 
                (connectionString.Contains("UseDevelopmentStorage=true") || connectionString.Contains("AccountKey=")))
            {
                _logger.LogInformation("Using connection string authentication (local/dev)");
                _queueServiceClient = new QueueServiceClient(connectionString, new QueueClientOptions
                {
                    MessageEncoding = QueueMessageEncoding.Base64
                });
            }
            else if (!string.IsNullOrEmpty(storageAccountName))
            {
                // Queue endpoint is different from table endpoint
                var queueUri = $"https://{storageAccountName}.queue.core.windows.net/";
                _logger.LogInformation($"Using Managed Identity authentication with queue endpoint: {queueUri}");
                _queueServiceClient = new QueueServiceClient(
                    new Uri(queueUri),
                    new DefaultAzureCredential(),
                    new QueueClientOptions
                    {
                        MessageEncoding = QueueMessageEncoding.Base64
                    });
            }
            else
            {
                _logger.LogError("Neither StorageAccountName nor AzureWebJobsStorage is configured");
                throw new InvalidOperationException(
                    "Either StorageAccountName or AzureWebJobsStorage must be configured");
            }

            _userAggregationQueueName = settingsService.UserAggregationsQueueName;
            
            _logger.LogInformation($"QueueServiceClient created successfully. Queue name: {_userAggregationQueueName}");
            _logger.LogInformation($"Queue URI: {_queueServiceClient?.Uri?.ToString() ?? "NULL"}");
        }

        public async Task QueueUserAggregationsAsync(List<string> encryptedUPNs, string reportRefreshDate)
        {
            try
            {
                // Validate dependencies
                if (_queueServiceClient == null)
                {
                    _logger.LogError("QueueServiceClient is null");
                    throw new InvalidOperationException("QueueServiceClient is not initialized");
                }

                if (string.IsNullOrEmpty(_userAggregationQueueName))
                {
                    _logger.LogError("UserAggregationQueueName is null or empty");
                    throw new InvalidOperationException("UserAggregationQueueName is not configured");
                }

                if (encryptedUPNs == null)
                {
                    _logger.LogError("encryptedUPNs parameter is null");
                    throw new ArgumentNullException(nameof(encryptedUPNs));
                }

                if (string.IsNullOrEmpty(reportRefreshDate))
                {
                    _logger.LogError("reportRefreshDate parameter is null or empty");
                    throw new ArgumentException("reportRefreshDate cannot be null or empty", nameof(reportRefreshDate));
                }

                _logger.LogInformation($"Queuing {encryptedUPNs.Count} user aggregation messages for date: {reportRefreshDate}");

                // Get or create the queue
                _logger.LogInformation($"Getting queue client for queue: {_userAggregationQueueName}");
                var queueClient = _queueServiceClient.GetQueueClient(_userAggregationQueueName);
                
                if (queueClient == null)
                {
                    _logger.LogError("QueueClient is null after GetQueueClient call");
                    throw new InvalidOperationException("Failed to create QueueClient");
                }

                _logger.LogInformation("Creating queue if it doesn't exist...");
                await queueClient.CreateIfNotExistsAsync();

                int queuedCount = 0;
                foreach (var encryptedUPN in encryptedUPNs)
                {
                    if (string.IsNullOrEmpty(encryptedUPN))
                    {
                        _logger.LogWarning($"Skipping null or empty encryptedUPN at index {queuedCount}");
                        continue;
                    }

                    var message = new UserAggregationMessage
                    {
                        EncryptedUPN = encryptedUPN,
                        ReportRefreshDate = reportRefreshDate
                    };

                    var messageJson = JsonSerializer.Serialize(message);
                    
                    // Log first message for debugging
                    if (queuedCount == 0)
                    {
                        _logger.LogInformation("First message being queued: {messageJson}", messageJson);
                    }
                    
                    await queueClient.SendMessageAsync(messageJson);
                    queuedCount++;

                    // Log every 1000th message to track progress
                    if (queuedCount % 1000 == 0)
                    {
                        _logger.LogInformation($"Queued {queuedCount}/{encryptedUPNs.Count} messages...");
                    }
                }

                _logger.LogInformation($"Successfully queued {queuedCount} user aggregation messages");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error queuing user aggregation messages. Error details: {Message}, StackTrace: {StackTrace}", 
                    ex.Message, ex.StackTrace);
                throw;
            }
        }
    }
}
