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
            _queueServiceClient = new QueueServiceClient(
                new Uri(settingsService.StorageAccountUri),
                new DefaultAzureCredential());

            _logger = logger;
            _userAggregationQueueName = settingsService.UserAggregationsQueueName;
        }

        public async Task QueueUserAggregationsAsync(List<string> encryptedUPNs, string reportRefreshDate)
        {
            try
            {
                _logger.LogInformation($"Queuing {encryptedUPNs.Count} user aggregation messages for date: {reportRefreshDate}");

                // Get or create the queue
                var queueClient = _queueServiceClient.GetQueueClient(_userAggregationQueueName);
                await queueClient.CreateIfNotExistsAsync();

                int queuedCount = 0;
                foreach (var encryptedUPN in encryptedUPNs)
                {
                    var message = new UserAggregationMessage
                    {
                        EncryptedUPN = encryptedUPN,
                        ReportRefreshDate = reportRefreshDate
                    };

                    var messageJson = JsonSerializer.Serialize(message);
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
                _logger.LogError(ex, "Error queuing user aggregation messages");
                throw;
            }
        }
    }
}
