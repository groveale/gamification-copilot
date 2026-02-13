using System;
using System.Text.Json;
using Azure.Storage.Queues.Models;
using groveale.Models;
using groveale.Services;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace groveale;

public class ProcessUsersAggegations
{
    private readonly ILogger<ProcessUsersAggegations> _logger;
    private readonly IAzureTableService _azureTableService;

    public ProcessUsersAggegations(ILogger<ProcessUsersAggegations> logger, IAzureTableService azureTableService)
    {
        _logger = logger;
        _azureTableService = azureTableService;
    }

    [Function(nameof(ProcessUsersAggegations))]
    public async Task Run([QueueTrigger("%UserAggregationsQueueName%", Connection = "AzureWebJobsStorage")] QueueMessage message)
    {
        try
        {
            _logger.LogInformation("Processing user aggregation queue message: {messageId}", message.MessageId);

            // Deserialize the message
            var userAggregationMessage = JsonSerializer.Deserialize<UserAggregationMessage>(message.MessageText);

            if (userAggregationMessage == null)
            {
                _logger.LogError("Failed to deserialize queue message: {messageText}", message.MessageText);
                return;
            }

            _logger.LogInformation("Processing aggregation for UPN: {encryptedUPN}, Date: {date}", 
                userAggregationMessage.EncryptedUPN, userAggregationMessage.ReportRefreshDate);

            // Process the single user aggregation
            await _azureTableService.ProcessSingleUserAggregationAsync(
                userAggregationMessage.EncryptedUPN, 
                userAggregationMessage.ReportRefreshDate);

            _logger.LogInformation("Successfully processed user aggregation for message: {messageId}", message.MessageId);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing user aggregation queue message: {messageId}. Error: {error}", 
                message.MessageId, ex.Message);
            throw; // Re-throw to trigger retry logic
        }
    }
}