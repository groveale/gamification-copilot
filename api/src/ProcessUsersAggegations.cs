using System;
using System.Text.Json;
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
    public async Task Run([QueueTrigger("%UserAggregationsQueueName%", Connection = "AzureWebJobsStorage")] string message)
    {
        _logger.LogInformation("=== FUNCTION TRIGGERED ===");
        _logger.LogInformation("Message Text: {messageText}", message);
    
        try
        {
            _logger.LogInformation("Processing user aggregation queue message");

            // Deserialize the message
            var userAggregationMessage = JsonSerializer.Deserialize<UserAggregationMessage>(message);

            if (userAggregationMessage == null)
            {
                _logger.LogError("Failed to deserialize queue message: {messageText}", message);
                return;
            }

            _logger.LogInformation("Processing aggregation for UPN: {encryptedUPN}, Date: {date}", 
                userAggregationMessage.EncryptedUPN, userAggregationMessage.ReportRefreshDate);

            // Process the single user aggregation
            await _azureTableService.ProcessSingleUserAggregationAsync(
                userAggregationMessage.EncryptedUPN, 
                userAggregationMessage.ReportRefreshDate);

            _logger.LogInformation("Successfully processed user aggregation");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing user aggregation queue message. Error: {error}", ex.Message);
            throw; // Re-throw to trigger retry logic
        }
    }
}