using System;
using System.Text.Json;
using groveale.Models;
using groveale.Services;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace groveale;

public class ProcessCopilotEventAggregation
{
    private readonly ILogger<ProcessCopilotEventAggregation> _logger;
    private readonly IAzureTableService _azureTableService;

    public ProcessCopilotEventAggregation(ILogger<ProcessCopilotEventAggregation> logger, IAzureTableService azureTableService)
    {
        _logger = logger;
        _azureTableService = azureTableService;
    }

    [Function(nameof(ProcessCopilotEventAggregation))]
    public async Task Run([QueueTrigger("%CopilotEventAggregationsQueueName%", Connection = "AzureWebJobsStorage")] string message)
    {
        _logger.LogInformation("Processing copilot event aggregation message");

        try
        {
            var aggMessage = JsonSerializer.Deserialize<CopilotEventAggregationMessage>(message);

            if (aggMessage == null)
            {
                _logger.LogError("Failed to deserialize copilot event aggregation message: {Message}", message);
                return;
            }

            var userId = aggMessage.EncryptedUPN;
            var eventDate = aggMessage.EventDate;

            _logger.LogInformation("Processing aggregation for user {UserId}, date {Date}, total events {Count}",
                userId, eventDate, aggMessage.TotalInteractionCount);

            // Write per-app aggregation rows
            foreach (var (appTypeName, count) in aggMessage.AppTypeCounts)
            {
                if (Enum.TryParse<AppType>(appTypeName, out var appType))
                {
                    await _azureTableService.AddSpecificCopilotInteractionDailyAggregationForUserAsync(
                        appType, userId, count, eventDate);
                }
                else
                {
                    _logger.LogWarning("Unknown AppType in aggregation message: {AppType}", appTypeName);
                }
            }

            // Write WebPlugin aggregation
            if (aggMessage.WebPluginCount > 0)
            {
                await _azureTableService.AddSpecificCopilotInteractionDailyAggregationForUserAsync(
                    AppType.WebPlugin, userId, aggMessage.WebPluginCount, eventDate);
            }

            // Write AppType.All total
            await _azureTableService.AddSpecificCopilotInteractionDailyAggregationForUserAsync(
                AppType.All, userId, aggMessage.TotalInteractionCount, eventDate);

            // Write agent aggregations
            foreach (var agent in aggMessage.AgentCounts)
            {
                await _azureTableService.AddAgentInteractionsDailyAggregationForUserAsync(
                    agent.AgentId, userId, agent.Count, agent.AgentName, eventDate);
            }

            // Update global daily totals
            await _azureTableService.UpdateDailyTotalInteractionsAsync(
                eventDate, aggMessage.TotalInteractionCount, 1);

            _logger.LogInformation("Successfully processed aggregation for user {UserId}: {Total} total, {AppTypes} app types, {Agents} agents",
                userId, aggMessage.TotalInteractionCount, aggMessage.AppTypeCounts.Count, aggMessage.AgentCounts.Count);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing copilot event aggregation message. Error: {Error}", ex.Message);
            throw; // Re-throw to trigger retry logic
        }
    }
}
