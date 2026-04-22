using groveale.Models;
using Microsoft.Extensions.Logging;

namespace groveale.Services
{
    public static class EventAggregationHelper
    {
        /// <summary>
        /// Maps a CopilotEventData to its AppType. Returns null for unhandled AppHost values.
        /// This mirrors the logic in AddSingleCopilotInteractionDailyAggregationForUserAsync.
        /// </summary>
        public static AppType? ResolveAppType(CopilotEventData entity)
        {
            return entity.AppHost switch
            {
                "Word" => AppType.Word,
                "Excel" => AppType.Excel,
                "PowerPoint" => AppType.PowerPoint,
                "OneNote" => AppType.OneNote,
                "Outlook" => AppType.Outlook,
                "OutlookSidepane" => AppType.Outlook,
                "Loop" => AppType.Loop,
                "Whiteboard" => AppType.Whiteboard,
                "Teams" => entity.Contexts != null && entity.Contexts.Any(c => c.Type.StartsWith("Teams"))
                    ? AppType.Teams
                    : AppType.CopilotChat,
                "Microsoft Teams" => AppType.Teams,
                "Office" => entity.AgentId != null ? AppType.Agent : AppType.CopilotChat,
                "Edge" => AppType.CopilotChat,
                "Designer" => AppType.Designer,
                "SharePoint" => AppType.SharePoint,
                "M365AdminCenter" => AppType.MAC,
                "OAIAutomationAgent" => AppType.CopilotAction,
                "Copilot Studio" => AppType.CopilotStudio,
                "Forms" => AppType.Forms,
                _ => null
            };
        }

        /// <summary>
        /// Builds pre-aggregated queue messages from copilot audit records, one message per user per event date.
        /// Single pass over events — collects per-app counts, web plugin counts, and agent counts.
        /// </summary>
        public static List<CopilotEventAggregationMessage> BuildAggregationMessages(
            List<AuditData> copilotAuditRecords,
            ILogger logger)
        {
            var messages = new List<CopilotEventAggregationMessage>();

            // Group by user, then by event date
            var groupedByUser = copilotAuditRecords
                .GroupBy(r => r.UserId);

            foreach (var userGroup in groupedByUser)
            {
                var userId = userGroup.Key;

                // Further group by event date so we get one message per user per date
                var groupedByDate = userGroup
                    .GroupBy(r => r.CopilotEventData.EventDateString ?? DateTime.SpecifyKind(r.CreationTime, DateTimeKind.Utc).ToString("yyyy-MM-dd"));

                foreach (var dateGroup in groupedByDate)
                {
                    var eventDate = dateGroup.Key;
                    var events = dateGroup.Select(r => r.CopilotEventData).ToList();

                    var appTypeCounts = new Dictionary<string, int>();
                    var agentMap = new Dictionary<string, (string? Name, int Count)>();
                    int webPluginCount = 0;
                    var unhandledHosts = new HashSet<string>();

                    // Single pass over all events for this user+date
                    foreach (var evt in events)
                    {
                        // Resolve app type
                        var appType = ResolveAppType(evt);
                        if (appType.HasValue)
                        {
                            var key = appType.Value.ToString();
                            appTypeCounts[key] = appTypeCounts.GetValueOrDefault(key, 0) + 1;
                        }
                        else if (!string.IsNullOrEmpty(evt.AppHost))
                        {
                            unhandledHosts.Add(evt.AppHost);
                        }

                        // Check for web plugin
                        if (evt.AISystemPlugin != null && evt.AISystemPlugin.Any(p => p.Name == "BingWebSearch"))
                        {
                            webPluginCount++;
                        }

                        // Check for agent
                        if (!string.IsNullOrEmpty(evt.AgentId))
                        {
                            if (agentMap.TryGetValue(evt.AgentId, out var existing))
                            {
                                agentMap[evt.AgentId] = (existing.Name ?? evt.AgentName, existing.Count + 1);
                            }
                            else
                            {
                                agentMap[evt.AgentId] = (evt.AgentName, 1);
                            }
                        }
                    }

                    foreach (var host in unhandledHosts)
                    {
                        logger.LogWarning("Unhandled AppHost during pre-aggregation: {AppHost}", host);
                    }

                    var message = new CopilotEventAggregationMessage
                    {
                        EncryptedUPN = userId,
                        EventDate = eventDate,
                        TotalInteractionCount = events.Count,
                        WebPluginCount = webPluginCount,
                        AppTypeCounts = appTypeCounts,
                        AgentCounts = agentMap.Select(kvp => new AgentAggregation
                        {
                            AgentId = kvp.Key,
                            AgentName = kvp.Value.Name,
                            Count = kvp.Value.Count
                        }).ToList()
                    };

                    messages.Add(message);
                }
            }

            return messages;
        }
    }
}
