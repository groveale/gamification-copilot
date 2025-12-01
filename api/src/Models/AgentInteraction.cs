using Azure.Data.Tables;
using groveale.Models;

public class AgentInteraction : BaseTableEntity
{
    public const string AllTimePartitionKeyPrefix = "allTime";
    public string? UPN { get; set; }
    public string? AgentId { get; set; }
    public string? AgentName { get; set; }
    public int TotalDailyActivityCount { get; set; }
    public int TotalInteractionCount { get; set; }

    public TableEntity ToDailyTableEntity(string stringDate)
    {
        PartitionKey = $"{stringDate}-{UPN}";
        RowKey = AgentId;

        return new TableEntity(PartitionKey, RowKey)
        {
            { nameof(TotalDailyActivityCount), 1 },
            { nameof(TotalInteractionCount), TotalInteractionCount },
            { nameof(AgentName), AgentName }
        };
    }

    public TableEntity ToTimeFrameTableEntity(string stringDate)
        {
            PartitionKey = $"{stringDate}-{AgentId}";
            RowKey = UPN;

            return new TableEntity(PartitionKey, RowKey)
            {
                { nameof(TotalDailyActivityCount), 1 },
                { nameof(TotalInteractionCount), TotalInteractionCount },
                { nameof(AgentName), AgentName }
            };
        }

        public TableEntity ToAllTimeTableEntity()
        {
            PartitionKey = $"{AllTimePartitionKeyPrefix}-{AgentId}";
            RowKey = UPN;

            return new TableEntity(PartitionKey, RowKey)
            {
                { nameof(TotalDailyActivityCount), TotalDailyActivityCount },
                { nameof(TotalInteractionCount), TotalInteractionCount },
                { nameof(AgentName), AgentName }
            };
        }

        // No user info
        public TableEntity ToAgentOnlyTableEntity(string stringDate)
        {
            PartitionKey = stringDate;
            RowKey = AgentId;

            return new TableEntity(PartitionKey, RowKey)
            {
                { nameof(TotalDailyActivityCount), TotalDailyActivityCount },
                { nameof(TotalInteractionCount), TotalInteractionCount },
                { nameof(AgentName), AgentName }
            };
        }
}