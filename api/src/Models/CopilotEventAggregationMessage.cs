namespace groveale.Models
{
    public class CopilotEventAggregationMessage
    {
        public required string EncryptedUPN { get; set; }
        public required string EventDate { get; set; }
        public int TotalInteractionCount { get; set; }
        public int WebPluginCount { get; set; }

        /// <summary>
        /// Per-app interaction counts keyed by AppType name (e.g. "Word", "Teams", "CopilotChat")
        /// </summary>
        public Dictionary<string, int> AppTypeCounts { get; set; } = new();

        /// <summary>
        /// Agent interaction counts keyed by AgentId, with AgentName stored alongside
        /// </summary>
        public List<AgentAggregation> AgentCounts { get; set; } = new();
    }

    public class AgentAggregation
    {
        public required string AgentId { get; set; }
        public string? AgentName { get; set; }
        public int Count { get; set; }
    }
}
