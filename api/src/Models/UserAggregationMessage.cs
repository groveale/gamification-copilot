namespace groveale.Models
{
    /// <summary>
    /// Queue message for processing individual user aggregations
    /// </summary>
    public class UserAggregationMessage
    {
        /// <summary>
        /// Encrypted user principal name
        /// </summary>
        public required string EncryptedUPN { get; set; }

        /// <summary>
        /// Report refresh date in yyyy-MM-dd format
        /// </summary>
        public required string ReportRefreshDate { get; set; }
    }
}
