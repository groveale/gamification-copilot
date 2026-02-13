namespace groveale.Services
{
    public interface ISettingsService
    {
        string TenantId { get; }
        string AuthGuid { get; }
        string StorageAccountUri { get; }
        string StorageAccountName { get; }
        string KeyVaultUrl { get; }
        string KeyVaultEncryptionKeySecretName { get; }
        string SPOSiteId { get; }
        string SPOFieldName { get; }
        string SPOListId { get; }
        string ReminderDays { get; }
        bool IsEmailListExclusive { get; }
        string UserAggregationsQueueName { get; }
        
    }

    public class SettingsService : ISettingsService
    {
        public string TenantId => Environment.GetEnvironmentVariable("TenantId");
        public string AuthGuid => Environment.GetEnvironmentVariable("AuthGuid");

        public string StorageAccountUri => Environment.GetEnvironmentVariable("StorageAccountUri");
        public string StorageAccountName => Environment.GetEnvironmentVariable("StorageAccountName");
        public string KeyVaultUrl => Environment.GetEnvironmentVariable("KeyVault:Url");
        public string KeyVaultEncryptionKeySecretName => Environment.GetEnvironmentVariable("KeyVault:EncryptionKeySecretName");
        public string SPOSiteId => Environment.GetEnvironmentVariable("SPO:SiteId");
        public string SPOFieldName => Environment.GetEnvironmentVariable("SPO:FieldName");
        public string SPOListId => Environment.GetEnvironmentVariable("SPO:ListId");
        public string ReminderDays => Environment.GetEnvironmentVariable("ReminderDays") ?? "14"; // Default to 14 days if not set

        // If true, the email list is exclusive (inclusion list). If false, it's an exclusion list.
        // Default to false (exclusion list) to maintain backward compatibility
        public bool IsEmailListExclusive => bool.TryParse(Environment.GetEnvironmentVariable("IsEmailListExclusive"), out var result) ? result : false;

        public string UserAggregationsQueueName => Environment.GetEnvironmentVariable("UserAggregationsQueueName") ?? "user-aggregations";

    }
}