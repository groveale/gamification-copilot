using groveale.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace groveale
{
    public class GetInactiveUsers
    {
        private readonly ILogger<GetInactiveUsers> _logger;
        private readonly IAzureTableService _azureTableService;
        private readonly ISettingsService _settingsService;
        private readonly IKeyVaultService _keyVaultService;

        public GetInactiveUsers(ILogger<GetInactiveUsers> logger, IAzureTableService azureTableService, ISettingsService settingsService, IKeyVaultService keyVaultService)
        {
            _settingsService = settingsService;
            _keyVaultService = keyVaultService;
            _azureTableService = azureTableService;
            _logger = logger;
        }

        [Function("GetInactiveUsers")]
        public async Task <IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequest req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");

            // Get params, app, count, timeFrame, and startdate (optional)
            string daysString = req.Query["days"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            daysString = daysString ?? data?.days;

            // Validate params
            if (string.IsNullOrEmpty(daysString) || !int.TryParse(daysString, out int days))
            {
                return new BadRequestObjectResult("Please pass a days parameter on the query string or body");
            }

            var users = await _azureTableService.GetInactiveUsers(days);

            // Decrypt users
            if (users.Any())
            {
                // Create Encyption Service
                var encryptionService = await DeterministicEncryptionService.CreateAsync(_settingsService, _keyVaultService);
            
                foreach (var user in users)
                {
                    user.UPN = encryptionService.Decrypt(user.UPN);
                }
            }

            // Process the request and return the result
            return new OkObjectResult(users);

        }
    }
}
