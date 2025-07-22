using groveale.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace groveale
{
    public class SeedTestData
    {
        private readonly ILogger<SeedTestData> _logger;
        private readonly IUserActivitySeeder _userActivitySeeder;
        private readonly ISettingsService _settingsService;
        private readonly IKeyVaultService _keyVaultService;

        public SeedTestData(ILogger<SeedTestData> logger, IUserActivitySeeder userActivitySeeder, ISettingsService settingsService, IKeyVaultService keyVaultService)
        {
            _settingsService = settingsService;
            _keyVaultService = keyVaultService;
            _logger = logger;
            _userActivitySeeder = userActivitySeeder;
        }

        [Function("SeedTestData")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Admin, "get", "post")] HttpRequest req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");


            // also check in the body
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            
            List<string> userList = null;
            if (data?.users is Newtonsoft.Json.Linq.JArray)
            {
                userList = data.users.ToObject<List<string>>();
            }
            else
            {
                string users = req.Query["users"];
                userList = users?.Split(',').Select(a => a.Trim()).ToList();
            }

            // Validate params
            if (userList == null || !userList.Any())
            {
                return new BadRequestObjectResult("Please pass a 'users' list on the query string or body. It should be a comma-separated list (or array) of UPNs");
            }
        

            try
                {
                    // Create Encyption Service
                    var encryptionService = await DeterministicEncryptionService.CreateAsync(_settingsService, _keyVaultService);

                    // Seed daily activities for a tenant
                    await _userActivitySeeder.SeedWeeklyActivitiesAsync(encryptionService, userList);
                    await _userActivitySeeder.SeedMonthlyActivitiesAsync(encryptionService, userList);
                    await _userActivitySeeder.SeedAllTimeActivityAsync(encryptionService, userList);
                    await _userActivitySeeder.SeedInactiveUsersAsync(encryptionService, userList);

                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error processing usage data. {Message}", ex.Message);
                    _logger.LogError(ex, "Error processing usage data. {StackTrace}", ex.StackTrace);

                    return new BadRequestObjectResult("Error seeding activities");
                }

            return new OkObjectResult($"Data seeded successfully for users: {string.Join(", ", userList)}");
        }
    }
}
