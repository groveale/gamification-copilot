using groveale.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace groveale
{
    public class GetUsersWithStreak
    {
        private readonly ILogger<GetUsersWithStreak> _logger;
        private readonly IAzureTableService _azureTableService;
        private readonly ISettingsService _settingsService;
        private readonly IKeyVaultService _keyVaultService;

        public GetUsersWithStreak(ILogger<GetUsersWithStreak> logger, IAzureTableService azureTableService, ISettingsService settingsService, IKeyVaultService keyVaultService)
        {
            _logger = logger;
            _azureTableService = azureTableService;
            _settingsService = settingsService;
            _keyVaultService = keyVaultService;
        }

        [Function("GetUsersWithStreak")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Admin, "get", "post")] HttpRequest req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");

            // Parse query parameter
            string count = req.Query["count"];

            // Parse request body
            string requestBody = new StreamReader(req.Body).ReadToEndAsync().Result;
            dynamic data = Newtonsoft.Json.JsonConvert.DeserializeObject(requestBody);
            count = count ?? data?.count;

            // Handle app parameter as an array or a comma-separated string
            List<string> appList;
            if (data?.apps is Newtonsoft.Json.Linq.JArray)
            {
                appList = data.apps.ToObject<List<string>>();
            }
            else
            {
                string apps = req.Query["apps"];
                appList = apps?.Split(',').Select(a => a.Trim()).ToList();
            }

            // Validate params
            if (appList == null || !appList.Any() || string.IsNullOrEmpty(count))
            {
                return new BadRequestObjectResult("Please pass an 'apps' list and count parameter on the query string or body");
            }

            if (!int.TryParse(count, out int countValue))
            {
                return new BadRequestObjectResult("Please pass a valid integer for the count parameter");
            }

            List<string> usersThatHaveAchieved = new List<string>();

            try
            {
                // Get users with streak
                if (appList.Count == 1)
                {
                    var usersWithAppStreak = await _azureTableService.GetUsersWithStreakForApp(appList.FirstOrDefault(), countValue);
                    usersThatHaveAchieved = usersWithAppStreak;
                }
                else
                {

                    HashSet<string> intersection = null;

                    foreach (var app in appList)
                    {
                        var usersWithAppStreak = await _azureTableService.GetUsersWithStreakForApp(app, countValue);
                        var currentUsers = new HashSet<string>(usersWithAppStreak);

                        // Early exit: if any app has no qualifying users, final result will be empty
                        if (!currentUsers.Any())
                        {
                            usersThatHaveAchieved = new List<string>();
                            break;
                        }

                        if (intersection == null)
                        {
                            // First app - initialize intersection
                            intersection = currentUsers;
                        }
                        else
                        {
                            // Subsequent apps - intersect with existing set
                            intersection.IntersectWith(currentUsers);

                            // Early exit: if intersection becomes empty, no point continuing
                            if (!intersection.Any())
                            {
                                usersThatHaveAchieved = new List<string>();
                                break;
                            }
                        }
                    }

                    usersThatHaveAchieved = intersection?.ToList() ?? new List<string>();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An error occurred while processing the request.");
                return new StatusCodeResult(StatusCodes.Status500InternalServerError);
            }

            // Decrypt the users that have achieved
            if (usersThatHaveAchieved.Any())
            {
                // Create Encyption Service
                var encryptionService = await DeterministicEncryptionService.CreateAsync(_settingsService, _keyVaultService);
                usersThatHaveAchieved = usersThatHaveAchieved.Select(encryptionService.Decrypt).ToList();
            }

            return new OkObjectResult(usersThatHaveAchieved);

        }
    }
}
