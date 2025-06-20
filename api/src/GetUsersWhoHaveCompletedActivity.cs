using groveale.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace groveale
{
    public class GetUsersWhoHaveCompletedActivity
    {
        private readonly ILogger<GetUsersWhoHaveCompletedActivity> _logger;
        private readonly ISettingsService _settingsService;
        private readonly IAzureTableService _azureTableService;
        private readonly IKeyVaultService _keyVaultService;

        public GetUsersWhoHaveCompletedActivity(ILogger<GetUsersWhoHaveCompletedActivity> logger, ISettingsService settingsService, IAzureTableService azureTableService, IKeyVaultService keyVaultService)
        {
            _logger = logger;
            _settingsService = settingsService;
            _azureTableService = azureTableService;
            _keyVaultService = keyVaultService;
        }

        [Function("GetUsersWhoHaveCompletedActivity")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequest req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");

            // Get params, app, count, timeFrame, and startdate (optional)

            string dayCount = req.Query["dayCount"];
            string interactionCount = req.Query["interactionCount"];
            string timeFrame = req.Query["timeFrame"];
            string startDate = req.Query["startDate"];
            string demo = req.Query["demo"];

            // also check in the body
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            dayCount = dayCount ?? data?.dayCount;
            interactionCount = interactionCount ?? data?.interactionCount;
            timeFrame = timeFrame ?? data?.timeFrame;
            startDate = startDate ?? data?.startDate;
            demo = demo ?? data?.demo;


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
            if (appList == null || !appList.Any() || (string.IsNullOrEmpty(dayCount) && string.IsNullOrEmpty(interactionCount)) || string.IsNullOrEmpty(timeFrame))
            {
                return new BadRequestObjectResult("Please pass an app, dayCount and/or interactionCount, and timeFrame on the query string or body");
            }

            timeFrame = timeFrame.ToLower();

            // Further validation
            if (timeFrame != "alltime" && string.IsNullOrEmpty(startDate))
            {
                return new BadRequestObjectResult("Please pass a app, dayCount and/or interactionCount, timeFrame, and startDate on the query string or body");
            }


            // Get current start date form time frame
            var startDateForTimeFrame = await _azureTableService.GetStartDate(timeFrame);
            if (startDateForTimeFrame == null && timeFrame != "alltime" && demo != "true")
            {
                // if startDateForTimeFrame is null then we have no data yet for this time frame
                // return a 400 bad request with a message
                _logger.LogInformation("No data yet for this time frame - wait until tomorrow");
                return new BadRequestObjectResult("No data yet - wait until tomorrow");
            }

            // define an object to return
            var usersThatHaveAchieved = new List<string>();
            var startDateStatus = "Active";

            // if startDateForTimeFrame == startDate from input then we are good
            if (startDateForTimeFrame == startDate || timeFrame == "alltime" || demo == "true")
            {
                try
                {


                    if (appList.Count == 1)
                    {
                        // a simple one app challenge
                        var users = await _azureTableService.GetUsersWhoHaveCompletedActivityForApp(appList.FirstOrDefault(), dayCount, interactionCount, timeFrame, startDate);
                        usersThatHaveAchieved = users;
                    }
                    else
                    {
                        // we need to have multiple queries
                        // Most efficient approach using HashSet for O(1) lookups

                        HashSet<string> intersection = null;

                        foreach (var app in appList)
                        {
                            var usersForApp = await _azureTableService.GetUsersWhoHaveCompletedActivityForApp(app, dayCount, interactionCount, timeFrame, startDate);
                            var currentUsers = new HashSet<string>(usersForApp);

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
                    return new BadRequestObjectResult(ex.Message);
                }

            }
            else
            {
                // convert the strings into dates
                var startDateForTimeFrameDate = DateTime.Parse(startDateForTimeFrame);
                var startDateDate = DateTime.Parse(startDate);

                // work out if startDate from paramerter is before or after startDate from timeFrame, if before then return expired
                if (startDateDate < startDateForTimeFrameDate)
                {
                    startDateStatus = "Expired";
                }
                else
                {
                    startDateStatus = "Future";
                }

            }

            // Decrypt the users that have achieved
            if (usersThatHaveAchieved.Any())
            {
                // Create Encyption Service
                var encryptionService = await DeterministicEncryptionService.CreateAsync(_settingsService, _keyVaultService);
                usersThatHaveAchieved = usersThatHaveAchieved.Select(encryptionService.Decrypt).ToList();
            }

            return new OkObjectResult(new { Users = usersThatHaveAchieved, StartDateStatus = startDateStatus });
        }
    }
}
