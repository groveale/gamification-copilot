
using Microsoft.Extensions.Logging;

namespace groveale.Services
{
    public interface IUserActivitySeeder
    {
        Task SeedWeeklyActivitiesAsync(IDeterministicEncryptionService encryptionService, List<string> userList);
        Task SeedMonthlyActivitiesAsync(IDeterministicEncryptionService encryptionService, List<string> userList);
        Task SeedAllTimeActivityAsync(IDeterministicEncryptionService encryptionService, List<string> userList);
        Task SeedInactiveUsersAsync(IDeterministicEncryptionService encryptionService, List<string> userList);
    }

    public class UserActivitySeeder : IUserActivitySeeder
    {
        private readonly ILogger<UserActivitySeeder> _logger;
        private readonly IAzureTableService _azureTableService;



        public UserActivitySeeder(ILogger<UserActivitySeeder> logger, IAzureTableService azureTableService)
        {
            _logger = logger;
            _azureTableService = azureTableService;
        }


        public async Task SeedMonthlyActivitiesAsync(IDeterministicEncryptionService encryptionService, List<string> userList)
        {
            _logger.LogInformation("Seeding monthly activities...");

            // Get all users
            var random = new Random();
            var activities = new List<CopilotTimeFrameUsage>();

            var firstOfMonthForSnapshot = DateTime.UtcNow.Date
                .AddDays(-1 * DateTime.UtcNow.Date.Day + 1)
                .ToString("yyyy-MM-dd");

            foreach (var upn in userList)
            {
                foreach (var app in Enum.GetValues(typeof(AppType)).Cast<AppType>())
                {
                    // Create a simple activity entry with random boolean values
                    var activity = new CopilotTimeFrameUsage
                    {
                        UPN = encryptionService.Encrypt(upn),
                        App = app,
                        TotalDailyActivityCount = random.Next(30),
                        TotalInteractionCount = random.Next(1000),
                    };

                    // Best can only be as high as the total daily activity count
                    activity.BestDailyStreak = random.Next(0, activity.TotalDailyActivityCount + 1);

                    // Current can only be as high as best daily streak
                    activity.CurrentDailyStreak = random.Next(0, (activity.BestDailyStreak ?? 0) + 1);


                    // add
                    activities.Add(activity);
                }
            }

            // Save user activities
            await _azureTableService.SeedMonthlyTimeFrameActivitiesAsync(activities, firstOfMonthForSnapshot);

            _logger.LogInformation("Seeding monthly activities completed.");
        }

        

        public async Task SeedWeeklyActivitiesAsync(IDeterministicEncryptionService encryptionService, List<string> userList)
        {
            _logger.LogInformation("Seeding weekly activities ...");

            // Get all users
            var random = new Random();


            // Get the Monday of the current week
            var dayOfWeek = (int)DateTime.UtcNow.Date.DayOfWeek;
            //var daysToSubtract = dayOfWeek == 0 ? 6 : dayOfWeek - 1; // Adjust for Sunday
            var daysToSubtract = dayOfWeek; // Sunday is 0, so subtract directly
            var monday = DateTime.UtcNow.Date.AddDays(-1 * daysToSubtract);

            var activities = new List<CopilotTimeFrameUsage>();

            foreach (var upn in userList)
            {
                foreach (var app in Enum.GetValues(typeof(AppType)).Cast<AppType>())
                {
                    // Create a simple activity entry with random boolean values
                    var activity = new CopilotTimeFrameUsage
                    {
                        UPN = encryptionService.Encrypt(upn),
                        App = app,
                        TotalDailyActivityCount = random.Next(5),
                        TotalInteractionCount = random.Next(100),

                    };

                    // Best can only be as high as the total daily activity count
                    activity.BestDailyStreak = random.Next(0, activity.TotalDailyActivityCount + 1);

                    // Current can only be as high as best daily streak
                    activity.CurrentDailyStreak = random.Next(0, (activity.BestDailyStreak ?? 0) + 1);

                    // add
                    activities.Add(activity);
                }
            }

            // Save user activities
            await _azureTableService.SeedWeeklyTimeFrameActivitiesAsync(activities, monday.ToString("yyyy-MM-dd"));

            _logger.LogInformation("Seeding weekly activities completed.");
        }

        public async Task SeedAllTimeActivityAsync(IDeterministicEncryptionService encryptionService, List<string> userList)
        {
            _logger.LogInformation("Seeding all time activities ...");



            // Get all users
            var random = new Random();
            var activities = new List<CopilotTimeFrameUsage>();

            foreach (var upn in userList)
            {
                foreach (var app in Enum.GetValues(typeof(AppType)).Cast<AppType>())
                {
                    // Create a simple activity entry with random boolean values
                    var activity = new CopilotTimeFrameUsage
                    {
                        UPN = encryptionService.Encrypt(upn),
                        App = app,
                        TotalInteractionCount = random.Next(5000),
                        TotalDailyActivityCount = random.Next(100),
                        
                    };

                    // Best can only be as high as the total daily activity count
                    activity.BestDailyStreak = random.Next(0, activity.TotalDailyActivityCount + 1);

                    // Current can only be as high as best daily streak
                    activity.CurrentDailyStreak = random.Next(0, (activity.BestDailyStreak ?? 0) + 1);

                    // add
                    activities.Add(activity);
                }
            }

            // Save user activities
            await _azureTableService.SeedAllTimeActivityAsync(activities);

            _logger.LogInformation("Seeding all time completed.");
        }

        public async Task SeedInactiveUsersAsync(IDeterministicEncryptionService encryptionService, List<string> userList)
        {
            _logger.LogInformation("Seeding inactive users ...");

            // Get all users
            var random = new Random();
            var activities = new List<InactiveUser>();

            int[] possibleDays = { 7, 14, 30, 60, 90 };

            // Today
            var today = DateTime.UtcNow.Date;

            foreach (var upn in userList)
            {
                // Create a simple activity entry with random boolean values
                var activity = new InactiveUser
                {
                    UPN = encryptionService.Encrypt(upn),
                    DaysSinceLastActivity = possibleDays[random.Next(0, possibleDays.Length)],
                };

                // Calculate LastActivityDate based on DaysSinceLastActivity
                activity.LastActivityDate = today.AddDays(-activity.DaysSinceLastActivity);

                activities.Add(activity);
            }

            // Save user activities
            await _azureTableService.SeedInactiveUsersAsync(activities);

            _logger.LogInformation("Seeding inactive users completed.");
        }
    }
}
