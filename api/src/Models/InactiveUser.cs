public class InactiveUser
{
    public string UPN { get; set; }
    public double DaysSinceLastActivity { get; set; }
    public DateTime LastActivityDate { get; set; }
    public string DisplayName { get; set; }
}