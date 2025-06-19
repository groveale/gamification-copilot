using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

public class M365CopilotUsage
{
    [JsonPropertyName("reportRefreshDate")]
    public string ReportRefreshDate { get; set; }

    [JsonPropertyName("userPrincipalName")]
    public string UserPrincipalName { get; set; }

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; }

    [JsonPropertyName("lastActivityDate")]
    public string LastActivityDate { get; set; }

    [JsonPropertyName("copilotChatLastActivityDate")]
    public string CopilotChatLastActivityDate { get; set; }

    [JsonPropertyName("microsoftTeamsCopilotLastActivityDate")]
    public string MicrosoftTeamsCopilotLastActivityDate { get; set; }

    [JsonPropertyName("wordCopilotLastActivityDate")]
    public string WordCopilotLastActivityDate { get; set; }

    [JsonPropertyName("excelCopilotLastActivityDate")]
    public string ExcelCopilotLastActivityDate { get; set; }

    [JsonPropertyName("powerPointCopilotLastActivityDate")]
    public string PowerPointCopilotLastActivityDate { get; set; }

    [JsonPropertyName("outlookCopilotLastActivityDate")]
    public string OutlookCopilotLastActivityDate { get; set; }

    [JsonPropertyName("oneNoteCopilotLastActivityDate")]
    public string OneNoteCopilotLastActivityDate { get; set; }

    [JsonPropertyName("loopCopilotLastActivityDate")]
    public string LoopCopilotLastActivityDate { get; set; }

    [JsonPropertyName("copilotActivityUserDetailsByPeriod")]
    public List<CopilotActivityUserDetail> CopilotActivityUserDetailsByPeriod { get; set; }
}

public class CopilotActivityUserDetail
{
    [JsonPropertyName("reportPeriod")]
    public int ReportPeriod { get; set; }
}