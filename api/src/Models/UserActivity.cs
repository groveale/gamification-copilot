using Azure.Data.Tables;
using groveale.Models;

public class UserActivity : BaseTableEntity
{
    public DateTime ReportDate { get; set; }
    public string UPN { get; set; }
    public string DisplayName { get; set; }

    public bool DailyTeamsActivity { get; set; }
    public int DailyTeamsInteractionCount { get; set; }

    public bool DailyCopilotChatActivity { get; set; }
    public int DailyCopilotChatInteractionCount { get; set; }

    public bool DailyOutlookActivity { get; set; }
    public int DailyOutlookInteractionCount { get; set; }

    public bool DailyWordActivity { get; set; }
    public int DailyWordInteractionCount { get; set; }

    public bool DailyExcelActivity { get; set; }
    public int DailyExcelInteractionCount { get; set; }

    public bool DailyPowerPointActivity { get; set; }
    public int DailyPowerPointInteractionCount { get; set; }

    public bool DailyOneNoteActivity { get; set; }
    public int DailyOneNoteInteractionCount { get; set; }

    public bool DailyLoopActivity { get; set; }
    public int DailyLoopInteractionCount { get; set; }

    public bool DailyAllUpActivity { get; set; }
    public int DailyAllInteractionCount { get; set; }

    // Additional interaction counts with corresponding bools
    public bool DailyMACActivity { get; set; }
    public int DailyMACInteractionCount { get; set; }

    public bool DailyDesignerActivity { get; set; }
    public int DailyDesignerInteractionCount { get; set; }

    public bool DailySharePointActivity { get; set; }
    public int DailySharePointInteractionCount { get; set; }

    public bool DailyPlannerActivity { get; set; }
    public int DailyPlannerInteractionCount { get; set; }

    public bool DailyWhiteboardActivity { get; set; }
    public int DailyWhiteboardInteractionCount { get; set; }

    public bool DailyStreamActivity { get; set; }
    public int DailyStreamInteractionCount { get; set; }

    public bool DailyFormsActivity { get; set; }
    public int DailyFormsInteractionCount { get; set; }

    public bool DailyCopilotActionActivity { get; set; }
    public int DailyCopilotActionCount { get; set; }

    public bool DailyWebPluginActivity { get; set; }
    public int DailyWebPluginInteractions { get; set; }
    public bool DailyAgentActivity { get; set; }
    public int DailyAgentInteractions { get; set; }

    // Converts the UserActivity object to a TableEntity
    public TableEntity ToTableEntity()
    {
        PartitionKey = ReportDate.ToString("yyyy-MM-dd");;
        RowKey = UPN;

        return new TableEntity(PartitionKey, RowKey)
        {
            { nameof(ReportDate), DateTime.SpecifyKind (ReportDate, DateTimeKind.Utc) },
            { nameof(DisplayName), DisplayName },
            { nameof(DailyTeamsActivity), DailyTeamsActivity },
            { nameof(DailyTeamsInteractionCount), DailyTeamsInteractionCount },
            { nameof(DailyCopilotChatActivity), DailyCopilotChatActivity },
            { nameof(DailyCopilotChatInteractionCount), DailyCopilotChatInteractionCount },
            { nameof(DailyOutlookActivity), DailyOutlookActivity },
            { nameof(DailyOutlookInteractionCount), DailyOutlookInteractionCount },
            { nameof(DailyWordActivity), DailyWordActivity },
            { nameof(DailyWordInteractionCount), DailyWordInteractionCount },
            { nameof(DailyExcelActivity), DailyExcelActivity },
            { nameof(DailyExcelInteractionCount), DailyExcelInteractionCount },
            { nameof(DailyPowerPointActivity), DailyPowerPointActivity },
            { nameof(DailyPowerPointInteractionCount), DailyPowerPointInteractionCount },
            { nameof(DailyOneNoteActivity), DailyOneNoteActivity },
            { nameof(DailyOneNoteInteractionCount), DailyOneNoteInteractionCount },
            { nameof(DailyLoopActivity), DailyLoopActivity },
            { nameof(DailyLoopInteractionCount), DailyLoopInteractionCount },
            { nameof(DailyAllUpActivity), DailyAllUpActivity },
            { nameof(DailyAllInteractionCount), DailyAllInteractionCount },
            { nameof(DailyMACActivity), DailyMACActivity },
            { nameof(DailyMACInteractionCount), DailyMACInteractionCount },
            { nameof(DailyDesignerActivity), DailyDesignerActivity },
            { nameof(DailyDesignerInteractionCount), DailyDesignerInteractionCount },
            { nameof(DailySharePointActivity), DailySharePointActivity },
            { nameof(DailySharePointInteractionCount), DailySharePointInteractionCount },
            { nameof(DailyPlannerActivity), DailyPlannerActivity },
            { nameof(DailyPlannerInteractionCount), DailyPlannerInteractionCount },
            { nameof(DailyWhiteboardActivity), DailyWhiteboardActivity },
            { nameof(DailyWhiteboardInteractionCount), DailyWhiteboardInteractionCount },
            { nameof(DailyStreamActivity), DailyStreamActivity },
            { nameof(DailyStreamInteractionCount), DailyStreamInteractionCount },
            { nameof(DailyFormsActivity), DailyFormsActivity },
            { nameof(DailyFormsInteractionCount), DailyFormsInteractionCount },
            { nameof(DailyCopilotActionActivity), DailyCopilotActionActivity },
            { nameof(DailyCopilotActionCount), DailyCopilotActionCount },
            { nameof(DailyWebPluginActivity), DailyWebPluginActivity },
            { nameof(DailyWebPluginInteractions), DailyWebPluginInteractions },
            { nameof(DailyAgentActivity), DailyAgentActivity },
            { nameof(DailyAgentInteractions), DailyAgentInteractions }
        };
    }
}