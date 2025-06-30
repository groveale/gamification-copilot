# Copilot Gamification API

## Core Features

## Additional Features

* Skip offline users



## Skip Offline Users

To ensure fairness in streak tracking and inactivity scoring, the system implements logic to **skip users who are offline**.

### Activity Detection

The system uses the **Microsoft 365 activity report** to determine whether a user was active on the day Copilot data is processed. This is done via the API:

`GET https://graph.microsoft.com/beta/reports/getOffice365ActiveUserDetail(date=2025-06-27)`

The `date` parameter corresponds to the **report refresh date** returned by the Copilot activity API.

### Online Criteria

A user is considered **online** if they show activity in any of the following services on the report date:

- Microsoft Teams
- Outlook
- OneDrive
- SharePoint

If no activity is detected, the user is **excluded** from gamification data for that day. This ensures they do not lose streaks or accrue inactivity unfairly.

### Fallback Handling

The Copilot API may be ahead of the usage API. If the API returns no data for a given day, a fallback is triggered using:

`GET https://graph.microsoft.com/beta/reports/getOffice365ActiveUserDetail(period='D7')`


This allows retrieval of Copilot license information. If it's still not possible to determine whether a user is offline or inactive, the user is treated as **inactive** and will **lose streaks**.

### Implementation

This logic is implemented in the `GraphService.cs` file, specifically in the methods:

```csharp
GetM365CopilotUsersAsync()
GetM365CopilotUserFallBackAsync()