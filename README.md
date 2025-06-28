# WorkProgramManager
A C# program I use alongside BarRaider's "Advanced Launcher" StreamDeck plugin to give me some sense of "at work" and "at home" while working from home full time.

Example work-programs.json
```json
{
  "logDirectory": "C:\\Logs\\work-programs",
  "programs": [
    {
      "name": "Outlook",
      "type": "uwp",
      "aumid": "Microsoft.OutlookForWindows_8wekyb3d8bbwe!Microsoft.OutlookforWindows",
      "processName": "olk"
    },
    {
      "name": "Teams",
      "type": "uwp",
      "aumid": "MSTeams_8wekyb3d8bbwe!MSTeams",
      "processName": "ms-teams"
    },
    {
      "name": "OneNote",
      "type": "uwp",
      "aumid": "Microsoft.Office.ONENOTE.EXE.15",
      "processName": "onenote"
    },
    {
      "name": "Send to OneNote Tool",
      "type": "exe",
      "path": "C:\\Program Files\\Microsoft Office\\root\\Office16\\ONENOTEM.EXE",
      "processName": "onenotem"
    }
  ]
}
```