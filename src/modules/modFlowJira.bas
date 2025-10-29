Attribute VB_Name = "modFlowJira"
Option Explicit
'
' Split Wrapper Module (Flow + Jira)
' Thin wrappers that forward to existing implementations in
' modCapacityPlanner. This lets users start calling Flow/Jira entrypoints
' from a second module while we migrate the full implementations.
' See docs/MODULE_SPLIT_PLAN.md for the detailed migration plan.

' -------- Flow Metrics --------
Public Sub Flow_BuildCharts(Optional ByVal loSelected As ListObject)
    On Error GoTo Fallback
    modCapacityPlanner.Flow_BuildCharts loSelected
    Exit Sub
Fallback:
    modCapacityPlanner.Flow_BuildCharts
End Sub

' WIP CSV Sanitizer
Public Sub WIP_ImportCSV()
    modCapacityPlanner.WIP_ImportCSV
End Sub

' Orchestration
Public Sub SanitizeRawAndBuildInsights()
    modCapacityPlanner.SanitizeRawAndBuildInsights
End Sub

Public Sub RefreshSamples()
    modCapacityPlanner.RefreshSamples
End Sub

' -------- Jira Integration / Insights --------
Public Sub Jira_PopulateMetrics()
    modCapacityPlanner.Jira_PopulateMetrics
End Sub

Public Sub BuildJiraInsights()
    modCapacityPlanner.BuildJiraInsights
End Sub

' NOTE:
' - As we complete the split, Flow_* and Jira_* implementations (and helpers)
'   will move here from modCapacityPlanner.

