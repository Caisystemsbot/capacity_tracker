Attribute VB_Name = "modCapacityCore"
Option Explicit
'
' Split Wrapper Module (Core)
' This module provides thin wrappers for commonly-invoked Core macros so
' teams can start using a two‑module layout immediately. The full split
' plan and the list of procedures to migrate out of modCapacityPlanner is
' documented in docs/MODULE_SPLIT_PLAN.md. As we migrate, wrappers will be
' replaced by real implementations here.
'
' Public entrypoints (Core)
Public Sub Bootstrap()
    modCapacityPlanner.Bootstrap
End Sub

Public Sub HealthCheck()
    modCapacityPlanner.HealthCheck
End Sub

Public Sub Diagnostics_RunBootstrap()
    modCapacityPlanner.Diagnostics_RunBootstrap
End Sub

Public Sub CreateTeamAvailability()
    modCapacityPlanner.CreateTeamAvailability
End Sub

Public Sub CreateOrAdvanceAvailability()
    modCapacityPlanner.CreateOrAdvanceAvailability
End Sub

Public Sub BuildMetricsSkeleton()
    modCapacityPlanner.BuildMetricsSkeleton
End Sub

Public Sub DeleteOldConfigSheets()
    modCapacityPlanner.DeleteOldConfigSheets
End Sub

' NOTE:
' - As we complete the split, shared helpers (EnsureSheet, EnsureTable, logging,
'   settings/roster, etc.) will live here and be removed from modCapacityPlanner.
' - For now, wrappers keep compatibility with existing buttons/macros.

