Option Explicit

' Excel constants (avoid references)
Private Const xlSrcRange As Long = 1
Private Const xlYes As Long = 1
Private Const msoFileDialogFilePicker As Long = 3
Private Const xlValidateList As Long = 3
Private Const xlValidAlertStop As Long = 1
Private Const xlBetween As Long = 1
Private Const xlValidateDecimal As Long = 2

Public Sub Bootstrap()
    On Error GoTo Fail
    Application.ScreenUpdating = False

    EnsureSheets
    EnsureTables
    SeedNamedValues
    SeedSamplesIfPresent

    Application.ScreenUpdating = True
    MsgBox "Bootstrap complete.", vbInformation
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Bootstrap failed: " & Err.Description, vbExclamation
End Sub

 ' PTO import is deferred; Calendars sheet removed for now.

Public Sub HealthCheck()
    Dim issues As String
    If Not SheetExists("Config") Then issues = issues & "- Missing sheet: Config" & vbCrLf
    If Not SheetExists("Logs") Then issues = issues & "- Missing sheet: Logs" & vbCrLf
    If Len(issues) = 0 Then
        MsgBox "Health check OK.", vbInformation
    Else
        MsgBox "Health check issues:" & vbCrLf & issues, vbExclamation
    End If
End Sub

 ' Calendars sheet removed in this minimal profile

' -------------------- internals --------------------

Private Sub EnsureSheets()
    RemoveSheetIfExists "Calendars"
    Dim cfg As Worksheet: Set cfg = EnsureConfig()
    EnsureSheet "Getting_Started"
    EnsureSheet "Dashboard"
    EnsureSheet "Logs"
End Sub

Private Sub EnsureTables()
    Dim ws As Worksheet
    Set ws = EnsureConfig()
    EnsureRosterTable ws

    Set ws = EnsureSheet("Logs")
    EnsureTable ws, "tblLogs", Array("Timestamp", "User", "Action", "Outcome", "Details")
End Sub

Private Sub SeedNamedValues()
    Dim ws As Worksheet: Set ws = EnsureConfig()
    EnsureNamedValue "ActiveTeam", ws.Range("H2"), "CraicForce"
    EnsureNamedValue "TemplateVersion", ws.Range("H3"), "0.1.0"
    EnsureNamedValue "SprintLengthDays", ws.Range("H4"), 10
    EnsureNamedValue "DefaultHoursPerDay", ws.Range("H5"), 6.5
    EnsureNamedValue "DefaultAllocationPct", ws.Range("H6"), 1
    EnsureNamedValue "DefaultHoursPerPoint", ws.Range("H7"), 6
    EnsureNamedValue "RolesWithVelocity", ws.Range("H8"), "Developer,QA"
    WriteSettingsLabels ws
    WriteGettingStarted
    EnsureDashboard
End Sub

Private Sub SeedSamplesIfPresent()
    Dim base As String
    base = ThisWorkbook.Path
    If Len(base) = 0 Then Exit Sub

    Dim roster As String: roster = PathJoin(base, "data\roster_example.csv")
    Dim lo As ListObject
    If FileExists(roster) Then
        Set lo = EnsureRosterTable(EnsureConfig())
        If lo.ListRows.Count = 0 Then ImportCsvToTable roster, lo, True
    End If
End Sub

Private Function EnsureSheet(ByVal name As String) As Worksheet
    If SheetExists(name) Then
        Set EnsureSheet = Worksheets(name)
    Else
        Set EnsureSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        On Error Resume Next: EnsureSheet.Name = name: On Error GoTo 0
    End If
End Function

Private Function SheetExists(ByVal name As String) As Boolean
    On Error Resume Next
    SheetExists = Not Worksheets(name) Is Nothing
    On Error GoTo 0
End Function

 ' Overlap-safe table creation: finds a free row and retries if needed
Private Function EnsureTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal headers As Variant) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If Not lo Is Nothing Then
        Set EnsureTable = lo
        Exit Function
    End If

    Dim startRow As Long: startRow = NextFreeRow(ws)
    Dim lastCol As Long: lastCol = UBound(headers) - LBound(headers) + 1
    Dim c As Long, attempts As Long

RetryPlacement:
    attempts = attempts + 1
    If attempts > 25 Then Err.Raise 1004, , "Could not find free space for table: " & tableName

    ' Write header row at candidate position
    For c = LBound(headers) To UBound(headers)
        ws.Cells(startRow, 1 + c - LBound(headers)).Value = headers(c)
    Next c

    On Error GoTo Overlap
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, lastCol)), , xlYes)
    On Error GoTo 0
    lo.Name = tableName
    Set EnsureTable = lo
    Exit Function

Overlap:
    ' Clear the header cells we just wrote (to avoid debris), shift down, and retry
    On Error Resume Next
    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, lastCol)).ClearContents
    On Error GoTo 0
    startRow = startRow + 5
    GoTo RetryPlacement
End Function

' Compute the first safe row below all existing tables and used cells
Private Function NextFreeRow(ByVal ws As Worksheet) As Long
    Dim r As Long: r = 1
    Dim lo As ListObject, bottom As Long: bottom = 0
    For Each lo In ws.ListObjects
        Dim b As Long
        b = lo.Range.Row + lo.Range.Rows.Count - 1
        If b > bottom Then bottom = b
    Next lo
    If bottom > 0 Then r = bottom + 2

    Dim usedBottom As Long
    If Not ws.UsedRange Is Nothing Then
        usedBottom = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        If usedBottom >= r Then r = usedBottom + 2
    End If
    If r < 1 Then r = 1
    NextFreeRow = r
End Function

Private Sub EnsureNamedValue(ByVal nm As String, ByVal target As Range, ByVal defaultValue As Variant)
    On Error Resume Next
    Dim n As Name: Set n = ThisWorkbook.Names(nm)
    On Error GoTo 0
    If n Is Nothing Then
        target.Value = defaultValue
        ThisWorkbook.Names.Add nm, target
    ElseIf Len(CStr(n.RefersToRange.Value)) = 0 Then
        n.RefersToRange.Value = defaultValue
    End If
End Sub

Private Sub WriteSettingsLabels(ByVal ws As Worksheet)
    ws.Range("G2").Value = "ActiveTeam"
    ws.Range("G3").Value = "TemplateVersion"
    ws.Range("G4").Value = "SprintLengthDays"
    ws.Range("G5").Value = "DefaultHoursPerDay"
    ws.Range("G6").Value = "DefaultAllocationPct (optional)"
    ws.Range("G7").Value = "DefaultHoursPerPoint"
    ws.Range("G8").Value = "RolesWithVelocity (comma list)"
    ws.Columns("G:H").AutoFit
End Sub

Private Sub RemoveSheetIfExists(ByVal name As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(name).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Sub WriteGettingStarted()
    Dim ws As Worksheet: Set ws = EnsureSheet("Getting_Started")
    ws.Cells.Clear
    ws.Range("A1").Value = "Getting Started"
    ws.Range("A2").Value = "1) Populate your team roster on sheet 'Config' table 'tblRoster'."
    ws.Range("A3").Value = "   Columns: Member, Role, ContributesToVelocity (Yes/No)"
    ws.Range("A4").Value = "2) Only roles listed in named cell 'RolesWithVelocity' contribute to velocity (default: Developer, QA)."
    ws.Range("A5").Value = "3) Settings live on 'Config' H2:H8 (ActiveTeam, SprintLengthDays, etc.)."
    ws.Range("A6").Value = "4) Run 'HealthCheck' anytime to validate required pieces."
End Sub

Private Sub ImportCsvToTable(ByVal filePath As String, ByVal lo As ListObject, ByVal hasHeader As Boolean)
    If Not FileExists(filePath) Then Err.Raise 5, , "CSV not found: " & filePath
    Dim fso As Object, ts As Object, line As String, parts As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1)
    Dim rowIdx As Long: rowIdx = 0
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        rowIdx = rowIdx + 1
        If hasHeader And rowIdx = 1 Then GoTo ContinueLoop
        If Len(Trim$(line)) = 0 Then GoTo ContinueLoop
        parts = Split(line, ",")
        Dim lr As ListRow: Set lr = lo.ListRows.Add
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            If i + 1 <= lo.ListColumns.Count Then lr.Range(1, i + 1).Value = Trim$(parts(i))
        Next i
ContinueLoop:
    Loop
    ts.Close
End Sub

' Build or repair roster table with new schema and data validation
Private Function EnsureRosterTable(ByVal ws As Worksheet) As ListObject
    Dim expected As Variant
    expected = Array("Member", "Role", "ContributesToVelocity")

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects("tblRoster")
    On Error GoTo 0

    Dim rebuild As Boolean: rebuild = False
    If Not lo Is Nothing Then
        ' Check header names
        Dim ok As Boolean: ok = True
        Dim i As Long
        If lo.ListColumns.Count <> (UBound(expected) - LBound(expected) + 1) Then
            ok = False
        Else
            For i = 1 To lo.ListColumns.Count
                If StrComp(lo.ListColumns(i).Name, expected(LBound(expected) + i - 1), vbTextCompare) <> 0 Then
                    ok = False: Exit For
                End If
            Next i
        End If
        If Not ok Then rebuild = True
    Else
        rebuild = True
    End If

    If rebuild Then
        If Not lo Is Nothing Then On Error Resume Next: lo.Delete: On Error GoTo 0
        Set lo = EnsureTable(ws, "tblRoster", expected)
        ApplyRosterValidation lo
    Else
        ApplyRosterValidation lo
    End If
    Set EnsureRosterTable = lo
End Function

Private Sub ApplyRosterValidation(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    Dim colRole As Long, colVel As Long
    colRole = lo.ListColumns("Role").Index
    colVel = lo.ListColumns("ContributesToVelocity").Index

    Dim rngRole As Range, rngVel As Range
    Set rngRole = BodyOrReserveRange(lo, colRole, 1000)
    Set rngVel = BodyOrReserveRange(lo, colVel, 1000)

    With rngRole.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:="QA,Developer,Analyst,Squad Leader,Project Manager"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    With rngVel.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:="Yes,No"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Function BodyOrReserveRange(ByVal lo As ListObject, ByVal colIndex As Long, ByVal reserveRows As Long) As Range
    On Error Resume Next
    Dim r As Range
    If Not lo.DataBodyRange Is Nothing Then
        Set r = lo.DataBodyRange.Columns(colIndex)
    End If
    On Error GoTo 0
    If r Is Nothing Then
        Set r = lo.HeaderRowRange.Columns(colIndex).Offset(1).Resize(reserveRows, 1)
    End If
    Set BodyOrReserveRange = r
End Function

Private Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir(p)) > 0)
    On Error GoTo 0
End Function

Private Function PathJoin(ByVal a As String, ByVal b As String) As String
    If Right$(a, 1) = "\" Or Right$(a, 1) = "/" Then
        PathJoin = a & b
    Else
        PathJoin = a & Application.PathSeparator & b
    End If
End Function

Private Function PickFile(ByVal title As String) As String
    On Error Resume Next
    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    If Err.Number <> 0 Or fd Is Nothing Then
        On Error GoTo 0
        PickFile = InputBox(title & vbCrLf & "Enter full path to CSV:")
        Exit Function
    End If
    On Error GoTo 0
    With fd
        .Title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV", "*.csv"
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

Private Sub LogEvent(ByVal action As String, ByVal outcome As String, ByVal details As String)
    Dim lo As ListObject
    Set lo = EnsureTable(EnsureSheet("Logs"), "tblLogs", Array("Timestamp", "User", "Action", "Outcome", "Details"))
    Dim r As ListRow: Set r = lo.ListRows.Add
    r.Range(1, 1).Value = Now
    r.Range(1, 2).Value = Environ$("USERNAME")
    r.Range(1, 3).Value = action
    r.Range(1, 4).Value = outcome
    r.Range(1, 5).Value = details
End Sub

' -------------------- Availability sheet (simple) --------------------

Private Sub EnsureDashboard()
    Dim ws As Worksheet: Set ws = EnsureSheet("Dashboard")
    ' simple labels
    ws.Range("A1").Value = "Capacity Tracker – Dashboard"
    ws.Range("A2").Value = "Team:"
    ws.Range("B2").Formula = "=ActiveTeam"
    ws.Range("A4").Value = "Actions"
    ws.Range("A6").Value = "Sprint Length (workdays)"
    ws.Range("B6").Formula = "=SprintLengthDays"
    ws.Range("A1:A6").Font.Bold = True

    ' Create or refresh the button (single action)
    On Error Resume Next
    ws.Buttons("btnCreateAvailability").Delete
    ws.Buttons("btnAdvanceAvailability").Delete
    On Error GoTo 0
    Dim btn2 As Button
    Set btn2 = ws.Buttons.Add(Left:=20, Top:=80, Width:=240, Height:=28)
    btn2.Name = "btnAdvanceAvailability"
    btn2.OnAction = "modCapacityPlanner.CreateOrAdvanceAvailability"
    btn2.Characters.Text = "Create/Advance Availability"
End Sub

Public Sub CreateTeamAvailability()
    On Error GoTo Fail
    Dim yr As Integer, q As Integer, s As Integer
    If Not PromptForQuarterSprint(yr, q, s) Then Exit Sub
    Dim sStart As Date: sStart = QuarterStartDate(yr, q) + (s - 1) * 14

    CreateTeamAvailabilityAtDate sStart, Nothing
    Exit Sub
Fail:
    MsgBox "CreateTeamAvailability failed: " & Err.Description, vbExclamation
End Sub

Public Sub CreateOrAdvanceAvailability()
    On Error GoTo Fail
    Dim last As Worksheet: Set last = FindLatestAvailability()
    If last Is Nothing Then
        ' none exists → prompt
        CreateTeamAvailability
        Exit Sub
    End If

    Dim lastStart As Date
    On Error Resume Next
    lastStart = CDate(last.Cells(6, 2).Value)
    On Error GoTo 0
    If lastStart = 0 Then
        ' fallback: derive from sprint tag
        Dim yr As Integer, q As Integer, s As Integer
        If ParseTagFromName(last.Name, yr, q, s) Then
            lastStart = DateSerial(yr, (q - 1) * 3 + 1, 1) + (s - 1) * 14
        Else
            lastStart = Date
        End If
    End If
    Dim nextStart As Date: nextStart = DateAdd("d", 14, lastStart)
    CreateTeamAvailabilityAtDate nextStart, last
    Exit Sub
Fail:
    MsgBox "CreateOrAdvanceAvailability failed: " & Err.Description, vbExclamation
End Sub

Private Sub CreateTeamAvailabilityAtDate(ByVal sStart As Date, ByVal toHide As Worksheet)
    Dim sheetName As String
    sheetName = FormatSprintTag(sStart) & " Team Availability"
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = NextUniqueName(sheetName)

    Dim members As Variant, roles As Variant, contrib As Variant
    members = GetRosterColumn("Member")
    roles = GetRosterColumn("Role")
    contrib = GetRosterColumn("ContributesToVelocity")
    If IsEmpty(members) Then
        MsgBox "No roster members found in tblRoster.", vbExclamation
        Exit Sub
    End If

    ' Build ordered index: contributors first
    Dim order() As Long, count As Long, yesCount As Long
    order = BuildRosterOrder(members, contrib, roles, count, yesCount)

    ' Headers
    ws.Range("A5").Value = "Day of Week"
    ws.Range("B5").Value = "Date"
    ws.Range("C5").Value = "Sprint Day"
    Dim i As Long, col As Long: col = 4
    For i = 1 To count
        Dim idx As Long: idx = order(i)
        Dim hdr As String: hdr = CStr(members(idx))
        If Not IsEmpty(roles) Then
            If idx <= UBound(roles) And Len(roles(idx) & "") > 0 Then hdr = hdr & " (" & roles(idx) & ")"
        End If
        ws.Cells(5, col).Value = hdr
        col = col + 1
    Next i

    ' Top-left metrics placeholders
    ws.Range("A1").Value = "Average Velocity per Available Day"
    ws.Range("B1").Value = 0
    ws.Range("A2").Value = "Available Days"
    ws.Range("A3").Value = "Full Capacity Points"

    ' Fill 14 calendar days starting at sprint start
    Dim targetDays As Long: targetDays = CLng(GetNameValueOr("SprintLengthDays", "10"))
    Dim dayIndex As Long, row As Long: row = 6
    Dim sprintDay As Long: sprintDay = 0
    For dayIndex = 0 To 13
        Dim d As Date: d = sStart + dayIndex
        ws.Cells(row, 1).Value = Format$(d, "dddd")
        ws.Cells(row, 2).Value = d
        ws.Cells(row, 2).NumberFormat = "m/d"
        If IsWorkday(d) And sprintDay < targetDays Then
            sprintDay = sprintDay + 1
            ws.Cells(row, 3).Value = sprintDay
        Else
            ws.Cells(row, 3).Value = 0
        End If
        ' default availability: 1 on workdays (only while sprint days not exceeded), 0 on weekends
        col = 4
        For i = 1 To count
            ws.Cells(row, col).Value = IIf(IsWorkday(d) And sprintDay > 0, 1, 0)
            col = col + 1
        Next i
        row = row + 1
    Next dayIndex

    ' Totals row
    ws.Cells(row, 1).Value = "Total Days"
    For col = 4 To 3 + count
        ws.Cells(row, col).FormulaR1C1 = "=SUM(R[-14]C:R[-1]C)"
    Next col

    ' Now bind top-left metrics
    Dim rngYes As String
    If yesCount > 0 Then
        rngYes = ws.Range(ws.Cells(row, 4), ws.Cells(row, 3 + yesCount)).Address(False, False)
        ws.Range("B2").Formula = "=SUM(" & rngYes & ")"
        ws.Range("B3").Formula = "=IFERROR(B2*(DefaultHoursPerDay/DefaultHoursPerPoint),0)"
    Else
        ws.Range("B2").Value = 0
        ws.Range("B3").Value = 0
    End If

    ' Basic formatting (light)
    ws.Range(ws.Cells(5, 1), ws.Cells(5, 3 + count)).Font.Bold = True
    ws.Columns("A:A").ColumnWidth = 14
    ws.Columns("B:B").ColumnWidth = 8
    ws.Columns("C:C").ColumnWidth = 10
    ws.Columns("D:Z").ColumnWidth = 10
    ws.Range("A5").Select
    If Not toHide Is Nothing Then toHide.Visible = 0 ' xlSheetHidden
    MsgBox "Availability sheet created: " & ws.Name, vbInformation
End Sub

Private Function FormatSprintTag(ByVal startDate As Date) As String
    Dim yr As Integer: yr = Year(startDate)
    Dim q As Integer: q = Int((Month(startDate) - 1) / 3) + 1
    Dim qStart As Date: qStart = DateSerial(yr, (q - 1) * 3 + 1, 1)
    Dim daysFromQ As Long: daysFromQ = CLng(startDate - qStart)
    Dim s As Integer: s = Int(daysFromQ / 14) + 1
    If s < 1 Then s = 1
    If s > 7 Then s = 7
    FormatSprintTag = yr & " Q" & q & " S" & s
End Function

Private Function ParseTagFromName(ByVal nm As String, ByRef yr As Integer, ByRef q As Integer, ByRef s As Integer) As Boolean
    On Error GoTo Fail
    Dim base As String
    Dim p As Long: p = InStr(1, nm, " Team Availability", vbTextCompare)
    If p > 1 Then
        base = Trim$(Left$(nm, p - 1))
    Else
        base = nm
    End If
    Dim parts() As String: parts = Split(base, " ")
    If UBound(parts) < 2 Then GoTo Fail
    yr = CInt(parts(0))
    q = CInt(Replace(parts(1), "Q", ""))
    s = CInt(Replace(parts(2), "S", ""))
    ParseTagFromName = True
    Exit Function
Fail:
    ParseTagFromName = False
End Function

Private Function IsWorkday(ByVal d As Date) As Boolean
    Dim dow As Integer: dow = Weekday(d, vbSunday)
    IsWorkday = (dow >= vbMonday And dow <= vbFriday)
End Function

Private Function FindLatestAvailability() As Worksheet
    Dim ws As Worksheet
    Dim best As Worksheet
    Dim by As Integer, bq As Integer, bs As Integer
    by = -1: bq = -1: bs = -1
    For Each ws In ThisWorkbook.Worksheets
        Dim y As Integer, q As Integer, s As Integer
        If ParseTagFromName(ws.Name, y, q, s) Then
            If (y > by) Or (y = by And q > bq) Or (y = by And q = bq And s > bs) Then
                Set best = ws: by = y: bq = q: bs = s
            End If
        End If
    Next ws
    Set FindLatestAvailability = best
End Function

Private Function GetNameValueOr(ByVal nm As String, ByVal fallback As String) As String
    On Error Resume Next
    Dim v As String
    v = CStr(ThisWorkbook.Names(nm).RefersToRange.Value)
    On Error GoTo 0
    If Len(v) = 0 Then v = fallback
    GetNameValueOr = v
End Function

Private Function PromptForDate(ByVal prompt As String) As Date
    Dim s As String
    s = InputBox(prompt, "Sprint Start Date", Format$(Date, "m/d/yyyy"))
    If Len(s) = 0 Then Exit Function
    On Error Resume Next
    PromptForDate = CDate(s)
    On Error GoTo 0
End Function

Private Function QuarterStartDate(ByVal yr As Integer, ByVal q As Integer) As Date
    If q < 1 Then q = 1
    If q > 4 Then q = 4
    QuarterStartDate = DateSerial(yr, (q - 1) * 3 + 1, 1)
End Function

Private Function PromptForQuarterSprint(ByRef yr As Integer, ByRef q As Integer, ByRef s As Integer) As Boolean
    On Error GoTo Fail
    Dim defYr As Integer: defYr = Year(Date)
    Dim defQ As Integer: defQ = Int((Month(Date) - 1) / 3) + 1
    Dim defS As Integer: defS = 1
    Dim tmp As Variant
    tmp = Application.InputBox("Year", "Sprint Year", defYr, Type:=1)
    If tmp = False Then Exit Function
    yr = CInt(tmp)
    tmp = Application.InputBox("Quarter (1-4)", "Sprint Quarter", defQ, Type:=1)
    If tmp = False Then Exit Function
    q = CInt(tmp)
    tmp = Application.InputBox("Sprint in Quarter (1-7)", "Sprint Number", defS, Type:=1)
    If tmp = False Then Exit Function
    s = CInt(tmp)
    If q < 1 Or q > 4 Or s < 1 Or s > 7 Then GoTo Fail
    PromptForQuarterSprint = True
    Exit Function
Fail:
    PromptForQuarterSprint = False
End Function

Private Function NextUniqueName(ByVal base As String) As String
    Dim sanitized As String
    sanitized = SafeSheetName(base)
    If Len(sanitized) = 0 Then sanitized = "Sheet"

    Dim name As String: name = sanitized
    Dim n As Integer: n = 1
    Do While SheetExists(name)
        n = n + 1
        Dim suffix As String: suffix = " (" & n & ")"
        Dim allowed As Long: allowed = 31 - Len(suffix)
        If allowed < 1 Then allowed = 1
        name = Left$(sanitized, allowed) & suffix
    Loop
    NextUniqueName = name
End Function

Private Function GetRosterColumn(ByVal colName As String) As Variant
    Dim lo As ListObject
    On Error Resume Next
    Set lo = EnsureConfig().ListObjects("tblRoster")
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    Dim idx As Long
    On Error Resume Next: idx = lo.ListColumns(colName).Index: On Error GoTo 0
    If idx = 0 Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    Dim arr() As Variant
    arr = lo.ListColumns(idx).DataBodyRange.Value
    ' Flatten to 1-D variant array
    Dim i As Long, tmp() As Variant
    ReDim tmp(1 To UBound(arr, 1))
    For i = 1 To UBound(arr, 1)
        tmp(i) = arr(i, 1)
    Next i
    GetRosterColumn = tmp
End Function

' Return the primary Config sheet; migrate from old sheets if needed
Private Function EnsureConfig() As Worksheet
    Dim cfg As Worksheet
    If SheetExists("Config") Then
        Set cfg = Worksheets("Config")
    Else
        Set cfg = EnsureSheet("Config")
        ' migrate roster if present
        If SheetExists("Config_Teams") Then
            Dim src As Worksheet: Set src = Worksheets("Config_Teams")
            Dim loSrc As ListObject
            On Error Resume Next: Set loSrc = src.ListObjects("tblRoster"): On Error GoTo 0
            Dim loDst As ListObject: Set loDst = EnsureRosterTable(cfg)
            If Not loSrc Is Nothing Then
                If loSrc.ListRows.Count > 0 Then
                    loDst.DataBodyRange.ClearContents
                    loDst.DataBodyRange.Resize(loSrc.DataBodyRange.Rows.Count, loDst.ListColumns.Count).Value = loSrc.DataBodyRange.Value
                End If
            End If
            src.Visible = 0 ' hide old
        Else
            Call EnsureRosterTable(cfg)
        End If
        ' migrate named values if old sheet exists
        If SheetExists("Config_Sprints") Then
            Dim s As Worksheet: Set s = Worksheets("Config_Sprints")
            Dim namesArr As Variant: namesArr = Array("ActiveTeam","TemplateVersion","SprintLengthDays","DefaultHoursPerDay","DefaultAllocationPct","DefaultHoursPerPoint","RolesWithVelocity")
            Dim i As Long
            For i = LBound(namesArr) To UBound(namesArr)
                On Error Resume Next
                Dim nm As Name: Set nm = ThisWorkbook.Names(CStr(namesArr(i)))
                On Error GoTo 0
                If Not nm Is Nothing Then
                    ' keep value and rebind to Config sheet same H-row
                    Dim rowOff As Long: rowOff = 2 + i
                    cfg.Range("H" & rowOff).Value = nm.RefersToRange.Value
                    ' Rebind to sheet-qualified (no external workbook path)
                    nm.RefersTo = "='" & cfg.Name & "'!" & cfg.Range("H" & rowOff).Address(True, True, xlA1)
                End If
            Next i
            s.Visible = 0 ' hide old
        End If
    End If
    Set EnsureConfig = cfg
End Function

Private Function BuildRosterOrder(ByVal members As Variant, ByVal contrib As Variant, ByVal roles As Variant, ByRef outCount As Long, ByRef yesCount As Long) As Long()
    Dim n As Long: n = UBound(members)
    Dim devY() As Long, qaY() As Long, anaN() As Long, slN() As Long, pmN() As Long, other() As Long
    ReDim devY(1 To n): ReDim qaY(1 To n): ReDim anaN(1 To n): ReDim slN(1 To n): ReDim pmN(1 To n): ReDim other(1 To n)
    Dim cd As Long, cq As Long, ca As Long, cs As Long, cp As Long, co As Long
    Dim i As Long
    For i = 1 To n
        Dim isYes As Boolean: isYes = False
        Dim c As String: c = ""
        If Not IsEmpty(contrib) Then If i <= UBound(contrib) Then c = CStr(contrib(i))
        If UCase$(Left$(Trim$(c), 1)) = "Y" Then isYes = True

        Dim r As String: r = ""
        If Not IsEmpty(roles) Then If i <= UBound(roles) Then r = UCase$(Trim$(CStr(roles(i))))

        If isYes And r = "DEVELOPER" Then cd = cd + 1: devY(cd) = i
        ElseIf isYes And r = "QA" Then cq = cq + 1: qaY(cq) = i
        ElseIf Not isYes And r = "ANALYST" Then ca = ca + 1: anaN(ca) = i
        ElseIf Not isYes And r = "SQUAD LEADER" Then cs = cs + 1: slN(cs) = i
        ElseIf Not isYes And (r = "PROJECT MANAGER" Or r = "PROJECT MANAGER (SCRUM MASTER)") Then cp = cp + 1: pmN(cp) = i
        Else
            co = co + 1: other(co) = i
        End If
    Next i
    yesCount = cd + cq
    outCount = yesCount + ca + cs + cp + co
    Dim order() As Long: ReDim order(1 To outCount)
    Dim k As Long: k = 0
    For i = 1 To cd: k = k + 1: order(k) = devY(i): Next i
    For i = 1 To cq: k = k + 1: order(k) = qaY(i): Next i
    For i = 1 To ca: k = k + 1: order(k) = anaN(i): Next i
    For i = 1 To cs: k = k + 1: order(k) = slN(i): Next i
    For i = 1 To cp: k = k + 1: order(k) = pmN(i): Next i
    For i = 1 To co: k = k + 1: order(k) = other(i): Next i
    BuildRosterOrder = order
End Function

Private Function SafeSheetName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array(":", "\\", "/", "?", "*", "[", "]")
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), " ")
    Next i
    s = Trim$(s)
    If Len(s) > 31 Then s = Left$(s, 31)
    SafeSheetName = s
End Function

' Optional: permanently delete old config sheets after migration
Public Sub DeleteOldConfigSheets()
    RemoveSheetIfExists "Config_Teams"
    RemoveSheetIfExists "Config_Sprints"
    MsgBox "Old config sheets removed (if present).", vbInformation
End Sub
