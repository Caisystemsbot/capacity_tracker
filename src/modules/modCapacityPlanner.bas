Option Explicit

' Excel constants (avoid references)
Private Const xlSrcRange As Long = 1
Private Const xlYes As Long = 1
Private Const msoFileDialogFilePicker As Long = 3

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

Public Sub ImportPTO_CSV()
    On Error GoTo Fail
    Dim ws As Worksheet: Set ws = EnsureSheet("Calendars")
    Dim lo As ListObject: Set lo = EnsureTable(ws, "tblPTO", Array("Team", "Member", "Date", "Hours", "Source"))

    Dim path As String
    path = PickFile("Select PTO CSV (Team,Member,Date,Hours,Source)")
    If Len(path) = 0 Then Exit Sub

    ImportCsvToTable path, lo, True
    LogEvent "ImportPTO_CSV", "OK", path
    MsgBox "PTO rows imported.", vbInformation
    Exit Sub
Fail:
    LogEvent "ImportPTO_CSV", "ERROR", Err.Description
    MsgBox "PTO import failed: " & Err.Description, vbExclamation
End Sub

Public Sub HealthCheck()
    Dim issues As String
    If Not SheetExists("Calendars") Then issues = issues & "- Missing sheet: Calendars" & vbCrLf
    If Not SheetExists("Config_Teams") Then issues = issues & "- Missing sheet: Config_Teams" & vbCrLf
    If Not SheetExists("Logs") Then issues = issues & "- Missing sheet: Logs" & vbCrLf
    If Len(issues) = 0 Then
        MsgBox "Health check OK.", vbInformation
    Else
        MsgBox "Health check issues:" & vbCrLf & issues, vbExclamation
    End If
End Sub

' -------------------- internals --------------------

Private Sub EnsureSheets()
    EnsureSheet "Calendars"
    EnsureSheet "Config_Teams"
    EnsureSheet "Config_Sprints"
    EnsureSheet "Logs"
End Sub

Private Sub EnsureTables()
    Dim ws As Worksheet
    Set ws = EnsureSheet("Config_Teams")
    EnsureTable ws, "tblRoster", Array("Team", "Member", "Role", "HoursPerDay", "AllocationPct")

    Set ws = EnsureSheet("Calendars")
    EnsureTable ws, "tblHolidays", Array("Date", "Region", "Name")
    EnsureTable ws, "tblPTO", Array("Team", "Member", "Date", "Hours", "Source")

    Set ws = EnsureSheet("Logs")
    EnsureTable ws, "tblLogs", Array("Timestamp", "User", "Action", "Outcome", "Details")
End Sub

Private Sub SeedNamedValues()
    Dim ws As Worksheet: Set ws = EnsureSheet("Config_Sprints")
    EnsureNamedValue "ActiveTeam", ws.Range("H2"), "CraicForce"
    EnsureNamedValue "TemplateVersion", ws.Range("H3"), "0.1.0"
    EnsureNamedValue "SprintLengthDays", ws.Range("H4"), 10
    EnsureNamedValue "DefaultHoursPerDay", ws.Range("H5"), 6.5
    EnsureNamedValue "DefaultAllocationPct", ws.Range("H6"), 0.8
    EnsureNamedValue "DefaultHoursPerPoint", ws.Range("H7"), 6
End Sub

Private Sub SeedSamplesIfPresent()
    Dim base As String
    base = ThisWorkbook.Path
    If Len(base) = 0 Then Exit Sub

    Dim roster As String: roster = PathJoin(base, "data\roster_example.csv")
    Dim hol As String: hol = PathJoin(base, "data\holidays.csv")

    Dim lo As ListObject
    If FileExists(roster) Then
        Set lo = EnsureTable(EnsureSheet("Config_Teams"), "tblRoster", Array("Team", "Member", "Role", "HoursPerDay", "AllocationPct"))
        If lo.ListRows.Count = 0 Then ImportCsvToTable roster, lo, True
    End If
    If FileExists(hol) Then
        Set lo = EnsureTable(EnsureSheet("Calendars"), "tblHolidays", Array("Date", "Region", "Name"))
        If lo.ListRows.Count = 0 Then ImportCsvToTable hol, lo, True
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

Private Function EnsureTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal headers As Variant) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then
        Dim c As Long
        For c = LBound(headers) To UBound(headers)
            ws.Cells(1, 1 + c - LBound(headers)).Value = headers(c)
        Next c
        Dim lastCol As Long: lastCol = UBound(headers) - LBound(headers) + 1
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)), , xlYes)
        lo.Name = tableName
    End If
    Set EnsureTable = lo
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

