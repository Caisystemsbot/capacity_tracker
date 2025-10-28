Option Explicit

' Excel constants (avoid references)
Private Const xlSrcRange As Long = 1
Private Const xlYes As Long = 1
Private Const msoFileDialogFilePicker As Long = 3
Private Const xlValidateList As Long = 3
Private Const xlValidAlertStop As Long = 1
Private Const xlBetween As Long = 1
Private Const xlValidateDecimal As Long = 2
Private Const xlDatabase As Long = 1
Private Const xlRowField As Long = 1
Private Const xlColumnField As Long = 2
Private Const xlDataField As Long = 4
Private Const xlSum As Long = -4157
Private Const xlCount As Long = -4112
Private Const xlColumnClustered As Long = 51
Private Const xlAreaStacked As Long = 76
Private Const xlXYScatter As Long = -4169

Public Sub Bootstrap()
    On Error GoTo Fail
    LogStart "Bootstrap"
    Application.ScreenUpdating = False

    Step_EnsureSheets
    Step_EnsureTables
    Step_SeedNamedValues
    Step_SeedSamplesIfPresent

    Application.ScreenUpdating = True
    If IsVerbose() Then MsgBox "Bootstrap complete.", vbInformation
    LogOk "Bootstrap"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    LogErr "Bootstrap", "Err " & Err.Number & ": " & Err.Description
    MsgBox "Bootstrap failed: " & Err.Description, vbExclamation
End Sub

Private Sub Step_EnsureSheets()
    On Error GoTo Fail
    LogStart "EnsureSheets"
    EnsureSheets
    LogOk "EnsureSheets"
    Exit Sub
Fail:
    LogErr "EnsureSheets", "Err " & Err.Number & " (Erl=" & Erl & "): " & Err.Description
    Err.Raise Err.Number, , Err.Description
End Sub

' -------------------- WIP CSV Sanitizer --------------------

' Import and sanitize a WIP CSV containing time-in-status durations and dates.
' Expected columns (case/spacing flexible; extras ignored):
' - Created (date-time)
' - Resolved (date-time, optional)
' - Time In Todo (days, decimal)
' - Time In Progress (days, decimal)
' - Time In Testing (days, decimal)
' - Time In Review (days, decimal)
' Optional pass-through if present: Issue key, Summary
' Output goes to sheet WIP_Facts with table tblWIPFacts.
Public Sub WIP_ImportCSV()
    On Error GoTo Fail
    LogStart "WIP_ImportCSV"

    Dim path As String
    path = PickFile("Pick WIP CSV with Created/Resolved and time-in-status columns")
    If Len(Trim$(path)) = 0 Then GoTo CancelOp
    If Not FileExists(path) Then Err.Raise 5, , "CSV not found: " & path

    Dim ws As Worksheet: Set ws = EnsureSheet("WIP_Facts")
    Dim lo As ListObject: Set lo = EnsureWIPFactsTable(ws)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object: Set ts = fso.OpenTextFile(path, 1)
    If ts.AtEndOfStream Then Err.Raise 9, , "CSV empty: " & path

    Dim header As String: header = ts.ReadLine
    Dim cols As Variant: cols = Split(header, ",")
    Dim names As Object: Set names = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        names(Norm(CStr(cols(i)))) = i + 1 ' 1-based index for Split array
    Next i

    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    ' Map minimal and optional columns
    WIP_MapCol idx, names, "IssueKey", Array("issue key","key","id","ticket")
    WIP_MapCol idx, names, "Summary", Array("summary","title")
    WIP_MapCol idx, names, "Created", Array("created","created date","created on")
    WIP_MapCol idx, names, "Resolved", Array("resolved","resolved date","done date","resolution date","closed")
    WIP_MapCol idx, names, "TimeInTodo", Array("time in todo","time in to do","in todo","todo days")
    WIP_MapCol idx, names, "TimeInProgress", Array("time in progress","in progress","in progress days")
    WIP_MapCol idx, names, "TimeInTesting", Array("time in testing","in testing","testing days")
    WIP_MapCol idx, names, "TimeInReview", Array("time in review","in review","review days")

    If Not idx.Exists("Created") Then Err.Raise 1004, , "CSV missing 'Created' column"

    Dim rowText As String
    Do While Not ts.AtEndOfStream
        rowText = ts.ReadLine
        If Len(Trim$(rowText)) = 0 Then GoTo ContinueLoop
        Dim parts As Variant: parts = Split(rowText, ",")
        Dim r As ListRow: Set r = lo.ListRows.Add

        ' Read values defensively
        Dim created As Date, resolved As Date
        created = ToDateSafe(WIP_Get(parts, idx, "Created"))
        resolved = ToDateSafe(WIP_Get(parts, idx, "Resolved"))
        Dim tTodo As Double, tProg As Double, tTest As Double, tRev As Double
        tTodo = CDbl(Val(WIP_Get(parts, idx, "TimeInTodo")))
        tProg = CDbl(Val(WIP_Get(parts, idx, "TimeInProgress")))
        tTest = CDbl(Val(WIP_Get(parts, idx, "TimeInTesting")))
        tRev = CDbl(Val(WIP_Get(parts, idx, "TimeInReview")))
        Dim total As Double: total = 0#
        If tTodo > 0 Then total = total + tTodo
        If tProg > 0 Then total = total + tProg
        If tTest > 0 Then total = total + tTest
        If tRev > 0 Then total = total + tRev

        ' Write output row
        r.Range(1, 1).Value = WIP_Get(parts, idx, "IssueKey")
        r.Range(1, 2).Value = WIP_Get(parts, idx, "Summary")
        r.Range(1, 3).Value = created
        r.Range(1, 4).Value = resolved
        r.Range(1, 5).Value = tTodo
        r.Range(1, 6).Value = tProg
        r.Range(1, 7).Value = tTest
        r.Range(1, 8).Value = tRev
        r.Range(1, 9).Value = total ' WIPTotalDays
        If created <> 0 And resolved <> 0 Then r.Range(1, 10).Value = DateDiff("d", created, resolved) ' CycleCalDays
ContinueLoop:
    Loop
    ts.Close

    ws.Columns("A:J").AutoFit
    LogOk "WIP_ImportCSV"
    If IsVerbose() Then MsgBox "WIP CSV imported to 'WIP_Facts'.", vbInformation
    Exit Sub
CancelOp:
    LogErr "WIP_ImportCSV", "Cancelled by user"
    Exit Sub
Fail:
    LogErr "WIP_ImportCSV", "Err " & Err.Number & ": " & Err.Description
    MsgBox "WIP_ImportCSV failed: " & Err.Description, vbExclamation
End Sub

Private Function EnsureWIPFactsTable(ByVal ws As Worksheet) As ListObject
    Dim headers As Variant
    headers = Array("IssueKey","Summary","Created","Resolved","TimeInTodo","TimeInProgress","TimeInTesting","TimeInReview","WIPTotalDays","CycleCalDays")
    Dim lo As ListObject
    On Error Resume Next: Set lo = ws.ListObjects("tblWIPFacts"): On Error GoTo 0
    If lo Is Nothing Then
        Set lo = EnsureTable(ws, "tblWIPFacts", headers)
    Else
        ' if header mismatch, rebuild
        Dim ok As Boolean: ok = (lo.ListColumns.Count = (UBound(headers) - LBound(headers) + 1))
        If ok Then
            Dim i As Long
            For i = 1 To lo.ListColumns.Count
                If StrComp(lo.ListColumns(i).Name, headers(LBound(headers) + i - 1), vbTextCompare) <> 0 Then ok = False: Exit For
            Next i
        End If
        If Not ok Then
            On Error Resume Next: lo.Delete: On Error GoTo 0
            Set lo = EnsureTable(ws, "tblWIPFacts", headers)
        End If
    End If
    Set EnsureWIPFactsTable = lo
End Function

Private Sub WIP_MapCol(ByVal idx As Object, ByVal names As Object, ByVal key As String, ByVal candidates As Variant)
    Dim j As Long
    For j = LBound(candidates) To UBound(candidates)
        Dim k As String: k = Norm(CStr(candidates(j)))
        Dim found As Variant: found = FindByContains(names, k)
        If Not IsEmpty(found) Then idx(key) = CLng(found): Exit Sub
    Next j
End Sub

Private Function WIP_Get(ByVal parts As Variant, ByVal idx As Object, ByVal key As String) As String
    On Error Resume Next
    Dim p As Long: p = idx(key)
    If p > 0 And p <= UBound(parts) + 1 Then WIP_Get = CStr(parts(p - 1))
End Function

Private Sub Step_EnsureTables()
    On Error GoTo Fail
    LogStart "EnsureTables"
    EnsureTables
    LogOk "EnsureTables"
    Exit Sub
Fail:
    LogErr "EnsureTables", "Err " & Err.Number & " (Erl=" & Erl & "): " & Err.Description
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub Step_SeedNamedValues()
    On Error GoTo Fail
    LogStart "SeedNamedValues"
    SeedNamedValues
    LogOk "SeedNamedValues"
    Exit Sub
Fail:
    LogErr "SeedNamedValues", "Err " & Err.Number & " (Erl=" & Erl & "): " & Err.Description
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub Step_SeedSamplesIfPresent()
    On Error GoTo Fail
    LogStart "SeedSamplesIfPresent"
    SeedSamplesIfPresent
    LogOk "SeedSamplesIfPresent"
    Exit Sub
Fail:
    LogErr "SeedSamplesIfPresent", "Err " & Err.Number & " (Erl=" & Erl & "): " & Err.Description
    Err.Raise Err.Number, , Err.Description
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
    On Error GoTo CoreFail
    RemoveSheetIfExists "Calendars"
    Dim cfg As Worksheet: Set cfg = EnsureConfig()

    ' Create core sheets (ignore if already there)
    On Error Resume Next: Set cfg = EnsureSheet("Getting_Started"): On Error GoTo CoreFail
    On Error Resume Next: EnsureSheet "Dashboard": On Error GoTo CoreFail
    On Error Resume Next: EnsureSheet "Logs": On Error GoTo CoreFail

    ' Ensure Metrics sheet exists (build skeleton once)
    Dim m As Worksheet: Set m = EnsureSheet("Metrics")
    If m Is Nothing Then Err.Raise 91, , "Failed to create Metrics sheet"
    If Not HasTable(m, "tblMetrics") Then EnsureMetricsSheet

    ' Provide a single Raw_Data sample with time-in-status columns
    On Error GoTo ExpandedFail
    LogStart "EnsureRawDataSheet"
    EnsureRawDataSheet
    LogOk "EnsureRawDataSheet"
    Exit Sub

CoreFail:
    LogErr "EnsureSheets", "Core sheet creation failed: Err " & Err.Number & ": " & Err.Description
    Err.Raise Err.Number
ExpandedFail:
    LogErr "EnsureSheets", "EnsureRawDataSheet failed: Err " & Err.Number & ": " & Err.Description
    Resume Next
End Sub

Private Sub EnsureTables()
    Dim ws As Worksheet
    Set ws = EnsureConfig()
    On Error GoTo RosterFail
    LogStart "EnsureRosterTable"
    EnsureRosterTable ws
    LogOk "EnsureRosterTable"

    On Error GoTo LogsFail
    LogStart "EnsureLogsTable"
    Call EnsureSheetTable("Logs", "tblLogs", Array("Timestamp", "User", "Action", "Outcome", "Details"))
    LogOk "EnsureLogsTable"
    Exit Sub
RosterFail:
    LogErr "EnsureRosterTable", "Err " & Err.Number & ": " & Err.Description
    Resume Next
LogsFail:
    LogErr "EnsureLogsTable", "Err " & Err.Number & ": " & Err.Description
End Sub

' -------------------- Flow Metrics (pasteable entrypoint) --------------------

' Build core Flow Metrics charts (CFD/WIP, Throughput, Cycle Time) from a
' sanitized facts table. No setup required: detects a likely facts table
' automatically and creates a new sheet named "Flow_Metrics".
' - CFD/WIP (stacked To Do / In Progress / Done by day)
' - Throughput run chart (items completed per day)
' - Cycle Time scatter (Completed date vs days)
' The code is defensive: if a required column is missing, that chart is skipped.
Public Sub Flow_BuildCharts()
    On Error GoTo Fail
    LogStart "Flow_BuildCharts"

    Dim lo As ListObject
    Set lo = Flow_FindFactsTable()
    If lo Is Nothing Then
        MsgBox "Could not find a facts table with Created/Resolved columns.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = EnsureSheet("Flow_Metrics")
    ws.Cells.Clear
    ws.Range("A1").Value = "Flow Metrics"
    ws.Range("A1").Font.Bold = True

    Dim nextTop As Long: nextTop = 3

    ' (Removed) Cumulative Flow Diagram table (To Do / In Progress / Done)
    ' Proceed with Throughput and Cycle Time charts only

    ' 1) Throughput Run Chart (completed per day)
    Flow_WriteThroughput_Data lo, ws, nextTop
    Flow_MakeThroughput_Chart ws, nextTop
    nextTop = Flow_NextFreeTop(ws)

    ' 2) Cycle Time Scatter (completed vs days)
    Flow_WriteCycleScatter_Data lo, ws, nextTop
    Flow_MakeCycleScatter_Chart ws, nextTop

    ws.Columns("A:Z").AutoFit
    LogOk "Flow_BuildCharts"
    If IsVerbose() Then MsgBox "Flow metrics charts created on 'Flow_Metrics'.", vbInformation
    Exit Sub
Fail:
    LogErr "Flow_BuildCharts", "Err " & Err.Number & ": " & Err.Description
    MsgBox "Flow_BuildCharts failed: " & Err.Description, vbExclamation
End Sub

Private Function Flow_FindFactsTable() As ListObject
    ' Prefer a non-empty Jira_Facts!tblJiraFacts; else first table with Created and Resolved/CycleCalDays with rows
    On Error Resume Next
    Dim ws As Worksheet, lo As ListObject
    Set ws = Worksheets("Jira_Facts"): If Not ws Is Nothing Then Set lo = ws.ListObjects("tblJiraFacts")
    On Error GoTo 0
    If Not lo Is Nothing Then
        On Error Resume Next
        If Not lo.DataBodyRange Is Nothing Then If lo.ListRows.Count > 0 Then Set Flow_FindFactsTable = lo: Exit Function
        On Error GoTo 0
    End If

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If Flow_HasColumn(lo, "Created") And (Flow_HasColumn(lo, "Resolved") Or Flow_HasColumn(lo, "CycleCalDays")) Then
                If Not lo.DataBodyRange Is Nothing Then If lo.ListRows.Count > 0 Then
                    Set Flow_FindFactsTable = lo
                    Exit Function
                End If
            End If
        Next lo
    Next ws
End Function

Private Function Flow_HasColumn(ByVal lo As ListObject, ByVal name As String) As Boolean
    Flow_HasColumn = (Flow_GetColIndex(lo, name) > 0)
End Function

Private Function Flow_GetColIndex(ByVal lo As ListObject, ByVal key As String) As Long
    Dim names As Variant
    Select Case LCase$(key)
        Case "created"
            names = Array("created", "created date", "created on", "created_date")
        Case "resolved"
            names = Array("resolved", "resolved date", "resolution date", "done date", "closed")
        Case "startprogress"
            names = Array("startprogress", "start progress", "started", "in progress date")
        Case "timeintodo"
            names = Array("time in todo", "time in to do", "in todo", "todo days")
        Case "timeinprogress"
            names = Array("time in progress", "in progress", "in progress days")
        Case "timeintesting"
            names = Array("time in testing", "in testing", "testing days")
        Case "timeinreview"
            names = Array("time in review", "in review", "review days")
        Case Else
            names = Array(key)
    End Select
    Flow_GetColIndex = Flow_Col(lo, names)
End Function

Private Function Flow_Col(ByVal lo As ListObject, ByVal candidates As Variant) As Long
    Dim i As Long, j As Long
    For i = 1 To lo.ListColumns.Count
        Dim nm As String
        nm = Norm(lo.ListColumns(i).Name)
        For j = LBound(candidates) To UBound(candidates)
            Dim cand As String
            cand = Norm(CStr(candidates(j)))
            If Len(cand) > 0 Then
                If StrComp(nm, cand, vbTextCompare) = 0 _
                   Or InStr(1, nm, cand, vbTextCompare) > 0 _
                   Or InStr(1, cand, nm, vbTextCompare) > 0 Then
                    Flow_Col = i
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

Private Sub Flow_WriteCFD_Data(ByVal lo As ListObject, ByVal ws As Worksheet, ByVal topRow As Long)
    ' Writes [Date | To Do | In Progress | Done] starting at topRow.
    ' Uses Created, StartProgress (optional), Resolved (optional). If StartProgress is
    ' missing, the In Progress series is skipped.
    Dim idxC As Long, idxS As Long, idxR As Long, idxTodo As Long
    idxC = Flow_GetColIndex(lo, "Created")
    idxS = Flow_GetColIndex(lo, "StartProgress")
    idxR = Flow_GetColIndex(lo, "Resolved")
    idxTodo = Flow_GetColIndex(lo, "TimeInTodo")
    If idxC = 0 Then Exit Sub

    Dim minD As Date, maxD As Date, todayD As Date: todayD = Date
    Dim i As Long, c As Variant, s As Variant, r As Variant
    Dim hasAny As Boolean: hasAny = False
    For i = 1 To lo.ListRows.Count
        c = lo.DataBodyRange.Cells(i, idxC).Value
        If IsDate(c) Then
            If Not hasAny Then minD = CDate(c): maxD = CDate(c): hasAny = True
            If CDate(c) < minD Then minD = CDate(c)
            If CDate(c) > maxD Then maxD = CDate(c)
        End If
        If idxR > 0 Then
            r = lo.DataBodyRange.Cells(i, idxR).Value
            If IsDate(r) Then If CDate(r) > maxD Then maxD = CDate(r)
        End If
    Next i
    If Not hasAny Then Exit Sub
    If maxD < minD Then maxD = minD

    ' Limit to last ~90 days for readability if range is too long
    If DateDiff("d", minD, maxD) > 120 Then minD = DateAdd("d", -90, maxD)

    ' Headers
    ws.Cells(topRow, 1).Value = "CFD (WIP)"
    ws.Cells(topRow, 1).Font.Bold = True
    ws.Cells(topRow + 1, 1).Resize(1, 4).Value = Array("Date", "To Do", "In Progress", "Done")
    ws.Cells(topRow + 1, 1).Resize(1, 4).Font.Bold = True

    Dim row As Long: row = topRow + 2
    Dim d As Date
    For d = minD To maxD
        Dim toDo As Long: toDo = 0
        Dim inProg As Long: inProg = 0
        Dim doneC As Long: doneC = 0
        For i = 1 To lo.ListRows.Count
            c = lo.DataBodyRange.Cells(i, idxC).Value
            If IsDate(c) Then
                s = 0: r = 0
                If idxS > 0 Then
                    s = lo.DataBodyRange.Cells(i, idxS).Value
                ElseIf idxTodo > 0 Then
                    ' Derive a start-progress date from Created + TimeInTodo (days)
                    If IsDate(c) Then s = CDate(c) + CDbl(Val(lo.DataBodyRange.Cells(i, idxTodo).Value))
                End If
                If idxR > 0 Then r = lo.DataBodyRange.Cells(i, idxR).Value

                If (IsDate(r) And CDate(r) <= d) Then
                    doneC = doneC + 1
                ElseIf (idxS > 0 And IsDate(s) And CDate(s) <= d And (Not IsDate(r) Or CDate(r) > d)) Then
                    inProg = inProg + 1
                ElseIf CDate(c) <= d Then
                    toDo = toDo + 1
                End If
            End If
        Next i
        ws.Cells(row, 1).Value = d
        ws.Cells(row, 2).Value = toDo
        ws.Cells(row, 3).Value = IIf(idxS > 0, inProg, Empty)
        ws.Cells(row, 4).Value = doneC
        row = row + 1
    Next d
End Sub

Private Sub Flow_MakeCFD_Chart(ByVal ws As Worksheet, ByVal topRow As Long)
    ' Build stacked area chart over the CFD block
    Dim hdr As Range: Set hdr = ws.Cells(topRow + 1, 1)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, hdr.Column).End(xlUp).Row
    If lastRow <= hdr.Row + 1 Then Exit Sub
    Dim ch As ChartObject
    Set ch = ws.ChartObjects.Add(Left:=380, Top:=ws.Cells(topRow + 1, 1).Top, Width:=540, Height:=280)
    ch.Chart.ChartType = xlAreaStacked
    ch.Chart.SetSourceData ws.Range(ws.Cells(hdr.Row, hdr.Column), ws.Cells(lastRow, hdr.Column + 3))
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Cumulative Flow (WIP)"
End Sub

Private Sub Flow_WriteThroughput_Data(ByVal lo As ListObject, ByVal ws As Worksheet, ByVal topRow As Long)
    ' Completed items per day based on Resolved
    Dim idxR As Long
    idxR = Flow_GetColIndex(lo, "Resolved")
    If idxR = 0 Then Exit Sub

    Dim counts As Object: Set counts = CreateObject("Scripting.Dictionary")
    Dim i As Long, r As Variant, d As Date, key As String
    For i = 1 To lo.ListRows.Count
        r = lo.DataBodyRange.Cells(i, idxR).Value
        If IsDate(r) Then
            d = DateSerial(Year(r), Month(r), Day(r))
            key = CStr(d)
            If Not counts.Exists(key) Then counts(key) = 0
            counts(key) = counts(key) + 1
        End If
    Next i
    If counts.Count = 0 Then Exit Sub

    ' Sort keys
    Dim keys() As Variant: keys = counts.Keys
    Dim j As Long, k As Long
    For j = LBound(keys) To UBound(keys) - 1
        For k = j + 1 To UBound(keys)
            If CDate(keys(k)) < CDate(keys(j)) Then
                Dim t As Variant: t = keys(j): keys(j) = keys(k): keys(k) = t
            End If
        Next k
    Next j

    ws.Cells(topRow, 1).Value = "Throughput"
    ws.Cells(topRow, 1).Font.Bold = True
    ws.Cells(topRow + 1, 1).Resize(1, 2).Value = Array("Date", "Completed")
    ws.Cells(topRow + 1, 1).Resize(1, 2).Font.Bold = True
    Dim row As Long: row = topRow + 2
    For j = LBound(keys) To UBound(keys)
        ws.Cells(row, 1).Value = CDate(keys(j))
        ws.Cells(row, 2).Value = counts(keys(j))
        row = row + 1
    Next j
End Sub

Private Sub Flow_MakeThroughput_Chart(ByVal ws As Worksheet, ByVal topRow As Long)
    Dim hdr As Range: Set hdr = ws.Cells(topRow + 1, 1)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, hdr.Column).End(xlUp).Row
    If lastRow <= hdr.Row + 1 Then Exit Sub
    Dim ch As ChartObject
    Set ch = ws.ChartObjects.Add(Left:=380, Top:=ws.Cells(topRow + 1, 1).Top, Width:=540, Height:=240)
    ch.Chart.ChartType = xlColumnClustered
    ch.Chart.SetSourceData ws.Range(ws.Cells(hdr.Row, hdr.Column), ws.Cells(lastRow, hdr.Column + 1))
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Throughput (Completed per Day)"
End Sub

Private Sub Flow_WriteCycleScatter_Data(ByVal lo As ListObject, ByVal ws As Worksheet, ByVal topRow As Long)
    ' Completed date vs cycle time (calendar days preferred), fallback to Resolved-Created
    Dim idxR As Long, idxC As Long, idxCal As Long
    idxR = Flow_GetColIndex(lo, "Resolved")
    idxC = Flow_GetColIndex(lo, "Created")
    On Error Resume Next: idxCal = lo.ListColumns("CycleCalDays").Index: On Error GoTo 0
    If idxR = 0 Or idxC = 0 Then Exit Sub

    ws.Cells(topRow, 1).Value = "Cycle Time Scatter"
    ws.Cells(topRow, 1).Font.Bold = True
    ws.Cells(topRow + 1, 1).Resize(1, 2).Value = Array("ResolvedDate", "Days")
    ws.Cells(topRow + 1, 1).Resize(1, 2).Font.Bold = True

    Dim row As Long: row = topRow + 2
    Dim i As Long, r As Variant, c As Variant, d As Double
    For i = 1 To lo.ListRows.Count
        r = lo.DataBodyRange.Cells(i, idxR).Value
        c = lo.DataBodyRange.Cells(i, idxC).Value
        If IsDate(r) And IsDate(c) Then
            If idxCal > 0 Then
                d = Val(lo.DataBodyRange.Cells(i, idxCal).Value)
                If d <= 0 Then d = DateDiff("d", CDate(c), CDate(r))
            Else
                d = DateDiff("d", CDate(c), CDate(r))
            End If
            If d >= 0 Then
                ws.Cells(row, 1).Value = CDate(r)
                ws.Cells(row, 2).Value = d
                row = row + 1
            End If
        End If
    Next i
End Sub

Private Sub Flow_MakeCycleScatter_Chart(ByVal ws As Worksheet, ByVal topRow As Long)
    Dim hdr As Range: Set hdr = ws.Cells(topRow + 1, 1)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, hdr.Column).End(xlUp).Row
    If lastRow <= hdr.Row + 1 Then Exit Sub
    Dim rngX As Range, rngY As Range
    Set rngX = ws.Range(ws.Cells(hdr.Row + 1, hdr.Column), ws.Cells(lastRow, hdr.Column))
    Set rngY = ws.Range(ws.Cells(hdr.Row + 1, hdr.Column + 1), ws.Cells(lastRow, hdr.Column + 1))
    Dim ch As ChartObject
    Set ch = ws.ChartObjects.Add(Left:=20, Top:=ws.Cells(topRow + 1, 1).Top, Width:=340, Height:=260)
    ch.Chart.ChartType = xlXYScatter
    With ch.Chart.SeriesCollection.NewSeries
        .XValues = rngX
        .Values = rngY
        .Name = "=Cycle Time"
    End With
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "Cycle Time Scatter"
End Sub

Private Function Flow_NextFreeTop(ByVal ws As Worksheet) As Long
    ' Find the next free Y position to place another block ~ 16 rows below last used row
    Dim last As Long
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If last < 1 Then last = 1
    Flow_NextFreeTop = last + 3
End Function

Private Sub SeedNamedValues()
    Dim ws As Worksheet: Set ws = EnsureConfig()
    EnsureNamedValue "ActiveTeam", ws.Range("H2"), "CraicForce"
    EnsureNamedValue "TemplateVersion", ws.Range("H3"), "0.1.0"
    EnsureNamedValue "SprintLengthDays", ws.Range("H4"), 10
    EnsureNamedValue "DefaultHoursPerDay", ws.Range("H5"), 6.5
    EnsureNamedValue "DefaultAllocationPct", ws.Range("H6"), 1
    EnsureNamedValue "DefaultHoursPerPoint", ws.Range("H7"), 6
    EnsureNamedValue "RolesWithVelocity", ws.Range("H8"), "Developer,QA"
    EnsureNamedValue "VerboseLogging", ws.Range("H9"), True
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
    ' Safe naming: avoid workbook-level name collision (1004)
    On Error Resume Next
    lo.Name = tableName
    If Err.Number <> 0 Then
        Err.Clear
        lo.Name = UniqueTableName(ws.Parent, tableName)
    End If
    On Error GoTo 0
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
    ' UsedRange always returns a range; compute its bottom row
    usedBottom = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If usedBottom >= r Then r = usedBottom + 2

    ' Clamp to sheet bounds; if last cell glitch pushes us past the bottom, restart at top
    If r < 1 Then r = 1
    If r > ws.Rows.Count - 5 Then r = 1
    NextFreeRow = r
End Function

Private Sub EnsureNamedValue(ByVal nm As String, ByVal target As Range, ByVal defaultValue As Variant)
    Dim n As Name
    On Error Resume Next
    Set n = ThisWorkbook.Names(nm)
    On Error GoTo 0
    If n Is Nothing Then
        target.Value = defaultValue
        ThisWorkbook.Names.Add Name:=nm, RefersTo:=target
        Exit Sub
    End If
    Dim ref As Range
    On Error Resume Next
    Set ref = n.RefersToRange
    On Error GoTo 0
    If ref Is Nothing Then
        ' Rebind broken external/invalid name to the provided target cell
        n.RefersTo = "='" & target.Worksheet.Name & "'!" & target.Address(True, True, xlA1)
        target.Value = defaultValue
    ElseIf Len(CStr(ref.Value)) = 0 Then
        ref.Value = defaultValue
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
    ws.Range("G9").Value = "VerboseLogging (TRUE/FALSE)"
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
    On Error GoTo SafeExit
    Dim ws As Worksheet
    Set ws = EnsureSheet("Logs")

    Dim lo As ListObject
    Set lo = EnsureSheetTable("Logs", "tblLogs", Array("Timestamp", "User", "Action", "Outcome", "Details"))
    If lo Is Nothing Then GoTo SafeExit

    Dim r As ListRow
    On Error Resume Next
    Set r = lo.ListRows.Add
    If r Is Nothing Then GoTo SafeExit
    On Error GoTo 0
    r.Range(1, 1).Value = Now
    r.Range(1, 2).Value = Environ$("USERNAME")
    r.Range(1, 3).Value = action
    r.Range(1, 4).Value = outcome
    r.Range(1, 5).Value = details
SafeExit:
End Sub

Private Function IsVerbose() As Boolean
    On Error Resume Next
    Dim v As Variant
    v = ThisWorkbook.Names("VerboseLogging").RefersToRange.Value
    On Error GoTo 0
    If VarType(v) = vbString Then
        IsVerbose = (UCase$(CStr(v)) = "TRUE")
    ElseIf IsNumeric(v) Then
        IsVerbose = (CDbl(v) <> 0)
    Else
        IsVerbose = True
    End If
End Function

Private Function UniqueTableName(ByVal wb As Workbook, ByVal base As String) As String
    Dim name As String: name = base
    Dim i As Long: i = 1
    Do While TableNameExists(wb, name)
        i = i + 1
        name = base & "_" & CStr(i)
    Loop
    UniqueTableName = name
End Function

Private Function TableNameExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim sh As Worksheet, lo As ListObject
    For Each sh In wb.Worksheets
        For Each lo In sh.ListObjects
            If StrComp(lo.Name, nm, vbTextCompare) = 0 Then TableNameExists = True: Exit Function
        Next lo
    Next sh
    TableNameExists = False
End Function

Private Function HasTable(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    On Error Resume Next
    HasTable = Not ws.ListObjects(tableName) Is Nothing
    On Error GoTo 0
End Function

Private Function EnsureSheetTable(ByVal sheetName As String, ByVal tableName As String, ByVal headers As Variant) As ListObject
    Dim ws As Worksheet
    Set ws = EnsureSheet(sheetName)
    Dim lo As ListObject
    On Error Resume Next: Set lo = ws.ListObjects(tableName): On Error GoTo 0
    If lo Is Nothing Then Set lo = EnsureTable(ws, tableName, headers)
    Set EnsureSheetTable = lo
End Function

Private Sub LogStart(ByVal action As String, Optional ByVal details As String = "")
    If IsVerbose() Then LogEvent action, "START", details
End Sub

Private Sub LogOk(ByVal action As String, Optional ByVal details As String = "")
    If IsVerbose() Then LogEvent action, "OK", details
End Sub

Private Sub LogErr(ByVal action As String, Optional ByVal details As String = "")
    LogEvent action, "ERROR", details
End Sub

Public Sub Diagnostics_RunBootstrap()
    On Error Resume Next
    ThisWorkbook.Names("VerboseLogging").RefersToRange.Value = True
    On Error GoTo 0
    Bootstrap
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

    ' Create or refresh buttons, defensively removing any shape with same name
    Dim nm As Variant, shp As Shape
    For Each nm In Array("btnCreateAvailability", "btnAdvanceAvailability", "btnSanitizeRawAndBuild")
        On Error Resume Next
        ws.Buttons(CStr(nm)).Delete
        On Error GoTo 0
        For Each shp In ws.Shapes
            If StrComp(shp.Name, CStr(nm), vbTextCompare) = 0 Then shp.Delete
        Next shp
    Next nm

    ' Button: Create/Advance Availability
    Dim btn2 As Button
    Set btn2 = ws.Buttons.Add(Left:=20, Top:=80, Width:=240, Height:=28)
    On Error Resume Next
    btn2.Name = UniqueShapeName(ws, "btnAdvanceAvailability")
    On Error GoTo 0
    btn2.OnAction = "CreateOrAdvanceAvailability"
    btn2.Characters.Text = "Create/Advance Availability"

    ' (Removed) Build Jira Insights button; use Sanitize Raw + Build Insights instead

    ' Button: Sanitize Raw + Build Insights
    Dim btn4 As Button
    Set btn4 = ws.Buttons.Add(Left:=20, Top:=120, Width:=240, Height:=28)
    On Error Resume Next
    btn4.Name = UniqueShapeName(ws, "btnSanitizeRawAndBuild")
    On Error GoTo 0
    btn4.OnAction = "SanitizeRawAndBuildInsights"
    btn4.Characters.Text = "Sanitize Raw + Build Insights"

    ' Button: Refresh Samples
    Dim btn5 As Button
    Set btn5 = ws.Buttons.Add(Left:=20, Top:=200, Width:=240, Height:=28)
    On Error Resume Next
    btn5.Name = UniqueShapeName(ws, "btnRefreshSamples")
    On Error GoTo 0
    btn5.OnAction = "RefreshSamples"
    btn5.Characters.Text = "Refresh Samples"
End Sub

Private Function UniqueShapeName(ByVal ws As Worksheet, ByVal base As String) As String
    Dim nameCandidate As String: nameCandidate = base
    Dim i As Long: i = 1
    Dim exists As Boolean
    Do
        exists = False
        Dim s As Shape
        For Each s In ws.Shapes
            If StrComp(s.Name, nameCandidate, vbTextCompare) = 0 Then
                exists = True: Exit For
            End If
        Next s
        If Not exists Then Exit Do
        i = i + 1
        nameCandidate = base & "_" & CStr(i)
    Loop
    UniqueShapeName = nameCandidate
End Function

Public Sub CreateTeamAvailability()
    On Error GoTo Fail
    LogStart "CreateTeamAvailability"
    Dim yr As Integer, q As Integer, s As Integer
    If Not PromptForQuarterSprint(yr, q, s) Then Exit Sub
    Dim sStart As Date: sStart = QuarterStartDate(yr, q) + (s - 1) * 14

    CreateTeamAvailabilityAtDate sStart, Nothing
    LogOk "CreateTeamAvailability"
    Exit Sub
Fail:
    LogErr "CreateTeamAvailability", "Err " & Err.Number & " (Erl=" & Erl & "): " & Err.Description
    MsgBox "CreateTeamAvailability failed: " & Err.Description, vbExclamation
End Sub

Public Sub CreateOrAdvanceAvailability()
    On Error GoTo Fail
    LogStart "CreateOrAdvanceAvailability"
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
    LogOk "CreateOrAdvanceAvailability"
    Exit Sub
Fail:
    LogErr "CreateOrAdvanceAvailability", "Err " & Err.Number & " (Erl=" & Erl & "): " & Err.Description
    MsgBox "CreateOrAdvanceAvailability failed: " & Err.Description, vbExclamation
End Sub

Private Sub CreateTeamAvailabilityAtDate(ByVal sStart As Date, ByVal toHide As Worksheet)
    On Error GoTo Fail
    Dim phase As String
    Dim sheetName As String
    sheetName = FormatSprintTag(sStart) & " Team Availability"
    Dim ws As Worksheet
    phase = "AddSheet"
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = NextUniqueName(sheetName)

    Dim members As Variant, roles As Variant, contrib As Variant
    phase = "ReadRoster"
    members = GetRosterColumn("Member")
    roles = GetRosterColumn("Role")
    contrib = GetRosterColumn("ContributesToVelocity")
    If IsEmpty(members) Then
        MsgBox "No roster members found in tblRoster.", vbExclamation
        Exit Sub
    End If

    ' Build ordered index: contributors first, grouped by role
    Dim order() As Long, count As Long, yesCount As Long
    phase = "BuildOrder"
    order = BuildRosterOrder(members, contrib, roles, count, yesCount)

    ' Headers
    phase = "WriteHeaders"
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
    ws.Range("A1:A3").Font.Bold = True
    ws.Range("A1:B3").Interior.Color = RGB(221, 235, 247)
    ws.Range("A1:B3").Borders.Weight = 2

    ' Fill 14 calendar days starting at sprint start
    Dim targetDays As Long: targetDays = CLng(GetNameValueOr("SprintLengthDays", "10"))
    Dim dayIndex As Long, row As Long: row = 6
    Dim sprintDay As Long: sprintDay = 0
    phase = "FillDays"
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
    phase = "Totals"
    ws.Cells(row, 1).Value = "Total Days"
    For col = 4 To 3 + count
        ws.Cells(row, col).FormulaR1C1 = "=SUM(R[-14]C:R[-1]C)"
    Next col

    ' Bind metrics block to contributor columns (Developer + QA first)
    If yesCount > 0 Then
        Dim rngYes As String
        rngYes = ws.Range(ws.Cells(row, 4), ws.Cells(row, 3 + yesCount)).Address(False, False)
        ws.Range("B2").Formula = "=SUM(" & rngYes & ")"
        ws.Range("B3").Formula = "=IFERROR(B2*(DefaultHoursPerDay/DefaultHoursPerPoint),0)"
    Else
        ws.Range("B2").Value = 0
        ws.Range("B3").Value = 0
    End If

    ' Basic formatting (light)
    phase = "Format"
    ws.Range(ws.Cells(5, 1), ws.Cells(5, 3 + count)).Font.Bold = True
    ws.Range(ws.Cells(5, 1), ws.Cells(5, 3 + count)).Interior.Color = RGB(217, 225, 242)
    ' Title over member headers
    With ws.Range(ws.Cells(4, 4), ws.Cells(4, 3 + count))
        .Merge
        .Value = "Squad Member Availability"
        .HorizontalAlignment = -4108 ' xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders.Weight = 2
    End With
    ' Grid borders
    With ws.Range(ws.Cells(5, 1), ws.Cells(row, 3 + count)).Borders
        .LineStyle = 1
        .Weight = 2
    End With
    ' Role-based column coloring: contributors (first yesCount) light green, others light yellow
    If count > 0 Then
        Dim rngContrib As Range, rngNon As Range
        If yesCount > 0 Then
            Set rngContrib = ws.Range(ws.Cells(5, 4), ws.Cells(row, 3 + yesCount))
            rngContrib.Interior.Color = RGB(198, 239, 206) ' light green
        End If
        If count > yesCount Then
            Set rngNon = ws.Range(ws.Cells(5, 4 + yesCount), ws.Cells(row, 3 + count))
            rngNon.Interior.Color = RGB(255, 242, 204) ' light yellow
        End If
    End If
    ' Center availability cells
    If count > 0 Then ws.Range(ws.Cells(6, 4), ws.Cells(row - 1, 3 + count)).HorizontalAlignment = -4108
    ' Totals row emphasis
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 3 + count)).Font.Bold = True
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 3 + count)).Interior.Color = RGB(235, 241, 222)

    ' Auto-size columns
    ws.Columns("A:A").ColumnWidth = 14
    ws.Columns("B:B").ColumnWidth = 8
    ws.Columns("C:C").ColumnWidth = 10
    If count > 0 Then
        Dim ccol As Long
        ws.Range(ws.Cells(5, 4), ws.Cells(row, 3 + count)).EntireColumn.AutoFit
        ' enforce minimum width for member columns for readability
        For ccol = 4 To 3 + count
            If ws.Columns(ccol).ColumnWidth < 8 Then ws.Columns(ccol).ColumnWidth = 8
        Next ccol
    End If
    ' Freeze header row
    ws.Range("A6").Select
    ActiveWindow.FreezePanes = True
    If Not toHide Is Nothing Then toHide.Visible = 0 ' xlSheetHidden
    If IsVerbose() Then MsgBox "Availability sheet created: " & ws.Name, vbInformation
    Exit Sub
Fail:
    LogErr "CreateTeamAvailabilityAtDate", "Phase=" & phase & "; Err " & Err.Number & " (Erl=" & Erl & "): " & Err.Description
    Err.Raise Err.Number, , Err.Description
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
        ' migrate named values if old sheet exists (guard against broken externals)
        If SheetExists("Config_Sprints") Then
            Dim s As Worksheet: Set s = Worksheets("Config_Sprints")
            Dim namesArr As Variant: namesArr = Array("ActiveTeam","TemplateVersion","SprintLengthDays","DefaultHoursPerDay","DefaultAllocationPct","DefaultHoursPerPoint","RolesWithVelocity")
            Dim i As Long
            For i = LBound(namesArr) To UBound(namesArr)
                Dim nm As Name
                On Error Resume Next: Set nm = ThisWorkbook.Names(CStr(namesArr(i))): On Error GoTo 0
                If Not nm Is Nothing Then
                    Dim rowOff As Long: rowOff = 2 + i
                    Dim v As Variant, ref As String
                    On Error Resume Next
                    v = nm.RefersToRange.Value
                    If Err.Number <> 0 Then
                        Err.Clear
                        ref = nm.RefersTo
                        If Len(ref) > 1 And Left$(ref, 1) = "=" Then ref = Mid$(ref, 2)
                        ' Best-effort evaluate; ignore failures
                        v = v
                    End If
                    On Error GoTo 0
                    cfg.Range("H" & rowOff).Value = v
                    ' Rebind strictly to this workbook's Config sheet
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
        If Not IsEmpty(contrib) Then
            If i <= UBound(contrib) Then c = CStr(contrib(i))
        End If
        If UCase$(Left$(Trim$(c), 1)) = "Y" Then isYes = True

        Dim r As String: r = ""
        If Not IsEmpty(roles) Then
            If i <= UBound(roles) Then r = UCase$(NormalizeRole(CStr(roles(i))))
        End If

        If isYes And r = "DEVELOPER" Then
            cd = cd + 1: devY(cd) = i
        ElseIf isYes And r = "QA" Then
            cq = cq + 1: qaY(cq) = i
        ElseIf Not isYes And r = "ANALYST" Then
            ca = ca + 1: anaN(ca) = i
        ElseIf Not isYes And r = "SQUAD LEADER" Then
            cs = cs + 1: slN(cs) = i
        ElseIf Not isYes And (r = "PROJECT MANAGER" Or r = "PROJECT MANAGER (SCRUM MASTER)") Then
            cp = cp + 1: pmN(cp) = i
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

Private Function NormalizeRole(ByVal raw As String) As String
    Dim u As String: u = UCase$(Trim$(raw))
    If u = "DEV" Or InStr(u, "DEVELOPER") > 0 Or InStr(u, "ENGINEER") > 0 Then NormalizeRole = "Developer": Exit Function
    If u = "QA" Or InStr(u, "QUALITY") > 0 Or InStr(u, "TEST") > 0 Then NormalizeRole = "QA": Exit Function
    If u = "SL" Or InStr(u, "SQUAD") > 0 Then NormalizeRole = "Squad Leader": Exit Function
    If u = "PM" Or InStr(u, "SCRUM") > 0 Or InStr(u, "PROJECT MANAGER") > 0 Then NormalizeRole = "Project Manager": Exit Function
    If u = "BA" Or u = "ANALYST" Or InStr(u, "ANALYST") > 0 Or InStr(u, "BUSINESS ANALYST") > 0 Then NormalizeRole = "Analyst": Exit Function
    NormalizeRole = raw
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

' -------------------- Metrics sheet (skeleton) --------------------

Public Sub BuildMetricsSkeleton()
    On Error GoTo Fail
    LogStart "BuildMetricsSkeleton"
    EnsureMetricsSheet
    LogOk "BuildMetricsSkeleton"
    Exit Sub
Fail:
    LogErr "BuildMetricsSkeleton", "Err " & Err.Number & ": " & Err.Description
    MsgBox "Metrics build failed: " & Err.Description, vbExclamation
End Sub

' -------------------- Jira integration (basic) --------------------

Public Sub Jira_PopulateMetrics()
    ' Token-based API is not used. Build Power Query web queries instead
    On Error GoTo Fail
    LogStart "Jira_PopulateMetrics"
    Jira_CreateQueries
    ThisWorkbook.RefreshAll
    ' After refresh, if Jira_Metrics sheet exists, apply into Metrics table
    Jira_ApplyMetricsFromQuery
    LogOk "Jira_PopulateMetrics"
    If IsVerbose() Then MsgBox "Jira queries created/refreshed. Metrics updated from Jira_Metrics.", vbInformation
    Exit Sub
Fail:
    LogErr "Jira_PopulateMetrics", "Err " & Err.Number & ": " & Err.Description
    MsgBox "Jira import failed: " & Err.Description, vbExclamation
End Sub

' -------------------- Jira issues analysis (mock + normalize) --------------------

Public Sub Jira_BuildSampleIssues()
    Dim ws As Worksheet: Set ws = EnsureSheet("Jira_Issues_Sample")
    ws.Cells.Clear
    Dim headers As Variant
    headers = Array("Summary","Issue key","Issue id","Issue Type","Status","Created","Start Progress","Resolved","Fix Version/s","Custom field (Epic Link)","Custom field (Story Points)")
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    Dim r As Long: r = 2
    Call W(ws, r, Array("Onboard Tools", "FIINT-4000", 330001, "Story", "Done", #9/10/2025#, #9/23/2025#, "2025.09.23", "EPIC-100", 3)): r = r + 1
    Call W(ws, r, Array("Automation Cleanup", "FIINT-4010", 330010, "Story", "Done", #9/25/2025#, #10/7/2025#, "2025.10.07", "EPIC-100", 5)): r = r + 1
    Call W(ws, r, Array("Improve Logs", "FIINT-4020", 330020, "Task", "In Progress", #10/8/2025#, "", "2025.10.21", "EPIC-120", 2)): r = r + 1
    Call W(ws, r, Array("Release Steps", "FIINT-4071", 333071, "Story", "Done", #9/21/2025#, #10/15/2025#, "2025.10.15", "EPIC-140", 3)): r = r + 1
    ws.Rows(1).Font.Bold = True
    Dim lo As ListObject
    On Error Resume Next: Set lo = ws.ListObjects("tblJiraIssuesSample"): On Error GoTo 0
    If Not lo Is Nothing Then lo.Delete
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    lo.Name = "tblJiraIssuesSample"
End Sub

Private Sub W(ws As Worksheet, ByVal r As Long, ByVal arr As Variant)
    Dim c As Long
    For c = LBound(arr) To UBound(arr)
        ws.Cells(r, c + 1).Value = arr(c)
    Next c
End Sub

Public Sub BuildJiraInsights()
    On Error GoTo Fail
    LogStart "BuildJiraInsights"
    ' Ensure sample exists and normalize (placeholder until live query wired)
    EnsureSampleIssuesSheet
    Jira_NormalizeIssues_FromSample
    Jira_CreatePivotsAndCharts
    LogOk "BuildJiraInsights"
    If IsVerbose() Then MsgBox "Jira insights built.", vbInformation
    Exit Sub
Fail:
    LogErr "BuildJiraInsights", "Err " & Err.Number & ": " & Err.Description
    MsgBox "BuildJiraInsights failed: " & Err.Description, vbExclamation
End Sub

Private Sub Jira_CreatePivotsAndCharts()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim srcWs As Worksheet: Set srcWs = EnsureSheet("Jira_Facts")
    Dim lo As ListObject
    On Error Resume Next: Set lo = srcWs.ListObjects("tblJiraFacts"): On Error GoTo 0
    If lo Is Nothing Or lo.ListRows.Count = 0 Then Err.Raise 1004, , "Jira_Facts empty"

    Dim ws As Worksheet: Set ws = EnsureSheet("Jira_Insights")
    ws.Cells.Clear
    ws.Range("A1").Value = "Jira Insights"
    ws.Range("A1").Font.Bold = True

    ' Per-story-point completion time statistics (Done only)
    Dim startStats As Range: Set startStats = ws.Range("A3")
    startStats.Offset(0, 0).Value = "Story Points"
    startStats.Offset(0, 1).Value = "Avg Days"
    startStats.Offset(0, 2).Value = "StDev Days"
    startStats.Offset(0, 3).Value = "Count"
    ws.Range(ws.Cells(startStats.Row, 1), ws.Cells(startStats.Row, 4)).Font.Bold = True
    Call WritePerPointTimeStats(lo, startStats.Offset(1, 0))
    ' Chart for Avg Days by Story Points
    Dim lastRowStats As Long
    lastRowStats = ws.Cells(ws.Rows.Count, startStats.Column).End(xlUp).Row
    Dim ch0 As ChartObject
    Set ch0 = ws.ChartObjects.Add(Left:=400, Top:=ws.Range("A3").Top, Width:=420, Height:=260)
    ch0.Chart.ChartType = xlColumnClustered
    ch0.Chart.SetSourceData ws.Range(ws.Cells(startStats.Row, startStats.Column), ws.Cells(lastRowStats, startStats.Column + 2))
    ch0.Chart.HasTitle = True
    ch0.Chart.ChartTitle.Text = "Avg Days by Story Points"

    ' Cycle Time Analysis (calendar days)
    Dim meanCT As Double, medCT As Double, sdCT As Double, outCT As Long
    Call ComputeCycleCalendarStats(lo, meanCT, medCT, sdCT, outCT)
    Dim cycStart As Long: cycStart = lastRowStats + 2
    ws.Cells(cycStart, 1).Value = "Cycle Time (days)"
    ws.Cells(cycStart, 1).Font.Bold = True
    ws.Cells(cycStart + 1, 1).Value = "Mean"
    ws.Cells(cycStart + 1, 2).Value = Round(meanCT, 2)
    ws.Cells(cycStart + 2, 1).Value = "Median"
    ws.Cells(cycStart + 2, 2).Value = Round(medCT, 2)
    ws.Cells(cycStart + 3, 1).Value = "StDev"
    ws.Cells(cycStart + 3, 2).Value = Round(sdCT, 2)
    ws.Cells(cycStart + 4, 1).Value = "Outliers (> mean + 2*stdev)"
    ws.Cells(cycStart + 4, 2).Value = outCT

    ' Bottleneck Detection (calendar days)
    Dim avgWait As Double, avgExec As Double
    Call ComputeBottlenecks(lo, avgWait, avgExec)
    ws.Cells(cycStart + 6, 1).Value = "Bottlenecks"
    ws.Cells(cycStart + 6, 1).Font.Bold = True
    ws.Cells(cycStart + 7, 1).Value = "Avg To Do → In Progress (days)"
    ws.Cells(cycStart + 7, 2).Value = IIf(avgWait > 0, Round(avgWait, 2), "N/A")
    ws.Cells(cycStart + 8, 1).Value = "Avg In Progress → Done (days)"
    ws.Cells(cycStart + 8, 2).Value = IIf(avgExec > 0, Round(avgExec, 2), "N/A")

    ' Summary replacement: Average Cycle Time by Story Points (1,2,3,5,8,13)
    Dim thr As Range: Set thr = startStats.Offset(0, 6) ' move further right to avoid overlap
    thr.Value = "Cycle Time by Story Points"
    thr.Font.Bold = True
    thr.Offset(1, 0).Value = "Story Points"
    thr.Offset(1, 1).Value = "Avg Days"
    thr.Offset(1, 0).Resize(1, 2).Font.Bold = True

    Dim idxSP2 As Long, idxCycle2 As Long
    On Error Resume Next
    idxSP2 = lo.ListColumns("StoryPoints").Index
    idxCycle2 = lo.ListColumns("CycleDays").Index
    On Error GoTo 0
    If idxSP2 > 0 And idxCycle2 > 0 Then
        Dim cats As Variant: cats = Array(1, 2, 3, 5, 8, 13)
        Dim sums2 As Object: Set sums2 = CreateObject("Scripting.Dictionary")
        Dim counts2 As Object: Set counts2 = CreateObject("Scripting.Dictionary")
        Dim i2 As Long
        For i2 = 1 To lo.ListRows.Count
            Dim sp2 As Long: sp2 = CLng(Val(lo.DataBodyRange.Cells(i2, idxSP2).Value))
            Dim days2 As Double: days2 = Val(lo.DataBodyRange.Cells(i2, idxCycle2).Value)
            If days2 > 0 Then
                Dim j As Long
                For j = LBound(cats) To UBound(cats)
                    If sp2 = CLng(cats(j)) Then
                        If Not sums2.Exists(sp2) Then sums2(sp2) = 0#: counts2(sp2) = 0
                        sums2(sp2) = CDbl(sums2(sp2)) + days2
                        counts2(sp2) = CLng(counts2(sp2)) + 1
                        Exit For
                    End If
                Next j
            End If
        Next i2
        Dim row2 As Long: row2 = thr.Row + 2
        For i2 = LBound(cats) To UBound(cats)
            ws.Cells(row2, thr.Column).Value = cats(i2)
            If counts2.Exists(CLng(cats(i2))) And CLng(counts2(CLng(cats(i2)))) > 0 Then
                ws.Cells(row2, thr.Column + 1).Value = Round(CDbl(sums2(CLng(cats(i2)))) / CLng(counts2(CLng(cats(i2)))), 2)
            Else
                ws.Cells(row2, thr.Column + 1).Value = "N/A"
            End If
            row2 = row2 + 1
        Next i2
    End If

    ' Build pivots (Epic summary removed by request)
    Dim pc As PivotCache
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=lo.Range)

    ' Pivot 1 (renumbered): Story Point Distribution (rows SP, count issues)
    Dim pt2 As PivotTable
    Dim rowStart As Long
    rowStart = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 2
    Set pt2 = ws.PivotTables.Add(PivotCache:=pc, TableDestination:=ws.Cells(rowStart, 1), TableName:="ptSPDist")
    With pt2
        On Error Resume Next
        .PivotFields("StoryPoints").Orientation = xlRowField
        .PivotFields("IssueKey").Orientation = xlDataField
        .PivotFields("IssueKey").Function = xlCount
        On Error GoTo 0
    End With
    Dim ch2 As ChartObject
    Set ch2 = ws.ChartObjects.Add(Left:=400, Top:=ws.Cells(rowStart, 1).Top, Width:=420, Height:=260)
    ch2.Chart.ChartType = xlColumnClustered
    ch2.Chart.SetSourceData pt2.TableRange2
    ch2.Chart.HasTitle = True
    ch2.Chart.ChartTitle.Text = "Story Point Distribution"

    ' Pivot 2: Quarter summary
    Dim pt3 As PivotTable
    rowStart = pt2.TableRange2.Row + pt2.TableRange2.Rows.Count + 2
    Set pt3 = ws.PivotTables.Add(PivotCache:=pc, TableDestination:=ws.Cells(rowStart, 1), TableName:="ptQuarter")
    With pt3
        On Error Resume Next
        .PivotFields("QuarterTag").Orientation = xlRowField
        .PivotFields("StoryPoints").Orientation = xlDataField
        .PivotFields("StoryPoints").Function = xlSum
        .PivotFields("IssueKey").Orientation = xlDataField
        .PivotFields("IssueKey").Function = xlCount
        On Error GoTo 0
    End With

    ws.Columns("A:K").AutoFit
End Sub

Private Sub WritePerPointTimeStats(ByVal lo As ListObject, ByVal dest As Range)
    Dim idxSP As Long, idxCycle As Long
    On Error Resume Next
    idxSP = lo.ListColumns("StoryPoints").Index
    idxCycle = lo.ListColumns("CycleDays").Index
    On Error GoTo 0
    If idxSP = 0 Or idxCycle = 0 Then Exit Sub

    Dim allSP As Object: Set allSP = CreateObject("Scripting.Dictionary")
    Dim stats As Object: Set stats = CreateObject("Scripting.Dictionary")
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
    Dim sumsSq As Object: Set sumsSq = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        Dim sp As Long: sp = CLng(Val(lo.DataBodyRange.Cells(i, idxSP).Value))
        If sp > 0 Then
            If Not allSP.Exists(sp) Then allSP(sp) = True
            Dim days As Double: days = Val(lo.DataBodyRange.Cells(i, idxCycle).Value)
            If days > 0 Then
                If Not stats.Exists(sp) Then
                    stats(sp) = 0: sums(sp) = 0#: sumsSq(sp) = 0#
                End If
                stats(sp) = CLng(stats(sp)) + 1
                sums(sp) = CDbl(sums(sp)) + days
                sumsSq(sp) = CDbl(sumsSq(sp)) + days * days
            End If
        End If
    Next i

    ' Use all distinct SP values (even if no resolved items yet)
    Dim keys() As Variant: keys = allSP.Keys
    Dim j As Long, k As Long
    For j = LBound(keys) To UBound(keys) - 1
        For k = j + 1 To UBound(keys)
            If CLng(keys(k)) < CLng(keys(j)) Then
                Dim tmp As Variant: tmp = keys(j): keys(j) = keys(k): keys(k) = tmp
            End If
        Next k
    Next j

    Dim row As Long: row = dest.Row
    For j = LBound(keys) To UBound(keys)
        Dim spv As Long: spv = CLng(keys(j))
        Dim n As Long: If stats.Exists(spv) Then n = CLng(stats(spv)) Else n = 0
        Dim s As Double: If sums.Exists(spv) Then s = CDbl(sums(spv)) Else s = 0#
        Dim ss As Double: If sumsSq.Exists(spv) Then ss = CDbl(sumsSq(spv)) Else ss = 0#
        Dim avg As Double: If n > 0 Then avg = s / n
        Dim sd As Double: If n > 1 Then sd = Sqr((ss - (s * s) / n) / (n - 1))
        dest.Worksheet.Cells(row, dest.Column + 0).Value = spv
        dest.Worksheet.Cells(row, dest.Column + 1).Value = Round(avg, 2)
        dest.Worksheet.Cells(row, dest.Column + 2).Value = Round(sd, 2)
        dest.Worksheet.Cells(row, dest.Column + 3).Value = n
        row = row + 1
    Next j
End Sub

Public Sub Jira_NormalizeIssues_FromSample()
    Jira_NormalizeIssues "Jira_Issues_Sample", "tblJiraIssuesSample"
End Sub

Public Sub SanitizeRawAndBuildInsights()
    On Error GoTo Fail
    LogStart "SanitizeRawAndBuildInsights"

    Dim srcSheet As String, srcTable As String
    Dim v As Variant
    v = Application.InputBox( _
        Prompt:="Enter source sheet name (e.g., Raw_Data or WIP_Facts)", _
        Title:="Sanitize Source", Type:=2)
    If VarType(v) = vbBoolean And v = False Then GoTo CancelOp
    srcSheet = Trim$(CStr(v))
    If Len(srcSheet) = 0 Then GoTo CancelOp

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Trim$(srcSheet))
    On Error GoTo 0
    If ws Is Nothing Then
        ' Offer to create known samples by name
        If StrComp(Trim$(srcSheet), "Raw_Data", vbTextCompare) = 0 Then
            EnsureRawDataSheet
            Set ws = ThisWorkbook.Worksheets("Raw_Data")
        ElseIf StrComp(Trim$(srcSheet), "WIP_Facts", vbTextCompare) = 0 Then
            Dim w As Worksheet: Set w = EnsureSheet("WIP_Facts")
            Call EnsureWIPFactsTable(w)
            Set ws = w
        End If
        If ws Is Nothing Then Err.Raise 9, , "Sheet not found: " & srcSheet
    End If

    ' Resolve table name
    If ws.ListObjects.Count = 1 Then
        srcTable = ws.ListObjects(1).Name
    Else
        ' Try common table names first
        Dim tryName As Variant, lo As ListObject
        For Each tryName In Array("tblRawData", "tblWIPFacts", "tblJiraFacts")
            On Error Resume Next
            Set lo = ws.ListObjects(CStr(tryName))
            On Error GoTo 0
            If Not lo Is Nothing Then srcTable = lo.Name: Exit For
        Next tryName
        If Len(srcTable) = 0 Then
            v = Application.InputBox( _
                Prompt:="Enter table name on sheet '" & ws.Name & "' (ListObject name)", _
                Title:="Sanitize Source Table", Type:=2)
            If VarType(v) = vbBoolean And v = False Then GoTo CancelOp
            srcTable = Trim$(CStr(v))
            If Len(srcTable) = 0 Then GoTo CancelOp
        End If
    End If

    ' If the selected table looks like WIP (time-in-status), skip Jira normalization and build flow charts
    Dim loSrc As ListObject, isWip As Boolean
    On Error Resume Next: Set loSrc = ws.ListObjects(srcTable): On Error GoTo 0
    If Not loSrc Is Nothing Then
        isWip = Flow_HasColumn(loSrc, "TimeInTodo") Or Flow_HasColumn(loSrc, "TimeInProgress") Or _
                Flow_HasColumn(loSrc, "TimeInTesting") Or Flow_HasColumn(loSrc, "TimeInReview")
    End If

    If isWip Then
        ' Build Flow metrics from WIP-like table only
        On Error Resume Next
        Flow_BuildCharts
        On Error GoTo 0
    Else
        ' Normalize and build insights from selected sheet/table
        Jira_NormalizeIssues ws.Name, srcTable
        Jira_CreatePivotsAndCharts
        ' Also build Flow Metrics charts (CFD, Throughput, Cycle Scatter)
        On Error Resume Next
        Flow_BuildCharts
        On Error GoTo 0
    End If
    LogOk "SanitizeRawAndBuildInsights"
    If IsVerbose() Then MsgBox "Sanitized '" & ws.Name & "' (" & srcTable & ") and updated insights.", vbInformation
    Exit Sub

CancelOp:
    LogErr "SanitizeRawAndBuildInsights", "Cancelled by user"
    Exit Sub

Fail:
    LogErr "SanitizeRawAndBuildInsights", "Err " & Err.Number & ": " & Err.Description
    MsgBox "SanitizeRawAndBuildInsights failed: " & Err.Description, vbExclamation
End Sub

Public Sub RefreshSamples()
    On Error GoTo Fail
    LogStart "RefreshSamples"
    EnsureSampleIssuesSheet
    EnsureRawSampleSheet
    EnsureSampleIssuesSheetExpanded
    LogOk "RefreshSamples"
    If IsVerbose() Then MsgBox "Sample sheets regenerated.", vbInformation
    Exit Sub
Fail:
    LogErr "RefreshSamples", "Err " & Err.Number & ": " & Err.Description
    MsgBox "RefreshSamples failed: " & Err.Description, vbExclamation
End Sub

Private Sub EnsureSampleIssuesSheet()
    Dim ws As Worksheet: Set ws = EnsureSheet("Jira_Issues_Sample")
    Dim lo As ListObject
    On Error Resume Next: Set lo = ws.ListObjects("tblJiraIssuesSample"): On Error GoTo 0
    If Not lo Is Nothing Then
        If lo.ListRows.Count >= 40 Then Exit Sub ' already populated
    End If
    ws.Cells.Clear
    Dim headers As Variant
    headers = Array("Summary","Issue key","Issue id","Issue Type","Status","Created","Resolved","Fix Version/s","Custom field (Epic Link)","Custom field (Story Points)")
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    Dim base As Date: base = DateSerial(2025, 9, 10)
    Dim ep As Long, sp As Variant, t As String, st As String
    Dim r As Long: r = 2
    Dim spPool As Variant
    spPool = Array(1,1,1,1,1,1, 2,2,2,2,2,2,2,2, 3,3,3,3,3,3,3,3,3,3, 5,5,5,5,5,5,5,5, 8,8,8,8,8, 13,13,13)
    For i = 1 To 40
        t = IIf(i Mod 10 = 0, "Bug", IIf(i Mod 5 = 0, "Task", "Story"))
        ' Majority Done, some In Progress/To Do
        If i Mod 10 = 0 Or i Mod 13 = 0 Then
            st = "To Do"
        ElseIf i Mod 6 = 0 Or i Mod 7 = 0 Then
            st = "In Progress"
        Else
            st = "Done"
        End If
        ep = 100 + ((i - 1) Mod 6)
        sp = spPool(i - 1)
        Dim created As Date: created = base + ((i - 1) * 2 Mod 70)
        Dim startProg As Variant, resolved As Variant
        If st = "Done" Or sp = 13 Then
            st = "Done"
            Dim baseDays As Long: baseDays = sp * 2 + ((i Mod 5) - 2) ' +/- 2 days variability
            If i Mod 12 = 0 Then baseDays = baseDays + sp ' overrun outlier
            If i Mod 11 = 0 Then baseDays = baseDays - sp ' underrun outlier
            If baseDays < 1 Then baseDays = 1
            Dim startOffset As Long: startOffset = Application.WorksheetFunction.Max(1, baseDays \ 3)
            startProg = created + startOffset
            resolved = created + baseDays
        Else
            startProg = ""
            resolved = ""
        End If
        Dim fixv As String
        If Month(created) = 9 Then
            fixv = "2025.09.23"
        ElseIf Month(created) = 10 Then
            fixv = "2025.10.07"
        Else
            fixv = "2025.10.21"
        End If
        Call W(ws, r, Array( _
            "Sample Work Item " & i, _
            "FIINT-" & CStr(4000 + i), _
            330000 + i * 10, _
            t, _
            st, _
            created, _
            startProg, _
            resolved, _
            fixv, _
            "EPIC-" & ep, _
            sp))
        r = r + 1
    Next i
    ws.Rows(1).Font.Bold = True
    Set lo = Nothing
    On Error Resume Next: Set lo = ws.ListObjects("tblJiraIssuesSample"): On Error GoTo 0
    If Not lo Is Nothing Then lo.Delete
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    lo.Name = "tblJiraIssuesSample"
End Sub

Private Sub EnsureRawSampleSheet()
    ' Create a verbose raw sheet simulating a Jira export with many columns
    Dim ws As Worksheet: Set ws = EnsureSheet("Jira_Raw")
    Dim lo As ListObject
    On Error Resume Next: Set lo = ws.ListObjects("tblJiraRaw"): On Error GoTo 0
    If Not lo Is Nothing Then
        If lo.ListRows.Count >= 40 Then Exit Sub
    End If

    ws.Cells.Clear
    Dim hdrStr As String
    hdrStr = "Summary|Issue key|Issue id|Issue Type|Status|Priority|Assignee|Reporter|" & _
             "Created|Start Progress|Resolved|Fix Version/s|Affects Version/s|Component/s|Labels|" & _
             "Custom field (Epic Link)|Custom field (Story Points)|Custom field (Target Date)|URL|Extra1|Extra2"
    Dim headers As Variant: headers = Split(hdrStr, "|")
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

    Dim base As Date: base = DateSerial(2025, 9, 10)
    Dim spPool As Variant
    spPool = Array(1,1,1,1,1,1, 2,2,2,2,2,2,2,2, 3,3,3,3,3,3,3,3,3,3, 5,5,5,5,5,5,5,5, 8,8,8,8,8, 13,13,13)
    Dim r As Long: r = 2
    For i = 1 To 40
        Dim t As String, st As String, pri As String
        t = IIf(i Mod 10 = 0, "Bug", IIf(i Mod 5 = 0, "Task", "Story"))
        pri = Choose(((i Mod 3) + 1), "Medium","High","Low")
        If i Mod 10 = 0 Or i Mod 13 = 0 Then
            st = "To Do"
        ElseIf i Mod 6 = 0 Or i Mod 7 = 0 Then
            st = "In Progress"
        Else
            st = "Done"
        End If
        Dim created As Date: created = base + ((i - 1) * 2 Mod 70)
        Dim sp As Variant: sp = spPool(i - 1)
        Dim startProg As Variant, resolved As Variant
        If st = "Done" Or sp = 13 Then
            st = "Done"
            Dim baseDays As Long: baseDays = sp * 2 + ((i Mod 5) - 2)
            If baseDays < 1 Then baseDays = 1
            Dim startOffset As Long: startOffset = Application.WorksheetFunction.Max(1, baseDays \ 3)
            startProg = created + startOffset
            resolved = created + baseDays
        Else
            startProg = ""
            resolved = ""
        End If
        Dim fixv As String
        If Month(created) = 9 Then
            fixv = "2025.09.23"
        ElseIf Month(created) = 10 Then
            fixv = "2025.10.07"
        Else
            fixv = "2025.10.21"
        End If

        ' Write row values directly to avoid line continuation limits
        ws.Cells(r, 1).Value = "Raw Item " & i
        ws.Cells(r, 2).Value = "FIINT-" & CStr(4400 + i)
        ws.Cells(r, 3).Value = 440000 + i * 15
        ws.Cells(r, 4).Value = t
        ws.Cells(r, 5).Value = st
        ws.Cells(r, 6).Value = pri
        ws.Cells(r, 7).Value = "User" & ((i Mod 5) + 1)
        ws.Cells(r, 8).Value = "Reporter" & ((i Mod 3) + 1)
        ws.Cells(r, 9).Value = created
        ws.Cells(r, 10).Value = startProg
        ws.Cells(r, 11).Value = resolved
        ws.Cells(r, 12).Value = fixv
        ws.Cells(r, 13).Value = "v1.0"
        ws.Cells(r, 14).Value = "Core"
        ws.Cells(r, 15).Value = IIf(i Mod 4 = 0, "urgent", "")
        ws.Cells(r, 16).Value = "EPIC-" & (100 + ((i - 1) Mod 6))
        ws.Cells(r, 17).Value = sp
        ws.Cells(r, 18).Value = created + 30
        ws.Cells(r, 19).Value = "https://jira.example/browse/FIINT-" & CStr(4400 + i)
        ws.Cells(r, 20).Value = "extraA"
        ws.Cells(r, 21).Value = "extraB"
        r = r + 1
    Next i
    ws.Rows(1).Font.Bold = True
    On Error Resume Next: Set lo = ws.ListObjects("tblJiraRaw"): On Error GoTo 0
    If Not lo Is Nothing Then lo.Delete
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    lo.Name = "tblJiraRaw"
End Sub

Private Sub EnsureSampleIssuesSheetExpanded()
    ' Deprecated in favor of EnsureRawDataSheet
End Sub

Private Sub EnsureRawDataSheet()
    ' Create a Raw_Data sheet including time-in-status durations and date-time fields
    Dim ws As Worksheet: Set ws = EnsureSheet("Raw_Data")
    Dim lo As ListObject
    On Error Resume Next: Set lo = ws.ListObjects("tblRawData"): On Error GoTo 0
    If Not lo Is Nothing Then
        If lo.ListRows.Count >= 40 Then Exit Sub
    End If

    ws.Cells.Clear
    Dim hdr As String
    hdr = "Summary|Issue key|Issue id|Issue Type|Status|Priority|Assignee|Reporter|" & _
          "Created|Start Progress|Updated|Resolved|" & _
          "Time In Todo|Time In Progress|Time In Testing|Time In Review|" & _
          "Fix Version/s|Component/s|Labels|" & _
          "Custom field (Epic Link)|Custom field (Story Points)|URL"
    Dim headers As Variant: headers = Split(hdr, "|")

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

    Dim base As Date: base = DateSerial(2025, 9, 10)
    Dim spPool As Variant
    spPool = Array(1,1,1,1,1,1, 2,2,2,2,2,2,2,2, 3,3,3,3,3,3,3,3,3,3, 5,5,5,5,5,5,5,5, 8,8,8,8,8, 13,13,13)

    Dim r As Long: r = 2
    For i = 1 To 40
        Dim t As String, st As String, pri As String
        t = IIf(i Mod 10 = 0, "Bug", IIf(i Mod 5 = 0, "Task", "Story"))
        pri = Choose(((i Mod 3) + 1), "Medium","High","Low")
        st = "Done"

        ' Base created with time-of-day
        Dim created As Date: created = base + ((i - 1) * 2 Mod 70) + TimeSerial((i * 2) Mod 24, (i * 7) Mod 60, 0)
        Dim sp As Variant: sp = spPool(i - 1)

        ' Time-in-status (days, decimals)
        Dim tTodo As Double, tProg As Double, tTest As Double, tRev As Double
        tTodo = Round(0.05 + ((i Mod 5) * 0.11), 2)
        tProg = Round(0.25 + ((sp Mod 8) * 0.35), 2)
        tTest = Round(((i Mod 4) * 0.12), 2)
        tRev = Round(((i Mod 3) * 0.10), 2)

        Dim startProg As Date: startProg = created + tTodo
        Dim updated As Date: updated = startProg + tProg
        Dim resolved As Date: resolved = created + tTodo + tProg + tTest + tRev

        ' Write values
        ws.Cells(r, 1).Value = "Raw Item " & i
        ws.Cells(r, 2).Value = "FIINT-" & CStr(5800 + i)
        ws.Cells(r, 3).Value = 580000 + i * 11
        ws.Cells(r, 4).Value = t
        ws.Cells(r, 5).Value = st
        ws.Cells(r, 6).Value = pri
        ws.Cells(r, 7).Value = "User" & ((i Mod 5) + 1)
        ws.Cells(r, 8).Value = "Reporter" & ((i Mod 3) + 1)
        ws.Cells(r, 9).Value = created
        ws.Cells(r,10).Value = startProg
        ws.Cells(r,11).Value = updated
        ws.Cells(r,12).Value = resolved
        ws.Cells(r,13).Value = tTodo
        ws.Cells(r,14).Value = tProg
        ws.Cells(r,15).Value = tTest
        ws.Cells(r,16).Value = tRev
        ws.Cells(r,17).Value = Choose(((Month(created)-8) Mod 3)+1, "2025.09.23","2025.10.07","2025.10.21")
        ws.Cells(r,18).Value = "Core"
        ws.Cells(r,19).Value = IIf(i Mod 4 = 0, "urgent", "")
        ws.Cells(r,20).Value = "EPIC-" & (100 + ((i - 1) Mod 6))
        ws.Cells(r,21).Value = sp
        ws.Cells(r,22).Value = "https://jira.example/browse/FIINT-" & CStr(5800 + i)
        r = r + 1
    Next i
    ws.Rows(1).Font.Bold = True
    On Error Resume Next: Set lo = ws.ListObjects("tblRawData"): On Error GoTo 0
    If Not lo Is Nothing Then lo.Delete
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    lo.Name = "tblRawData"
End Sub

Public Sub Jira_NormalizeIssues(ByVal rawSheet As String, ByVal rawTable As String)
    On Error GoTo Fail
    Dim wsRaw As Worksheet: Set wsRaw = Worksheets(rawSheet)
    Dim loRaw As ListObject: Set loRaw = wsRaw.ListObjects(rawTable)
    Dim map As Object: Set map = Jira_BuildHeaderMap(loRaw)
    If map Is Nothing Then Err.Raise 1004, , "Could not map required Jira columns"

    Dim ws As Worksheet: Set ws = EnsureSheet("Jira_Facts")
    Dim lo As ListObject
    On Error Resume Next: Set lo = ws.ListObjects("tblJiraFacts"): On Error GoTo 0
    If lo Is Nothing Then
        Dim headers As Variant
        headers = Array("IssueKey","Summary","IssueType","Status","Epic","Created","StartProgress","Resolved","StoryPoints","CycleDays","SprintSpan","IsCrossSprint","QuarterTag","YearTag","FixVersion","CreatedMonth","CycleCalDays","LeadCalDays")
        Set lo = EnsureTable(ws, "tblJiraFacts", headers)
    Else
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    End If

    Dim r As ListRow, out As ListRow
    Dim created As Date, startProg As Date, resolved As Date, d As Double, span As Long, sp As Double
    Dim sprintLen As Long: sprintLen = CLng(Val(GetNameValueOr("SprintLengthDays", "10")))
    For Each r In loRaw.ListRows
        Set out = lo.ListRows.Add
        out.Range(1, 1).Value = GetCellBy(loRaw, r, map, "IssueKey")
        out.Range(1, 2).Value = GetCellBy(loRaw, r, map, "Summary")
        out.Range(1, 3).Value = GetCellBy(loRaw, r, map, "IssueType")
        out.Range(1, 4).Value = GetCellBy(loRaw, r, map, "Status")
        out.Range(1, 5).Value = GetCellBy(loRaw, r, map, "Epic")
        created = ToDateSafe(GetCellBy(loRaw, r, map, "Created"))
        out.Range(1, 6).Value = created
        startProg = ToDateSafe(GetCellBy(loRaw, r, map, "StartProgress"))
        out.Range(1, 7).Value = startProg
        resolved = ToDateSafe(GetCellBy(loRaw, r, map, "Resolved"))
        out.Range(1, 8).Value = resolved
        sp = Val(GetCellBy(loRaw, r, map, "StoryPoints"))
        out.Range(1, 9).Value = sp
        d = WorkdaysBetween(created, IIf(resolved = 0, Date, resolved))
        out.Range(1, 10).Value = d
        span = Application.WorksheetFunction.RoundUp(d / sprintLen, 0)
        out.Range(1, 11).Value = span
        out.Range(1, 12).Value = (span > 1)
        Dim refDate As Date: refDate = IIf(resolved = 0, created, resolved)
        out.Range(1, 13).Value = Year(refDate) & " Q" & Int((Month(refDate) - 1) / 3 + 1)
        out.Range(1, 14).Value = Year(refDate)
        ' FixVersion and CreatedMonth
        out.Range(1, 15).Value = GetCellBy(loRaw, r, map, "FixVersion")
        If created <> 0 Then out.Range(1, 16).Value = DateSerial(Year(created), Month(created), 1)
        ' Calendar day metrics
        If created <> 0 And resolved <> 0 Then
            out.Range(1, 17).Value = DateDiff("d", created, resolved)
            If startProg <> 0 Then out.Range(1, 18).Value = DateDiff("d", created, resolved) ' lead time same; cycle from start can be DateDiff("d", startProg, resolved)
        End If
    Next r

    If IsVerbose() Then MsgBox "Jira facts built from " & rawTable & ".", vbInformation
    Exit Sub
Fail:
    MsgBox "Normalize issues failed: " & Err.Description, vbExclamation
End Sub

Private Function Jira_BuildHeaderMap(ByVal lo As ListObject) As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Dim names As Object: Set names = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        names(Norm(lo.ListColumns(i).Name)) = i
    Next i
    ' map keys
    Call MapCol(idx, names, "IssueKey", Array("issue key","key"))
    Call MapCol(idx, names, "Summary", Array("summary","title"))
    Call MapCol(idx, names, "IssueType", Array("issue type","type"))
    Call MapCol(idx, names, "Status", Array("status"))
    Call MapCol(idx, names, "Epic", Array("epic link","epic"))
    Call MapCol(idx, names, "Created", Array("created","created date","created on"))
    Call MapCol(idx, names, "Resolved", Array("resolved","resolved date","done date"))
    Call MapCol(idx, names, "StoryPoints", Array("story points","story point","story point estimate","custom field (story points)"))
    Call MapCol(idx, names, "FixVersion", Array("fix version/s","fix version"))
    Call MapCol(idx, names, "StartProgress", Array("start progress","started","in progress","in progress date","start date"))
    If idx.Count = 0 Then Set idx = Nothing
    Set Jira_BuildHeaderMap = idx
End Function

Private Sub MapCol(ByVal idx As Object, ByVal names As Object, ByVal key As String, ByVal candidates As Variant)
    Dim j As Long
    For j = LBound(candidates) To UBound(candidates)
        Dim k As String: k = Norm(CStr(candidates(j)))
        Dim col As Variant
        col = FindByContains(names, k)
        If Not IsEmpty(col) Then idx(key) = CLng(col): Exit Sub
    Next j
End Sub

Private Function FindByContains(ByVal names As Object, ByVal needle As String) As Variant
    Dim k As Variant
    For Each k In names.Keys
        If InStr(1, k, needle, vbTextCompare) > 0 Then
            FindByContains = names(k): Exit Function
        End If
    Next k
End Function

Private Function Norm(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace$(s, "_", " ")
    s = Replace$(s, "/", " ")
    s = Replace$(s, "-", " ")
    s = Replace$(s, "(" , " ")
    s = Replace$(s, ")" , " ")
    s = Replace$(s, ":" , " ")
    s = Application.WorksheetFunction.Trim(s)
    Norm = s
End Function

Private Function GetCellBy(ByVal lo As ListObject, ByVal r As ListRow, ByVal map As Object, ByVal key As String) As String
    On Error Resume Next
    Dim idx As Long: idx = map(key)
    If idx > 0 Then GetCellBy = CStr(r.Range(1, idx).Value)
End Function

Private Function ToDateSafe(ByVal v As Variant) As Date
    On Error Resume Next
    If IsDate(v) Then ToDateSafe = CDate(v)
End Function

Private Function WorkdaysBetween(ByVal d1 As Date, ByVal d2 As Date) As Long
    If d1 = 0 Or d2 = 0 Then Exit Function
    Dim s As Date, e As Date
    If d2 < d1 Then s = d2: e = d1 Else s = d1: e = d2
    Dim n As Long: n = 0
    Do While s <= e
        Dim w As VbDayOfWeek: w = Weekday(s, vbSunday)
        If w >= vbMonday And w <= vbFriday Then n = n + 1
        s = s + 1
    Loop
    WorkdaysBetween = n
End Function

Private Sub ComputeSummaryMetrics(ByVal lo As ListObject, ByRef avgCycle As Double, ByRef totalItems As Long, ByRef avgSPPerSprint As Double)
    Dim idxCycle As Long, idxResolved As Long, idxSP As Long, idxFix As Long
    On Error Resume Next
    idxCycle = lo.ListColumns("CycleDays").Index
    idxResolved = lo.ListColumns("Resolved").Index
    idxSP = lo.ListColumns("StoryPoints").Index
    idxFix = lo.ListColumns("FixVersion").Index
    On Error GoTo 0
    If idxCycle = 0 Or idxSP = 0 Then Exit Sub

    Dim sumDays As Double, nDays As Long
    Dim fixSums As Object: Set fixSums = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        Dim days As Double: days = Val(lo.DataBodyRange.Cells(i, idxCycle).Value)
        If days > 0 Then
            sumDays = sumDays + days
            nDays = nDays + 1
        End If
        Dim sp As Double: sp = Val(lo.DataBodyRange.Cells(i, idxSP).Value)
        Dim fx As String
        If idxFix > 0 Then fx = CStr(lo.DataBodyRange.Cells(i, idxFix).Value)
        If sp > 0 And Len(fx) > 0 And days > 0 Then
            If Not fixSums.Exists(fx) Then fixSums(fx) = 0#
            fixSums(fx) = CDbl(fixSums(fx)) + sp
        End If
    Next i
    totalItems = lo.ListRows.Count
    If nDays > 0 Then avgCycle = sumDays / nDays
    If fixSums.Count > 0 Then
        Dim sumSP As Double, k As Variant
        For Each k In fixSums.Keys
            sumSP = sumSP + CDbl(fixSums(k))
        Next k
        avgSPPerSprint = sumSP / fixSums.Count
    End If
End Sub

Private Sub ComputeCycleCalendarStats(ByVal lo As ListObject, ByRef meanCT As Double, ByRef medCT As Double, ByRef sdCT As Double, ByRef outCount As Long)
    Dim idx As Long
    On Error Resume Next
    idx = lo.ListColumns("CycleCalDays").Index
    On Error GoTo 0
    If idx = 0 Then Exit Sub
    Dim vals() As Double
    Dim n As Long: n = 0
    Dim i As Long, v As Double
    For i = 1 To lo.ListRows.Count
        v = Val(lo.DataBodyRange.Cells(i, idx).Value)
        If v > 0 Then
            n = n + 1
            ReDim Preserve vals(1 To n)
            vals(n) = v
        End If
    Next i
    If n = 0 Then Exit Sub
    Dim wf As WorksheetFunction
    Set wf = Application.WorksheetFunction
    meanCT = wf.Average(vals)
    medCT = wf.Median(vals)
    If n > 1 Then sdCT = wf.StDev_S(vals)
    Dim thr As Double: thr = meanCT + 2 * sdCT
    For i = 1 To n
        If vals(i) > thr Then outCount = outCount + 1
    Next i
End Sub

Private Sub ComputeConsistencyPredictability(ByVal lo As ListObject, ByRef velocityStdev As Double, ByRef predictability As Variant)
    ' velocity stdev based on completed SP by FixVersion in facts
    Dim idxSP As Long, idxFix As Long
    On Error Resume Next
    idxSP = lo.ListColumns("StoryPoints").Index
    idxFix = lo.ListColumns("FixVersion").Index
    On Error GoTo 0
    If idxSP > 0 And idxFix > 0 Then
        Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
        Dim i As Long
        For i = 1 To lo.ListRows.Count
            Dim fx As String: fx = CStr(lo.DataBodyRange.Cells(i, idxFix).Value)
            Dim sp As Double: sp = Val(lo.DataBodyRange.Cells(i, idxSP).Value)
            If Len(fx) > 0 And sp > 0 Then
                If Not sums.Exists(fx) Then sums(fx) = 0#
                sums(fx) = CDbl(sums(fx)) + sp
            End If
        Next i
        If sums.Count > 1 Then
            Dim vals() As Double: ReDim vals(1 To sums.Count)
            Dim k As Variant, j As Long: j = 0
            For Each k In sums.Keys
                j = j + 1: vals(j) = CDbl(sums(k))
            Next k
            velocityStdev = Application.WorksheetFunction.StDev_S(vals)
        End If
    End If
    ' predictability from Jira_Metrics if present
    predictability = Empty
    On Error Resume Next
    Dim ws As Worksheet: Set ws = Worksheets("Jira_Metrics")
    Dim loM As ListObject: Set loM = ws.ListObjects("tblJiraMetrics")
    On Error GoTo 0
    If Not loM Is Nothing Then
        Dim idxC As Long, idxK As Long, sumC As Double, sumK As Double
        On Error Resume Next
        idxC = loM.ListColumns("Completed").Index
        idxK = loM.ListColumns("Committed").Index
        On Error GoTo 0
        If idxC > 0 And idxK > 0 Then
            Dim r As ListRow
            For Each r In loM.ListRows
                sumC = sumC + Val(r.Range(1, idxC).Value)
                sumK = sumK + Val(r.Range(1, idxK).Value)
            Next r
            If sumK > 0 Then predictability = sumC / sumK
        End If
    End If
End Sub

Private Function ComputeCorrelationSPCycle(ByVal lo As ListObject) As Variant
    Dim idxSP As Long, idxCycle As Long
    On Error Resume Next
    idxSP = lo.ListColumns("StoryPoints").Index
    idxCycle = lo.ListColumns("CycleDays").Index
    On Error GoTo 0
    If idxSP = 0 Or idxCycle = 0 Then Exit Function
    Dim n As Long: n = 0
    Dim sumX As Double, sumY As Double, sumXY As Double, sumX2 As Double, sumY2 As Double
    Dim i As Long, x As Double, y As Double
    For i = 1 To lo.ListRows.Count
        x = Val(lo.DataBodyRange.Cells(i, idxSP).Value)
        y = Val(lo.DataBodyRange.Cells(i, idxCycle).Value)
        If x > 0 And y > 0 Then
            n = n + 1
            sumX = sumX + x
            sumY = sumY + y
            sumXY = sumXY + x * y
            sumX2 = sumX2 + x * x
            sumY2 = sumY2 + y * y
        End If
    Next i
    If n < 2 Then Exit Function
    Dim num As Double, den As Double
    num = n * sumXY - sumX * sumY
    den = Sqr((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY))
    If den = 0 Then Exit Function
    ComputeCorrelationSPCycle = num / den
End Function

Private Sub ComputeBottlenecks(ByVal lo As ListObject, ByRef avgWait As Double, ByRef avgExec As Double)
    Dim idxCreated As Long, idxStart As Long, idxResolved As Long
    On Error Resume Next
    idxCreated = lo.ListColumns("Created").Index
    idxStart = lo.ListColumns("StartProgress").Index
    idxResolved = lo.ListColumns("Resolved").Index
    On Error GoTo 0
    If idxCreated = 0 Or idxResolved = 0 Then Exit Sub
    Dim sumWait As Double, cntWait As Long
    Dim sumExec As Double, cntExec As Long
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        Dim c As Variant, s As Variant, r As Variant
        c = lo.DataBodyRange.Cells(i, idxCreated).Value
        s = IIf(idxStart > 0, lo.DataBodyRange.Cells(i, idxStart).Value, 0)
        r = lo.DataBodyRange.Cells(i, idxResolved).Value
        If IsDate(c) And IsDate(r) Then
            If IsDate(s) Then
                sumWait = sumWait + DateDiff("d", CDate(c), CDate(s))
                cntWait = cntWait + 1
                sumExec = sumExec + DateDiff("d", CDate(s), CDate(r))
                cntExec = cntExec + 1
            End If
        End If
    Next i
    If cntWait > 0 Then avgWait = sumWait / cntWait
    If cntExec > 0 Then avgExec = sumExec / cntExec
End Sub

Public Sub Jira_CreateQueries()
    On Error GoTo Fail
    Dim qNameS As String: qNameS = "JiraSprints"
    Dim qNameF As String: qNameF = "JiraSprintReport"
    Dim qNameM As String: qNameM = "JiraSprintMetrics"

    Dim mS As String, mF As String, mM As String
    mS = _
        "let" & vbCrLf & _
        "  Base = Excel.CurrentWorkbook(){[Name=""JiraBaseUrl""]}[Content]{0}[Column1]," & vbCrLf & _
        "  Board = Text.From(Excel.CurrentWorkbook(){[Name=""JiraBoardId""]}[Content]{0}[Column1])," & vbCrLf & _
        "  Source = Json.Document(Web.Contents(Base, [RelativePath=""rest/agile/1.0/board/"" & Board & ""/sprint"", Query=[state=""active,closed"", maxResults=""50""]]))," & vbCrLf & _
        "  Values = if Record.HasFields(Source, ""values"") then Source[values] else if Record.HasFields(Source, ""sprints"") then Source[sprints] else {}," & vbCrLf & _
        "  T = Table.FromList(Values, Splitter.SplitByNothing(), null, null, ExtraValues.Error)," & vbCrLf & _
        "  E = Table.ExpandRecordColumn(T, ""Column1"", {""id"",""name"",""startDate"",""endDate"",""state""}, {""id"",""name"",""startDate"",""endDate"",""state""})" & vbCrLf & _
        "in" & vbCrLf & _
        "  E"

    mF = _
        "(sprintId as text) =>" & vbCrLf & _
        "let" & vbCrLf & _
        "  Base = Excel.CurrentWorkbook(){[Name=""JiraBaseUrl""]}[Content]{0}[Column1]," & vbCrLf & _
        "  Board = Text.From(Excel.CurrentWorkbook(){[Name=""JiraBoardId""]}[Content]{0}[Column1])," & vbCrLf & _
        "  R = Json.Document(Web.Contents(Base, [RelativePath=""rest/greenhopper/1.0/rapid/charts/sprintreport"", Query=[rapidViewId=Board, sprintId=sprintId]]))," & vbCrLf & _
        "  C = try R[contents] otherwise null," & vbCrLf & _
        "  committedInit = try C[completedIssuesInitialEstimateSum][value] otherwise 0," & vbCrLf & _
        "  notCompletedInit = try C[issuesNotCompletedInitialEstimateSum][value] otherwise 0," & vbCrLf & _
        "  completed = try C[completedIssuesEstimateSum][value] otherwise 0," & vbCrLf & _
        "  committed = committedInit + notCompletedInit," & vbCrLf & _
        "  T = #table({""sprintId"",""Committed"",""Completed""}, {{sprintId, committed, completed}})" & vbCrLf & _
        "in" & vbCrLf & _
        "  T"

    mM = _
        "let" & vbCrLf & _
        "  S = " & qNameS & "," & vbCrLf & _
        "  Add = Table.AddColumn(S, ""Metrics"", each " & qNameF & "(Text.From([id])))," & vbCrLf & _
        "  E = Table.ExpandTableColumn(Add, ""Metrics"", {""Committed"",""Completed""}, {""Committed"",""Completed""})" & vbCrLf & _
        "in" & vbCrLf & _
        "  E"

    Call AddOrUpdateQuery(qNameS, mS)
    Call AddOrUpdateQuery(qNameF, mF)
    Call AddOrUpdateQuery(qNameM, mM)

    ' Load JiraSprintMetrics to a sheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets("Jira_Metrics")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        ws.Name = "Jira_Metrics"
    Else
        ws.Cells.Clear
    End If

    Dim src As String
    src = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qNameM & ";Extended Properties="""""""

    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(SourceType:=0, Source:=src, Destination:=ws.Range("A1"))
    lo.Name = "tblJiraMetrics"
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & qNameM & "]")
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
    End With
    Exit Sub
Fail:
    MsgBox "Failed to create Power Query connections: " & Err.Description, vbExclamation
End Sub

Private Sub AddOrUpdateQuery(ByVal qName As String, ByVal mCode As String)
    On Error Resume Next
    Dim q As Object
    Set q = ActiveWorkbook.Queries(qName)
    On Error GoTo 0
    If Not q Is Nothing Then
        ' update
        ActiveWorkbook.Queries(qName).Formula = mCode
    Else
        ActiveWorkbook.Queries.Add Name:=qName, Formula:=mCode
    End If
End Sub

Public Sub Jira_ApplyMetricsFromQuery()
    On Error GoTo Fail
    Dim wsQ As Worksheet, wsM As Worksheet
    Set wsQ = Worksheets("Jira_Metrics")
    Set wsM = EnsureSheet("Metrics")
    Dim loQ As ListObject, loM As ListObject
    On Error Resume Next: Set loQ = wsQ.ListObjects("tblJiraMetrics"): On Error GoTo 0
    On Error Resume Next: Set loM = wsM.ListObjects("tblMetrics"): On Error GoTo 0
    If loQ Is Nothing Or loM Is Nothing Then Exit Sub

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As ListRow
    For Each r In loQ.ListRows
        Dim tag As String: tag = CStr(r.Range(1, loQ.ListColumns("name").Index).Value)
        Dim sid As Variant: sid = r.Range(1, loQ.ListColumns("id").Index).Value
        Dim committed As Double: committed = CDbl(r.Range(1, loQ.ListColumns("Committed").Index).Value)
        Dim completed As Double: completed = CDbl(r.Range(1, loQ.ListColumns("Completed").Index).Value)
        ' Prefer matching by tag (name might not match our tag) fallback: use startDate
        If Len(tag) > 0 Then dict(tag) = Array(committed, completed)
    Next r
    ' Apply to Metrics by Sprint column
    For Each r In loM.ListRows
        Dim sprintTag As String: sprintTag = CStr(r.Range(1, loM.ListColumns("Sprint").Index).Value)
        If dict.Exists(sprintTag) Then
            Dim v: v = dict(sprintTag)
            r.Range(1, loM.ListColumns("Points Committed").Index).Value = v(0)
            r.Range(1, loM.ListColumns("Points Completed").Index).Value = v(1)
        End If
    Next r
    Exit Sub
Fail:
End Sub

Private Sub MetricsApplyJiraForSprint(ByVal sStart As Date, ByVal sEnd As Date, ByVal committed As Double, ByVal completed As Double)
    Dim ws As Worksheet
    Set ws = EnsureSheet("Metrics")
    Dim lo As ListObject
    On Error Resume Next: Set lo = ws.ListObjects("tblMetrics"): On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    Dim tag As String: tag = FormatSprintTag(sStart)
    Dim r As ListRow, found As Boolean
    For Each r In lo.ListRows
        Dim cSprint As String: cSprint = CStr(r.Range(1, 2).Value)
        If StrComp(Trim$(cSprint), tag, vbTextCompare) = 0 Then
            ' timeframe if empty
            If Len(CStr(r.Range(1, 1).Value)) = 0 Then r.Range(1, 1).Value = Format$(sStart, "m/d") & " - " & Format$(sEnd, "m/d")
            ' write points
            r.Range(1, 4).Value = completed
            r.Range(1, 5).Value = committed
            found = True
            Exit For
        End If
    Next r
    If Not found Then
        ' append a row to metrics table
        Set r = lo.ListRows.Add
        r.Range(1, 1).Value = Format$(sStart, "m/d") & " - " & Format$(sEnd, "m/d")
        r.Range(1, 2).Value = tag
        r.Range(1, 4).Value = completed
        r.Range(1, 5).Value = committed
    End If
End Sub

"" ' token-based HTTP removed; using Power Query From Web instead

Private Sub JiraParseSprints(ByVal json As String, ByRef ids() As Long, ByRef starts() As Date, ByRef ends() As Date, ByRef outCount As Long)
    ' Lightweight parser for /board/{id}/sprint listing (assumes fields id/startDate/endDate)
    Dim p As Long: p = 1
    Dim cap As Long: cap = 0
    Do
        Dim idx As Long: idx = InStr(p, json, "{""id"":")
        If idx = 0 Then Exit Do
        Dim idVal As Long: idVal = CLng(ParseNumberAfter(json, idx + 6))
        Dim sStart As Date: sStart = ParseIsoDate(FindJsonString(json, idx, "startDate"))
        Dim sEnd As Date: sEnd = ParseIsoDate(FindJsonString(json, idx, "endDate"))
        cap = cap + 1
        ReDim Preserve ids(1 To cap): ReDim Preserve starts(1 To cap): ReDim Preserve ends(1 To cap)
        ids(cap) = idVal: starts(cap) = sStart: ends(cap) = sEnd
        p = idx + 8
    Loop
    outCount = cap
End Sub

Private Function JiraExtractEstimate(ByVal json As String, ByVal key As String) As Double
    Dim pat As String
    pat = "\"" & key & "\":{""value"":" ' not real regex; just find start
    Dim i As Long: i = InStr(1, json, "\"" & key & "\":{", vbTextCompare)
    If i = 0 Then Exit Function
    Dim j As Long: j = InStr(i, json, "\"value\":" , vbTextCompare)
    If j = 0 Then Exit Function
    JiraExtractEstimate = CDbl(ParseNumberAfter(json, j + 8))
End Function

Private Function ParseNumberAfter(ByVal s As String, ByVal pos As Long) As Double
    Dim i As Long: i = pos
    Do While i <= Len(s) And Mid$(s, i, 1) Like "[ \t\r\n:]"
        i = i + 1
    Loop
    Dim j As Long: j = i
    Do While j <= Len(s)
        Dim ch As String: ch = Mid$(s, j, 1)
        If Not (ch Like "[0-9.-]") Then Exit Do
        j = j + 1
    Loop
    If j > i Then ParseNumberAfter = Val(Mid$(s, i, j - i))
End Function

Private Function FindJsonString(ByVal s As String, ByVal startAt As Long, ByVal key As String) As String
    Dim i As Long: i = InStr(startAt, s, "\"" & key & "\":\"", vbTextCompare)
    If i = 0 Then Exit Function
    i = i + Len(key) + 4
    Dim j As Long: j = InStr(i, s, "\"")
    If j > i Then FindJsonString = Mid$(s, i, j - i)
End Function

Private Function ParseIsoDate(ByVal s As String) As Date
    On Error Resume Next
    If Len(s) = 0 Then Exit Function
    Dim t As String: t = s
    ' Expect formats like 2025-01-02T12:34:56.000+0000
    Dim y As Integer, m As Integer, d As Integer
    y = CInt(Left$(t, 4))
    m = CInt(Mid$(t, 6, 2))
    d = CInt(Mid$(t, 9, 2))
    ParseIsoDate = DateSerial(y, m, d)
End Function

Private Function B64Encode(ByVal plain As String) As String
    Dim bytes() As Byte: bytes = StrConv(plain, vbFromUnicode)
    Dim out As String
    Dim enc As String: enc = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim i As Long
    For i = 0 To UBound(bytes) Step 3
        Dim b0 As Long, b1 As Long, b2 As Long, n As Long
        b0 = bytes(i)
        If i + 1 <= UBound(bytes) Then b1 = bytes(i + 1) Else b1 = -1
        If i + 2 <= UBound(bytes) Then b2 = bytes(i + 2) Else b2 = -1
        If b1 >= 0 And b2 >= 0 Then
            n = (b0 And &HFF) * &H10000 + (b1 And &HFF) * &H100 + (b2 And &HFF)
            out = out & Mid$(enc, (n \ &H40000) + 1, 1)
            out = out & Mid$(enc, ((n And &H3F000) \ &H1000) + 1, 1)
            out = out & Mid$(enc, ((n And &HFC0) \ &H40) + 1, 1)
            out = out & Mid$(enc, (n And &H3F) + 1, 1)
        ElseIf b1 >= 0 Then
            n = (b0 And &HFF) * &H100 + (b1 And &HFF)
            out = out & Mid$(enc, ((n And &HFC00) \ &H400) + 1, 1)
            out = out & Mid$(enc, ((n And &H3F0) \ &H10) + 1, 1)
            out = out & Mid$(enc, ((n And &HF) * 4) + 1, 1)
            out = out & "="
        Else
            n = (b0 And &HFF)
            out = out & Mid$(enc, ((n And &HF0) \ &H10) + 1, 1)
            out = out & Mid$(enc, ((n And &HF) * 4) + 1, 1)
            out = out & "=="
        End If
    Next i
    B64Encode = out
End Function

Private Sub EnsureMetricsSheet()
    Dim ws As Worksheet: Set ws = EnsureSheet("Metrics")
    ' If table already exists, assume user has content and formatting—do not rebuild
    If HasTable(ws, "tblMetrics") Then Exit Sub

    ws.Cells.Clear

    ' Headers
    ws.Range("A1").Value = "Timeframe"
    ws.Range("B1").Value = "Sprint"
    ws.Range("C1").Value = "Days of Availability"
    ws.Range("D1").Value = "Points Completed"
    ws.Range("E1").Value = "Points Committed"
    ws.Range("F1").Value = "Velocity Per Available Day"
    ws.Range("G1").Value = "3-Sprint Average"
    ws.Range("H1").Value = "Comments 1"
    ws.Range("A1:H1").Font.Bold = True
    ws.Range("A1:H1").Interior.Color = RGB(198, 239, 206)
    ws.Range("A1:H1").Borders.Weight = 2

    ' Create a table for easier entry/formatting
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects("tblMetrics")
    On Error GoTo 0
    If Not lo Is Nothing Then lo.Delete
    Dim lastCol As Long: lastCol = 8
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, lastCol)), , xlYes)
    lo.Name = "tblMetrics"
    lo.TableStyle = "TableStyleMedium2"

    ' Pre-seed 32 sprint rows starting from current sprint
    Dim yr As Integer, q As Integer, s As Integer
    Dim tag As String: tag = FormatSprintTag(Date)
    If Not ParseTagFromName(tag, yr, q, s) Then
        yr = Year(Date): q = Int((Month(Date) - 1) / 3) + 1: s = 1
    End If

    Dim i As Long, r As Long: r = 2
    For i = 1 To 32
        Dim sStart As Date: sStart = QuarterStartDate(yr, q) + (s - 1) * 14
        Dim sEnd As Date: sEnd = DateAdd("d", 13, sStart)
        ws.Cells(r, 1).Value = Format$(sStart, "m/d") & " - " & Format$(sEnd, "m/d")
        ws.Cells(r, 2).Value = yr & " Q" & q & " S" & s
        ' F: Velocity per Day = Completed / Days
        ws.Cells(r, 6).FormulaR1C1 = "=IFERROR(RC[-2]/RC[-3],0)"
        ' G: 3-sprint moving average of F (including this row)
        ' Show N/A until three rows exist; then average last 3 velocities
        ws.Cells(r, 7).FormulaR1C1 = "=IF(ROW()<4,""N/A"",AVERAGE(R[-2]C[-1]:RC[-1]))"

        ' advance sprint counters
        s = s + 1
        If s > 7 Then s = 1: q = q + 1: If q > 4 Then q = 1: yr = yr + 1
        r = r + 1
    Next i

    ' Resize table over the filled rows
    Dim lastRow As Long: lastRow = r - 1
    lo.Resize ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Column widths similar to screenshot
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 22
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 24

    ' Freeze header row
    With ws
        .Activate
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With
End Sub
