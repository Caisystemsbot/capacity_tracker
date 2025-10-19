Attribute VB_Name = "modInstaller"
Option Explicit

' Paste this single module into any workbook (Insert > Module),
' then run InstallFromFolder to import all .bas/.cls/.frm from your repo.
'
' Requirements:
' - Excel Trust Center: enable "Trust access to the VBA project object model".
' - Repo layout: <chosen-root>\src\{modules,classes,forms}

Public Sub InstallFromFolder()
    Dim root As String
    If Not HasVBOMAccess() Then Exit Sub

    root = PickFolder("Select the repo root (containing 'src' folder)")
    If Len(root) = 0 Then Exit Sub

    Dim srcPath As String
    srcPath = PathJoin(root, "src")
    If Dir(srcPath, vbDirectory) = vbNullString Then
        MsgBox "Folder 'src' not found under: " & root, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableCancelKey = xlDisabled
    On Error GoTo CleanFail

    RemoveExistingCodeExcept Me.Name
    ImportFolder PathJoin(srcPath, "modules"), "*.bas"
    ImportFolder PathJoin(srcPath, "classes"), "*.cls"
    ImportFolder PathJoin(srcPath, "forms"), "*.frm"

    Application.ScreenUpdating = True
    Application.EnableCancelKey = xlInterrupt
    MsgBox "Import complete.", vbInformation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.EnableCancelKey = xlInterrupt
    MsgBox "Import failed: " & Err.Description, vbExclamation
End Sub

Private Function HasVBOMAccess() As Boolean
    On Error GoTo Denied
    Dim n$: n$ = ThisWorkbook.VBProject.Name
    HasVBOMAccess = True
    Exit Function
Denied:
    MsgBox "Access to the VBA project is blocked. Enable 'Trust access to the VBA project object model' in Excel Trust Center and retry.", vbCritical
    HasVBOMAccess = False
End Function

Private Sub RemoveExistingCodeExcept(ByVal keepModule As String)
    Dim vbcomp As Object, toRemove As Collection
    Set toRemove = New Collection
    For Each vbcomp In ThisWorkbook.VBProject.VBComponents
        Select Case vbcomp.Type
            Case 1, 2, 3 ' StdModule, ClassModule, MSForm
                If StrComp(vbcomp.Name, keepModule, vbTextCompare) <> 0 Then toRemove.Add vbcomp
            Case Else
                ' keep document modules (ThisWorkbook/Sheets)
        End Select
    Next vbcomp
    Dim c As Variant
    For Each c In toRemove
        ThisWorkbook.VBProject.VBComponents.Remove c
    Next c
End Sub

Private Sub ImportFolder(ByVal folder As String, ByVal pattern As String)
    If Dir(folder, vbDirectory) = vbNullString Then Exit Sub
    Dim file As String
    file = Dir(PathJoin(folder, pattern))
    Do While Len(file) > 0
        ThisWorkbook.VBProject.VBComponents.Import PathJoin(folder, file)
        file = Dir()
    Loop
End Sub

Private Function PathJoin(ByVal a As String, ByVal b As String) As String
    If Right$(a, 1) = "\" Or Right$(a, 1) = "/" Then
        PathJoin = a & b
    Else
        PathJoin = a & Application.PathSeparator & b
    End If
End Function

Private Function PickFolder(ByVal title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = title
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        Else
            PickFolder = vbNullString
        End If
    End With
End Function

