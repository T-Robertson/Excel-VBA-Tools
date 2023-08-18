Attribute VB_Name = "UnprotectAll"
Sub UnprotectAll()
Dim wSheet As Worksheet
Dim wBook As Workbook
Dim Pwd As String
Set wBook = ActiveWorkbook

    Pwd = InputBox("Enter your password to unprotect document", "Password Input")
    On Error Resume Next
    For Each wSheet In Worksheets
        wSheet.Unprotect Password:=Pwd
    Next wSheet
    If Err <> 0 Then
        MsgBox "You have entered an incorect password. All document could not " & _
        "be unprotected.", vbCritical, "Incorect Password"
    End If
    On Error GoTo 0
    wBook.Unprotect Password:=Pwd
End Sub


