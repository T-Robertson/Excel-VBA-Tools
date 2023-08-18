Attribute VB_Name = "ProtectAll"
Sub ProtectAll()
Dim wSheet As Worksheet
Dim wBook As Workbook
Dim Pwd As String
     
    Pwd = InputBox("Enter your password to protect document", "Password Input")
    For Each wSheet In Worksheets
        wSheet.Protect Password:=Pwd
        wBook.Protect Password:=Pwd
    Next wSheet
     
End Sub


