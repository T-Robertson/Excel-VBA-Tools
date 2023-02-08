Sub Quick_Cell_Copy()
Dim wb As Workbook
Dim ws As Worksheet
Dim ac As Range

Set wb = ThisWorkbook
Set ws = wb.ActiveSheet
Set ac = ws.Application.Selection

'Copy text to the clipboard
  Clipboard1 ac.Address

End Sub
Function Clipboard1(Optional StoreText As String) As String
'PURPOSE: Read/Write to Clipboard

Dim x As Variant

'Store as variant for 64-bit VBA support
  x = StoreText

'Create HTMLFile Object
  With CreateObject("htmlfile")
    With .parentWindow.clipboardData
      Select Case True
        Case Len(StoreText)
          'Write to the clipboard
            .setData "text", Replace(x, "$", "")
        Case Else
          'Read from the clipboard (no variable passed through)
            Clipboard = .GetData("text")
      End Select
    End With
  End With

End Function
