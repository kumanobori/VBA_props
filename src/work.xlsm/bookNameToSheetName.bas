Attribute VB_Name = "bookNameToSheetName"
Option Explicit

' *********************************
' * 開いているブック全てに対して、
' * そのブックのシートが1個である場合に限り、
' * シート名をブック名と同じにする。
' *********************************

Sub booknameToSheetnameOnAllBooks()
    Dim wb As Workbook
    For Each wb In Workbooks
        Call renameSheetnameIfBookHasOnlyOneSheet(wb)
    Next wb
End Sub

Private Function renameSheetnameIfBookHasOnlyOneSheet(wb As Workbook)
    If wb.Worksheets.count = 1 Then
        Call bookNameToSheetName(wb.Worksheets(1))
    End If
End Function

Private Function bookNameToSheetName(ws As Worksheet)
    ws.Name = Replace(ws.Parent.Name, InStrRev(ws.Parent.Name, "."), "")
End Function
