Attribute VB_Name = "bookNameToSheetName"
Option Explicit

' *********************************
' * 開いているブック全てに対して、
' * そのブックのシートが1個である場合に限り、
' * シート名をブック名と同じにする。
' *********************************

Sub booknameToSheetnameOnAllBooks()
    Dim WB As Workbook
    For Each WB In Workbooks
        Call renameSheetnameIfBookHasOnlyOneSheet(WB)
    Next WB
End Sub

Private Function renameSheetnameIfBookHasOnlyOneSheet(WB As Workbook)
    If WB.Worksheets.count = 1 Then
        Call bookNameToSheetName(WB.Worksheets(1))
    End If
End Function

Private Function bookNameToSheetName(ws As Worksheet)
    ws.name = Replace(ws.Parent.name, InStrRev(ws.Parent.name, "."), "")
End Function
