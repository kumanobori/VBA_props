Attribute VB_Name = "bookNameToSheetName"
Option Explicit

' *********************************
' * �J���Ă���u�b�N�S�Ăɑ΂��āA
' * ���̃u�b�N�̃V�[�g��1�ł���ꍇ�Ɍ���A
' * �V�[�g�����u�b�N���Ɠ����ɂ���B
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
