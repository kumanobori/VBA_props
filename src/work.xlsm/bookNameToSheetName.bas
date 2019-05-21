Attribute VB_Name = "bookNameToSheetName"
Option Explicit

' *********************************
' * �J���Ă���u�b�N�S�Ăɑ΂��āA
' * ���̃u�b�N�̃V�[�g��1�ł���ꍇ�Ɍ���A
' * �V�[�g�����u�b�N���Ɠ����ɂ���B
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
