Attribute VB_Name = "aaaModule1"
Option Explicit

Sub PageChange()
    '********************************************
    '*** �w�肵��������̂���s�ŉ��y�[�W���� ***
    '********************************************
    Dim HDR As String
    HDR = inpubox("header string?", , "REC. NO.")
    Call Fn_PageChange(HDR)
End Sub

Function Fn_PageChange(HDR As String)
    Dim wr As Range
    Set wr = Range("A1")
    Dim WRsav As Range
    Set WRsav = wr
    Do While wr.Row >= WRsav.Row
        Set WRsav = wr
        Set wr = Cells.Find(what:=HDR, after:=WRsav)
        ActiveSheet.HPageBreaks.Add before:=wr
    Loop
End Function

Sub reOpenAsReadOnly()
    ' ���̃u�b�N����U���A�ǂݎ���p�ŊJ���Ȃ���
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim wbFullPath As String
    wbFullPath = wb.Path & "\" & wb.Name
    wb.Close
    Workbooks.Open Filename:=wbFullPath, ReadOnly:=True
    
End Sub


Sub deleteFormatConditionsAll()
    If MsgBox("���̃u�b�N�̏����t��������S���폜���܂��B�����͂ł��܂���B", vbYesNoCancel) <> vbYes Then
        End
    End If
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Call fn_deleteFormatConditions(ws)
    Next ws
End Sub
Sub deleteFormatConditions()
    If MsgBox("���̃V�[�g�̏����t��������S���폜���܂��B�����͂ł��܂���B", vbYesNoCancel) <> vbYes Then
        End
    End If
    Call fn_deleteFormatConditions(ActiveSheet)
End Sub
' �V�[�g�̏����t��������S�폜����
Function fn_deleteFormatConditions(ws As Worksheet)
    ws.Cells.FormatConditions.Delete
End Function
' �V�[�g�̏����t���������w�萔���c���č폜
Function wk_FormatConditions()
    Dim conditionsLeft As Long
    conditionsLeft = 7
    Application.ScreenUpdating = False
    Dim count As Long
    count = Cells.FormatConditions.count
    MsgBox count
    Dim fm As FormatCondition
    Dim i As Long
    For i = count - conditionsLeft To 1 Step -1
        Set fm = Cells.FormatConditions.Item(i)
        fm.Delete
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "done"
End Function



Sub selectWorksheet()
    ' �_�C�A���O�œ��͂���������Ɉ�v���閼�̂̃��[�N�V�[�g��I������B
    ' �Ȃ������ꍇ�͂����΂񍶂̃V�[�g��I������B
    ' �_�C�A���O�̏����l�͂����΂񍶂̃V�[�g�B
    Dim wsName As String
    wsName = ActiveWorkbook.Worksheets(1).Name
    wsName = InputBox(prompt:="Worksheet Name?", default:=wsName)
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If wsName = ws.Name Then
            ws.Select
            ws.Cells(1, 1).Select
            Exit Sub
        End If
    Next ws
    MsgBox prompt:="not found", Buttons:=vbOKOnly
    ActiveWorkbook.Worksheets(1).Select
    Cells(1, 1).Select
End Sub


' �F��\��Long�l��RGB�ɕϊ�����
Private Function fn_longToRgb(i As Long)
    Dim DIV As Long
    DIV = 256
    Dim r As Long, g As Long, b As Long, wk As Long
    wk = i
    r = wk Mod DIV
    wk = (wk - r) / DIV
    g = wk Mod DIV
    wk = (wk - g) / DIV
    b = wk
    fn_longToRgb = "RGB(" & r & ", " & g & ", " & b & ")" & vbCrLf & _
    "in=" & Str(i) & vbCrLf & _
    "out=" & Str(r * DIV * DIV + g * DIV + b)
End Function


