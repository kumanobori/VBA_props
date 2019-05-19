Attribute VB_Name = "listIndex"
Option Explicit

' ************************************************************
' * ���[�N�V�[�g�ꗗ�Ƃ��̃����N���쐬����B
' * �쐬�ꏊ�́A���݂̃A�N�e�B�u�Z�����牺�����B
' ************************************************************

Sub listIndex()
    
    ' �t�H���g�ݒ�
    Dim FONT_NAME As String: FONT_NAME = "�l�r �S�V�b�N"
    Dim FONT_SIZE As Long: FONT_SIZE = 9
    Dim FONT_COLOR As Long: FONT_COLOR = RGB(0, 0, 0)
    
    If MsgBox("���݂̃Z�����牺�ɁA���݂̃u�b�N�̃V�[�g���ꗗ���o�͂��܂��B" & vbCrLf & "���͉ӏ��Ɋ����̓��͂������Ă���؍l�����܂���B" & "���s���܂����H", vbYesNoCancel) <> vbYes Then
        MsgBox "�L�����Z�����܂���"
        End
    End If
    
    Dim wsOut As Worksheet: Set wsOut = ActiveSheet
    Dim wrOut As Range: Set wrOut = ActiveCell
    Dim wrOutStart As Range: Set wrOutStart = wrOut
    
    Dim wbIn As Workbook: Set wbIn = ActiveWorkbook
    Dim wsIn As Worksheet, i As Long
    For i = 1 To wbIn.Worksheets.Count
        Set wsIn = wbIn.Worksheets(i)
        
        wsOut.Hyperlinks.Add anchor:=wrOut, Address:="", SubAddress:=wsIn.Name & "!A1", _
                            ScreenTip:=" ", TextToDisplay:=wsIn.Name
        
        Set wrOut = wrOut.Offset(1, 0)
    Next i
    
    ' �t�H���g��K�p
    With wsOut.Range(wrOutStart, wrOut).Font
        .Name = FONT_NAME
        .Size = FONT_SIZE
        .Color = FONT_COLOR
    End With
    
    MsgBox "done"
End Sub





