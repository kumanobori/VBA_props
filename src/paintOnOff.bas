Attribute VB_Name = "paintOnOff"
Option Explicit

' ********************************************************************************
' * �I�����Ă���Z���S�Ăɑ΂��āA�F�����̃Z���͎w��F�A�w��F�̃Z���͐F�����ɓh��ւ���B
' * �\���h�~�[�u�Ƃ��āA��萔�ȏ�̃Z����I�����Ă���Ƃ��͏������Ȃ��B
' ********************************************************************************
Sub paintInterior()
    
    ' �ݒ�: �w��F���`
    Dim TARGET_COLOR As Long: TARGET_COLOR = RGB(255, 255, 0)  ' ���F
    ' �ݒ�:
    Dim MAX_COUNT As Long: MAX_COUNT = 10
    
    
    If Selection.Count > MAX_COUNT Then
        MsgBox "paintInterior: �I�𒆂̃Z������" & MAX_COUNT & "�𒴂��Ă��܂��B�������s�킸�I�����܂��B"
        End
    End If
    
    Dim wr As Range
    For Each wr In Selection
        ' ���F�̏ꍇ�͎w��F��
        If wr.Interior.ColorIndex = xlNone Then
            wr.Interior.Color = TARGET_COLOR
        ' �w��F�̏ꍇ�͖��F��
        Else
            If wr.Interior.Color = TARGET_COLOR Then
                wr.Interior.ColorIndex = xlNone
            End If
        End If
    Next wr
    
End Sub

