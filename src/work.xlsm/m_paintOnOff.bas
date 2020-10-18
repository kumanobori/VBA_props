Attribute VB_Name = "m_paintOnOff"
Option Explicit
' ********************************************************************************
' * �I�����Ă���Z���S�Ăɑ΂��āA�F�����̃Z���͎w��F�A�w��F�̃Z���͐F�����ɓh��ւ���B
' * �\���h�~�[�u�Ƃ��āA��萔�ȏ�̃Z����I�����Ă���Ƃ��͏������Ȃ��B
' ********************************************************************************
Sub paintOnOff()
    
    
    
    Dim result As Variant
    result = fn_paintOnOff(Selection)
    
    If Not result(0) Then
        MsgBox result(1)
    End If
    
    
End Sub

' @param targetRange ����Ɋ܂܂��Z���S�Ă��A�F���]�̑ΏۂƂ���B
' @return Array(result as boolean, message as string)
Public Function fn_paintOnOff(targetRange As Range)
    
    ' �ݒ�: �w��F���`
    Dim TARGET_COLOR As Long: TARGET_COLOR = RGB(255, 255, 0)  ' ���F
    Dim MAX_COUNT As Long: MAX_COUNT = 100
    
    
    If targetRange.count > MAX_COUNT Then
        fn_paintOnOff = Array(False, "paintOnOff: �I�𒆂̃Z������" & MAX_COUNT & "�𒴂��Ă��܂��B�������s�킸�I�����܂��B")
        Exit Function
    End If
    
    Dim wr As Range
    For Each wr In targetRange
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
    fn_paintOnOff = Array(True)
End Function


