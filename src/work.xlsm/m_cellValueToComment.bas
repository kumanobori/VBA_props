Attribute VB_Name = "m_cellValueToComment"
Option Explicit

' ****************************************
' * ���݂̃Z�����e���R�����g�ɓ]�L����B
' * �R�����g���Ȃ��ꍇ�͍쐬����B
' ****************************************
Sub cellValueToComment()
    
    Dim wr As Range: Set wr = ActiveCell
    
    If wr.Value = "" Then
        MsgBox "�Z������Ȃ̂ŉ������܂���"
        End
    End If
    
    
    If wr.Comment Is Nothing Then
        wr.AddComment
    Else
        wr.Comment.Text Text:=wr.Comment.Text & vbCrLf
    End If
    
    wr.Comment.Text Text:=wr.Comment.Text & wr.Value
    
End Sub

