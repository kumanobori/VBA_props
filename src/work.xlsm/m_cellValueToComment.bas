Attribute VB_Name = "m_cellValueToComment"
Option Explicit

' ****************************************
' * 現在のセル内容をコメントに転記する。
' * コメントがない場合は作成する。
' ****************************************
Sub cellValueToComment()
    
    Dim wr As Range: Set wr = ActiveCell
    
    If wr.Value = "" Then
        MsgBox "セルが空なので何もしません"
        End
    End If
    
    
    If wr.Comment Is Nothing Then
        wr.AddComment
    Else
        wr.Comment.Text Text:=wr.Comment.Text & vbCrLf
    End If
    
    wr.Comment.Text Text:=wr.Comment.Text & wr.Value
    
End Sub

