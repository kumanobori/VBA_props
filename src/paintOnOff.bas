Attribute VB_Name = "paintOnOff"
Option Explicit

' ********************************************************************************
' * 選択しているセル全てに対して、色無しのセルは指定色、指定色のセルは色無しに塗り替える。
' * 暴走防止措置として、一定数以上のセルを選択しているときは処理しない。
' ********************************************************************************
Sub paintInterior()
    
    ' 設定: 指定色を定義
    Dim TARGET_COLOR As Long: TARGET_COLOR = RGB(255, 255, 0)  ' 黄色
    ' 設定:
    Dim MAX_COUNT As Long: MAX_COUNT = 10
    
    
    If Selection.Count > MAX_COUNT Then
        MsgBox "paintInterior: 選択中のセル数が" & MAX_COUNT & "を超えています。処理を行わず終了します。"
        End
    End If
    
    Dim wr As Range
    For Each wr In Selection
        ' 無色の場合は指定色に
        If wr.Interior.ColorIndex = xlNone Then
            wr.Interior.Color = TARGET_COLOR
        ' 指定色の場合は無色に
        Else
            If wr.Interior.Color = TARGET_COLOR Then
                wr.Interior.ColorIndex = xlNone
            End If
        End If
    Next wr
    
End Sub

