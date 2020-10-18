Attribute VB_Name = "m_paintOnOff"
Option Explicit
' ********************************************************************************
' * 選択しているセル全てに対して、色無しのセルは指定色、指定色のセルは色無しに塗り替える。
' * 暴走防止措置として、一定数以上のセルを選択しているときは処理しない。
' ********************************************************************************
Sub paintOnOff()
    
    
    
    Dim result As Variant
    result = fn_paintOnOff(Selection)
    
    If Not result(0) Then
        MsgBox result(1)
    End If
    
    
End Sub

' @param targetRange これに含まれるセル全てを、色反転の対象とする。
' @return Array(result as boolean, message as string)
Public Function fn_paintOnOff(targetRange As Range)
    
    ' 設定: 指定色を定義
    Dim TARGET_COLOR As Long: TARGET_COLOR = RGB(255, 255, 0)  ' 黄色
    Dim MAX_COUNT As Long: MAX_COUNT = 100
    
    
    If targetRange.count > MAX_COUNT Then
        fn_paintOnOff = Array(False, "paintOnOff: 選択中のセル数が" & MAX_COUNT & "を超えています。処理を行わず終了します。")
        Exit Function
    End If
    
    Dim wr As Range
    For Each wr In targetRange
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
    fn_paintOnOff = Array(True)
End Function


