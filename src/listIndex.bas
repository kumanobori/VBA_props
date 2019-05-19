Attribute VB_Name = "listIndex"
Option Explicit

' ************************************************************
' * ワークシート一覧とそのリンクを作成する。
' * 作成場所は、現在のアクティブセルから下方向。
' ************************************************************

Sub listIndex()
    
    ' フォント設定
    Dim FONT_NAME As String: FONT_NAME = "ＭＳ ゴシック"
    Dim FONT_SIZE As Long: FONT_SIZE = 9
    Dim FONT_COLOR As Long: FONT_COLOR = RGB(0, 0, 0)
    
    If MsgBox("現在のセルから下に、現在のブックのシート名一覧を出力します。" & vbCrLf & "入力箇所に既存の入力があっても一切考慮しません。" & "実行しますか？", vbYesNoCancel) <> vbYes Then
        MsgBox "キャンセルしました"
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
    
    ' フォントを適用
    With wsOut.Range(wrOutStart, wrOut).Font
        .Name = FONT_NAME
        .Size = FONT_SIZE
        .Color = FONT_COLOR
    End With
    
    MsgBox "done"
End Sub





