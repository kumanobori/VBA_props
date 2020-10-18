Attribute VB_Name = "aaaModule1"
Option Explicit

Sub PageChange()
    '********************************************
    '*** 指定した文字列のある行で改ページする ***
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
    ' 今のブックを一旦閉じ、読み取り専用で開きなおす
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim wbFullPath As String
    wbFullPath = wb.Path & "\" & wb.Name
    wb.Close
    Workbooks.Open Filename:=wbFullPath, ReadOnly:=True
    
End Sub


Sub deleteFormatConditionsAll()
    If MsgBox("このブックの条件付き書式を全部削除します。復元はできません。", vbYesNoCancel) <> vbYes Then
        End
    End If
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Call fn_deleteFormatConditions(ws)
    Next ws
End Sub
Sub deleteFormatConditions()
    If MsgBox("このシートの条件付き書式を全部削除します。復元はできません。", vbYesNoCancel) <> vbYes Then
        End
    End If
    Call fn_deleteFormatConditions(ActiveSheet)
End Sub
' シートの条件付き書式を全削除する
Function fn_deleteFormatConditions(ws As Worksheet)
    ws.Cells.FormatConditions.Delete
End Function
' シートの条件付き書式を指定数を残して削除
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
    ' ダイアログで入力した文字列に一致する名称のワークシートを選択する。
    ' なかった場合はいちばん左のシートを選択する。
    ' ダイアログの初期値はいちばん左のシート。
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


' 色を表すLong値をRGBに変換する
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


