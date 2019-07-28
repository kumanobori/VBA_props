Attribute VB_Name = "editHeaderFooter"
Option Explicit

' **********************************************
' * 全シートのヘッダ・フッタを同内容で編集する
' **********************************************

Sub editHeaderFooter()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        With ws.PageSetup
            .LeftHeader = ""
            .CenterHeader = "&""ＭＳ ゴシック,標準""&9&F_&A"   ' 上段中央にブック名・シート名
            .RightHeader = "&""ＭＳ ゴシック,標準""&6&P/&N"    ' 上段右側にページ通番・総ページ数
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = "&""ＭＳ ゴシック,標準""&6&Z&F" & Chr(10) & "Printed: " & "&D_&T" ' 下段右側にファイルフルパス・印刷日時
        End With
    Next ws
    MsgBox "done"
End Sub
