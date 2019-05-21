Attribute VB_Name = "prepareAppearance"
Option Explicit

' ****************************************
' * 現在のブックの全シートで以下を行う。
' * ・指定した表示倍率を適用
' * ・指定したページ表示方法を適用
' * ・セルA1を選択
' ****************************************

Sub prepareView()

    ' 設定値取得_表示倍率
    Dim zoomRate As Long: zoomRate = Val(InputBox("zoom rate?", "", 100))
    If 10 <= zoomRate And zoomRate <= 400 Then
        ' OK, do nothing
    Else
        MsgBox "表示倍率の有効値は10～400です。"
        End
    End If
    
    ' 設定値取得_表示方法
    Dim viewType_wk As Long: viewType_wk = MsgBox("view type?" & "  Yes=NormalView," & "  No=PageBreakView," & "  Cancel=PageLayoutView", vbYesNoCancel)
    Dim viewType As Long
    Select Case viewType_wk
        Case vbYes:
            viewType = xlNormalView
        Case vbNo:
            viewType = xlPageBreakPreview
        Case vbCancel:
            viewType = xlPageLayoutView
        Case Else:
            MsgBox "unexpected input. exit."
            End
    End Select
    
    ' 各シートに、表示方法を適用
    Dim WB As Workbook: Set WB = ActiveWorkbook
    Dim ws As Worksheet
    For Each ws In WB.Worksheets
        ws.Activate
        ActiveWindow.View = viewType
        ActiveWindow.Zoom = zoomRate
        ws.Cells(1, 1).Activate
    Next ws
    
    ' シートのうち、可視かつ一番左のものを選択
    Dim i As Long
    For i = 1 To WB.Worksheets.count
        If WB.Worksheets(i).Visible Then
            WB.Select
            End
        End If
    Next i
    
End Sub

