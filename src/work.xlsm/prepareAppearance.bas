Attribute VB_Name = "prepareAppearance"
Option Explicit

' ****************************************
' * ���݂̃u�b�N�̑S�V�[�g�ňȉ����s���B
' * �E�w�肵���\���{����K�p
' * �E�w�肵���y�[�W�\�����@��K�p
' * �E�Z��A1��I��
' ****************************************

Sub prepareView()

    ' �ݒ�l�擾_�\���{��
    Dim zoomRate As Long: zoomRate = Val(InputBox("zoom rate?", "", 100))
    If 10 <= zoomRate And zoomRate <= 400 Then
        ' OK, do nothing
    Else
        MsgBox "�\���{���̗L���l��10�`400�ł��B"
        End
    End If
    
    ' �ݒ�l�擾_�\�����@
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
    
    ' �e�V�[�g�ɁA�\�����@��K�p
    Dim WB As Workbook: Set WB = ActiveWorkbook
    Dim ws As Worksheet
    For Each ws In WB.Worksheets
        ws.Activate
        ActiveWindow.View = viewType
        ActiveWindow.Zoom = zoomRate
        ws.Cells(1, 1).Activate
    Next ws
    
    ' �V�[�g�̂����A������ԍ��̂��̂�I��
    Dim i As Long
    For i = 1 To WB.Worksheets.count
        If WB.Worksheets(i).Visible Then
            WB.Select
            End
        End If
    Next i
    
End Sub

