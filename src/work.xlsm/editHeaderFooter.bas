Attribute VB_Name = "editHeaderFooter"
Option Explicit

' **********************************************
' * �S�V�[�g�̃w�b�_�E�t�b�^�𓯓��e�ŕҏW����
' **********************************************

Sub editHeaderFooter()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        With ws.PageSetup
            .LeftHeader = ""
            .CenterHeader = "&""�l�r �S�V�b�N,�W��""&9&F_&A"   ' ��i�����Ƀu�b�N���E�V�[�g��
            .RightHeader = "&""�l�r �S�V�b�N,�W��""&6&P/&N"    ' ��i�E���Ƀy�[�W�ʔԁE���y�[�W��
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = "&""�l�r �S�V�b�N,�W��""&6&Z&F" & Chr(10) & "Printed: " & "&D_&T" ' ���i�E���Ƀt�@�C���t���p�X�E�������
        End With
    Next ws
    MsgBox "done"
End Sub
