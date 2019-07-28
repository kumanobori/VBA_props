Attribute VB_Name = "searchAndPaintTest"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private wb As Workbook
Private ws As Worksheet

Private STR_COLOR_FOUND As Double
Private BG_COLOR_FOUND As Double
Private STR_COLOR_NORMAL As Double
Private BG_COLORINDEX_NORMAL As Long

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    
    Application.ScreenUpdating = False
    
    STR_COLOR_FOUND = RGB(255, 0, 0)
    BG_COLOR_FOUND = RGB(255, 255, 0)
    STR_COLOR_NORMAL = RGB(0, 0, 0)
    BG_COLORINDEX_NORMAL = xlNone
    
    
    Workbooks.Add
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    ws.Range("B1").Value = "abcdefg"
    ws.Range("B2").Value = "ABCDEFG"
    ws.Range("B3").Value = "��������������"
    ws.Range("B4").Value = "�`�a�b�c�d�e�f"
    ws.Range("B5").Value = "abcdefg"
    ws.Range("C1").Value = "�͂Ђӂւ�"
    ws.Range("C2").Value = "�n�q�t�w�z"
    ws.Range("C3").Value = "�����"
    ws.Range("C4").Value = "�o�r�u�x�{"
    ws.Range("C5").Value = "����������"
    
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    wb.Close SaveChanges:=False
    
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    ws.Cells.Interior.ColorIndex = BG_COLORINDEX_NORMAL
    ws.Cells.Font.Color = STR_COLOR_NORMAL
    ws.Cells.Font.Bold = False
    ws.Range("A:A").Clear
End Sub

'@TestMethod("searchAndPaintTest")
' columnAasIndex=false�̂Ƃ��ɁAA�񂪌�������邱��
Private Sub columnAasIndex_false()
    On Error GoTo TestFail
    
    'Arrange:
    ws.Range("A1").Value = "�͂Ђӂւ�"
    ws.Range("A5").Value = "�͂Ђӂւ�"

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="�͂Ђӂւ�", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=False, toClear:=False, _
                            columnAasIndex:=False, matchCase:=False, matchByte:=False)

    'Assert:
    
    Call Assert.AreEqual(3, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("A1").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("A5").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C1").Interior.Color)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("searchAndPaintTest")
' columnAasIndex=true�̂Ƃ��ɁAA�񂪌�������Ȃ�����
Private Sub columnAasIndex_true()
    On Error GoTo TestFail
    
    'Arrange:
    ws.Range("A1").Value = "�͂Ђӂւ�"
    ws.Range("A5").Value = "�͂Ђӂւ�"

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="�͂Ђӂւ�", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=False, toClear:=False, _
                            columnAasIndex:=True, matchCase:=False, matchByte:=False)

    'Assert:
    Call Assert.AreEqual(1, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C1").Interior.Color)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("searchAndPaintTest")
' �����܂��������ׂ�OFF�̌���
Private Sub search_strict()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="bcd", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=True, toClear:=False, _
                            columnAasIndex:=False, matchCase:=True, matchByte:=True)

    'Assert:
    ' ���������Z����w�i�F�Ń`�F�b�N
    Call Assert.AreEqual(2, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B1").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B5").Interior.Color)
    ' B1�ŕ���������̑������`�F�b�N
    Call Assert.AreEqual(STR_COLOR_NORMAL, ws.Range("B1").Characters(1, 1).Font.Color)
    Call Assert.AreEqual(STR_COLOR_FOUND, ws.Range("B1").Characters(2, 3).Font.Color)
    Call Assert.AreEqual(STR_COLOR_NORMAL, ws.Range("B1").Characters(5, 3).Font.Color)
    Call Assert.AreEqual(False, ws.Range("B1").Characters(1, 1).Font.Bold)
    Call Assert.AreEqual(True, ws.Range("B1").Characters(2, 3).Font.Bold)
    Call Assert.AreEqual(False, ws.Range("B1").Characters(5, 3).Font.Bold)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("searchAndPaintTest")
' �啶������������ʂ��Ȃ�����
Private Sub search_matchCase_false()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="bcd", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=True, toClear:=False, _
                            columnAasIndex:=False, matchCase:=False, matchByte:=True)

    'Assert:
    ' ���������Z����w�i�F�Ń`�F�b�N
    Call Assert.AreEqual(3, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B1").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B2").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B5").Interior.Color)
    ' �����܂������Ńq�b�g�����Z���̕���������̑������`�F�b�N
    Call Assert.AreEqual(STR_COLOR_NORMAL, ws.Range("B2").Characters(1, 1).Font.Color)
    Call Assert.AreEqual(STR_COLOR_FOUND, ws.Range("B2").Characters(2, 3).Font.Color)
    Call Assert.AreEqual(STR_COLOR_NORMAL, ws.Range("B2").Characters(5, 3).Font.Color)
    Call Assert.AreEqual(False, ws.Range("B2").Characters(1, 1).Font.Bold)
    Call Assert.AreEqual(True, ws.Range("B2").Characters(2, 3).Font.Bold)
    Call Assert.AreEqual(False, ws.Range("B2").Characters(5, 3).Font.Bold)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("searchAndPaintTest")
' �S�p���p����ʂ��Ȃ�����_�A���t�@�x�b�g
Private Sub search_matchByte_false_01()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="bcd", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=True, toClear:=False, _
                            columnAasIndex:=False, matchCase:=False, matchByte:=True)

    'Assert:
    ' ���������Z����w�i�F�Ń`�F�b�N
    Call Assert.AreEqual(3, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B1").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B2").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B5").Interior.Color)
    ' �����܂������Ńq�b�g�����Z���̕���������̑������`�F�b�N
    Call Assert.AreEqual(STR_COLOR_NORMAL, ws.Range("B2").Characters(1, 1).Font.Color)
    Call Assert.AreEqual(STR_COLOR_FOUND, ws.Range("B2").Characters(2, 3).Font.Color)
    Call Assert.AreEqual(STR_COLOR_NORMAL, ws.Range("B2").Characters(5, 3).Font.Color)
    Call Assert.AreEqual(False, ws.Range("B2").Characters(1, 1).Font.Bold)
    Call Assert.AreEqual(True, ws.Range("B2").Characters(2, 3).Font.Bold)
    Call Assert.AreEqual(False, ws.Range("B2").Characters(5, 3).Font.Bold)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("searchAndPaintTest")
' �S�p���p����ʂ��Ȃ�����_�J�i
Private Sub search_matchByte_false_02()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="�q�t", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=True, toClear:=False, _
                            columnAasIndex:=False, matchCase:=True, matchByte:=False)

    'Assert:
    ' ���������Z����w�i�F�Ń`�F�b�N
    Call Assert.AreEqual(2, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C2").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C3").Interior.Color)
    ' �����܂������Ńq�b�g�����Z���̕���������̑������`�F�b�N
    Call Assert.AreEqual(STR_COLOR_NORMAL, ws.Range("C3").Characters(1, 1).Font.Color)
    Call Assert.AreEqual(STR_COLOR_FOUND, ws.Range("C3").Characters(2, 2).Font.Color)
    Call Assert.AreEqual(STR_COLOR_NORMAL, ws.Range("C3").Characters(4, 2).Font.Color)
    Call Assert.AreEqual(False, ws.Range("C3").Characters(1, 1).Font.Bold)
    Call Assert.AreEqual(True, ws.Range("C3").Characters(2, 2).Font.Bold)
    Call Assert.AreEqual(False, ws.Range("C3").Characters(4, 2).Font.Bold)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("searchAndPaintTest")
' �S�p���p����ʂ��Ȃ�����_�S�J�i�����J�i
Private Sub search_matchByte_false_03()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="�r�u", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=True, toClear:=False, _
                            columnAasIndex:=False, matchCase:=True, matchByte:=False)

    'Assert:
    ' ���������Z����w�i�F�Ń`�F�b�N
    Call Assert.AreEqual(2, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C4").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C5").Interior.Color)
    
    '**************************************
    '* ���p�J�i�ő��_���g���Ă���ꍇ�̕��������񑕏��͕s���S�Ȃ��߁A�e�X�g���R�����g�����Ă���
    '* (���s����Ǝ��s����)
    '**************************************
    ' �����܂������Ńq�b�g�����Z���̕���������̑������`�F�b�N
'    Call Assert.AreEqual(STR_COLOR_NORMAL, ws.Range("C5").Characters(1, 2).Font.Color)
'    Call Assert.AreEqual(STR_COLOR_FOUND, ws.Range("C5").Characters(3, 4).Font.Color)
'    Call Assert.AreEqual(STR_COLOR_NORMAL, ws.Range("C5").Characters(7, 4).Font.Color)
'    Call Assert.AreEqual(False, ws.Range("C5").Characters(1, 2).Font.Bold)
'    Call Assert.AreEqual(True, ws.Range("C5").Characters(3, 4).Font.Bold)
'    Call Assert.AreEqual(False, ws.Range("C5").Characters(7, 4).Font.Bold)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("searchAndPaintTest")
' ���ʃN���A�����Ȃ�
Private Sub toClear_invalid()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="bcd", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=False, toClear:=False, _
                            columnAasIndex:=True, matchCase:=True, matchByte:=True)
    Call Assert.AreEqual(2, count)
    count = fn_searchAndPaint(ws:=ws, strTarget:="�͂Ђ�", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=False, toClear:=False, _
                            columnAasIndex:=True, matchCase:=True, matchByte:=True)
    Call Assert.AreEqual(1, count)

    'Assert:
    ' ���������Z����w�i�F�Ń`�F�b�N
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B1").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B5").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C1").Interior.Color)
    ' A����`�F�b�N
    Call Assert.AreEqual("bcd, �͂Ђ�", ws.Range("A1").Value)
    Call Assert.AreEqual("bcd", ws.Range("A5").Value)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("searchAndPaintTest")
' ���ʃN���A������AA��C���f�b�N�X�L��
Private Sub toClear_valid_columnA_valid()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="bcd", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=False, toClear:=False, _
                            columnAasIndex:=True, matchCase:=True, matchByte:=True)
    Call Assert.AreEqual(2, count)
    count = fn_searchAndPaint(ws:=ws, strTarget:="�͂Ђ�", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=False, toClear:=True, _
                            columnAasIndex:=True, matchCase:=True, matchByte:=True)
    Call Assert.AreEqual(1, count)

    'Assert:
    ' ���������Z����w�i�F�Ń`�F�b�N
    Call Assert.AreEqual(BG_COLORINDEX_NORMAL, ws.Range("B1").Interior.ColorIndex)
    Call Assert.AreEqual(BG_COLORINDEX_NORMAL, ws.Range("B5").Interior.ColorIndex)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C1").Interior.Color)
    ' A����`�F�b�N
    Call Assert.AreEqual("�͂Ђ�", ws.Range("A1").Value)
    Call Assert.AreEqual("", ws.Range("A5").Value)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("searchAndPaintTest")
' ���ʃN���A������AA��C���f�b�N�X����
Private Sub toClear_valid_columnA_invalid()
    On Error GoTo TestFail
    
    'Arrange:
    ws.Range("A1").Value = "abcdefg"

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="bcd", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=False, toClear:=False, _
                            columnAasIndex:=False, matchCase:=True, matchByte:=True)
    Call Assert.AreEqual(3, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("A1").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B1").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B5").Interior.Color)
    count = fn_searchAndPaint(ws:=ws, strTarget:="�q�t�w", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=False, toClear:=True, _
                            columnAasIndex:=False, matchCase:=True, matchByte:=True)
    Call Assert.AreEqual(1, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C2").Interior.Color)

    'Assert:
    ' ���������Z����w�i�F�Ń`�F�b�N
    Call Assert.AreEqual("abcdefg", ws.Range("A1").Value)
    Call Assert.AreEqual(BG_COLORINDEX_NORMAL, ws.Range("A1").Interior.ColorIndex)
    Call Assert.AreEqual(STR_COLOR_NORMAL, ws.Range("A1").Font.Color)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub




Sub a()
    Dim wr As Range: Set wr = ActiveCell
    Dim c1 As Double, c2 As Double, c3 As Double, c4 As Double
    c1 = wr.Characters(1, 1).Font.Color
    c2 = wr.Characters(2, 3).Font.Color
    c3 = wr.Characters(5, 3).Font.Color
    c4 = wr.Characters(1, 2).Font.Color
End Sub
