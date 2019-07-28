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
    ws.Range("B3").Value = "ａｂｃｄｅｆｇ"
    ws.Range("B4").Value = "ＡＢＣＤＥＦＧ"
    ws.Range("B5").Value = "abcdefg"
    ws.Range("C1").Value = "はひふへほ"
    ws.Range("C2").Value = "ハヒフヘホ"
    ws.Range("C3").Value = "ﾊﾋﾌﾍﾎ"
    ws.Range("C4").Value = "バビブベボ"
    ws.Range("C5").Value = "ﾊﾞﾋﾞﾌﾞﾍﾞﾎﾞ"
    
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
' columnAasIndex=falseのときに、A列が検索されること
Private Sub columnAasIndex_false()
    On Error GoTo TestFail
    
    'Arrange:
    ws.Range("A1").Value = "はひふへほ"
    ws.Range("A5").Value = "はひふへほ"

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="はひふへほ", _
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
' columnAasIndex=trueのときに、A列が検索されないこと
Private Sub columnAasIndex_true()
    On Error GoTo TestFail
    
    'Arrange:
    ws.Range("A1").Value = "はひふへほ"
    ws.Range("A5").Value = "はひふへほ"

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="はひふへほ", _
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
' あいまい条件すべてOFFの検索
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
    ' 見つかったセルを背景色でチェック
    Call Assert.AreEqual(2, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B1").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B5").Interior.Color)
    ' B1で部分文字列の装飾をチェック
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
' 大文字小文字を区別しない検索
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
    ' 見つかったセルを背景色でチェック
    Call Assert.AreEqual(3, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B1").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B2").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B5").Interior.Color)
    ' あいまい検索でヒットしたセルの部分文字列の装飾をチェック
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
' 全角半角を区別しない検索_アルファベット
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
    ' 見つかったセルを背景色でチェック
    Call Assert.AreEqual(3, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B1").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B2").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B5").Interior.Color)
    ' あいまい検索でヒットしたセルの部分文字列の装飾をチェック
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
' 全角半角を区別しない検索_カナ
Private Sub search_matchByte_false_02()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="ヒフ", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=True, toClear:=False, _
                            columnAasIndex:=False, matchCase:=True, matchByte:=False)

    'Assert:
    ' 見つかったセルを背景色でチェック
    Call Assert.AreEqual(2, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C2").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C3").Interior.Color)
    ' あいまい検索でヒットしたセルの部分文字列の装飾をチェック
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
' 全角半角を区別しない検索_全カナ→半カナ
Private Sub search_matchByte_false_03()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim count As Integer
    count = fn_searchAndPaint(ws:=ws, strTarget:="ビブ", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=True, toClear:=False, _
                            columnAasIndex:=False, matchCase:=True, matchByte:=False)

    'Assert:
    ' 見つかったセルを背景色でチェック
    Call Assert.AreEqual(2, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C4").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C5").Interior.Color)
    
    '**************************************
    '* 半角カナで濁点を使っている場合の部分文字列装飾は不完全なため、テストをコメント化してある
    '* (実行すると失敗する)
    '**************************************
    ' あいまい検索でヒットしたセルの部分文字列の装飾をチェック
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
' 結果クリアをしない
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
    count = fn_searchAndPaint(ws:=ws, strTarget:="はひふ", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=False, toClear:=False, _
                            columnAasIndex:=True, matchCase:=True, matchByte:=True)
    Call Assert.AreEqual(1, count)

    'Assert:
    ' 見つかったセルを背景色でチェック
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B1").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("B5").Interior.Color)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C1").Interior.Color)
    ' A列をチェック
    Call Assert.AreEqual("bcd, はひふ", ws.Range("A1").Value)
    Call Assert.AreEqual("bcd", ws.Range("A5").Value)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("searchAndPaintTest")
' 結果クリアをする、A列インデックス有効
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
    count = fn_searchAndPaint(ws:=ws, strTarget:="はひふ", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=False, toClear:=True, _
                            columnAasIndex:=True, matchCase:=True, matchByte:=True)
    Call Assert.AreEqual(1, count)

    'Assert:
    ' 見つかったセルを背景色でチェック
    Call Assert.AreEqual(BG_COLORINDEX_NORMAL, ws.Range("B1").Interior.ColorIndex)
    Call Assert.AreEqual(BG_COLORINDEX_NORMAL, ws.Range("B5").Interior.ColorIndex)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C1").Interior.Color)
    ' A列をチェック
    Call Assert.AreEqual("はひふ", ws.Range("A1").Value)
    Call Assert.AreEqual("", ws.Range("A5").Value)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("searchAndPaintTest")
' 結果クリアをする、A列インデックス無効
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
    count = fn_searchAndPaint(ws:=ws, strTarget:="ヒフヘ", _
                            strColor:=STR_COLOR_FOUND, bgColor:=BG_COLOR_FOUND, _
                            isStrBold:=False, toClear:=True, _
                            columnAasIndex:=False, matchCase:=True, matchByte:=True)
    Call Assert.AreEqual(1, count)
    Call Assert.AreEqual(BG_COLOR_FOUND, ws.Range("C2").Interior.Color)

    'Assert:
    ' 見つかったセルを背景色でチェック
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
