Attribute VB_Name = "SearchAndPaint"
Option Explicit

Const DO_CLEAR_ALL = False ' 最初に装飾を初期化するフラグ
Const DO_CHECK_COL_A = True ' ヒットしたセルのA列に印を入れるフラグ
Const DO_CHECK_COL_A_AUTO = True ' ヒットしたセルのA列の印を検索語にするフラグ
Const DO_PAINT_INTERIOR_YELLOW = False ' ヒットした列を黄色に塗るフラグ。
Const DO_PAINT_CHARACTER_RED = True ' ヒットしたセルの文字列に色をつけるフラグ


Sub SearchAndPaint()
    Dim target As String
    target = InputBox("target string?")
    Call fn_searchAndPaintCurrentSheet(aStrTarget:=target)
    MsgBox "done"
End Sub
Function fn_searchAndPaintCurrentSheet(aStrTarget As String)
    Call fn_searchAndPaint(aTargetRange:=ActiveSheet.Cells, aStrTarget:=aStrTarget)
End Function
Function fn_searchAndPaint(aTargetRange As Range, aStrTarget As String)
    ' 指定範囲内を検索し、ヒットしたセルを装飾する。
    ' aTargetRange:検索範囲
    ' aStrTarget:検索語
    
    ' 装飾クリア
    If DO_CLEAR_ALL Then
        aTargetRange.Interior.ColorIndex = xlNone
        aTargetRange.Font.Bold = False
        aTargetRange.Font.Color = RGB(0, 0, 0)
        If DO_CHECK_COL_A Then
            aTargetRange.Range("A:A").ClearContents
        End If
    End If
    
    ' 検索開始セルを定義
    Dim startRange As Range
    Set startRange = aTargetRange.Cells(1, 1)
    
    ' 1回目の検索
    Dim firstRange As Range
    Set firstRange = aTargetRange.Find(what:=aStrTarget, after:=startRange)
    
    ' 1回目の検索結果セルがNothingでなければ、検索開始セルまたは1回目の結果セルに来るまで続ける
    If Not firstRange Is Nothing Then
        Dim wr As Range
        Set wr = firstRange
        Do
            wr.Activate
            ' セット
            Call fn_paintSearchResult(aWr:=wr, aSearchString:=aStrTarget)
            ' 次を検索
            Set wr = aTargetRange.FindNext(after:=wr)
        Loop Until (wr.row = startRange.row And wr.Column = startRange.Column) Or (wr.row = firstRange.row And wr.Column = firstRange.Column)
        ' 最後が検索開始セルだった場合のセット
        If fn_isSameCell(startRange, wr) Then
            Call fn_paintSearchResult(aWr:=wr, aSearchString:=aStrTarget)
        End If
    End If
    
    
End Function
Function fn_isSameCell(aWr1 As Range, aWr2 As Range)
    ' 引数の2つのセルが同じセルを指し示しているか否か判定する
    
    ' ワークシートが違えばfalse
    If aWr1.Parent.name <> aWr2.Parent.name Then
        fn_isSameCell = False
        Exit Function
    End If
    ' 行が違えばfalse
    If aWr1.row <> aWr2.row Then
        fn_isSameCell = False
        Exit Function
    End If
    ' 列が違えばfalse
    If aWr1.Column <> aWr2.Column Then
        fn_isSameCell = False
        Exit Function
    End If
    ' チェックを全て通過したらtrue
    fn_isSameCell = True
End Function
Function fn_paintSearchResult(aWr As Range, aSearchString As String)
    ' セルを装飾する。
    ' aWr:対象セル
    ' aSearchString:検索に用いた文字列。
    
    ' A列書き込みモードの場合、A列のヒットは無視
    If DO_CHECK_COL_A And aWr.Column = 1 Then
        Exit Function
    End If
    
    ' ワークシート取得
    Dim ws As Worksheet
    Set ws = aWr.Parent
    
    ' 黄色塗り
    If DO_PAINT_INTERIOR_YELLOW Then
        aWr.Interior.Color = RGB(255, 255, 0)
    End If
    
    ' A列チェック
    If DO_CHECK_COL_A Then
        Dim colA As Range
        Set colA = ws.Cells(aWr.row, 1)
        If DO_CHECK_COL_A_AUTO Then
            colA.Value = colA.Value & "#" & aSearchString
        Else
            colA.Value = "○"
        End If
    End If
    
    ' 文字塗り
    If DO_PAINT_CHARACTER_RED Then
        Call fn_paintCharacter(aWr, aSearchString, RGB(255, 0, 0))
    End If
    
    
End Function
Function fn_paintCharacter(aWr As Range, aSearchString As String, charColor As Long)
    Dim indexes As Collection
    Set indexes = fn_getCharMatchedIndexes(aWr.Value, aSearchString)
    Dim idx As Variant
    For Each idx In indexes
        With aWr.Characters(start:=idx, length:=Len(aSearchString)).Font
            .Color = charColor
            .Bold = True
        End With
    Next idx
    
    
    
End Function
Function fn_getCharMatchedIndexes(aHaystack As String, aNeedle As String)
    ' Haystack内にNeedleがあれば、それが何文字目からかを取得する。
    ' 返り値はLong値のCollection。要素は、ヒットした開始index。
    Dim lengthNeedle As Long
    
    Dim result As New Collection
        
    Dim i As Long
    i = 1
    Do While i <> 0
        i = InStr(i, aHaystack, aNeedle)
        If i <> 0 Then
            result.Add i
            i = i + 1
        End If
    Loop
    Set fn_getCharMatchedIndexes = result
End Function
Sub test()
    Dim a As String
    a = "abcdefg"
    Dim b As Variant
    Dim c As Variant
    Dim d As Variant
    b = InStr(a, "a")
    c = InStr(1, a, "a")
    d = InStr(0, a, "a")
    
    
End Sub
