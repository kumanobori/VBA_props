Attribute VB_Name = "m_searchAndPaint"
Option Explicit


Sub searchAndPaint()
    
    Dim strColor As Double: strColor = RGB(255, 0, 0)
    Dim bgColor As Double: bgColor = RGB(255, 255, 0)
    Dim isStrBold As Boolean: isStrBold = True
    Dim columnAasIndex As Boolean: columnAasIndex = True
    Dim matchCase As Boolean: matchCase = False
    Dim matchByte As Boolean: matchByte = False
    
    ' 検索前にクリアするフラグ
    Dim toClear As Boolean
    toClear = MsgBox("clear existing result?", vbYesNo) = vbYes
    
    ' 検索文字列
    Dim strTarget As String
    strTarget = InputBox("target string?")
    
    Dim count As Long
    count = fn_searchAndPaint(ws:=ActiveSheet _
                        , strTarget:=strTarget _
                        , strColor:=strColor _
                        , bgColor:=bgColor _
                        , isStrBold:=isStrBold _
                        , toClear:=toClear _
                        , columnAasIndex:=columnAasIndex _
                        , matchCase:=matchCase _
                        , matchByte:=matchByte _
                        )
    
    MsgBox "done. found count is " & count
End Sub

' ***********************************************************
' * ワークシート内を指定文字列で検索し、見つかったセルや文字列を装飾する。
' * @param ws 検索対象のワークシート
' * @param strTarget 検索文字列
' * @param strColor 見つかった文字列につける文字色
' * @param bgColor 文字列が見つかったセルにつける背景色
' * @param isStrBold trueなら、見つかった文字列を太字にする
' * @param toClear trueなら、検索する前に、すべての装飾を除去する
' * @param columnAasIndex trueなら、A列を結果行として、見つかった文字列を表示する
' * @param matchCase 大文字小文字を区別する
' * @param matchByte 全角半角を区別する
' * @return as long 検索語のヒット数
' ***********************************************************
Public Function fn_searchAndPaint(ws As Worksheet, strTarget As String, strColor As Double, bgColor As Double, isStrBold As Boolean, toClear As Boolean, columnAasIndex As Boolean, matchCase As Boolean, matchByte As Boolean)
    
    ' 検索対象エリアを設定
    Dim searchTargetArea As Range: Set searchTargetArea = IIf(columnAasIndex, ws.Range(ws.Range("B1"), ws.Cells.SpecialCells(xlLastCell)), ws.Cells)
    
    ' クリアフラグがついてる場合はクリア
    If toClear Then
        searchTargetArea.Interior.ColorIndex = xlNone
        searchTargetArea.Font.ColorIndex = xlAutomatic
        searchTargetArea.Font.Bold = False
        If columnAasIndex Then
            ws.Range("A:A").Clear
        End If
    End If
    
    Dim count As Long: count = 0
    
    ' 最初の検索
    Dim wr As Range: Set wr = searchTargetArea.Find(what:=strTarget, lookat:=xlPart, matchCase:=matchCase, matchByte:=matchByte)
    If wr Is Nothing Then
        fn_searchAndPaint = count
        Exit Function
    End If
    
    ' 装飾および、2回目以降の検索
    Dim foundFirst As Range: Set foundFirst = wr
    Do
        count = count + 1
        Call fn_decorateFoundCell(wr, strTarget, strColor, bgColor, isStrBold, columnAasIndex, matchCase, matchByte)
        
        Set wr = searchTargetArea.FindNext(wr)
    Loop Until wr.Address = foundFirst.Address
    
    fn_searchAndPaint = count
End Function


' ***********************************************************
' * セルの背景や部分文字列を装飾する。
' * あとおまけでA列への追記もやる。
' * @param wr 装飾対象のセル
' * @param strTarget 検索文字列
' * @param strColor 見つかった文字列につける文字色
' * @param bgColor 文字列が見つかったセルにつける背景色
' * @param isStrBold trueなら、見つかった文字列を太字にする
' * @param columnAasIndex trueなら、A列を結果行として、見つかった文字列を表示する
' * @param matchCase 大文字小文字を区別する
' * @param matchByte 全角半角を区別する
' * @return void
' ***********************************************************
Private Function fn_decorateFoundCell(wr As Range, strTarget As String, strColor As Double, bgColor As Double, isStrBold As Boolean, columnAasIndex As Boolean, matchCase As Boolean, matchByte As Boolean)
    
    Const COL_INDEX = 1
    Dim ws As Worksheet: Set ws = wr.Parent
    
    ' A列に検索文字列を追記する
    If columnAasIndex Then
        Dim wrIndex As Range: Set wrIndex = ws.Cells(wr.Row, COL_INDEX)
        If wrIndex.Value <> "" Then
            wrIndex.Value = wrIndex.Value & ", "
        End If
        wrIndex.Value = wrIndex.Value & strTarget
    End If
    
    ' 背景色を変える
    wr.Interior.Color = bgColor
    ' ヒット文字列の開始位置をコレクションで取得する
    Dim matchedIndexes As New Collection: Set matchedIndexes = fn_findStrMatchedIndexes(wr.Value, strTarget, matchCase, matchByte)
    
    ' 取得した開始位置ごとに、装飾する
    Dim i As Long
    For i = 1 To matchedIndexes.count
        Call fn_decorateStrInCell(wr:=wr, startIndex:=matchedIndexes(i), strLength:=Len(strTarget), strColor:=strColor, isStrBold:=isStrBold)
    Next i

End Function
' ****************************************************
' * haystack内にneedleがあれば、その開始位置を取得する。
' * needleが複数ある場合は、開始位置を全て取得する。
' * @param haystack needleを含む文字列
' * @param needle haystack内に含まれる部分文字列
' * @param matchCase 大文字小文字を区別する
' * @param matchByte 全角半角を区別する
' * @return as Collection<long> haystack内におけるneedleの開始位置を要素としたコレクション
' ****************************************************
Private Function fn_findStrMatchedIndexes(haystack As String, needle As String, matchCase As Boolean, matchByte As Boolean)
    
    Dim wkHaystack As String: wkHaystack = haystack
    Dim wkNeedle As String: wkNeedle = needle
    
    ' 大文字小文字を区別しない場合は全部大文字にする
    If Not matchCase Then
        wkHaystack = UCase(wkHaystack)
        wkNeedle = UCase(wkNeedle)
    End If
    ' 全角半角を区別しない場合は全部半角にする
    If Not matchByte Then
        wkHaystack = StrConv(wkHaystack, vbNarrow)
        wkNeedle = StrConv(wkNeedle, vbNarrow)
    End If
    
    Dim result As New Collection
    Dim i As Long
    i = 1
    Do While i <> 0
        i = InStr(i, wkHaystack, wkNeedle)
        If i <> 0 Then
            result.Add i
            i = i + 1
        End If
    Loop
    Set fn_findStrMatchedIndexes = result
End Function

' ****************************************************
' * セル内の部分文字列を装飾する。
' * @param wr 装飾対象のセル
' * @param startIndex 装飾対象の部分文字列の開始位置
' * @param strLength 装飾対象の文字列の文字数
' * @param strColor 文字列につける色
' * @param isStrBold trueであれば文字列を太字にする
' * @return void
' ****************************************************
Private Function fn_decorateStrInCell(wr As Range, startIndex As Long, strLength As Long, strColor As Double, isStrBold As Boolean)
    wr.Select
    With wr.Characters(Start:=startIndex, length:=strLength).Font
        .Color = strColor
        .Bold = isStrBold
    End With
End Function

