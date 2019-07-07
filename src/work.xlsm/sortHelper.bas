Attribute VB_Name = "sortHelper"
Option Explicit

' ********************************************
' * 【概略】
' * ソートを半自動化したもの。
' * キーを色々変えながらソートを繰り返すことを想定している。
' *
' * 【使い方】
' * 1. ソート順指定行を作る。
' * 1-1. ヘッダ行の上に空行をつくり、これをソート順指定行とする。
' * 2. ソート順指定行に、ソート順を記載する。
' * 2-1. 書式は「数字」+「AまたはD」。
' * 2-2. 「数字」は、その列のソートキー順位を表す。
' * 2-3. 「AまたはD」は、並び順が昇順か降順かを表す。Aは昇順、Dは降順。省略した場合は昇順となる。
' * 2-4. 例: 「1」「1A」「1D」「2」「2A」「2D」
' * 3. ソート範囲を指定する。方法は2通り。
' * 3-1. 方法1:ヘッダ行を含んだ、ソート範囲を選択する。
' * 3-2. 方法2:ヘッダ行を含めたソート範囲の文字列を、ソート範囲外の任意のセルに入力し、そのセルを選択する。(例：B4:E10)
' * 3-3. 注釈:いずれの方法の場合も、ソート範囲には、「ソート順指定行」は含まない。
' * 4. マクロ(Sub sortHelper)を実行する。
' ********************************************


Sub sortHelper()

    ' ソート範囲の取得
    Dim sortRange As Range
    If Selection.count > 1 Then
        ' 1. 複数セルを選択している場合、その選択範囲をソート範囲とする
        Set sortRange = Selection
    Else
        ' 2. 選択セルが1つの場合、そこに書かれているアドレス範囲をソート範囲とする
        Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = "[a-zA-Z]+[0-9]+:[a-zA-Z]+[0-9]+"
        Dim matches As Variant
        Set matches = regex.Execute(ActiveCell.Value)
        If matches.count = 1 Then
            Set sortRange = ActiveSheet.Range(ActiveCell.Value)
        Else
            ' 3. 1と2のいずれでも取得できなかった場合、エラー終了とする
            MsgBox "ソート範囲を取得できませんでした。" & vbCrLf & _
                   "ソート範囲をヘッダ行込みで選択するか、" & vbCrLf & _
                   "ソート範囲のアドレスを記入したセルを選択した状態で実行してください。"
            End
        End If
    End If
        
    ' ソート実施
    Call fn_sortByHead(sortRange)
    
    ' ソート範囲の最終セルへ移動する。(理由：ソート範囲を間違っていたとき気づきやすいようにするため)
    sortRange.Select
    sortRange(sortRange.count).Activate
    
End Sub


Private Function fn_sortByHead(sortTarget As Range)
    
    Dim ws As Worksheet: Set ws = sortTarget.Parent
    ' 領域情報を取得
    Dim sortTargetFirst As Range: Set sortTargetFirst = sortTarget(1)
    Dim sorttargetLast As Range: Set sorttargetLast = sortTarget(sortTarget.count)
    
    ' ソート対象領域ヘッダ行の1行上を、ソート情報RANGEとして取得し、ソート条件を生成
    Dim rangeSortCondition As Range: Set rangeSortCondition = ws.Range(sortTargetFirst.Offset(-1, 0), ws.Cells(sortTargetFirst.Row - 1, sorttargetLast.Column))
    Dim sortCondition As New Collection: Set sortCondition = fn_generateSortCondition(rangeSortCondition)
    
    ws.Sort.SortFields.Clear
    ws.Sort.SetRange sortTarget
    ws.Sort.Header = xlYes
    
    Dim i As Long
    ' sortConditionの内容を、ソート条件に組み込む
    ' Collectionの要素は、Array(列指定セルのrange as variant, 並び順 as long)
    For i = 1 To sortCondition.count
        ws.Sort.SortFields.Add Key:=sortCondition.Item(i)(0), Order:=sortCondition.Item(i)(1)
    Next i
    
    ws.Sort.Apply
    
End Function
' 設定値のセルを読み込んで、設定のCollectionを返す。
' Collectionの要素は、Array(列指定セルのrange as variant, 並び順 as long)
Private Function fn_generateSortCondition(rangeCondition As Range)

    ' 設定値をDictionaryに格納
    Dim conditionDictionary As Object: Set conditionDictionary = CreateObject("Scripting.Dictionary")
    Dim wr As Range
    For Each wr In rangeCondition
        If wr.Value <> "" Then
            Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
            regex.Pattern = "([0-9]+)([ad]?)"
            Dim matches As Variant
            Set matches = regex.Execute(wr.Value)
            Dim priority As Long: priority = Val(matches(0).submatches(0))
            Dim orderby As Long: orderby = IIf(UCase(matches(0).submatches(1)) = "D", xlDescending, xlAscending)
            Call conditionDictionary.Add(priority, Array(wr.Offset(1, 0), orderby))
        End If
    Next wr
    
    ' DictionaryのキーをArrayListに移してソート
    Dim wkSort As Object: Set wkSort = CreateObject("System.Collections.ArrayList")
    Dim i As Long
    For i = 0 To conditionDictionary.count - 1
        Call wkSort.Add(conditionDictionary.keys()(i))
    Next i
    Call wkSort.Sort
    
    ' ソートした順でDictionaryを取得し、Collectionに移す
    Dim result As New Collection
    Dim j As Long
    For j = 0 To wkSort.count - 1
        Call result.Add(conditionDictionary.Item(wkSort.Item(j)))
    Next j
    
    Set fn_generateSortCondition = result
End Function
