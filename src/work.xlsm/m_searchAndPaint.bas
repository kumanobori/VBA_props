Attribute VB_Name = "m_searchAndPaint"
Option Explicit


Sub searchAndPaint()
    
    Dim strColor As Double: strColor = RGB(255, 0, 0)
    Dim bgColor As Double: bgColor = RGB(255, 255, 0)
    Dim isStrBold As Boolean: isStrBold = True
    Dim columnAasIndex As Boolean: columnAasIndex = True
    Dim matchCase As Boolean: matchCase = False
    Dim matchByte As Boolean: matchByte = False
    
    ' �����O�ɃN���A����t���O
    Dim toClear As Boolean
    toClear = MsgBox("clear existing result?", vbYesNo) = vbYes
    
    ' ����������
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
' * ���[�N�V�[�g�����w�蕶����Ō������A���������Z���╶����𑕏�����B
' * @param ws �����Ώۂ̃��[�N�V�[�g
' * @param strTarget ����������
' * @param strColor ��������������ɂ��镶���F
' * @param bgColor �����񂪌��������Z���ɂ���w�i�F
' * @param isStrBold true�Ȃ�A��������������𑾎��ɂ���
' * @param toClear true�Ȃ�A��������O�ɁA���ׂĂ̑�������������
' * @param columnAasIndex true�Ȃ�AA������ʍs�Ƃ��āA���������������\������
' * @param matchCase �啶������������ʂ���
' * @param matchByte �S�p���p����ʂ���
' * @return as long ������̃q�b�g��
' ***********************************************************
Public Function fn_searchAndPaint(ws As Worksheet, strTarget As String, strColor As Double, bgColor As Double, isStrBold As Boolean, toClear As Boolean, columnAasIndex As Boolean, matchCase As Boolean, matchByte As Boolean)
    
    ' �����ΏۃG���A��ݒ�
    Dim searchTargetArea As Range: Set searchTargetArea = IIf(columnAasIndex, ws.Range(ws.Range("B1"), ws.Cells.SpecialCells(xlLastCell)), ws.Cells)
    
    ' �N���A�t���O�����Ă�ꍇ�̓N���A
    If toClear Then
        searchTargetArea.Interior.ColorIndex = xlNone
        searchTargetArea.Font.ColorIndex = xlAutomatic
        searchTargetArea.Font.Bold = False
        If columnAasIndex Then
            ws.Range("A:A").Clear
        End If
    End If
    
    Dim count As Long: count = 0
    
    ' �ŏ��̌���
    Dim wr As Range: Set wr = searchTargetArea.Find(what:=strTarget, lookat:=xlPart, matchCase:=matchCase, matchByte:=matchByte)
    If wr Is Nothing Then
        fn_searchAndPaint = count
        Exit Function
    End If
    
    ' ��������сA2��ڈȍ~�̌���
    Dim foundFirst As Range: Set foundFirst = wr
    Do
        count = count + 1
        Call fn_decorateFoundCell(wr, strTarget, strColor, bgColor, isStrBold, columnAasIndex, matchCase, matchByte)
        
        Set wr = searchTargetArea.FindNext(wr)
    Loop Until wr.Address = foundFirst.Address
    
    fn_searchAndPaint = count
End Function


' ***********************************************************
' * �Z���̔w�i�╔��������𑕏�����B
' * ���Ƃ��܂���A��ւ̒ǋL�����B
' * @param wr �����Ώۂ̃Z��
' * @param strTarget ����������
' * @param strColor ��������������ɂ��镶���F
' * @param bgColor �����񂪌��������Z���ɂ���w�i�F
' * @param isStrBold true�Ȃ�A��������������𑾎��ɂ���
' * @param columnAasIndex true�Ȃ�AA������ʍs�Ƃ��āA���������������\������
' * @param matchCase �啶������������ʂ���
' * @param matchByte �S�p���p����ʂ���
' * @return void
' ***********************************************************
Private Function fn_decorateFoundCell(wr As Range, strTarget As String, strColor As Double, bgColor As Double, isStrBold As Boolean, columnAasIndex As Boolean, matchCase As Boolean, matchByte As Boolean)
    
    Const COL_INDEX = 1
    Dim ws As Worksheet: Set ws = wr.Parent
    
    ' A��Ɍ����������ǋL����
    If columnAasIndex Then
        Dim wrIndex As Range: Set wrIndex = ws.Cells(wr.Row, COL_INDEX)
        If wrIndex.Value <> "" Then
            wrIndex.Value = wrIndex.Value & ", "
        End If
        wrIndex.Value = wrIndex.Value & strTarget
    End If
    
    ' �w�i�F��ς���
    wr.Interior.Color = bgColor
    ' �q�b�g������̊J�n�ʒu���R���N�V�����Ŏ擾����
    Dim matchedIndexes As New Collection: Set matchedIndexes = fn_findStrMatchedIndexes(wr.Value, strTarget, matchCase, matchByte)
    
    ' �擾�����J�n�ʒu���ƂɁA��������
    Dim i As Long
    For i = 1 To matchedIndexes.count
        Call fn_decorateStrInCell(wr:=wr, startIndex:=matchedIndexes(i), strLength:=Len(strTarget), strColor:=strColor, isStrBold:=isStrBold)
    Next i

End Function
' ****************************************************
' * haystack����needle������΁A���̊J�n�ʒu���擾����B
' * needle����������ꍇ�́A�J�n�ʒu��S�Ď擾����B
' * @param haystack needle���܂ޕ�����
' * @param needle haystack���Ɋ܂܂�镔��������
' * @param matchCase �啶������������ʂ���
' * @param matchByte �S�p���p����ʂ���
' * @return as Collection<long> haystack���ɂ�����needle�̊J�n�ʒu��v�f�Ƃ����R���N�V����
' ****************************************************
Private Function fn_findStrMatchedIndexes(haystack As String, needle As String, matchCase As Boolean, matchByte As Boolean)
    
    Dim wkHaystack As String: wkHaystack = haystack
    Dim wkNeedle As String: wkNeedle = needle
    
    ' �啶������������ʂ��Ȃ��ꍇ�͑S���啶���ɂ���
    If Not matchCase Then
        wkHaystack = UCase(wkHaystack)
        wkNeedle = UCase(wkNeedle)
    End If
    ' �S�p���p����ʂ��Ȃ��ꍇ�͑S�����p�ɂ���
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
' * �Z�����̕���������𑕏�����B
' * @param wr �����Ώۂ̃Z��
' * @param startIndex �����Ώۂ̕���������̊J�n�ʒu
' * @param strLength �����Ώۂ̕�����̕�����
' * @param strColor ������ɂ���F
' * @param isStrBold true�ł���Ε�����𑾎��ɂ���
' * @return void
' ****************************************************
Private Function fn_decorateStrInCell(wr As Range, startIndex As Long, strLength As Long, strColor As Double, isStrBold As Boolean)
    wr.Select
    With wr.Characters(Start:=startIndex, length:=strLength).Font
        .Color = strColor
        .Bold = isStrBold
    End With
End Function

