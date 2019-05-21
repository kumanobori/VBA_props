Attribute VB_Name = "SearchAndPaint"
Option Explicit

Const DO_CLEAR_ALL = False ' �ŏ��ɑ���������������t���O
Const DO_CHECK_COL_A = True ' �q�b�g�����Z����A��Ɉ������t���O
Const DO_CHECK_COL_A_AUTO = True ' �q�b�g�����Z����A��̈��������ɂ���t���O
Const DO_PAINT_INTERIOR_YELLOW = False ' �q�b�g����������F�ɓh��t���O�B
Const DO_PAINT_CHARACTER_RED = True ' �q�b�g�����Z���̕�����ɐF������t���O


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
    ' �w��͈͓����������A�q�b�g�����Z���𑕏�����B
    ' aTargetRange:�����͈�
    ' aStrTarget:������
    
    ' �����N���A
    If DO_CLEAR_ALL Then
        aTargetRange.Interior.ColorIndex = xlNone
        aTargetRange.Font.Bold = False
        aTargetRange.Font.Color = RGB(0, 0, 0)
        If DO_CHECK_COL_A Then
            aTargetRange.Range("A:A").ClearContents
        End If
    End If
    
    ' �����J�n�Z�����`
    Dim startRange As Range
    Set startRange = aTargetRange.Cells(1, 1)
    
    ' 1��ڂ̌���
    Dim firstRange As Range
    Set firstRange = aTargetRange.Find(what:=aStrTarget, after:=startRange)
    
    ' 1��ڂ̌������ʃZ����Nothing�łȂ���΁A�����J�n�Z���܂���1��ڂ̌��ʃZ���ɗ���܂ő�����
    If Not firstRange Is Nothing Then
        Dim wr As Range
        Set wr = firstRange
        Do
            wr.Activate
            ' �Z�b�g
            Call fn_paintSearchResult(aWr:=wr, aSearchString:=aStrTarget)
            ' ��������
            Set wr = aTargetRange.FindNext(after:=wr)
        Loop Until (wr.row = startRange.row And wr.Column = startRange.Column) Or (wr.row = firstRange.row And wr.Column = firstRange.Column)
        ' �Ōオ�����J�n�Z���������ꍇ�̃Z�b�g
        If fn_isSameCell(startRange, wr) Then
            Call fn_paintSearchResult(aWr:=wr, aSearchString:=aStrTarget)
        End If
    End If
    
    
End Function
Function fn_isSameCell(aWr1 As Range, aWr2 As Range)
    ' ������2�̃Z���������Z�����w�������Ă��邩�ۂ����肷��
    
    ' ���[�N�V�[�g���Ⴆ��false
    If aWr1.Parent.name <> aWr2.Parent.name Then
        fn_isSameCell = False
        Exit Function
    End If
    ' �s���Ⴆ��false
    If aWr1.row <> aWr2.row Then
        fn_isSameCell = False
        Exit Function
    End If
    ' �񂪈Ⴆ��false
    If aWr1.Column <> aWr2.Column Then
        fn_isSameCell = False
        Exit Function
    End If
    ' �`�F�b�N��S�Ēʉ߂�����true
    fn_isSameCell = True
End Function
Function fn_paintSearchResult(aWr As Range, aSearchString As String)
    ' �Z���𑕏�����B
    ' aWr:�ΏۃZ��
    ' aSearchString:�����ɗp����������B
    
    ' A�񏑂����݃��[�h�̏ꍇ�AA��̃q�b�g�͖���
    If DO_CHECK_COL_A And aWr.Column = 1 Then
        Exit Function
    End If
    
    ' ���[�N�V�[�g�擾
    Dim ws As Worksheet
    Set ws = aWr.Parent
    
    ' ���F�h��
    If DO_PAINT_INTERIOR_YELLOW Then
        aWr.Interior.Color = RGB(255, 255, 0)
    End If
    
    ' A��`�F�b�N
    If DO_CHECK_COL_A Then
        Dim colA As Range
        Set colA = ws.Cells(aWr.row, 1)
        If DO_CHECK_COL_A_AUTO Then
            colA.Value = colA.Value & "#" & aSearchString
        Else
            colA.Value = "��"
        End If
    End If
    
    ' �����h��
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
    ' Haystack����Needle������΁A���ꂪ�������ڂ��炩���擾����B
    ' �Ԃ�l��Long�l��Collection�B�v�f�́A�q�b�g�����J�nindex�B
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
