Attribute VB_Name = "sortHelper"
Option Explicit

' ********************************************
' * �y�T���z
' * �\�[�g�𔼎������������́B
' * �L�[��F�X�ς��Ȃ���\�[�g���J��Ԃ����Ƃ�z�肵�Ă���B
' *
' * �y�g�����z
' * 1. �\�[�g���w��s�����B
' * 1-1. �w�b�_�s�̏�ɋ�s������A������\�[�g���w��s�Ƃ���B
' * 2. �\�[�g���w��s�ɁA�\�[�g�����L�ڂ���B
' * 2-1. �����́u�����v+�uA�܂���D�v�B
' * 2-2. �u�����v�́A���̗�̃\�[�g�L�[���ʂ�\���B
' * 2-3. �uA�܂���D�v�́A���я����������~������\���BA�͏����AD�͍~���B�ȗ������ꍇ�͏����ƂȂ�B
' * 2-4. ��: �u1�v�u1A�v�u1D�v�u2�v�u2A�v�u2D�v
' * 3. �\�[�g�͈͂��w�肷��B���@��2�ʂ�B
' * 3-1. ���@1:�w�b�_�s���܂񂾁A�\�[�g�͈͂�I������B
' * 3-2. ���@2:�w�b�_�s���܂߂��\�[�g�͈͂̕�������A�\�[�g�͈͊O�̔C�ӂ̃Z���ɓ��͂��A���̃Z����I������B(��FB4:E10)
' * 3-3. ����:������̕��@�̏ꍇ���A�\�[�g�͈͂ɂ́A�u�\�[�g���w��s�v�͊܂܂Ȃ��B
' * 4. �}�N��(Sub sortHelper)�����s����B
' ********************************************


Sub sortHelper()

    ' �\�[�g�͈͂̎擾
    Dim sortRange As Range
    If Selection.count > 1 Then
        ' 1. �����Z����I�����Ă���ꍇ�A���̑I��͈͂��\�[�g�͈͂Ƃ���
        Set sortRange = Selection
    Else
        ' 2. �I���Z����1�̏ꍇ�A�����ɏ�����Ă���A�h���X�͈͂��\�[�g�͈͂Ƃ���
        Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = "[a-zA-Z]+[0-9]+:[a-zA-Z]+[0-9]+"
        Dim matches As Variant
        Set matches = regex.Execute(ActiveCell.Value)
        If matches.count = 1 Then
            Set sortRange = ActiveSheet.Range(ActiveCell.Value)
        Else
            ' 3. 1��2�̂�����ł��擾�ł��Ȃ������ꍇ�A�G���[�I���Ƃ���
            MsgBox "�\�[�g�͈͂��擾�ł��܂���ł����B" & vbCrLf & _
                   "�\�[�g�͈͂��w�b�_�s���݂őI�����邩�A" & vbCrLf & _
                   "�\�[�g�͈͂̃A�h���X���L�������Z����I��������ԂŎ��s���Ă��������B"
            End
        End If
    End If
        
    ' �\�[�g���{
    Call fn_sortByHead(sortRange)
    
    ' �\�[�g�͈͂̍ŏI�Z���ֈړ�����B(���R�F�\�[�g�͈͂��Ԉ���Ă����Ƃ��C�Â��₷���悤�ɂ��邽��)
    sortRange.Select
    sortRange(sortRange.count).Activate
    
End Sub


Private Function fn_sortByHead(sortTarget As Range)
    
    Dim ws As Worksheet: Set ws = sortTarget.Parent
    ' �̈�����擾
    Dim sortTargetFirst As Range: Set sortTargetFirst = sortTarget(1)
    Dim sorttargetLast As Range: Set sorttargetLast = sortTarget(sortTarget.count)
    
    ' �\�[�g�Ώۗ̈�w�b�_�s��1�s����A�\�[�g���RANGE�Ƃ��Ď擾���A�\�[�g�����𐶐�
    Dim rangeSortCondition As Range: Set rangeSortCondition = ws.Range(sortTargetFirst.Offset(-1, 0), ws.Cells(sortTargetFirst.Row - 1, sorttargetLast.Column))
    Dim sortCondition As New Collection: Set sortCondition = fn_generateSortCondition(rangeSortCondition)
    
    ws.Sort.SortFields.Clear
    ws.Sort.SetRange sortTarget
    ws.Sort.Header = xlYes
    
    Dim i As Long
    ' sortCondition�̓��e���A�\�[�g�����ɑg�ݍ���
    ' Collection�̗v�f�́AArray(��w��Z����range as variant, ���я� as long)
    For i = 1 To sortCondition.count
        ws.Sort.SortFields.Add Key:=sortCondition.Item(i)(0), Order:=sortCondition.Item(i)(1)
    Next i
    
    ws.Sort.Apply
    
End Function
' �ݒ�l�̃Z����ǂݍ���ŁA�ݒ��Collection��Ԃ��B
' Collection�̗v�f�́AArray(��w��Z����range as variant, ���я� as long)
Private Function fn_generateSortCondition(rangeCondition As Range)

    ' �ݒ�l��Dictionary�Ɋi�[
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
    
    ' Dictionary�̃L�[��ArrayList�Ɉڂ��ă\�[�g
    Dim wkSort As Object: Set wkSort = CreateObject("System.Collections.ArrayList")
    Dim i As Long
    For i = 0 To conditionDictionary.count - 1
        Call wkSort.Add(conditionDictionary.keys()(i))
    Next i
    Call wkSort.Sort
    
    ' �\�[�g��������Dictionary���擾���ACollection�Ɉڂ�
    Dim result As New Collection
    Dim j As Long
    For j = 0 To wkSort.count - 1
        Call result.Add(conditionDictionary.Item(wkSort.Item(j)))
    Next j
    
    Set fn_generateSortCondition = result
End Function
