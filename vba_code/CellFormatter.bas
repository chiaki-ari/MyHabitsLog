Attribute VB_Name = "CellFormatter"
' �r����ʒ�`
Public Const BORDER_INSIDE = 0
Public Const BORDER_OUTSIDE = 1
Public Const BORDER_TOP = 2
Public Const BORDER_LEFT = 3
Public Const BORDER_RIGHT = 4
Public Const BORDER_BOTTOM = 5
Public Const BORDER_VERTICAL = 6
Public Const BORDER_HORIZONAL = 7

'�t�H���g�T�C�Y�̒萔
Public Const FONT_SIZE_DEFAULT As Long = 10
Public Const FONT_SIZE_SMALL As Long = 9
Public Const FONT_SIZE_VERY_SMALL As Long = 8

' �r�b�g��`
Private Const DECORATE_BACKGROUND As Long = &H1    ' �w�i�F����
Private Const DECORATE_BORDER_OUTER As Long = &H2  ' �O�g������
Private Const DECORATE_BORDER_INNER As Long = &H4  ' ���g������
Private Const DECORATE_BORDER_HORIZONTAL As Long = &H8  ' ���������̌r��
Private Const DECORATE_BORDER_VERTICAL As Long = &H10  ' ���������̌r��

' ��Ԃ��`
Public Const STATE_HIGHLIGHT_BORDER_ALL As Long = DECORATE_BACKGROUND Or DECORATE_BORDER_OUTER Or DECORATE_BORDER_INNER
Public Const STATE_HIGHLIGHT_NO_BORDER As Long = DECORATE_BACKGROUND
Public Const STATE_HIGHLIGHT_OUTER_ONLY As Long = DECORATE_BACKGROUND Or DECORATE_BORDER_OUTER
Public Const STATE_NO_BACKGROUND_OUTER_ONLY As Long = DECORATE_BORDER_OUTER
Public Const STATE_NO_BACKGROUND_BORDER_ALL As Long = DECORATE_BORDER_OUTER Or DECORATE_BORDER_INNER
Public Const STATE_NO_BACKGROUND_BORDER_HORIZONTAL As Long = DECORATE_BORDER_OUTER Or DECORATE_BORDER_HORIZONTAL ' ���������̂ݐ�
Public Const STATE_NO_BACKGROUND_BORDER_VERTICAL As Long = DECORATE_BORDER_OUTER Or DECORATE_BORDER_VERTICAL ' ���������̂ݐ�
' �F�̒�`
Public Const COLOR_GRAY As Long = &H808080 ' RGB(128, 128, 128)
Public Const COLOR_LIGHT_GRAY As Long = &HBFBFBF ' RGB(191, 191, 191)
Public Const COLOR_VERY_LIGHT_GRAY As Long = &HF2F2F2 ' RGB(242, 242, 242)
Public Const COLOR_RED As Long = &HFF ' RGB(255, 0, 0)
Public Const COLOR_BLUE As Long = &H985221 ' RGB(33, 92, 152)
Public Const COLOR_DARK_GRAY As Long = &H262626
Public Const COLOR_BLACK As Long = &O0

' �F�̒�`
Public Const UNCOLOR_CELL As Boolean = False
Public Const COLOR_CELL As Boolean = True

Public Const BORDER_CLEAR = xlLineStyleNone

Public Const RULE_ON As Boolean = True
Public Const RULE_OFF As Boolean = False

Public Const MERGE As Boolean = True
Public Const UNMERGE As Boolean = False

Public Const WRAP As Boolean = True
Public Const UNWRAP As Boolean = False
Public Sub ChangeSheetFontSize(targetSheet As Worksheet, fontSize As Long)
    If targetSheet Is Nothing Then Exit Sub
    targetSheet.Cells.Font.Size = fontSize
End Sub
Public Sub ChangeRangeFontSize(startCell As Range, endCell As Range, fontSize As Long)
    Dim targetRange As Range
    '�Z���͈͂�ݒ�
    Set targetRange = Range(startCell, endCell)
    targetRange.Font.Size = fontSize
End Sub

' �F�t���g�̐ݒ�
Public Sub setDecorateCell(startCell As Range, endCell As Range, cellDecoType As Long)
    '�w�i�F
    If (cellDecoType And DECORATE_BACKGROUND) <> 0 Then
        Call setBackgroundColor(startCell, endCell, COLOR_VERY_LIGHT_GRAY)
    Else
        'Call setBackgroundColor(startCell, endCell, xlNone) ' �w�i�F������
    End If

    '�O�g�Ɠ��g
    If (cellDecoType And DECORATE_BORDER_OUTER) <> 0 Then
        Call setBorders(startCell, endCell, COLOR_GRAY, BORDER_OUTSIDE) ' �O�g
    End If

    If (cellDecoType And DECORATE_BORDER_INNER) <> 0 Then
        Call setBorders(startCell, endCell, COLOR_LIGHT_GRAY, BORDER_INSIDE) ' ���g
    End If
    
    If (cellDecoType And DECORATE_BORDER_HORIZONTAL) <> 0 Then
        Call setBorders(startCell, endCell, COLOR_LIGHT_GRAY, BORDER_HORIZONAL) ' ��������
    End If

    If (cellDecoType And DECORATE_BORDER_VERTICAL) <> 0 Then
        Call setBorders(startCell, endCell, COLOR_LIGHT_GRAY, BORDER_VERTICAL) ' ��������
    End If
End Sub


' �Z�������E����
Public Sub mergeCells(startCell As Range, endCell As Range, isMerge As Boolean)
    Dim targetRange As Range
    '�Z���͈͂�ݒ�
    Set targetRange = Range(startCell, endCell)
    If isMerge Then
        targetRange.MERGE
    Else
        targetRange.UNMERGE
    End If
End Sub
' �g���w�i���N���A
Public Sub clearDecorateCell(startCell As Range, endCell As Range)
    Call setBorders(startCell, endCell, BORDER_CLEAR, BORDER_OUTSIDE)
    Call setBorders(startCell, endCell, BORDER_CLEAR, BORDER_INSIDE)
    Call setBackgroundColor(startCell, endCell, xlLineStyleNone) '�w�i�F������
End Sub
' �����ʒu�w��
Public Sub setTextAlign(startCell As Range, endCell As Range, alignType As Long)
    Range(startCell, endCell).HorizontalAlignment = alignType
End Sub
' �Z���r����ݒ�/�N���A(targetColor��BORDER_CLEAR�Ōr���폜�j
Sub setBorders(startCell As Range, endCell As Range, targetColor As Long, targetEdge As Long)
    Select Case targetEdge
        Case BORDER_OUTSIDE
            Range(startCell, endCell).Borders(xlEdgeTop).color = targetColor
            Range(startCell, endCell).Borders(xlEdgeLeft).color = targetColor
            Range(startCell, endCell).Borders(xlEdgeRight).color = targetColor
            Range(startCell, endCell).Borders(xlEdgeBottom).color = targetColor
        Case BORDER_INSIDE
            Range(startCell, endCell).Borders(xlInsideVertical).color = targetColor
            Range(startCell, endCell).Borders(xlInsideHorizontal).color = targetColor
        Case BORDER_VERTICAL
            Range(startCell, endCell).Borders(xlInsideVertical).color = targetColor
        Case BORDER_HORIZONAL
            Range(startCell, endCell).Borders(xlInsideHorizontal).color = targetColor
        Case BORDER_TOP
            Range(startCell, endCell).Borders(xlEdgeTop).color = targetColor
        Case BORDER_RIGHT
            Range(startCell, endCell).Borders(xlEdgeRight).color = targetColor
        Case BORDER_LEFT
            Range(startCell, endCell).Borders(xlEdgeLeft).color = targetColor
        Case BORDER_BOTTOM
            Range(startCell, endCell).Borders(BORDER_BOTTOM).color = targetColor
        Case Else
            '�������Ȃ�
    End Select
End Sub
' �Z���w�i�F��ݒ�/�N���A
Sub setBackgroundColor(startCell As Range, endCell As Range, targetColor As Long)
    Range(startCell, endCell).Interior.color = targetColor
End Sub
Public Sub setFontColor(startCell As Range, endCell As Range, color As Long)
    Dim targetRange As Range
    '�Z���͈͂�ݒ�
    Set targetRange = Range(startCell, endCell)
    targetRange.Font.color = color
End Sub
' �j���Z���̃t�H���g�F��ݒ�
Public Sub setWeekdayColor(targetCell As Range, curDate As Date)

    Select Case Weekday(curDate)
        Case vbSaturday
            targetCell.Font.color = COLOR_BLUE ' �y�j: ��
        Case vbSunday
            targetCell.Font.color = COLOR_RED ' ���j: ��
        Case Else
            'targetCell.Font.color = COLOR_BLACK ' ����: ��
    End Select
End Sub
' �f�[�^�Z���̌r���N���A
Public Sub clearCellBorders(targetCell As Range)
    Dim j As Long
    Dim borderCell As Range

    For j = 0 To NUM_OF_ITEMS + CELL_OFFSET
        Set borderCell = GetCellOffset(targetCell.row, targetCell.Column, j, ALT_DIRECTION)
        targetCell.Borders.LineStyle = xlLineStyleNone
        targetCell.Validation.Delete '���͋K�����N���A
    Next j
End Sub
'�Z���̓��͋K��
Public Sub setRules(startCell As Range, endCell As Range, isEnabled As Boolean)
    Dim targetRange As Range
    If startCell.Address = endCell.Address Then
        Set targetRange = startCell ' �P�ƃZ���̏ꍇ
    Else
        Set targetRange = Range(startCell, endCell) ' �����Z���̏ꍇ
    End If

    If isEnabled Then
        targetRange.Validation.Delete '�N���A���Ă���łȂ��ƃo�O��
        targetRange.Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=0, Formula2:=1
    Else
        targetRange.Validation.Delete '���͋K�����N���A
    End If
End Sub

'�������낦
Public Sub setCenter(startCell As Range, endCell As Range)
    Dim targetRange As Range
    If startCell.Address = endCell.Address Then
        Set targetRange = startCell ' �P�ƃZ���̏ꍇ
    Else
        Set targetRange = Range(startCell, endCell) ' �����Z���̏ꍇ
    End If
    targetRange.HorizontalAlignment = xlCenter
End Sub
'�����낦
Public Sub setBottom(startCell As Range, endCell As Range)
    Dim targetRange As Range
    If startCell.Address = endCell.Address Then
        Set targetRange = startCell ' �P�ƃZ���̏ꍇ
    Else
        Set targetRange = Range(startCell, endCell) ' �����Z���̏ꍇ
    End If
    targetRange.VerticalAlignment = xlBottom
End Sub
'�܂�Ԃ��\��
Public Sub setWrap(startCell As Range, endCell As Range, isEnabled As Boolean)
    Dim targetRange As Range
    If startCell.Address = endCell.Address Then
        Set targetRange = startCell ' �P�ƃZ���̏ꍇ
    Else
        Set targetRange = Range(startCell, endCell) ' �����Z���̏ꍇ
    End If
    targetRange.WrapText = isEnabled
End Sub
Sub ChangeAllCellFontColor(ws As Worksheet, fontColor As Long)
    ws.Cells.Font.color = fontColor
End Sub
