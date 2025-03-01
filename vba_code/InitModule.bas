Attribute VB_Name = "InitModule"
Private Const MAIN_SHEET_NAME As String = "�\�̐���"

'�N�E���̓��͈ʒu
Private Const YEAR_LABEL_ROW As Long = 4
Private Const MONTH_LABEL_ROW As Long = 5

'�e���[�h�̐ݒ�ʒu
Private Const SETTING_COMMON_COLUMN As Long = 5
Private Const NUM_OF_ITEMS_ROW As Long = 7
Private Const MODE_DIRECTION_ROW As Long = 9
Private Const MODE_WEEK_AVERAGE_ROW As Long = 11
Private Const MODE_START_WEEKDAY_ROW As Long = 12
Private Const MODE_WEEK_AVERAGE_GRAPH_ROW As Long = 13

Private Const HOR_START_ROW As Long = 10
Private Const HOR_START_COLUMN As Long = 6
Private Const VER_START_ROW As Long = 8
Private Const VER_START_COLUMN As Long = 7

'���X�g���̒萔
Private Const HORIZONTAL_ITEMS_LIST_WIDTH As Long = 4
Private Const VERTICAL_ITEMS_LIST_WIDTH As Long = 6
Public ITEMS_LIST_WIDTH As Long

'���X�g���̒萔
Public Const DATE_LIST_WIDTH As Long = 2
Public Const CELL_OFFSET As Long = 1

'�t���O��` (�r�b�g�Ǘ�) MODE_PARAM
Public Const MODE_DIRECTION_HORIZONTAL As Long = &H1   ' ����: 0001
Public Const MODE_WEEK_AVERAGE As Long = &H2          ' �T����: 0010
Public Const MODE_WEEK_AVERAGE_GRAPH As Long = &H4    ' �T���σO���t: 0100

'�O���[�o���ϐ�
Public MODE_PARAM As Long
Public START_WEEKDAY As Long
Public LAST_WEEKDAY As Long
Public YEAR_VALUE As Long
Public MONTH_VALUE As Long
Public FIRST_DATE As Date
Public LAST_DATE As Date
Public NUM_OF_DAYS As Long
Public NUM_OF_ITEMS As Long

Public TITLE_OFFSET_ROW As Long
Public TITLE_OFFSET_COLUMN As Long
Public TITLE_WIDTH As Long
Private Const HOR_TITLE_OFFSET_ROW As Long = 8
Private Const HOR_TITLE_OFFSET_COLUMN As Long = 3
Private Const HOR_TITLE_WIDTH As Long = 1
Private Const VER_TITLE_OFFSET_ROW As Long = 6
Private Const VER_TITLE_OFFSET_COLUMN As Long = 5
Private Const VER_TITLE_WIDTH As Long = 1

Public Const COL_WIDTH As Long = 3
Public COL_WIDTHS(4) As Long
Private Const HOR_COL_WIDTH_1 As Long = 1
Private Const HOR_COL_WIDTH_2 As Long = 2
Private Const HOR_COL_WIDTH_3 As Long = 9
Private Const HOR_COL_WIDTH_4 As Long = 9
Private Const HOR_COL_WIDTH_5 As Long = 2

Private Const VER_COL_WIDTH_1 As Long = 1
Private Const VER_COL_WIDTH_2 As Long = 9
Private Const VER_COL_WIDTH_3 As Long = 9
Private Const VER_COL_WIDTH_4 As Long = 1
Private Const VER_COL_WIDTH_5 As Long = 3

Public Const ROW_HEIGHT  As Long = 20
Public ROW_HEIGHTS(7) As Long
Private Const HOR_ROW_HEIGHT_1 As Long = 10
Private Const HOR_ROW_HEIGHT_2 As Long = 20
Private Const HOR_ROW_HEIGHT_3 As Long = 20
Private Const HOR_ROW_HEIGHT_4 As Long = 20
Private Const HOR_ROW_HEIGHT_5 As Long = 20
Private Const HOR_ROW_HEIGHT_6 As Long = 15
Private Const HOR_ROW_HEIGHT_7 As Long = 15
Private Const HOR_ROW_HEIGHT_8 As Long = 20

Private Const VER_ROW_HEIGHT_1 As Long = 10
Private Const VER_ROW_HEIGHT_2 As Long = 20
Private Const VER_ROW_HEIGHT_3 As Long = 15
Private Const VER_ROW_HEIGHT_4 As Long = 15
Private Const VER_ROW_HEIGHT_5 As Long = 20
Private Const VER_ROW_HEIGHT_6 As Long = 15
Private Const VER_ROW_HEIGHT_7 As Long = 15
Private Const VER_ROW_HEIGHT_8 As Long = 20

'���������C���֐�
Public Function initData() As Variant
    Dim sheetIndex As Long
    sheetIndex = ThisWorkbook.Sheets(MAIN_SHEET_NAME).Index
    With Worksheets(sheetIndex)
        Call initParam
        Call initDirectionMode
        Call initYearMonthValue
        Call initNumberOfItems
        Call InitWeekAverageMode
        Call InitWeekAverageGraphMode
        Call InitStartWeekday
        Call initFisrtDate
        Call initLastDate
        Call initNumOfDays
        Call initItemsListWidth
        Call initTitleWidth
        Call initColRowSize
    End With
    '�f�[�^�J�n�ʒu��Ԃ�
    Dim anchorCell As Variant
    If MODE_PARAM And MODE_DIRECTION_HORIZONTAL Then
        anchorCell = Array(HOR_START_ROW, HOR_START_COLUMN)
    Else
        anchorCell = Array(VER_START_ROW, VER_START_COLUMN)
    End If
    initData = anchorCell
End Function
Sub initParam()
    MODE_PARAM = 0
End Sub
'�����E�������[�h�擾
Sub initDirectionMode()
    Select Case Cells(MODE_DIRECTION_ROW, SETTING_COMMON_COLUMN)
        Case "����"
            MODE_PARAM = MODE_PARAM Or MODE_DIRECTION_HORIZONTAL ' �����t���O��ON
        Case "����"
            MODE_PARAM = MODE_PARAM And (Not MODE_DIRECTION_HORIZONTAL) ' �����t���O��OFF
        Case Else
            ' �������Ȃ��i�܂��̓G���[�����j
    End Select
End Sub
'�T���σ��[�h�̐ݒ�
Sub InitWeekAverageMode()
    If Cells(MODE_WEEK_AVERAGE_ROW, SETTING_COMMON_COLUMN) = "ON" Then
        MODE_PARAM = MODE_PARAM Or MODE_WEEK_AVERAGE  ' �T���σt���O��ON
    Else
        MODE_PARAM = MODE_PARAM And (Not MODE_WEEK_AVERAGE)  ' �T���σt���O��OFF
    End If
End Sub
'�T���σO���t���[�h�̐ݒ�
Sub InitWeekAverageGraphMode()
    If Cells(MODE_WEEK_AVERAGE_GRAPH_ROW, SETTING_COMMON_COLUMN) = "ON" Then
        MODE_PARAM = MODE_PARAM Or MODE_WEEK_AVERAGE_GRAPH  ' �T���σO���t�t���O��ON
    Else
        MODE_PARAM = MODE_PARAM And (Not MODE_WEEK_AVERAGE_GRAPH) ' �T���σO���t�t���O��OFF
    End If
End Sub
'�N�E���̎擾
Sub initYearMonthValue()
    YEAR_VALUE = Cells(YEAR_LABEL_ROW, SETTING_COMMON_COLUMN).value
    MONTH_VALUE = Cells(MONTH_LABEL_ROW, SETTING_COMMON_COLUMN).value
End Sub

'���ڐ��擾
Sub initNumberOfItems()
    If Cells(NUM_OF_ITEMS_ROW, SETTING_COMMON_COLUMN) <> "" Then
        NUM_OF_ITEMS = Cells(NUM_OF_ITEMS_ROW, SETTING_COMMON_COLUMN)
    Else
            ' �������Ȃ��i�܂��̓G���[�����j
    End If
End Sub
'�J�n�j���̐ݒ�
Sub InitStartWeekday()
    Select Case Cells(MODE_START_WEEKDAY_ROW, SETTING_COMMON_COLUMN)
        Case "��"
            START_WEEKDAY = vbMonday
            LAST_WEEKDAY = vbSunday
        Case "��"
            START_WEEKDAY = vbTuesday
            LAST_WEEKDAY = vbMonday
        Case "��"
            START_WEEKDAY = vbWednesday
            LAST_WEEKDAY = vbTuesday
        Case "��"
            START_WEEKDAY = vbThursday
            LAST_WEEKDAY = vbWednesday
        Case "��"
            START_WEEKDAY = vbFriday
            LAST_WEEKDAY = vbThursday
        Case "�y"
            START_WEEKDAY = vbSaturday
            LAST_WEEKDAY = vbFriday
        Case "��"
            START_WEEKDAY = vbSunday
            LAST_WEEKDAY = vbSaturday
        Case Else
            ' �������Ȃ��i�܂��̓G���[�����j
    End Select
End Sub
'�J�n����ݒ�
Sub initFisrtDate()
    Dim firstDate As Date
    firstDate = DateSerial(YEAR_VALUE, MONTH_VALUE, 1)
    If MODE_PARAM And MODE_WEEK_AVERAGE Then
        If Weekday(curDate) <> START_WEEKDAY Then
            firstDate = firstDate - (Weekday(firstDate, START_WEEKDAY + 1))
        End If
    End If
    FIRST_DATE = CDate(firstDate)
End Sub
'�ŏI����ݒ�
Sub initLastDate()
    Dim endOfMonth As Date
    endOfMonth = DateSerial(YEAR_VALUE, MONTH_VALUE + 1, 0)
    If MODE_PARAM And MODE_WEEK_AVERAGE Then
        endOfMonth = endOfMonth - (Weekday(endOfMonth, LAST_WEEKDAY) - 2)
    End If
    LAST_DATE = CDate(endOfMonth)
End Sub
'������ݒ�
Sub initNumOfDays()
    days = LAST_DATE - FIRST_DATE
    NUM_OF_DAYS = days
End Sub
'���ڃ��X�g����ݒ�
Sub initItemsListWidth()
    If MODE_PARAM And MODE_DIRECTION_HORIZONTAL Then
        ITEMS_LIST_WIDTH = HORIZONTAL_ITEMS_LIST_WIDTH
    Else
        ITEMS_LIST_WIDTH = VERTICAL_ITEMS_LIST_WIDTH

    End If
End Sub
'�^�C�g������ݒ�
Sub initTitleWidth()
    If MODE_PARAM And MODE_DIRECTION_HORIZONTAL Then
        TITLE_OFFSET_ROW = HOR_TITLE_OFFSET_ROW
        TITLE_OFFSET_COLUMN = HOR_TITLE_OFFSET_COLUMN
        TITLE_WIDTH = HOR_TITLE_WIDTH
    Else
        TITLE_OFFSET_ROW = VER_TITLE_OFFSET_ROW
        TITLE_OFFSET_COLUMN = VER_TITLE_OFFSET_COLUMN
        TITLE_WIDTH = VER_TITLE_WIDTH

    End If
End Sub

'�^�C�g������ݒ�
Sub initColRowSize()
    If MODE_PARAM And MODE_DIRECTION_HORIZONTAL Then
        COL_WIDTHS(0) = HOR_COL_WIDTH_1
        COL_WIDTHS(1) = HOR_COL_WIDTH_2
        COL_WIDTHS(2) = HOR_COL_WIDTH_3
        COL_WIDTHS(3) = HOR_COL_WIDTH_4
        COL_WIDTHS(4) = HOR_COL_WIDTH_5
        
        ROW_HEIGHTS(0) = HOR_ROW_HEIGHT_1
        ROW_HEIGHTS(1) = HOR_ROW_HEIGHT_2
        ROW_HEIGHTS(2) = HOR_ROW_HEIGHT_3
        ROW_HEIGHTS(3) = HOR_ROW_HEIGHT_4
        ROW_HEIGHTS(4) = HOR_ROW_HEIGHT_5
        ROW_HEIGHTS(5) = HOR_ROW_HEIGHT_6
        ROW_HEIGHTS(6) = HOR_ROW_HEIGHT_7
        ROW_HEIGHTS(7) = HOR_ROW_HEIGHT_8
    Else
        COL_WIDTHS(0) = VER_COL_WIDTH_1
        COL_WIDTHS(1) = VER_COL_WIDTH_2
        COL_WIDTHS(2) = VER_COL_WIDTH_3
        COL_WIDTHS(3) = VER_COL_WIDTH_4
        COL_WIDTHS(4) = VER_COL_WIDTH_5
        
        ROW_HEIGHTS(0) = VER_ROW_HEIGHT_1
        ROW_HEIGHTS(1) = VER_ROW_HEIGHT_2
        ROW_HEIGHTS(2) = VER_ROW_HEIGHT_3
        ROW_HEIGHTS(3) = VER_ROW_HEIGHT_4
        ROW_HEIGHTS(4) = VER_ROW_HEIGHT_5
        ROW_HEIGHTS(5) = VER_ROW_HEIGHT_6
        ROW_HEIGHTS(6) = VER_ROW_HEIGHT_7
        ROW_HEIGHTS(7) = VER_ROW_HEIGHT_8

    End If
End Sub
