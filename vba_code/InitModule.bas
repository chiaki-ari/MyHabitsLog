Attribute VB_Name = "InitModule"
Private Const MAIN_SHEET_NAME As String = "表の生成"

'年・月の入力位置
Private Const YEAR_LABEL_ROW As Long = 4
Private Const MONTH_LABEL_ROW As Long = 5

'各モードの設定位置
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

'リスト幅の定数
Private Const HORIZONTAL_ITEMS_LIST_WIDTH As Long = 4
Private Const VERTICAL_ITEMS_LIST_WIDTH As Long = 6
Public ITEMS_LIST_WIDTH As Long

'リスト幅の定数
Public Const DATE_LIST_WIDTH As Long = 2
Public Const CELL_OFFSET As Long = 1

'フラグ定義 (ビット管理) MODE_PARAM
Public Const MODE_DIRECTION_HORIZONTAL As Long = &H1   ' 水平: 0001
Public Const MODE_WEEK_AVERAGE As Long = &H2          ' 週平均: 0010
Public Const MODE_WEEK_AVERAGE_GRAPH As Long = &H4    ' 週平均グラフ: 0100

'グローバル変数
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

'初期化メイン関数
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
    'データ開始位置を返す
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
'水平・垂直モード取得
Sub initDirectionMode()
    Select Case Cells(MODE_DIRECTION_ROW, SETTING_COMMON_COLUMN)
        Case "水平"
            MODE_PARAM = MODE_PARAM Or MODE_DIRECTION_HORIZONTAL ' 水平フラグをON
        Case "垂直"
            MODE_PARAM = MODE_PARAM And (Not MODE_DIRECTION_HORIZONTAL) ' 水平フラグをOFF
        Case Else
            ' 何もしない（またはエラー処理）
    End Select
End Sub
'週平均モードの設定
Sub InitWeekAverageMode()
    If Cells(MODE_WEEK_AVERAGE_ROW, SETTING_COMMON_COLUMN) = "ON" Then
        MODE_PARAM = MODE_PARAM Or MODE_WEEK_AVERAGE  ' 週平均フラグをON
    Else
        MODE_PARAM = MODE_PARAM And (Not MODE_WEEK_AVERAGE)  ' 週平均フラグをOFF
    End If
End Sub
'週平均グラフモードの設定
Sub InitWeekAverageGraphMode()
    If Cells(MODE_WEEK_AVERAGE_GRAPH_ROW, SETTING_COMMON_COLUMN) = "ON" Then
        MODE_PARAM = MODE_PARAM Or MODE_WEEK_AVERAGE_GRAPH  ' 週平均グラフフラグをON
    Else
        MODE_PARAM = MODE_PARAM And (Not MODE_WEEK_AVERAGE_GRAPH) ' 週平均グラフフラグをOFF
    End If
End Sub
'年・月の取得
Sub initYearMonthValue()
    YEAR_VALUE = Cells(YEAR_LABEL_ROW, SETTING_COMMON_COLUMN).value
    MONTH_VALUE = Cells(MONTH_LABEL_ROW, SETTING_COMMON_COLUMN).value
End Sub

'項目数取得
Sub initNumberOfItems()
    If Cells(NUM_OF_ITEMS_ROW, SETTING_COMMON_COLUMN) <> "" Then
        NUM_OF_ITEMS = Cells(NUM_OF_ITEMS_ROW, SETTING_COMMON_COLUMN)
    Else
            ' 何もしない（またはエラー処理）
    End If
End Sub
'開始曜日の設定
Sub InitStartWeekday()
    Select Case Cells(MODE_START_WEEKDAY_ROW, SETTING_COMMON_COLUMN)
        Case "月"
            START_WEEKDAY = vbMonday
            LAST_WEEKDAY = vbSunday
        Case "火"
            START_WEEKDAY = vbTuesday
            LAST_WEEKDAY = vbMonday
        Case "水"
            START_WEEKDAY = vbWednesday
            LAST_WEEKDAY = vbTuesday
        Case "木"
            START_WEEKDAY = vbThursday
            LAST_WEEKDAY = vbWednesday
        Case "金"
            START_WEEKDAY = vbFriday
            LAST_WEEKDAY = vbThursday
        Case "土"
            START_WEEKDAY = vbSaturday
            LAST_WEEKDAY = vbFriday
        Case "日"
            START_WEEKDAY = vbSunday
            LAST_WEEKDAY = vbSaturday
        Case Else
            ' 何もしない（またはエラー処理）
    End Select
End Sub
'開始日を設定
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
'最終日を設定
Sub initLastDate()
    Dim endOfMonth As Date
    endOfMonth = DateSerial(YEAR_VALUE, MONTH_VALUE + 1, 0)
    If MODE_PARAM And MODE_WEEK_AVERAGE Then
        endOfMonth = endOfMonth - (Weekday(endOfMonth, LAST_WEEKDAY) - 2)
    End If
    LAST_DATE = CDate(endOfMonth)
End Sub
'日数を設定
Sub initNumOfDays()
    days = LAST_DATE - FIRST_DATE
    NUM_OF_DAYS = days
End Sub
'項目リスト幅を設定
Sub initItemsListWidth()
    If MODE_PARAM And MODE_DIRECTION_HORIZONTAL Then
        ITEMS_LIST_WIDTH = HORIZONTAL_ITEMS_LIST_WIDTH
    Else
        ITEMS_LIST_WIDTH = VERTICAL_ITEMS_LIST_WIDTH

    End If
End Sub
'タイトル幅を設定
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

'タイトル幅を設定
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
