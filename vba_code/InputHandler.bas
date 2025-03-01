Attribute VB_Name = "InputHandler"
Private Const LABEL_MOTHLY_SUM As String = "累計"
Private Const LABEL_DAILY_SUM As String = "日計"
Private Const LABEL_WEEKLY_SUM As String = "週計"

Private Const LABEL_GOAL As String = "目標"
Private Const LABEL_TERM As String = "期間"
Private Const LABEL_GOAL_SAMPLE As String = "例)習慣つける！"

Private Const LABEL_ITEMS_LIST As String = "行動目標"
Private Const LABEL_ITEMS_LIST_POINT As String = "※達成可能性80%以上"

Private Const LABEL_ITEM_SAMPLE1 As String = "例)朝7時起床"
Private Const LABEL_ITEM_SAMPLE2 As String = "例)瞑想"
Private Const LABEL_ITEM_SAMPLE3 As String = "例)腹筋10回"

Private Const ITEMS_LIST_TITLE_OFFSET_ROW As Long = 2


Public Const STATE_EMPTY As Long = &H0             ' 0000 すべて空
Public Const STATE_DATE As Long = &H1              ' 0001 日付あり
Public Const STATE_DATA As Long = &H2              ' 0010 データあり
Public Const STATE_ITEM As Long = &H4              ' 0100 項目あり
Public Const STATE_DATE_AND_DATA As Long = STATE_DATE Or STATE_DATA  ' 0011 日付 + データ
Public Const STATE_DATE_AND_ITEM As Long = STATE_DATE Or STATE_ITEM  ' 0101 日付 + 項目
Public Const STATE_ITEM_AND_DATA As Long = STATE_ITEM Or STATE_DATA  ' 0110 項目 + データ
Public Const STATE_ALL_FILLED As Long = STATE_DATE Or STATE_DATA Or STATE_ITEM  ' 0111 全部あり


Public Sub inputDate(anchorCellRow As Long, anchorCellColumn As Long)
    Dim rslt As VbMsgBoxResult
    Dim startDateCell As Range
    Dim dataState As Long
    dataState = getDataState(Cells(anchorCellRow, anchorCellColumn))

    Select Case True
    Case dataState = STATE_EMPTY
        Call writeDate(anchorCellRow, anchorCellColumn)
        Call writeItemList(anchorCellRow, anchorCellColumn)
        Call writeDataArea(anchorCellRow, anchorCellColumn)
        Call writeDailySum(anchorCellRow, anchorCellColumn)
        Call writeMonthlySum(anchorCellRow, anchorCellColumn)
        Call writeTitle(anchorCellRow, anchorCellColumn)
        Call writeItemTitleList(anchorCellRow, anchorCellColumn)
        If MODE_PARAM And MODE_WEEK_AVERAGE Then
            Call writeWeeklySum(anchorCellRow, anchorCellColumn)
        End If
        ' 完了メッセージ
        MsgBox (YEAR_VALUE & "年" & MONTH_VALUE & "月の日付を自動入力しました！")
        
    Case (dataState + STATE_DATE) <> 0
    Case (dataState + STATE_DATA) <> 0
    Case (dataState + STATE_ITEM) <> 0
        MsgBox ("先にクリアしてください。")
    End Select

End Sub

Function getDataState(baseCell As Range) As Long
    Dim anchorDateCell As Range, anchorDataCell As Range, anchorItemsListCell As Range
    Dim endDateCell As Range, endDataCell As Range, endItemsListCell As Range
    Dim result As Long
    
    Set anchorDateCell = CellMover.GetCellOffset(baseCell.row, baseCell.Column, -DATE_LIST_WIDTH, ALT_DIRECTION)
    Set endDateCell = CellMover.GetCellOffset(anchorDateCell.row, anchorDateCell.Column, NUM_OF_DAYS + CELL_OFFSET, DIRECTION)
    Set endDateCell = CellMover.GetCellOffset(endDateCell.row, endDateCell.Column, DATE_LIST_WIDTH - CELL_OFFSET, ALT_DIRECTION)
    
    Set anchorDataCell = baseCell
    Set endDataCell = CellMover.GetCellOffset(anchorDataCell.row, anchorDataCell.Column, NUM_OF_DAYS + CELL_OFFSET, DIRECTION)
    Set endDataCell = CellMover.GetCellOffset(endDataCell.row, endDataCell.Column, NUM_OF_ITEMS - CELL_OFFSET, ALT_DIRECTION)

    Set anchorItemsListCell = CellMover.GetCellOffset(baseCell.row, baseCell.Column, -ITEMS_LIST_WIDTH, DIRECTION)
    Set endItemsListCell = CellMover.GetCellOffset(anchorItemsListCell.row, anchorItemsListCell.Column, NUM_OF_ITEMS, ALT_DIRECTION)
    Set endItemsListCell = CellMover.GetCellOffset(endItemsListCell.row, endItemsListCell.Column, ITEMS_LIST_WIDTH - CELL_OFFSET, DIRECTION)

 
    If checkEmpty(anchorDateCell, endDateCell) Then
        result = result + STATE_DATE
    End If
    If checkEmpty(anchorDataCell, endDataCell) Then
        result = result + STATE_DATA
    End If
    If checkEmpty(anchorItemsListCell, endItemsListCell) Then
        result = result + STATE_ITEM
    End If
    
    getDataState = result
    
End Function


Function checkEmpty(baseCell As Range, endCell As Range) As Boolean
    Dim rng As Range
    Dim cell As Range

    Set rng = Range(baseCell, endCell)

    ' 範囲内のセルをチェック
    For Each cell In rng
        If Not IsEmpty(cell.value) Then ' 空でないセルが1つでもあればTrue
            checkEmpty = True
            Exit Function ' 早期リターン
        End If
    Next cell

    ' すべてのセルが空の場合はFalse
    checkEmpty = False
End Function

Sub writeDate(startRow As Long, startColumn As Long)
    Dim i As Long
    Dim curDate As Date
    Dim cellStartDate As Range, currentDateCell As Range, currentWeekdayCell As Range
    Dim startCell As Range, endCell As Range
    curDate = FIRST_DATE
    Set cellStartDate = CellMover.GetCellOffset(startRow, startColumn, -DATE_LIST_WIDTH, ALT_DIRECTION)
    
    ' 日付と曜日
    For i = 0 To NUM_OF_DAYS - CELL_OFFSET
        Set currentDateCell = CellMover.GetCellOffset(cellStartDate.row, cellStartDate.Column, i, DIRECTION)
        Set currentWeekdayCell = getWeekdayCell(currentDateCell.row, currentDateCell.Column, CELL_OFFSET)
        ' 日付と曜日を入力
        Call writeCellData(currentDateCell, currentWeekdayCell, curDate)
        curDate = curDate + 1
    Next i

    ' 日付の罫線/背景色を設定
    Set startCell = CellMover.GetCellOffset(startRow, startColumn, -DATE_LIST_WIDTH, ALT_DIRECTION)
    Set endCell = CellMover.GetCellOffset(startRow, startColumn, -CELL_OFFSET, ALT_DIRECTION)
    Set endCell = CellMover.GetCellOffset(endCell.row, endCell.Column, NUM_OF_DAYS, DIRECTION)
    Call setBackgroundColorBorders(startCell, endCell, STATE_HIGHLIGHT_BORDER_ALL)
End Sub
' 習慣項目欄タイトル
Sub writeItemTitleList(startRow As Long, startColumn As Long)
    Dim baseCell As Range, startCell As Range, i As Long, endCell As Range
      
    Set startCell = CellMover.GetCellOffset(startRow, startColumn, -ITEMS_LIST_WIDTH, DIRECTION)
    Set startCell = CellMover.GetCellOffset(startCell.row, startCell.Column, -ITEMS_LIST_TITLE_OFFSET_ROW, ALT_DIRECTION)
    Set baseCell = startCell
    If MODE_PARAM And MODE_DIRECTION_HORIZONTAL Then
        Set endCell = CellMover.GetCellOffset(startCell.row, startCell.Column, ITEMS_LIST_WIDTH - CELL_OFFSET, DIRECTION)
        Call CellFormatter.mergeCells(startCell, endCell, MERGE)
        Call CellFormatter.setWrap(startCell, endCell, WRAP)
        Call CellFormatter.setBottom(startCell, startCell)
        startCell.value = LABEL_ITEMS_LIST
        
        Set startCell = CellMover.GetCellOffset(baseCell.row, baseCell.Column, CELL_OFFSET, ALT_DIRECTION)
        Set endCell = CellMover.GetCellOffset(endCell.row, endCell.Column, CELL_OFFSET, ALT_DIRECTION)
        Call CellFormatter.mergeCells(startCell, endCell, MERGE)
        Call CellFormatter.setWrap(startCell, endCell, WRAP)
        startCell.value = LABEL_ITEMS_LIST_POINT
        Call CellFormatter.ChangeRangeFontSize(startCell, startCell, 7)
    Call CellFormatter.setFontColor(startCell, startCell, COLOR_GRAY)
    Else
        Set endCell = CellMover.GetCellOffset(startCell.row, startCell.Column, DATE_LIST_WIDTH - CELL_OFFSET, ALT_DIRECTION)
        Set endCell = CellMover.GetCellOffset(endCell.row, endCell.Column, (ITEMS_LIST_WIDTH / 2) - CELL_OFFSET, DIRECTION)
        Call CellFormatter.mergeCells(startCell, endCell, MERGE)
        Call CellFormatter.setWrap(startCell, endCell, WRAP)
        Call CellFormatter.setBottom(startCell, startCell)
        startCell.value = LABEL_ITEMS_LIST
        
        Set startCell = CellMover.GetCellOffset(baseCell.row, baseCell.Column, ITEMS_LIST_WIDTH / 2, DIRECTION)
        Set endCell = CellMover.GetCellOffset(startCell.row, startCell.Column, DATE_LIST_WIDTH - CELL_OFFSET, ALT_DIRECTION)
        Set endCell = CellMover.GetCellOffset(endCell.row, endCell.Column, (ITEMS_LIST_WIDTH / 2) - CELL_OFFSET, DIRECTION)
        Call CellFormatter.mergeCells(startCell, endCell, MERGE)
        Call CellFormatter.setWrap(startCell, endCell, WRAP)
        startCell.value = LABEL_ITEMS_LIST_POINT
        Call CellFormatter.ChangeRangeFontSize(startCell, startCell, 7)
        

    Call CellFormatter.setFontColor(startCell, startCell, COLOR_GRAY)

    End If

    
    Call CellFormatter.setCenter(baseCell, endCell)
    Call setBackgroundColorBorders(baseCell, endCell, STATE_HIGHLIGHT_OUTER_ONLY)
End Sub
' 習慣項目欄の罫線/背景色を設定
Sub writeItemList(startRow As Long, startColumn As Long)
    Dim currentCell As Range, currentNextCell As Range, targetCell As Range, startCell As Range, i As Long, endCell As Range
      
    Set startCell = CellMover.GetCellOffset(startRow, startColumn, -ITEMS_LIST_WIDTH, DIRECTION)
    Set endCell = CellMover.GetCellOffset(startRow, startColumn, NUM_OF_ITEMS - CELL_OFFSET, ALT_DIRECTION)
    Set endCell = CellMover.GetCellOffset(endCell.row, endCell.Column, -CELL_OFFSET, DIRECTION)
    If MODE_PARAM And MODE_DIRECTION_HORIZONTAL Then
        Call setBackgroundColorBorders(startCell, endCell, STATE_NO_BACKGROUND_BORDER_HORIZONTAL)

    Else
        Call setBackgroundColorBorders(startCell, endCell, STATE_NO_BACKGROUND_BORDER_VERTICAL)

    End If
        For i = 0 To NUM_OF_ITEMS - CELL_OFFSET
        Set currentCell = CellMover.GetCellOffset(startCell.row, startCell.Column, i, ALT_DIRECTION)
        Set currentNextCell = CellMover.GetCellOffset(currentCell.row, currentCell.Column, CELL_OFFSET, DIRECTION)
        Set targetCell = CellMover.GetCellOffset(currentCell.row, currentCell.Column, ITEMS_LIST_WIDTH - CELL_OFFSET, DIRECTION)
        Call CellFormatter.mergeCells(currentNextCell, targetCell, MERGE)
        Call CellFormatter.setWrap(currentNextCell, targetCell, WRAP)
        Call CellFormatter.ChangeRangeFontSize(currentCell, currentCell, FONT_SIZE_SMALL)
        currentCell.value = i + 1
        Call CellFormatter.setCenter(currentCell, currentCell)
    Next i
    'Set currentCell = startCell
    'Set currentNextCell = CellMover.GetCellOffset(currentCell.row, currentCell.Column, CELL_OFFSET, DIRECTION)
    'currentNextCell.value = LABEL_ITEM_SAMPLE1
    'Set currentNextCell = CellMover.GetCellOffset(currentNextCell.row, currentNextCell.Column, CELL_OFFSET, ALT_DIRECTION)
    'currentNextCell.value = LABEL_ITEM_SAMPLE2
    'Set currentNextCell = CellMover.GetCellOffset(currentNextCell.row, currentNextCell.Column, CELL_OFFSET, ALT_DIRECTION)
    'currentNextCell.value = LABEL_ITEM_SAMPLE3
    
End Sub
' データ欄の罫線/背景色を設定 'データ欄の入力規則
Sub writeDataArea(startRow As Long, startColumn As Long)
    Dim startCell As Range, endCell As Range
    Set startCell = Cells(startRow, startColumn)
    Set endCell = CellMover.GetCellOffset(startRow, startColumn, NUM_OF_ITEMS - CELL_OFFSET, ALT_DIRECTION)
    Set endCell = CellMover.GetCellOffset(endCell.row, endCell.Column, NUM_OF_DAYS - CELL_OFFSET, DIRECTION)
    Call setBackgroundColorBorders(startCell, endCell, STATE_NO_BACKGROUND_BORDER_ALL)
    Call CellFormatter.setRules(startCell, endCell, RULE_ON)
    Call CellFormatter.ChangeRangeFontSize(startCell, endCell, FONT_SIZE_VERY_SMALL)
    Call CellFormatter.setCenter(startCell, endCell)
End Sub
'デイリー合計
Sub writeDailySum(startRow As Long, startColumn As Long)
    Dim i As Long, baseCell As Range, currentCell As Range, targetCell As Range, startCell As Range, endCell As Range
    Set baseCell = CellMover.GetCellOffset(startRow, startColumn, NUM_OF_ITEMS, ALT_DIRECTION)
    Set startCell = CellMover.GetCellOffset(startRow, startColumn, -ITEMS_LIST_WIDTH, DIRECTION)
    Set startCell = CellMover.GetCellOffset(startCell.row, startCell.Column, NUM_OF_ITEMS, ALT_DIRECTION)
    Set endCell = CellMover.GetCellOffset(startCell.row, startCell.Column, ITEMS_LIST_WIDTH - CELL_OFFSET, DIRECTION)
    
    Call CellFormatter.mergeCells(startCell, endCell, MERGE)
    Call CellFormatter.setWrap(startCell, endCell, WRAP)
    startCell.value = LABEL_DAILY_SUM
    Call CellFormatter.setCenter(startCell, endCell)
    Call setBackgroundColorBorders(startCell, endCell, STATE_HIGHLIGHT_BORDER_ALL)
    
    For i = 0 To NUM_OF_DAYS - CELL_OFFSET
        Set currentCell = CellMover.GetCellOffset(baseCell.row, baseCell.Column, i, DIRECTION)
        If MODE_PARAM And MODE_DIRECTION_HORIZONTAL Then
            currentCell.Formula = "=SUM(R[-" & NUM_OF_ITEMS & "]C:R[-1]C)"
        Else
            currentCell.Formula = "=SUM(RC[-" & NUM_OF_ITEMS & "]:RC[-1])"
        End If
        
        currentCell.HorizontalAlignment = xlCenter
    Next i
    Call setBackgroundColorBorders(baseCell, currentCell, STATE_NO_BACKGROUND_BORDER_ALL)
End Sub

'週単位合計
Sub writeWeeklySum(startRow As Long, startColumn As Long)
    Dim i As Long, j As Long
    Dim baseCell As Range
    Dim startCell As Range, endCell As Range, targetDateCell As Range
    
    Dim currentCell As Range, targetCell As Range
    Set targetDataCell = Cells(startRow, startColumn)
    
    Set baseCell = CellMover.GetCellOffset(startRow, startColumn, NUM_OF_ITEMS + CELL_OFFSET, ALT_DIRECTION)
    
    Set startCell = CellMover.GetCellOffset(baseCell.row, baseCell.Column, -ITEMS_LIST_WIDTH, DIRECTION)
    Set endCell = CellMover.GetCellOffset(startCell.row, startCell.Column, ITEMS_LIST_WIDTH - CELL_OFFSET, DIRECTION)
    Call CellFormatter.mergeCells(startCell, endCell, MERGE)
    Call CellFormatter.setWrap(startCell, endCell, WRAP)
    startCell.value = LABEL_WEEKLY_SUM
    Call CellFormatter.setCenter(startCell, endCell)
    Call setBackgroundColorBorders(startCell, endCell, STATE_HIGHLIGHT_BORDER_ALL)
    Set startCell = baseCell
    Set currentCell = startCell
    Set targetCell = startCell
    Set targetDateCell = CellMover.GetCellOffset(targetDataCell.row, targetDataCell.Column, -DATE_LIST_WIDTH, ALT_DIRECTION)
    For i = CELL_OFFSET To NUM_OF_DAYS
        If i Mod 7 = 0 Then
            Call CellFormatter.mergeCells(targetCell, currentCell, MERGE)
            If MODE_PARAM And MODE_DIRECTION_HORIZONTAL Then
                targetCell.Formula = "=SUM(R[-1]C:R[-1]C[6])"
            Else
                targetCell.Formula = "=SUM(RC[-1]:R[6]C[-1])"
            End If
            
            Call CellFormatter.setCenter(targetCell, currentCell)
            Call setBackgroundColorBorders(targetDateCell, currentCell, STATE_NO_BACKGROUND_OUTER_ONLY)
            Set targetCell = CellMover.GetCellOffset(currentCell.row, currentCell.Column, CELL_OFFSET, DIRECTION)
            Set targetDateCell = CellMover.GetCellOffset(targetDateCell.row, targetDateCell.Column, 7, DIRECTION)
            'Set targetDateCell = CellMover.GetCellOffset(targetDateCell.row, targetDateCell.Column, 7, ALT_DIRECTION)'広がるバグ
        End If
        Set currentCell = CellMover.GetCellOffset(currentCell.row, currentCell.Column, CELL_OFFSET, DIRECTION)
    Next i
End Sub

'月単位合計
Sub writeMonthlySum(startRow As Long, startColumn As Long)
    Dim i As Long, startCell As Range, currentCell As Range, endCell As Range
    Set startCell = CellMover.GetCellOffset(startRow, startColumn, NUM_OF_DAYS, DIRECTION)
    Set currentCell = CellMover.GetCellOffset(startCell.row, startCell.Column, -DATE_LIST_WIDTH, ALT_DIRECTION)
    currentCell.value = LABEL_MOTHLY_SUM
    Set endCell = CellMover.GetCellOffset(currentCell.row, currentCell.Column, CELL_OFFSET, ALT_DIRECTION)
    Call CellFormatter.setCenter(currentCell, currentCell)
    Call CellFormatter.mergeCells(currentCell, endCell, MERGE)
    Call setBackgroundColorBorders(currentCell, endCell, STATE_HIGHLIGHT_BORDER_ALL)
    Call CellFormatter.setWrap(currentCell, endCell, WRAP)
    
    Set endCell = CellMover.GetCellOffset(currentCell.row, currentCell.Column, CELL_OFFSET, ALT_DIRECTION)
    For i = 0 To NUM_OF_ITEMS
        Set currentCell = CellMover.GetCellOffset(startCell.row, startCell.Column, i, ALT_DIRECTION)
        If MODE_PARAM And MODE_DIRECTION_HORIZONTAL Then
            currentCell.Formula = "=SUM(RC[-" & NUM_OF_DAYS & "]:RC[-1])"
        Else
            currentCell.Formula = "=SUM(R[-" & NUM_OF_DAYS & "]C:R[-1]C)"
        End If
        currentCell.HorizontalAlignment = xlCenter
    Next i
    Call setBackgroundColorBorders(startCell, currentCell, STATE_NO_BACKGROUND_BORDER_ALL)
    Call setBackgroundColorBorders(currentCell, currentCell, STATE_HIGHLIGHT_BORDER_ALL)
End Sub


Sub writeCellData(cellDate As Range, cellWeekday As Range, curDate As Date)
    ' 日付を入力
    cellDate.value = Format(curDate, "dd")
    cellDate.HorizontalAlignment = xlCenter

    ' 曜日を入力
    cellWeekday.value = Format(curDate, "aaa")
    cellWeekday.HorizontalAlignment = xlCenter

    ' 曜日の色を設定
    Call setWeekdayColor(cellWeekday, curDate)
    Call CellFormatter.ChangeRangeFontSize(cellDate, cellWeekday, FONT_SIZE_SMALL)
End Sub


Function getWeekdayCell(baseRow As Long, baseCol As Long, offset As Long) As Range
    Set getWeekdayCell = CellMover.GetCellOffset(baseRow, baseCol, offset, ALT_DIRECTION)
End Function


'タイトル入力
Sub writeTitle(startRow As Long, startColumn As Long)
    Dim startCell As Range, i As Long, endCell As Range
      
    Set startCell = Cells(startRow - TITLE_OFFSET_ROW, startColumn - TITLE_OFFSET_COLUMN)
    
    Set endCell = Cells(startCell.row, startCell.Column + TITLE_WIDTH)
    Call CellFormatter.mergeCells(startCell, endCell, MERGE)
    Call CellFormatter.setCenter(startCell, endCell)
    Call setBackgroundColorBorders(startCell, endCell, STATE_HIGHLIGHT_NO_BORDER)
    startCell.value = LABEL_GOAL
    
    Set startCell = Cells(startCell.row + CELL_OFFSET, startCell.Column)
    Set endCell = Cells(startCell.row + CELL_OFFSET, startCell.Column + TITLE_WIDTH)
    Call CellFormatter.mergeCells(startCell, endCell, MERGE)
    Call CellFormatter.setCenter(startCell, endCell)
    startCell.value = LABEL_GOAL_SAMPLE
    
    Set startCell = Cells(startCell.row + CELL_OFFSET + CELL_OFFSET, startCell.Column)
    Set endCell = Cells(startCell.row, startCell.Column + TITLE_WIDTH)
    Call CellFormatter.mergeCells(startCell, endCell, MERGE)
    Call CellFormatter.setCenter(startCell, endCell)
    Call setBackgroundColorBorders(startCell, endCell, STATE_HIGHLIGHT_NO_BORDER)
    startCell.value = LABEL_TERM
    
    Set startCell = Cells(startCell.row + CELL_OFFSET, startCell.Column)
    Set endCell = Cells(startCell.row + CELL_OFFSET, startCell.Column + TITLE_WIDTH)
    Call CellFormatter.mergeCells(startCell, endCell, MERGE)
    Call CellFormatter.setCenter(startCell, endCell)
    startCell.value = Format(DateSerial(YEAR_VALUE, MONTH_VALUE, 1), "YYYY年MM月")
End Sub

Sub renameSheet()
    Dim sheetName As String
    Dim nameDate As Date
    ' 始まりの日を設定
    nameDate = DateSerial(YEAR_VALUE, MONTH_VALUE, 1)
    sheetName = ActiveSheet.Name
    'シート名変更
    Worksheets(sheetName).Name = Format(nameDate, "yyyymm")

End Sub
'　背景/枠線セット
Sub setBackgroundColorBorders(startCell As Range, endCell As Range, cellDecoType As Long)
    
    Call CellFormatter.setDecorateCell(startCell, endCell, cellDecoType)
End Sub


