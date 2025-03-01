Attribute VB_Name = "SheetManager"
Public NEW_WORKSHEET As Worksheet

Public Const DONE As Boolean = False
Public Const NG As Boolean = True


Public Function createNewSheet(anchorCellRow As Long, anchorCellColumn As Long) As Boolean
    Dim result As Boolean
    result = setNameSheet
    If result Then
        createNewSheet = NG
    Else
        Dim anchorCell As Range
        Set anchorCell = Cells(anchorCellRow, anchorCellColumn)
        NEW_WORKSHEET.Activate
        Call CellFormatter.ChangeAllCellFontColor(NEW_WORKSHEET, COLOR_DARK_GRAY)
        ActiveWindow.DisplayGridlines = False
        Call CellFormatter.ChangeSheetFontSize(NEW_WORKSHEET, FONT_SIZE_DEFAULT)
        Call setCellColWidth
        Call setCellRowHeight
        Call setFreeze(anchorCell)
        createNewSheet = DONE
    End If
End Function

Function setNameSheet() As Boolean
    Dim newWorkSheet
    Dim nameDate As String
    ' 始まりの日を設定
    nameDate = Format(DateSerial(YEAR_VALUE, MONTH_VALUE, 1), "YYYYMM")
    
    '既存のシートがあるかチェック
    If sheetExists(nameDate) Then
        MsgBox "シート [" & nameDate & "] は既に存在します！", vbCritical, "エラー"
        setNameSheet = NG
        Exit Function
    End If
    '新しいワークシートを追加
    Set NEW_WORKSHEET = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    NEW_WORKSHEET.Name = nameDate
    
End Function

Sub setCellColWidth()
    Dim colWidth As Long
    Dim colWidths As Variant
    Cells.ColumnWidth = COL_WIDTH
    colWidths = COL_WIDTHS
    Dim i As Long
    For i = LBound(colWidths) To UBound(colWidths)
        Columns(i + 1).ColumnWidth = colWidths(i)
    Next i
End Sub
Sub setCellRowHeight()
    Dim rowHeight As Long
    Dim rowHeights As Variant
    Cells.rowHeight = ROW_HEIGHT
    rowHeights = ROW_HEIGHTS
    Dim i As Long
    For i = LBound(rowHeights) To UBound(rowHeights)
        Rows(i + 1).rowHeight = rowHeights(i)
    Next i
End Sub
Sub setFreeze(anchorDataCell As Range)
    Dim freezeRow As Long
    Dim freezeColumn As Long
    ' ウィンドウ枠を固定
    freezeRow = anchorDataCell.row - CELL_OFFSET
    freezeColumn = anchorDataCell.Column - CELL_OFFSET
    NEW_WORKSHEET.Activate
    NEW_WORKSHEET.Cells(freezeRow + 1, freezeColumn + 1).Select
    ActiveWindow.FreezePanes = True
End Sub
Function sheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    sheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            sheetExists = True
            Exit Function
        End If
    Next ws
End Function
