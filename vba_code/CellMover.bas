Attribute VB_Name = "CellMover"
Public Const DIRECTION As Boolean = False
Public Const ALT_DIRECTION As Boolean = True

Public Function GetCellOffset(baseRow As Long, baseCol As Long, offset As Long, isAlternateDirection As Boolean) As Range
    If isAlternateDirection Then
        If MODE_PARAM And MODE_DIRECTION_HORIZONTAL Then
            Set GetCellOffset = Cells(baseRow + offset, baseCol)
        Else
            Set GetCellOffset = Cells(baseRow, baseCol + offset)
        End If
    Else
        If MODE_PARAM And MODE_DIRECTION_HORIZONTAL Then
            Set GetCellOffset = Cells(baseRow, baseCol + offset)
        Else
            Set GetCellOffset = Cells(baseRow + offset, baseCol)
        End If
    End If
End Function

