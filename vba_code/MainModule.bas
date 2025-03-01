Attribute VB_Name = "MainModule"
'ï\ê∂ê¨É{É^Éì
Sub createNewSheetButton_Click()
    Dim anchorDataCell As Variant
    Dim anchorRow As Long, anchorColumn As Long
    Dim ws As Worksheet
    dataAnchorCell = InitModule.initData
    anchorRow = dataAnchorCell(0)
    anchorColumn = dataAnchorCell(1)
    Dim result As Boolean
    result = SheetManager.createNewSheet(anchorRow, anchorColumn)
    If result Then
    Else
        With NEW_WORKSHEET
            Call InputHandler.inputDate(anchorRow, anchorColumn)
        End With
    End If
End Sub
