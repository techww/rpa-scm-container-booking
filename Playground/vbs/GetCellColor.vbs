Function GetCellColor(sheetName As String, cellAddress As String) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    GetCellColor = ws.Range(cellAddress).Interior.Color
End Function
