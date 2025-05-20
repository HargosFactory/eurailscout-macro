' Description: This function defines the last row of a worksheet.
' Parameters:
'   - ws: The worksheet to analyze.
'   - column: The column to analyze.
' Returns: The last row of the worksheet.
Function define_end_sheet(ws As Worksheet, column As Range) As Integer
    define_end_sheet = ws.Cells(ws.Rows.Count, column.column).End(xlUp).row
End Function

' Description: This function searches for a value in a column.
' Parameters:
'   - ws: The worksheet to analyze.
'   - column: The column to analyze.
'   - Value: The value to search for.
' Returns: The row number of the value.
Function search_end_point_in_column(ws As Worksheet, column As Range, value As String) As Long
    Dim i As Long
    Dim end_ws As Long
    end_ws = define_end_sheet(ws, column)
    For i = 1 To end_ws
        If ws.Cells(i, column.column).value = value Then
            search_end_point_in_column = i
            Exit For
        End If
    Next i
End Function

' Description: This function gets the line number of a value in a column.
' Parameters:
'   - ws: The worksheet to analyze.
'   - col: The column to search.
'   - val: The value to search for.
' Returns: The line number of the value.
Function get_line_nb(ws As Worksheet, col As Integer, val As Variant) As Long
    Dim rng As Range
    Dim line As Long
    
    Set rng = ws.Columns(col).Find(What:=val, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rng Is Nothing Then
        line = rng.row
    Else
        line = 0
    End If
    
    get_line_nb = line
End Function

' Description: This function gets the value of a cell.
' Parameters:
'   - ws: The worksheet to analyze.
'   - row: The row of the cell.
'   - column: The column of the cell.
' Returns: The value of the cell.
Function get_value_of_cell(ws As Worksheet, row As Integer, column As Integer) As Variant
    get_value_of_cell = ws.Cells(row, column).value
End Function

' Description: This function write the value in a cell.
' Parameters:
'   - ws: The worksheet to analyze.
'   - line: The line of the cell.
'   - column: The column of the cell.
'   - value: The value to write.
' Returns: True if the value was written successfully, False otherwise.
Function report_value(ws As Worksheet, ByVal line As Long, ByVal column As Long, ByVal value As String) As Boolean
    Dim cell As Range

    Set cell = ws.Cells(line, column)
    If cell.value = "" Then
        cell.value = value
        report_value = True
    Else
        report_value = False
    End If
End Function

' Description: This function checks if a pattern is present in a text.
' Parameters:
'   - text: The text to search.
'   - pattern: The pattern to search for.
' Returns: True if the pattern is found, False otherwise.
Function regex(text As String, pattern As String) As Boolean
    Dim regexObj As Object
    Set regexObj = CreateObject("VBScript.RegExp")
    regexObj.Global = True
    regexObj.pattern = pattern
    regex = regexObj.test(text)
End Function

' Description: This function defines the last row of a worksheet table section.
' Parameters:
'   - ws: The worksheet to analyze.
'   - column: The column to analyze.
'   - startRow: Optional parameter for the starting row (default is 1).
' Returns: The last row of the table section before hitting an empty cell
Function search_last_row(ws As Worksheet, column As Range, Optional startRow As Integer = 1) As Integer
    Dim currentRow As Integer
    Dim columnNumber As Integer
    
    columnNumber = column.Column
    currentRow = startRow
    
    ' Parcourir la colonne à partir de la ligne de départ
    Do While Not IsEmpty(ws.Cells(currentRow, columnNumber).Value)
        currentRow = currentRow + 1
    Loop
    
    ' Comme on s'est arrêté à la première cellule vide, la dernière ligne avec des données est la précédente
    search_last_row = currentRow - 1
End Function

' Trouve une feuille contenant le texte spécifié dans son nom
' Retourne Nothing si aucune feuille correspondante n'est trouvée
Function FindSheetByNameContaining(searchText As String) As Worksheet
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, searchText, vbTextCompare) > 0 Then
            Set FindSheetByNameContaining = ws
            Exit Function
        End If
    Next ws
    
    ' Aucune feuille correspondante trouvée
    Set FindSheetByNameContaining = Nothing
End Function