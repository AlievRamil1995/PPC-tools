Sub CSVToXML()
    'Coordinates of data table (first row reserved for column names)
    Dim Row As Integer
    Dim Column As Integer
    
    'Coordinates of xml code
    Dim RowXML As Integer
    Dim ColumnXML As Integer
    
    'Starting coordinate for XML code
    RowXML = 1
    ColumnXML = 11
    
    'Replace 470 with last row of your table and 9 with last column of your table
    Dim LastRow As Integer
    LastRow = 470
    Dim LastColumn As Integer
    LastColumn = 9
    
    Cells(RowXML, ColumnXML) = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>"
    RowXML = RowXML + 1
    Cells(RowXML, ColumnXML) = "<rss xmlns:g=" & Chr(34) & "http://base.google.com/ns/1.0" & Chr(34) & " version=" & Chr(34) & "2.0" & Chr(34) & ">"
    RowXML = RowXML + 1
    Cells(RowXML, ColumnXML) = "<channel>"
    RowXML = RowXML + 1
    'Replace text with your title
    Cells(RowXML, ColumnXML) = "<title>TITLE OF YOUR SITE</title>"
    RowXML = RowXML + 1
    'Replace text with your site address
    Cells(RowXML, ColumnXML) = "<link>link-to-your-site.com</link>"
    RowXML = RowXML + 1
    For Row = 2 To LastRow
        Cells(RowXML, ColumnXML) = "<item>"
        RowXML = RowXML + 1
        For Column = 1 To LastColumn
            Cells(RowXML, ColumnXML) = "<" & Cells(1, Column) & ">" & Cells(Row, Column) & "</" & Cells(1, Column) & ">"
            RowXML = RowXML + 1
        Next Column
        Cells(RowXML, ColumnXML) = "</item>"
        RowXML = RowXML + 1
    Next Row
    Cells(RowXML, ColumnXML) = "</channel>"
    RowXML = RowXML + 1
    Cells(RowXML, ColumnXML) = "</rss>"
End Sub
