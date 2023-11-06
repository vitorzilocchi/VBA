Attribute VB_Name = "ExcelData"
Function FindColuns() As String()
    Dim BaseColun As Integer
    Dim Position() As String
    BaseColun = 1
    
    With ThisWorkbook.Sheets("baseXML")
        Do While .Cells(1, BaseColun).Value <> ""
            ReDim Preserve Position(1 To BaseColun)
            Position(BaseColun) = .Cells(1, BaseColun).Value
            BaseColun = BaseColun + 1
        Loop
    End With
    FindColuns = Position
End Function

Function LastRow() As Long
    LastRow = ThisWorkbook.Sheets("BaseXML").Range("A" & Rows.Count).End(xlUp).Row
End Function
