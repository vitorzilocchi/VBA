Attribute VB_Name = "XML_Control"
Sub LoadFile(file As String)
    Dim CommonInformation() As XML_CommonInformation.CommonInformation
    Dim ItemInformation() As XML_DetItemInformation.detCommonInformation
    Dim LastRow As Integer
    Dim ColunsArray() As String
    Dim index As Integer
    
    ColunsArray = ExcelData.FindColuns
    LastRow = ExcelData.LastRow
    CommonInformation = XML_CommonInformation.Upload(file)
    ItemInformation = XML_DetItemInformation.Upload(file)

    For index = 1 To UBound(ItemInformation, 2)
        Colun = Application.Match(ItemInformation(1, index).ParentNode & "." & ItemInformation(1, index).ItemName, ColunsArray, 0)
        If CStr(Colun) <> "Error 2042" Then
            For i = 1 To UBound(ItemInformation, 1)
                ThisWorkbook.Sheets("BaseXML").Cells(i + LastRow, Colun).Value = ItemInformation(i, index).ItemValue
                
            Next i
            
        End If
        
    Next index
    
    For index = 1 To UBound(CommonInformation, 1)
    On Error GoTo Ends
        Colun = Application.Match(CommonInformation(index).ParentNode & "." & CommonInformation(index).ItemName, ColunsArray, 0)
        
        If CStr(Colun) <> "Error 2042" Then
            For i = 1 To UBound(ItemInformation, 1)
                ThisWorkbook.Sheets("BaseXML").Cells(i + LastRow, Colun).Value = CommonInformation(index).ItemValue
                
            Next i
            
        End If
Ends:
    Next index

End Sub


