Attribute VB_Name = "CheckFile"
Dim KeyNumber As String

Function GetInvoiceKey(file As String)
xmlUrl = file
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Dim MainNode As Object

    xmlDoc.async = False
    
    If xmlDoc.Load(xmlUrl) Then
        Set elements = xmlDoc.DocumentElement
        For Each MainNode In elements.ChildNodes
            If MainNode.BaseName = "protNFe" Then
                 Call NodeCheck(MainNode)
            End If
        Next
    End If
    GetInvoiceKey = KeyNumber
End Function

Sub NodeCheck(MainNode As Object)
    Dim Node As Object
    
    For Each Node In MainNode.ChildNodes
        If Node.BaseName = "chNFe" Then
            KeyNumber = Node.Text

        ElseIf Node.HasChildNodes Then
            Call NodeCheck(Node)
        End If
    Next Node
End Sub


Function ListLoadedInvoice() As String()
    Dim KeyList() As String
    Dim line As Integer, index As Integer
    line = 2
    index = 2
    
    With ThisWorkbook.Sheets("BaseXML")
        ReDim Preserve KeyList(1 To 1)
        KeyList(1) = .Cells(line, "A").Value
        
        Do While .Cells(line, "A").Value <> ""
            If CStr(Application.Match(.Cells(line, "A").Value, KeyList, 0)) = "Error 2042" Then
                ReDim Preserve KeyList(1 To index)
                KeyList(index) = .Cells(line, "A").Value
                index = index + 1
            End If
            line = line + 1
        Loop
    End With
    ListLoadedInvoice = KeyList

End Function

