Attribute VB_Name = "XML_CommonInformation"
Public Type CommonInformation
    ParentNode As String
    ItemName As String
    ItemValue As String

End Type

Dim CommonInformation() As CommonInformation
Dim index As Integer
Function Upload(xmlUrl As String) As CommonInformation()

    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Dim MainNode As Object
    Dim Position() As String
    index = 0
    xmlDoc.async = False
    
    Position = ExcelData.FindColuns
    
    If xmlDoc.Load(xmlUrl) Then
        Set elements = xmlDoc.DocumentElement
        For Each MainNode In elements.ChildNodes
            Call NodeCheck(MainNode, Position)
        Next

    End If
    Upload = CommonInformation

End Function

Sub NodeCheck(MainNode As Object, Positon() As String)
    Dim Node As Object
    For Each Node In MainNode.ChildNodes
        If Node.BaseName <> "det" Then
            If Node.HasChildNodes Then
                Call NodeCheck(Node, Positon)
            End If
            If Node.Text <> "" And Node.BaseName <> "" Then
                Call LoadInformation(Node.ParentNode.BaseName, Node.BaseName, Node.Text)
            End If
        End If
    Next Node
End Sub

Function LoadInformation(ParentNode As String, Name As String, Text As String)
   index = index + 1
    ReDim Preserve CommonInformation(1 To index)
    CommonInformation(index).ParentNode = ParentNode
    CommonInformation(index).ItemName = Name
    CommonInformation(index).ItemValue = Text

End Function

