Attribute VB_Name = "XML_DetItemInformation"
Public Type detCommonInformation
    ParentNode As String
    ItemName As String
    ItemValue As String

End Type

Dim detCommonInformation() As detCommonInformation
Dim total() As detCommonInformation
Dim index As Integer, MainIndex As Integer, MaxIndex As Integer, MaxIndex2 As Integer


Function Upload(xmlUrl As String) As detCommonInformation()

    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Dim MainNode As Object
    index = 0
    MainIndex = 0
    MaxIndex = 0
    xmlDoc.async = False
    MaxIndex2 = 1
    
    If xmlDoc.Load(xmlUrl) Then
        Set elements = xmlDoc.DocumentElement
        
        For Each MainNode In elements.ChildNodes
            Call DetItemCount(MainNode)
        Next
        ReDim total(1 To MaxIndex, 1 To 1)
        For Each MainNode In elements.ChildNodes
            Call NodeCheck(MainNode)
        Next
    End If
    Upload = total


End Function

Sub NodeCheck(MainNode As Object)
    Dim Node As Object
    
    For Each Node In MainNode.ChildNodes
        If Node.BaseName = "det" Then
            index = 0
            MainIndex = MainIndex + 1
            Call DetChildNodes(Node)
        End If
        If Node.HasChildNodes Then
            Call NodeCheck(Node)
        End If
    Next Node
   
End Sub

Sub DetChildNodes(MainNode As Object)
    Dim Node As Object
    For Each Node In MainNode.ChildNodes
        If Node.HasChildNodes Then
            Call DetChildNodes(Node)
        End If
        If Node.Text <> "" And Node.BaseName <> "" Then
            Call LoadInformation(Node.ParentNode.BaseName, Node.BaseName, Node.Text)
        End If
    Next Node
End Sub

Function LoadInformation(ParentNode As String, Name As String, Text As String)
    index = index + 1
    If index > MaxIndex2 Then
        MaxIndex2 = index
    End If
    
    ReDim Preserve total(1 To MaxIndex, 1 To MaxIndex2)
    total(MainIndex, index).ParentNode = ParentNode
    total(MainIndex, index).ItemName = Name
    total(MainIndex, index).ItemValue = Text
End Function

Sub DetItemCount(MainNode As Object)
    Dim Node As Object
    For Each Node In MainNode.ChildNodes
        If Node.BaseName = "det" Then
            
            MaxIndex = MaxIndex + 1
            
        End If
        
        If Node.HasChildNodes Then
            Call DetItemCount(Node)
        End If
    Next Node

End Sub

