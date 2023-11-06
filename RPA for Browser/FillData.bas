Attribute VB_Name = "FillData"
Public Type DadosMaster
    
    Agendamento As String
    PONumber As String
    CustomerName As String
    PlNumber As String
    Vehicle As String
    FreightValue As String
    MaterialType As String
    InvoiceNumber As Double
    ShippingDate As Date
    ArrivingDate As Date
    ReceiverName As String
    ReceiverTelephone As String
    ReceiverAddress As String
    LocationFID As String
    LocationTID As String
    CBM As Double
    Kg As Double
    TotalCases As Double
    DetailedAdress As String
    Remark As String
    InvoiceKeys As String
    DN As String
    DNcol As Integer
    DNdatecol As Integer
    QtyAgend As Integer
    DType As String
    Centro_de_Custo As String
    KM As String
    city As String
    UF As String
    CTE As String
    
End Type

Function GetMasterinformation(linha As Variant) As DadosMaster

On Error Resume Next
    
    Application.ScreenUpdating = False
    Dim coluna As Integer
    Dim linhatitulo As Integer
    coluna = 1
    linhatitulo = 1
       
   With ThisWorkbook.Sheets("Controle")
   
        Do While ThisWorkbook.Sheets("Controle").Cells(linhatitulo, coluna).Value <> ""
   
   
            Select Case ThisWorkbook.Sheets("Controle").Cells(linhatitulo, coluna).Value
                
                Case Is = "Agendamento"
                    GetMasterinformation.Agendamento = .Cells(linha, coluna).Value
   
                Case Is = "Numero RMA"
                    GetMasterinformation.PONumber = .Cells(linha, coluna).Value
   
                Case Is = "PL"
                    GetMasterinformation.PlNumber = .Cells(linha, coluna).Value
   
                Case Is = "Projeto"
                    GetMasterinformation.CustomerName = .Cells(linha, coluna).Value
                
                Case Is = "NF"
                    GetMasterinformation.InvoiceNumber = .Cells(linha, coluna).Value

                Case Is = "Tipo de Material"
                    GetMasterinformation.MaterialType = .Cells(linha, coluna).Value
                    
                Case Is = "M3"
                    GetMasterinformation.CBM = .Cells(linha, coluna).Value
                    
                Case Is = "Peso"
                    GetMasterinformation.Kg = .Cells(linha, coluna).Value
                    
                Case Is = "Qtde de volumes"
                    GetMasterinformation.TotalCases = .Cells(linha, coluna).Value
                    
                Case Is = "Endereco Entrega"
                    GetMasterinformation.ReceiverAddress = .Cells(linha, coluna).Value
                    
                Case Is = "Recebedor"
                    GetMasterinformation.ReceiverName = .Cells(linha, coluna).Value
                    
                Case Is = "Telefone"
                    GetMasterinformation.ReceiverTelephone = .Cells(linha, coluna).Value
                    
                Case Is = "Endereco Coleta"
                    GetMasterinformation.DetailedAdress = .Cells(linha, coluna).Value
                    
                Case Is = "Data e hora de solicitacao do agendamento"
                    GetMasterinformation.ShippingDate = .Cells(linha, coluna).Value
                    
                Case Is = "Date e hora de agendamento"
                    GetMasterinformation.ArrivingDate = .Cells(linha, coluna).Value
                                        
                Case Is = "DN"
                    GetMasterinformation.DN = .Cells(linha, coluna).Value
                    GetMasterinformation.DNcol = coluna
                    
                Case Is = "DN Date"
                    GetMasterinformation.DNdatecol = coluna
                    
                Case Is = "Agendamento"
                    GetMasterinformation.QtyAgend = Qty_Agend(.Cells(linha, coluna).Value)

                Case Is = "Tipo de Veiculo"
                    GetMasterinformation.Vehicle = .Cells(linha, coluna).Value
                                        
                Case Is = "Valor do Frete"
                    GetMasterinformation.FreightValue = .Cells(linha, coluna).Value
                    
                Case Is = "Tipo de Entrega"
                    GetMasterinformation.DType = .Cells(linha, coluna).Value
                                        
                Case Is = "LocationFID"
                    GetMasterinformation.LocationFID = .Cells(linha, coluna).Value
                    
                Case Is = "LocationTID"
                    GetMasterinformation.LocationTID = .Cells(linha, coluna).Value
                       
                Case Is = "Centro de Custo Spare Parts"
                    GetMasterinformation.Centro_de_Custo = .Cells(linha, coluna).Value
                    
                Case Is = "KM"
                    GetMasterinformation.KM = .Cells(linha, coluna).Value
                    
                Case Is = "Cidade Coleta"
                    GetMasterinformation.city = .Cells(linha, coluna).Value

                Case Is = "UF Coleta"
                    GetMasterinformation.UF = .Cells(linha, coluna).Value
                    
                Case Is = "CTE"
                    GetMasterinformation.CTE = .Cells(linha, coluna).Value

            End Select
   
        coluna = coluna + 1
   
        Loop
   
    End With

    GetMasterinformation.InvoiceKeys = FileName(GetMasterinformation.InvoiceNumber, GetMasterinformation.city, GetMasterinformation.UF)

On Error GoTo 0

End Function

Function Qty_Agend(cod As String) As Integer

    Qty_Agend = Application.WorksheetFunction.CountIf(ThisWorkbook.Sheets("Controle").Range("A:A"), cod)
    
End Function

Function FileName(nf As Double, city As String, UF As String) As String

    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    Dim ssss As String
    
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(ThisWorkbook.Path & "\" & city & "-" & UF & "\REMESSA PARA A OPERADORA GOOD")
    
    For Each oFile In oFolder.Files
        If Len(oFile.Name) = 48 Then
            If Int(Mid(oFile.Name, 26, 9)) = nf Then
                FileName = oFile.Name
                Exit For
            End If
        ElseIf Len(oFile.Name) = 51 Then
            If Int(Mid(oFile.Name, 29, 9)) = nf Then
                FileName = oFile.Name
                Exit For
            End If
        ElseIf Len(oFile.Name) = 71 Then
            If Int(Mid(oFile.Name, 49, 9)) = nf Then
                FileName = oFile.Name
                Exit For
            End If
        ElseIf Len(oFile.Name) = 70 Then
            If Int(Mid(oFile.Name, 48, 9)) = nf Then
                FileName = oFile.Name
                Exit For
            End If
        End If
    Next oFile
    
End Function

Function getLinesToRegister(NumeroAgendamento As String, FistRow As Integer, LastRow As Integer) As Integer()
    Dim Lines() As Integer
    Dim ColunaAgendamento As Integer: ColunaAgendamento = Application.Match("Agendamento", ThisWorkbook.Sheets("Controle").Rows(1), 0)
    Dim Index As Integer: Index = 0
    
    For i = FistRow To LastRow
        If ThisWorkbook.Sheets("Controle").Cells(i, ColunaAgendamento).Value = NumeroAgendamento Then
            Index = Index + 1
            ReDim Preserve Lines(1 To Index)
            Lines(Index) = i
        End If
    Next i
    getLinesToRegister = Lines
End Function
