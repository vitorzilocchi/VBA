Attribute VB_Name = "M_EncontraDistancias"
Option Explicit

Function GetXML(url As String) As String

Dim XMLHttpRequest As Object 'xmlhttp

Set XMLHttpRequest = Interaction.CreateObject("MSXML2.xmlhttp") ' New MSXML2.xmlhttp
XMLHttpRequest.Open "GET", url, False
XMLHttpRequest.setRequestHeader "Content-Type", "text/xml"
XMLHttpRequest.send
GetXML = XMLHttpRequest.responseText

End Function

Sub BuscaIEdistUpdate()

Dim inicio As String, MyStr As String, swap1 As String, swap2 As String, strAddress As String
Dim linha As Integer, j As Integer, n1 As Long, tam As Long, n2 As Long
Dim kilometragem As Long, tempos As Long
Dim IE As InternetExplorer 'Declara objeto internet explorer
Dim isOK As Boolean

linha = 5

With Plan2
  
  .Activate

  While .Cells(linha, 1).Value <> ""
    
    'troca virgula por ponto
    swap1 = Strings.Replace(.Cells(linha, 1), " ", "+") 'origem
    swap2 = Strings.Replace(.Cells(linha, 2), " ", "+") 'destino
    
    swap1 = RetirarAcento(swap1)
    swap2 = RetirarAcento(swap2)
    
    '\\ Define a consulta no google maps
    strAddress = "https://maps.google.com/maps/api/directions/xml?origin=" & swap1 & ",Brasil&destination=" & swap2 & "&key=AIzaSyAjYg43jxigvqCzbf8QbdI1DnILi4olF6Y"
    
    'strAddress = "https://maps.googleapis.com/maps/api/directions/json?origin=Toronto&destination=Montreal&key=AIzaSyAjYg43jxigvqCzbf8QbdI1DnILi4olF6Y"
    
    'https://maps.googleapis.com/maps/api/directions/json?origin=Toronto&destination=Montreal&key=YOUR_API_KEY
    '\\ l� o xml
    MyStr = GetXML(strAddress)
    
    'Verificador se o limite de pesquisas foi atingido
    ' Se der pau, tem que cadastrar outro codigo de acesso ao gmaps
    If Strings.InStr(1, MyStr, "OVER_QUERY_LIMIT") > 0 Then
      Set IE = Nothing
      Interaction.MsgBox ("Limite de pesquisa do google atingido")
      Exit Sub
    End If
    
    isOK = False
    
    'acha o ultimo duration
    While isOK = False
      n1 = Strings.InStr(1, MyStr, "<duration>")
      tam = Strings.Len(MyStr)
      MyStr = Strings.Mid(MyStr, n1 + 5, tam - n1)
      If n1 = 0 Then
        isOK = True
      End If
    Wend

    n1 = Strings.InStr(1, MyStr, "value")
    n2 = Strings.InStr(n1 + 5, MyStr, "/value")
    tempos = Strings.Mid(MyStr, n1 + 6, n2 - 1 - (n1 + 6))
                
    'acha ultimo distance
    isOK = False
    While isOK = False
      n1 = Strings.InStr(1, MyStr, "<distance>")
      tam = Strings.Len(MyStr)
      MyStr = Strings.Mid(MyStr, n1 + 5, tam - n1)
      If n1 = 0 Then
        isOK = True
      End If
    Wend
    
    'Encontra a palavra "text"
    n1 = Strings.InStr(1, MyStr, "value")
    n2 = Strings.InStr(n1 + 5, MyStr, "/value")
    kilometragem = Strings.Mid(MyStr, n1 + 6, n2 - 1 - (n1 + 6))
                
    '\\ grava os resultados
    .Cells(linha, 4) = kilometragem / 1000
    .Cells(linha, 5) = tempos / 60
    
    linha = linha + 1
    
  Wend

End With

End Sub

Function RetirarAcento(Palavra As String) As String
    
  RetirarAcento = Strings.Replace(Palavra, "�", "c")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "C")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "a")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "a")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "a")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "a")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "e")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "e")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "e")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "i")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "i")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "i")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "o")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "o")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "o")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "o")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "u")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "u")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "u")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "A")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "A")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "A")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "A")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "E")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "E")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "E")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "I")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "I")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "I")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "O")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "O")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "O")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "O")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "U")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "U")
  RetirarAcento = Strings.Replace(RetirarAcento, "�", "U")
       
End Function

Sub BuscaIEdist(linha As Integer)

Dim inicio As String, MyStr As String, distTemp As String
Dim swap1 As String, swap2 As String, swap3 As String, swap4 As String
Dim n1 As Long, n2 As Long, tam As Long
Dim IE As Object

'Cria o objeto para navega��o Internet Explorer
Set IE = Interaction.CreateObject("InternetExplorer.Application")

'troca espa�o por positivo
swap1 = Replace(Cells(linha, 1), " ", "+")
swap2 = Replace(Cells(linha, 2), " ", "+")

'retira acentos
swap1 = RetirarAcento(swap1)
swap2 = RetirarAcento(swap2)

'cria o objeto para visualiza��o do maps
Set IE = CreateObject("InternetExplorer.Application")
IE.Navigate "http://maps.google.com/maps?f=d&saddr=" & swap1 & "&daddr=" & swap2

'aguarda at� que a pagina seja totalmente carregada
Do Until Not IE.Busy And IE.ReadyState = READYSTATE_COMPLETE
Loop

IE.Visible = True
        
End Sub

Sub Remover_Acentos()

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Controle")
    Dim LastRow As Integer
    Dim i As Integer
    
    LastRow = Sheets("Controle").Range("A" & Sheets("Controle").Rows.Count).End(xlUp).Row
    
    For i = 2 To LastRow
        Cells(i, "M") = RetirarAcento(Cells(i, "M"))
        Cells(i, "N") = RetirarAcento(Cells(i, "N"))
        Cells(i, "Q") = RetirarAcento(Cells(i, "Q"))
        Cells(i, "R") = RetirarAcento(Cells(i, "R"))
    Next i

End Sub

