Attribute VB_Name = "Main"

Option Explicit
    Dim IE As InternetExplorer
    Dim user As String
    Dim passw As String
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    
Declare PtrSafe Function apiShowWindow Lib "user32" Alias "ShowWindow" _
            (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
 
Global Const SW_MAXIMIZE = 3
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
  
Sub Internet()
Attribute Internet.VB_ProcData.VB_Invoke_Func = "m\n14"

    
    'Declara variaveis
    Dim AllInvoice As String
    Dim AllCBM As Double
    Dim AllCases As Double
    Dim AllWeight As Double
    Dim LinesToRegister() As Integer
    Dim InformationOfLinesToRegister() As FillData.DadosMaster
    Dim Master As FillData.DadosMaster
    Dim IE As Object: Set IE = CreateObject("InternetExplorer.Application")
    Dim i As Variant: i = 0
    Dim FileName As String
    Dim LastRow As Integer
    Dim DN As String
    Dim Agreement As Variant
    Dim decisao As Integer
    Dim objShell As Object
    Dim objAllWindows As Object
    Set objShell = CreateObject("Shell.Application")
    Set objAllWindows = objShell.Windows
    Dim posicao As Integer
    Dim wsh As Object
    Dim BatchPath As String
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim windowStyle As Integer: windowStyle = 0
    Dim ErroCount As Integer: ErroCount = 0
    Dim j As Integer
    Dim k As Integer
    Dim loopp As Integer
    Dim UploadPage As Object: Set UploadPage = CreateObject("InternetExplorer.Application")
    BatchPath = ThisWorkbook.Path & "\Address.bat"

        
    LastRow = Sheets("Controle").Range("A" & Sheets("Controle").Rows.Count).End(xlUp).Row
  
    
    'Abre pagina login
    IE.Visible = True
    apiShowWindow IE.hwnd, SW_MAXIMIZE
    IE.Navigate "https://isc.huawei.com/web/cds/#/cds/edit_qyecysProcessSubmit"
    
    'Realiza Login
    'IE.document.getElementsByClassName("user").uid.Value = "LWX410612"
    'IE.document.getElementsByClassName("psw").Password.Value = "@DSV_2022"
    'IE.document.getElementsByClassName("btn").submit.Click

    'Verifica se  necessario abrir página e fazer login
    
    decisao = MsgBox("Aguarde a pagina abrir, faça logi e aguarde ela estabilizar antes de prosseguir!")

    
    For Each IE In objAllWindows
        If IE.LocationURL <> "" Then
            If IE.document.Title = "miniapp runtime" Then
                Exit For
            End If
        End If
    Next
        
    'Lopping realiza a rotina

    Dim coluna As Integer: coluna = Application.Match("DN", ThisWorkbook.Sheets("Controle").Rows(1), 0)
    For i = 2 To LastRow
    posicao = i
        ReDim LinesToRegister(1 To 1)
        
        If Cells(i, coluna) <> "" Then
            GoTo proximo
        End If

    Master = FillData.GetMasterinformation(i)

    'Loop para agrupar NF, CBM, Weight, Cases
    
    If Master.DN = "" And Master.MaterialType = "GOOD" Then
        AllInvoice = ""
        AllCBM = 0
        AllWeight = 0
        AllCases = 0
        LinesToRegister = FillData.getLinesToRegister(Master.Agendamento, posicao, LastRow)
        For j = 1 To UBound(LinesToRegister)
            ReDim Preserve InformationOfLinesToRegister(1 To j)
            InformationOfLinesToRegister(j) = FillData.GetMasterinformation(LinesToRegister(j))
            AllInvoice = AllInvoice + CStr(InformationOfLinesToRegister(j).InvoiceNumber)
            AllCBM = AllCBM + InformationOfLinesToRegister(j).CBM
            AllWeight = AllWeight + InformationOfLinesToRegister(j).Kg
            AllCases = AllCases + InformationOfLinesToRegister(j).TotalCases
             
            If j <> UBound(LinesToRegister) Then
                AllInvoice = AllInvoice + ","
            End If
            
            If InformationOfLinesToRegister(j).InvoiceKeys = "" Then
            
                For k = 1 To UBound(LinesToRegister)
                
                    ThisWorkbook.Sheets("Controle").Cells(i, Master.DNcol).Value = "PDF Nao encontrado"
                
                Next k
                
                GoTo proximo
                
            End If
        Next j
        

        IE.Visible = False
        IE.Visible = True


On Error GoTo ERRO
InicioPreenchimento:

        'Campos Fixos
        Sleep (2000)
        IE.document.getElementsByClassName("input-padding").Item(0).Focus
        Sleep (1000)
        SendKeys "LWX410612"
        Sleep (1000)
        For loopp = 1 To 20
            SendKeys "{DEL}"
        Next loopp
        Sleep (4000)
        SendKeys "{TAB}"
      
        Sleep (2000)
        SendKeys "TILGATE RODRIGUES"
        Sleep (1000)
        SendKeys "{TAB}"
        Sleep (1000)
        SendKeys "551533161391"
        Sleep (1000)
        IE.document.getElementsByClassName("hae-input")(3).Children(0).Value = "SPARE PARTS"
        Sleep (500)
        IE.document.getElementsByClassName("hae-input")(3).Children(0).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (2000)
        SendKeys "{ENTER}"
        
        
        IE.document.getElementsByClassName("hae-icon icon-popup")(0).Click 'Country
        Sleep (1000)
        IE.document.getElementsByClassName("hae-ui-input")(16).Focus 'Country
        Sleep (1000)
        SendKeys "BRAZIL" 'Pesquisa país
        Sleep (1000)
        SendKeys "{ENTER}"
        Sleep (1000)
        IE.document.getElementsByClassName("hae-btn hae-btn btn-small")(0).Click
        SendKeys "{ENTER}"
        Sleep (3000)
        IE.document.getElementsByClassName("grid-input")(0).Click 'Seleciona Brazil
        Sleep (1000)
        IE.document.getElementsByClassName("hae-btn btn-primary")(0).Click 'OK
        Sleep (1000)
        IE.document.getElementsByClassName("hae-input")(9).Children(0).Value = "BY TRUCK" 'Transport Mode
        Sleep (500)
        IE.document.getElementsByClassName("hae-input")(9).Children(0).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (2500)
        SendKeys "{ENTER}"
        Sleep (1000)
        
        If Master.MaterialType = "GOOD" Then
            IE.document.getElementsByClassName("hae-input")(8).Children(0).Focus 'Consigned Delivery Reason
            SendKeys "FSLDELIVERY"
            Sleep (500)
            SendKeys "{TAB}"
            Sleep (500)
            SendKeys "{TAB}"
            Sleep (500)
            SendKeys "{TAB}"
            Sleep (500)
            SendKeys "FSLDELIVERY"
            SendKeys "{ENTER}"
            Sleep (500)
        ElseIf Master.MaterialType = "FAULTY" Then
            IE.document.getElementsByClassName("hae-input")(8).Children(0).Focus 'Consigned Delivery Reason
            SendKeys "FSLREVERSE"
            Sleep (500)
            SendKeys "{TAB}"
            Sleep (500)
            SendKeys "{TAB}"
            Sleep (500)
            SendKeys "{TAB}"
            Sleep (500)
            SendKeys "FSLREVERSE"
            SendKeys "{ENTER}"
            Sleep (500)
        End If
                   
        'Campos Variaveis
        IE.document.getElementsByClassName("hae-ui-input")(6).Value = Master.PONumber
        IE.document.getElementsByClassName("hae-ui-input")(6).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (500)
        IE.document.getElementsByClassName("hae-ui-input")(7).Value = Master.CustomerName
        IE.document.getElementsByClassName("hae-ui-input")(7).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (500)
        IE.document.getElementsByClassName("hae-ui-input")(8).Value = Master.Agendamento
        IE.document.getElementsByClassName("hae-ui-input")(8).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (500)
        IE.document.getElementsByClassName("hae-ui-input")(9).Value = Master.CTE
        IE.document.getElementsByClassName("hae-ui-input")(9).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (500)
        IE.document.getElementsByClassName("hae-ui-input")(10).Value = AllInvoice
        IE.document.getElementsByClassName("hae-ui-input")(10).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (500)
        IE.document.getElementsByClassName("hae-input")(18).Children(0).Value = Format(Master.ShippingDate, "yyyy-mm-dd hh:mm:ss")
        IE.document.getElementsByClassName("hae-input")(18).Children(0).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (500)
        IE.document.getElementsByClassName("hae-input")(19).Children(0).Value = Format(Master.ArrivingDate, "yyyy-mm-dd hh:mm:ss")
        IE.document.getElementsByClassName("hae-input")(19).Children(0).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (500)
        IE.document.getElementsByClassName("hae-ui-input")(12).Value = Master.ReceiverName
        IE.document.getElementsByClassName("hae-ui-input")(12).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (500)
        IE.document.getElementsByClassName("hae-ui-input")(13).Value = Master.ReceiverTelephone
        IE.document.getElementsByClassName("hae-ui-input")(13).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (500)
        IE.document.getElementsByClassName("hae-input")(22).Children(0).Value = Master.ReceiverAddress
        IE.document.getElementsByClassName("hae-input")(22).Children(0).Focus
        SendKeys "^{END}"
        SendKeys "{ENTER}"
        Sleep (500)
        
        IE.document.getElementsByClassName("hae-input")(25).Children(0).Value = "BRAZIL"
        Sleep (500)
        IE.document.getElementsByClassName("hae-input")(25).Children(0).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (2000)
        SendKeys "{ENTER}"
        
        
        IE.document.getElementsByClassName("hae-input")(26).Children(0).Value = Master.LocationFID
        Sleep (500)
        IE.document.getElementsByClassName("hae-input")(26).Children(0).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (2000)
        SendKeys "{ENTER}"
        
        IE.document.getElementsByClassName("hae-input")(27).Children(0).Value = "BRAZIL"
        Sleep (500)
        IE.document.getElementsByClassName("hae-input")(27).Children(0).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (2000)
        SendKeys "{ENTER}"
        
        IE.document.getElementsByClassName("hae-input")(28).Children(0).Value = Master.LocationTID
        Sleep (500)
        IE.document.getElementsByClassName("hae-input")(28).Children(0).Focus
        Sleep (500)
        SendKeys "{ENTER}"
        Sleep (2000)
        SendKeys "{ENTER}"
        
        
        'Department Approver
        
        IE.document.getElementsByClassName("hae-icon icon-popup")(2).Click
        Sleep (1000)
        IE.document.getElementsByClassName("input-padding")(2).Focus
        Sleep (700)
        SendKeys "00746232"
        Sleep (1000)
        SendKeys "{ENTER}"
        Sleep (2000)
        IE.document.getElementsByClassName("hae-btn hae-btn btn-small")(0).Click
        Sleep (2000)
        IE.document.getElementsByClassName("grid-input")(0).Click
        Sleep (700)
        IE.document.getElementsByClassName("hae-btn btn-primary")(0).Click 'OK
        Sleep (1000)
        
        'Logistic Manager
        IE.document.getElementsByClassName("hae-icon icon-popup")(3).Click
        Sleep (1000)
        IE.document.getElementsByClassName("input-padding")(2).Focus
        Sleep (700)
        SendKeys "00715521"
        Sleep (100)
        SendKeys "{ENTER}"
        Sleep (2000)
        IE.document.getElementsByClassName("hae-btn hae-btn btn-small")(0).Click
        Sleep (2000)
        IE.document.getElementsByClassName("grid-input")(0).Click
        Sleep (700)
        IE.document.getElementsByClassName("hae-btn btn-primary")(0).Click 'OK
        Sleep (1000)
                
        'Copy users
        IE.document.getElementsByClassName("input-padding")(1).Focus
        'Sleep (700)
        'SendKeys "AWX743081"
        Sleep (2000)
                  
        IE.document.getElementsByClassName("hae-input")(38).Children(1).Focus
        Sleep (1000)
        IE.document.getElementsByClassName("hae-input")(38).Children(1).Focus
        SendKeys AllCBM
        Sleep (1000)
        SendKeys "{TAB}"
        Sleep (1000)
        SendKeys AllWeight
        Sleep (1000)
        SendKeys "{TAB}"
        Sleep (1000)
        SendKeys AllCases
        Sleep (1000)
        SendKeys "{TAB}"
        
        IE.document.getElementsByClassName("hae-btn")(1).Click
        Sleep (2000)
        IE.document.getElementsByClassName("grid-cell consignLineGridId-col1 cell-tip")(0).Click
        Sleep (1000)
        SendKeys "NETWORK"
        Sleep (1000)
        SendKeys "{ENTER}"
        Sleep (2000)
        
       
        IE.document.getElementsByClassName("hae-input")(41).Children(0).Value = Master.DetailedAdress
        IE.document.getElementsByClassName("hae-input")(41).Children(0).Focus
        SendKeys "^{END}"
        SendKeys "{ENTER}"
        Sleep (500)
        IE.document.getElementsByClassName("hae-input")(42).Children(0).Value = RemarkFunc(Master.DType, Master.Vehicle, Master.FreightValue, Master.KM)
        IE.document.getElementsByClassName("hae-input")(42).Children(0).Focus
        SendKeys "^{END}"
        SendKeys "{ENTER}"
        SendKeys "{TAB}"
        Sleep (500)
        
        IE.document.getElementsByClassName("grid-cell consignLineGridId-col2 cell-tip  table-number")(0).Click
        Sleep (1000)
        SendKeys "1"
        Sleep (1000)
        
        For j = 1 To UBound(LinesToRegister)
                
            Open BatchPath For Output As #1
                Print #1, "TIMEOUT 3"
                Print #1, "powershell -c " & Chr(34) & "$wshell = New-Object -ComObject wscript.shell; $wshell.SendKeys('" & ThisWorkbook.Path & "\" & Master.city & "-" & Master.UF & "\REMESSA PARA A OPERADORA GOOD\" & InformationOfLinesToRegister(j).InvoiceKeys & "')"
                Print #1, "powershell -c " & Chr(34) & "$wshell = New-Object -ComObject wscript.shell; $wshell.SendKeys('{ENTER}')"
                Close #1
                
            'Roda batch
            BatchPath = ThisWorkbook.Path & "\Address.bat"
            wsh.Run (Chr(34) & BatchPath & Chr(34)), windowStyle
            
            'abre seletor de arquivo enquanto a batch est rodando
            
            IE.document.getElementsByClassName("aui-button aui-button--default")(0).Click
        
        Sleep (7000)
        Next j
           

        IE.document.getElementsByClassName("hae-icon icon-check")(0).Click
        Sleep (500)
        
        If Master.Centro_de_Custo = "N" Then 'diferente de spareparts
        MsgBox ("Preencher centro de custo (Beneficiary Dept) manualmente")
        Exit Sub
        End If
        
        IE.document.getElementsByClassName("hae-radio")(1).Children(0).Click
               
        'Submit
        IE.document.getElementsByClassName("hae-btn")(8).Click
        Sleep (1000)
        IE.document.getElementsByClassName("hae-btn btn-primary")(0).Click
        Sleep (1000)
        IE.document.getElementsByClassName("hae-btn btn-primary")(0).Click
        Sleep (4000)

        
        Sleep (5000)
        IE.document.getElementsByClassName("hae-tabs__item is-top")(0).Click
        
        
        Sleep (4000)

        DN = IE.document.getElementsByClassName("grid-cell")(44).innerText
            
            
        If Left(DN, 4) <> "DBRA" Then
            Application.Speech.Speak ("Problem, Please Check")
            Exit Sub
        End If
        
            For j = 1 To UBound(LinesToRegister)
                
                ThisWorkbook.Sheets("Controle").Cells(LinesToRegister(j), Master.DNcol).Value = DN
                ThisWorkbook.Sheets("Controle").Cells(LinesToRegister(j), Master.DNdatecol).Value = Now
        
            Next j
        
        IE.Navigate "https://isc.huawei.com/web/cds/#/cds/edit_qyecysProcessSubmit"
    
        ThisWorkbook.Save
            
        Sleep (10000)


    End If
 '####################################################################################################################
    
proximo:
    Next i
    
    Application.Speech.Speak ("All DN were created successfully")
    
    Exit Sub
    
ERRO:
    ErroCount = ErroCount + 1
    If ErroCount < 3 Then
    If IE.LocationURL = "https://isc.huawei.com/web/cds/#/cds/edit_qyecysProcessSubmit" Then
            IE.Navigate "https://isc.huawei.com/"
            Sleep (1000)
            IE.Navigate "https://isc.huawei.com/web/cds/#/cds/edit_qyecysProcessSubmit"
            Sleep (15000)
            
            GoTo InicioPreenchimento
        Else
            MsgBox ("Falha apos submit, verifique se a ultima DN foi salva no excel")
        End If
    Else
        MsgBox ("A macro tentou rodar essa linha 3 vezes, verifique se o dados estao corretos e tente novamente")
    End If
    
End Sub
 
Private Sub Wait_Page(IE As Object)
    'Do While IE.Busy Or IE.ReadyState <> READYSTATE_COMPLETE
    '    Application.StatusBar = "Carregando Pagina"
    'Loop
    Sleep (5000)
End Sub


Private Sub digitar(id As String, elemento As String, IE As Object)
Dim i As Integer
    IE.document.parentWindow.execScript (elemento)
    Application.Wait (Now + TimeValue("00:00:01"))
    For i = 1 To Len(id)
        Sleep (150)
        SendKeys (Mid(id, i, 1)), True
    Next i
    Application.Wait (Now + TimeValue("00:00:01"))
End Sub

Function RemarkFunc(DType As String, Vehicle As String, FreightValue As String, KM As String) As String
    
    DType = Replace(DType, "Dedicado", "FTL")
    DType = Replace(DType, "Fracionado", "LTL")
    
    If Vehicle = "AA" And DType <> "FSL - Aereo Prox Voo" And DType <> "FSL - Aereo Others" Then
        RemarkFunc = "AEREO - NBD"
    ElseIf Vehicle = "AA" And DType = "FSL - Aereo Prox Voo" Then
        RemarkFunc = "AEREO - Next Flight"
    ElseIf Vehicle = "AA" And DType = "FSL - Aereo Others" Then
        RemarkFunc = "AEREO - OTHERS"
    Else
        RemarkFunc = DType & " - " & Vehicle & " - " & KM
    End If
    
End Function

Sub DataHora1(MasterData As Date, IE As Object)

Dim ano As Integer: ano = Year(MasterData)
Dim mes As Integer: mes = Month(MasterData)
Dim dia As Integer: dia = Day(MasterData)
Dim hora As Integer: hora = Hour(MasterData)
Dim minuto As Integer: minuto = Minute(MasterData)
Dim Press As Integer
Dim QtyDias As Integer
Dim i As Integer
Dim MesSelecionado As String
Dim GeneralVar As String
Dim MAtual As Boolean: MAtual = False

If minuto <> 0 Then
    minuto = Int(minuto / 5)
End If

Press = ((Year(Now) - 1) * 12 + Month(Now)) - ((ano - 1) * 12 + mes)

Sleep (1500)
IE.document.parentWindow.execScript "document.getElementsByClassName('webix_input_icon fa-calendar')[0].click()"

Sleep (1000)
IE.document.parentWindow.execScript "document.getElementsByClassName('webix_icon fa-clock-o')[0].click()"

Sleep (1000)
GeneralVar = "document.getElementsByClassName('webix_hours')[0].children[" & hora & "].click()"
IE.document.parentWindow.execScript (GeneralVar)

Sleep (1000)
GeneralVar = "document.getElementsByClassName('webix_minutes')[0].children[" & minuto & "].click()"
IE.document.parentWindow.execScript (GeneralVar)

Sleep (1000)
IE.document.parentWindow.execScript "document.getElementsByClassName('webix_cal_done')[0].click()"

Sleep (1500)
IE.document.parentWindow.execScript "document.getElementsByClassName('webix_input_icon fa-calendar')[0].click()"


While Press <> 0
    IE.document.parentWindow.execScript "document.getElementsByClassName('webix_cal_prev_button')[0].click()"
    Press = Press - 1
    Sleep (300)
Wend

For i = 0 To 41
    
    If IE.document.getElementsByClassName("webix_cal_day_inner")(i).innerText = 1 And MAtual = False Then
        MAtual = True
    End If
    
    If IE.document.getElementsByClassName("webix_cal_day_inner")(i).innerText = dia And MAtual = True Then
        IE.document.parentWindow.execScript (Replace("document.getElementsByClassName('webix_cal_day_inner')[DiaSelecionado].click()", "DiaSelecionado", i))
        Exit For
    End If

Next i

End Sub

Sub DataHora2(MasterData As Date, IE As Object)

Dim ano As Integer: ano = Year(MasterData)
Dim mes As Integer: mes = Month(MasterData)
Dim dia As Integer: dia = Day(MasterData)
Dim hora As Integer: hora = Hour(MasterData)
Dim minuto As Integer: minuto = Minute(MasterData)
Dim Press As Integer
Dim QtyDias As Integer
Dim i As Integer
Dim MesSelecionado As String
Dim GeneralVar As String
Dim MAtual As Boolean: MAtual = False

If minuto <> 0 Then
    minuto = Int(minuto / 5)
End If

Press = ((Year(Now) - 1) * 12 + Month(Now)) - ((ano - 1) * 12 + mes)

Sleep (1500)
IE.document.parentWindow.execScript "document.getElementsByClassName('webix_input_icon fa-calendar')[1].click()"

Sleep (1000)
IE.document.parentWindow.execScript "document.getElementsByClassName('webix_icon fa-clock-o')[1].click()"
IE.document.parentWindow.execScript "document.getElementsByClassName('webix_icon fa-clock-o')[1].click()"

Sleep (1000)
GeneralVar = "document.getElementsByClassName('webix_hours')[0].children[" & hora & "].click()"
IE.document.parentWindow.execScript (GeneralVar)

Sleep (1000)
GeneralVar = "document.getElementsByClassName('webix_minutes')[0].children[" & minuto & "].click()"
IE.document.parentWindow.execScript (GeneralVar)

Sleep (1000)
IE.document.parentWindow.execScript "document.getElementsByClassName('webix_cal_done')[0].click()"

Sleep (1500)
IE.document.parentWindow.execScript "document.getElementsByClassName('webix_input_icon fa-calendar')[1].click()"

Sleep (2000)

IE.document.parentWindow.execScript "document.getElementsByClassName('webix_cal_prev_button')[1].click()"

While Press <> 0
    IE.document.parentWindow.execScript "document.getElementsByClassName('webix_cal_prev_button')[1].click()"
    Press = Press - 1
    Sleep (1000)
Wend

For i = 42 To 83

    If IE.document.getElementsByClassName("webix_cal_day_inner")(i).innerText = 1 And MAtual = False Then
        MAtual = True
    End If
    
    If IE.document.getElementsByClassName("webix_cal_day_inner")(i).innerText = dia And MAtual = True Then
        IE.document.parentWindow.execScript (Replace("document.getElementsByClassName('webix_cal_day_inner')[DiaSelecionado].click()", "DiaSelecionado", i))
        Exit For
    End If

Next i

End Sub



