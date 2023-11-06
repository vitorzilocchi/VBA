Attribute VB_Name = "MainControl"

Sub main()
Attribute main.VB_ProcData.VB_Invoke_Func = "m\n14"
Application.ScreenUpdating = False
Debug.Print "inicio:" & Now()
    'Save Curent StatusBar
    oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True

    Dim Arq As String
    Arq = Dir(ThisWorkbook.Path & "\" & "*.xml")
    Dim ultimaLinha As Integer
    Dim KeyList() As String
    
    
    Do Until Arq = ""
        'Show the file name in StatusBar
        Application.StatusBar = Arq
    
        'Open the file and get the KeyNumber
        KeyNumber = CheckFile.GetInvoiceKey(ThisWorkbook.Path & "\" & Arq)
        
        'If the key dont was loaded yet, load it
        If CStr(Application.Match(KeyNumber, Columns("A:A"), 0)) = "Error 2042" Then
            XML_Control.LoadFile (ThisWorkbook.Path & "\" & Arq)
        End If
        
        'Get next file name
        Arq = Dir
        
        'Avoid "Excel Not Responding" duriing the macros runs
        DoEvents
    Loop
    
    'Return the StatusBar to original
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
Debug.Print "inicio:" & Now()
Application.ScreenUpdating = True
End Sub

Sub Limpar()
    Sheets("baseXML").UsedRange.Offset(1).ClearContents
End Sub



