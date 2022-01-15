Attribute VB_Name = "LoginSAP"
Sub SAP_Logon()
    
    Dim usuario, senha
    'usuario = UCase(InputBox("Digite seu login SAP: "))
    'senha = InputBox("Digite sua senha SAP: ")
    'debug
    usuario = "BOMARQUES"
    senha = "321654987Leo!"
    
    Dim SapGui, Applic, connection, session, WSHShell
    
    'Abre o Sap instalado na sua máquina
    Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus
    
    'Inicia a variável com o objeto SAP
    Set WSHShell = CreateObject("WScript.Shell")
    Do Until WSHShell.AppActivate("SAP Logon ")
        Application.Wait Now + TimeValue("0:00:01")
    Loop
    
    Set WSHShell = Nothing
    Set SapGui = GetObject("SAPGUI")
    Set Applic = SapGui.GetScriptingEngine
    Set connection = Applic.OpenConnection("14 - ECC PRD - EP1", True)
    Set session = connection.Children(0)
    
    session.findById("wnd[0]").maximize
    'DADOS PARA FAZER O LOGIN NO SISTEMA
    session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "500"
    session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = senha
    session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "PT"
    session.findById("wnd[0]").sendVKey 0

End Sub

