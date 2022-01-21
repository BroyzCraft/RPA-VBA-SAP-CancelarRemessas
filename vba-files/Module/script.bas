Attribute VB_Name = "script"
Sub cancelarRemessas()

login = 0
If login = 1 Then
sap_login:
    Call SAP_Logon
End If

If Not IsObject(app) Then
    Set SapGuiAuto = GetObject("SAPGUI")
    Set app = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(connection) Then
    On Error GoTo sap_login
    Set connection = app.Children(0)
End If

If Not IsObject(session) Then
    Set session = connection.Children(0)
End If

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject app, "on"
End If

tot = Range("A100000").End(xlUp).Row
x = 0
'script
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVL03N"
session.findById("wnd[0]").sendVKey 0
Do While x < tot
session.findById("wnd[0]/usr/ctxtLIKP-VBELN").Text = Range("A" & x + 1).Value
On Error Resume Next
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[25]").press
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
x = x + 1
Loop

MsgBox ("Remessas Canceladas")

End Sub
