Attribute VB_Name = "Module2"
Option Explicit
Dim sServer As String
Dim sModoAutenticacion  As String
Dim sUserName As String
Dim sPassword As String
Dim conn As New ADODB.Connection

Sub Main()

vusu = "SISTEMAS"
vper = "0001"
vemp = "01"
vRuta = App.Path
iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
   
 'sServer = GetSetting("Visuales", "Settings", "Server")
 
 sServer = GetSetting("Visuales", "Settings", "Server")
 sModoAutenticacion = GetSetting("Visuales", "Settings", "AutenticacionMode")
 
     If UCase(sModoAutenticacion) = "SQL" Then
        sUserName = GetSetting("Visuales", "Settings", "UserName")
        sPassword = GetSetting("Visuales", "Settings", "Password")
        cCONNECT = "Provider=SQLOLEDB;User ID=" & RTrim(sUserName) & ";Password=" & RTrim(sPassword) & ";Server=" & sServer & ";Database=HIALPESA;Use Procedure for Prepare=0;Auto Translate=FALSE;Packet Size=4096;Use Encryption for Data=FALSE;Tag with column collation when possible=FALSE"
     Else
        cCONNECT = "Provider=SQLOLEDB;Integrated Security=SSPI;Server=" & sServer & ";Database=HIALPESA;Use Procedure for Prepare=0;Auto Translate=FALSE;Packet Size=4096;Use Encryption for Data=FALSE;Tag with column collation when possible=FALSE"
     End If
     conn.ConnectionString = cCONNECT
    
     conn.Open cCONNECT
     
     frmShowCtaCteCopy.Show 1

End Sub
