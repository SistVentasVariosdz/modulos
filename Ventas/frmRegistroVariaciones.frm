VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmRegistroVariaciones 
   Caption         =   "Registro de Variaciones"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1500
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   4380
      Begin VB.TextBox txtCobranzas 
         Height          =   315
         Left            =   1965
         TabIndex        =   4
         Top             =   795
         Width           =   1635
      End
      Begin VB.TextBox txtVentas 
         Height          =   315
         Left            =   1965
         TabIndex        =   3
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label Label2 
         Caption         =   "Cobranzas (%)"
         Height          =   300
         Left            =   750
         TabIndex        =   2
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Ventas  (%)"
         Height          =   270
         Left            =   765
         TabIndex        =   1
         Top             =   435
         Width           =   1125
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1065
      TabIndex        =   5
      Top             =   1725
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmRegistroVariaciones.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmRegistroVariaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sAnio As String
Public smes As String
Dim VENTAS As String
Dim COBRANZAS As String
Dim Tipo As String
Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
    If Trim(txtVentas.Text) = "" Then
        VENTAS = 0
    Else
         VENTAS = Trim(txtVentas.Text)
    End If
    If Trim(txtCobranzas.Text) = "" Then
        COBRANZAS = 0
    Else
        COBRANZAS = Trim(txtCobranzas.Text)
    End If
    Call Salvar_Datos
Case "CANCELAR"
    Unload Me
End Select
End Sub


 

Sub Salvar_Datos()
    Dim Con As New ADODB.Connection
    Dim RS As Object
    Set RS = CreateObject("ADODB.Recordset")
    On Error GoTo Salvar_DatosErr
    Dim strSQL As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
        Con.BeginTrans
        
        
       If Tipo = "U" Then
       
         
         If MsgBox("Esta seguro de Actualizar los datos", vbInformation + vbYesNo, "AVISO") = vbYes Then
            
            strSQL = "EXEC CN_INGRESO_VARIACION_VENTAS_COBRANZAS '" & Tipo & "','" & _
            sAnio & "','" & _
            smes & "'," & _
            VENTAS & "," & _
            COBRANZAS & ""
         
         End If
               
       Else
       
             strSQL = "EXEC CN_INGRESO_VARIACION_VENTAS_COBRANZAS '" & Tipo & "','" & _
             sAnio & "','" & _
             smes & "'," & _
             VENTAS & "," & _
             COBRANZAS & ""
            
       End If
        

        
        Con.Execute strSQL
       
        Con.CommitTrans
        
        Aviso "Proceso Culminó Satisfactoriamente", 2
        Unload Me
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler err, "Salvar_Datos"
End Sub

 
 Public Sub CargarDatos()
 
 Dim verifica As Integer
 
 verifica = DevuelveCampo("select count(*) from CN_VENTAS_VARIACION_VEN_COBRA where  ano='" & sAnio & "' and mes ='" & smes & "'", cCONNECT)
 If verifica = 0 Then
 Tipo = "I"
 Else
 Tipo = "U"
 
 txtVentas.Text = Trim(DevuelveCampo("select imp_por_ventas from cn_ventas_variacion_ven_cobra where ANO='" & sAnio & "' and MES='" & smes & "' ", cCONNECT))
 txtCobranzas.Text = Trim(DevuelveCampo("select imp_por_cobranza from cn_ventas_variacion_ven_cobra where ANO='" & sAnio & "' and MES='" & smes & "' ", cCONNECT))
 
 End If
 End Sub
 
 
 Public Sub txtCobranzas_GotFocus()
 SelectionText txtCobranzas
 End Sub

  Public Sub txtVentas_GotFocus()
 SelectionText txtVentas
 End Sub



 Private Sub txtCobranzas_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     FunctButt1.SetFocus
  End If
  
End Sub

 Private Sub txtVentas_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtCobranzas.SetFocus
  End If
  
End Sub
 
