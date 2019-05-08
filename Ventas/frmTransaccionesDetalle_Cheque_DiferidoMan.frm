VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Object = "{144A86C7-1AF0-44BA-9AA8-AF3AAF6043B8}#1.0#0"; "NumBox.ocx"
Begin VB.Form frmTransaccionesDetalle_Cheque_DiferidoMan 
   ClientHeight    =   2130
   ClientLeft      =   1125
   ClientTop       =   1980
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2130
   ScaleWidth      =   3150
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.TextBox txtNro 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   1200
      End
      Begin NumBoxProject.NumBox txt_Importe 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         TypeVal         =   2
         Mask            =   "9,999,999,999.99"
         Formato         =   "#,###,###,###.##"
         AllowedMask     =   -1
         MaskLen         =   10
         Aling           =   3
         Text            =   "0.00"
         CanEmpty        =   -1
         ShowError       =   0
         Locked          =   0   'False
         Enabled         =   -1  'True
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DecimalNumber   =   2
      End
      Begin VB.Label Label5 
         Caption         =   "Numero :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   255
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe :"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   765
         Width           =   615
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmTransaccionesDetalle_Cheque_DiferidoMan.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmTransaccionesDetalle_Cheque_DiferidoMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lfAceptar As Boolean, strStore As String, strCod_Banco As String, strCod_Moneda As String
Public codigo As String, Descripcion As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sMessage As Long
Dim StrSql As String

On Error GoTo dprError

Select Case ActionName
  Case "GRABAR"
    If MsgBox("Esta seguro de Aplicar este Cheque ", vbYesNo, "IMPORTANTE") = vbYes Then
      If lfSalvar_Datos Then
        Unload Me
        lfAceptar = True
      End If
    End If
  Case "CANCELAR"
      Unload Me
      lfAceptar = False
End Select

Exit Sub

dprError:

errores Err.Number
End Sub

Private Function lfSalvar_Datos() As Boolean

On Error GoTo hand

SQL = strStore & ",'" & strCod_Banco & "','" & strCod_Moneda & "','" & txtNro & "'," & txt_Importe.Text
      
Call ExecuteCommandSQL(cCONNECT, SQL)

lfSalvar_Datos = True

Exit Function
Resume
hand:

errores Err.Number

lfSalvar_Datos = False

End Function

Private Sub txt_Importe_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNro_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call Busca_Opcion_Cheque
  If KeyAscii >= 48 And KeyAscii <= 57 _
    Or KeyAscii = 8 Or KeyAscii = 13 Then Else KeyAscii = 0
End Sub

Public Sub Busca_Opcion_Cheque()

On Error GoTo Fin

Dim rstAux As ADODB.Recordset, StrSql As String

    StrSql = "Select a.Cod_Banco,Banco = b.Nom_Banco,Moneda = a.Cod_Moneda,Nro_Cheque,Importe_Pendiente = Importe -Importe_Cancelado From Cn_Ventas_Cheques_Diferidos a, Tg_Banco b Where a.Cod_Banco = b.Cod_Banco and a.Flg_Status = 'P' "

    txt_Importe.Text = "0"
    txtNro = "0"
    strCod_Banco = ""
    strCod_Moneda = ""
    
    With frmBusqGeneral
        Set .oParent = Me
        .SQuery = StrSql
        .Cargar_Datos
        
        Me.codigo = ""
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then
          .DGridLista.Columns("Cod_Banco").Visible = False
          .DGridLista.Columns("Banco").Width = 2550
          .DGridLista.Columns("Moneda").Width = 900
          .DGridLista.Columns("Importe_Pendiente").Width = 1515
          .DGridLista.Columns("Importe_Pendiente").Caption = "Imp Pendiente"
          .DGridLista.Columns("Importe_Pendiente").Format = "###,###.00"
          .Show vbModal
        Else
          Me.codigo = ".."
        End If
        
        If Me.codigo <> "" And rstAux.RecordCount > 0 Then
            txtNro = Trim(rstAux!Nro_Cheque)
            txt_Importe.Text = Trim(rstAux!Importe_Pendiente)
            strCod_Banco = Trim(rstAux!Cod_Banco)
            strCod_Moneda = Trim(rstAux!Moneda)
            SendKeys "{TAB}"
        End If
        
    End With
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Resume
Fin:
On Error Resume Next
    Unload frmBusqGeneral
    Set frmBusqGeneral = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub


