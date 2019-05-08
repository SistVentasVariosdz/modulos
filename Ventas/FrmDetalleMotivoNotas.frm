VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmDetalleMotivoNotas 
   Caption         =   "Detalle Motivos Notas"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   90
      TabIndex        =   4
      Top             =   0
      Width           =   10215
      Begin VB.TextBox txtCuenta2010 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         MaxLength       =   14
         TabIndex        =   17
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtDescripcion2010 
         Height          =   315
         Left            =   3435
         TabIndex        =   16
         Top             =   2040
         Width           =   6495
      End
      Begin VB.TextBox TXTANO 
         Height          =   315
         Left            =   1920
         TabIndex        =   14
         Top             =   1080
         Width           =   1410
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   3435
         TabIndex        =   13
         Top             =   1500
         Width           =   6495
      End
      Begin VB.TextBox txtCuenta 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         MaxLength       =   14
         TabIndex        =   12
         Top             =   1500
         Width           =   1455
      End
      Begin VB.CheckBox chkCondonacion_Deuda 
         Caption         =   "Condonacion Deuda"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5160
         TabIndex        =   11
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CheckBox chkGasto_Financiero 
         Caption         =   "Gasto Financiero"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CheckBox chkNo_Cantidad_Gupo 
         Caption         =   "No Cantidad Gupo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5160
         TabIndex        =   9
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CheckBox chkMostrar_Grupo 
         Caption         =   "Mostrar Grupo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox txtdescripcion1 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   650
         Width           =   7530
      End
      Begin VB.TextBox txtCod_TipDoc 
         Height          =   315
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   0
         Top             =   240
         Width           =   600
      End
      Begin VB.TextBox txtDes_TipDoc 
         Height          =   315
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta Hasta 2010 :"
         Height          =   210
         Left            =   240
         TabIndex        =   18
         Top             =   2100
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Año :"
         Height          =   210
         Left            =   240
         TabIndex        =   15
         Top             =   1125
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta :"
         Height          =   210
         Left            =   240
         TabIndex        =   7
         Top             =   1565
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Descripción :"
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   705
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Doc :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   255
         Width           =   855
      End
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   3840
      TabIndex        =   3
      Top             =   3600
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"FrmDetalleMotivoNotas.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "FrmDetalleMotivoNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public codigo As String, sOpcion As String, Sid_proyeccion As String
Public Descripcion As String, TipoAdd As String
Dim strSQL As String
Dim SGRUPO As String
Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim sMessage As Long
Dim strSQL As String

On Error GoTo dprError

Select Case ActionName
  Case "ACEPTAR"
     
    If txtCod_TipDoc.Text = "" Then
        MsgBox "Debe ingresar el Tipo de Documento"
        Exit Sub
    End If
    
  
    If MsgBox("Esta seguro de grabar... ", vbYesNo, "IMPORTANTE") = vbYes Then
        Salvar_Datos
        Unload Me
      End If
    
    
  Case "CANCELAR"
      Unload Me

End Select

Exit Sub


dprError:

errores err.Number
End Sub

Sub Salvar_Datos()
On Error GoTo ErrSalvarDatos
Dim VCOD_MOT_NOTA As String, flag As String

If sOpcion = "I" Then
    VCOD_MOT_NOTA = "0"
Else
    VCOD_MOT_NOTA = Sid_proyeccion
End If

If chkGasto_Financiero.Value = 1 Then
    flag = "N"
Else
    flag = "S"
End If


    strSQL = "exec VENTAS_MAN_MOTIVOS_NOTAS '" & sOpcion & "','" & UCase(Trim(txtCod_TipDoc.Text)) & "','" & VCOD_MOT_NOTA & "','" & Trim(txtdescripcion1.Text) & "','" & Trim(txtCuenta.Text) & "','" & flag & "','" & Trim(txtCuenta2010.Text) & "'"
    ExecuteSQL cCONNECT, strSQL
    MsgBox "Se guardó correctamente"
    
        
Exit Sub
ErrSalvarDatos:
    ErrorHandler err, "SALVAR_DATOS"
End Sub



Private Sub txtAno_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      txtCuenta.SetFocus
  End If

End Sub


Private Sub txtCuenta2010_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
            If RTrim(txtCuenta2010.Text) = "" Then
                BUSCA_CUENTACONTABLE_HASTA2010 3
            Else
                BUSCA_CUENTACONTABLE_HASTA2010 1
            End If
    End If
End Sub

Private Sub txtDesCRIPCION_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And TXTANO.Text <> "" Then
        If Len(txtDescripcion.Text) > 3 Then
            BUSCA_CUENTACONTABLE 2
        End If
    End If
End Sub



Private Sub txtCUENTA_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And TXTANO.Text <> "" Then
            If RTrim(txtCuenta.Text) = "" Then
                BUSCA_CUENTACONTABLE 3
            Else
                BUSCA_CUENTACONTABLE 1
            End If
    End If

End Sub

Private Sub BUSCA_CUENTACONTABLE(Tipo As Integer)
On Error GoTo errx
Dim strSQL  As String

    Select Case Tipo
    Case 1:
        strSQL = "SELECT Cod_CtaCont as 'Código', Des_CtaCont as 'Descripción' " & _
                 "FROM CN_PlanContable WHERE ano='" & TXTANO.Text & "' and  Cod_CtaCont like '" & Trim(txtCuenta.Text) & "%' ORDER BY Cod_CtaCont "
                
    Case 2, 3:
            strSQL = "SELECT Cod_CtaCont AS 'Código', " & _
            " Des_CtaCont as 'Descripción' " & _
            "FROM CN_PlanContable " & _
            "WHERE ano='" & TXTANO.Text & "' and  Des_CtaCont LIKE '%" & Trim(Me.txtDescripcion.Text) _
            & "%' AND DATALENGTH(RTRIM(Cod_CtaCont)) = 8 ORDER BY 2"
    End Select
    
    With frmBusqGeneral3
        .Caption = "Buscar Cuenta"
        .sQuery = strSQL
        .Cargar_Datos
        Set .oParent = Me
        
        .gexLista.Columns("Código").Caption = "Código"
        .gexLista.Columns("Descripción").Caption = "Desc. Cuenta"
        
        .gexLista.Columns("Descripción").Width = 4800
                
        If .gexLista.RowCount > 1 Then
            .Show vbModal
        Else
            codigo = .gexLista.Value(.gexLista.Columns("Código").Index)
            Descripcion = .gexLista.Value(.gexLista.Columns("Descripción").Index)
        End If
            
        
        If .gexLista.RowCount > 0 And Not .bCancel Then
            txtCuenta = codigo
            txtDescripcion = Descripcion

            txtCuenta2010.SetFocus
      
            codigo = "": Descripcion = ""
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    Exit Sub
    'Unload frmBusqGeneral3
errx:
    errores err.Number
End Sub



Private Sub BUSCA_CUENTACONTABLE_HASTA2010(Tipo As Integer)
On Error GoTo errx
Dim strSQL  As String

    Select Case Tipo
    Case 1:
        strSQL = "SELECT Cod_CtaCont as 'Código', Des_CtaCont as 'Descripción' " & _
                 "FROM CN_PlanContable WHERE ano='2010' and  Cod_CtaCont like '" & Trim(txtCuenta2010.Text) & "%' ORDER BY Cod_CtaCont "
                
    Case 2, 3:
            strSQL = "SELECT Cod_CtaCont AS 'Código', " & _
            " Des_CtaCont as 'Descripción' " & _
            "FROM CN_PlanContable " & _
            "WHERE ano='2010' and  Des_CtaCont LIKE '%" & Trim(Me.txtDescripcion2010.Text) _
            & "%' AND DATALENGTH(RTRIM(Cod_CtaCont)) = 8 ORDER BY 2"
    End Select
    
    With frmBusqGeneral3
        .Caption = "Buscar Cuenta"
        .sQuery = strSQL
        .Cargar_Datos
        Set .oParent = Me
        
        .gexLista.Columns("Código").Caption = "Código"
        .gexLista.Columns("Descripción").Caption = "Desc. Cuenta"
        
        .gexLista.Columns("Descripción").Width = 4800
                
        If .gexLista.RowCount > 1 Then
            .Show vbModal
        Else
            codigo = .gexLista.Value(.gexLista.Columns("Código").Index)
            Descripcion = .gexLista.Value(.gexLista.Columns("Descripción").Index)
        End If
            
        
        If .gexLista.RowCount > 0 And Not .bCancel Then
            txtCuenta2010 = codigo
            txtDescripcion2010 = Descripcion

            FunctButt1.SetFocus
      
            codigo = "": Descripcion = ""
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    Exit Sub
    'Unload frmBusqGeneral3
errx:
    errores err.Number
End Sub



Private Sub txtCod_TipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtCod_TipDoc.Text) = "" Then
            Call BUSCA_tipo(3)
        Else
            Call BUSCA_tipo(1)
        End If
    End If
End Sub

Public Sub BUSCA_tipo(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    strSQL = "select Des_TipDoc as Descripcion from CN_TiposDocum  where  Flg_Doc_Ventas = '*' and Cod_TipDoc='" & txtCod_TipDoc.Text & "'"
                    txtDes_TipDoc.Text = Trim(DevuelveCampo(strSQL, cCONNECT))
                    txtdescripcion1.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim rs As Object
                    Set rs = CreateObject("ADODB.Recordset")
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.sQuery = "select Cod_TipDoc as codigo, Des_TipDoc as Descripcion  from CN_TiposDocum  where  Flg_Doc_Ventas = '*' and Des_TipDoc like '%" & txtDes_TipDoc.Text & "%'"
                    Else
                        oTipo.sQuery = "select Cod_TipDoc as Codigo , Des_TipDoc as Descripcion from CN_TiposDocum  where  Flg_Doc_Ventas = '*'"
                    End If
                    
                    oTipo.Cargar_Datos
                   ' oTipo.DGridLista.Columns(2).Width = 2500
                    oTipo.Show 1
                    If codigo <> "" Then
                         txtCod_TipDoc.Text = Trim(codigo)
                         txtDes_TipDoc.Text = Trim(Descripcion)

                         codigo = "": Descripcion = ""
                        txtdescripcion1.SetFocus
                    End If
                    Set oTipo = Nothing
                    Set rs = Nothing
    End Select
    
End Sub


Private Sub txtDes_TipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDes_TipDoc.Text) = "" Then
            Call BUSCA_tipo(3)
        Else
            Call BUSCA_tipo(2)
        End If
    End If

End Sub


Private Sub txtdescripcion1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      TXTANO.SetFocus
  End If
End Sub



Private Sub txtDescripcion2010_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(txtDescripcion2010.Text) > 3 Then
            BUSCA_CUENTACONTABLE_HASTA2010 2
        End If
    End If
End Sub
