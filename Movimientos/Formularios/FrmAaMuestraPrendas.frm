VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form FrmAaMuestraPrendas 
   BackColor       =   &H00FFC0C0&
   Caption         =   "MUESTRA PRENDAS"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   1305
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12225
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "X COLOR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   23
         Top             =   900
         Width           =   975
      End
      Begin VB.TextBox txtCod_ColCli 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   3600
         TabIndex        =   22
         Top             =   870
         Width           =   1245
      End
      Begin VB.TextBox txtDes_ColCli 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   4860
         TabIndex        =   21
         Top             =   870
         Width           =   4575
      End
      Begin VB.TextBox txtDes_PurOrd 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   8100
         TabIndex        =   20
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtCod_PurOrd 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   7320
         TabIndex        =   19
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox txtDes_TemCli 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   5220
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtCod_TemCli 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   4440
         TabIndex        =   15
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox txtDes_Cliente 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1950
         TabIndex        =   13
         Top             =   270
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "X CLIENTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "X ESTILO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   11
         Top             =   600
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtCod_Cliente 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1170
         TabIndex        =   10
         Top             =   270
         Width           =   765
      End
      Begin VB.TextBox txtcod_Estilo 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1170
         TabIndex        =   9
         Top             =   570
         Width           =   1245
      End
      Begin VB.TextBox txtCod_Ordpro 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1170
         MaxLength       =   5
         TabIndex        =   8
         Top             =   870
         Width           =   1245
      End
      Begin VB.TextBox txtDes_Estilo 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   2430
         TabIndex        =   7
         Top             =   570
         Width           =   4575
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   11040
         TabIndex        =   6
         Top             =   360
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "X OP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   5
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   18
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TEMP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4040
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "ACEPTAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8880
      TabIndex        =   3
      Top             =   8040
      Width           =   1515
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "CANCELAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10680
      TabIndex        =   2
      Top             =   8040
      Width           =   1515
   End
   Begin VB.CheckBox chkExpandir 
      BackColor       =   &H00FFC0C0&
      Caption         =   "EXPANDIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   11040
      TabIndex        =   1
      Top             =   1320
      Width           =   1155
   End
   Begin VB.CheckBox chkTodos 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   30
      TabIndex        =   0
      Top             =   1320
      Width           =   1155
   End
   Begin GridEX20.GridEX grxDatos 
      Height          =   6435
      Left            =   0
      TabIndex        =   14
      Top             =   1560
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   11351
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      HeaderFontName  =   "Verdana"
      HeaderFontBold  =   -1  'True
      HeaderFontSize  =   6.75
      HeaderFontWeight=   700
      ColumnHeaderHeight=   270
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      FormatStylesCount=   7
      FormatStyle(1)  =   "FrmAaMuestraPrendas.frx":0000
      FormatStyle(2)  =   "FrmAaMuestraPrendas.frx":0138
      FormatStyle(3)  =   "FrmAaMuestraPrendas.frx":01E8
      FormatStyle(4)  =   "FrmAaMuestraPrendas.frx":029C
      FormatStyle(5)  =   "FrmAaMuestraPrendas.frx":0374
      FormatStyle(6)  =   "FrmAaMuestraPrendas.frx":042C
      FormatStyle(7)  =   "FrmAaMuestraPrendas.frx":050C
      ImageCount      =   0
      PrinterProperties=   "FrmAaMuestraPrendas.frx":052C
   End
End
Attribute VB_Name = "FrmAaMuestraPrendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CODIGO As String
Public DESCRIPCION As String
Public oParent As Object
Public Indice As Integer
Public sFec_movActual As String
Public snum_movistkActual As String
Public sCod_ClaMov  As String
Public sCod_TipMov As String
Public sCod_Almacen  As String
Public sTipo_MovConfec As String
Public sTip_Accion As String
Public Cod_ClaMov As String
Private strSQL As String
Private sCod_AlmacenSel As String
Public num_guia As String

Private Sub chkTodos_Click()
    If grxDatos.RowCount = 0 Then Exit Sub
    Dim rs As New ADODB.Recordset
    Dim valor As Boolean
    Dim i As Long

    If chkTodos.Value = Checked Then
        valor = True
    Else
        valor = False
    End If

    grxDatos.Update
    Set rs = grxDatos.ADORecordset
    rs.MoveFirst
    Do While Not rs.EOF
        rs("SEL") = valor
        rs.MoveNext
    Loop

    rs.MoveFirst
    rs.Update
    Set grxDatos.ADORecordset = rs

    CONFIGURA_GRILLA

End Sub
Private Sub Form_Load()
    On Error GoTo SALTO_ERROR
        Dim strSQL As String
        Dim sSeguridad As String
        limpiarctrlBus
        habilitaText (False)
        
        Indice = 2
        strSQL = " select  Cod_ClaMov  from lg_tiposmov where cod_tipmov  = '" & sCod_TipMov & "'"
        Cod_ClaMov = DevuelveCampo(strSQL, cConnect)
        
        txtcod_Estilo.Enabled = True
        txtDes_Estilo.Enabled = True
        
    Exit Sub
SALTO_ERROR:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub chkExpandir_Click()
    If grxDatos.RowCount = 0 Then Exit Sub
    With grxDatos
        Select Case CBool(chkExpandir.Value)
            Case True: .ExpandAll
            Case False: .CollapseAll
        End Select
        If .RowCount > 0 Then
            .Row = -1
            .Col = .Columns.Count - 1
        End If
    End With
End Sub
Private Sub cmdBuscar_Click()
chkExpandir.Value = 1
Call buscarPrendas
End Sub

Private Sub buscarPrendas()

On Error GoTo SALTO_ERROR
  
strSQL = "EXEC CF_MUESTRA_PRENDAS_MOV_TIENDA_ITEMS   '" & Indice & _
                                                    "','" & sCod_Almacen & _
                                                    "','" & Trim(snum_movistkActual) & _
                                                    "','" & Trim(TxtCod_Cliente.Text) & _
                                                    "','" & Trim(txtCod_TemCli.Text) & _
                                                    "','" & Trim(txtDes_ColCli.Text) & _
                                                    "','" & Trim(txtCod_PurOrd.Text) & _
                                                    "','" & Trim(txtcod_Estilo.Text) & _
                                                    "','" & Trim(txtcod_ordpro.Text) & _
                                                    "','" & sCod_ClaMov & _
                                                    "','" & sFec_movActual & "'"
Set grxDatos.ADORecordset = CargarRecordSetDesconectado(strSQL, cConnect)

Call CONFIGURA_GRILLA
    
Exit Sub

SALTO_ERROR:
    MsgBox err.Description, vbCritical, Me.Caption
    
End Sub
Private Sub CONFIGURA_GRILLA()
    On Error GoTo SALTO_ERROR
    Dim C As Integer
    With grxDatos
    
        For C = 1 To .Columns.Count
            .Columns(C).Visible = False
            .Columns(C).HeaderAlignment = jgexAlignCenter
            .Columns(C).TextAlignment = jgexAlignCenter
        Next C

        With .Columns("NOM_CLIENTE")
            .Visible = False
            .Width = 600
            .Caption = "CLIENTE"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("NOM_TEMCLI")
            .Visible = False
            .Width = 1000
            .Caption = "TEMP"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("COD_PURORD")
            .Visible = True
            .Width = 1500
            .Caption = "PO"
            .TextAlignment = jgexAlignLeft
        End With
                        
        With .Columns("DES_ESTCLI")
            .Visible = True
            .Width = 3000
            .Caption = "ESTILO"
            .TextAlignment = jgexAlignLeft
        End With
                        
                        
        With .Columns("Cod_OrdPro")
            .Visible = True
            .Width = 800
            .Caption = "OP"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("Cod_Present")
            .Visible = False
            .Width = 500
            .Caption = "COD_PRESENT"
            .TextAlignment = jgexAlignLeft
        End With
        
        'Presentacion
        With .Columns("DES_PRESENT")
            .Visible = True
            .Width = 2000
            .Caption = "PRESENTACION"
            .TextAlignment = jgexAlignLeft
        End With
        With .Columns("Cod_Talla")
            .Visible = True
            .Width = 600
            .Caption = "TALLA"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("NUM_PRENDAS")
            .Visible = True
            .Width = 1300
            .Caption = "PRENDAS"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("NUM_PRENDASTRANF")
            .Visible = True
            .Width = 1300
            .Caption = "PRENDASTRANF"
            .TextAlignment = jgexAlignLeft
        End With
        
        With .Columns("SEL")
            .Visible = True
            .Width = 500
            .Caption = "SEL"
            .TextAlignment = jgexAlignLeft
        End With
        
        Dim oGroup01 As GridEX20.JSGroup
        Dim oGroup02 As GridEX20.JSGroup
        Dim oGroup03 As GridEX20.JSGroup
    
        Dim colPRENDAS As JSColumn
        Dim colPRENDAtranf As JSColumn
        
    
        With grxDatos
            Set oGroup01 = .Groups.Add(.Columns("NOM_CLIENTE").Index, jgexSortAscending)
            Set oGroup02 = .Groups.Add(.Columns("DES_ESTCLI").Index, jgexSortAscending)
            Set oGroup03 = .Groups.Add(.Columns("COD_ORDPRO").Index, jgexSortAscending)
            
            If chkExpandir.Value = Checked Then
                .DefaultGroupMode = jgexDGMExpanded
            Else
                .DefaultGroupMode = jgexDGMCollapsed
            End If
                
            .GroupFooterStyle = jgexTotalsGroupFooter
            
            Set colPRENDAS = .Columns("NUM_PRENDAS")
            Set colPRENDAtranf = .Columns("NUM_PRENDASTRANF")
            
            With colPRENDAS
                .AggregateFunction = jgexSum
                .TotalRowPrefix = "TOT: "
            End With
            
            With colPRENDAtranf
                .AggregateFunction = jgexSum
                .TotalRowPrefix = "TOT: "
            End With
            
            
        End With
        
        End With
    
    Call SetColores
    
    Exit Sub
    
SALTO_ERROR:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub SetColores()

        grxDatos.Columns("NUM_PRENDAS").CellStyle = "prendas"
        grxDatos.Columns("NUM_PRENDASTRANF").CellStyle = "prendasTrans"
        
        
End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub
Private Sub cmdAceptar_Click()
If grxDatos.RowCount > 0 Then
Call AceptaMovimiento
'frmShowCF_Detalle.BUSCAR
End If
End Sub

Private Sub AceptaMovimiento()
    On Error GoTo SALTO_ERROR
    Dim strSQL As String
    Dim sSeguridad As String
    Dim sEmpresa  As String
    Dim TIPMOV As String
    sEmpresa = "001"
  Dim rs As ADODB.Recordset
  If grxDatos.RowCount > 0 Then
        grxDatos.Update
        Set rs = grxDatos.ADORecordset
  
        rs.MoveFirst
        
        Do While Not rs.EOF
           
           If rs.Fields("sel").Value <> 0 Then
                strSQL = " EXEC LG_UP_MAN_LG_MOVISTK_ITEM_PRENDAS_VENTAS_DIRECTA '" & sCod_Almacen & "'," & _
                                                             "'" & snum_movistkActual & "'," & _
                                                             "'" & Trim(rs.Fields("cod_item").Value) & "'," & _
                                                             "'" & Trim(rs.Fields("cod_comb").Value) & "'," & _
                                                             "'" & Trim(rs.Fields("cod_color").Value) & "'," & _
                                                             "'" & Trim(rs.Fields("cod_estcli").Value) & "'," & _
                                                             "'" & Trim(rs.Fields("COD_ORDPRO").Value) & "'," & _
                                                             "'" & Trim(rs.Fields("COD_PRESENT").Value) & "'," & _
                                                             "'" & Trim(rs.Fields("COD_TALLA").Value) & "'," & _
                                                             "'" & Trim(rs.Fields("codigo_barra").Value) & "'," & _
                                                             "'" & Trim(rs.Fields("NUM_PRENDASTRANF").Value) & "'," & _
                                                             "'I','','','' "
                ExecuteSQL cConnect, strSQL
           End If
           rs.MoveNext
        Loop
        rs.MoveFirst
        Set grxDatos.ADORecordset = rs
        Call buscarPrendas
        grxDatos.SetFocus
    End If
    Call MsgBox("Las se Registro con exito", vbOKOnly, "Mensaje")
    
    Exit Sub
SALTO_ERROR:
    MsgBox err.Description, vbCritical, Me.Caption
    
End Sub

''''******************************HABILITA LA EDICION SOLO DE ALGUNAS COLUMNAS LAS TIENEN CANCEL=FALSE***********************
Private Sub grxDatos_BeforeColEdit(ByVal ColIndex As Integer, ByVal Cancel As GridEX20.JSRetBoolean)
  Select Case ColIndex
    Case Is = grxDatos.Columns("NUM_PRENDASTRANF").Index
      Cancel = False
    Case Is = grxDatos.Columns("SEL").Index
      Cancel = False
    Case Else
      Cancel = True
  End Select
End Sub

Private Sub grxDatos_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
   'grxDatos.Col = 14
End Sub
Private Sub grxDatos_RowFormat(RowBuffer As GridEX20.JSRowData)
'    If grxDatos.RowCount = 0 Then Exit Sub
'    Dim fmtConDIA_Programado As JSFmtCondition
'
'    Set fmtConDIA_Programado = grxDatos.FmtConditions.Add(grxDatos.Columns("pintar").Index, jgexEqual, "S")
'
'    With fmtConDIA_Programado.FormatStyle
'        .ForeColor = &H8000&
'        .FontSize = 8
'        .BackColor = &H80000018 'vbYellow
'    End With
End Sub
Private Sub Option1_Click(Index As Integer)
Indice = Index
Call limpiarctrlBus
Call habilitaText(False)
Set grxDatos.Recordset = Nothing

Select Case Indice
Case 0 'color cliente
    txtCod_ColCli.Enabled = True
    txtDes_ColCli.Enabled = True
               
Case 1 'CLIENTE temp po
  
    TxtCod_Cliente.Enabled = True
    txtDes_Cliente.Enabled = True
    txtCod_TemCli.Enabled = True
    txtDes_TemCli.Enabled = True
    txtCod_PurOrd.Enabled = True
    txtDes_PurOrd.Enabled = True
Case 2 'ESTILO
    txtcod_Estilo.Enabled = True
    txtDes_Estilo.Enabled = True
 
Case 3 'OP
    txtcod_ordpro.Enabled = True
    
End Select

End Sub
Private Sub limpiarctrlBus()
    
    TxtCod_Cliente.Text = ""
    txtDes_Cliente.Text = ""
    txtCod_TemCli.Text = ""
    txtDes_TemCli.Text = ""
    txtCod_PurOrd.Text = ""
    txtDes_PurOrd.Text = ""
    txtcod_Estilo.Text = ""
    txtDes_Estilo.Text = ""
    txtcod_ordpro.Text = ""
    txtCod_ColCli.Text = ""
    txtDes_ColCli.Text = ""
    
End Sub
Private Sub habilitaText(valor As Boolean)
    
    TxtCod_Cliente.Enabled = valor
    txtDes_Cliente.Enabled = valor
    txtCod_TemCli.Enabled = valor
    txtDes_TemCli.Enabled = valor
    txtCod_PurOrd.Enabled = valor
    txtDes_PurOrd.Enabled = valor
    txtcod_Estilo.Enabled = valor
    txtDes_Estilo.Enabled = valor
    txtcod_ordpro.Enabled = valor
    txtCod_ColCli.Enabled = valor
    txtDes_ColCli.Enabled = valor

End Sub
Private Sub txtCod_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_Cliente", "Nom_Cliente", "tg_cliente where  ", TxtCod_Cliente, txtDes_Cliente, 1, Me)
    'cmdBuscar.SetFocus
    txtCod_TemCli.SetFocus
End If
End Sub
Private Sub TxtDes_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Busca_Opcion("Cod_Cliente", "Nom_Cliente", "tg_cliente where  ", TxtCod_Cliente, txtDes_Cliente, 2, Me)
       ' cmdBuscar.SetFocus
       txtCod_TemCli.SetFocus
    End If
End Sub
Private Sub txtCod_TemCli_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Opcion("cod_temcli", "Nom_TemCli", "tg_temcli where cod_cliente='" & TxtCod_Cliente.Text & "' and ", txtCod_TemCli, txtDes_TemCli, 1, Me)
    txtCod_PurOrd.SetFocus
End If
End Sub
Private Sub txtDes_TemCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Busca_Opcion("cod_temcli", "Nom_TemCli", "tg_temcli where cod_cliente='" & TxtCod_Cliente.Text & "' and ", txtCod_TemCli, txtDes_TemCli, 2, Me)
        txtCod_PurOrd.SetFocus
    End If
End Sub
Private Sub txtCod_PurOrd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Opcion("cod_purord", "cod_purord", "TG_PURORD where cod_temcli='" & txtCod_TemCli.Text & "' and ", txtCod_PurOrd, txtDes_PurOrd, 1, Me)
    CmdBuscar.SetFocus
End If
End Sub
Private Sub txtDes_PurOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Busca_Opcion("cod_purord", "cod_purord", "TG_PURORD where cod_temcli='" & txtCod_TemCli.Text & "' and ", txtCod_PurOrd, txtDes_PurOrd, 2, Me)
     CmdBuscar.SetFocus
    End If
End Sub
Private Sub txtcod_Estilo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_EstCli", "Des_EstCli", "Tg_EstCliTem where  ", txtcod_Estilo, txtDes_Estilo, 1, Me)
    CmdBuscar.SetFocus
End If
End Sub
Private Sub txtDes_Estilo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Opcion("Cod_EstCli", "Des_EstCli", "Tg_EstCliTem where ", txtcod_Estilo, txtDes_Estilo, 2, Me)
    CmdBuscar.SetFocus
End If
End Sub
Private Sub txtCod_ColCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Busca_Opcion("DES_PRESENT ", "DES_PRESENT", " ES_ESTPROPRE  where ", txtCod_ColCli, txtDes_ColCli, 1, Me)
        CmdBuscar.SetFocus
    End If
End Sub
Private Sub txtdes_colcli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Busca_Opcion("DES_PRESENT", "DES_PRESENT", " ES_ESTPROPRE  where ", txtCod_ColCli, txtDes_ColCli, 2, Me)
        CmdBuscar.SetFocus
    End If
End Sub

Private Sub txtcod_ordpro_KeyPress(KeyAscii As Integer)
    Dim iLen As Integer
    Dim sSQl As String
    Dim varCod_Fabrica  As String
    
    varCod_Fabrica = "001"
    
    If KeyAscii = vbKeyReturn Then
    
        If RTrim(txtcod_ordpro.Text) <> "" Then
            txtcod_ordpro.Text = txtcod_ordpro
            If BUSCA_OP(1) Then
            
                'txtPresentacion.SetFocus
                CmdBuscar.SetFocus
                
            Else
                Aviso "La O/P ingresada no es valida o no existe. Sirvase verificar", 1
            End If
        Else
            If BUSCA_OP(2) Then
               ' txtPresentacion.SetFocus
               CmdBuscar.SetFocus
            End If
        End If
    End If
End Sub
Private Function BUSCA_OP(ByVal iModo As Integer)
    Dim sSQl As String
    Dim sRet As String
    Dim oTipo As New frmBusqGeneralJanus
    Dim rs As New ADODB.Recordset
    Dim varCod_Fabrica As String
    
    Dim varCod_EstPro As String
    
    varCod_Fabrica = "001"
    Select Case iModo
        Case 1
            'sSQl = "UP_MUESTRA_OP '" & varCod_Fabrica & "','" & Trim(txtop.Text) & "','" & Trim(txtCodCliente.Text) & "','','',''"
            sSQl = "UP_MUESTRA_OP '" & varCod_Fabrica & "','" & Trim(txtcod_ordpro.Text) & "','','','',''"
            
            Set rs = CargarRecordSetDesconectado(cConnect, sSQl)
            
            If Not rs.EOF Then
                txtcod_ordpro.Text = rs!CODIGO ' rs!Cod_OrdPro
                'txtDes_OrdPro.Text = rs!DESCRIPCION 'rs!Des_OrdPro
                varCod_EstPro = rs!Cod_Estpro
                'FunctBuscar.SetFocus
                BUSCA_OP = True
            Else
                Exit Function
            End If
            
        Case 2
            sSQl = "UP_MUESTRA_OP '" & varCod_Fabrica & "','" & Trim(txtcod_ordpro.Text) & "','','','',''"
            
            Set oTipo.oParent = Me
            oTipo.sQuery = sSQl
            
            oTipo.Cargar_Datos
            oTipo.Show 1
            If CODIGO <> "" Then
                txtcod_ordpro.Text = Trim(CODIGO)
                'txtDes_OrdP.Text = Trim(DESCRIPCION)
'                Me.varCod_EstPro = Cod_Estpro
                'FunctBuscar.SetFocus
                BUSCA_OP = True
            End If
            Set oTipo = Nothing
            Set rs = Nothing
            
    End Select
    
End Function

Public Sub Busca_Opcion(strCampo1 As String, strCampo2 As String, strTabla As String, txtCod As TextBox, txtDes As TextBox, Opcion As Integer, frmME As Form)

On Error GoTo Fin

Dim rstAux As ADODB.Recordset, strSQL As String

    strSQL = "Select DISTINCT " & strCampo1 & " AS Cod," & strCampo2 & " as Descripcion from " & strTabla

    txtCod = Trim(txtCod)
    txtDes = Trim(txtDes)
    
    
    Select Case Opcion
    Case 1: strSQL = strSQL & strCampo1 & " like '%" & txtCod & "%'"
    Case 2: strSQL = strSQL & strCampo2 & " like '%" & txtDes & "%'"
   
    
    End Select
    txtCod = ""
    txtDes = ""
    
    With frmBusqGeneral
        Set .oParent = frmME
        .sQuery = strSQL
        .Cargar_Datos
        
        frmME.CODIGO = ""
        Set rstAux = .gexList.ADORecordset
        If rstAux.RecordCount > 1 Then
          .Show vbModal
        Else
          frmME.CODIGO = ".."
        End If
        
        If frmME.CODIGO <> "" And rstAux.RecordCount > 0 Then
            txtCod = frmME.CODIGO 'Trim(rstAux!Cod)
            txtDes = frmME.DESCRIPCION  'Trim(rstAux!Descripcion)
            
            If txtCod = ".." Or frmME.DESCRIPCION = "" Then
                txtCod = Trim(rstAux!cod)
                txtDes = Trim(rstAux!DESCRIPCION)
            End If
            
            Select Case Opcion
            Case 1: SendKeys "{TAB}": SendKeys "{TAB}"
            Case 2: SendKeys "{TAB}"
            End Select
            
        Else
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
    MsgBox err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Búsqueda de Descuento (" & Opcion & ")"
End Sub



