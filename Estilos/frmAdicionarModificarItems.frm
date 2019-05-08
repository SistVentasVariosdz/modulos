VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmAdicionarModificarItems 
   Caption         =   "Adicionar"
   ClientHeight    =   6576
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11568
   LinkTopic       =   "Form1"
   ScaleHeight     =   6576
   ScaleWidth      =   11568
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTecnicaEstampado 
      Height          =   288
      Left            =   5400
      TabIndex        =   15
      Top             =   3720
      Width           =   6012
   End
   Begin VB.TextBox txtPrecioComercial 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   1800
      TabIndex        =   14
      Text            =   "0"
      Top             =   3720
      Width           =   1812
   End
   Begin VB.ComboBox cboIde_TallaX 
      Height          =   288
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   5520
      Width           =   585
   End
   Begin VB.CommandButton cmdGrafico 
      Caption         =   "..."
      Height          =   315
      Left            =   10440
      TabIndex        =   12
      Top             =   2700
      Width           =   435
   End
   Begin VB.TextBox txtDirIcono 
      Height          =   300
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2700
      Width           =   8490
   End
   Begin VB.TextBox txtDesStatus 
      Height          =   285
      Left            =   2280
      TabIndex        =   30
      Top             =   8520
      Width           =   2895
   End
   Begin VB.TextBox txtCodStatus 
      Height          =   285
      Left            =   1560
      TabIndex        =   29
      Top             =   8520
      Width           =   615
   End
   Begin VB.TextBox txtDesTipoVersion 
      Height          =   285
      Left            =   2280
      TabIndex        =   32
      Top             =   8880
      Width           =   2895
   End
   Begin VB.TextBox txtCodTipoVersion 
      Height          =   285
      Left            =   1560
      TabIndex        =   31
      Top             =   8880
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   120
      TabIndex        =   56
      Top             =   4080
      Width           =   11130
      Begin VB.TextBox txtUniMedProv 
         Height          =   285
         Left            =   9720
         MaxLength       =   2
         TabIndex        =   19
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txtNombreProveedor 
         Height          =   285
         Left            =   2640
         TabIndex        =   17
         Top             =   315
         Width           =   4200
      End
      Begin VB.TextBox txtCodProveedor 
         Height          =   285
         Left            =   1350
         TabIndex        =   16
         Top             =   315
         Width           =   1245
      End
      Begin VB.TextBox txtCodItemProv 
         Height          =   285
         Left            =   7755
         MaxLength       =   15
         TabIndex        =   18
         Top             =   315
         Width           =   1335
      End
      Begin VB.TextBox txtObservacionesProv 
         Height          =   285
         Left            =   4815
         TabIndex        =   21
         Top             =   765
         Width           =   6210
      End
      Begin VB.TextBox txtPrecio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1350
         MaxLength       =   8
         TabIndex        =   20
         Text            =   "0"
         Top             =   765
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código/Nombre :"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Código del Prov:"
         Height          =   390
         Left            =   6975
         TabIndex        =   60
         Top             =   285
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "UniMed"
         Height          =   195
         Left            =   9075
         TabIndex        =   59
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Precio Cotizado$:"
         Height          =   195
         Left            =   135
         TabIndex        =   58
         Top             =   795
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones del Proveedor"
         Height          =   195
         Left            =   2670
         TabIndex        =   57
         Top             =   795
         Width           =   2100
      End
   End
   Begin VB.TextBox TxtModo 
      Height          =   285
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox TxtDes_modo 
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   1455
      Width           =   2895
   End
   Begin VB.TextBox txtDesOrigen 
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Top             =   1900
      Width           =   2895
   End
   Begin VB.TextBox txtCodOrigen 
      Height          =   285
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1900
      Width           =   615
   End
   Begin VB.TextBox txtDesMotivo 
      Height          =   285
      Left            =   9120
      TabIndex        =   34
      Top             =   8640
      Width           =   2895
   End
   Begin VB.TextBox txtCodMotivo 
      Height          =   285
      Left            =   8400
      TabIndex        =   33
      Top             =   8640
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   40
      Top             =   8040
      Width           =   11175
      Begin VB.ComboBox cboIde_Talla 
         Height          =   315
         Left            =   1118
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   360
         Width           =   1185
      End
      Begin VB.ComboBox cboIde_EsCli 
         Height          =   315
         Left            =   3038
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   360
         Width           =   1185
      End
      Begin VB.ComboBox cboIde_Color 
         Height          =   315
         Left            =   5558
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   360
         Width           =   1185
      End
      Begin VB.ComboBox cboIde_Destino 
         Height          =   315
         Left            =   7598
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   360
         Width           =   1185
      End
      Begin VB.ComboBox CboIde_PO 
         Height          =   315
         Left            =   9518
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Talla :"
         Height          =   195
         Left            =   638
         TabIndex        =   53
         Top             =   405
         Width           =   435
      End
      Begin VB.Label Label17 
         Caption         =   "Estilo Cliente :"
         Height          =   375
         Left            =   2438
         TabIndex        =   52
         Top             =   255
         Width           =   765
      End
      Begin VB.Label Label16 
         Caption         =   "Color Cliente :"
         Height          =   270
         Left            =   4478
         TabIndex        =   51
         Top             =   435
         Width           =   960
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Destino :"
         Height          =   195
         Left            =   6878
         TabIndex        =   50
         Top             =   420
         Width           =   630
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "P.O. :"
         Height          =   195
         Left            =   9038
         TabIndex        =   49
         Top             =   420
         Width           =   405
      End
   End
   Begin VB.TextBox txtUbicacion 
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Top             =   2250
      Width           =   9135
   End
   Begin VB.TextBox txtComentario 
      Height          =   495
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   3120
      Width           =   9672
   End
   Begin VB.TextBox txtDesGrupo 
      Height          =   285
      Left            =   9120
      TabIndex        =   28
      Top             =   9000
      Width           =   2535
   End
   Begin VB.TextBox txtCodGrupo 
      Height          =   285
      Left            =   8400
      TabIndex        =   27
      Top             =   9000
      Width           =   615
   End
   Begin VB.TextBox txtDesClase 
      Height          =   285
      Left            =   2385
      TabIndex        =   26
      Top             =   6765
      Width           =   2880
   End
   Begin VB.TextBox txtCodClase 
      Height          =   285
      Left            =   1665
      TabIndex        =   25
      Top             =   6765
      Width           =   615
   End
   Begin VB.TextBox txtDesUM 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   1050
      Width           =   2865
   End
   Begin VB.TextBox txtCodUM 
      Height          =   285
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1050
      Width           =   615
   End
   Begin VB.TextBox txtDesFamilia 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   570
      Width           =   2880
   End
   Begin VB.TextBox txtCodFamilia 
      Height          =   285
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   2
      Top             =   580
      Width           =   615
   End
   Begin VB.TextBox txtcoditem 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   0
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox txtDesItem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      MaxLength       =   100
      TabIndex        =   1
      Top             =   90
      Width           =   8160
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   516
      Left            =   4680
      TabIndex        =   23
      Top             =   5880
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   910
      Custom          =   $"frmAdicionarModificarItems.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Label Label23 
      Caption         =   "Tecnica Estampado"
      Height          =   252
      Left            =   3840
      TabIndex        =   66
      Top             =   3720
      Width           =   1572
   End
   Begin VB.Label Label22 
      Caption         =   "Precio Comercial($)"
      Height          =   252
      Left            =   240
      TabIndex        =   65
      Top             =   3720
      Width           =   1452
   End
   Begin VB.Label Label19 
      Caption         =   "Identificador de Talla"
      Height          =   276
      Left            =   180
      TabIndex        =   64
      Top             =   5520
      Width           =   1668
   End
   Begin VB.Label Label18 
      Caption         =   "Dirrección de Icono:"
      Height          =   375
      Left            =   240
      TabIndex        =   63
      Top             =   2670
      Width           =   1080
   End
   Begin VB.Label Label12 
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   62
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo Versión"
      Height          =   255
      Left            =   120
      TabIndex        =   55
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Modo Proceso"
      Height          =   255
      Left            =   240
      TabIndex        =   54
      Top             =   1455
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Ubicación en la Prenda"
      Height          =   390
      Left            =   240
      TabIndex        =   48
      Top             =   2205
      Width           =   1215
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Comentario :"
      Height          =   195
      Left            =   240
      TabIndex        =   47
      Top             =   3210
      Width           =   885
   End
   Begin VB.Label Etiqueta 
      AutoSize        =   -1  'True
      Caption         =   "Familia Item:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   240
      TabIndex        =   46
      Tag             =   "Mat. Prima :"
      Top             =   540
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Unidad de Medida :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   45
      Tag             =   "Porcentaje :"
      Top             =   1050
      Width           =   1365
   End
   Begin VB.Label Label9 
      Caption         =   "Grupo de Item :"
      Height          =   255
      Left            =   7110
      TabIndex        =   44
      Top             =   9000
      Width           =   1095
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Clase de Item :"
      Height          =   195
      Left            =   345
      TabIndex        =   43
      Top             =   6795
      Width           =   1050
   End
   Begin VB.Label Label13 
      Caption         =   "Origen :"
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   1890
      Width           =   735
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Motivo Preproduc :"
      Height          =   195
      Left            =   6960
      TabIndex        =   41
      Top             =   8640
      Width           =   1350
   End
   Begin VB.Label lblCod_Item 
      AutoSize        =   -1  'True
      Caption         =   "Item :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   240
      TabIndex        =   24
      Tag             =   "Hilado :"
      Top             =   165
      Width           =   375
   End
End
Attribute VB_Name = "frmAdicionarModificarItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As String
Public sTipo As String
Public Codigo      As String
Public Descripcion As String
Public Abr_Cliente As String
Public sTemporada  As String
Public ruta As String
Public strImagenCambio As String
Public oParent As Object
Dim StrSQL As String



Private Sub cboIde_TallaX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       Me.FunctButt1.SetFocus
End If
End Sub

Private Sub cmdGrafico_Click()
Dim ofIcon As New frmCambiaGrafico


ofIcon.StrImagen1_Origen = strImagenCambio
ofIcon.Codigo_item = txtcoditem.Text
ofIcon.Show 1

If RTrim(ofIcon.ruta_imagenes <> "") Then
    txtDirIcono.Text = Trim(ofIcon.ruta_imagenes)
    strImagenCambio = Trim(ofIcon.StrImagen_cambio)
End If
Set ofIcon = Nothing
txtComentario.SetFocus

End Sub



Private Sub Form_Load()
Call CargarCombos

    cboIde_TallaX.Clear
    cboIde_TallaX.AddItem ("N")
    cboIde_TallaX.AddItem ("S")
    cboIde_TallaX.ListIndex = 0
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPTAR"
Call Grabar

Unload Me
 
Case "CANCELAR"
    Unload Me
End Select
End Sub

Sub Grabar()
On Error GoTo hand
Dim Rs As ADODB.Recordset

  If MsgBox("Esta seguro de actualizar los datos", vbInformation + vbYesNo, "AVISO") = vbYes Then
  
  
    Dim sql As String
    sql = "SELECT Cod_Cliente FROM TG_CLIENTE WHERE Abr_Cliente='" & Trim(Abr_Cliente) & "'"
       
             
    StrSQL = "UP_MAN_ITEMS2 " & _
        Opcion & ",'" & _
        sTipo & "','" & _
        Trim(txtcoditem.Text) & "','" & Trim(txtCodFamilia.Text) & "','" & _
        Trim(txtCodGrupo.Text) & "','" & Trim(txtCodUM.Text) & "','" & _
        Trim(txtDesItem.Text) & "','" & Trim(txtCodClase.Text) & "','" & _
        Trim(txtCodOrigen.Text) & "','" & Trim(cboIde_TallaX.Text) & "','" & _
        Trim(cboIde_Color.Text) & "','" & Trim(cboIde_EsCli.Text) & "','" & _
        Trim(cboIde_Destino.Text) & "','" & _
        Trim(txtCodMotivo.Text) & "','" & _
        DevuelveCampo(sql, cCONNECT) & "','" & _
        Trim(sTemporada) & "','" & _
        Trim(txtComentario.Text) & "','" & _
        Trim(CboIde_PO.Text) & "','" & _
        vusu & "','" & _
        Trim(txtUbicacion.Text) & "','" & _
        Trim(txtCodStatus.Text) & "','" & _
        Trim(txtCodTipoVersion.Text) & "','" & _
        Trim(TxtModo.Text) & "','" & Trim(txtDirIcono.Text) & "','" & Trim(txtCodProveedor.Text) & "','" & _
        Trim(txtCodItemProv.Text) & "','" & Trim(txtUniMedProv.Text) & "','" & _
        Trim(txtPrecio.Text) & "','" & Trim(txtObservacionesProv.Text) & "','" & _
        Trim(Me.txtPrecioComercial) & "','" & Trim(Me.txtTecnicaEstampado) & "'"
           
   
      
   Set Rs = GetRecordset(cCONNECT, StrSQL)
   
    If Not Rs Is Nothing And sTipo = "I" Then
        txtcoditem = Rs!Cod_Item
    End If
    
     ChangeName
     
     Call Move_Files("")
     
    End If
    
    oParent.CargaLista
    oParent.FindItem txtcoditem
    
    Exit Sub
    
hand:
   ErrorHandler Err, "GRABAR"
   'Err.Raise Err.Number, Err.Source, Err.Description
End Sub




 
Private Sub txtCodGrupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaGrupo 1
    txtCodStatus.SetFocus
End If
End Sub


Private Sub txtCodItemProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtUniMedProv.SetFocus
End If
End Sub

 

Private Sub txtDesGrupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaGrupo 2
    txtCodStatus.SetFocus
End If
End Sub


 

Private Sub BuscaGrupo(Opcion As Integer)
Dim sField As String, iRows As Long
Dim rstAux As ADODB.Recordset
    
    StrSQL = "Select Cod_Gruitem, des_famgruite From LG_FamGruIte WHERE  Cod_Famitem='" & txtCodFamilia.Text & "' and "
    txtCodGrupo = Trim(txtCodGrupo)
    txtDesGrupo = Trim(txtDesGrupo)
    sField = txtCodGrupo
    Select Case Opcion
    Case 1: StrSQL = StrSQL & "Cod_Gruitem like '%" & txtCodGrupo & "%'"
    Case 2: StrSQL = StrSQL & "des_famgruite like '%" & txtDesGrupo & "%'"
    End Select
    
    txtCodGrupo = ""
    txtDesGrupo = ""
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = StrSQL
        .Cargar_Datos
        
        Codigo = ""
        Descripcion = ""

        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            txtCodGrupo = rstAux!Cod_Gruitem
            txtDesGrupo = rstAux!des_famgruite
            If iRows = 1 And Opcion = 1 And _
            sField = "" Then
                'txtCodGrupo.Enabled = False
                'txtDesGrupo.Enabled = False
            End If
            'SendKeys "{TAB}"
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub

Private Sub txtCodFamilia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaFamilia 1
    txtCodUM.SetFocus
End If
End Sub



Private Sub txtDesFamilia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaFamilia 2
    txtCodUM.SetFocus
End If
End Sub

Private Sub BuscaFamilia(Opcion As Integer)
Dim sField As String, iRows As Long
Dim rstAux As ADODB.Recordset
    
    StrSQL = "Select cod_famitem, des_famitem From LG_FamIte WHERE flg_proceso_confec  = 'S' AND "
    txtCodFamilia = Trim(txtCodFamilia)
    txtDesFamilia = Trim(txtDesFamilia)
    sField = txtCodFamilia
    Select Case Opcion
    Case 1: StrSQL = StrSQL & "cod_famitem like '%" & txtCodFamilia & "%'"
    Case 2: StrSQL = StrSQL & "des_famitem like '%" & txtDesFamilia & "%'"
    End Select
    
    txtCodFamilia = ""
    txtDesFamilia = ""
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = StrSQL
        .Caption = "Seleccionar - Familia Item"
        .Cargar_Datos
        
        Codigo = ""
        Descripcion = ""
        
        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If .DGridLista.RowCount > 1 Then
            .Show vbModal
        ElseIf .DGridLista.RowCount = 1 Then
            Codigo = .DGridLista.Value(.DGridLista.Columns("COD_FAMITEM").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("DES_FAMITEM").Index)
        End If
        
        If Codigo <> "" Then
            txtCodFamilia = RTrim(Codigo)
            txtDesFamilia = RTrim(Descripcion)
            Codigo = "": Descripcion = ""
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub


Private Sub txtCodUM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaUniMedida 1
    TxtModo.SetFocus
    If RTrim(txtUniMedProv.Text) = "" Then
    txtUniMedProv.Text = txtCodUM.Text
    End If
End If
End Sub

Private Sub txtDesItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCodFamilia.SetFocus
End If

End Sub

Private Sub txtDesUM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaUniMedida 2
    TxtModo.SetFocus
End If
End Sub

Private Sub BuscaUniMedida(Opcion As Integer)
Dim sField As String, iRows As Long
Dim rstAux As ADODB.Recordset
    
    StrSQL = "Select Cod_UniMed, Des_UniMed From TG_UniMed WHERE "
    txtCodUM = Trim(txtCodUM)
    txtDesUM = Trim(txtDesUM)
    sField = txtCodUM
    Select Case Opcion
    Case 1: StrSQL = StrSQL & "Cod_UniMed like '%" & txtCodUM & "%'"
    Case 2: StrSQL = StrSQL & "Des_UniMed like '%" & txtDesUM & "%'"
    End Select
    
    txtCodUM = ""
    txtDesUM = ""
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = StrSQL
         .Caption = "Seleccionar - Unidad Medida"
        .Cargar_Datos
        
        Codigo = ""
        Descripcion = ""

        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If .DGridLista.RowCount > 1 Then
            .Show vbModal
        ElseIf .DGridLista.RowCount = 1 Then
            Codigo = .DGridLista.Value(.DGridLista.Columns("COD_unimed").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("DES_unimed").Index)
        End If
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            txtCodUM = Codigo
            txtDesUM = Descripcion
            If iRows = 1 And Opcion = 1 And _
            sField = "" Then
               ' txtCodUM.Enabled = False
                'txtDesUM.Enabled = False
            End If
            'SendKeys "{TAB}"
            Codigo = "": Descripcion = ""
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub



 
 
Private Sub txtCodClase_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        BuscaClaseItem 1
        TxtModo.SetFocus
End If
End Sub

 Private Sub txtDesClase_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaClaseItem 2
    TxtModo.SetFocus
End If
End Sub

Private Sub BuscaClaseItem(Opcion As Integer)
Dim sField As String, iRows As Long
Dim rstAux As ADODB.Recordset
    
    StrSQL = "Select cod_claitem, des_claitem From LG_Claitem WHERE "
    txtCodClase = Trim(txtCodClase)
    txtDesClase = Trim(txtDesClase)
    sField = txtCodClase
    Select Case Opcion
    Case 1: StrSQL = StrSQL & "cod_claitem like '%" & txtCodClase & "%'"
    Case 2: StrSQL = StrSQL & "des_claitem like '%" & txtDesClase & "%'"
    End Select
    
    txtCodClase = ""
    txtDesClase = ""
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = StrSQL
         .Caption = "Seleccionar - Clase Item"
        .Cargar_Datos
        
        Codigo = ""
        Descripcion = ""

        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            txtCodClase = rstAux!cod_claitem
            txtDesClase = rstAux!des_claitem
            If iRows = 1 And Opcion = 1 And _
            sField = "" Then
                'txtCodClase.Enabled = False
               ' txtDesClase.Enabled = False
            End If
            'SendKeys "{TAB}"
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub




 
Private Sub txtCodMotivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        BuscaMotivoProduccion 1
        txtCodOrigen.SetFocus
End If
End Sub

 Private Sub txtDesMotivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaMotivoProduccion 2
     txtCodOrigen.SetFocus
End If
End Sub

   
   
Private Sub BuscaMotivoProduccion(Opcion As Integer)
Dim sField As String, iRows As Long
Dim rstAux As ADODB.Recordset
    
    StrSQL = "Select cod_motprepro, des_motprepro From TG_MotPrePro WHERE  "
    txtCodMotivo = Trim(txtCodMotivo)
    txtDesMotivo = Trim(txtDesMotivo)
    sField = txtCodGrupo
    Select Case Opcion
    Case 1: StrSQL = StrSQL & "cod_motprepro like '%" & txtCodMotivo & "%'"
    Case 2: StrSQL = StrSQL & "des_motprepro like '%" & txtDesMotivo & "%'"
    End Select
    
    txtCodMotivo = ""
    txtDesMotivo = ""
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = StrSQL
        .Cargar_Datos
        
        Codigo = ""
        Descripcion = ""
        
        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            txtCodMotivo = rstAux!cod_motprepro
            txtDesMotivo = rstAux!des_motprepro
            If iRows = 1 And Opcion = 1 And _
            sField = "" Then
               ' txtCodMotivo.Enabled = False
               ' txtDesMotivo.Enabled = False
            End If
            'SendKeys "{TAB}"
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub
       
   
   
Private Sub txtCodOrigen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        BuscaOrigen 1
        txtUbicacion.SetFocus
End If
End Sub

 Private Sub txtDesOrigen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaOrigen 2
       txtUbicacion.SetFocus
End If
End Sub

   
   
   
Private Sub BuscaOrigen(Opcion As Integer)
Dim sField As String, iRows As Long
Dim rstAux As ADODB.Recordset
    
    StrSQL = "Select cod_Origen, des_origen From LG_Origen WHERE "
    txtCodOrigen = Trim(txtCodOrigen)
    txtDesOrigen = Trim(txtDesOrigen)
    sField = txtCodOrigen
    Select Case Opcion
    Case 1: StrSQL = StrSQL & "cod_Origen like '%" & txtCodOrigen & "%'"
    Case 2: StrSQL = StrSQL & "des_origen like '%" & txtDesOrigen & "%'"
    End Select
    
    txtCodOrigen = ""
    txtDesOrigen = ""
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = StrSQL
         .Caption = "Seleccionar - Busca Origen"
        .Cargar_Datos
        
        Codigo = ""
        Descripcion = ""
        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If .DGridLista.RowCount > 1 Then
            .Show vbModal
        ElseIf .DGridLista.RowCount = 1 Then
            Codigo = .DGridLista.Value(.DGridLista.Columns("COD_ORIGEN").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("DES_ORIGEN").Index)
        End If
        
        If Codigo <> "" Then
            txtCodOrigen = Codigo
            txtDesOrigen = Descripcion
            Codigo = "": Descripcion = ""
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub

Private Sub txtCodStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        BuscaStatus 1
         txtCodTipoVersion.SetFocus
End If
End Sub

 Private Sub txtDesStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaStatus 2
        txtCodTipoVersion.SetFocus
End If
End Sub

   
 'StrSQL = "SELECT des_status + space(100) + flg_status  FROM TG_StaDes"
   
   Private Sub BuscaStatus(Opcion As Integer)
   Dim sField As String, iRows As Long
   Dim rstAux As ADODB.Recordset
    
    StrSQL = "Select flg_status, des_status From TG_StaDes WHERE "
    txtCodStatus = Trim(txtCodStatus)
    txtDesStatus = Trim(txtDesStatus)
    sField = txtCodStatus
    Select Case Opcion
    Case 1: StrSQL = StrSQL & "flg_status like '%" & txtCodStatus & "%'"
    Case 2: StrSQL = StrSQL & "des_status like '%" & txtDesStatus & "%'"
    End Select
    
    txtCodStatus = ""
    txtDesStatus = ""
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = StrSQL
        .Cargar_Datos
        
        Codigo = ""
        Descripcion = ""

        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            txtCodStatus = rstAux!flg_status
            txtDesStatus = rstAux!des_status
            If iRows = 1 And Opcion = 1 And _
            sField = "" Then
                'txtCodOrigen.Enabled = False
                'txtDesOrigen.Enabled = False
            End If
            'SendKeys "{TAB}"
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub
   


    
Sub CargarCombos()
     'Combo Identificador Talla
    cboIde_Talla.Clear
    cboIde_Talla.AddItem ("N")
    cboIde_Talla.AddItem ("S")
    cboIde_Talla.ListIndex = 0
    'Combo Identificador Color
    cboIde_Color.Clear
    cboIde_Color.AddItem ("N")
    cboIde_Color.AddItem ("S")
    cboIde_Color.ListIndex = 0
    'Combo Identificador Estilo Cliente
    cboIde_EsCli.Clear
    cboIde_EsCli.AddItem ("N")
    cboIde_EsCli.AddItem ("S")
    cboIde_EsCli.ListIndex = 0
    'Combo Identificador de Destino
    cboIde_Destino.Clear
    cboIde_Destino.AddItem ("N")
    cboIde_Destino.AddItem ("S")
    cboIde_Destino.ListIndex = 0
    'Combo Identificador de p.o.
    CboIde_PO.Clear
    CboIde_PO.AddItem ("N")
    CboIde_PO.AddItem ("S")
    CboIde_PO.ListIndex = 0
End Sub
    
Private Sub txtDirIcono_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   
    txtComentario.SetFocus
End If
End Sub

 
'Private Sub TxtIdeTalla_GotFocus()
'SelectionText TxtIdeTalla
'End Sub

Private Sub TxtIdeTalla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       FunctButt1.SetFocus
End If
 
End Sub

    Private Sub TxtModo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Modo(1, 1)
    txtCodOrigen.SetFocus
End If
End Sub




Private Sub TxtDes_modo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Modo(2, 1)
     txtCodOrigen.SetFocus
End If
End Sub
    
    Public Sub Busca_Modo(Opcion As Integer, tipo As Integer)
Dim rstAux As ADODB.Recordset
On Error GoTo Fin
Dim iCol As Long

    StrSQL = "SELECT Flg_ModoProceso as Codigo, Des_ModoProceso as Descripcion FROM ES_ModoProceso where "
    
    Select Case Opcion
    Case 1: StrSQL = StrSQL & " Flg_ModoProceso like '%" & Trim(TxtModo.Text) & "%'"
    Case 2: StrSQL = StrSQL & " Des_ModoProceso like '%" & Trim(TxtDes_modo.Text) & "%'"
    End Select
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = StrSQL
        .Cargar_Datos
        .Caption = "Seleccionar - Modo Proceso"
        Codigo = ""
        Descripcion = ""
        
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Codigo").Width = 700
        .DGridLista.Columns("Descripcion").Width = 5000
        
        If .DGridLista.RowCount > 1 Then
            .Show vbModal
        ElseIf .DGridLista.RowCount = 1 Then
            Codigo = .DGridLista.Value(.DGridLista.Columns("Codigo").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("Descripcion").Index)
        End If

        
        If Codigo <> "" Then
            TxtModo = Codigo
            TxtDes_modo = Descripcion
           
        End If
    End With
    Codigo = "": Descripcion = ""
    Unload frmBusqGeneral3
    Set frmBusqGeneral = Nothing
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral3
    Set frmBusqGeneral = Nothing
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Busca Unidad Medida (" & Opcion & ")"
End Sub

Private Sub txtCodTipoVersion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaTipoVersion 1
    TxtModo.SetFocus
End If
End Sub
Private Sub txtdesTipoVersion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaTipoVersion 2
    TxtModo.SetFocus
End If
End Sub



 Private Sub BuscaTipoVersion(Opcion As Integer)
   Dim sField As String, iRows As Long
   Dim rstAux As ADODB.Recordset
    
    StrSQL = "Select Tip_Version, Descripcion From Es_TiposVersion WHERE "
    txtCodTipoVersion = Trim(txtCodTipoVersion)
    txtDesTipoVersion = Trim(txtDesTipoVersion)
    sField = txtCodTipoVersion
    Select Case Opcion
    Case 1: StrSQL = StrSQL & "Tip_Version like '%" & txtCodTipoVersion & "%'"
    Case 2: StrSQL = StrSQL & "Descripcion like '%" & txtDesTipoVersion & "%'"
    End Select
    
    txtCodTipoVersion = ""
    txtDesTipoVersion = ""
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = StrSQL
        .Cargar_Datos
        
        Codigo = ""
        Descripcion = ""

        iRows = .DGridLista.RowCount
        Set rstAux = .DGridLista.ADORecordset
        If rstAux.RecordCount > 1 Then .Show vbModal
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            txtCodTipoVersion = rstAux!Tip_Version
            txtDesTipoVersion = rstAux!Descripcion
            If iRows = 1 And Opcion = 1 And _
            sField = "" Then
                'txtCodOrigen.Enabled = False
                'txtDesOrigen.Enabled = False
            End If
            'SendKeys "{TAB}"
        End If
    End With
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
End Sub
   




 
Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
        If Trim(txtCodProveedor.Text) = "" Then
            Call Me.BUSCA_PROVEEDOR(3)
        Else
            txtCodProveedor.Text = Right("0000000000000" & Trim(txtCodProveedor.Text), 12)
            Call Me.BUSCA_PROVEEDOR(1)
        End If
        txtCodItemProv.SetFocus
End If
 
End Sub


Private Sub txtNombreProveedor_KeyPress(KeyAscii As Integer)
 
  If KeyAscii = 13 Then
        Call Me.BUSCA_PROVEEDOR(2)
             txtCodItemProv.SetFocus
  End If
 
End Sub


Public Sub BUSCA_PROVEEDOR(tipo As Integer)
    Select Case tipo
        Case 1:
                    'Strsql = "SELECT Des_Proveedor as 'Descripción' FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & Trim(Me.txtCod_Proveedor.Text) & "' AND Cod_Proveedor IN (SELECT DISTINCT(Cod_Proveedor) FROM cf_acumulado_proveedores where Flg_Status = 'P')"
                    StrSQL = "EXEC UP_SEL_PROVEEDORES_CF_ACUMULADOS '" & CInt(tipo) & "','" & Me.txtCodProveedor.Text & "','" & Me.txtNombreProveedor.Text & "'"
                    txtNombreProveedor.Text = Trim(DevuelveCampo(StrSQL, cCONNECT))
                    'txtCod_TemCli.SetFocus
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral3
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If tipo = 2 Then
                        'oTipo.sQuery = "SELECT Cod_Proveedor AS  'Código', Des_Proveedor as 'Descripción' FROM LG_PROVEEDOR WHERE Des_Proveedor LIKE  '%" & Trim(Me.txtDes_Proveedor.Text) & "%' AND Cod_Proveedor IN (SELECT DISTINCT(Cod_Proveedor) FROM cf_acumulado_proveedores where Flg_Status = 'P')"
                        oTipo.sQuery = "EXEC UP_SEL_PROVEEDORES_CF_ACUMULADOS '" & CInt(tipo) & "','" & Me.txtCodProveedor.Text & "','" & Me.txtNombreProveedor.Text & "'"
                    Else
                        'oTipo.sQuery = "SELECT Cod_Proveedor AS  'Código', Des_Proveedor as 'Descripción' FROM LG_PROVEEDOR WHERE Cod_Proveedor IN (SELECT DISTINCT(Cod_Proveedor) FROM cf_acumulado_proveedores where Flg_Status = 'P')"
                        oTipo.sQuery = "EXEC UP_SEL_PROVEEDORES_CF_ACUMULADOS '" & CInt(tipo) & "','" & Me.txtCodProveedor.Text & "','" & Me.txtNombreProveedor.Text & "'"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.DGridLista.Columns(2).Width = 5000
                    oTipo.Show 1
                    If Codigo <> "" Then
                        txtCodProveedor.Text = Trim(Codigo)
                        txtNombreProveedor.Text = Trim(Descripcion)
                        Codigo = "": Descripcion = ""
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
                    
    End Select
    FunctButt1.SetFocus
End Sub


Private Sub txtObservacionesProv_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
       cboIde_TallaX.SetFocus
End If
End Sub

Private Sub txtPrecio_GotFocus()
    SelectionText txtPrecio
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
       txtObservacionesProv.SetFocus
End If
End Sub

Private Sub txtPrecioComercial_GotFocus()
    SelectionText Me.txtPrecioComercial
End Sub

Private Sub txtPrecioComercial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTecnicaEstampado.SetFocus
    End If
End Sub

Private Sub txtTecnicaEstampado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        KeyAscii = 0
       If Frame3.Enabled = True Then
            Me.txtCodProveedor.SetFocus
            Else
            cboIde_TallaX.SetFocus
       End If
End If
End Sub

Private Sub txtUbicacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       cmdGrafico.SetFocus
End If
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            Me.txtPrecioComercial.SetFocus
End If
End Sub
Private Sub txtUniMedProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       txtPrecio.SetFocus
End If
End Sub
 

Sub ChangeName()
On Error Resume Next
Dim fso As New FileSystemObject
Dim sNombreFileOrigen
Dim iUbicacionNombreFileOrigen As Long
    
If RTrim(txtDirIcono) <> "" Then
    sNombreFileOrigen = Mid(fso.GetFileName(txtDirIcono.Text), 1, InStr(fso.GetFileName(txtDirIcono.Text), ".") - 1)
    iUbicacionNombreFileOrigen = InStr(txtDirIcono, sNombreFileOrigen)
    txtDirIcono.Text = Replace(txtDirIcono.Text, sNombreFileOrigen, txtcoditem)
End If
Exit Sub
Resume

ErrHandler:
ErrorHandler Err, "Move_Files"
    

End Sub

Sub Move_Files(strOption As String)
On Error Resume Next
    
If RTrim(txtDirIcono) <> "" Then
    FileCopy strImagenCambio, txtDirIcono.Text
    Guarda_Imagen
End If
Exit Sub
Resume

ErrHandler:
ErrorHandler Err, "Move_Files"
    
End Sub

Public Function Guarda_Imagen() As Boolean
    Dim Con As New ADODB.Connection
    Dim StrSQL As String
    On Error GoTo Guarda_ImagenErr

    Guarda_Imagen = False
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
       
        
    StrSQL = "EXEC SP_ActualizaDir_Icono '" & txtcoditem & "','" & txtDirIcono.Text & "'"

    
    Con.Execute StrSQL
        
    Con.CommitTrans
    'Move_Files (StrImagen_cambio)
    Guarda_Imagen = True
    Exit Function
Guarda_ImagenErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Guarda_Imagen"



End Function

