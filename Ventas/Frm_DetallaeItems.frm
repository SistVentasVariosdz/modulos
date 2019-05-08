VERSION 5.00
Begin VB.Form Frm_DetallaeItems 
   Caption         =   "Detalle Items"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3862
         TabIndex        =   6
         Tag             =   "&Cancel"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2617
         TabIndex        =   5
         Tag             =   "&OK"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtCod_Producto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   25
         TabIndex        =   2
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1680
         MaxLength       =   53
         TabIndex        =   1
         Top             =   645
         Width           =   5505
      End
      Begin VB.Label Label12 
         Caption         =   "Codigo de Producto :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Descripcion :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   645
         Width           =   975
      End
   End
End
Attribute VB_Name = "Frm_DetallaeItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigo As String, sunidad As String
Public Descripcion As String, TipoAdd As String, Tipoa As String, Tipob As String
Dim strSQL As String

Private Sub cmdaceptar_Click()
'    If Trim(txtCod_Producto.Text) = "" Or txtDescripcion.Text = "" Then
'        MsgBox "El campo Codigo o la descripcion no pueden estar en blanco, verificar. "
'        Exit Sub
'    Else
        If Trim(txtCod_Producto) = "" Then
            If MsgBox("Desea crear un nuevo Item?", vbQuestion + vbYesNo, Mantenimiento) = vbYes Then
                Load Frm_ManteItem
                Frm_ManteItem.Show vbModal
            End If
        End If
        frmAdicionaDetalleDocum.txtCod_Producto.Text = txtCod_Producto.Text
        frmAdicionaDetalleDocum.txtDescripcion.Text = txtDescripcion.Text
        frmAdicionaDetalleDocum.txtUnida_Medida.Text = sunidad
        frmAdicionaDetalleDocum.txtCod_Producto.Enabled = False
        frmAdicionaDetalleDocum.txtDescripcion.Enabled = False
        Unload Me
'    End If

End Sub

Private Sub cmdcancelar_Click()
If Trim(txtDescripcion.Text) = "" Or Trim(txtCod_Producto.Text) = "" Then
     If MsgBox("Desea crear un nuevo Item?", vbQuestion + vbYesNo, Mantenimiento) = vbYes Then
        Load Frm_ManteItem
        Frm_ManteItem.Show vbModal
    Else
        Unload Me
    End If
Else
    Unload Me
End If
End Sub

Private Sub txtCod_Producto_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        txtDescripcion.SetFocus

End If
End Sub


Public Sub BuscaItem()

    Dim oTipo As New frmBusqGeneral3
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    Set oTipo.oParent = Me


    oTipo.sQuery = "VENTAS_BUSCA_ITEMS_DIVERSOS '" & Trim(txtDescripcion.Text) & "'"


    oTipo.Caption = "Buscar Items"
    oTipo.Cargar_Datos

    oTipo.gexLista.Columns("COD_ITEM").Width = 1400
    oTipo.gexLista.Columns("DES_ITEM").Width = 5000

    If oTipo.gexLista.RowCount > 1 Then
        oTipo.Show vbModal
    Else
        codigo = Trim(oTipo.gexLista.Value(oTipo.gexLista.Columns("COD_ITEM").Index))
        Descripcion = Trim(oTipo.gexLista.Value(oTipo.gexLista.Columns("DES_ITEM").Index))
        TipoAdd = Trim(oTipo.gexLista.Value(oTipo.gexLista.Columns("COD_UNIMED").Index))
    End If

    If Trim(codigo) <> "" Then
        txtCod_Producto.Text = ""
        txtDescripcion = ""
        sunidad = ""
        txtCod_Producto.Text = Trim(codigo)
        txtDescripcion.Text = Trim(Descripcion)
        sunidad = Trim(TipoAdd)

        codigo = "": Descripcion = "": TipoAdd = ""
        'vCod_costo = DevuelveCampo("select Cod_CenCost from Rh_Centros_Costos where Cod_CenCost ='" & Trim(Txt_centro.Text) & "' order by Cod_CenCost", cCONNECT)
        cmdAceptar.SetFocus
    End If
    Unload oTipo
    Set oTipo = Nothing
    Set rs = Nothing
End Sub

Private Sub txtDesCRIPCION_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BuscaItem
End If

End Sub
