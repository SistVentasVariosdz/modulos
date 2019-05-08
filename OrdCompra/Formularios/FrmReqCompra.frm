VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmReqCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requerimientos por comprar"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Detalles de Requerimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   45
      TabIndex        =   11
      Top             =   1080
      Width           =   10035
      Begin SSDataWidgets_B.SSDBGrid DGridLista 
         Height          =   2925
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   9825
         _Version        =   196617
         DataMode        =   2
         AllowColumnShrinking=   0   'False
         SelectTypeRow   =   1
         BackColorOdd    =   13434879
         RowHeight       =   423
         ExtraHeight     =   106
         Columns.Count   =   22
         Columns(0).Width=   1138
         Columns(0).Caption=   "Flag"
         Columns(0).Name =   "Flag"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   2
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "Cod_Fabrica"
         Columns(1).Name =   "Cod_Fabrica"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Caption=   "Fábrica"
         Columns(2).Name =   "Fabrica"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         Columns(2).Style=   4
         Columns(3).Width=   1958
         Columns(3).Caption=   "O/P"
         Columns(3).Name =   "Cod_OrdPro"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(3).Locked=   -1  'True
         Columns(3).Style=   4
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
         Columns(4).Caption=   "Cod_Present"
         Columns(4).Name =   "Cod_Present"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "Cod_CompEst"
         Columns(5).Name =   "Cod_CompEst"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3200
         Columns(6).Visible=   0   'False
         Columns(6).Caption=   "Cod_Item"
         Columns(6).Name =   "Cod_Item"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   3122
         Columns(7).Caption=   "Item"
         Columns(7).Name =   "ITEM"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(7).Locked=   -1  'True
         Columns(7).Style=   4
         Columns(8).Width=   3200
         Columns(8).Caption=   "Cod. Prov. "
         Columns(8).Name =   "Cod_Prov"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   20
         Columns(9).Width=   3200
         Columns(9).Visible=   0   'False
         Columns(9).Caption=   "Cod_Comb"
         Columns(9).Name =   "Cod_Comb"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(10).Width=   3200
         Columns(10).Caption=   "Combinación"
         Columns(10).Name=   "COMBINACION"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(10).Locked=   -1  'True
         Columns(10).Style=   4
         Columns(11).Width=   3200
         Columns(11).Visible=   0   'False
         Columns(11).Caption=   "Cod_Color"
         Columns(11).Name=   "Cod_Color"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         Columns(12).Width=   2937
         Columns(12).Caption=   "Color"
         Columns(12).Name=   "COLOR"
         Columns(12).DataField=   "Column 12"
         Columns(12).DataType=   8
         Columns(12).FieldLen=   256
         Columns(12).Locked=   -1  'True
         Columns(12).Style=   4
         Columns(13).Width=   1217
         Columns(13).Caption=   "Talla"
         Columns(13).Name=   "Cod_Talla"
         Columns(13).DataField=   "Column 13"
         Columns(13).DataType=   8
         Columns(13).FieldLen=   256
         Columns(13).Locked=   -1  'True
         Columns(13).Style=   4
         Columns(14).Width=   3200
         Columns(14).Visible=   0   'False
         Columns(14).Caption=   "Cod_Destino"
         Columns(14).Name=   "Cod_Destino"
         Columns(14).DataField=   "Column 14"
         Columns(14).DataType=   8
         Columns(14).FieldLen=   256
         Columns(15).Width=   3200
         Columns(15).Caption=   "Destino"
         Columns(15).Name=   "DESTINO"
         Columns(15).DataField=   "Column 15"
         Columns(15).DataType=   8
         Columns(15).FieldLen=   256
         Columns(15).Locked=   -1  'True
         Columns(15).Style=   4
         Columns(16).Width=   3200
         Columns(16).Visible=   0   'False
         Columns(16).Caption=   "cod_estcli"
         Columns(16).Name=   "cod_estcli"
         Columns(16).DataField=   "Column 16"
         Columns(16).DataType=   8
         Columns(16).FieldLen=   256
         Columns(17).Width=   1614
         Columns(17).Caption=   "Uni. Med"
         Columns(17).Name=   "Cod_UniMed"
         Columns(17).DataField=   "Column 17"
         Columns(17).DataType=   8
         Columns(17).FieldLen=   256
         Columns(17).Locked=   -1  'True
         Columns(17).Style=   4
         Columns(18).Width=   2752
         Columns(18).Caption=   "Cant. por Comprar"
         Columns(18).Name=   "CANTXCOMPRAR"
         Columns(18).DataField=   "Column 18"
         Columns(18).DataType=   8
         Columns(18).FieldLen=   256
         Columns(19).Width=   3200
         Columns(19).Visible=   0   'False
         Columns(19).Caption=   "CANTIDAD"
         Columns(19).Name=   "CANTIDAD"
         Columns(19).DataField=   "Column 19"
         Columns(19).DataType=   8
         Columns(19).FieldLen=   256
         Columns(20).Width=   1958
         Columns(20).Caption=   "Medida"
         Columns(20).Name=   "MEDIDA"
         Columns(20).DataField=   "Column 20"
         Columns(20).DataType=   8
         Columns(20).FieldLen=   256
         Columns(21).Width=   3200
         Columns(21).Caption=   "Presentacion"
         Columns(21).Name=   "Presentacion"
         Columns(21).DataField=   "Column 21"
         Columns(21).DataType=   8
         Columns(21).FieldLen=   256
         Columns(21).Style=   4
         _ExtentX        =   17330
         _ExtentY        =   5159
         _StockProps     =   79
         Caption         =   "Resultados de la Busqueda"
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   480
      Left            =   5115
      TabIndex        =   9
      Top             =   4500
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opción de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   50
      TabIndex        =   1
      Top             =   15
      Width           =   10035
      Begin VB.TextBox txtDes_Grupo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1170
         TabIndex        =   13
         Top             =   255
         Width           =   2370
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Buscar"
         Height          =   405
         Left            =   5445
         TabIndex        =   10
         Top             =   555
         Width           =   1245
      End
      Begin VB.TextBox TxtFamilia 
         BackColor       =   &H00FFFFFF&
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
         Left            =   4695
         MaxLength       =   20
         TabIndex        =   8
         Top             =   240
         Width           =   2010
      End
      Begin VB.CommandButton cmdBuscaFamilia 
         Caption         =   "..."
         Height          =   330
         Left            =   6705
         TabIndex        =   7
         Tag             =   "..."
         Top             =   240
         Width           =   330
      End
      Begin VB.TextBox TxtOp 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1170
         MaxLength       =   20
         TabIndex        =   5
         Top             =   570
         Width           =   2010
      End
      Begin VB.CommandButton cmdBuscaColor 
         Caption         =   "..."
         Height          =   330
         Left            =   3210
         TabIndex        =   4
         Tag             =   "..."
         Top             =   570
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Familias"
         Height          =   195
         Index           =   1
         Left            =   3810
         TabIndex        =   6
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "O/P"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   3
         Top             =   675
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
         Height          =   195
         Left            =   225
         TabIndex        =   2
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   480
      Left            =   3135
      TabIndex        =   0
      Top             =   4500
      Width           =   1380
   End
End
Attribute VB_Name = "FrmReqCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String
Public Descripcion As String
Dim Strsql As String
Dim Rs_Lista As New ADODB.Recordset
Dim CadConn  As New ADODB.Connection

'Variables para la ejecucion del store
Public varCod_GrupoLog As String
'Dim vTotal As Double
Dim vCantAnt As Double


'Definicion de variables que seran pasadas por nuestro master
Public varSer_OrdComp, varCod_OrdComp, varSec_OrdComp As String
Public varTip_Presentacion, varCod_ClaOrdComp, varCod_Proveedor As String
Public varCod_Descuento As String
Public varCod_TipRequ As Integer
Public varPorc_IGV As Double
'Variables para la ejecucion del super mega store de generacion de requerimientos
Public varAccion As String
Dim Cadena As String

Sub Buscar(Grupo As String, Orden As String, Familia As String)
    On Error GoTo hand
    Dim Rs_Prov As New ADODB.Recordset
    Dim i As Integer
    
    Strsql = " UP_SEL_REQUEXCOMPRARLOG '" & Grupo & "','" & Orden & "','" & Familia & "'"
      
    Set Rs_Lista = Nothing
    Rs_Lista.ActiveConnection = cConnect
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    Rs_Lista.Open Strsql
    
    Set Rs_Prov = Rs_Lista.Clone
    
    'vTotal = 0
    If Rs_Lista.RecordCount > 0 Then
            
            Me.DGridLista.Redraw = False
            SSDBGridSetGrid Me.DGridLista
            ADODBToSSDBGridOC Rs_Prov, DGridLista
            DGridLista.ActiveRowStyleSet = "RowActive"
            DGridLista.SelectTypeRow = ssSelectionTypeMultiSelectRange
            DGridLista.Visible = True
            
            For i = 0 To DGridLista.Rows
                If i >= 6 Then DGridLista.Scroll 0, 1
                DGridLista.Columns(0).Value = 1
                'vTotal = vTotal + Me.DGridLista.Columns("CANTXCOMPRAR").Value
                If i = DGridLista.Rows - 1 Then Exit Sub
                DGridLista.Row = DGridLista.Row + 1
            Next
            'txtTotal.Text = vTotal
    Else
    
        Me.DGridLista.Redraw = False
        SSDBGridSetGrid Me.DGridLista
        ADODBToSSDBGridOC Rs_Prov, DGridLista
        DGridLista.ActiveRowStyleSet = "RowActive"
        DGridLista.SelectTypeRow = ssSelectionTypeMultiSelectRange
        DGridLista.Visible = True
    
        MsgBox "No se encontraron registros ", vbInformation, "Ordenes de Compra"
    End If
    
    Rs_Prov.Close
    Set Rs_Prov = Nothing
    Exit Sub
    
hand:
    ErrorHandler Err, "Buscar"
    Set Rs_Prov = Nothing
End Sub


Public Sub SoloNumeros(ByVal pTextbox As TextBox, _
                       ByRef pKeyAscii As Integer, _
                       Optional ByVal pConDecimales As Boolean, _
                       Optional ByVal pNumDecimales As Integer, _
                       Optional ByVal pNumEnteros As Integer)
   If pNumEnteros = 0 Then pNumEnteros = 10
   If pKeyAscii = 8 Then
      If pConDecimales And pTextbox.SelStart > 0 Then
         If Mid(pTextbox, pTextbox.SelStart, 1) = "." Then
            If Len(Mid(pTextbox, 1, pTextbox.SelStart - 1)) >= pNumEnteros And Len(Mid(pTextbox, pTextbox.SelStart + 1)) > 0 Then pKeyAscii = 0
         End If
      End If
      Exit Sub
   End If
   If pKeyAscii = 46 Then
      If pConDecimales Then
         If InStr(1, pTextbox, ".") > 0 Then
            pKeyAscii = 0
         Else
            If Len(Mid(pTextbox, pTextbox.SelStart + 1)) > pNumDecimales Then pKeyAscii = 0
            If pTextbox.SelStart > 0 Then If Len(Mid(pTextbox, 1, pTextbox.SelStart - 1)) >= pNumEnteros Then pKeyAscii = 0
         End If
      Else
         pKeyAscii = 0
      End If
   Else
      If Not (pKeyAscii >= 48 And pKeyAscii <= 57) Then pKeyAscii = 0
      If pKeyAscii = 39 Or pKeyAscii = 13 Then
         pKeyAscii = 0
      End If
      
      Dim iPos As Integer
      iPos = InStr(1, pTextbox, ".")
      If iPos > 0 And pConDecimales Then _
         If Len(Mid(pTextbox, iPos)) > pNumDecimales Then _
            If InStr(pTextbox.SelStart + 1, pTextbox, ".") = 0 Then pKeyAscii = 0
            
      If pTextbox.SelStart < iPos Or iPos = 0 Then
         If pNumEnteros > 0 Then
            If InStr(pTextbox.SelStart + 1, pTextbox, ".") > 0 Then
               If Len(Mid(pTextbox, 1, InStr(pTextbox.SelStart + 1, pTextbox, ".") - 1)) >= pNumEnteros Then pKeyAscii = 0
            Else
               If Len(pTextbox) >= pNumEnteros Then pKeyAscii = 0
            End If
         End If
      End If
   End If
End Sub

Private Sub cmdAceptar_Click()
Dim j As Integer
On Error GoTo ErrorAceptar:
    'Cadena = DevuelveCampo("SELECT COD_FAMITEM FROM LG_ITEM WHERE COD_ITEM='" & TxtFamilia & "'", cConnect)
    If TxtFamilia = "HI" Then
        For j = 0 To DGridLista.Rows - 1
            If DGridLista.Columns("flag").Value <> 0 Then
                If RTrim(DGridLista.Columns("cod_prov").Value) = "" Then
                    MsgBox "Debe ingresar el campo Cod. Prov.", vbInformation
                    DGridLista.Bookmark = 1
                    Exit Sub
                End If
            End If
            DGridLista.Bookmark = (j + 1)
        Next j
            Strsql = "SELECT ISNULL(MAX(Sec_OrdComp),0) FROM lg_ordcompitem WHERE Ser_OrdComp='" & Me.varSer_OrdComp & "' AND Cod_OrdComp='" & Me.varCod_OrdComp & "'"
            varSec_OrdComp = DevuelveCampo(Strsql, cConnect)
        
            'Llamando a este form obtendremos el tipo de insercion
            Load frmOpcionReq
            Set frmOpcionReq.frmmaster = Me
            frmOpcionReq.Show 1
            
            If varAccion = "" Then
                Exit Sub
            End If
            
            Set CadConn = Nothing
            CadConn.Open cConnect
        '    DGridLista.Row = 0
            DGridLista.Bookmark = 0
            For j = 0 To DGridLista.Rows - 1
                'DGridLista.Row = j
                'Grilla.Bookmark = j
                
                If Abs(DGridLista.Columns(0).Value) = 1 Then
                
                    Strsql = "exec UP_ACTUALIZA_REQ_OC  '" & _
                    varSer_OrdComp & "','" & _
                    varCod_OrdComp & "','" & _
                    varSec_OrdComp & "','" & _
                    varAccion & "','" & _
                    DGridLista.Columns("cod_fabrica").Text & "','" & _
                    DGridLista.Columns("Cod_OrdPro").Text & "','" & _
                    DGridLista.Columns("Cod_Present").Text & "','" & _
                    DGridLista.Columns("Cod_CompEst").Text & "','" & _
                    DGridLista.Columns("Cod_Item").Text & "','" & _
                    DGridLista.Columns("Cod_Comb").Text & "','" & _
                    DGridLista.Columns("Cod_Color").Text & "','" & _
                    DGridLista.Columns("Cod_Talla").Text & "','" & _
                    DGridLista.Columns("cod_destino").Text & "','" & _
                    DGridLista.Columns("cod_estcli").Text & "'," & _
                    DGridLista.Columns("CANTXCOMPRAR").Text & ",'" & _
                    DGridLista.Columns("Cod_Prov").Text & "'"
                
                    CadConn.Execute Strsql
                End If
                
        '        If j >= 6 Then
        '            DGridLista.Scroll 0, 1
        '            DGridLista.Row = 5
        '        End If
        '        DGridLista.Row = DGridLista.Row + 1
        
                DGridLista.Bookmark = (j + 1)
            Next
            Set CadConn = Nothing
            Unload Me
     Else
        Strsql = "SELECT ISNULL(MAX(Sec_OrdComp),0) FROM lg_ordcompitem WHERE Ser_OrdComp='" & Me.varSer_OrdComp & "' AND Cod_OrdComp='" & Me.varCod_OrdComp & "'"
            varSec_OrdComp = DevuelveCampo(Strsql, cConnect)
        
            'Llamando a este form obtendremos el tipo de insercion
            Load frmOpcionReq
            Set frmOpcionReq.frmmaster = Me
            frmOpcionReq.Show 1
            
            If varAccion = "" Then
                Exit Sub
            End If
            
            Set CadConn = Nothing
            CadConn.Open cConnect
            
            
        '    DGridLista.Row = 0
            DGridLista.Bookmark = 0
            For j = 0 To DGridLista.Rows - 1
                'DGridLista.Row = j
                'Grilla.Bookmark = j
                
                If Abs(DGridLista.Columns(0).Value) = 1 Then
                
                    Strsql = "exec UP_ACTUALIZA_REQ_OC  '" & _
                    varSer_OrdComp & "','" & _
                    varCod_OrdComp & "','" & _
                    varSec_OrdComp & "','" & _
                    varAccion & "','" & _
                    DGridLista.Columns("cod_fabrica").Text & "','" & _
                    DGridLista.Columns("Cod_OrdPro").Text & "','" & _
                    DGridLista.Columns("Cod_Present").Text & "','" & _
                    DGridLista.Columns("Cod_CompEst").Text & "','" & _
                    DGridLista.Columns("Cod_Item").Text & "','" & _
                    DGridLista.Columns("Cod_Comb").Text & "','" & _
                    DGridLista.Columns("Cod_Color").Text & "','" & _
                    DGridLista.Columns("Cod_Talla").Text & "','" & _
                    DGridLista.Columns("cod_destino").Text & "','" & _
                    DGridLista.Columns("cod_estcli").Text & "'," & _
                    DGridLista.Columns("CANTXCOMPRAR").Text & ",'" & _
                    DGridLista.Columns("Cod_Prov").Text & "'"
                
                    CadConn.Execute Strsql
                End If
                
        '        If j >= 6 Then
        '            DGridLista.Scroll 0, 1
        '            DGridLista.Row = 5
        '        End If
        '        DGridLista.Row = DGridLista.Row + 1
        
                DGridLista.Bookmark = (j + 1)
                
            Next
            Set CadConn = Nothing
            
            Unload Me
     End If
    Exit Sub
ErrorAceptar:
    Set CadConn = Nothing
    ErrorHandler Err, "Error Aceptar"
End Sub

Private Sub cmdBuscaColor_Click()
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = "Select cod_ordpro as Codigo, convert(char(10),fec_creacion,103) as [Fecha Creacion] from ES_ORDPRO where Cod_GrupoLog='" & varCod_GrupoLog & "' order by 1"
    frmBusqGeneral.CARGAR_DATOS
    frmBusqGeneral.Show 1
    
    If TxtOp <> Codigo Then
        TxtFamilia.Text = ""
    End If
    
    TxtOp = Codigo
    Codigo = ""

    

End Sub

Private Sub cmdBuscaFamilia_Click()

    '         Strsql = "SELECT  DISTINCT (SUBSTRING(RI.Cod_Item,1,2)) as 'Código' "
    'Strsql = Strsql & "FROM    ES_ORDPRO       OP, ES_ORDPROREQ_ITEMS RI "
    'Strsql = Strsql & "WHERE   OP.Cod_OrdPro = RI.Cod_OrdPro   AND "
    'Strsql = Strsql & "OP.Cod_GrupoLog = '" & varCod_GrupoLog & "' AND "
    'Strsql = Strsql & "RI.Cod_OrdPro   = '" & Trim(TxtOp.Text) & "'"
    
    Strsql = "EXEC UP_SEL_FAMORDCOMPLOG '" & varCod_GrupoLog & "','" & Trim(TxtOp.Text) & "'"
    
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = Strsql
    frmBusqGeneral.CARGAR_DATOS
    frmBusqGeneral.Show 1
    TxtFamilia.Text = Codigo
    Codigo = ""

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    Buscar varCod_GrupoLog, TxtOp, TxtFamilia
'    CalculaTotal
End Sub

'Private Sub DGridLista_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
'vCantAnt = Me.DGridLista.Columns("CANTXCOMPRAR").Value
'End Sub

Private Sub Form_Load()
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"
    'LlenaCombo varCod_GrupoLog, "select Des_Grupo + space(100) + Cod_GrupoLog from ES_GRUPOlog order by 1", cCONNECT
    'If varCod_GrupoLog.ListCount > 0 Then
    '    TxtOp.Enabled = True
    '    TxtOp = ""
    '    cmdBuscaColor.Enabled = True
    'Else
    '    TxtOp.Enabled = False
    '    TxtOp = ""
    '    cmdBuscaColor.Enabled = False
    'End If
End Sub

'Private Sub DGridLista_AfterUpdate(RtnDispErrMsg As Integer)
'    If Len(Trim(DGridLista.Columns(17).Text)) = 0 Then DGridLista.Columns(17).Text = "0"
'    CalculaTotal
'End Sub

Private Sub DGridLista_Change()
'    Dim ubic_cant As Integer
'    Select Case Index
'        Case 0:
'                ubic_cant = 7
'        Case 1:
'                ubic_cant = 10
'        Case 2:
'                ubic_cant = 12
'        Case 3:
'                ubic_cant = 16
'    End Select
    If Val(DGridLista.Columns(17).Text) > Val(DGridLista.Columns(18).Text) Then
        MsgBox "El valor comprado no puede ser mayor a la disponible. Sirvase verificar", vbInformation, "Ordenes de Compra"
        DGridLista.Columns(17).Text = DGridLista.Columns(18).Text
    End If
    
End Sub

'Private Sub DGridLista_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
'    Dim position As Integer
'    'If LastCol = 17 Then
'    position = DGridLista.Row
'    DGridLista.Row = CInt(LastRow) + 1
'    'DGridLista.Col = LastCol
'    'DGridLista.
'        If DGridLista.Columns(17).Text > DGridLista.Columns(18).Text Then
'            MsgBox "El valor comprado no puede ser mayor a la disponible. Sirvase verificar"
'            'DGridLista.Columns(17).Text = DGridLista.Columns(18).Text
'            'LastRow = 0
'        Else
'            DGridLista.Row = position
'        End If
'
'End Sub

Private Sub DGridLista_KeyPress(KeyAscii As Integer)
    If DGridLista.Col = 17 Then
       Select Case KeyAscii
            Case 48 To 57
                    If Len(Trim(DGridLista.Columns(17).Text)) >= 14 Then KeyAscii = 0: Exit Sub
                    KeyAscii = KeyAscii
            Case 46
                    If Len(Trim(DGridLista.Columns(17).Text)) >= 14 Then KeyAscii = 0: Exit Sub
                    If InStr(1, DGridLista.Columns(17).Text, ".") > 0 Then
                        KeyAscii = 0
                    Else
                        KeyAscii = KeyAscii
                    End If
            Case 8
                    KeyAscii = KeyAscii
            Case Else
                    KeyAscii = 0
        End Select
    End If
End Sub

Private Sub TxtOp_KeyPress(KeyAscii As Integer)
Dim temp As String
If KeyAscii = 13 Then
    TxtFamilia.Text = ""
    TxtOp = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(5," & IIf(Trim(TxtOp) = "", 0, TxtOp) & ")", cConnect))
    If DevuelveCampo("select count(*) from ES_ORDPRO where cod_ordpro ='" & TxtOp & "' and Cod_GrupoLog='" & Right(varCod_GrupoLog, 8) & "'", cConnect) <= 0 Then
            MsgBox "Codigo no existe", vbInformation
    End If
    
End If
End Sub

'Sub CalculaTotal()
'Dim i As Integer
'Dim vTotal As Double
'vTotal = 0
'DGridLista.Row = 0
'DGridLista.Bookmark = 0
'For i = 0 To DGridLista.Rows - 1
'    If DGridLista.Columns(0).Value = 1 Or DGridLista.Columns(0).Value = -1 Then
'        vTotal = vTotal + DGridLista.Columns("CANTXCOMPRAR").Value
'    End If
'        If i >= 6 Then
'            DGridLista.Scroll 0, 1
'            DGridLista.Row = 5
'        End If
'
'    DGridLista.Row = DGridLista.Row + 1
'Next
'txtTotal.Text = vTotal
'End Sub

