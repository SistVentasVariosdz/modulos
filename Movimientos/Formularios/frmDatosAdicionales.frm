VERSION 5.00
Begin VB.Form frmDatosAdicionales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emision de la Guia"
   ClientHeight    =   5280
   ClientLeft      =   3435
   ClientTop       =   3195
   ClientWidth     =   7200
   Icon            =   "frmDatosAdicionales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7200
   Begin VB.Frame Frame3 
      Caption         =   "Datos de la Guia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   6975
      Begin VB.CommandButton Command1 
         Caption         =   "&Añadir"
         Height          =   320
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TxtNom_Transportista 
         Height          =   315
         Left            =   2400
         TabIndex        =   24
         Top             =   1560
         Width           =   3735
      End
      Begin VB.CommandButton CmdTransportista 
         Caption         =   "..."
         Height          =   315
         Left            =   2040
         TabIndex        =   23
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox TxtSec_Transportista 
         Height          =   315
         Left            =   1320
         TabIndex        =   22
         Top             =   1560
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Caption         =   "Guia a Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   4215
         Begin VB.TextBox TxtSerie 
            Height          =   315
            Left            =   720
            TabIndex        =   4
            Top             =   260
            Width           =   615
         End
         Begin VB.TextBox TxtNumero 
            Height          =   315
            Left            =   2280
            TabIndex        =   5
            Top             =   260
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Serie:"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   320
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Número:"
            Height          =   255
            Left            =   1560
            TabIndex        =   17
            Top             =   320
            Width           =   735
         End
      End
      Begin VB.TextBox TxtCod_Motivo 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton CmdTraslado 
         Caption         =   "..."
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox TxtDes_Motivo 
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label8 
         Caption         =   "Transportista:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Mot.Traslado:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3720
      Picture         =   "frmDatosAdicionales.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de Transportista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6975
      Begin VB.TextBox TxtPlaca 
         Height          =   315
         Left            =   1335
         TabIndex        =   3
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox TxtRuc 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   1305
         Width           =   1815
      End
      Begin VB.TextBox TxtDomicilio 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox TxtTransportista 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label7 
         Caption         =   "Nº de Placa"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "RUC:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Domicilio:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmDatosAdicionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String
Public Descripcion As String

Public NumMovStk As String
Public CodAlmacen As String
Public CodProveedor As String
Public CodCenCost As String
Public Ser_OrdComp As String
Public Cod_OrdComp As String

Public varMoviStk_Guia As Boolean
Public varReferencia As String
Public varPedido As String
Public varOpt As String
Public Paso As String
Public sTrans As String
Public vNumConosHilos As Integer
Public sTipImpresion As String
Public vRespuesta As String, sDoc As String
Public sguia As String


Dim oPrint As LibraryVB.clsPrintFile
'Dim oPrint As LibraryVB.clsPrintFile
Dim StrSql As String
Dim iLin As Integer

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
On Error GoTo FinImp
Dim Rs As ADODB.Recordset
Dim oTipo As New frmTipImpresion
If ValidaDatos = False Then Exit Sub
    
    Set Rs = New ADODB.Recordset
    Rs.ActiveConnection = cConnect
    Rs.CursorType = adOpenStatic
    Rs.CursorLocation = adUseClient
    
    If varMoviStk_Guia = False Then
        Rs.Open "EXEC UP_MAN_IMPRESION_GUIA '" & CodAlmacen & "','" & NumMovStk & "','" & Trim(TxtTransportista.Text) & "','" & Trim(TxtDomicilio.Text) & "','" & Trim(TxtRuc.Text) & "','" & Trim(TxtSerie.Text) & "','" & TxtNumero.Text & "','" & Trim(TxtPlaca.Text) & "','" & Trim(Me.TxtSec_Transportista.Text) & "', '" & sDoc & "','" & TxtCod_Motivo.Text & "'"
    Else
        Rs.Open "EXEC UP_MAN_IMPRESION_GUIA_NUEVO '" & CodAlmacen & "','" & NumMovStk & "','" & Trim(TxtTransportista.Text) & "','" & Trim(TxtDomicilio.Text) & "','" & Trim(TxtRuc.Text) & "','" & Trim(TxtSerie.Text) & "','" & TxtNumero.Text & "','" & Trim(TxtPlaca.Text) & "','" & Trim(Me.TxtCod_Motivo.Text) & "','" & Trim(Me.TxtSec_Transportista.Text) & "'"
        'preguntamos que tipo de impresion es
        Set oTipo.oParent = Me
        oTipo.Show 1
        Set oTipo = Nothing
    End If
    sguia = DevuelveCampo("select Flg_Guia from lg_almacen where cod_almacen='" & CodAlmacen & "'", cConnect)
    
    StrSql = "SELECT Tipo_Impresion_Guia FROM TG_CONTROL"
    Select Case DevuelveCampo(StrSql, cConnect)
    
    
    
    Case "2"
        If sTipImpresion = "RECOJO" Then
            IMPRIMIR_REPORTE
            sTipImpresion = ""
        Else
            IMPRIMIR_REPORTE2
        End If
    Case "1"
        Reporte
    Case "3"
        If Not IMPRIMIR_REPORTE2_SUMIT Then Exit Sub
    End Select
    Set Rs = Nothing
    Unload Me
Exit Sub
FinImp:
    MsgBox err.Description, vbCritical + vbOKOnly, "Impresion de Guia"
End Sub


Private Sub CmdTransportista_Click()
    Call Me.BUSCA_TRANSPORTISTA(3)
End Sub

Private Sub CmdTraslado_Click()
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = "select cod_mottra as codigo, des_mottra as Descripcion from tg_mottra"
    frmBusqGeneral.Cargar_Datos
    frmBusqGeneral.Show 1
    TxtDes_Motivo = Descripcion
    TxtCod_Motivo = Codigo
    If Codigo <> "" Then
        If Me.TxtSec_Transportista.Enabled Then AVANZA 13
        Else: cmdImprimir.SetFocus
    End If
    Codigo = ""
    Descripcion = ""

End Sub

Private Sub Command1_Click()
    frmManTransportistas.Show 1
End Sub

Private Sub Form_Load()
    varMoviStk_Guia = False
End Sub

Private Sub TxtCod_Motivo_KeyPress(KeyAscii As Integer)
Dim StrSql As String
If KeyAscii = 13 Then
    If TxtCod_Motivo <> "" Then
        If ExisteCampo("cod_mottra", "tg_mottra", TxtCod_Motivo, cConnect) Then
            TxtDes_Motivo = DevuelveCampo("select des_mottra from tg_mottra where cod_mottra='" & TxtCod_Motivo & "'", cConnect)
            If Me.TxtSec_Transportista.Enabled Then
                AVANZA 13
            Else
                cmdImprimir.SetFocus
            End If
        Else
            MsgBox "El codigo no existe", vbInformation, "Guia de Remisión"
            TxtCod_Motivo.Text = ""
        End If
    Else
        CmdTraslado_Click
    End If
End If
End Sub

Private Sub TxtDomicilio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtRuc.SetFocus
End Sub

Private Sub TxtNom_Transportista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Me.BUSCA_TRANSPORTISTA(2)
    End If
End Sub

Private Sub TxtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
        'TxtCod_Motivo.SetFocus
    End If
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(TxtNumero, KeyAscii, False, 0, 8)
End Sub

Private Sub TxtNumero_LostFocus()
    TxtNumero = Format(TxtNumero, "00000000")
End Sub

Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtSerie.SetFocus
End Sub

Private Sub TxtRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtPlaca.SetFocus
    Else
        Call SoloNumeros(TxtRuc, KeyAscii, False, 0, 11)
    End If
End Sub

Private Sub TxtSec_Transportista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.TxtSec_Transportista.Text) = "" Then
            Call Me.BUSCA_TRANSPORTISTA(3)
        Else
            Call Me.BUSCA_TRANSPORTISTA(1)
        End If
    End If
End Sub

'Private Sub TxtSerie_Change()
'    If varMoviStk_Guia = False Then
'        If Len(TxtSerie) = 3 Then TxtNumero.SetFocus
'    End If
'End Sub



Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    Call SoloNumeros(TxtSerie, KeyAscii, False, 0, 3)
End Sub

Function ValidaDatos() As Boolean
    If Trim(TxtSerie.Text) = "" Then
        MsgBox "Ingrese la serie de la guia", vbInformation, "Guia de Remision"
        TxtSerie.SetFocus
        ValidaDatos = False
        Exit Function
    End If

    If Trim(TxtNumero.Text) = "" Then
        MsgBox "Ingrese el número de la guia", vbInformation, "Guia de Remision"
        TxtNumero.SetFocus
        ValidaDatos = False
        Exit Function
    End If

    If Trim(TxtCod_Motivo.Text) = "" Then
        MsgBox "Ingrese el motivo de la guia", vbInformation, "Guia de Remision"
        TxtCod_Motivo.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If UCase(sDoc) = "Guia" Then
        If Trim(Me.TxtSec_Transportista.Text) = "" Then
            MsgBox "Ingrese el Transportista", vbInformation, "Guia de Remision"
            TxtSec_Transportista.SetFocus
            ValidaDatos = False
            Exit Function
        Else
            StrSql = "SELECT REG_TRANSPORTISTA FROM LG_TRANSPORTISTA WHERE SECUENCIA='" & TxtSec_Transportista & "'"
            If Trim(DevuelveCampo(StrSql, cConnect)) = "" Then
                MsgBox "Vehiculo no tiene registrado Transportista Asociado", vbExclamation + vbOKOnly, "Datos Adicionales"
                ValidaDatos = False
                Exit Function
            End If
        End If
    End If
ValidaDatos = True
End Function

Sub Reporte()
On Error GoTo ErrorImpresion
Dim oo As Object
Dim Ruta As String
        Ruta = vRuta & "\GRemision.xlt"
        'Ruta = App.Path & "\GRemision.xlt"
        'Ruta = "C:\Archivos de Programa\Gestion de pedidos\GRemision.xlt"

    Set oo = CreateObject("excel.application")
    oo.Workbooks.Open Ruta
    'oo.Visible = True
    oo.DisplayAlerts = False
    
    If varMoviStk_Guia = False Then
        oo.Run "reporte", cConnect, vemp1, CodProveedor, CodAlmacen, NumMovStk, TxtTransportista.Text, TxtDomicilio, Trim(TxtRuc.Text), Trim(TxtPlaca.Text), TxtSerie & "-" & TxtNumero, TxtCod_Motivo, Ser_OrdComp, Cod_OrdComp, 0
    Else
        oo.Run "reporte", cConnect, vemp1, Me.varOpt, CodAlmacen, NumMovStk, TxtTransportista.Text, TxtDomicilio, Trim(TxtRuc.Text), Trim(TxtPlaca.Text), TxtSerie & "-" & TxtNumero, TxtCod_Motivo, Ser_OrdComp, Cod_OrdComp, 1
    End If
    
    
    oo.Workbooks.Close
    Exit Sub
ErrorImpresion:
    Set oo = Nothing
    MsgBox "Hubo error en la impresion del Reporte de Guia de Remisión " & err.Description, vbCritical, "Impresion"

End Sub

Sub IMPRIMIR_REPORTE()
Dim RsPro As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim oPrint As LibraryVB.clsPrintFile
Dim i As Integer
Dim strCadena As String
Dim varNroReg As Integer
Dim NroReg As Integer
Dim NumTotReg As Integer
Dim varDescripcion As String
Dim VarCociente As Integer
Dim sOrden As String
Dim sTip_Item As String
Dim sDireccion3 As String

sTip_Item = DevuelveCampo("SELECT TIP_ITEM FROM LG_ALMACEN WHERE COD_ALMACEN ='" & CodAlmacen & "'", cConnect)
sDireccion3 = DevuelveCampo("SELECT DIRECCION_DE_PARTIDA_EMPRESA3 FROM TG_CONTROL ", cConnect)

iLin = 0
Set oPrint = New clsPrintFile

    Open "c:\GUIA.txt" For Output As #1
    
    Plin Chr(15) & "   "

    Plin sDoc
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
   ' Plin "     "
    
    
Set RsPro = New ADODB.Recordset
RsPro.ActiveConnection = cConnect
RsPro.CursorLocation = adUseClient
RsPro.CursorType = adOpenStatic


If varMoviStk_Guia = False Then
    StrSql = "SELECT Des_Proveedor AS 'NOMBRE', Dom_Proveedor AS 'DIRECCION', Num_Ruc AS 'RUC' FROM LG_PROVEEDOR WHERE Cod_Proveedor ='" & CodProveedor & "'"
Else
    Select Case Mid(Me.varOpt, 1, 1)
        Case "0":     'TipoProveedor
                        StrSql = "SELECT Des_Proveedor AS 'NOMBRE', Dom_Proveedor AS 'DIRECCION', Num_Ruc AS 'RUC' FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) & "'"
        Case "1":     'TipoCliente
                        StrSql = "SELECT Nom_Cliente AS 'NOMBRE', Direccion AS 'DIRECCION', Num_Ruc AS 'RUC' FROM TG_CLIENTE WHERE Cod_Cliente = '" & Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) & "'"
        Case "2":     'Destinatario
                        StrSql = "SELECT Destinatario as 'NOMBRE', Dom_Destinatario as 'DIRECCION', Ruc_Destinatario as 'RUC' FROM LG_MOVISTK_GUI WHERE COD_ALMACEN = '" & CodAlmacen & "' and Num_MovStk = '" & NumMovStk & "'"
    End Select
   
End If

strCadena = Space(100) & Trim(Str(Day(Date))) & Space(8)
strCadena = strCadena & DevuelveMes(IIf(Month(Date) < 10, "0" & Trim(Str(Month(Date))), Trim(Str(Month(Date)))), 1)
strCadena = strCadena & Space(5) & Trim(Str(Year(Date)))
Plin strCadena
Plin "     "
'Aqui imprimimos los datos obtenidos
    RsPro.Open StrSql
    If Not RsPro.EOF Then
        strCadena = Space(18) & Trim(RsPro.Fields("NOMBRE").Value)
        strCadena = strCadena & Space(95 - Len(strCadena)) & Me.varReferencia
        Plin strCadena
        
        strCadena = Space(15) & Trim(RsPro.Fields("RUC").Value) & Space(90) & IIf(Trim(Ser_OrdComp) <> "", Ser_OrdComp & "-" & Cod_OrdComp, Cod_OrdComp)
        Plin strCadena
        
        If varMoviStk_Guia = False Or Mid(Me.varOpt, 1, 1) = "0" Then
           If varMoviStk_Guia = False Then
              If CodProveedor = "000000000054" And sTip_Item = "T" Then
                 strCadena = Space(22) & Trim(sDireccion3)
              Else
                 strCadena = Space(22) & Trim(RsPro.Fields("DIRECCION").Value)
              End If
           Else
              If Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) = "000000000054" And sTip_Item = "T" Then
                 strCadena = Space(22) & Trim(sDireccion3)
              Else
                 strCadena = Space(22) & Trim(RsPro.Fields("DIRECCION").Value)
              End If
           End If
        Else
           strCadena = Space(22) & Trim(RsPro.Fields("DIRECCION").Value)
        End If
        
        strCadena = strCadena & Space(95 - Len(strCadena))
        
        StrSql = "UP_EXTRAE_PEDIDO_PARA_GUIAREM '" & CodAlmacen & "' , '" & NumMovStk & "'"
        sOrden = RTrim(DevuelveCampo(StrSql, cConnect))
        
        If RTrim(sOrden) <> "" Then
            Me.varPedido = sOrden
        End If
        
        strCadena = strCadena & Space(5) & Me.varPedido
        
        Plin strCadena
        Plin "     "
    Else
        strCadena = ""
        strCadena = strCadena & Space(95 - Len(strCadena))
        
        StrSql = "UP_EXTRAE_PEDIDO_PARA_GUIAREM '" & CodAlmacen & "' , '" & NumMovStk & "'"
        sOrden = RTrim(DevuelveCampo(StrSql, cConnect))
        
        If RTrim(sOrden) <> "" Then
            Me.varPedido = sOrden
        End If
        
        strCadena = strCadena & Space(5) & Me.varPedido
        Plin strCadena
    End If

Plin "     "


strCadena = Space(100) & Mid(Trim(TxtTransportista.Text), 1, 30)
Plin strCadena
strCadena = Space(95) & Mid(Trim(TxtRuc.Text), 1, 40) & Space(20) & Trim(TxtPlaca.Text)
Plin strCadena
strCadena = Space(98) & Mid(Trim(TxtDomicilio.Text), 1, 30)
Plin strCadena

If varMoviStk_Guia = False Then

    RsPro.Close
    RsPro.Open "select ser_ordcomp,cod_ordcomp,des_protex,des_claordcomp from lg_ordComp a,tx_procesos b,lg_claordcomp c where a.cod_protex *=b.cod_protex and a.cod_claordcomp=c.cod_claordcomp and ser_ordcomp='" & Ser_OrdComp & "' and cod_ordcomp='" & Cod_OrdComp & "'"
    Plin "     "
    
    If RsPro.RecordCount Then
        Plin "     "
        Plin "     "
        strCadena = Space(5) & "O/C: " & Ser_OrdComp + "-" + Cod_OrdComp & " - " & RsPro("des_claordcomp") & "   Proceso: " & RsPro("des_protex") & Space(10) & "Almacen: " & DevuelveCampo("select rtrim(nom_almacen) from lg_almacen where cod_almacen='" & CodAlmacen & "'", cConnect) & "   Mov.#: " & NumMovStk
        Plin strCadena
        If DevuelveCampo("select tip_fabrica from tg_control", cConnect) = "2" Then
             strCadena = "      N/P : " & DevuelveCampo("EXEC SM_BUSCA_OPS_OC '" & Ser_OrdComp & "','" & Cod_OrdComp & "',''", cConnect)
             Plin strCadena
        End If
    Else
        Plin "     "
        Plin "     "
        strCadena = Space(5) & "Almacen: " & DevuelveCampo("select rtrim(nom_almacen) from lg_almacen where cod_almacen='" & CodAlmacen & "'", cConnect) & "   Mov.#: " & NumMovStk
        Plin strCadena
    End If


Else
    Plin "     "
    Plin "     "
    StrSql = "SELECT isnull(Linea1,'') as 'Linea1', isnull(Linea2,'') as 'Linea2' FROM LG_MOVISTK_GUI WHERE COD_ALMACEN = '" & CodAlmacen & "' and Num_MovStk = '" & NumMovStk & "'"

    RsPro.Close
    RsPro.Open StrSql

    strCadena = Space(3) & Trim(RsPro("Linea1").Value) & Space(20) & "Almacen: " & DevuelveCampo("SELECT RTRIM(nom_almacen) FROM LG_ALMACEN_GUIAS WHERE cod_almacen='" & CodAlmacen & "'", cConnect) & "   Mov.#: " & NumMovStk
    Plin strCadena
    strCadena = Space(3) & Trim(RsPro("Linea2").Value)
    Plin strCadena
    
End If

Plin "     "

Set RsPro = Nothing

Set Rs = New ADODB.Recordset
Rs.ActiveConnection = cConnect
Rs.CursorLocation = adUseClient
Rs.CursorType = adOpenStatic

If varMoviStk_Guia = False Then
    Rs.Open "EXEC UP_SEL_GUIA_REMISION '" & CodAlmacen & "','" & NumMovStk & "'"
Else
    Rs.Open "EXEC UP_SEL_GUIA_REMISION '" & CodAlmacen & "','" & NumMovStk & "','*'"
End If


If Rs.RecordCount Then

varNroReg = 1
NroReg = 1
NumTotReg = Rs.RecordCount

'Imprimimos titulos del detalle
strCadena = Space(2) & "CANTIDAD" & Space(2) & "UNIDAD" & Space(20) & "DESCRIPCION"
Plin strCadena
strCadena = "-----------------------------------------------------------------------------------------------------------------------------------"
Plin strCadena
    For i = 1 To NumTotReg
        'VarDescripcion = Trim(Rs.Fields("Descripcion").Value)
        'VarCociente = 1 + (Len(VarDescripcion) / 90)
        
        strCadena = Space(5) & Rs.Fields("Cantidad").Value
        strCadena = strCadena & Space(14 - Len(strCadena)) & Rs.Fields("Uni Med").Value & Space(5) & Trim(Rs.Fields("Descripcion").Value)
        
        Plin strCadena
        If Trim(Rs.Fields("Descripcion").Value) <> "" Then
            Plin Space(38) & Trim(Rs.Fields("Descripcion").Value)
        End If
    
        'For varNroReg = 2 To VarCociente
'             strCadena = Trim(Rs.Fields("Descripcion").Value)
'             Plin strCadena
        'Next
       
        NroReg = NroReg + VarCociente - 1
        Rs.MoveNext
        NroReg = NroReg + 1
        'VarCociente = 0
    Next
    
    If varMoviStk_Guia = False Then
        StrSql = Trim(DevuelveCampo("select observaciones from Lg_MoviStk where cod_almacen='" & CodAlmacen & "' and num_movstk='" & NumMovStk & "'", cConnect))
        If Trim(StrSql) <> "" Then
            strCadena = "==================================================================================================================================="
            Plin strCadena
            strCadena = Space(3) & "Observacion:"
            Plin strCadena
            strCadena = Space(15) & StrSql
            Plin strCadena
        End If
    End If
End If


    Plin "                 "
    Plin "                 "
    
    Close #1
    oPrint.SendPrint "c:\GUIA.txt"
    Set oPrint = Nothing

End Sub

Sub IMPRIMIR_REPORTE2()
Dim StrSql As String
iLin = 0
Set oPrint = New clsPrintFile

    'Open "LPT1:" For Output As #1
    Open "C:\Guia.txt" For Output As #1
    
    Plin Chr(15) & "   "
    
    Plin sDoc
    Plin "     "
    Plin "     "
    Plin "     "
    If sguia = "S" Then
        Plin "     "
    End If

    'IMPRIME_CABECERA
    ' IMPRIME_DETALLE
    IMPRIME_CABECERA_02
    IMPRIME_DETALLE_02
    
    
    Close #1
    oPrint.SendPrint "c:\GUIA.txt"
    Set oPrint = Nothing

End Sub

Sub IMPRIME_CABECERA()
Dim RsPro As ADODB.Recordset
Dim strCadena As String
Dim varDescripcion As String
Dim sOrden As String
Dim sTip_Item As String
Dim sDireccion3 As String
Dim strSQLFecha  As String
Dim sFecha As String

sTip_Item = DevuelveCampo("SELECT TIP_ITEM FROM LG_ALMACEN WHERE COD_ALMACEN ='" & CodAlmacen & "'", cConnect)
sDireccion3 = DevuelveCampo("SELECT DIRECCION_DE_PARTIDA_EMPRESA3 FROM TG_CONTROL ", cConnect)

Set RsPro = New ADODB.Recordset
RsPro.ActiveConnection = cConnect
RsPro.CursorLocation = adUseClient
RsPro.CursorType = adOpenStatic


If varMoviStk_Guia = False Then
    If Trim(CodProveedor) = "" Then
        StrSql = "SELECT Des_CenCost AS 'NOMBRE', ' ' AS 'DIRECCION', ' ' AS 'RUC' FROM tg_cencosto WHERE Cod_CenCost ='" & CodCenCost & "'"
    Else
        StrSql = "SELECT Des_Proveedor AS 'NOMBRE', Dom_Proveedor AS 'DIRECCION', Num_Ruc AS 'RUC' FROM LG_PROVEEDOR WHERE Cod_Proveedor ='" & CodProveedor & "'"
    End If
Else
    Select Case Mid(Me.varOpt, 1, 1)
        Case "0":     'TipoProveedor
                        StrSql = "SELECT Des_Proveedor AS 'NOMBRE', Dom_Proveedor AS 'DIRECCION', Num_Ruc AS 'RUC' FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) & "'"
        Case "1":     'TipoCliente
                        StrSql = "SELECT Nom_Cliente AS 'NOMBRE', Direccion AS 'DIRECCION', Num_Ruc AS 'RUC' FROM TG_CLIENTE WHERE Cod_Cliente = '" & Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) & "'"
        Case "2":     'Destinatario
                        StrSql = "SELECT Destinatario as 'NOMBRE', Dom_Destinatario as 'DIRECCION', Ruc_Destinatario as 'RUC' FROM LG_MOVISTK_GUI WHERE COD_ALMACEN = '" & CodAlmacen & "' and Num_MovStk = '" & NumMovStk & "'"
    End Select
   
End If
    


'Plin "     "
'Plin "     "
'strCadena = Space(105) & Trim(TxtSerie.Text) & "-" & Trim(TxtNumero.Text)
'Plin strCadena
'Plin "     "
'***************************************************************************************************************************
'==> CAMBIOS EN LA IMPRESION DE LA GUIA [16/04/2008]
'***************************************************************************************************************************
Plin "     ": Plin "     ": Plin "     ": Plin "     ": Plin "     ": Plin "     "
strCadena = Space(105) & Trim(TxtSerie.Text) & "-" & Trim(TxtNumero.Text)
Plin strCadena

strCadena = " Motivo de Traslado: " & RPad(TxtDes_Motivo.Text, 50, " ")
strCadena = strCadena & Space(25) & Trim(Str(Day(Date))) & Space(8)
strCadena = strCadena & DevuelveMes(IIf(Month(Date) < 10, "0" & Trim(Str(Month(Date))), Trim(Str(Month(Date)))), 1)
strCadena = strCadena & Space(8) & Trim(Str(Year(Date)))
Plin strCadena

Plin "     "


If varMoviStk_Guia = False Then
    strCadena = Space(52) & DevuelveCampo("select isnull(Nom_Almacen_Guia,'') From Lg_Almacen Where cod_almacen='" & CodAlmacen & "'", cConnect)
Else
    strCadena = Space(52) & DevuelveCampo("select isnull(Nom_Almacen_Guia,'') From Lg_Almacen_guias Where cod_almacen='" & CodAlmacen & "'", cConnect)
End If
Plin strCadena


'strCadena = Space(100) & Trim(Str(Day(Date))) & Space(8)
'strCadena = strCadena & DevuelveMes(IIf(Month(Date) < 10, "0" & Trim(Str(Month(Date))), Trim(Str(Month(Date)))), 1)
'strCadena = strCadena & Space(8) & Trim(Str(Year(Date)))
'
'Plin strCadena
'Plin "     "
'Aqui imprimimos los datos obtenidos
    RsPro.Open StrSql
    If Not RsPro.EOF Then
        strCadena = Space(17) & Trim(RsPro.Fields("NOMBRE").Value)
        strCadena = strCadena & Space(95 - Len(strCadena)) & Me.varReferencia
        Plin strCadena
        


        StrSql = "UP_EXTRAE_PEDIDO_PARA_GUIAREM '" & CodAlmacen & "' , '" & NumMovStk & "'"
        sOrden = RTrim(DevuelveCampo(StrSql, cConnect))

        If RTrim(sOrden) <> "" Then
            Me.varPedido = sOrden
        End If


        strCadena = Space(17) & Trim(RsPro.Fields("RUC").Value) & Space(95 - Len(varPedido)) & varPedido
        Plin strCadena
        
        If varMoviStk_Guia = False Or Mid(Me.varOpt, 1, 1) = "0" Then
           If varMoviStk_Guia = False Then
              If CodProveedor = "000000000054" And sTip_Item = "T" Then
                 strCadena = Space(22) & Trim(sDireccion3)
              Else
                 strCadena = Space(22) & Trim(RsPro.Fields("DIRECCION").Value)
              End If
           Else
              If Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) = "000000000054" And sTip_Item = "T" Then
                 strCadena = Space(22) & Trim(sDireccion3)
              Else
                 strCadena = Space(22) & Trim(RsPro.Fields("DIRECCION").Value)
              End If
           End If
        Else
            strCadena = Space(22) & Trim(RsPro.Fields("DIRECCION").Value)
        End If
        
        strCadena = strCadena & Space(120 - Len(strCadena)) & IIf(Trim(Ser_OrdComp) <> "", Ser_OrdComp & "-" & Cod_OrdComp, Cod_OrdComp)
        
        Plin strCadena
       
    Else
        Plin "       "
        strCadena = ""
        strCadena = strCadena & Space(95 - Len(strCadena))
        
        StrSql = "UP_EXTRAE_PEDIDO_PARA_GUIAREM '" & CodAlmacen & "' , '" & NumMovStk & "'"
        sOrden = RTrim(DevuelveCampo(StrSql, cConnect))
        
        If RTrim(sOrden) <> "" Then
            Me.varPedido = sOrden
        End If
        
        strCadena = strCadena & Space(5) & Me.varPedido
        Plin strCadena
    End If

Plin "     "
Plin "     "

strCadena = Space(100) & Mid(Trim(TxtTransportista.Text), 1, 30)
Plin strCadena
strCadena = Space(95) & Mid(Trim(TxtRuc.Text), 1, 40) & Space(20) & Trim(TxtPlaca.Text)
Plin strCadena
strCadena = Space(98) & Mid(Trim(TxtDomicilio.Text), 1, 30)
Plin strCadena

Plin "     "
Plin "     "


'Imprimimos titulos del detalle
strCadena = Space(2) & "CANTIDAD" & Space(2) & "UNIDAD" & Space(20) & "DESCRIPCION"
Plin strCadena
Plin "     "

End Sub

Sub IMPRIME_DETALLE()
Dim Rs As ADODB.Recordset
Dim varNroReg As Integer
Dim NroReg As Integer
Dim NumTotReg As Integer
Dim strCadena As String
Dim varObserv As String
Dim i As Integer
Dim iMaxLen As Integer
Dim varDescripcion As String
Dim vFila As Integer
Dim vExcede As Integer
Dim varLoteHilado As String

Set Rs = New ADODB.Recordset
Rs.ActiveConnection = cConnect
Rs.CursorLocation = adUseClient
Rs.CursorType = adOpenStatic

If varMoviStk_Guia = False Then
    Rs.Open "EXEC UP_SEL_GUIA_REMISION '" & CodAlmacen & "','" & NumMovStk & "'"
Else
    Rs.Open "EXEC UP_SEL_GUIA_REMISION '" & CodAlmacen & "','" & NumMovStk & "','*'"
End If

iMaxLen = 130

If Rs.RecordCount Then
varNroReg = 1
NroReg = 1

varObserv = Trim(DevuelveCampo("select observaciones from Lg_MoviStk where cod_almacen='" & CodAlmacen & "' and num_movstk='" & NumMovStk & "'", cConnect))
varLoteHilado = Trim(DevuelveCampo("select CASE ISNULL(Glosa_Hilado,'') WHEN '' THEN '' ELSE 'Lote Hilado:' + ISNULL(Glosa_Hilado,'') END from Lg_MoviStk where cod_almacen='" & CodAlmacen & "' and num_movstk='" & NumMovStk & "'", cConnect))
If varObserv = "" Then
    NumTotReg = Rs.RecordCount
Else
    NumTotReg = Rs.RecordCount + 3
End If

Dim TCANTIDAD As Long

Rs.MoveFirst
    'If NumTotReg < 21 Then
    vExcede = 0
        For i = 1 To Rs.RecordCount
            strCadena = Space(5) & Rs.Fields("Cantidad").Value
            TCANTIDAD = TCANTIDAD + Rs.Fields("Cantidad").Value
            strCadena = strCadena & Space(14 - Len(strCadena)) & Rs.Fields("Uni Med").Value & Space(5) & Trim(Rs.Fields("Descripcion").Value)
            
            If Len(strCadena) > 0 Then
                vFila = 1
                 Do While strCadena <> ""
                     varDescripcion = Mid(strCadena, 1, iMaxLen)
                     If vFila = 1 Then
                        Plin varDescripcion
                     Else
                        Plin Space(19) & varDescripcion
                     End If
                     strCadena = Mid(strCadena, iMaxLen + 1, iMaxLen)
                     NroReg = NroReg + 1
                     vFila = vFila + 1
                 Loop
             Else
                 NroReg = NroReg + 1
                 Plin strCadena
             End If
        
            If NroReg > 18 Then
                vExcede = 1
                Exit For
            Else
                Rs.MoveNext
            End If
        Next
        
'         strCadena = Space(5) & TCANTIDAD
         Plin "       "
         Plin "       "
         Plin "       "
         Plin "       "
         strCadena = "  Cantidad Total " & TCANTIDAD
         Plin strCadena
         
        If varMoviStk_Guia = False Then
            If varObserv <> "" Then
                strCadena = "======================================================================================================================================="
                Plin strCadena
                strCadena = Space(3) & "Observacion:"
                Plin strCadena
                strCadena = Space(15) & varObserv
                Plin strCadena
                Plin Space(3) & varLoteHilado
                
                If vNumConosHilos <> 0 Then
                strCadena = Space(15) & "Numero de Conos enviados " & vNumConosHilos
                Plin strCadena
                End If
                NroReg = NroReg + 4
            End If
        End If
        
        'For i = NroReg To 20
        For i = NroReg To 18
            Plin "     "
        Next
        
        'IMPRIME_REFERENCIA
        Plin Chr(12)
        
        If vExcede = 1 Then MsgBox "La cantidad de detalle excede el tamaño de la Guia, algunos datos no se imprimieron, verifique", vbInformation, Me.Caption

End If


End Sub

Sub IMPRIME_REFERENCIA()
Dim RsObs As New Recordset
Dim Rs As New ADODB.Recordset
Dim strCadena As String, iCount As Integer, sNumPed As String, sLineaFin As String
Dim sMarca As String, sPeso As String, sMTC As String, sNoLic As String, _
    sNomTransp As String, sPlaca As String


If varMoviStk_Guia = False Then

    RsObs.Open "select ser_ordcomp,cod_ordcomp,des_protex,des_claordcomp from lg_ordComp a,tx_procesos b,lg_claordcomp c where a.cod_protex *=b.cod_protex and a.cod_claordcomp=c.cod_claordcomp and ser_ordcomp='" & Ser_OrdComp & "' and cod_ordcomp='" & Cod_OrdComp & "'", cConnect, adOpenStatic, adLockReadOnly
    
    If RsObs.RecordCount Then
'        Plin "     "
        strCadena = "-----------------------------------------------------------------------------------------------------------------------------------"
        Plin strCadena
        strCadena = Space(5) & "O/C: " & Ser_OrdComp + "-" + Cod_OrdComp & " - " & RsObs("des_claordcomp") & "   Proceso: " & RsObs("des_protex") & Space(10) & "Almacen: " & DevuelveCampo("select rtrim(nom_almacen) from lg_almacen where cod_almacen='" & CodAlmacen & "'", cConnect) & "   Mov.#: " & NumMovStk
        Plin strCadena
        If DevuelveCampo("select tip_fabrica from tg_control", cConnect) = "2" Then
             strCadena = "      O/P : " & DevuelveCampo("EXEC SM_BUSCA_OPS_OC '" & Ser_OrdComp & "','" & Cod_OrdComp & "',''", cConnect)
             Plin strCadena
        End If
    Else
 '       Plin "     "
        Plin "     "
        strCadena = Space(5) & "Almacen: " & DevuelveCampo("select rtrim(nom_almacen) from lg_almacen where cod_almacen='" & CodAlmacen & "'", cConnect) & "   Mov.#: " & NumMovStk
        Plin strCadena
    End If
    

Else
'    Plin "     "
    Plin "     "
    StrSql = "SELECT isnull(Linea1,'') as 'Linea1', isnull(Linea2,'') as 'Linea2' FROM LG_MOVISTK_GUI WHERE COD_ALMACEN = '" & CodAlmacen & "' and Num_MovStk = '" & NumMovStk & "'"

    RsObs.Open StrSql, cConnect, adOpenStatic, adLockReadOnly

    strCadena = Space(3) & Trim(RsObs("Linea1").Value) & Space(20) & "Almacen: " & DevuelveCampo("SELECT RTRIM(nom_almacen) FROM LG_ALMACEN_GUIAS WHERE cod_almacen='" & CodAlmacen & "'", cConnect) & "   Mov.#: " & NumMovStk
    Plin strCadena
    strCadena = Space(3) & Trim(RsObs("Linea2").Value)
    Plin strCadena
    

End If
    Plin "PUNTO DE PARTIDA:" & DevuelveCampo("SELECT Direccion_de_Partida_Empresa2 FROM TG_CONTROL", cConnect)
    Plin "     "
    
'COMENTARIZADO EL 28/11/2003
'    Rs.Open "Select * from Lg_Transportista where secuencia='" & Trim(Me.TxtSec_Transportista.Text) & "'", cConnect, adOpenStatic, adLockReadOnly
'    If Rs.RecordCount Then
'        strCadena = "DATOS DE LA UNIDAD DE TRANSPORTE:     Marca Y Placa: " & Trim(Rs.Fields("marca_y_placa").Value) & "  Peso: " & CStr(Rs.Fields("peso_vehiculo_kg").Value) & " KG"
'        Plin strCadena
'        strCadena = "Cons.Inscr.MTC: " & Trim(Rs.Fields("reg_transportista").Value) & " NoLic.Cond.: " & Trim(Rs.Fields("num_licencia").Value) & "  " & Trim(Rs.Fields("nom_conductor").Value) & "  -  " & "  -  " & "  -  "
'        Plin strCadena
'    End If
    
    'Agregado el 28/11/2003
    
    Set Rs = New ADODB.Recordset
    StrSql = "SELECT * from Lg_Transportista WHERE secuencia = '" & _
             Trim(Me.TxtSec_Transportista.Text) & "' "
    Rs.Open StrSql, cConnect, adOpenStatic, adLockReadOnly
    
    sMarca = "": sPlaca = "": sPeso = "": sMTC = "": sNoLic = "": sNomTransp = """"
    If Rs.RecordCount > 0 Then
        'sMarca = Trim(IIf(IsNull(Rs!Marca), "", Rs!Marca))
        'sPlaca = Trim(IIf(IsNull(Rs!Placa), "", Rs!Placa))
        sMarca = Trim(IIf(IsNull(Rs!Marca_y_placa), "", Rs!Marca_y_placa))
        sPeso = Format(IIf(IsNull(Rs!Peso_Vehiculo_Kg), "", Rs!Peso_Vehiculo_Kg), "0.00")
        sMTC = Trim(IIf(IsNull(Rs!reg_transportista), "", Rs!reg_transportista))
        sNoLic = Trim(IIf(IsNull(Rs!num_licencia), "", Rs!num_licencia))
        sNomTransp = Trim(IIf(IsNull(Rs!nom_conductor), "", Rs!nom_conductor))
    End If
    If sMTC <> "" Then
        sLineaFin = Space(18) & "Constancia de Inscripcion MTC : " & sMTC
        Plin sLineaFin
    End If
    sLineaFin = Space(18) & "Marca : " & sMarca & " Placa : " & sPlaca & _
                " Peso Seco : " & sPeso
    Plin sLineaFin
    sLineaFin = Space(18) & "Transportista : " & sNomTransp & " Nro.Licencia : " & _
                sNoLic
    Plin sLineaFin

    Set Rs = Nothing

End Sub

Function IMPRIMIR_REPORTE2_SUMIT() As Boolean
Dim StrSql As String, sNomPartida As String
iLin = 0
Set oPrint = New clsPrintFile
    IMPRIMIR_REPORTE2_SUMIT = False
    If Trim(TxtSec_Transportista) = "" Then
        MsgBox "Se debe specificar un Transportista", vbOKOnly + vbExclamation, "Imprimir Guia"
        Exit Function
    End If
    Open "C:\Guia.txt" For Output As #1
    
    'Plin Chr(15) & "   "
    Plin sDoc
    Plin "     "
    Plin "     "
    Plin "     "
    
    sNomPartida = IMPRIME_CABECERA_SUMIT
    IMPRIME_DETALLE_SUMIT sNomPartida
    
    Close #1
    oPrint.SendPrint "c:\GUIA.txt"
    Set oPrint = Nothing
    IMPRIMIR_REPORTE2_SUMIT = True
End Function

Function IMPRIME_CABECERA_SUMIT() As String
Dim RsPro As ADODB.Recordset, Rs As ADODB.Recordset
Dim strCadena As String
Dim varDescripcion As String
Dim sOrden As String

Dim sDirPartida As String, sNomPartida As String, sRucPartida As String, _
    sDirDestino As String, sNomDestino As String, sRucDestino As String
Dim sMarcaPlaca As String, sPeso As String, sMTC As String, sNoLic As String, _
    sNomTransp As String


    Plin "     "
    Plin "     "
    'Plin "     "
    

Set RsPro = New ADODB.Recordset
RsPro.ActiveConnection = cConnect
RsPro.CursorLocation = adUseClient
RsPro.CursorType = adOpenStatic
StrSql = "SELECT Cod_Empresa, Des_Empresa, Direccion, Num_Ruc " & _
         "FROM SEGURIDAD.dbo.SEG_EMPRESAS " & _
         "WHERE Cod_Empresa = '" & vemp1 & "'"
RsPro.Open StrSql

sNomPartida = ""
sRucPartida = ""
sDirPartida = ""
If RsPro.RecordCount > 0 Then
'    sNomPartida = Mid(Trim(TxtTransportista.Text), 1, 40)
'    sRucPartida = Mid(Trim(TxtRuc.Text), 1, 40)
'    sDirPartida = Mid(Trim(TxtDomicilio.Text), 1, 40)
'Else
    sNomPartida = Mid(Trim(IIf(IsNull(RsPro!Des_Empresa), "", RsPro!Des_Empresa)), 1, 40)
    sRucPartida = Mid(Trim(IIf(IsNull(RsPro!Num_Ruc), "", RsPro!Num_Ruc)), 1, 40)
    sDirPartida = Mid(Trim(IIf(IsNull(RsPro!Direccion), "", RsPro!Direccion)), 1, 40)
End If
    
Set RsPro = New ADODB.Recordset
RsPro.ActiveConnection = cConnect
RsPro.CursorLocation = adUseClient
RsPro.CursorType = adOpenStatic

If varMoviStk_Guia = False Then
    If Trim(CodProveedor) = "" Then
        StrSql = "SELECT Des_CenCost AS 'NOMBRE', ' ' AS 'DIRECCION', ' ' AS 'RUC' FROM tg_cencosto WHERE Cod_CenCost ='" & CodCenCost & "'"
    Else
        StrSql = "SELECT Des_Proveedor AS 'NOMBRE', Dom_Proveedor AS 'DIRECCION', Num_Ruc AS 'RUC' FROM LG_PROVEEDOR WHERE Cod_Proveedor ='" & CodProveedor & "'"
    End If
Else
    Select Case Mid(Me.varOpt, 1, 1)
        Case "0":     'TipoProveedor
                        StrSql = "SELECT Des_Proveedor AS 'NOMBRE', Dom_Proveedor AS 'DIRECCION', Num_Ruc AS 'RUC' FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) & "'"
        Case "1":     'TipoCliente
                        StrSql = "SELECT Nom_Cliente AS 'NOMBRE', Direccion AS 'DIRECCION', Num_Ruc AS 'RUC' FROM TG_CLIENTE WHERE Cod_Cliente = '" & Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) & "'"
        Case "2":     'Destinatario
                        StrSql = "SELECT Destinatario as 'NOMBRE', Dom_Destinatario as 'DIRECCION', Ruc_Destinatario as 'RUC' FROM LG_MOVISTK_GUI WHERE COD_ALMACEN = '" & CodAlmacen & "' and Num_MovStk = '" & NumMovStk & "'"
    End Select
   
End If

RsPro.Open StrSql
If Not RsPro.EOF Then
    sNomDestino = Trim(RsPro.Fields("NOMBRE").Value)
    sDirDestino = Trim(RsPro.Fields("DIRECCION").Value)
   
    StrSql = "UP_EXTRAE_PEDIDO_PARA_GUIAREM '" & CodAlmacen & "' , '" & NumMovStk & "'"
    sOrden = RTrim(DevuelveCampo(StrSql, cConnect))

    If RTrim(sOrden) <> "" Then
        Me.varPedido = sOrden
    End If


    
    sRucDestino = Trim(RsPro.Fields("RUC").Value)
    
    strCadena = strCadena & Space(120 - Len(strCadena)) & IIf(Trim(Ser_OrdComp) <> "", Ser_OrdComp & "-" & Cod_OrdComp, Cod_OrdComp)
    
Else

    strCadena = ""
    strCadena = strCadena & Space(95 - Len(strCadena))
    
    StrSql = "UP_EXTRAE_PEDIDO_PARA_GUIAREM '" & CodAlmacen & "' , '" & NumMovStk & "'"
    sOrden = RTrim(DevuelveCampo(StrSql, cConnect))
    
    If RTrim(sOrden) <> "" Then
        Me.varPedido = sOrden
    End If
    
    strCadena = strCadena & Space(5) & Me.varPedido

End If


Plin "     "
strCadena = Space(16) & Format(Date, "dd/mm/yyyy") & Space(24) & Format(Date, "dd/mm/yyyy")
Plin strCadena
strCadena = Space(105) & Trim(TxtSerie.Text) & "-" & Trim(TxtNumero.Text)
Plin strCadena
Plin "     "

Plin "     "
Plin "     "

strCadena = Space(15) & sDirPartida & Space(65 - Len(sDirPartida)) & sDirDestino
Plin strCadena
        
'Aqui imprimimos los datos obtenidos


Plin "     "
Plin "     "
Plin "     "
Plin "     "


Set Rs = New ADODB.Recordset
StrSql = "SELECT * from Lg_Transportista WHERE secuencia = '" & _
         Trim(Me.TxtSec_Transportista.Text) & "' "
Rs.Open StrSql, cConnect, adOpenStatic, adLockReadOnly

sMarcaPlaca = "": sPeso = "": sMTC = "": sNoLic = "": sNomTransp = ""
If Rs.RecordCount Then
    sMarcaPlaca = Trim(IIf(IsNull(Rs!Marca_y_placa), "", Rs!Marca_y_placa))
    sPeso = Format(IIf(IsNull(Rs!Peso_Vehiculo_Kg), "", Rs!Peso_Vehiculo_Kg), "0.00")
    sMTC = Trim(IIf(IsNull(Rs!reg_transportista), "", Rs!reg_transportista))
    sNoLic = Trim(IIf(IsNull(Rs!num_licencia), "", Rs!num_licencia))
    sNomTransp = Trim(IIf(IsNull(Rs!nom_conductor), "", Rs!nom_conductor))
Else
    sMarcaPlaca = TxtPlaca
    sNomTransp = TxtTransportista
End If

sTrans = sNomTransp
IMPRIME_CABECERA_SUMIT = IIf(vRespuesta = "S", sNomTransp, sNomPartida)

strCadena = Space(15) & sNomDestino & Space(85 - Len(sNomDestino)) & sMarcaPlaca
Plin strCadena
strCadena = Space(15) & sRucDestino & Space(85 - Len(sRucDestino)) & sMTC
Plin strCadena
strCadena = Space(100) & sNoLic
Plin strCadena
Plin "     "
Plin Space(9) & String(120, "-")

'Imprimimos titulos del detalle
strCadena = Space(12) & "CANTIDAD" & Space(2) & "UNIDAD" & Space(20) & "DESCRIPCION"
Plin strCadena
Plin Space(9) & String(120, "-")

End Function

Sub IMPRIME_DETALLE_SUMIT(sNomPartida As String)
Dim Rs As ADODB.Recordset
Dim varNroReg As Integer
Dim NroReg As Integer
Dim NumTotReg As Integer
Dim strCadena As String
Dim varObserv As String, varObserv1 As String 'Para Observaciones en Guia Manual
Dim i As Integer
Dim iMaxLen As Integer
Dim varDescripcion As String
Dim vFila As Integer
Dim vExcede As Integer

Set Rs = New ADODB.Recordset
Rs.ActiveConnection = cConnect
Rs.CursorLocation = adUseClient
Rs.CursorType = adOpenStatic

If varMoviStk_Guia = False Then
    Rs.Open "EXEC UP_SEL_GUIA_REMISION '" & CodAlmacen & "','" & NumMovStk & "'"
Else
    Rs.Open "EXEC UP_SEL_GUIA_REMISION '" & CodAlmacen & "','" & NumMovStk & "','*'"
End If

iMaxLen = 130

If Rs.RecordCount Then
    varNroReg = 1
    NroReg = 1
    varObserv1 = ""
    If varMoviStk_Guia = False Then
        StrSql = "select observaciones from Lg_MoviStk where cod_almacen='" & CodAlmacen & "' and num_movstk='" & NumMovStk & "'"
    Else
        StrSql = "Select ISNULL(Linea1, '') FROM LG_MOVISTK_GUI WHERE Num_MovStk = '" & NumMovStk & "' AND Cod_Almacen = '" & CodAlmacen & "'"
        varObserv1 = Trim(DevuelveCampo(StrSql, cConnect))
        StrSql = "Select ISNULL(Linea2, '') FROM LG_MOVISTK_GUI WHERE Num_MovStk = '" & NumMovStk & "' AND Cod_Almacen = '" & CodAlmacen & "'"
    End If
    
    varObserv = Trim(DevuelveCampo(StrSql, cConnect))
    
    If varObserv = "" Then
        NumTotReg = Rs.RecordCount
    Else
        NumTotReg = Rs.RecordCount + 3
    End If
    
    Rs.MoveFirst
    'Observaciones 1 para La guia Manual
    Plin Space(15) & varObserv1
    vExcede = 0
    For i = 1 To Rs.RecordCount
        strCadena = Space(15) & Rs.Fields("Cantidad").Value
        strCadena = strCadena & Space(23 - Len(strCadena)) & Rs.Fields("Uni Med").Value & Space(5) & Trim(Rs.Fields("Descripcion").Value)
        
        If Len(strCadena) > 0 Then
            vFila = 1
             Do While strCadena <> ""
                 varDescripcion = Mid(strCadena, 1, iMaxLen)
                 If vFila = 1 Then
                    Plin varDescripcion
                 Else
                    Plin Space(19) & varDescripcion
                 End If
                 strCadena = Mid(strCadena, iMaxLen + 1, iMaxLen)
                 NroReg = NroReg + 1
                 vFila = vFila + 1
             Loop
         Else
             NroReg = NroReg + 1
             Plin strCadena
         End If
    
        If NroReg > 23 Then
            vExcede = 1
            Exit For
        Else
            Rs.MoveNext
        End If
    Next
    
    If varObserv <> "" Then
        strCadena = "     "
        Plin strCadena
        strCadena = Space(15) & "Obs:"
        Plin strCadena
        strCadena = Space(15) & varObserv
        Plin strCadena
        NroReg = NroReg + 3
    End If
    
    'For i = NroReg To 20
    For i = NroReg To 23
        Plin "     "
    Next
    
    IMPRIME_REFERENCIA_SUMIT sNomPartida
    Plin Chr(12)
    
    If vExcede = 1 Then MsgBox "La cantidad de detalle excede el tamaño de la Guia, algunos datos no se imprimieron, verifique", vbInformation, Me.Caption

End If


End Sub

Sub IMPRIME_REFERENCIA_SUMIT(sNomPartida As String)
Dim RsObs As New Recordset
Dim Rs As New ADODB.Recordset
Dim strCadena As String, iCount As Integer, sNumPed As String

sNumPed = ""
If varMoviStk_Guia = False Then
    
    RsObs.Open "select ser_ordcomp,cod_ordcomp,des_protex,des_claordcomp " & _
               "from lg_ordComp a,tx_procesos b,lg_claordcomp c " & _
               "where a.cod_protex *= b.cod_protex " & _
               "and a.cod_claordcomp = c.cod_claordcomp " & _
               "and ser_ordcomp = '" & Ser_OrdComp & "' " & _
               "and cod_ordcomp = '" & Cod_OrdComp & "'", _
               cConnect, adOpenStatic, adLockReadOnly
    
    If RsObs.RecordCount Then
        Plin "     "
        Plin "     "
        strCadena = Space(25) & "O/C: " & Ser_OrdComp + "-" + Cod_OrdComp & " - " & RsObs("des_claordcomp") & "   Proceso: " & RsObs("des_protex") & Space(10) & "Almacen: " & DevuelveCampo("select rtrim(nom_almacen) from lg_almacen where cod_almacen='" & CodAlmacen & "'", cConnect) & "   Mov.#: " & NumMovStk
        Plin strCadena
'        If DevuelveCampo("select tip_fabrica from tg_control", cConnect) = "2" Then
'             strCadena = Space(25) & "N/P : " & DevuelveCampo("EXEC SM_BUSCA_OPS_OC '" & Ser_OrdComp & "','" & Cod_OrdComp & "',''", cConnect)
'             Plin strCadena
'        End If
    Else
        Plin "     "
        Plin "     "
        strCadena = Space(25) & "Almacen: " & DevuelveCampo("select rtrim(nom_almacen) from lg_almacen where cod_almacen='" & CodAlmacen & "'", cConnect) & "   Mov.#: " & NumMovStk
        Plin strCadena
    End If
Else
    Plin "PUNTO DE PARTIDA:" & DevuelveCampo("SELECT Direccion_de_Partida_Empresa2 FROM TG_CONTROL", cConnect)
    Plin "     "
    Plin "     "
    strCadena = "SELECT ISNULL(Pedido, '') FROM LG_MOVISTK_GUI WHERE Num_MovStk = '" & NumMovStk & "' AND Cod_Almacen = '" & CodAlmacen & "'"
    sNumPed = DevuelveCampo(strCadena, cConnect)
    strCadena = Space(25) & "Mov.#: " & NumMovStk
    Plin strCadena
End If
For iCount = 1 To 5
    Plin "     "
Next iCount
strCadena = Space(15) & sNomPartida
Plin strCadena
strCadena = Space(15) & IIf((vRespuesta = "N" And varMoviStk_Guia), sTrans, "")
Plin strCadena
For iCount = 1 To 2
    Plin "     "
Next iCount
strCadena = Space(15) & sNumPed
Plin strCadena
End Sub

Private Sub TxtSerie_LostFocus()
    TxtSerie = Format(TxtSerie, "000")
End Sub

Private Sub TxtTransportista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtDomicilio.SetFocus
End Sub


Sub Plin(ByVal Text)
If IsNull(Text) Then
       Text = ""
    End If
    Print #1, Text
    iLin = iLin + 1
End Sub

Public Sub BUSCA_TRANSPORTISTA(Tipo As Integer)
    Select Case Tipo
        Case 1:
                    StrSql = "Select nom_conductor as Descripcion from Lg_Transportista where secuencia='" & Trim(Me.TxtSec_Transportista.Text) & "'"
                    Me.TxtNom_Transportista.Text = Trim(DevuelveCampo(StrSql, cConnect))
        Case 2, 3:
                    Dim oTipo As New frmBusqGeneral
                    Dim Rs As New ADODB.Recordset
                    Set oTipo.oParent = Me
                    
                    If Tipo = 2 Then
                        oTipo.sQuery = "Select secuencia as Codigo,nom_conductor as Descripcion from lg_transportista Where nom_conductor like '%" & Trim(Me.TxtNom_Transportista.Text) & "%' order by 2"
                    Else
                        oTipo.sQuery = "Select secuencia as Codigo,nom_conductor as Descripcion from lg_transportista order by 2"
                    End If
                    
                    oTipo.Cargar_Datos
                    oTipo.Show 1
                    If Codigo <> "" Then
                        Me.TxtSec_Transportista.Text = Trim(Codigo)
                        Me.TxtNom_Transportista.Text = Trim(Descripcion)
                        Codigo = "": Descripcion = ""
                    End If
                    Set oTipo = Nothing
                    Set Rs = Nothing
    End Select
    Me.cmdImprimir.SetFocus
End Sub

Sub IMPRIME_CABECERA_02()
    Dim RsPro As ADODB.Recordset, Rs As Recordset
    Dim RsUbiGeo As ADODB.Recordset
    Dim sLINEA As String
    Dim Scliente_Especial As String
    
    Dim sdesplaza_dist_partida As Integer, sdesplaza_prov_partida As Integer, sdesplaza_dpto_partida As Integer
    Dim sdesplaza_dist_llegada As Integer, sdesplaza_prov_llegada As Integer, sdesplaza_dpto_llegada As Integer
    
    Dim sdesplazat_dist_partida As Integer, sdesplazat_prov_partida As Integer, sdesplazat_dpto_partida As Integer
    Dim sdesplazat_dist_llegada As Integer, sdesplazat_prov_llegada As Integer, sdesplazat_dpto_llegada As Integer
    
    Dim SOrigen_Pto_Partida As Integer, SOrigen_Pto_LLegada As Integer

    
    Dim sOrigen_Dist_LLegada As Integer, sAncho_Dist_LLegada As Integer
    Dim sOrigen_Prov_LLegada As Integer, sAncho_Prov_LLegada As Integer
    Dim sOrigen_Dpto_LLegada As Integer, sAncho_Dpto_LLegada As Integer
    
    Dim sOrigen_Dist_Partida As Integer, sancho_dist_partida As Integer
    Dim sOrigen_Prov_Partida As Integer, sAncho_Prov_Partida As Integer
    Dim sOrigen_Dpto_Partida As Integer, sAncho_Dpto_Partida As Integer
    
    
    Dim sMarca_Transporte_u As Integer
    Dim sPlaca_Transporte_u As Integer
    Dim sCodigo_MTC_u As Integer
    Dim SDestinatario_u As Integer
    Dim sRuc_Destinatario_u As Integer
    Dim sNro_Licencia_u As Integer
    
    Dim sNro_Lineas_Adicionales_Header As Integer
    
    Dim sPuntoDePartida As String, sPuntoDeLlegada As String
    Dim sDistrito_Partida As String, sDistrito_LLegada As String
    Dim sProvincia_Partida As String, sProvincia_LLegada As String
    Dim sDpto_Partida As String, sDpto_LLegada As String
    
    Dim SDestinatario As String, Sruc As String
    Dim sMarca As String, sPeso As String, sMTC As String, sNoLic As String, sChofer As String, sPlaca As String
    
    Dim sguia As String, sFecha As String, sMotivo As String, sAlmacen As String
    Dim sTransportista As String, sRUC_Transportista As String, sDireccionTransportista As String
    
    
    Set RsUbiGeo = New ADODB.Recordset
    RsUbiGeo.ActiveConnection = cConnect
    RsUbiGeo.CursorLocation = adUseClient
    RsUbiGeo.CursorType = adOpenStatic

    RsUbiGeo.Open "EXEC lg_extrae_datos_ubigeo_guia_remision '" & CodAlmacen & "','" & NumMovStk & "'"
    
    sDpto_Partida = RsUbiGeo.Fields("dpto_partida")
    sProvincia_Partida = RsUbiGeo.Fields("prov_partida")
    sDistrito_Partida = RsUbiGeo.Fields("dist_partida")
    
    sDpto_LLegada = RTrim(RsUbiGeo.Fields("dpto_LLegada"))
    sProvincia_LLegada = RTrim(RsUbiGeo.Fields("prov_LLegada"))
    sDistrito_LLegada = RTrim(RsUbiGeo.Fields("dist_LLegada"))
    
    sNro_Lineas_Adicionales_Header = DevuelveCampo("SELECT Nro_Lineas_Adicionales_Header from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    
   sOrigen_Dist_Partida = DevuelveCampo("SELECT Origen_Dist_Partida from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    sancho_dist_partida = DevuelveCampo("SELECT Ancho_Dist_Partida from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    
    sOrigen_Prov_Partida = DevuelveCampo("SELECT Origen_Prov_Partida from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    sAncho_Prov_Partida = DevuelveCampo("SELECT Ancho_Prov_Partida from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    sOrigen_Dpto_Partida = DevuelveCampo("SELECT Origen_Dpto_Partida from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    sAncho_Dpto_Partida = DevuelveCampo("SELECT Ancho_Dpto_Partida from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    
    sOrigen_Dist_LLegada = DevuelveCampo("SELECT Origen_Dist_LLegada from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    sAncho_Dist_LLegada = DevuelveCampo("SELECT Ancho_Dist_LLegada from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
   
    sOrigen_Prov_LLegada = DevuelveCampo("SELECT Origen_Prov_LLegada from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    sAncho_Prov_LLegada = DevuelveCampo("SELECT Ancho_Prov_LLegada from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)

   sOrigen_Dpto_LLegada = DevuelveCampo("SELECT Origen_Dpto_LLegada from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    sAncho_Dpto_LLegada = DevuelveCampo("SELECT Ancho_Dpto_LLegada from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)

    SOrigen_Pto_Partida = DevuelveCampo("SELECT Origen_pto_partida from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    SOrigen_Pto_LLegada = DevuelveCampo("SELECT Origen_pto_llegada from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    
    
    sMarca_Transporte_u = DevuelveCampo("SELECT Marca_Transporte from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
 
    sCodigo_MTC_u = DevuelveCampo("SELECT Codigo_MTC from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    SDestinatario_u = DevuelveCampo("SELECT Destinatario from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    sRuc_Destinatario_u = DevuelveCampo("SELECT ruc_Destinatario from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    sNro_Licencia_u = DevuelveCampo("SELECT Nro_Licencia from  Cf_Coordenadas_Emision_Guia where tipo = '1' ", cConnect)
    
    Set RsPro = New ADODB.Recordset
    RsPro.ActiveConnection = cConnect
    RsPro.CursorLocation = adUseClient
    RsPro.CursorType = adOpenStatic

    If varMoviStk_Guia = False Then
        StrSql = "SELECT Des_Proveedor AS 'NOMBRE', Dom_Proveedor AS 'DIRECCION', Num_Ruc AS 'RUC' FROM LG_PROVEEDOR WHERE Cod_Proveedor ='" & CodProveedor & "'"
    Else
        Select Case Mid(Me.varOpt, 1, 1)
            '--+--------------------------------------------------+--
            'TipoProveedor
            '--+--------------------------------------------------+--
            Case "0": StrSql = "SELECT Des_Proveedor AS 'NOMBRE', Dom_Proveedor AS 'DIRECCION', Num_Ruc AS 'RUC' FROM LG_PROVEEDOR WHERE Cod_Proveedor = '" & Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) & "'"
            '--+--------------------------------------------------+--
            'TipoCliente
            '--+--------------------------------------------------+--
            Case "1": StrSql = "SELECT Nom_Cliente AS 'NOMBRE', Direccion AS 'DIRECCION', Num_Ruc AS 'RUC' FROM TG_CLIENTE WHERE Cod_Cliente = '" & Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) & "'"
            '--+--------------------------------------------------+--
            'Destinatario
            '--+--------------------------------------------------+--
            Case "2": StrSql = "SELECT Destinatario as 'NOMBRE', Dom_Destinatario as 'DIRECCION', Ruc_Destinatario as 'RUC' FROM LG_MOVISTK_GUI WHERE COD_ALMACEN = '" & CodAlmacen & "' and Num_MovStk = '" & NumMovStk & "'"
            '--+--------------------------------------------------+--
            'Sector Propio
            '--+--------------------------------------------------+--
            Case "3": StrSql = "SELECT Des_SecConf as 'NOMBRE', '' as 'DIRECCION', ''  as 'RUC' FROM CF_SECTORES_CONFECCION  WHERE  Cod_SecConf= '" & Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) & "'"
            '--+--------------------------------------------------+--
            'ALMACEN
            '--+--------------------------------------------------+--
            Case "6": StrSql = "SELECT Destino_Guia  as 'NOMBRE', Dir_Almacen as 'DIRECCION', Num_Ruc as 'RUC' FROM CF_ALMACEN WHERE  Cod_Almacen = '" & Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) & "'"
            '--+--------------------------------------------------+--
            'ALMACEN ADUANA
            '--+--------------------------------------------------+--
            Case "7": StrSql = "SELECT Nom_AlmacenAduana  as 'NOMBRE', Dir_AlmacenAduana as 'DIRECCION', RUC_AlamcenAduana as 'RUC' FROM CF_ALMACEN_ADUANA WHERE  Cod_AlmacenAduana = '" & Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) & "'"
            '--+--------------------------------------------------+--
            'CENTRO DE COSTO
            '--+--------------------------------------------------+--
            Case "8": StrSql = "SELECT Des_CenCost  as 'NOMBRE', '' as 'DIRECCION', '' as 'RUC' FROM tg_cencosto WHERE  Cod_CenCost = '" & Trim(Mid(Me.varOpt, 2, Len(Me.varOpt) - 1)) & "'"
        End Select
    End If
    
    
    '****************************************************************************************************************************************************************************************************************************************************************************************
    '== OBTENGO LOS DATOS
    '****************************************************************************************************************************************************************************************************************************************************************************************
    
    
     Scliente_Especial = DevuelveCampo("SELECT Flg_Direccion_Especial FROM Lg_Almacen WHERE Flg_Direccion_Especial='S' and COD_ALMACEN = '" & CodAlmacen & "' ", cConnect)
     
     If Scliente_Especial = "S" Then
        sPuntoDePartida = DevuelveCampo("SELECT Direccion_de_Partida_Empresa_Especial FROM TG_CONTROL", cConnect)
    Else
        sPuntoDePartida = DevuelveCampo("SELECT Direccion_de_Partida_Empresa2 FROM TG_CONTROL", cConnect)
    End If

       
    RsPro.Open StrSql
    If Not RsPro.EOF Then
             SDestinatario = Trim(RsPro.Fields("NOMBRE").Value)
             Sruc = Trim(RsPro.Fields("RUC").Value)
             sPuntoDeLlegada = Trim(RsPro.Fields("DIRECCION").Value) & Space(4) & IIf(Trim(Ser_OrdComp) <> "", Ser_OrdComp & "-" & Cod_OrdComp, Cod_OrdComp)

    End If

    sguia = Trim(TxtSerie.Text) & "-" & Trim(TxtNumero.Text)
    sAlmacen = DevuelveCampo("select isnull(Nom_Almacen_Guia,'') From LG_Almacen Where cod_almacen='" & CodAlmacen & "'", cConnect)
    sMotivo = RPad(TxtDes_Motivo.Text, 50, " ")
    
    Dim sFec_MovStk As Date
        
    sFec_MovStk = DevuelveCampo("SELECT fec_movstk FROM LG_MOVISTK WHERE COD_ALMACEN = '" & CodAlmacen & "' and Num_MovStk = '" & NumMovStk & "'", cConnect)
       
    sFecha = FormatDateTime(sFec_MovStk, vbShortDate)


    sTransportista = Mid(Trim(TxtTransportista.Text), 1, 30)
    sRUC_Transportista = Mid(Trim(TxtRuc.Text), 1, 40)
    sDireccionTransportista = Mid(Trim(TxtDomicilio.Text), 1, 30)

    
    Set Rs = New ADODB.Recordset
    StrSql = "SELECT * from Lg_Transportista WHERE secuencia = '" & Trim(Me.TxtSec_Transportista.Text) & "' "
    Rs.Open StrSql, cConnect, adOpenStatic, adLockReadOnly

    
    sMarca = "": sPlaca = "": sPeso = "": sMTC = "": sNoLic = "": sChofer = ""
    If Rs.RecordCount > 0 Then
        sMarca = Trim(Rs!Marca)
        sPlaca = Trim(Rs!Placa)
        sPeso = Format(Rs!Peso_Vehiculo_Kg, "0.00")
        sMTC = Trim(Rs!reg_transportista)
        sNoLic = Trim(Rs!num_licencia)
        sChofer = Trim(Rs!nom_conductor)

    Else
        sPlaca = TxtPlaca
        sChofer = TxtTransportista
    End If

    Rs.Close
    
    sdesplaza_dist_partida = 0
    If sancho_dist_partida > Len(RTrim(sDistrito_Partida)) Then
        sdesplaza_dist_partida = sancho_dist_partida - Len(RTrim(sDistrito_Partida))
    End If
    
        
    sdesplaza_prov_partida = 0
    If sAncho_Prov_Partida > Len(RTrim(sProvincia_Partida)) Then
        sdesplaza_prov_partida = sAncho_Prov_Partida - Len(RTrim(sProvincia_Partida))
    End If
    
    sdesplaza_dpto_partida = 0
    If sAncho_Dpto_Partida > Len(RTrim(sDpto_Partida)) Then
        sdesplaza_dpto_partida = sAncho_Dpto_Partida - Len(RTrim(sDpto_Partida))
    End If
    
    
    sdesplaza_dist_llegada = 0
    If sAncho_Dist_LLegada > Len(RTrim(sDistrito_LLegada)) Then
        sdesplaza_dist_llegada = sAncho_Dist_LLegada - Len(RTrim(sDistrito_LLegada))
    End If
    
        
    sdesplaza_prov_llegada = 0
    If sAncho_Prov_LLegada > Len(RTrim(sProvincia_LLegada)) Then
        sdesplaza_prov_llegada = sAncho_Prov_LLegada - Len(RTrim(sProvincia_LLegada))
    End If
    
    sdesplaza_dpto_llegada = 0
    If sAncho_Dpto_LLegada > Len(RTrim(sDpto_LLegada)) Then
        sdesplaza_dpto_llegada = sAncho_Dpto_LLegada - Len(RTrim(sDpto_LLegada))
    End If
        
        
    sdesplazat_prov_partida = sOrigen_Prov_Partida - sOrigen_Dist_Partida - sancho_dist_partida + sdesplaza_dist_partida
    sdesplazat_dpto_partida = sOrigen_Dpto_Partida - sOrigen_Prov_Partida - sAncho_Prov_Partida + sdesplaza_prov_partida
    
    sdesplazat_dist_llegada = sOrigen_Dist_LLegada - sOrigen_Dpto_Partida - sAncho_Dpto_Partida + sdesplaza_dpto_partida
    sdesplazat_prov_llegada = sOrigen_Prov_LLegada - sOrigen_Dist_LLegada - sAncho_Dist_LLegada + sdesplaza_dist_llegada
    sdesplazat_dpto_llegada = sOrigen_Dpto_LLegada - sOrigen_Prov_LLegada - sAncho_Prov_LLegada + sdesplaza_prov_llegada
    

    '****************************************************************************************************************************************************************************************************************************************************************************************
    '==> IMPRIMO LOS VALORES OBTENIDOS
    '****************************************************************************************************************************************************************************************************************************************************************************************
    Dim Scontador As Integer, SFila As Integer, SFila_Anterior As Integer
        
    SFila = DevuelveCampo("SELECT Fila from  Cf_Coordenadas_Emision_Guia_Lineas where tipo = '1' and nro_linea = '1' ", cConnect)
    
    If sguia = "S" Then
        Plin "     "
    End If
    
       
    SFila = 5
    For Scontador = 1 To 5
    'For Scontador = 1 To SFila - 1
        Plin "     "
    Next
    
    sLINEA = Space(8) & Format(sFecha, "dd") & Space(6) & Format(sFecha, "mm") & Space(6) & Format(sFecha, "yyyy") & Space(10) & Format(sFecha, "dd") & Space(6) & Format(sFecha, "mm") & Space(6) & Format(sFecha, "yyyy") & Space(60) & sguia
    'sLINEA = Space(5) & sFecha & Space(20) & sFecha & Space(60) & sguia
    Plin sLINEA
    
    SFila_Anterior = SFila
    SFila = DevuelveCampo("SELECT Fila from  Cf_Coordenadas_Emision_Guia_Lineas where tipo = '1' and nro_linea = '2' ", cConnect)
    'For Scontador = 1 To SFila - SFila_Anterior
    For Scontador = 1 To 1
        Plin "     "
    Next
     
    sLINEA = Space(7) & Space(SOrigen_Pto_Partida - 1) & RTrim(SDestinatario) & Space(SOrigen_Pto_LLegada - SOrigen_Pto_Partida - 21) & Trim(Sruc)
    Plin sLINEA
    Plin "     "
    sLINEA = Space(10) & RTrim(sPuntoDePartida) & Space(13) & Trim(sPuntoDeLlegada)
    Plin sLINEA
    
    SFila_Anterior = SFila
    SFila = DevuelveCampo("SELECT Fila from  Cf_Coordenadas_Emision_Guia_Lineas where tipo = '1' and nro_linea = '3' ", cConnect)
    'For Scontador = 1 To SFila - SFila_Anterior - 1
    For Scontador = 1 To 2
        Plin "     "
    Next
      
    Plin " "
    'Plin Space(sOrigen_Dist_Partida - 1) & Mid(RTrim(sDistrito_Partida), 1, sancho_dist_partida) & Space(sdesplazat_prov_partida) & Mid(RTrim(sProvincia_Partida), 1, sAncho_Prov_Partida) & Space(sdesplazat_dpto_partida) & Mid(RTrim(sDpto_Partida), 1, sAncho_Dpto_Partida) & Space(sdesplazat_dist_llegada) & Mid(RTrim(sDistrito_LLegada), 1, sAncho_Dist_LLegada) & Space(sdesplazat_prov_llegada) & Mid(RTrim(sProvincia_LLegada), 1, sAncho_Prov_LLegada) & Space(sdesplazat_dpto_llegada) & Mid(RTrim(sDpto_LLegada), 1, sAncho_Dpto_LLegada) & Space(15) & Sruc
    
    SFila_Anterior = SFila
    SFila = DevuelveCampo("SELECT Fila from  Cf_Coordenadas_Emision_Guia_Lineas where tipo = '1' and nro_linea = '4' ", cConnect)
    For Scontador = 1 To SFila - SFila_Anterior - 1
        Plin "     "
    Next
      
    'sLINEA = Space(sMarca_Transporte_u - 1) & RTrim(sMarca) & "  placa : " & UCase(sPlaca)
    sLINEA = " "
    Plin sLINEA
    
    SFila_Anterior = SFila
    SFila = DevuelveCampo("SELECT Fila from  Cf_Coordenadas_Emision_Guia_Lineas where tipo = '1' and nro_linea = '5' ", cConnect)
    For Scontador = 1 To SFila - SFila_Anterior - 1
        Plin "     "
    Next
    
    
   ' Plin Space(SDestinatario_u - 1) & Space(20) & SDestinatario
    
    SFila_Anterior = SFila
    SFila = DevuelveCampo("SELECT Fila from  Cf_Coordenadas_Emision_Guia_Lineas where tipo = '1' and nro_linea = '6' ", cConnect)
    For Scontador = 1 To SFila - SFila_Anterior - 1
        Plin "     "
    Next
    Plin "     "
    'sLINEA = Space(sRuc_Destinatario_u - 1) & Sruc & Space(sCodigo_MTC_u - sRuc_Destinatario_u - Len(Sruc)) & sMTC
 '   sLINEA = Space(sRuc_Destinatario_u - 1) & Sruc
  '  Plin sLINEA
    
    SFila_Anterior = SFila
    SFila = DevuelveCampo("SELECT Fila from  Cf_Coordenadas_Emision_Guia_Lineas where tipo = '1' and nro_linea = '7' ", cConnect)
    For Scontador = 1 To SFila - SFila_Anterior - 1
        Plin "     "
    Next
    
    
    'sLINEA = Space(sNro_Licencia_u - 1) & sNoLic
    sLINEA = ""
    Plin sLINEA
    Plin sLINEA
    For Scontador = 1 To 2
    'For Scontador = 1 To sNro_Lineas_Adicionales_Header
        Plin "     "
    Next


End Sub
Sub IMPRIME_DETALLE_02()
Dim Rs As ADODB.Recordset
Dim varNroReg As Integer
Dim NroReg As Integer
Dim NumTotReg As Integer
Dim strCadena As String
Dim varObserv As String
Dim i As Integer
Dim iMaxLen As Integer
Dim varDescripcion As String
Dim vFila As Integer
Dim vExcede As Integer
Dim varLoteHilado As String
Dim sLINEA As String

'Dim snro_linea_trailer As Integer

Set Rs = New ADODB.Recordset
Rs.ActiveConnection = cConnect
Rs.CursorLocation = adUseClient
Rs.CursorType = adOpenStatic

If varMoviStk_Guia = False Then
    Rs.Open "EXEC UP_SEL_GUIA_REMISION '" & CodAlmacen & "','" & NumMovStk & "'"
Else
    Rs.Open "EXEC UP_SEL_GUIA_REMISION '" & CodAlmacen & "','" & NumMovStk & "','*'"
End If


Dim sFilaMAx As Integer, sNroAdicHeader As Integer, snro_linea_razon_social_transportista As Integer
Dim snro_linea_trailer As Integer
   
        
sFilaMAx = DevuelveCampo("select max(fila) from Cf_Coordenadas_Emision_Guia_lineas where tipo='1'", cConnect)
sNroAdicHeader = DevuelveCampo("select nro_lineas_adicionales_header from Cf_Coordenadas_Emision_Guia where tipo='1'", cConnect)
    
snro_linea_trailer = DevuelveCampo("select nro_linea_trailer from Cf_Coordenadas_Emision_Guia where tipo='1'", cConnect)
snro_linea_razon_social_transportista = DevuelveCampo("select nro_linea_razon_social_transportista from Cf_Coordenadas_Emision_Guia where tipo='1'", cConnect)
    
    
sFilaMAx = sFilaMAx + sNroAdicHeader



Dim SInicio_Bultos As Integer, SInicio_Descripcion_Textil As Integer, sAncho_Descripcion_Textil As Integer
Dim SInicio_Grupo_Textil  As Integer, sUnidad_MEdida_Textil As Integer, SPeso_Textil As Integer
Dim sNro_Reg_Textil As Integer, sTrailer_Textil As Integer



SInicio_Bultos = DevuelveCampo("select inicio_bultos from Cf_Coordenadas_Emision_Guia where tipo = '1'", cConnect)
SInicio_Descripcion_Textil = DevuelveCampo("select Inicio_Descripcion_Textil from Cf_Coordenadas_Emision_Guia where tipo = '1'", cConnect)
sAncho_Descripcion_Textil = DevuelveCampo("select Ancho_Descripcion_Textil from Cf_Coordenadas_Emision_Guia where tipo = '1'", cConnect)
SInicio_Grupo_Textil = DevuelveCampo("select Inicio_Grupo_Textil  from Cf_Coordenadas_Emision_Guia where tipo = '1'", cConnect)

sUnidad_MEdida_Textil = DevuelveCampo("select Unidad_MEdida_Textil  from Cf_Coordenadas_Emision_Guia where tipo = '1'", cConnect)
SPeso_Textil = DevuelveCampo("select Peso_Textil  from Cf_Coordenadas_Emision_Guia where tipo = '1'", cConnect)

sNro_Reg_Textil = DevuelveCampo("select Nro_Reg_Textil  from Cf_Coordenadas_Emision_Guia where tipo = '1'", cConnect)
sTrailer_Textil = DevuelveCampo("select Trailer_Textil  from Cf_Coordenadas_Emision_Guia where tipo = '1'", cConnect)


iMaxLen = sAncho_Descripcion_Textil

If Rs.RecordCount Then
varNroReg = 1
NroReg = 1

varObserv = Trim(DevuelveCampo("select observaciones from Lg_MoviStk where cod_almacen='" & CodAlmacen & "' and num_movstk='" & NumMovStk & "'", cConnect))
varLoteHilado = Trim(DevuelveCampo("select CASE ISNULL(Glosa_Hilado,'') WHEN '' THEN '' ELSE 'Lote Hilado:' + ISNULL(Glosa_Hilado,'') END from Lg_MoviStk where cod_almacen='" & CodAlmacen & "' and num_movstk='" & NumMovStk & "'", cConnect))
If varObserv = "" Then
    NumTotReg = Rs.RecordCount
Else
    NumTotReg = Rs.RecordCount + 3
End If

Dim TCANTIDAD As Long

Rs.MoveFirst

    vExcede = 0
        For i = 1 To Rs.RecordCount
            
            TCANTIDAD = TCANTIDAD + Rs.Fields("Cantidad").Value
            strCadena = Trim(Rs.Fields("Descripcion").Value)
          
            
            If Len(strCadena) > 0 Then
                 vFila = 1
                 Do While strCadena <> ""
                     varDescripcion = Mid(strCadena, 1, iMaxLen)
                     If vFila = 1 Then
                        'sLINEA = Space(SInicio_Bultos - 1) & Rs.Fields("Num_Bultos").Value
                        'sLINEA = Space(SInicio_Bultos - 1) & Rs.Fields("Num_Bultos").Value
                        sLINEA = Space(29) & Str(Rs.Fields("Cantidad").Value) & Space(4) & "Kgs"
                        sLINEA = sLINEA & Space(5) & RTrim(varDescripcion)
                        'sLINEA = sLINEA & Space(SInicio_Grupo_Textil - Len(sLINEA) - 1) & Trim(Rs.Fields("Grupo").Value)
                        'sLINEA = sLINEA & Space(sUnidad_MEdida_Textil - Len(sLINEA) - 1) & Trim(Rs.Fields("uni med").Value)
                        'sLINEA = sLINEA & Space(SPeso_Textil - Len(sLINEA) - 1) & Trim(Rs.Fields("cantidad").Value)
                        Plin sLINEA
                        sFilaMAx = sFilaMAx + 1
                     Else
                        Plin Space(46) & varDescripcion
                        sFilaMAx = sFilaMAx + 1
                     End If
                     strCadena = Mid(strCadena, iMaxLen + 1, iMaxLen)
                     NroReg = NroReg + 1
                     vFila = vFila + 1
                 Loop
             Else
                 NroReg = NroReg + 1
                 Plin strCadena
             End If
          
        
            If NroReg > sNro_Reg_Textil Then
                vExcede = 1
                Exit For
            Else
                Rs.MoveNext
            End If
        Next
        
        Dim swa As Integer ', i As Integer
        
           Plin ""
        Plin Space(25) & varObserv
        
        
        For i = 1 To snro_linea_trailer - sFilaMAx
            Plin "     "
        Next
    
        
        'IMPRIME_REFERENCIA_02
        Plin Chr(12)
        
        If vExcede = 1 Then MsgBox "La cantidad de detalle excede el tamaño de la Guia, algunos datos no se imprimieron, verifique", vbInformation, Me.Caption

End If


End Sub

Sub IMPRIME_REFERENCIA_02()
Dim RsObs As New Recordset
Dim Rs As New ADODB.Recordset
Dim strCadena As String, iCount As Integer, sNumPed As String, sLineaFin As String
Dim sMarca As String, sPeso As String, sMTC As String, sNoLic As String, _
    sNomTransp As String, sPlaca As String
Dim varObserv As String



    RsObs.Open "select ser_ordcomp,cod_ordcomp,des_protex,des_claordcomp from lg_ordComp a,tx_procesos b,lg_claordcomp c where a.cod_protex *=b.cod_protex and a.cod_claordcomp=c.cod_claordcomp and ser_ordcomp='" & Ser_OrdComp & "' and cod_ordcomp='" & Cod_OrdComp & "'", cConnect, adOpenStatic, adLockReadOnly
    
    varObserv = Trim(DevuelveCampo("select observaciones from Lg_MoviStk where cod_almacen='" & CodAlmacen & "' and num_movstk='" & NumMovStk & "'", cConnect))
    
     
    If RsObs.RecordCount Or Len(varObserv) > 1 Then

        strCadena = Space(2) & varObserv
        Plin strCadena
        strCadena = "-----------------------------------------------------------------------------------------------------------------------------------"
        Plin strCadena
        
        Plin strCadena
        If DevuelveCampo("select tip_fabrica from tg_control", cConnect) = "2" Then
             strCadena = "      O/P : " & DevuelveCampo("EXEC SM_BUSCA_OPS_OC '" & Ser_OrdComp & "','" & Cod_OrdComp & "',''", cConnect) & Space(2)
             Plin strCadena
        End If
        
        Dim sTransportista As String, sRUC_Transportista As String, i As Integer, snro_linea_trailer As Integer, snro_linea_razon_social_transportista As Integer
    
    'only no va esto
        'sTransportista = Mid(Trim(TxtTransportista.Text), 1, 30)
        'sRUC_Transportista = Mid(Trim(TxtRuc.Text), 1, 40)
    
        'snro_linea_trailer = DevuelveCampo("select nro_linea_trailer from Cf_Coordenadas_Emision_Guia where tipo='1'", cConnect)
        'snro_linea_razon_social_transportista = DevuelveCampo("select nro_linea_razon_social_transportista from Cf_Coordenadas_Emision_Guia where tipo='1'", cConnect)
    
    
        'For i = 1 To snro_linea_razon_social_transportista - snro_linea_trailer - 3
        '    Plin "     "
        'Next
            
        'Plin Space(1) & sTransportista
        'Plin Space(5) & sRUC_Transportista

        
        
    End If
    


    


    
        


End Sub


