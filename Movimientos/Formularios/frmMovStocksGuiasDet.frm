VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMovStocksGuiasDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2085
      TabIndex        =   9
      Top             =   5580
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMovStocksGuiasDet.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin VB.Frame fraFlechas 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   15
      TabIndex        =   11
      Top             =   5580
      Width           =   2310
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1485
         Picture         =   "frmMovStocksGuiasDet.frx":0160
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ultimo"
         Top             =   45
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   1005
         Picture         =   "frmMovStocksGuiasDet.frx":02D2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Siguiente"
         Top             =   45
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   525
         Picture         =   "frmMovStocksGuiasDet.frx":0444
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Anterior"
         Top             =   45
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   45
         Picture         =   "frmMovStocksGuiasDet.frx":05B6
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Primero"
         Top             =   45
         Width           =   495
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   60
      TabIndex        =   2
      Top             =   3330
      Width           =   5565
      Begin VB.TextBox txtNum_Secuencia 
         Height          =   300
         Left            =   45
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1815
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   1380
         Left            =   1125
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   615
         Width           =   4320
      End
      Begin VB.TextBox txtcod_unimed 
         Height          =   285
         Left            =   4155
         MaxLength       =   5
         TabIndex        =   6
         Top             =   210
         Width           =   1290
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   1125
         TabIndex        =   4
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   705
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "U.M. :"
         Height          =   195
         Left            =   3450
         TabIndex        =   5
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad :"
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   330
         Width           =   720
      End
   End
   Begin VB.Frame fraLista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3330
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   5565
      Begin GridEX20.GridEX gexLista 
         Height          =   2925
         Left            =   120
         TabIndex        =   1
         Top             =   225
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   5159
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmMovStocksGuiasDet.frx":0728
         Column(2)       =   "frmMovStocksGuiasDet.frx":07F0
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmMovStocksGuiasDet.frx":0894
         FormatStyle(2)  =   "frmMovStocksGuiasDet.frx":09CC
         FormatStyle(3)  =   "frmMovStocksGuiasDet.frx":0A7C
         FormatStyle(4)  =   "frmMovStocksGuiasDet.frx":0B30
         FormatStyle(5)  =   "frmMovStocksGuiasDet.frx":0C08
         FormatStyle(6)  =   "frmMovStocksGuiasDet.frx":0CC0
         ImageCount      =   0
         PrinterProperties=   "frmMovStocksGuiasDet.frx":0DA0
      End
   End
End
Attribute VB_Name = "frmMovStocksGuiasDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSql As String
Dim sTipo As String

Public varCod_Almacen As String
Public varNum_MovStk As String
Public varNum_Secuencia As String
Public Codigo As String, Descripcion As String

Public varBloqueado As Boolean
Dim varBusqueda As String

Public Sub HABILITA_DATOS()
    Me.txtCantidad.Enabled = True
    Me.txtcod_unimed.Enabled = True
    Me.txtDescripcion.Enabled = True
    
    Me.fraLista.Enabled = False
    Me.fraFlechas.Enabled = False
    
End Sub

Public Sub INHABILITA_DATOS()
    Me.txtCantidad.Enabled = False
    Me.txtcod_unimed.Enabled = False
    Me.txtDescripcion.Enabled = False
    
    Me.fraLista.Enabled = True
    Me.fraFlechas.Enabled = True
End Sub

Sub CARGA_GRID()

    'Esta cadena es para devolver el Codigo de Cliente
    StrSql = "EXEC UP_SEL_LGMOVISTKGUIDET '" & Me.varCod_Almacen & "','" & Me.varNum_MovStk & "'"
    
    Set gexLista.ADORecordset = CargarRecordSetDesconectado(StrSql, cConnect)
    
    SetGeneralGridEX gexLista, 0, 1
    
    If Me.gexLista.RowCount = 0 Then
        varBusqueda = ""
    End If
    
    Call Me.gexLista.Find(3, jgexEqual, varBusqueda)
    
    Call Configurar_Grid
    
    If Me.gexLista.RowCount > 0 Then
        gexLista.Enabled = True
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call CARGA_DATOS
    Else
        gexLista.Enabled = False
        HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIA_DATOS
    End If

End Sub

Public Sub CARGA_DATOS()
    If gexLista.RowCount > 0 Then
        Me.txtNum_Secuencia = Trim(gexLista.Value(gexLista.Columns("Num_Secuencia").Index))
        Me.txtCantidad = Trim(gexLista.Value(gexLista.Columns("Cantidad").Index))
        Me.txtcod_unimed = Trim(gexLista.Value(gexLista.Columns("Cod_unimed").Index))
        Me.txtDescripcion = Trim(gexLista.Value(gexLista.Columns("Descripcion").Index))
        
    End If
End Sub

Public Sub LIMPIA_DATOS()
    Me.txtNum_Secuencia.Text = ""
    Me.txtCantidad.Text = ""
    Me.txtcod_unimed.Text = ""
    Me.txtDescripcion.Text = ""
End Sub

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If sTipo <> "U" Then
    Else
    End If
End Function

Sub SALVAR_DATOS()
    Dim Con As New ADODB.Connection
    Dim Rs As New ADODB.Recordset
    On Error GoTo Salvar_DatosErr
    Dim StrSql As String
    
    Con.ConnectionString = cConnect
    Con.Open
    
        Con.BeginTrans

        StrSql = "EXEC UP_MAN_LGMOVISTKGUIDET '" & _
        sTipo & "','" & _
        Me.varCod_Almacen & "','" & _
        Me.varNum_MovStk & "','" & _
        Trim(Me.txtNum_Secuencia.Text) & "','" & _
        Trim(Me.txtcod_unimed.Text) & "','" & _
        Trim(Me.txtCantidad.Text) & "','" & _
        Trim(Me.txtDescripcion.Text) & "'"
        
        Con.Execute StrSql
       
        Con.CommitTrans
        
'        Dim amensaje As New clsMessages
'        amensaje.Codigo = CodeMsg.kMSG_INF_DATA_SAVE
'        Informa "", amensaje
'        Mensaje kMSG_INF_DATA_SAVE
        
    Exit Sub
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
   
    Con.ConnectionString = cConnect
    Con.Open
    Con.BeginTrans
       
        StrSql = "EXEC UP_MAN_LGMOVISTKGUIDET '" & _
        sTipo & "','" & _
        Me.varCod_Almacen & "','" & _
        Me.varNum_MovStk & "','" & _
        Trim(Me.txtNum_Secuencia.Text) & "','" & _
        Trim(Me.txtcod_unimed.Text) & "','" & _
        Trim(Me.txtCantidad.Text) & "','" & _
        Trim(Me.txtDescripcion.Text) & "'"
        
        Con.Execute StrSql
    
    Con.CommitTrans
    
'    Dim amensaje As New clsMessages
'    amensaje.Codigo = CodeMsg.kMSG_INF_DATA_DELETE
'    Informa "", amensaje
'    Mensaje kMSG_INF_DATA_DELETE
    
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"

End Sub

Private Sub Form_Load()
    Call INHABILITA_DATOS
    varBloqueado = False
End Sub

Private Sub gexLista_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
    Call Me.CARGA_DATOS
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)


    Dim eliminar As Integer
    Dim vRow As Long
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            Call LIMPIA_DATOS
            Call HABILITA_DATOS
            Me.txtCantidad.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "MODIFICAR"
            sTipo = "U"
            varBusqueda = Trim(gexLista.Value(gexLista.Columns("Num_Secuencia").Index))
            Call HABILITA_DATOS
            Me.txtCantidad.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Case "ELIMINAR"
            varBusqueda = ""
            If varBloqueado Then 'And ActionName <> "SALIR" Then
                'MsgBox "No se puede acceder a ninguna opción. Los registros estan bloqueados", vbInformation, "Mensaje"
                MsgBox "No se puede eliminar. Los registros estan bloqueados", vbInformation, "Mensaje"
                Exit Sub
            End If
            eliminar = MsgBox("¿Esta usted seguro de eliminar el registro seleccionado?", vbInformation + vbYesNo, "Combinación-Detalle")
            If eliminar = vbYes Then
                sTipo = "D"
                If VALIDA_DATOS Then
                    Call ELIMINAR_DATOS
                    Call CARGA_GRID
                    sTipo = ""
                End If
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                Call SALVAR_DATOS
                Call CARGA_GRID
                If sTipo = "I" Then
                    Me.gexLista.MoveLast
                End If
                Call INHABILITA_DATOS
                HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                sTipo = ""
            End If
        Case "DESHACER"
            Call LIMPIA_DATOS
            Call CARGA_DATOS
            Call INHABILITA_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            sTipo = ""
         Case "SALIR"
            Unload Me
    End Select
End Sub

Public Sub Configurar_Grid()
    Me.gexLista.Columns("Cod_Almacen").Visible = False
    Me.gexLista.Columns("Num_MovStk").Visible = False
    Me.gexLista.Columns("Num_Secuencia").Visible = False


    Me.gexLista.Columns("cod_unimed").Caption = "U.M."
    Me.gexLista.Columns("cod_unimed").Width = 1000
    Me.gexLista.Columns("Cantidad").Caption = "Cantidad"
    Me.gexLista.Columns("Cantidad").Width = 1500
    Me.gexLista.Columns("Descripcion").Caption = "Decripción"
    Me.gexLista.Columns("Descripcion").Width = 4000
    
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtcod_unimed.SetFocus
    Else
        Call SoloNumeros(Me.txtCantidad, KeyAscii, True, 5, 9)
    End If
End Sub

Private Sub txtcod_unimed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDescripcion.SetFocus
    End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        MantFunc1.SetFocus
    End If
End Sub


Private Sub cmdFirst_Click()
    gexLista.MoveFirst
End Sub

Private Sub cmdLast_Click()
    gexLista.MoveLast
End Sub

Private Sub cmdNext_Click()
    gexLista.MoveNext
End Sub

Private Sub cmdPrevious_Click()
    gexLista.MovePrevious
End Sub

