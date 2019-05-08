VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEstCliTemp_Estampados 
   Caption         =   "Estampados del Estilo Cliente"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDetalles 
      Caption         =   "Detalles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   0
      TabIndex        =   10
      Top             =   2330
      Width           =   5940
      Begin VB.TextBox txtDes_Estampado 
         Height          =   300
         Left            =   1185
         TabIndex        =   1
         Top             =   630
         Width           =   3360
      End
      Begin VB.TextBox txtCod_Estampado 
         Height          =   300
         Left            =   1200
         TabIndex        =   0
         Top             =   225
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Estilo Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   645
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   290
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   495
      Left            =   105
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1470
         Picture         =   "frmEstCliTemp_Estampados.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ultimo"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   -30
         Picture         =   "frmEstCliTemp_Estampados.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Primero"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   975
         Picture         =   "frmEstCliTemp_Estampados.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Siguiente"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   465
         Picture         =   "frmEstCliTemp_Estampados.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Anterior"
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Frame FraListado 
      Caption         =   "Listado"
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
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5940
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3413
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Cod_Estampado"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Des_Estampado"
            Caption         =   "Descripción"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4020.095
            EndProperty
         EndProperty
      End
   End
   Begin Mantenimientos.MantFunc MFEstCli 
      Height          =   540
      Left            =   2310
      TabIndex        =   2
      Top             =   3570
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmEstCliTemp_Estampados.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "frmEstCliTemp_Estampados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Codigo, Descripcion As String
Dim Opcion As Integer
Dim sTipo As String
Dim strSQL As String
Dim Rs_Lista As ADODB.Recordset
Public varCod_Cliente, varCod_TemCli As String

Dim TipoFab As Integer

Private Sub cmdFirst_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MoveFirst
    End If
End Sub

Private Sub cmdLast_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveLast
    End If
End Sub

Private Sub cmdNext_Click()
    If Not Rs_Lista.EOF Then
        Rs_Lista.MoveNext
        If Rs_Lista.EOF Then
            Rs_Lista.MoveLast
        End If
    End If
End Sub

Private Sub cmdPrevious_Click()
    If Not Rs_Lista.BOF Then
        Rs_Lista.MovePrevious
        If Rs_Lista.BOF Then
            Rs_Lista.MoveFirst
        End If
    End If
End Sub

Public Sub RECARGA_LISTA()
    Set Rs_Lista = Nothing
    Call CARGA_LISTA
End Sub

Public Sub CARGA_LISTA()
    Dim strSQL As String
    Set Rs_Lista = New ADODB.Recordset
    Rs_Lista.ActiveConnection = cCONNECT
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    
    'Esta cadena es la que nos devolvera los items segun la seleccion establecida
    strSQL = "EXEC UP_SEL_ESTCLITEM_ESTAMPADOS '" & varCod_Cliente & "','" & varCod_TemCli & "'"
    Rs_Lista.Open strSQL
    Set DGridLista.DataSource = Rs_Lista

    If Rs_Lista.RecordCount > 0 Then
        HabilitaMant Me.MFEstCli, "ADICIONAR/MODIFICAR/ELIMINAR"
    Else
        HabilitaMant Me.MFEstCli, "ADICIONAR"
        Call LIMPIA_DATOS
    End If
End Sub

Public Sub Carga_Datos()
    If Rs_Lista.RecordCount > 0 Then
        txtCod_Estampado = Trim(Rs_Lista("Cod_Estampado").Value)
        txtDes_Estampado = Trim(Rs_Lista("Des_Estampado").Value)
    End If
End Sub
Public Sub HABILITA_DATOS()
    txtCod_Estampado.Enabled = True
    txtDes_Estampado.Enabled = True
End Sub
Public Sub DESABILITA_DATOS()
    txtCod_Estampado.Enabled = False
    txtDes_Estampado.Enabled = False
End Sub

Public Sub LIMPIA_DATOS()
    txtCod_Estampado.Text = ""
    txtDes_Estampado.Text = ""
End Sub

Public Sub CARGA_COMBOS()
End Sub

Public Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
        If Trim(txtCod_Estampado.Text) = "" Then
            Call MsgBox("Sirvase ingresar un codigo de Estampado", vbExclamation)
            VALIDA_DATOS = False
            Exit Function
        End If
        If Trim(txtDes_Estampado.Text) = "" Then
            Call MsgBox("Descripción de Estampado no puede estar vacia. Sirvase verificar", vbExclamation)
            VALIDA_DATOS = False
            Exit Function
        End If
        
End Function

Public Sub SALVAR_DATOS()
    Dim con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    con.ConnectionString = cCONNECT
    con.Open
          
    
     'Esta es la sentencia que realizara el salvado de datos
     strSQL = "UP_MAN_ESTCLITEM_ESTAMPADOS " & _
     sTipo & ",'" & _
     varCod_Cliente & "','" & _
     varCod_TemCli & "','" & _
     Trim(txtCod_Estampado.Text) & "','" & _
     Trim(txtDes_Estampado.Text) & "'"
     
     con.Execute strSQL
                
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
    Informa "", amensaje
    Call DESABILITA_DATOS
    Call LIMPIA_DATOS

    Exit Sub
Salvar_DatosErr:
    con.RollbackTrans
    Set con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Sub

Public Sub ELIMINAR_DATOS()
    Dim con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
    
    strSQL = "SELECT Cod_Estampado FROM TG_COLCLITEM WHERE Cod_Cliente='" & varCod_Cliente & "' AND Cod_TemCli='" & varCod_TemCli & "' AND Cod_Estampado = '" & txtCod_Estampado & "'"

    If DevuelveCampo(strSQL, cCONNECT) <> "" Then
        MsgBox ("No se puede eliminar el Registro por que posee registros relacionados")
        Exit Sub
    End If
    
    con.ConnectionString = cCONNECT
    con.Open
    
    'Esta es la sentencia que realizara el salvado de datos
    strSQL = "UP_MAN_ESTCLITEM_ESTAMPADOS " & _
    sTipo & ",'" & _
    varCod_Cliente & "','" & _
    varCod_TemCli & "','" & _
    Trim(txtCod_Estampado.Text) & "','" & _
    Trim(txtDes_Estampado.Text) & "'"
     con.Execute strSQL
        
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje

    LIMPIA_DATOS
Exit Sub
Eliminar_DatosErr:
    con.RollbackTrans
    Set con = Nothing
    ErrorHandler Err, "Eliminar_Datos"
End Sub

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Rs_Lista.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Lista.BOF And Not Rs_Lista.EOF Then
        Call Carga_Datos
    End If
End Sub

Private Sub Form_Load()
On Error GoTo hand
    Call FormSet(Me)
    Call DESABILITA_DATOS
    Call FormateaGrid(DGridLista)
    'MFEstCli.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
Exit Sub
hand:
ErrorHandler Err, "Form_Load()"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Rs_Lista.RecordCount > 0 Then
        'With oParent
        '   .Valor = DGridLista.Columns(0)
        'End With
    End If
End Sub


Private Sub MFEstCli_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIA_DATOS
            HABILITA_DATOS
            HabilitaMant Me.MFEstCli, "GRABAR/DESHACER"
            DGridLista.Enabled = False
            txtCod_Estampado.SetFocus
            
        Case "MODIFICAR"
            sTipo = "U"
            HABILITA_DATOS
            txtCod_Estampado.Enabled = False
            txtDes_Estampado.SetFocus
            'Aqui guardamos en esta varialbe temporal el codigo antiguo del estilo
            HabilitaMant Me.MFEstCli, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "ELIMINAR"
            sTipo = "D"
            Eliminar = MsgBox("Usted desea eliminar el registro seleccionado", vbExclamation + vbYesNo)
            If Eliminar = vbYes Then
                Call ELIMINAR_DATOS
                Call RECARGA_LISTA
            Else
                Exit Sub
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                SALVAR_DATOS
                sTipo = ""
                Call RECARGA_LISTA
                HabilitaMant Me.MFEstCli, "ADICIONAR/MODIFICAR/ELIMINAR"
                DGridLista.Enabled = True
                Call Carga_Datos
            End If
        Case "DESHACER"
            DESABILITA_DATOS
            sTipo = ""
            LIMPIA_DATOS
            Call Carga_Datos
            HabilitaMant Me.MFEstCli, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
        Case "SALIR"
            sTipo = ""
            Unload Me
    End Select
End Sub

Private Sub txtCod_Estampado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call AVANZA(13)
    End If
End Sub

Private Sub txtDes_Estampado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Call AVANZA(13)
    End If
End Sub



