VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmMantItemComb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Combinaciones"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Hilados"
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Height          =   1800
      Left            =   1680
      TabIndex        =   20
      Top             =   1080
      Visible         =   0   'False
      Width           =   3630
      Begin VB.ComboBox cboEsta 
         Height          =   315
         Left            =   1065
         TabIndex        =   21
         Top             =   405
         Width           =   2385
      End
      Begin FunctionsButtons.FunctButt FunctButt3 
         Height          =   510
         Left            =   690
         TabIndex        =   22
         Top             =   1080
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   900
         Custom          =   "0~0~ACEPT~True~True~Aceptar~0~0~1~~0~False~False~Aceptar~~1~0~CANCE~True~True~Cancelar~0~0~2~~0~False~False~Cancelar~"
         Orientacion     =   0
         Style           =   0
         Language        =   0
         TypeImageList   =   0
         ControlWidth    =   1155
         ControlHeigth   =   490
         ControlSeparator=   110
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Estado"
         Height          =   345
         Left            =   375
         TabIndex        =   23
         Top             =   420
         Width           =   765
      End
   End
   Begin FunctionsButtons.FunctButt FunctDetalles 
      Height          =   495
      Left            =   2040
      TabIndex        =   15
      Top             =   6120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      Custom          =   "0~0~DETALLES~True~True~&Detalles~0~0~1~~0~False~False~&Detalles~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   5445
      Width           =   1965
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "frmMantItemComb.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "frmMantItemComb.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "frmMantItemComb.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "frmMantItemComb.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame Fralista 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   90
      TabIndex        =   4
      Tag             =   "List"
      Top             =   105
      Width           =   6855
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2925
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5159
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Cod_Comb"
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
            DataField       =   "Des_Comb"
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
         BeginProperty Column02 
            DataField       =   "Observaciones"
            Caption         =   "Observaciones"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Precio_Cotizado"
            Caption         =   "Precio_Cotizado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "flg_status"
            Caption         =   "flg_status"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "des_status"
            Caption         =   "des_status"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4694.74
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fradetalle 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Tag             =   "Detail"
      Top             =   3470
      Width           =   6855
      Begin VB.TextBox txtstatus 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4560
         TabIndex        =   26
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtPrecioCotizado 
         Height          =   300
         Left            =   1560
         TabIndex        =   19
         Text            =   "0"
         Top             =   1490
         Width           =   1575
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1035
         Width           =   4935
      End
      Begin VB.TextBox txtDes_Comb 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   3
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txtDes_Item 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtCod_Comb 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Status :"
         Height          =   255
         Left            =   3960
         TabIndex        =   25
         Top             =   1500
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Precio Cotizado  $:"
         Height          =   375
         Left            =   165
         TabIndex        =   18
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Detalle Pinturas :"
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   1110
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   675
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Item :"
         Height          =   195
         Left            =   3600
         TabIndex        =   13
         Top             =   345
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   345
         Width           =   585
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2160
      TabIndex        =   6
      Top             =   5520
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"frmMantItemComb.frx":05C8
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   495
      Left            =   5880
      TabIndex        =   24
      Top             =   5520
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      Custom          =   "0~0~ESTADO~True~True~&Cambiar Estado~0~0~1~~0~False~False~&Cambiar Estado~"
      Orientacion     =   1
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
End
Attribute VB_Name = "frmMantItemComb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oParent As Object
Public Codigo, Descripcion As String
Dim sTipo As String
Dim Rs_Grid As New ADODB.Recordset
Dim StrSQL As String
Public Codigo_item As String

Private Sub cmdFirst_Click()
    If Not Rs_Grid.BOF Then
        Rs_Grid.MoveFirst
    End If
End Sub
Private Sub cmdLast_Click()
    If Not Rs_Grid.EOF Then
        Rs_Grid.MoveLast
    End If
End Sub
Private Sub cmdNext_Click()
    If Not Rs_Grid.EOF Then
        Rs_Grid.MoveNext
    End If
End Sub
Private Sub cmdPrevious_Click()
    If Not Rs_Grid.BOF Then
        Rs_Grid.MovePrevious
    End If
End Sub
Public Sub CargaCombos()
    Dim StrSQL As String
        'Combo Flag Estatus
'    StrSQL = "SELECT Des_Status + space(100) + Flg_Status  FROM LG_Status_Servicios"
'    Call LlenaCombo(cboFlg_Status, StrSQL, cCONNECT)
    
    cboEsta.Clear
    cboEsta.AddItem ("Aprobado")
    cboEsta.AddItem ("Aprobado Parcial")
    cboEsta.ListIndex = 0
  End Sub
  
Private Sub Form_Load()
    Call FormSet(Me)
    Call CargaCombos
    FormateaGrid Me.DGridLista
    HabilitaMant Me.MantFunc1, ""
    DESHABILITA_DATOS
    MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
    'FunctDetalles.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub

Private Sub DGridLista_KeyDown(KeyCode As Integer, Shift As Integer)
    AVANZA (KeyCode)
End Sub
Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Rs_Grid.State <> 1 Then
        Exit Sub
    End If
    If Not Rs_Grid.EOF And Not Rs_Grid.BOF Then
        Call Carga_Datos
        DESHABILITA_DATOS
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Rs_Grid = Nothing
End Sub



Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)

On Error GoTo AceptarErr

    Dim StrSQL As String
    Dim vericono As Integer
    Select Case ActionName
 
    Case "ESTADO"
    
        'If DGridLista.RowCount > 0 And (Not DGridLista.IsGroupItem(DGridLista.Row)) Then
               
                  
                   
                   If Rs_Grid("flg_status") <> "P" Then
                     If MsgBox("Esta seguro de cambiar de estado", vbInformation + vbYesNo, "AVISO") = vbYes Then
                
                        Frame4.Visible = False
                        StrSQL = " exec ES_LG_ItemComb_Cambia_Status_Servicios '" & Codigo_item & "','" & Rs_Grid("Cod_Comb").Value & "','P' "
                        Call ExecuteSQL(cCONNECT, StrSQL)
                        Call CARGA_GRID
                        
                      End If
                   Else
                        DGridLista.Enabled = False
                        Frame4.Visible = True
                   End If
               
                
'        Else
'                MsgBox "Debe seleccionar un item para acceder a esta opcion", vbInformation
'        End If
        End Select
Exit Sub
AceptarErr:
    ErrorHandler Err, "Aceptar"
    Screen.MousePointer = vbNormal
End Sub

Private Sub FunctButt3_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Select Case ActionName
Case "ACEPT"

   If MsgBox("Esta seguro de cambiar de estado", vbInformation + vbYesNo, "AVISO") = vbYes Then
                
        Dim sEstado As String
        Dim StrSQL As String
        If cboEsta.Text = "Aprobado" Then
        sEstado = "A"
        End If
        
        If cboEsta.Text = "Aprobado Parcial" Then
        sEstado = "B"
        End If
      StrSQL = " exec ES_LG_ItemComb_Cambia_Status_Servicios '" & Codigo_item & "','" & Rs_Grid("Cod_Comb").Value & "','" & sEstado & "'"
      Call ExecuteSQL(cCONNECT, StrSQL)
      DGridLista.Enabled = True
      Call CARGA_GRID
      Frame4.Visible = False
      End If
      
Case "CANCE"
    DGridLista.Enabled = True
    Frame4.Visible = False
End Select
End Sub

Private Sub FunctDetalles_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "DETALLES"
            If Not Rs_Grid.EOF Then
                Load frmMantItemCombDet
                frmMantItemCombDet.Caption = "DETALLE DE COMBINACION:" & Rs_Grid("Cod_Comb") & " " & Rs_Grid("Des_Comb")
                frmMantItemCombDet.Codigo_item = Rs_Grid("Cod_Item")
                frmMantItemCombDet.Codigo_Comb = Rs_Grid("Cod_Comb")
                frmMantItemCombDet.CargaCombos
                frmMantItemCombDet.CARGA_GRID
                frmMantItemCombDet.Show 1
            Else
                MsgBox ("Debe seleccionar una Tela para acceder a esta opcion")
            End If
    End Select
End Sub

Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Dim Eliminar As Integer
    Select Case ActionName
        Case "ADICIONAR"
            sTipo = "I"
            LIMPIAR_DATOS
            HABILITA_DATOS
            txtDes_Comb.SetFocus
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "MODIFICAR"
            sTipo = "U"
            HABILITA_DATOS
            HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
            DGridLista.Enabled = False
        Case "ELIMINAR"
            sTipo = "D"
            If VALIDA_DATOS Then
                Eliminar = MsgBox("Esta seguro de eliminar el registro", vbInformation + vbYesNo)
                If Eliminar = vbYes Then
                    ELIMINAR_DATOS
                    RECARGAR_DATOS
                    sTipo = ""
                End If
            End If
        Case "GRABAR"
            If VALIDA_DATOS Then
                If SALVAR_DATOS Then
                    RECARGAR_DATOS
                    HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
                    DGridLista.Enabled = True
                    If sTipo = "I" Then
                        MantFunc1_ActionClick 0, 0, "ADICIONAR"
                    End If

                End If
            End If
        Case "DESHACER"
            LIMPIAR_DATOS
            RECARGAR_DATOS
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            DGridLista.Enabled = True
            sTipo = ""
         Case "SALIR"
            Unload Me
    End Select
End Sub

Sub LIMPIAR_DATOS()
    txtCod_Comb.Text = ""
    txtDes_Comb.Text = ""
    txtObservaciones.Text = ""
    txtPrecioCotizado.Text = 0
End Sub


Sub Carga_Datos()
    If Not Rs_Grid.EOF Then
        txtCod_Comb.Text = Trim(Rs_Grid("Cod_Comb").Value)
        txtDes_Comb.Text = Trim(Rs_Grid("Des_Comb").Value)
        txtObservaciones.Text = Trim(Rs_Grid("Observaciones").Value)
        txtPrecioCotizado.Text = Trim(Rs_Grid("Precio_Cotizado").Value)
        txtstatus.Text = Trim(Rs_Grid("des_status").Value)

    End If

End Sub

Sub RECARGAR_DATOS()
    
    Rs_Grid.Close
    CARGA_GRID
    
End Sub

Public Sub CARGA_GRID()
    Dim StrSQL As String
    Dim xRow As Long
    
    Set Rs_Grid = New ADODB.Recordset
    Rs_Grid.ActiveConnection = cCONNECT
    Rs_Grid.CursorType = adOpenStatic
    Rs_Grid.CursorLocation = adUseClient
    Rs_Grid.LockType = adLockReadOnly
    
    xRow = DGridLista.Row
    
    'Esta cadena es para devolver el Codigo de Cliente
    StrSQL = "EXEC UP_SEL_ITEMCOMB '" & Codigo_item & "'"
    
    Rs_Grid.Open StrSQL
    Set DGridLista.DataSource = Rs_Grid
    DGridLista.Refresh
    
    If xRow > 0 And xRow <= Rs_Grid.RecordCount Then
        DGridLista.Row = xRow
    End If

    If Rs_Grid.RecordCount > 0 Then
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Call Carga_Datos
    Else
        HabilitaMant Me.MantFunc1, "ADICIONAR"
        Call LIMPIAR_DATOS
    End If
    


    MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
  
  

    
End Sub

Function VALIDA_DATOS() As Boolean
    VALIDA_DATOS = True
    If sTipo <> "D" Then
        If Trim(txtDes_Comb.Text) = "" Then
            Call MsgBox("La descripción no puede estar vacia. Sirvase verificar", vbCritical)
            VALIDA_DATOS = False
            Exit Function
        End If
    Else
        StrSQL = "SELECT COUNT(Num_Secuencia) FROM LG_ITEMCOMBDET WHERE Cod_Item='" & Rs_Grid("Cod_Item").Value & "' AND Cod_Comb='" & Rs_Grid("Cod_Comb").Value & "'"
        If DevuelveCampo(StrSQL, cCONNECT) > 0 Then
            Call MsgBox("No se puede eliminar este Regitro por que posee registros relacionados", vbCritical)
            VALIDA_DATOS = False
            Exit Function
        End If
    End If
End Function

Sub HABILITA_DATOS()
    txtDes_Comb.Enabled = True
    txtObservaciones.Enabled = True
    txtPrecioCotizado.Enabled = True
End Sub

Sub DESHABILITA_DATOS()
    txtDes_Comb.Enabled = False
    txtObservaciones.Enabled = False
    txtPrecioCotizado.Enabled = False
End Sub

Function SALVAR_DATOS() As Boolean
    Dim Con As New ADODB.Connection
    On Error GoTo Salvar_DatosErr
    Dim StrSQL As String
    
    Con.ConnectionString = cCONNECT
    Con.Open
    
        Con.BeginTrans

        
        StrSQL = "EXEC UP_MAN_ITEMCOMB '" & _
        sTipo & "','" & _
        Codigo_item & "','" & _
        txtCod_Comb.Text & "','" & _
        txtDes_Comb.Text & "','" & _
        txtObservaciones & "','" & _
        txtPrecioCotizado & "'"
       
        Con.Execute StrSQL

        Con.CommitTrans
        Dim amensaje As New clsMessages
        amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_save
        Informa "", amensaje
        
        SALVAR_DATOS = True
        
    Exit Function
Salvar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Salvar_Datos"
End Function
Sub ELIMINAR_DATOS()
 Dim Con As New ADODB.Connection
    On Error GoTo Eliminar_DatosErr
   
    Con.ConnectionString = cCONNECT
    Con.Open
    Con.BeginTrans
       
        StrSQL = "EXEC UP_MAN_ITEMCOMB '" & _
        sTipo & "','" & _
        Codigo_item & "','" & _
        Rs_Grid("Cod_Comb").Value & "','" & _
        Rs_Grid("Des_Comb").Value & "','" & _
        Rs_Grid("Observaciones").Value & "','" & _
        txtPrecioCotizado & "'"
         
        Con.Execute StrSQL
        
    Con.CommitTrans
    
    Dim amensaje As New clsMessages
    amensaje.Codigo = CodeMsg.kMeSsaGe_INF_DATA_DELETE
    Informa "", amensaje
    
Exit Sub
Eliminar_DatosErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Eliminar_Datos"

End Sub

Private Sub txtDes_Comb_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtObservaciones.SetFocus
    End If
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       txtPrecioCotizado.SetFocus
    End If
End Sub
Private Sub txtMotSolicitud_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   MantFunc1.SetFocus
 End If
End Sub
