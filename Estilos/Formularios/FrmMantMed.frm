VERSION 5.00
Object = "{71ED96E1-5967-46DB-BB10-BD36D6EC1412}#1.0#0"; "Mantenimientos.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmMantMed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Medidas"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "FrmMantMed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2550
      Left            =   60
      TabIndex        =   15
      Tag             =   "Detail"
      Top             =   3315
      Width           =   6855
      Begin VB.TextBox txtPeso 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   5520
         TabIndex        =   6
         Text            =   "0"
         Top             =   2040
         Width           =   1005
      End
      Begin VB.OptionButton OptPulgadas 
         Caption         =   "Pulgadas"
         Height          =   195
         Left            =   3000
         TabIndex        =   24
         Top             =   1000
         Width           =   1215
      End
      Begin VB.OptionButton OptCentimetros 
         Caption         =   "Centimetros"
         Height          =   195
         Left            =   1320
         TabIndex        =   23
         Top             =   1000
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox TxtAgujas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5520
         TabIndex        =   5
         Text            =   "0"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox TxtPasadas 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5520
         TabIndex        =   4
         Text            =   "0"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TxtAlto 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtLargo 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtCod_Comb 
         BackColor       =   &H80000009&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1335
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDes_Comb 
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Peso (kgs) x Und. Cnf"
         Height          =   195
         Left            =   3720
         TabIndex        =   25
         Top             =   2145
         Width           =   1545
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Medida:"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1000
         Width           =   930
      End
      Begin VB.Label Label6 
         Caption         =   "Agujas :"
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Pasadas :"
         Height          =   255
         Left            =   4200
         TabIndex        =   20
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Alto :"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Largo :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   345
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   680
         Width           =   930
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
      Left            =   60
      TabIndex        =   13
      Tag             =   "List"
      Top             =   30
      Width           =   6855
      Begin MSDataGridLib.DataGrid DGridLista 
         Height          =   2925
         Left            =   240
         TabIndex        =   14
         Top             =   255
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5159
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   645
      TabIndex        =   8
      Top             =   5910
      Width           =   1965
      Begin VB.CommandButton cmdPrevious 
         Height          =   495
         Left            =   480
         Picture         =   "FrmMantMed.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Anterior"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   495
         Left            =   960
         Picture         =   "FrmMantMed.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Siguiente"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   495
         Left            =   15
         Picture         =   "FrmMantMed.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Primero"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Height          =   495
         Left            =   1440
         Picture         =   "FrmMantMed.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ultimo"
         Top             =   120
         Width           =   495
      End
   End
   Begin Mantenimientos.MantFunc MantFunc1 
      Height          =   540
      Left            =   2805
      TabIndex        =   7
      Top             =   5985
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   953
      Custom          =   $"FrmMantMed.frx":08D2
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
   End
End
Attribute VB_Name = "FrmMantMed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cod_Item As String
Public Tipo_Item As String

Dim Cod_Medida As String
Dim Estado As String
Dim Reg As New ADODB.Recordset
Sub Datos(Accion As String, TipoItem As String, Optional EsAccion As Boolean = False)
On Error GoTo hand
Set Reg = Nothing

If Trim(txtPeso) = "" Then txtPeso.Text = 0

Reg.CursorLocation = adUseClient
Reg.Open " UP_Medidas '" & Accion & "','" & TipoItem & "','" & Cod_Item & "','" & Cod_Medida & "','" & Me.txtDes_Comb & "','" & Trim(txtLargo.Text) & "','" & Trim(TxtAlto.Text) & "'," & IIf(Trim(TxtPasadas.Text) = "", 0, TxtPasadas.Text) & "," & IIf(Trim(TxtAgujas.Text) = "", 0, TxtAgujas.Text) & ",'" & IIf(OptCentimetros, "C", "P") & "'," & CDbl(txtPeso.Text), cCONNECT

If Not EsAccion Then
    Set Me.DGridLista.DataSource = Reg

End If

Exit Sub
hand:
ErrorHandler Err, "Datos"
Set Reg = Nothing
End Sub

Sub Habilita()
'txtCod_Comb.Enabled = True
txtDes_Comb.Enabled = True
txtLargo.Enabled = True
TxtAlto.Enabled = True
TxtPasadas.Enabled = True
TxtAgujas.Enabled = True
OptCentimetros.Enabled = True
OptPulgadas.Enabled = True
txtPeso.Enabled = True
End Sub

Sub DesHabilita()
txtCod_Comb.Enabled = False
txtDes_Comb.Enabled = False
TxtAlto.Enabled = False
txtLargo.Enabled = False
TxtPasadas.Enabled = False
TxtAgujas.Enabled = False
OptCentimetros.Enabled = False
OptPulgadas.Enabled = False
txtPeso.Enabled = False
End Sub

Sub Limpia()
Me.txtCod_Comb = ""
Me.txtDes_Comb = ""
TxtAlto.Text = ""
txtLargo.Text = ""
TxtPasadas.Text = 0
TxtAgujas.Text = 0
OptCentimetros.Value = True
txtPeso.Text = 0
End Sub

Private Sub DGridLista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

If Not Reg.EOF And Not Reg.BOF Then
    Me.txtCod_Comb = Reg(0)
    Me.txtDes_Comb = Reg(1)
    Cod_Medida = Reg(0)
    If Tipo_Item = "T" Then
        txtLargo.Text = Reg("Largo")
        TxtAlto.Text = Reg("Alto")
        TxtPasadas.Text = Reg("Pasadas")
        TxtAgujas.Text = Reg("Agujas")
        txtPeso.Text = Reg("Peso")
        If Reg("Tipo_Medida") = "C" Then
            OptCentimetros.Value = True
        Else
            OptPulgadas.Value = True
        End If
    End If
End If
End Sub


Private Sub Form_Load()
Datos "V", Tipo_Item
FormateaGrid Me.DGridLista
Limpia
Datos "V", Tipo_Item
DesHabilita
If Tipo_Item = "I" Then
    Me.Caption = Me.Caption & " " & "de Items"
    txtLargo.Visible = False
    TxtAlto.Visible = False
    Label2.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    TxtPasadas.Visible = False
    TxtAgujas.Visible = False
    OptCentimetros.Visible = False
    OptPulgadas.Value = False
Else
    Me.Caption = Me.Caption & " " & "de Telas"
    txtLargo.Visible = True
    TxtAlto.Visible = True
    Label2.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    TxtPasadas.Visible = True
    TxtAgujas.Visible = True
    OptCentimetros.Visible = True
    OptPulgadas.Value = True
    txtPeso.Visible = True
End If
MantFunc1.FunctionsUser = get_botones1(Me, vper, vemp, Me.Name)
End Sub


Private Sub MantFunc1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Cod_Medida = Me.txtCod_Comb
Select Case ActionName
    Case "ADICIONAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Habilita
        Me.txtCod_Comb.Enabled = True
        Estado = "NUEVO"
        Me.txtCod_Comb.SetFocus
        Limpia
    Case "MODIFICAR"
        HabilitaMant Me.MantFunc1, "GRABAR/DESHACER"
        Estado = "MODIFICAR"
        Habilita
        ''txtDes_Comb.Enabled = True
        txtDes_Comb.SetFocus
    Case "ELIMINAR"
        Datos "B", Tipo_Item, True
            Limpia
            DesHabilita
            Datos "V", Tipo_Item
    Case "GRABAR"
'            If Trim(TxtDescripcion) = "" Then MsgBox "Llene la descripcion", vbInformation: Exit Sub
            
            If Estado = "NUEVO" Then
                Datos "I", Tipo_Item, True
            Else
                Datos "A", Tipo_Item, True
            End If
            Limpia
            DesHabilita
            HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
            Datos "V", Tipo_Item
    Case "DESHACER"
        HabilitaMant Me.MantFunc1, "ADICIONAR/MODIFICAR/ELIMINAR"
        Datos "V", Tipo_Item
    Case "SALIR"
        Unload Me
End Select
End Sub

Private Sub TxtAgujas_GotFocus()
SelectionText TxtAgujas
End Sub

Private Sub TxtAgujas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtAgujas, KeyAscii, False)
End If
End Sub

Private Sub TxtAlto_GotFocus()
SelectionText TxtAlto
End Sub

Private Sub TxtAlto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtCod_Comb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtDes_Comb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtLargo_GotFocus()
SelectionText txtLargo
End Sub

Private Sub txtLargo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub TxtPasadas_GotFocus()
SelectionText TxtPasadas
End Sub

Private Sub TxtPasadas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
Else
    Call SoloNumeros(TxtPasadas, KeyAscii, False)
End If
End Sub

Private Sub txtPeso_GotFocus()
SelectionText txtPeso
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
Else
    Call SoloNumeros(txtPeso, KeyAscii, True, 5)
End If
End Sub
