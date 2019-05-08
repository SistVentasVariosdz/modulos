VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form frmAsignaGuias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignacion de Guias"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmAsignaGuias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1920
      TabIndex        =   6
      Top             =   1920
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmAsignaGuias.frx":0442
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   6255
      Begin VB.TextBox TxtNom_Usuario 
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   840
         Width           =   3375
      End
      Begin VB.CommandButton CmdUsuario 
         Caption         =   "..."
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox TxtCod_Usuario 
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxtCod_Fin 
         Height          =   375
         Left            =   4680
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtSer_Fin 
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtCod_Ini 
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtSer_Ini 
         Height          =   375
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Usuario"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmAsignaGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo As String
Public Descripcion As String
Dim Strsql As String

Private Sub CARGA_DATOS()
Dim Rs As New ADODB.Recordset

TxtSer_Ini = DevuelveCampo("select serie_guia from tg_control", cCONNECT)
TxtSer_Fin = TxtSer_Ini

Strsql = "Select isnull(max(ser_ordcomp+cod_ordcomp),'" & Trim(TxtSer_Ini.Text) & "' + '000000') + 1 as Numero from lg_userguia where ser_ordcomp='" & TxtSer_Ini.Text & "'"
Rs.Open Strsql, cCONNECT, adOpenStatic
If Rs.RecordCount Then
    txtCod_Ini = Right(Rs("Numero"), 6)
    TxtCod_Fin = txtCod_Ini
End If

Set Rs = Nothing
End Sub

Private Sub SALVAR_DATOS()
Dim Rs As ADODB.Recordset
'Dim RsBusca As New ADODB.Recordset
'Dim sMessage As Variant
Dim i As Integer
On Error GoTo hand

Set Rs = New ADODB.Recordset
Rs.ActiveConnection = cCONNECT
Rs.CursorLocation = adUseClient
If Val(txtCod_Ini.Text) = Val(TxtCod_Fin.Text) Then
    Rs.Open "exec UP_MAN_USERGUIA 'I','" & TxtSer_Ini & "','" & Format(TxtCod_Fin.Text, "000000") & "','" & Trim(TxtCod_Usuario.Text) & "',''"
Else
    For i = Val(txtCod_Ini.Text) To Val(TxtCod_Fin.Text)
            Rs.Open "exec UP_MAN_USERGUIA 'I','" & TxtSer_Ini & "','" & Format(i, "000000") & "','" & Trim(TxtCod_Usuario.Text) & "',''"
    Next
End If

MsgBox "Los datos se ingresaron satisfactoriamente", vbInformation, Me.Caption

Exit Sub
hand:
    ErrorHandler Err, "SALVAR_DATOS"
    Set Rs = Nothing
End Sub

Private Sub CmdUsuario_Click()
Set frmBusqGeneral2.oParent = Me
frmBusqGeneral2.sQuery = "select cod_usuario as Codigo ,nom_usuario as Descripcion from seguridad..seg_usuarios order by 2"
frmBusqGeneral2.CARGAR_DATOS
frmBusqGeneral2.Show 1
If Codigo <> "" Then
    Me.TxtCod_Usuario = Codigo
    Me.TxtNom_Usuario = Descripcion
End If
    Codigo = ""
    Descripcion = ""
End Sub

Private Sub Form_Load()
    CARGA_DATOS
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
    Select Case ActionName
        Case "ACEPTAR"
            If VALIDA_DATOS Then
                SALVAR_DATOS
            End If
        Case "CANCELAR"
            Unload Me
    End Select
End Sub

Private Sub TxtCod_Fin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtCod_Fin = Format(TxtCod_Fin, "000000")
    If Trim(TxtCod_Fin) <> "" Then
        TxtCod_Usuario.SetFocus
    End If
Else
    Call SoloNumeros(TxtCod_Fin, KeyAscii, False, 0, 6)
End If

End Sub

Private Sub txtCod_Ini_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCod_Ini = Format(txtCod_Ini, "000000")
    If Trim(txtCod_Ini) <> "" Then
        TxtCod_Fin.SetFocus
    End If
Else
    Call SoloNumeros(txtCod_Ini, KeyAscii, False, 0, 6)
End If
End Sub

Private Sub TxtCod_Usuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ExisteCampo("cod_usuario", "seg_usuarios", TxtCod_Usuario.Text, cSEGURIDAD) Then
        TxtNom_Usuario.Text = DevuelveCampo("select nom_usuario from seg_usuarios where cod_usuario='" & TxtCod_Usuario.Text & "'", cSEGURIDAD)
    Else
        MsgBox "Codigo no existe", Me.Caption
    End If

End If
End Sub

Private Function VALIDA_DATOS() As Boolean
    If Trim(txtCod_Ini) = "" Then
        MsgBox "El Numero inicial no puede ser vacio", vbInformation, Me.Caption
        txtCod_Ini.SetFocus
        VALIDA_DATOS = False
        Exit Function
    End If
    If Trim(TxtCod_Fin) = "" Then
        MsgBox "El Numero Final no puede ser vacio", vbInformation, Me.Caption
        TxtCod_Fin.SetFocus
        VALIDA_DATOS = False
        Exit Function
    End If
    If Val(txtCod_Ini.Text) > Val(TxtCod_Fin.Text) Then
        MsgBox "El Numero inicial no puede ser mayor que el final, verifique", vbInformation, Me.Caption
        txtCod_Ini.SetFocus
        VALIDA_DATOS = False
        Exit Function
    End If
    If Trim(TxtCod_Usuario.Text) = "" Then
        MsgBox "Ingrese el usuario", vbInformation, Me.Caption
        TxtCod_Usuario.SetFocus
        VALIDA_DATOS = False
        Exit Function
    End If
    
    VALIDA_DATOS = True
End Function

Private Sub TxtSer_Fin_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(TxtSer_Fin, KeyAscii, False, 0, 3)
End Sub

Private Sub TxtSer_Ini_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Rs As New ADODB.Recordset
    If KeyCode = 13 Then
        TxtSer_Fin.Text = Trim(TxtSer_Ini.Text)
        Strsql = "Select isnull(max(ser_ordcomp+cod_ordcomp),'" & Trim(TxtSer_Ini.Text) & "' + '000000') + 1 as Numero from lg_userguia where ser_ordcomp='" & TxtSer_Ini.Text & "'"
        Rs.Open Strsql, cCONNECT, adOpenStatic
        If Rs.RecordCount Then
            txtCod_Ini = Right(Rs("Numero"), 6)
            TxtCod_Fin = txtCod_Ini
        End If
    End If
End Sub

Private Sub TxtSer_Ini_KeyPress(KeyAscii As Integer)
    Call SoloNumeros(TxtSer_Ini, KeyAscii, False, 0, 3)
End Sub
