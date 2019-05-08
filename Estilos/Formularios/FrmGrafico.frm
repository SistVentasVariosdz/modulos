VERSION 5.00
Begin VB.Form FrmGrafico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gráfico"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   5010
      TabIndex        =   1
      Top             =   4530
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   4275
      Left            =   240
      ScaleHeight     =   4215
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   120
      Width           =   7245
   End
   Begin VB.TextBox label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4485
      Width           =   2415
   End
End
Attribute VB_Name = "FrmGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BuscaArchivo As New FileSystemObject
Public diricono As String

Private Sub CmdAceptar_Click()
    Unload Me
End Sub

Public Sub CARGA_ICONO()
On Error GoTo AceptarErr
    If Not BuscaArchivo.FileExists(diricono) Then
        'Hoja1.Image1.Visible = True
        label1.Text = "La ruta especificada para la imagen no existe"
    Else
        'Hoja1.Image1.Visible = True
        Picture1.Picture = LoadPicture(diricono)
    End If
    Set BuscaArchivo = Nothing
    Exit Sub
AceptarErr:
    MsgBox "El gráfico seleccionado no es valido. Sirvase verificar", vbInformation, "Gráfico"
   
End Sub

