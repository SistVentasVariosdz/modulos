VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCambiaGrafico 
   Caption         =   "GraficoItem"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox label1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   5040
      Width           =   4575
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6120
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Height          =   735
      Left            =   5520
      TabIndex        =   2
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton cmdIcono1 
      Caption         =   "..."
      Height          =   735
      Left            =   5760
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFCC&
      Caption         =   "Imagen Referencial"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.Image Image1 
         Height          =   4215
         Left            =   480
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmCambiaGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Codigo_item As String
Dim txtDes_Cliente As String
Public StrImagen1_Origen As String
Public StrImagen_cambio As String
Public ruta_imagenes As String

Dim BuscaArchivo As New FileSystemObject




Private Sub cmdIcono1_Click()

On Error GoTo ErrHandler

With cd
    .DialogTitle = "Encuentra Imagenes de Estilos Propios cliente " & txtDes_Cliente
    .Filter = "Imágenes (*.bmp;*.ico;*.jpeg;*.jpg)|*.bmp;*.ico;*.jpeg;*.jpg"
End With

cd.CancelError = True
'cd.CancelError = False
cd.ShowOpen

If Err = 0 Then
    StrImagen_cambio = cd.FileName
End If


Open_Imagen (StrImagen_cambio)

Exit Sub

Resume

ErrHandler:


StrImagen_cambio = ""
Open_Imagen (StrImagen_cambio)

End Sub

Sub Open_Imagen(StrRuta As String)

On Error GoTo AceptarErr
    If Not BuscaArchivo.FileExists(StrRuta) Then
        'Hoja1.Image1.Visible = True
        label1.Text = "La ruta especificada para la imagen no existe"
        StrImagen_cambio = ""
    Else
        Image1.Picture = LoadPicture(StrRuta)
    End If
    Set BuscaArchivo = Nothing
    Exit Sub
AceptarErr:
    MsgBox "El gráfico seleccionado no es valido. Sirvase verificar", vbInformation, "Gráfico"
   
End Sub




Private Sub Command2_Click()
 If StrImagen_cambio <> "" Then
  If Guarda_Imagen2 Then
    'MsgBox ("Se guardo la Imagen correctamente para el item: " + Codigo_item)
  End If
 End If
Unload Me
End Sub

Private Sub Form_Load()
 Open_Imagen (StrImagen1_Origen)
End Sub


Public Function Guarda_Imagen2() As Boolean
 Dim fso As New FileSystemObject
  
 
 ruta_imagenes = DevuelveCampo("Select Ruta_Iconos_Servicios from Tg_Control", cCONNECT) & fso.GetFileName(StrImagen_cambio)
 
 '& Replace(fso.GetFileName(StrImagen_cambio), Mid(fso.GetFileName(StrImagen_cambio), 1, InStr(fso.GetFileName(StrImagen_cambio), ".") - 1), Codigo_item)
 
 'Move_Files (StrImagen_cambio)
 Guarda_Imagen2 = True

End Function




Public Function Guarda_Imagen() As Boolean
    Dim Con As New ADODB.Connection
    Dim strSQL As String
    On Error GoTo Guarda_ImagenErr

    Guarda_Imagen = False
    Con.ConnectionString = cCONNECT
    Con.Open
    
    Con.BeginTrans
       
        
    strSQL = "EXEC SP_ActualizaDir_Icono '" & Codigo_item & "','" & StrImagen_cambio & "'"

    
    Con.Execute strSQL
        
    Con.CommitTrans
    'Move_Files (StrImagen_cambio)
    Guarda_Imagen = True
    Exit Function
Guarda_ImagenErr:
    Con.RollbackTrans
    Set Con = Nothing
    ErrorHandler Err, "Guarda_Imagen"



End Function


'Sub Move_Files(strOption As String)
'
'On Error Resume Next
'
''If strOption = "I" Or strOption = "A" Then
'  Dim fso As New FileSystemObject, fil1, fil2
'
'  If StrImagen_cambio = "" Then Exit Sub
'
''  Set fil1 = fso.GetFile(StrImagen1_cambio)
''  fil1.Delete
'
'  If StrImagen_cambio <> StrImagen1_Origen Then
'     ruta_imagenes = DevuelveCampo("Select Ruta_Iconos_Estilos from Tg_Control", cCONNECT) & fso.GetFileName(StrImagen_cambio)
'     ruta_imagenes = txtfile 'ruta_imagenes '"C:\mitemp\"
'
'     FileCopy StrImagen_cambio, txtfile
'  End If
''End If
'Exit Sub
'Resume
'
'ErrHandler:
'ErrorHandler Err, "Move_Files"
'
'End Sub
