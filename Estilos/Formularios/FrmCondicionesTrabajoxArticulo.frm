VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "FuncButt.ocx"
Begin VB.Form FrmCondicionesTrabajoxArticulo 
   Caption         =   "Condiciones de Trabajo por Artículo"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   540
      Left            =   7875
      TabIndex        =   45
      Top             =   9030
      Width           =   1170
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1050
      TabIndex        =   31
      Top             =   9030
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   "0~0~GUARDAR~True~True~&Guardar~0~0~1~~0~False~False~&Guardar~~1~0~IMPRIMIR~True~True~&Imprimir~0~0~2~~0~False~False~&Imprimir~"
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame6 
      Height          =   1590
      Left            =   105
      TabIndex        =   42
      Top             =   7350
      Width           =   9885
      Begin VB.Frame FrmGradoDoblez 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   2100
         TabIndex        =   44
         Top             =   315
         Width           =   2325
         Begin VB.OptionButton OptNinguna 
            Caption         =   "Ninguna"
            Height          =   210
            Left            =   315
            TabIndex        =   26
            Top             =   210
            Width           =   1425
         End
         Begin VB.OptionButton OptMangas 
            Caption         =   "Solo mangas"
            Height          =   210
            Left            =   315
            TabIndex        =   27
            Top             =   525
            Width           =   1425
         End
         Begin VB.OptionButton OptAmbas 
            Caption         =   "Mangas y/o espaldas"
            Height          =   210
            Left            =   315
            TabIndex        =   28
            Top             =   840
            Width           =   1845
         End
      End
      Begin VB.Frame Frame5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   4935
         TabIndex        =   43
         Top             =   315
         Width           =   2430
         Begin VB.OptionButton Opt1 
            Caption         =   "1% - 3 % (de 1º a 2.5º)"
            Height          =   210
            Left            =   315
            TabIndex        =   30
            Top             =   735
            Width           =   1950
         End
         Begin VB.OptionButton Opt0 
            Caption         =   "0% - 1% (de 0º a 1º)"
            Height          =   210
            Left            =   315
            TabIndex        =   29
            Top             =   315
            Width           =   1740
         End
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Inclinación de Trama"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5145
         TabIndex        =   49
         Top             =   105
         Width           =   1800
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Grado de linea doblez"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2310
         TabIndex        =   48
         Top             =   105
         Width           =   1875
      End
   End
   Begin VB.Frame FraMermas 
      Caption         =   "Mermas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   105
      TabIndex        =   2
      Top             =   5775
      Width           =   9885
      Begin VB.TextBox TxtDesPorcentaje 
         Height          =   285
         Left            =   3990
         TabIndex        =   25
         Top             =   1155
         Width           =   5685
      End
      Begin VB.TextBox TxtPorcentaje 
         Height          =   285
         Left            =   2730
         TabIndex        =   24
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mermas de Orillo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1260
         TabIndex        =   40
         Top             =   210
         Width           =   7155
         Begin VB.OptionButton OptOrillosEngomados 
            Caption         =   "Orillos engomados -3.5cm "
            Height          =   330
            Left            =   5040
            TabIndex        =   23
            Top             =   315
            Width           =   1590
         End
         Begin VB.OptionButton OptSinOrillos 
            Caption         =   "Sin orillos - 2.0 cm"
            Height          =   330
            Left            =   210
            TabIndex        =   21
            Top             =   315
            Width           =   960
         End
         Begin VB.OptionButton OptOrillosAguja 
            Caption         =   "Orillos de aguja - 2.5 cm"
            Height          =   330
            Left            =   2625
            TabIndex        =   22
            Top             =   315
            Width           =   1380
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje de Merma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   735
         TabIndex        =   41
         Top             =   1155
         Width           =   1815
      End
   End
   Begin VB.Frame FraCaracteristicas 
      Caption         =   "Características Físicas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   105
      TabIndex        =   1
      Top             =   2730
      Width           =   9885
      Begin VB.Frame Frame3 
         Caption         =   "Encogimientos Esperados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   210
         TabIndex        =   37
         Top             =   1365
         Width           =   9570
         Begin VB.Frame Frame7 
            Caption         =   "Ancho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   5880
            TabIndex        =   51
            Top             =   210
            Width           =   3570
            Begin VB.Frame Frame11 
               Caption         =   "Porcentual"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   105
               TabIndex        =   63
               Top             =   210
               Width           =   2115
               Begin VB.TextBox TxtAncho1 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1050
                  TabIndex        =   65
                  Top             =   210
                  Width           =   690
               End
               Begin VB.TextBox TxtAncho2 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1050
                  TabIndex        =   64
                  Top             =   525
                  Width           =   690
               End
               Begin VB.Label Label18 
                  AutoSize        =   -1  'True
                  Caption         =   "Desde"
                  Height          =   195
                  Left            =   210
                  TabIndex        =   69
                  Top             =   240
                  Width           =   465
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  Caption         =   "Hasta"
                  Height          =   195
                  Left            =   210
                  TabIndex        =   68
                  Top             =   555
                  Width           =   420
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "%"
                  Height          =   195
                  Left            =   1785
                  TabIndex        =   67
                  Top             =   255
                  Width           =   120
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "%"
                  Height          =   195
                  Left            =   1785
                  TabIndex        =   66
                  Top             =   630
                  Width           =   120
               End
            End
            Begin VB.Frame Frame9 
               Caption         =   "Particular"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   2310
               TabIndex        =   54
               Top             =   210
               Width           =   1170
               Begin VB.TextBox TxtAncho 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   210
                  TabIndex        =   55
                  Top             =   420
                  Width           =   795
               End
            End
         End
         Begin VB.Frame FraLargo 
            Caption         =   "Largo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   105
            TabIndex        =   50
            Top             =   210
            Width           =   3570
            Begin VB.Frame Frame10 
               Caption         =   "Porcentual"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   105
               TabIndex        =   56
               Top             =   210
               Width           =   2115
               Begin VB.TextBox TxtLargo2 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1050
                  TabIndex        =   58
                  Top             =   550
                  Width           =   690
               End
               Begin VB.TextBox TxtLargo1 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1050
                  TabIndex        =   57
                  Top             =   210
                  Width           =   690
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "%"
                  Height          =   195
                  Left            =   1785
                  TabIndex        =   62
                  Top             =   630
                  Width           =   120
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  Caption         =   "%"
                  Height          =   195
                  Left            =   1785
                  TabIndex        =   61
                  Top             =   255
                  Width           =   120
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  Caption         =   "Hasta"
                  Height          =   195
                  Left            =   210
                  TabIndex        =   60
                  Top             =   580
                  Width           =   420
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "Desde"
                  Height          =   195
                  Left            =   210
                  TabIndex        =   59
                  Top             =   240
                  Width           =   465
               End
            End
            Begin VB.Frame Frame8 
               Caption         =   "Particular"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   2310
               TabIndex        =   52
               Top             =   210
               Width           =   1170
               Begin VB.TextBox TxtLargo 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   210
                  TabIndex        =   53
                  Top             =   420
                  Width           =   795
               End
            End
         End
         Begin VB.TextBox TxtRevirado 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4725
            TabIndex        =   20
            Top             =   1110
            Width           =   900
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   5670
            TabIndex        =   39
            Top             =   1155
            Width           =   120
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Revirado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3780
            TabIndex        =   38
            Top             =   1155
            Width           =   780
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de tejido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   5355
         TabIndex        =   36
         Top             =   315
         Width           =   4215
         Begin VB.OptionButton OptConSentido 
            Caption         =   "Con Sentido Unico"
            Height          =   330
            Left            =   2205
            TabIndex        =   19
            Top             =   315
            Width           =   1695
         End
         Begin VB.OptionButton OptSinSentido 
            Caption         =   "Sin Sentido"
            Height          =   330
            Left            =   420
            TabIndex        =   18
            Top             =   315
            Width           =   1380
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Condiciones de llegada de tela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   420
         TabIndex        =   35
         Top             =   315
         Width           =   4215
         Begin VB.OptionButton OptAbierta 
            Caption         =   "Abierta"
            Height          =   330
            Left            =   2205
            TabIndex        =   17
            Top             =   315
            Width           =   1275
         End
         Begin VB.OptionButton OptTubular 
            Caption         =   "Tubular"
            Height          =   330
            Left            =   630
            TabIndex        =   16
            Top             =   315
            Width           =   960
         End
      End
   End
   Begin VB.Frame FraGenerales 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   9885
      Begin VB.TextBox TxtArticulo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   47
         Top             =   840
         Width           =   1065
      End
      Begin VB.TextBox TxtDescripcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2835
         TabIndex        =   46
         Top             =   840
         Width           =   6735
      End
      Begin VB.TextBox TxtAnchoEsta 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4935
         TabIndex        =   15
         Top             =   1995
         Width           =   1005
      End
      Begin VB.TextBox TxtAnchoHist 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   1995
         Width           =   1005
      End
      Begin VB.TextBox TxtDensHist 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8190
         TabIndex        =   11
         Top             =   1365
         Width           =   1005
      End
      Begin VB.TextBox TxtDensReq 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4935
         TabIndex        =   9
         Top             =   1365
         Width           =   1005
      End
      Begin VB.TextBox TxtDensEsta 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   1365
         Width           =   1005
      End
      Begin VB.TextBox txtFamTela 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   420
         Width           =   1845
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(g/m2)"
         Height          =   195
         Left            =   9240
         TabIndex        =   34
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "(g/m2)"
         Height          =   195
         Left            =   5985
         TabIndex        =   33
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "(g/m2)"
         Height          =   195
         Left            =   2730
         TabIndex        =   32
         Top             =   1470
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Estándar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   14
         Top             =   2100
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ancho Histórico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   2100
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Densidad Histórica"
         Height          =   195
         Left            =   6720
         TabIndex        =   10
         Top             =   1470
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Densidad Requerida"
         Height          =   195
         Left            =   3360
         TabIndex        =   8
         Top             =   1470
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Densidad Estandar"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   1470
         Width           =   1350
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Familia de tela"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   465
         Width           =   1005
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   840
         Width           =   555
      End
   End
End
Attribute VB_Name = "FrmCondicionesTrabajoxArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gradolinea, inclinaciontrama, condicionllegada, tipotejido, mermaorilla As String

Dim Rs As New ADODB.Recordset
Public Codigo As String
Public Descripcion As String

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
gradolinea = " "
inclinaciontrama = " "
condicionllegada = " "
tipotejido = " "
mermaorilla = " "
Select Case ActionName
    Case "GUARDAR"
        If SALVA_DATOS Then
            If OptTubular = True Then
                condicionllegada = "T"
            Else
                condicionllegada = "A"
            End If
            
            If OptSinSentido = True Then
                tipotejido = "S"
            Else
                tipotejido = "C"
            End If
            
            If OptNinguna = True Then
                gradolinea = "N"
            ElseIf OptMangas = True Then
                gradolinea = "M"
            Else
                gradomangas = "A"
            End If
            
            If Opt0 = True Then
                inclinaciontrama = "0"
            Else
                inclinaciontrama = "1"
            End If
            
            If OptSinOrillos = True Then
                mermaorilla = "S"
            ElseIf OptOrillosAguja = True Then
                mermaorilla = "A"
            Else
                mermaorilla = "E"
            End If
            
            Set Rs = New ADODB.Recordset
            Rs.ActiveConnection = cCONNECT
            Rs.CursorType = adOpenStatic
            Rs.CursorLocation = adUseClient
            Rs.LockType = adLockReadOnly
            
            strSQL = "UP_MAN_TX_TELA 'U','" & TxtArticulo & "'," & TxtDensEsta & "," & TxtDensReq & "," & TxtAnchoEsta & ",'" & condicionllegada & "','" & tipotejido & "'," & TxtLargo1 & "," & TxtLargo2 & "," & TxtAncho1 & "," & TxtAncho2 & ",'" & mermaorilla & "','" & TxtPorcentaje & "'," & TxtLargo & "," & TxtAncho & "," & TxtRevirado & ",'" & _
                                            ComputerName & "','" & vusu & "'"
                       
            Rs.Open strSQL, cCONNECT, 3, 3
            MsgBox "Se grabó correctamente", vbInformation
        End If
        
    Case "IMPRIMIR"
        Dim oo As Object
        
                If OptTubular = True Then
                condicionllegada = "T"
            Else
                condicionllegada = "A"
            End If
            
            If OptSinSentido = True Then
                tipotejido = "S"
            Else
                tipotejido = "C"
            End If
            
            If OptNinguna = True Then
                gradolinea = "N"
            ElseIf OptMangas = True Then
                gradolinea = "M"
            Else
                gradomangas = "A"
            End If
            
            If Opt0 = True Then
                inclinaciontrama = "0"
            Else
                inclinaciontrama = "1"
            End If
            
            If OptSinOrillos = True Then
                mermaorilla = "S"
            ElseIf OptOrillosAguja = True Then
                mermaorilla = "A"
            Else
                mermaorilla = "E"
            End If
        
            On Error GoTo AceptarErr
            Set oo = CreateObject("excel.application")
            oo.workbooks.Open vRuta & "\RptCondicionesTrabajo.xlt"
            oo.Visible = True
            oo.run "Reporte", TxtArticulo, TxtDensEsta, cCONNECT
            Screen.MousePointer = vbNormal
            oo.Visible = True
            Set oo = Nothing
    End Select
    Exit Sub
AceptarErr:
    
End Sub

Private Sub TxtAncho1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    AVANZA 13
Else
    SoloNumeros TxtAncho1, KeyAscii, True, 2, 5
End If
End Sub

Private Sub TxtAncho2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    AVANZA 13
Else
    SoloNumeros TxtAncho2, KeyAscii, True, 2, 5
        End If
End Sub

Private Sub TxtAnchoEsta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    AVANZA 13
Else
    SoloNumeros TxtAnchoEsta, KeyAscii, True, 2, 5
End If
End Sub

Private Sub TxtAnchoHist_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    AVANZA 13
Else
    SoloNumeros TxtAnchoHist, KeyAscii, True, 2, 5
End If
End Sub

Private Sub TxtDensEsta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AVANZA 13
    Else
        SoloNumeros TxtDensEsta, KeyAscii, True, 2, 5
    End If
End Sub

Private Sub TxtDensHist_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    AVANZA 13
Else
    SoloNumeros TxtDensHist, KeyAscii, True, 2, 5
End If
End Sub

Private Sub TxtDensReq_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    AVANZA 13
Else
    SoloNumeros TxtDensReq, KeyAscii, True, 2, 5
End If
End Sub

Private Sub TxtLargo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    AVANZA 13
Else
    SoloNumeros TxtLargo1, KeyAscii, True, 2, 5
End If
End Sub


Private Sub TxtLargo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    AVANZA 13
Else
    SoloNumeros TxtLargo2, KeyAscii, True, 2, 5
End If
End Sub

Private Sub TxtPorcentaje_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If txtCodigo = "" Then
            Call MUESTRA_AYUDA
        End If
    End If
End Sub

Public Sub MUESTRA_AYUDA()

    Dim oTipo As New frmBusqGeneral
    Dim Rs As New ADODB.Recordset
    Set oTipo.oParent = Me

    oTipo.sQuery = "SELECT por_merma_manufactura,descripcion FROM TX_PORCENTAJES"

    oTipo.Cargar_Datos
    oTipo.Show 1
    If Codigo <> "" Then
        TxtPorcentaje.Text = Trim(Codigo)
        TxtDesPorcentaje.Text = Trim(Descripcion)
        Codigo = "": Descripcion = ""
    End If
    Set oTipo = Nothing
    Set Rs = Nothing

    TxtPorcentaje.SetFocus
End Sub


Function SALVA_DATOS() As Boolean
    
    If Len(RTrim(TxtDensEsta)) = 0 Then
        MsgBox "Debe ingresar Densidad estandar o en su defecto 0 :Cero", vbInformation
        Exit Function
    End If
    
    If Len(RTrim(TxtDensReq)) = 0 Then
        MsgBox "Debe ingresar Densidad requerida o en su defecto 0 :Cero", vbInformation
        Exit Function
    End If
    
    If Len(RTrim(TxtAnchoEsta)) = 0 Then
        MsgBox "Debe ingresar Ancho estandar o en su defecto 0 :Cero", vbInformation
        Exit Function
    End If
    
    If Len(RTrim(TxtLargo1)) = 0 Then
        MsgBox "Debe ingresar porcentaje largo desde o en su defecto 0 :Cero", vbInformation
        Exit Function
    End If
    
    If Len(RTrim(TxtLargo2)) = 0 Then
        MsgBox "Debe ingresar porcentaje largo hasta o en su defecto 0 :Cero", vbInformation
        Exit Function
    End If
    
    If Len(RTrim(TxtAncho1)) = 0 Then
        MsgBox "Debe ingresar porcentaje ancho desde o en su defecto 0 :Cero", vbInformation
        Exit Function
    End If
    
    If Len(RTrim(TxtAncho2)) = 0 Then
        MsgBox "Debe ingresar porcentaje ancho hasta o  en su defecto 0 :Cero", vbInformation
        Exit Function
    End If
   If Len(RTrim(TxtRevirado)) = 0 Then
        MsgBox "Debe ingresar porcentaje revirado o en su defecto 0 :Cero", vbInformation
        Exit Function
    End If
    
    If Len(RTrim(TxtPorcentaje)) = 0 Then
        MsgBox "Debe ingresar porcentaje merma", vbInformation
        Exit Function
    End If
    SALVA_DATOS = True
End Function

