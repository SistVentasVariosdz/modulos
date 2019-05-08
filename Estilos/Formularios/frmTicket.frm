VERSION 5.00
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTicket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresion de Tickets"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmTicket.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   360
      TabIndex        =   22
      Top             =   6120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   900
      Custom          =   $"frmTicket.frx":0442
      Orientacion     =   0
      Style           =   0
      Language        =   0
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   6075
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   6705
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   600
         Visible         =   0   'False
         Width           =   6495
         Begin VB.TextBox txtruta 
            Height          =   285
            Left            =   1860
            MaxLength       =   4
            TabIndex        =   54
            Top             =   0
            Width           =   660
         End
         Begin VB.TextBox txtdesruta 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2550
            MaxLength       =   50
            TabIndex        =   53
            Top             =   0
            Width           =   3900
         End
         Begin VB.Label Label28 
            Caption         =   "Ruta                              :"
            Height          =   225
            Left            =   0
            TabIndex        =   55
            Top             =   75
            Width           =   1800
         End
      End
      Begin VB.TextBox TxtPeso_Lote 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1980
         TabIndex        =   20
         Text            =   "0"
         Top             =   5280
         Width           =   1035
      End
      Begin VB.TextBox TxtPartida 
         Height          =   285
         Left            =   1980
         TabIndex        =   3
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Frame FraComb 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3360
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   3255
         Begin VB.TextBox TxtDes_Comb 
            Height          =   285
            Left            =   1200
            TabIndex        =   2
            Top             =   80
            Width           =   2010
         End
         Begin VB.TextBox TxtCod_Comb 
            Height          =   285
            Left            =   600
            TabIndex        =   1
            Top             =   80
            Width           =   570
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Comb :"
            Height          =   195
            Left            =   0
            TabIndex        =   48
            Top             =   150
            Width           =   495
         End
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2340
         TabIndex        =   37
         Top             =   5715
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "TxtCantidad"
         BuddyDispid     =   196622
         OrigLeft        =   2580
         OrigTop         =   4515
         OrigRight       =   2820
         OrigBottom      =   4800
         Max             =   10000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox TxtCantidad 
         Height          =   285
         Left            =   2025
         TabIndex        =   21
         Text            =   "1"
         Top             =   5715
         Width           =   315
      End
      Begin VB.TextBox TxtcolorWay 
         Height          =   315
         Left            =   1995
         TabIndex        =   18
         Text            =   "Text17"
         Top             =   4515
         Width           =   4560
      End
      Begin VB.TextBox TxtDiamGalga 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4515
         TabIndex        =   17
         Text            =   "0"
         Top             =   4155
         Width           =   800
      End
      Begin VB.TextBox TxtGalga 
         Height          =   285
         Left            =   1995
         TabIndex        =   16
         Text            =   "Text15"
         Top             =   4155
         Width           =   1530
      End
      Begin VB.TextBox TxtPesoOZ2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4515
         TabIndex        =   15
         Text            =   "0"
         Top             =   3795
         Width           =   800
      End
      Begin VB.TextBox TxtPeso2 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1995
         TabIndex        =   14
         Text            =   "0"
         Top             =   3795
         Width           =   800
      End
      Begin VB.TextBox TxtPesoOZ 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4515
         TabIndex        =   13
         Text            =   "0"
         Top             =   3435
         Width           =   800
      End
      Begin VB.TextBox TxtPeso 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1995
         TabIndex        =   12
         Text            =   "0"
         Top             =   3435
         Width           =   800
      End
      Begin VB.TextBox TxtAnchoPulg 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4515
         TabIndex        =   11
         Text            =   "0"
         Top             =   3090
         Width           =   800
      End
      Begin VB.TextBox TxtAncho 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1995
         TabIndex        =   10
         Text            =   "0"
         Top             =   3090
         Width           =   800
      End
      Begin VB.TextBox TxtEncogAncho 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4515
         TabIndex        =   9
         Text            =   "0"
         Top             =   2745
         Width           =   800
      End
      Begin VB.TextBox TxtEncogLargo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1995
         TabIndex        =   8
         Text            =   "Text7"
         Top             =   2745
         Width           =   800
      End
      Begin VB.TextBox TxtProcesoLav 
         Height          =   285
         Left            =   1980
         TabIndex        =   19
         Text            =   "Text6"
         Top             =   4920
         Width           =   4600
      End
      Begin VB.TextBox TxtMetodoTen 
         Height          =   285
         Left            =   1980
         TabIndex        =   7
         Text            =   "Text5"
         Top             =   2415
         Width           =   4600
      End
      Begin VB.TextBox TxtComposicion 
         Height          =   285
         Left            =   1980
         TabIndex        =   6
         Text            =   "Text4"
         Top             =   2070
         Width           =   4600
      End
      Begin VB.TextBox TxtHilado 
         Height          =   285
         Left            =   1980
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   1725
         Width           =   4600
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   1980
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1380
         Width           =   4600
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   285
         Left            =   1980
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   210
         Width           =   1290
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "grms/mt2"
         Height          =   195
         Left            =   3120
         TabIndex        =   51
         Top             =   5400
         Width           =   660
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Lot Weight                  :"
         Height          =   195
         Left            =   150
         TabIndex        =   50
         Top             =   5385
         Width           =   1635
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Partida                          :"
         Height          =   195
         Left            =   150
         TabIndex        =   49
         Top             =   1035
         Width           =   1710
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "oz/yd2"
         Height          =   195
         Left            =   5355
         TabIndex        =   46
         Top             =   3870
         Width           =   495
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "oz/yd2"
         Height          =   195
         Left            =   5340
         TabIndex        =   45
         Top             =   3495
         Width           =   495
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "% width"
         Height          =   195
         Left            =   5325
         TabIndex        =   44
         Top             =   2790
         Width           =   540
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "incles"
         Height          =   195
         Left            =   5340
         TabIndex        =   43
         Top             =   3165
         Width           =   405
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Diametro :"
         Height          =   195
         Left            =   3735
         TabIndex        =   42
         Top             =   4200
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "grms/mt2"
         Height          =   195
         Left            =   2805
         TabIndex        =   41
         Top             =   3885
         Width           =   660
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "grms/mt2"
         Height          =   195
         Left            =   2820
         TabIndex        =   40
         Top             =   3525
         Width           =   660
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "mts."
         Height          =   195
         Left            =   2835
         TabIndex        =   39
         Top             =   3165
         Width           =   285
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "% lenght"
         Height          =   195
         Left            =   2805
         TabIndex        =   38
         Top             =   2835
         Width           =   600
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Copias a Imprimir         :"
         Height          =   195
         Left            =   150
         TabIndex        =   36
         Top             =   5730
         Width           =   1635
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Colorway                      :"
         Height          =   195
         Left            =   150
         TabIndex        =   35
         Top             =   4605
         Width           =   1680
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Gauge                          :"
         Height          =   195
         Left            =   150
         TabIndex        =   34
         Top             =   4245
         Width           =   1695
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Weitght STD a/w         :"
         Height          =   195
         Left            =   150
         TabIndex        =   33
         Top             =   3885
         Width           =   1710
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Weight STD b/w          :"
         Height          =   195
         Left            =   150
         TabIndex        =   32
         Top             =   3525
         Width           =   1710
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Width                            :"
         Height          =   195
         Left            =   150
         TabIndex        =   31
         Top             =   3195
         Width           =   1725
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Shrinkage                     :"
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   2835
         Width           =   1710
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Laundry process          :"
         Height          =   195
         Left            =   150
         TabIndex        =   29
         Top             =   4995
         Width           =   1665
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dyeing method             :"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   2460
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Composition                 :"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   2100
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Yarn                             :"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   1740
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fabric                           :"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   1395
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Article #                        :"
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Top             =   225
         Width           =   1710
      End
   End
End
Attribute VB_Name = "frmTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strCadena As String
Dim strSQL  As String
Dim iLin As Integer
Public vCod_FamTela As String
Dim strCadena1 As String
Dim strCadena2 As String
Dim strCadena3 As String
Dim strCadena4 As String
Dim strCadena5 As String
Dim strCadena6 As String
Dim strCadena7 As String
Dim strCadena8 As String
Dim strCadena9 As String
Dim strCadena10 As String
Dim strCadena11 As String
Dim strCadena12 As String
Dim strCadena13 As String
Dim strCadena14 As String
Public sMascara As String
Public Codigo As String, Descripcion As String, TipoAdd As String, TipoAdd2 As String

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim i As Integer
    Select Case ActionName
        Case "IMPRIMIR"
            ImprimirTicket
            If Me.vCod_FamTela = "DE" Then
                Call Act_Gramajes
            End If
'            Unload Me
        Case "CARTON"
            ImprimirCarton
        Case "ZEBRA"
                   For i = 1 To CDbl(TxtCantidad.Text)
                        ImprimirTicket1
                   Next
        Case "CANCELAR"
            Unload Me
        Case "NUEVA"
                   For i = 1 To CDbl(TxtCantidad.Text)
                        ImprimirTicket2
                   Next
    End Select
End Sub


Private Function ImprimirTicket2()
On Error GoTo errx
'Dim sSQL  As String
'Dim mRS As ADODB.Recordset
Dim oPrint As clsPrintFile
Dim M As Printer
MsgBox vRuta
sMascara = vRuta & "\HPLANTA_TELAS.EJF"

For Each M In Printers
   If UCase(M.DeviceName) = "ELTRON" Then
      Set Printer = M
      Exit For
   End If
Next

strCadena1 = Mid(Trim(Me.TxtEncogLargo), 1, 5) & "% length" & Space(1) & Space(5 - Len(Mid(Me.TxtEncogAncho, 1, 5))) & Trim(Mid(Me.TxtEncogAncho, 1, 5)) & "% width"
strCadena2 = Mid(Trim(Me.TxtAncho), 1, 5) & "mts" & Space(6)
strCadena2 = strCadena2 & Space(7 - Len(Mid(Me.TxtAnchoPulg, 1, 5))) & Trim(Mid(Me.TxtAnchoPulg, 1, 5)) & "inches"
strCadena3 = Mid(Trim(Me.TxtPeso), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ, 1, 5)) & " oz/yd2"
strCadena4 = Mid(Trim(Me.TxtPeso2), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ2, 1, 5)) & " oz/yd2"
strCadena5 = Mid(Trim(Me.TxtGalga), 1, 17) & Space(15 - Len(Mid(Me.TxtGalga, 1, 17))) & Space(1) & "Diam :" & Trim(Mid(Me.TxtDiamGalga, 1, 5))
strCadena11 = Mid(Me.TxtcolorWay, 1, 31)
strCadena12 = Mid(Me.TxtProcesoLav, 1, 31)
strCadena13 = Me.TxtPeso_Lote.Text
If Frame2.Visible = True Then
    strCadena14 = Me.txtruta.Text & "-" & Me.txtdesruta.Text
End If

   Close #1
    Open vRuta & "\ROLLO1.TXT" For Output As #1
    Plin "FR" & Chr(34) & "HPLANTA_TELAS" & Chr(34) & ""
    Plin "?"
    Plin "Article #"
    Plin ": " & Me.TxtCodigo
    Plin "Fabric"
    Plin ": " & Mid(Me.TxtDescripcion, 1, 31)
    Plin "Yarn"
    Plin ": " & Mid(Me.TxtHilado, 1, 31)
    Plin "Composition"
    Plin ": " & Mid(Me.TxtComposicion, 1, 31)
    Plin "Dyeing method"
    Plin ": " & Mid(Me.TxtMetodoTen, 1, 31)
    Plin "Shrinkage STD"
    Plin ": " & strCadena1
    Plin "Width"
    Plin ": " & strCadena2
    Plin "Weight STD b/w"
    Plin ": " & strCadena3
    Plin "Weight STD a/w"
    Plin ": " & strCadena4
    Plin "Gauge"
    Plin ": " & strCadena5
    Plin "Colorway"
    Plin ": " & strCadena11
    Plin "Laundry process"
    Plin ": " & strCadena12
    Plin "Lot Weight"
    Plin ": " & strCadena13
    If Frame2.Visible = True Then
    Plin "Path"
    Plin ": " & strCadena14
    End If
    Plin "P1"
    Close #1
    
    sMascara = vRuta & "\HPLANTA_TELAS.EJF"
    
    Set oPrint = New clsPrintFile
    oPrint.SendPrint sMascara ', "ELTRON"
    Set oPrint = Nothing
    
    Set oPrint = New clsPrintFile
    oPrint.SendPrint vRuta & "\ROLLO1.TXT" ', "ELTRON"
    Set oPrint = Nothing
    
    Exit Function
errx:
    Close #1
    errores Err.numer

'/*************************/
'Printer.Print " "
'Printer.Print "^XA"
'Printer.Print "^PRC"
'Printer.Print "^LH0, 0 ^ FS"
'Printer.Print "^LL500"
'Printer.Print "^MD0"
'Printer.Print "^MNY"
'Printer.Print "^LH0, 0 ^ FS"
'
'
'Printer.Print "^FO80,30^A0N,25,30^CI13^FR^FDArticle #^FS"
'Printer.Print "^FO300,30^A0N,25,30^CI13^FR^FD: " & Me.TxtCodigo & "^FS"
'
'Printer.Print "^FO80,59^A0N,25,30^CI13^FR^FDFabric^FS"
'Printer.Print "^FO300,59^A0N,25,30^CI13^FR^FD: " & Mid(Me.TxtDescripcion, 1, 31) & "^FS"
'
'Printer.Print "^FO80,88^A0N,25,27^CI13^FR^FDYarn^FS"
'Printer.Print "^FO300,88^A0N,25,27^CI13^FR^FD: " & Mid(Me.TxtHilado, 1, 31) & "^FS"
'
'Printer.Print "^FO80,117^A0N,25,30^CI13^FR^FDComposition^FS"
'Printer.Print "^FO300,117^A0N,25,30^CI13^FR^FD: " & Mid(Me.TxtComposicion, 1, 31) & "^FS"
'
'
'Printer.Print "^FO80,146^A0N,25,30^CI13^FR^FDDyeing method^FS"
'Printer.Print "^FO300,146^A0N,25,30^CI13^FR^FD: " & Mid(Me.TxtMetodoTen, 1, 31) & "^FS"
'
'strCadena1 = Mid(Trim(Me.TxtEncogLargo), 1, 5) & "% length" & Space(1) & Space(5 - Len(Mid(Me.TxtEncogAncho, 1, 5))) & Trim(Mid(Me.TxtEncogAncho, 1, 5)) & "% width"
'Printer.Print "^FO80,175^A0N,25,30^CI13^FR^FDShrinkage STD^FS"
'Printer.Print "^FO300,175^A0N,25,30^CI13^FR^FD: " & strCadena1 & "^FS"
'
'Printer.Print "^FO80,204^A0N,25,30^CI13^FR^FDWidth^FS"
'strCadena2 = Mid(Trim(Me.TxtAncho), 1, 5) & "mts" & Space(6)
'strCadena2 = strCadena2 & Space(7 - Len(Mid(Me.TxtAnchoPulg, 1, 5))) & Trim(Mid(Me.TxtAnchoPulg, 1, 5)) & "inches"
'Printer.Print "^FO300,204^A0N,25,30^CI13^FR^FD: " & strCadena2 & "^FS"
'
'
'Printer.Print "^FO80,233^A0N,25,30^CI13^FR^FDWeight STD b/w^FS"
'strCadena3 = Mid(Trim(Me.TxtPeso), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ, 1, 5)) & " oz/yd2"
'Printer.Print "^FO300,233^A0N,25,30^CI13^FR^FD: " & strCadena3 & "^FS"
'
'
'Printer.Print "^FO80,262^A0N,25,30^CI13^FR^FDWeight STD a/w^FS"
'strCadena4 = Mid(Trim(Me.TxtPeso2), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ2, 1, 5)) & " oz/yd2"
'Printer.Print "^FO300,262^A0N,25,30^CI13^FR^FD: " & strCadena4 & "^FS"
'
'Printer.Print "^FO80,291^A0N,25,30^CI13^FR^FDGauge^FS"
'strCadena5 = Mid(Trim(Me.TxtGalga), 1, 17) & Space(15 - Len(Mid(Me.TxtGalga, 1, 17))) & Space(1) & "Diam :" & Trim(Mid(Me.TxtDiamGalga, 1, 5))
'Printer.Print "^FO300,291^A0N,25,30^CI13^FR^FD: " & strCadena5 & "^FS"
'
'Printer.Print "^FO80,320^A0N,25,30^CI13^FR^FDColorway^FS"
'strCadena11 = Mid(Me.TxtcolorWay, 1, 31)
'Printer.Print "^FO300,320^A0N,25,30^CI13^FR^FD: " & strCadena11 & "^FS"
'
'Printer.Print "^FO80,349^A0N,25,30^CI13^FR^FDLaundry process^FS"
'strCadena12 = Mid(Me.TxtProcesoLav, 1, 31)
'Printer.Print "^FO300,349^A0N,25,30^CI13^FR^FD: " & strCadena12 & "^FS"
'
'Printer.Print "^FO80,375^A0N,25,30^CI13^FR^FDLot Weight^FS"
'strCadena13 = Me.TxtPeso_Lote.Text
'Printer.Print "^FO300,375^A0N,25,30^CI13^FR^FD: " & strCadena13 & "^FS"
'
'Printer.Print "^XZ"
'Printer.Print "^FX End of job"
'Printer.Print "^XA"
'Printer.Print "^IDR:ID*.*"
'Printer.Print "^XZ"
'Printer.EndDoc

End Function

Sub ImprimirTicket1()
Printer.Print " "
Printer.Print "^XA"
Printer.Print "^PRC"
Printer.Print "^LH0, 0 ^ FS"
Printer.Print "^LL500"
Printer.Print "^MD0"
Printer.Print "^MNY"
Printer.Print "^LH0, 0 ^ FS"


Printer.Print "^FO80,30^A0N,25,30^CI13^FR^FDArticle #^FS"
Printer.Print "^FO300,30^A0N,25,30^CI13^FR^FD: " & Me.TxtCodigo & "^FS"

Printer.Print "^FO80,59^A0N,25,30^CI13^FR^FDFabric^FS"
Printer.Print "^FO300,59^A0N,25,30^CI13^FR^FD: " & Mid(Me.TxtDescripcion, 1, 31) & "^FS"

Printer.Print "^FO80,88^A0N,25,27^CI13^FR^FDYarn^FS"
Printer.Print "^FO300,88^A0N,25,27^CI13^FR^FD: " & Mid(Me.TxtHilado, 1, 31) & "^FS"

Printer.Print "^FO80,117^A0N,25,30^CI13^FR^FDComposition^FS"
Printer.Print "^FO300,117^A0N,25,30^CI13^FR^FD: " & Mid(Me.TxtComposicion, 1, 31) & "^FS"


Printer.Print "^FO80,146^A0N,25,30^CI13^FR^FDDyeing method^FS"
Printer.Print "^FO300,146^A0N,25,30^CI13^FR^FD: " & Mid(Me.TxtMetodoTen, 1, 31) & "^FS"

strCadena1 = Mid(Trim(Me.TxtEncogLargo), 1, 5) & "% length" & Space(1) & Space(5 - Len(Mid(Me.TxtEncogAncho, 1, 5))) & Trim(Mid(Me.TxtEncogAncho, 1, 5)) & "% width"
Printer.Print "^FO80,175^A0N,25,30^CI13^FR^FDShrinkage STD^FS"
Printer.Print "^FO300,175^A0N,25,30^CI13^FR^FD: " & strCadena1 & "^FS"

Printer.Print "^FO80,204^A0N,25,30^CI13^FR^FDWidth^FS"
strCadena2 = Mid(Trim(Me.TxtAncho), 1, 5) & "mts" & Space(6)
strCadena2 = strCadena2 & Space(7 - Len(Mid(Me.TxtAnchoPulg, 1, 5))) & Trim(Mid(Me.TxtAnchoPulg, 1, 5)) & "inches"
Printer.Print "^FO300,204^A0N,25,30^CI13^FR^FD: " & strCadena2 & "^FS"


Printer.Print "^FO80,233^A0N,25,30^CI13^FR^FDWeight STD b/w^FS"
strCadena3 = Mid(Trim(Me.TxtPeso), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ, 1, 5)) & " oz/yd2"
Printer.Print "^FO300,233^A0N,25,30^CI13^FR^FD: " & strCadena3 & "^FS"


Printer.Print "^FO80,262^A0N,25,30^CI13^FR^FDWeight STD a/w^FS"
strCadena4 = Mid(Trim(Me.TxtPeso2), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ2, 1, 5)) & " oz/yd2"
Printer.Print "^FO300,262^A0N,25,30^CI13^FR^FD: " & strCadena4 & "^FS"

Printer.Print "^FO80,291^A0N,25,30^CI13^FR^FDGauge^FS"
strCadena5 = Mid(Trim(Me.TxtGalga), 1, 17) & Space(15 - Len(Mid(Me.TxtGalga, 1, 17))) & Space(1) & "Diam :" & Trim(Mid(Me.TxtDiamGalga, 1, 5))
Printer.Print "^FO300,291^A0N,25,30^CI13^FR^FD: " & strCadena5 & "^FS"

Printer.Print "^FO80,320^A0N,25,30^CI13^FR^FDColorway^FS"
strCadena11 = Mid(Me.TxtcolorWay, 1, 31)
Printer.Print "^FO300,320^A0N,25,30^CI13^FR^FD: " & strCadena11 & "^FS"

Printer.Print "^FO80,349^A0N,25,30^CI13^FR^FDLaundry process^FS"
strCadena12 = Mid(Me.TxtProcesoLav, 1, 31)
Printer.Print "^FO300,349^A0N,25,30^CI13^FR^FD: " & strCadena12 & "^FS"

Printer.Print "^FO80,375^A0N,25,30^CI13^FR^FDLot Weight^FS"
strCadena13 = Me.TxtPeso_Lote.Text
Printer.Print "^FO300,375^A0N,25,30^CI13^FR^FD: " & strCadena13 & "^FS"
 
If Frame2.Visible = True Then
Printer.Print "^FO80,404^A0N,25,30^CI13^FR^FDPath^FS"
strCadena14 = Me.txtruta.Text & "-" & Me.txtdesruta.Text
Printer.Print "^FO300,404^A0N,25,30^CI13^FR^FD: " & strCadena14 & "^FS"
End If

Printer.Print "^XZ"
Printer.Print "^FX End of job"
Printer.Print "^XA"
Printer.Print "^IDR:ID*.*"
Printer.Print "^XZ"
Printer.EndDoc

End Sub

Sub ImprimirTicket()
Dim oPrint As LibraryVB.clsPrintFile
Dim i As Integer

Set oPrint = New clsPrintFile

    Open "c:\PRUEBA.txt" For Output As #1
    
    Plin Chr(15) & "   "

    Plin "     "
    strCadena = " " & " Article #       : " & Me.TxtCodigo
    Plin strCadena

    strCadena = " " & " Fabric          : " & Mid(Me.TxtDescripcion, 1, 31)
    Plin strCadena

    strCadena = " " & " Yarn            : " & Mid(Me.TxtHilado, 1, 31)
    Plin strCadena

    strCadena = " " & " Composition     : " & Mid(Me.TxtComposicion, 1, 31)
    Plin strCadena

    strCadena = " " & " Dyeing method   : " & Mid(Me.TxtMetodoTen, 1, 31)
    Plin strCadena

    strCadena = " " & " Shrinkage STD   : " & Space(5 - Len(Mid(Me.TxtEncogLargo, 1, 5))) & Mid(Trim(Me.TxtEncogLargo), 1, 5) & "% length" & Space(1) & Space(5 - Len(Mid(Me.TxtEncogAncho, 1, 5))) & Trim(Mid(Me.TxtEncogAncho, 1, 5)) & "% width"
    Plin strCadena

    strCadena = " " & " Width           : " & Space(5 - Len(Mid(Me.TxtAncho, 1, 5)))
    strCadena = strCadena & Mid(Trim(Me.TxtAncho), 1, 5) & "mts" & Space(6)
    strCadena = strCadena & Space(7 - Len(Mid(Me.TxtAnchoPulg, 1, 5))) & Trim(Mid(Me.TxtAnchoPulg, 1, 5)) & "inches"
    Plin strCadena

    strCadena = " " & " Weight STD b/w  : " & Space(5 - Len(Mid(Me.TxtPeso, 1, 5))) & Mid(Trim(Me.TxtPeso), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ, 1, 5)) & " oz/yd2"
    Plin strCadena

    strCadena = " " & " Weight STD a/w  : " & Space(5 - Len(Mid(Me.TxtPeso2, 1, 5))) & Mid(Trim(Me.TxtPeso2), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ2, 1, 5)) & " oz/yd2"
    Plin strCadena

    strCadena = " " & " Gauge           : " & Mid(Trim(Me.TxtGalga), 1, 17) & Space(15 - Len(Mid(Me.TxtGalga, 1, 17))) & Space(1) & "Diam :" & Trim(Mid(Me.TxtDiamGalga, 1, 5))
    Plin strCadena

    strCadena = " " & " Colorway        : " & Mid(Me.TxtcolorWay, 1, 31)
    Plin strCadena

    strCadena = " " & " Laundry process : " & Mid(Me.TxtProcesoLav, 1, 31)
    Plin strCadena
    
    strCadena = " " & " Lot Weight      : " & Me.TxtPeso_Lote.Text
    Plin strCadena
    
    If Frame2.Visible = True Then
    
    strCadena = " " & " Path            : " & Me.txtruta & "-" & Me.txtdesruta.Text
    Plin strCadena
    End If
    Plin "   "
'------------------------------------------------------------

'    strCadena = "  " & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95)
'    strCadena = strCadena & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95)
'    Plin strCadena
'    Plin " " & Chr(124) & Space(50) & Chr(124)
'    strCadena = " " & Chr(124) & " Article #       : " & Me.TxtCodigo & Space(23) & Chr(124) & Space(1)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Fabric          : " & Mid(Me.TxtDescripcion, 1, 31) & Chr(124) & Space(1)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Yarn            : " & Mid(Me.TxtHilado, 1, 31) & Space(31 - Len(Me.TxtHilado)) & Chr(124) & Space(1)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Composition     : " & Mid(Me.TxtComposicion, 1, 31) & Space(31 - Len(Me.TxtComposicion)) & Chr(124) & Space(1)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Dyeing method   : " & Mid(Me.TxtMetodoTen, 1, 31) & Space(31 - Len(Me.TxtMetodoTen)) & Chr(124) & Space(1)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Laundry process : " & Mid(Me.TxtProcesoLav, 1, 31) & Space(31 - Len(Me.TxtProcesoLav)) & Chr(124)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Shrinkage       : " & Space(5 - Len(Mid(Me.TxtEncogLargo, 1, 5))) & Mid(Trim(Me.TxtEncogLargo), 1, 5) & "% length" & Space(1) & Space(5 - Len(Mid(Me.TxtEncogAncho, 1, 5))) & Trim(Mid(Me.TxtEncogAncho, 1, 5)) & "% width"
'    strCadena = strCadena & Space(52 - Len(strCadena)) & Chr(124)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Width           : " & Space(5 - Len(Mid(Me.TxtAncho, 1, 5)))
'    strCadena = strCadena & Mid(Trim(Me.TxtAncho), 1, 5) & "mts" & Space(6)
'    strCadena = strCadena & Space(7 - Len(Mid(Me.TxtAnchoPulg, 1, 5))) & Trim(Mid(Me.TxtAnchoPulg, 1, 5)) & "inches"
'    strCadena = strCadena & Space(52 - Len(strCadena)) & Chr(124)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Weight b/w      : " & Space(5 - Len(Mid(Me.TxtPeso, 1, 5))) & Mid(Trim(Me.TxtPeso), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ, 1, 5)) & " oz/yd2"
'    strCadena = strCadena & Space(52 - Len(strCadena)) & Chr(124)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Weight a/w      : " & Space(5 - Len(Mid(Me.TxtPeso2, 1, 5))) & Mid(Trim(Me.TxtPeso2), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ2, 1, 5)) & " oz/yd2"
'    strCadena = strCadena & Space(52 - Len(strCadena)) & Chr(124)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Gauge           : " & Mid(Trim(Me.TxtGalga), 1, 17) & Space(15 - Len(Mid(Me.TxtGalga, 1, 17))) & Space(1) & "Diam :" & Trim(Mid(Me.TxtDiamGalga, 1, 5))
'    strCadena = strCadena & Space(52 - Len(strCadena)) & Chr(124)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Colorway        : " & Mid(Me.TxtcolorWay, 1, 31) & Space(31 - Len(Me.TxtcolorWay)) & Chr(124)
'    Plin strCadena
'
'    strCadena = "  " & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95)
'    strCadena = strCadena & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95)
'    Plin strCadena
    
    Plin "                 "
    
    Close #1
    
    For i = 1 To CDbl(TxtCantidad.Text)
        oPrint.SendPrint "c:\PRUEBA.txt"
    Next
    Set oPrint = Nothing
End Sub

Sub Plin(ByVal Text)
If IsNull(Text) Then
       Text = ""
    End If
    Print #1, Text
    iLin = iLin + 1
End Sub

Private Sub TxtAncho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SoloNumeros TxtAncho, KeyAscii
End Sub

Private Sub TxtAnchoPulg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SoloNumeros TxtAnchoPulg, KeyAscii
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SoloNumeros TxtCantidad, KeyAscii, False
End Sub

Private Sub TxtCod_Comb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Ayuda_Comb("1")
End If
End Sub

Private Sub TxtDes_Comb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Ayuda_Comb("2")
End If
End Sub

Private Sub TxtDiamGalga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SoloNumeros TxtEncogAncho, KeyAscii, False
End Sub

Private Sub TxtEncogAncho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SoloNumeros TxtDiamGalga, KeyAscii, False
End Sub

Private Sub TxtEncogLargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SoloNumeros TxtEncogLargo, KeyAscii, False
End Sub



Sub ImprimirCarton()
Dim oPrint As LibraryVB.clsPrintFile
Dim i As Integer

Set oPrint = New clsPrintFile

    Open "c:\PRUEBA.txt" For Output As #1
    
    Plin Chr(15) & "   "

    Plin "             "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    Plin "     "
    
    strCadena = " " & Space(42) & " Article #       : " & Me.TxtCodigo
    Plin strCadena

    strCadena = " " & Space(42) & " Fabric          : " & Mid(Me.TxtDescripcion, 1, 31)
    Plin strCadena

    strCadena = " " & Space(42) & " Yarn            : " & Mid(Me.TxtHilado, 1, 31)
    Plin strCadena

    strCadena = " " & Space(42) & " Composition     : " & Mid(Me.TxtComposicion, 1, 31)
    Plin strCadena

    strCadena = " " & Space(42) & " Dyeing method   : " & Mid(Me.TxtMetodoTen, 1, 31)
    Plin strCadena

    strCadena = " " & Space(42) & " Laundry process : " & Mid(Me.TxtProcesoLav, 1, 31)
    Plin strCadena

    strCadena = " " & Space(42) & " Shrinkage       : " & Space(5 - Len(Mid(Me.TxtEncogLargo, 1, 5))) & Mid(Trim(Me.TxtEncogLargo), 1, 5) & "% length" & Space(1) & Space(5 - Len(Mid(Me.TxtEncogAncho, 1, 5))) & Trim(Mid(Me.TxtEncogAncho, 1, 5)) & "% width"
    Plin strCadena

    strCadena = " " & Space(42) & " Width           : " & Space(5 - Len(Mid(Me.TxtAncho, 1, 5)))
    strCadena = strCadena & Mid(Trim(Me.TxtAncho), 1, 5) & "mts" & Space(6)
    strCadena = strCadena & Space(7 - Len(Mid(Me.TxtAnchoPulg, 1, 5))) & Trim(Mid(Me.TxtAnchoPulg, 1, 5)) & "inches"
    Plin strCadena

    strCadena = " " & Space(42) & " Weight b/w      : " & Space(5 - Len(Mid(Me.TxtPeso, 1, 5))) & Mid(Trim(Me.TxtPeso), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ, 1, 5)) & " oz/yd2"
    Plin strCadena

    strCadena = " " & Space(42) & " Weight a/w      : " & Space(5 - Len(Mid(Me.TxtPeso2, 1, 5))) & Mid(Trim(Me.TxtPeso2), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ2, 1, 5)) & " oz/yd2"
    Plin strCadena

    strCadena = " " & Space(42) & " Gauge           : " & Mid(Trim(Me.TxtGalga), 1, 17) & Space(15 - Len(Mid(Me.TxtGalga, 1, 17))) & Space(1) & "Diam :" & Trim(Mid(Me.TxtDiamGalga, 1, 5))
    Plin strCadena

    strCadena = " " & Space(42) & " Colorway        : " & Mid(Me.TxtcolorWay, 1, 31)
    Plin strCadena

    Plin "   "
'------------------------------------------------------------

'    strCadena = "  " & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95)
'    strCadena = strCadena & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95)
'    Plin strCadena
'    Plin " " & Chr(124) & Space(50) & Chr(124)
'    strCadena = " " & Chr(124) & " Article #       : " & Me.TxtCodigo & Space(23) & Chr(124) & Space(1)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Fabric          : " & Mid(Me.TxtDescripcion, 1, 31) & Chr(124) & Space(1)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Yarn            : " & Mid(Me.TxtHilado, 1, 31) & Space(31 - Len(Me.TxtHilado)) & Chr(124) & Space(1)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Composition     : " & Mid(Me.TxtComposicion, 1, 31) & Space(31 - Len(Me.TxtComposicion)) & Chr(124) & Space(1)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Dyeing method   : " & Mid(Me.TxtMetodoTen, 1, 31) & Space(31 - Len(Me.TxtMetodoTen)) & Chr(124) & Space(1)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Laundry process : " & Mid(Me.TxtProcesoLav, 1, 31) & Space(31 - Len(Me.TxtProcesoLav)) & Chr(124)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Shrinkage       : " & Space(5 - Len(Mid(Me.TxtEncogLargo, 1, 5))) & Mid(Trim(Me.TxtEncogLargo), 1, 5) & "% length" & Space(1) & Space(5 - Len(Mid(Me.TxtEncogAncho, 1, 5))) & Trim(Mid(Me.TxtEncogAncho, 1, 5)) & "% width"
'    strCadena = strCadena & Space(52 - Len(strCadena)) & Chr(124)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Width           : " & Space(5 - Len(Mid(Me.TxtAncho, 1, 5)))
'    strCadena = strCadena & Mid(Trim(Me.TxtAncho), 1, 5) & "mts" & Space(6)
'    strCadena = strCadena & Space(7 - Len(Mid(Me.TxtAnchoPulg, 1, 5))) & Trim(Mid(Me.TxtAnchoPulg, 1, 5)) & "inches"
'    strCadena = strCadena & Space(52 - Len(strCadena)) & Chr(124)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Weight b/w      : " & Space(5 - Len(Mid(Me.TxtPeso, 1, 5))) & Mid(Trim(Me.TxtPeso), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ, 1, 5)) & " oz/yd2"
'    strCadena = strCadena & Space(52 - Len(strCadena)) & Chr(124)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Weight a/w      : " & Space(5 - Len(Mid(Me.TxtPeso2, 1, 5))) & Mid(Trim(Me.TxtPeso2), 1, 5) & "grms/mt2" & Space(1) & Space(5 - Len(Mid(Me.TxtPesoOZ, 1, 5))) & Trim(Mid(Me.TxtPesoOZ2, 1, 5)) & " oz/yd2"
'    strCadena = strCadena & Space(52 - Len(strCadena)) & Chr(124)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Gauge           : " & Mid(Trim(Me.TxtGalga), 1, 17) & Space(15 - Len(Mid(Me.TxtGalga, 1, 17))) & Space(1) & "Diam :" & Trim(Mid(Me.TxtDiamGalga, 1, 5))
'    strCadena = strCadena & Space(52 - Len(strCadena)) & Chr(124)
'    Plin strCadena
'
'    strCadena = " " & Chr(124) & " Colorway        : " & Mid(Me.TxtcolorWay, 1, 31) & Space(31 - Len(Me.TxtcolorWay)) & Chr(124)
'    Plin strCadena
'
'    strCadena = "  " & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95)
'    strCadena = strCadena & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95) & Chr(95)
'    Plin strCadena
    
    Plin "                 "
    Plin "                 "
    Plin "                 "
    Plin "                 "
    Plin "                 "
    'Plin Chr(12)
    
    Close #1
    
    'For i = 1 To CDbl(TxtCantidad.Text)
    oPrint.SendPrint "c:\PRUEBA.txt"
    'Next
    Set oPrint = Nothing
End Sub


Sub Act_Gramajes()
On Error GoTo errActualiza

strCadena = "Up_Actualiza_Gramajes_TelaComb '" & TxtCodigo.Text & "','" & TxtCod_Comb.Text & "'," & CDbl(TxtPeso.Text) & "," & CDbl(TxtPeso2.Text) & "," & CDbl(TxtAncho.Text)
ExecuteCommandSQL cCONNECT, strCadena

Exit Sub
errActualiza:
    MsgBox Err.Description, vbCritical, "Actrualizacion Gramajes Tela/Comb"
End Sub

Public Sub Ayuda_Comb(Opcion As Integer)
Dim rstAux As ADODB.Recordset
On Error GoTo Fin
Dim iCol As Long
    
    strSQL = "exec sm_muestra_ayuda_combinacion_tela '" & Opcion & "','" & Trim(TxtCodigo) & "','" & Trim(TxtCod_Comb) & "','" & Trim(TxtDes_Comb) & "'"
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        .Caption = "Seleccionar Comb."
        'Codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("codigo").Width = 900
        .DGridLista.Columns("Descripcion").Width = 5000
        
        If rstAux.RecordCount > 1 Then
            .Show vbModal
        Else
            Codigo = .DGridLista.Value(.DGridLista.Columns("Codigo").Index)
            Descripcion = .DGridLista.Value(.DGridLista.Columns("Descripcion").Index)
        End If
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtCod_Comb = Codigo
            TxtDes_Comb = Descripcion
            TxtPeso.Text = DevuelveCampo("select gramaje from tx_telacomb where cod_tela ='" & Trim(TxtCodigo) & "' and cod_comb='" & Codigo & "'", cCONNECT)
            TxtPeso2.Text = DevuelveCampo("select gramaje_despues_lavado from tx_telacomb where cod_tela ='" & Trim(TxtCodigo) & "' and cod_comb='" & Codigo & "'", cCONNECT)
            TxtAncho.Text = DevuelveCampo("select ancho from tx_telacomb where cod_tela ='" & Trim(TxtCodigo) & "' and cod_comb='" & Codigo & "'", cCONNECT)
            TxtPartida.SetFocus
        End If
    End With
    Codigo = "": Descripcion = ""
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Busca Comb (" & Opcion & ")"
End Sub

Private Sub TxtPartida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Busca_Partida
End If
End Sub

Public Sub Busca_Partida()
Dim rstAux As ADODB.Recordset
On Error GoTo Fin
Dim iCol As Long
    
    strSQL = "exec ti_muestra_datos_tecnicos_tela_acabada '" & Trim(TxtCodigo) & "','" & Trim(TxtCod_Comb) & "'"
    
    With frmBusqGeneral3
        Set .oParent = Me
        .sQuery = strSQL
        .Cargar_Datos
        .Caption = "Seleccionar Partida"
        'Codigo = ".."
        Set rstAux = .DGridLista.ADORecordset
        
        .DGridLista.Columns("Partida").Width = 900
        .DGridLista.Columns("Kgs_Tenidos_1ras").Width = 900
        .DGridLista.Columns("Color").Width = 2000
        .DGridLista.Columns("Gramaje_Acab").Width = 900
        .DGridLista.Columns("Ancho_Acab").Width = 900
        .DGridLista.Columns("Encog_Ancho").Width = 900
        .DGridLista.Columns("Encog_Largo").Width = 900
        .DGridLista.Columns("Cod_UsuarioUltmodDatTec").Width = 900
        .DGridLista.Columns("Revirado").Width = 900
        .DGridLista.Columns("Gramaje_Lavado").Width = 900
                
        If rstAux.RecordCount > 1 Then
            .Show vbModal
        Else
            Codigo = .DGridLista.Value(.DGridLista.Columns("PARTIDA").Index)
        End If
        
        If Codigo <> "" And rstAux.RecordCount > 0 Then
            TxtPartida = Codigo
            'TxtPeso.Text = rstAux!Gramaje_Acab
            'TxtPeso2.Text = rstAux!Gramaje_Lavado
            'TxtPeso_Lote.Text = rstAux!Gramaje_Acab
            'TxtAncho.Text = rstAux!Ancho_Acab
            TxtPeso_Lote.Text = .DGridLista.Value(.DGridLista.Columns("Gramaje_Acab").Index)
            TxtDescripcion.SetFocus
        End If
    End With
    Codigo = "": Descripcion = ""
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
Exit Sub
Fin:
On Error Resume Next
    Unload frmBusqGeneral3
    Set frmBusqGeneral3 = Nothing
    rstAux.Close
    Set rstAux = Nothing
    MsgBox Err.Description & ", No se puede Continuar", vbExclamation + vbOKOnly, _
    "Busca Comb"
End Sub

Private Sub txtruta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call BUSCARUTA(1)
    End If
End Sub

Sub BUSCARUTA(tipo As Integer)
Dim oTipo As New frmBusqGeneral3
Dim rs As New ADODB.Recordset

Set oTipo.oParent = Me

If tipo = 1 Then
    oTipo.sQuery = "select Cod_Ruta as Codigo,Descripcion   from Tx_Tela_DatTecnicos_cabecera where cod_tela  like '%" & Trim(TxtCodigo.Text) & "%'"
ElseIf tipo = 2 Then
    oTipo.sQuery = "select Cod_Ruta as Codigo, Descripcion from Tx_Tela_DatTecnicos_cabecera where cod_tela like '%" & Trim(TxtCodigo.Text) & "%'"
End If

oTipo.Caption = "Buscar Rutas"
oTipo.Cargar_Datos

oTipo.DGridLista.Columns("Codigo").Width = 1400
oTipo.DGridLista.Columns("Descripcion").Width = 5000

If oTipo.DGridLista.RowCount > 1 Then
    oTipo.Show vbModal
Else
    Codigo = oTipo.DGridLista.Value(oTipo.DGridLista.Columns("Codigo").Index)
    Descripcion = oTipo.DGridLista.Value(oTipo.DGridLista.Columns("Descripcion").Index)
End If

If Trim(Codigo) <> "" Then
    txtruta.Text = Codigo
    txtdesruta.Text = Descripcion
    Codigo = "": Descripcion = ""
    TxtPartida.SetFocus
End If

Unload oTipo
Set oTipo = Nothing
Set rs = Nothing
End Sub


