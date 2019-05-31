VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm004_ECNNET 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equipos conectados en la misma red"
   ClientHeight    =   3645
   ClientLeft      =   540
   ClientTop       =   5670
   ClientWidth     =   6480
   Icon            =   "ECNNET.frx":0000
   LinkTopic       =   "ECNNET"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6480
   Begin ComctlLib.ListView lvwNET 
      Height          =   3645
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   6429
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ILLarge"
      SmallIcons      =   "ILSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Comments"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "5"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ImageList ILSmall 
      Left            =   2910
      Top             =   4410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483643
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNNET.frx":08CA
            Key             =   "PC"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNNET.frx":0C1C
            Key             =   "RED"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ILLarge 
      Left            =   3510
      Top             =   4410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   -2147483643
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNNET.frx":0F6E
            Key             =   "PC"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ECNNET.frx":1288
            Key             =   "RED"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu 
      Caption         =   "Vista"
      Begin VB.Menu mnuTipoVista 
         Caption         =   "Iconos Grandes"
         Index           =   0
      End
      Begin VB.Menu mnuTipoVista 
         Caption         =   "Iconos Pequeños"
         Index           =   1
      End
      Begin VB.Menu mnuTipoVista 
         Caption         =   "Lista"
         Index           =   2
      End
      Begin VB.Menu mnuTipoVista 
         Caption         =   "Detalles"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frm004_ECNNET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NETRESOURCE
   dwScope As Long
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   lpLocalName As Long
   lpRemoteName As Long
   lpComment As Long
   lpProvider As Long
End Type

Private Type NetInfo       'UDT to store Data and use it for
   dwScope As Long 'filling ListView
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   LocalName As String
   RemoteName As String
   Comment As String
   Provider As String
End Type

Private Const RESOURCE_CONTEXT = &H5

Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCETYPE_PRINT = &H2

Private Const RESOURCEUSAGE_CONTAINER = &H2

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CopyPointer2String Lib "kernel32" Alias "lstrcpyA" (ByVal NewString As String, ByVal OldString As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, ByVal lpBuffer As Long, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long

Private NI() As NetInfo
Private NR As NETRESOURCE

Private Sub Form_Load()
    GO_004_COMPUTER_SEL = Empty
    GO_004_ENU_OPC_WIN_RESULT = WD_NULL
    Call CargarRed
End Sub

Private Sub mnuTipoVista_Click(Index As Integer)
    lvwNET.View = Index
End Sub

Private Sub lvwNET_ItemClick(ByVal item As ComctlLib.ListItem)
   On Error Resume Next
   GO_004_COMPUTER_SEL = item.Text
End Sub

Private Sub lvwNET_DblClick()
    On Error Resume Next
    GO_004_COMPUTER_SEL = lvwNET.SelectedItem.Text
    GO_004_ENU_OPC_WIN_RESULT = WD_ACCEPT
    Unload Me
End Sub


'*********************************************************************************************************************************************************************
' PROCEDIMIENTOS DE USUARIO
'*********************************************************************************************************************************************************************
Private Sub CargarRed()
    lvwNET.ListItems.Clear
    Call ObtenerRedLocal
    Call CargaListView
End Sub

Private Sub CargaListView()
    Dim itmx As Object 'ListItem
    Dim picNum As Integer
    Dim i As Integer
    Dim sNombre As String
    
    For i = 3 To 5
        lvwNET.ColumnHeaders(i).Width = 0
    Next i
    lvwNET.ColumnHeaders(2).Text = "Comments"
    For i = 1 To UBound(NI)
        picNum = NI(i).dwDisplayType
        If NI(i).dwType = RESOURCETYPE_PRINT Then picNum = 4
        If picNum > 5 Then picNum = 5
        sNombre = StripSlash(NI(i).RemoteName)
        If sNombre = "" Then sNombre = NI(i).Comment
        Set itmx = lvwNET.ListItems.Add(, , sNombre)
        With itmx
            .SubItems(1) = NI(i).Comment
            .SubItems(3) = NI(i).dwType
            .SubItems(4) = NI(i).dwUsage
            .Key = NI(i).RemoteName
            .Text = sNombre
            .Icon = ILLarge.ListImages(1).Key
            .SmallIcon = ILSmall.ListImages(1).Key
        End With
    Next i
    Set itmx = Nothing
    With lvwNET
        .SortOrder = lvwAscending
        .SortKey = 0
        .Sorted = True
        .SelectedItem.Selected = False
    End With
End Sub

Private Function StripSlash(sName As String) As String
   ' Obtiene parte de la cadena despues del ultimo slash (Item.text)
    Dim a As Integer
    Dim b As Integer
    Do
        b = a
        a = InStr(a + 1, sName, "\", vbTextCompare)
    Loop While a <> 0
    StripSlash = Mid$(sName, b + 1)
End Function


Private Sub ObtenerRedLocal()
    Dim hEnum As Long
    Dim lpBuff As Long
    Dim cbBuff As Long
    Dim cCount As Long
    
    Dim p As Long
    Dim res As Long
    Dim i As Long
    
    On Error GoTo SALTO_ERROR
    ClearNr
    cbBuff = 16384
    cCount = &HFFFFFFFF
    res = WNetOpenEnum(RESOURCE_CONTEXT, RESOURCETYPE_ANY, RESOURCEUSAGE_CONTAINER, NR, hEnum)
    If res = 0 Then
        lpBuff = GlobalAlloc(GPTR, cbBuff)
        res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
        If res = 0 Then
            ReDim NI(cCount)
            p = lpBuff
            For i = 1 To cCount
                CopyMemory NR, ByVal p, LenB(NR)
                FillInfo i
                p = p + LenB(NR)
            Next i
        End If
SALTO_ERROR:
    On Error Resume Next
    If lpBuff <> 0 Then GlobalFree (lpBuff)
    WNetCloseEnum (hEnum)
    End If
End Sub

Private Sub ClearNr()
  NR.dwDisplayType = 0&
  NR.dwScope = 0&
  NR.dwType = 0&
  NR.dwUsage = 0&
  NR.lpComment = 0&
  NR.lpLocalName = 0&
  NR.lpProvider = 0&
  NR.lpRemoteName = 0&
End Sub

Private Sub FillInfo(Index As Long)
    NI(Index).dwScope = NR.dwScope
    NI(Index).dwDisplayType = NR.dwDisplayType
    NI(Index).dwType = NR.dwType
    NI(Index).dwUsage = NR.dwUsage
    NI(Index).RemoteName = PunteroToCadena(NR.lpRemoteName)
    NI(Index).LocalName = PunteroToCadena(NR.lpLocalName)
    NI(Index).Comment = PunteroToCadena(NR.lpComment)
    NI(Index).Provider = PunteroToCadena(NR.lpProvider)
End Sub

Public Function PunteroToCadena(p As Long) As String
    Dim s As String
    s = String(255, Chr$(0))
    CopyPointer2String s, p
    PunteroToCadena = Left(s, InStr(s, Chr$(0)) - 1)
End Function
