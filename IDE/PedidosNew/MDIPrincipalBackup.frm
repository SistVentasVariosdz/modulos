VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{85202277-6C76-4228-BC56-7B3E69E8D5CA}#5.0#0"; "IGToolBars50.ocx"
Begin VB.MDIForm MDIPrincipal 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   Caption         =   "Menú Principal"
   ClientHeight    =   6675
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11175
   Icon            =   "MDIPrincipalBackup.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pictureMenuAccesos 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00B4C6C3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   0
      ScaleHeight     =   445
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   0
      Top             =   0
      Width           =   3555
      Begin VB.TextBox txtBuscarOpcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   60
         TabIndex        =   1
         Text            =   " Ingrese descripción de la opción a buscar"
         ToolTipText     =   "Ingrese descripción de la opción y luego hacer clic en el ícono de buscar."
         Top             =   60
         Width           =   3135
      End
      Begin MSComctlLib.TreeView tvMenuAccesos 
         Height          =   5835
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   10292
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView tvMenuAccesosBusqueda 
         Height          =   5835
         Left            =   60
         TabIndex        =   3
         Top             =   6330
         Visible         =   0   'False
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   10292
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   30
         Top             =   30
         Width           =   3495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Left            =   3270
         Picture         =   "MDIPrincipalBackup.frx":54482
         ToolTipText     =   "Use la caja de texto para buscar una opción del sistema."
         Top             =   60
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList imageListFondo 
      Left            =   4650
      Top             =   3270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   563
      ImageHeight     =   550
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":54A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":8C15B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":F64B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":19DFD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":27C63B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":388355
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":4B160F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   4650
      Top             =   2310
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327680
      Tools           =   "MDIPrincipalBackup.frx":726247
      ToolBars        =   "MDIPrincipalBackup.frx":72625F
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3750
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":726277
            Key             =   "mancli"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":7262D5
            Key             =   "manfab"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":726333
            Key             =   "manOrg"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":726391
            Key             =   "mantra"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":7263EF
            Key             =   "mancomisin"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":72644D
            Key             =   "manBan"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":7264AB
            Key             =   "mandestino"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":726509
            Key             =   "mantippre"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipalBackup.frx":726567
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5250
      Top             =   2250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   4800
      Top             =   1530
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sopcion              As String

Private boolBuscarOpcion As Boolean

Dim treeNode             As node

Sub BorrarTablas()

    On Error Resume Next

    Dim Reg As New ADODB.Recordset

    Set Reg = Nothing
    Set Reg = Nothing
    Reg.CursorLocation = adUseClient
    Reg.Open "drop table cf_clie", cCONNECT

    Set Reg = Nothing
    Reg.CursorLocation = adUseClient
    Reg.Open "drop table CF_DES", cCONNECT

    Set Reg = Nothing
    Reg.CursorLocation = adUseClient
    Reg.Open "drop table cf_pedd", cCONNECT

    Set Reg = Nothing
    Reg.CursorLocation = adUseClient
    Reg.Open "drop table cf_pedi", cCONNECT

    Set Reg = Nothing
    Reg.CursorLocation = adUseClient
    Reg.Open "drop table CF_PEDR", cCONNECT
    Set Reg = Nothing
End Sub

Sub CambiaCaptionMenu()

    On Error GoTo hand

    Dim ctl As Control

    Dim Reg As New ADODB.Recordset

    Reg.CursorLocation = adUseClient
    Reg.Open "select Cod_Opcion,Des_Opcion,Des_Opcion_Eng from seg_opciones order by 1", conn.ConnectionString

    If Reg.RecordCount > 0 Then

        For Each ctl In MDIPrincipal.Controls

            'Debug.Print ctl.Name
            If TypeOf ctl Is Menu Then
                If Mid(ctl.Name, 1, 3) <> "ISO" Then
                    If iLanguage = 1 Then
                        If DevuelveCampo("select Des_Opcion from seg_opciones where Cod_Opcion='" & ctl.Name & "'", sconnect) <> "" Then
                            ctl.Caption = DevuelveCampo("select Des_Opcion from seg_opciones where Cod_Opcion='" & ctl.Name & "'", sconnect)
                        End If

                    Else

                        If DevuelveCampo("select Des_Opcion_Eng from seg_opciones where Cod_Opcion='" & ctl.Name & "'", sconnect) <> "" Then
                            ctl.Caption = DevuelveCampo("select Des_Opcion_Eng from seg_opciones where Cod_Opcion='" & ctl.Name & "'", sconnect)
                        End If
                    End If
                End If
            End If

        Next

    End If

    Set Reg = Nothing

    Exit Sub

hand:
    ErrorHandler Err, "CambiaCaptionMenu"
    Set Reg = Nothing
End Sub

Private Sub ActEstCli_Click()
    EjecutaOpcionMenu "ActEstCli", Me.perfil, Me.pEmpresa
End Sub

'Private Sub Cierre_Click()
'EjecutaOpcionMenu "Cierre", Me.perfil, Me.pEmpresa
'End Sub

'Private Sub ConCierre_Click()
'EjecutaOpcionMenu "concierre", Me.perfil, Me.pEmpresa
'End Sub

Private Sub CondVent_Click()
    EjecutaOpcionMenu "CondVent", Me.perfil, Me.pEmpresa
End Sub

Private Sub DespPrendas_Click()
    EjecutaOpcionMenu "DespPrend", Me.perfil, Me.pEmpresa
End Sub

Private Sub Dscto_Click()
    EjecutaOpcionMenu "Dscto", Me.perfil, Me.pEmpresa
End Sub

Private Sub frmGruposReq_Click()
    EjecutaOpcionMenu "frmgruposreq", Me.perfil, Me.pEmpresa
End Sub

Private Sub frmImpSit_Click()
    EjecutaOpcionMenu "frmImpSit", Me.perfil, Me.pEmpresa
End Sub

Private Sub ISO4_Click(Index As Integer)

    sopcion = "ISO4" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO42_Click(Index As Integer)
    sopcion = "ISO42" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO5_Click(Index As Integer)
    sopcion = "ISO5" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO54_Click(Index As Integer)
    sopcion = "ISO54" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO55_Click(Index As Integer)
    sopcion = "ISO55" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO56_Click(Index As Integer)
    sopcion = "ISO56" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO6_Click(Index As Integer)
    sopcion = "ISO6" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO62_Click(Index As Integer)
    sopcion = "ISO62" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO7_Click(Index As Integer)
    sopcion = "ISO7" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO72_Click(Index As Integer)
    sopcion = "ISO72" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO73_Click(Index As Integer)
    sopcion = "ISO73" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO74_Click(Index As Integer)
    sopcion = "ISO74" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO75_Click(Index As Integer)
    sopcion = "ISO75" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO8_Click(Index As Integer)
    sopcion = "ISO8" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO82_Click(Index As Integer)
    sopcion = "ISO82" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

Private Sub ISO85_Click(Index As Integer)
    sopcion = "ISO85" & CStr(Index)
    EjecutaOpcionMenu sopcion, Me.perfil, Me.pEmpresa

End Sub

'Private Sub ISO41_Click()
'    EjecutaOpcionMenu "ISO41", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO42_Click()
'    EjecutaOpcionMenu "ISO42", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO5_Click()
'    EjecutaOpcionMenu "ISO5", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO61_Click()
'    EjecutaOpcionMenu "ISO61", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO621_Click()
'    EjecutaOpcionMenu "ISO621", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO622_Click()
'    EjecutaOpcionMenu "ISO622", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO63_Click()
'    EjecutaOpcionMenu "ISO63", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO64_Click()
'    EjecutaOpcionMenu "ISO64", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO731_Click()
'    EjecutaOpcionMenu "ISO731", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO732_Click()
'        EjecutaOpcionMenu "ISO732", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO733_Click()
'    EjecutaOpcionMenu "ISO733", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO734_Click()
'    EjecutaOpcionMenu "ISO734", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO735_Click()
'    EjecutaOpcionMenu "ISO735", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO736_Click()
'    EjecutaOpcionMenu "ISO736", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO737_Click()
'    EjecutaOpcionMenu "ISO737", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISO8_Click()
'    EjecutaOpcionMenu "ISO8", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOComCli_Click()
'    EjecutaOpcionMenu "ISOComCli", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOConDis_Click()
'EjecutaOpcionMenu "ISOConDis", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOControl_Click()
'EjecutaOpcionMenu "ISOControl", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISODDP_Click()
'
'End Sub
'
'Private Sub ISODetReq_Click()
'EjecutaOpcionMenu "ISODetReq", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOIdeTra_Click()
'EjecutaOpcionMenu "ISOIdeTra", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOInfor_Click()
'EjecutaOpcionMenu "ISOInfor", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOPlaRea_Click()
'EjecutaOpcionMenu "ISOPlaRea", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOPrePro_Click()
'EjecutaOpcionMenu "ISOPrePro", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOProce_Click()
'EjecutaOpcionMenu "ISOProce", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOProCli_Click()
'EjecutaOpcionMenu "ISOProCli", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISORevReq_Click()
'EjecutaOpcionMenu "ISORevReq", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOValPro_Click()
'EjecutaOpcionMenu "ISOValPro", Me.perfil, Me.pEmpresa
'End Sub
'
'Private Sub ISOVerPro_Click()
'EjecutaOpcionMenu "ISOVerPro", Me.perfil, Me.pEmpresa
'End Sub

Private Sub LugEntr_Click()
    EjecutaOpcionMenu "LugEntr", Me.perfil, Me.pEmpresa
End Sub

Private Sub manban_Click()
    EjecutaOpcionMenu "manBan", Me.perfil, Me.pEmpresa
End Sub

Private Sub mancli_Click()
    EjecutaOpcionMenu "MANCLI", Me.perfil, Me.pEmpresa
End Sub

Private Sub mancomisin_Click()
    EjecutaOpcionMenu "manComisin", Me.perfil, Me.pEmpresa
End Sub

Private Sub manDestino_Click()
    EjecutaOpcionMenu "manDestino", Me.perfil, Me.pEmpresa
End Sub

Private Sub mandivpre_Click()
    EjecutaOpcionMenu "mandivpre", Me.perfil, Me.pEmpresa
End Sub

Private Sub manDocs_Click()
    EjecutaOpcionMenu "MANDocs", Me.perfil, Me.pEmpresa
End Sub

Private Sub manFab_Click()
    EjecutaOpcionMenu "MANfab", Me.perfil, Me.pEmpresa
End Sub

Private Sub manfun_Click()
    EjecutaOpcionMenu "manfun", Me.perfil, Me.pEmpresa
End Sub

Private Sub manMon_Click()
    EjecutaOpcionMenu "manMon", Me.perfil, Me.pEmpresa
End Sub

Private Sub manmotatr_Click()
    EjecutaOpcionMenu "manmotatr", Me.perfil, Me.pEmpresa
End Sub

Private Sub manopc_Click()
    EjecutaOpcionMenu "manopc", Me.perfil, Me.pEmpresa
End Sub

Private Sub manorg_Click()
    EjecutaOpcionMenu "manOrg", Me.perfil, Me.pEmpresa
End Sub

Private Sub manPagemb_Click()
    EjecutaOpcionMenu "manPagEmb", Me.perfil, Me.pEmpresa
End Sub

Private Sub manper_Click()
    EjecutaOpcionMenu "manper", Me.perfil, Me.pEmpresa
End Sub

Private Sub manPOObs_Click()
    EjecutaOpcionMenu "manPoObs", Me.perfil, Me.pEmpresa
End Sub

Private Sub manseg_Click()
    EjecutaOpcionMenu "MANSEG", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantal_Click()
    EjecutaOpcionMenu "manTal", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantCargos_Click()
    EjecutaOpcionMenu "MANtcargos", Me.perfil, Me.pEmpresa

End Sub

Private Sub mantGruTal_Click()
    EjecutaOpcionMenu "mantgrutal", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantHil_Click()
    EjecutaOpcionMenu "manthil", Me.perfil, Me.pEmpresa
End Sub

Private Sub manTipEmb_Click()
    EjecutaOpcionMenu "manTipEmb", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantippre_Click()
    EjecutaOpcionMenu "manTipPre", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantra_Click()
    EjecutaOpcionMenu "manTra", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantitm_Click()
    EjecutaOpcionMenu "mantitm", Me.perfil, Me.pEmpresa
End Sub

Private Sub mantItem_Click()
    EjecutaOpcionMenu "mantitem", Me.perfil, Me.pEmpresa
End Sub

Private Sub MantTelas_Click()
    EjecutaOpcionMenu "manttelas", Me.perfil, Me.pEmpresa
End Sub

Private Sub manunimed_Click()
    EjecutaOpcionMenu "manUniMed", Me.perfil, Me.pEmpresa
End Sub

Private Sub manusu_Click()
    EjecutaOpcionMenu "manusu", Me.perfil, Me.pEmpresa
End Sub

Private Sub mConfecDr_Click()
    EjecutaOpcionMenu "mConfecDr", Me.perfil, Me.pEmpresa
End Sub
    
    

Private Sub imgBuscar_Click()

    If Trim(txtBuscarOpcion) <> "" Then
        tvMenuAccesosBusqueda.Nodes.Clear
        Call BUSCAR(tvMenuAccesos, txtBuscarOpcion, False, False)
    End If

End Sub

Sub BuscarTree(Tv As TreeView, _
               texto As String, _
               Optional bMayMin As Boolean = False, _
               Optional bCadenaComp As Boolean = True)

    Dim Nodo        As node

    Dim i           As Integer

    Dim bEncontrado As Boolean

    Dim istart      As Long
    
    If Tv.Tag = "" Then Tv.Tag = 0
    istart = CInt(Tv.Tag) + 1

    Dim sNodo As String
    
    With Tv

        For i = istart To .Nodes.count
            sNodo = .Nodes(i).Text
            
            If bCadenaComp = True Then
                If bMayMin = False Then
                    If LCase(texto) = LCase(sNodo) Then bEncontrado = True
                ElseIf bMayMin Then

                    If texto = sNodo Then
                        bEncontrado = True
                    End If
                End If

            Else

                Dim pos As Integer

                pos = InStr(1, .Nodes(i).Text, texto, Abs(Not bMayMin))

                If pos <> 0 Then
                    bEncontrado = True
                End If
            End If
            
            ' si se encontró ..
            If bEncontrado Then
               
                Set Nodo = .Nodes(i)

                ''
                ''               Do Until Nodo.Parent Is Nothing
                ''                  Nodo.Parent.Expanded = True
                ''                  Set Nodo = Nodo.Parent
                ''
                ''
                ''               Loop
                ''
                ''               .Nodes.Item(i).Selected = True
                ''
                ''               .SetFocus
                ''               Tv.Tag = .Nodes.Item(i).Index
                '   Exit Sub
                If Not Nodo.Parent Is Nothing Then

                    '   tvMenuAccesosBusqueda.Nodes.Add , , Nodo.Key, Nodo.Text
                    '  Else
                    If encuentraTreePadre(tvMenuAccesosBusqueda, Nodo.Parent.Text) = False Then
                
                        tvMenuAccesosBusqueda.Nodes.Add , , Nodo.Parent.Key, Nodo.Parent.Text
                    End If

                    tvMenuAccesosBusqueda.Nodes.Add Nodo.Parent.Key, tvwChild, Nodo.Key, Nodo.Text
                End If
                
                bEncontrado = False
            End If

        Next

    End With
    
    '   MsgBox "Dato no encontrado", vbExclamation
    Tv.Tag = 0
End Sub

Private Function encuentraTreePadre(Tv As TreeView, texto As String) As Boolean

    Dim i           As Integer

    Dim sNodo       As String

    Dim bEncontrado As Boolean

    bEncontrado = False

    With Tv
    
        If .Nodes.count > 0 Then

            For i = 1 To .Nodes.count
            
                If .Nodes(i).Parent Is Nothing Then
                    sNodo = .Nodes(i).Text
               
                    If LCase(texto) = LCase(sNodo) Then
                        bEncontrado = True

                        Exit For

                    End If
                End If
               
            Next i

        End If

    End With

    encuentraTreePadre = bEncontrado
End Function

Sub BUSCAR(Tv As TreeView, _
           texto As String, _
           Optional bMayMin As Boolean = False, _
           Optional bCadenaComp As Boolean = True)

    Dim Nodo        As node

    Dim i           As Integer

    Dim bEncontrado As Boolean

    Dim istart      As Long
    
    If Tv.Tag = "" Then Tv.Tag = 0
    istart = CInt(Tv.Tag) + 1

    Dim sNodo As String
    
    With Tv

        For i = istart To .Nodes.count
            sNodo = .Nodes(i).Text
            
            If bCadenaComp = True Then
                If bMayMin = False Then
                    If LCase(texto) = LCase(sNodo) Then bEncontrado = True
                ElseIf bMayMin Then

                    If texto = sNodo Then
                        bEncontrado = True
                    End If
                End If

            Else

                Dim pos As Integer

                pos = InStr(1, .Nodes(i).Text, texto, Abs(Not bMayMin))

                If pos <> 0 Then
                    bEncontrado = True
                End If
            End If
            
            ' si se encontró ..
            If bEncontrado Then
               
                Set Nodo = .Nodes(i)
               
                Do Until Nodo.Parent Is Nothing
                    Nodo.Parent.Expanded = True
                    Set Nodo = Nodo.Parent
                Loop
               
                .Nodes.item(i).Selected = True
               
                .SetFocus
                Tv.Tag = .Nodes.item(i).Index

                Exit Sub
               
            Else
               
                .Nodes.item(i).Selected = False
               
            End If

        Next

    End With
    
    MsgBox "Opción de acceso no encontrada", vbOKOnly + vbExclamation, Me.Caption
    Tv.Tag = 0
End Sub

'Option Explicit
Private Sub MDIForm_Load()
MsgBox 2
    Dim f As Form

    iLanguage = CInt(GetSetting("Visuales", "Settings", "Language", "1"))
    IdiomaEtiquetas1 Me
   ' MsgBox 3
    Set f = Me
    f.Caption = Caption & "-" & NEmpresa
 
    SeteaAnoPeriodoVigente

    'RMP
    'get_accesos3 pEmpresa, perfil, f
    'get_favoritos pEmpresa, pUsuario, f, iLanguage

    'set_barra (iLanguage)

    'RMP
    'CambiaCaptionMenu

    'InitMessages 'C.A.R.
    'FrmMantEmpUsuPer.Show
    'FrmMantopciones.Show

    Dim adoRsColor As ADODB.Recordset

    Set adoRsColor = CargarRecordSetDesconectado("SELECT ColorFondo_R, ColorFondo_G, ColorFondo_B FROM dbo.SEG_Empresas WHERE Cod_Empresa ='" & vemp & "'", sconnect)

    If Not adoRsColor.EOF Then
        SSActiveToolBars1.BackColor = RGB(adoRsColor("ColorFondo_R"), adoRsColor("ColorFondo_G"), adoRsColor("ColorFondo_B"))
        SSActiveToolBars1.Font.Bold = True
    
    End If

   ' SSActiveToolBars1.BackColor = RGB(188, 204, 185)

  '  SSActiveToolBars1.BackColor = RGB(223, 224, 218) ' rea
 
    'BuildMenuBar
    'BuildOptions

    Dim tooMenu As ActiveToolBars.SSTool

    SSActiveToolBars1.ToolBars.Add "Aplicacion", ssMenuBar
    SSActiveToolBars1.ToolBars("Aplicacion").AllowCustomize = True

    Set tooMenu = SSActiveToolBars1.ToolBars("Aplicacion").Tools.Add("Ventana", ssTypeMenu)
    tooMenu.Name = "&Ventana"
    
    
  '  SSActiveToolBars1.Tools("Ventana").Menu.Tools.Add "Mostrar ventanas en cascada"
  '  SSActiveToolBars1.Tools("Ventana").Menu.Tools.Add "Mostrar ventanas en mosaico"
    Set tooMenu = SSActiveToolBars1.Tools("Ventana")
    tooMenu.Type = ssTypeMenu
    Set tooMenu = tooMenu.Menu.Tools.Add("cascada", ssTypeButton)
    tooMenu.Name = "Mostrar ventanas en cascada"
    tooMenu.DisplayStyle = ssDisplayTextOnlyAlways
    
    Set tooMenu = SSActiveToolBars1.Tools("Ventana")
    tooMenu.Type = ssTypeMenu
    Set tooMenu = tooMenu.Menu.Tools.Add("mosaico", ssTypeButton)
    tooMenu.Name = "Mostrar ventanas en mosaico"
    tooMenu.DisplayStyle = ssDisplayTextOnlyAlways
    
'    Set tooMenu = SSActiveToolBars1.Tools("Ventana")
'    tooMenu.Type = ssTypeMenu
'    Set tooMenu = tooMenu.Menu.Tools.Add("mnuOcultarOpcionesSistema", ssTypeButton)
'    tooMenu.Name = "Ocultar menú de opciones del sistema"
'    tooMenu.DisplayStyle = ssDisplayTextOnlyAlways
        
  '  SSActiveToolBars1.Tools("Ventana").Menu.Tools.Add "Ocultar menú de opciones del sistema"
    

    Set tooMenu = SSActiveToolBars1.ToolBars("Aplicacion").Tools.Add("Seguridad", ssTypeMenu)
    tooMenu.Name = "&Seguridad"
    
    Set tooMenu = SSActiveToolBars1.Tools("Seguridad")
    tooMenu.Type = ssTypeMenu
    Set tooMenu = tooMenu.Menu.Tools.Add("mnuCambioContrasenia", ssTypeButton)
    tooMenu.Name = "Cambiar contraseña"
    tooMenu.DisplayStyle = ssDisplayTextOnlyAlways
    
    Set tooMenu = SSActiveToolBars1.ToolBars("Aplicacion").Tools.Add("Ayuda", ssTypeMenu)
    tooMenu.Name = "&Ayuda"
    
    Set tooMenu = SSActiveToolBars1.Tools("Ayuda")
    tooMenu.Type = ssTypeMenu
    Set tooMenu = tooMenu.Menu.Tools.Add("mnuAcercaDe", ssTypeButton)
    tooMenu.Name = "Acerca de Gestión de Telas"
    tooMenu.DisplayStyle = ssDisplayTextOnlyAlways
    
    
    Set tooMenu = SSActiveToolBars1.ToolBars("Aplicacion").Tools.Add("Exit", ssTypeButton)
    tooMenu.Name = "&Salir del sistema"
    tooMenu.DisplayStyle = ssDisplayTextOnlyAlways
    'SSActiveToolBars1.Tools("Window").Menu.Tools.Add "Exit"

    BuildStatusBar
'MsgBox 4
    Select Case fVerResol
        Case "800x600"
            Me.Picture = imageListFondo.ListImages(1).Picture
        Case "1024x768"
            Me.Picture = imageListFondo.ListImages(2).Picture
        Case "1280x1024"
            Me.Picture = imageListFondo.ListImages(3).Picture
        Case "1600x1024"
            Me.Picture = imageListFondo.ListImages(4).Picture
        Case "1600x1200"
            Me.Picture = imageListFondo.ListImages(5).Picture
        Case "1920x1080"
            Me.Picture = imageListFondo.ListImages(6).Picture
        Case "2560x1600"
            Me.Picture = imageListFondo.ListImages(7).Picture


    End Select
  '  MsgBox 5
    boolBuscarOpcion = False
    
    '  tvMenuAccesos.SingleSel = True
    Call BuildMenuBarTree
    'Call BuildOptionsTree
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next

    With pictureMenuAccesos
        .Height = Me.Height - 100
    End With

    With tvMenuAccesos
        .Width = pictureMenuAccesos.Width - 100
        .Height = pictureMenuAccesos.Height - 100
    End With

    
    
    picFondo.Width = Me.Width
    picFondo.Height = Me.Height
 MsgBox 3
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    conn.Close
    Set conn = Nothing
End Sub

Private Sub mnuBanco_Click()
    PopupMenu mnuPopmenu
End Sub

Private Sub mnuClien_Click()
    EjecutaOpcionMenu "MANCLI", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuDesti_Click()
    'frmMotivos.Show
End Sub

Private Sub mnuCieMe_Click()

End Sub

Private Sub mnu1_Click()
    EjecutaOpcionMenu "mnu1", Me.perfil, Me.pEmpresa
End Sub

Private Sub mDspTelCru_Click()
    EjecutaOpcionMenu "mDspTelCru", Me.perfil, Me.pEmpresa
End Sub

Private Sub mEvoCosGrp_Click()
    EjecutaOpcionMenu "mEvoCosGrp", Me.perfil, Me.pEmpresa
End Sub

Private Sub mLisOrdCrt_Click()
    EjecutaOpcionMenu "mLisOrdCrt", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuactinf_Click()
    EjecutaOpcionMenu "mnuactinf", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuactinfs_Click()
    EjecutaOpcionMenu "mnuactinfs", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuActTic_Click()
    EjecutaOpcionMenu "mnuActTic", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuanxcon_Click()
    EjecutaOpcionMenu "mnuanxcon", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuAOpDia_Click()
    EjecutaOpcionMenu "mnuAOpDia", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuAutorLT_Click()
    EjecutaOpcionMenu "mnuAutorLT", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuAutorPg_Click()
    EjecutaOpcionMenu "mnuAutorPg", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuautsob_Click()
    EjecutaOpcionMenu "mnuautsob", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuAvaAca_Click()
    EjecutaOpcionMenu "mnuAvaAca", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuAvaCn2_Click()
    EjecutaOpcionMenu "mnuAvaCn2", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuAvaGer_Click()
    EjecutaOpcionMenu "mnuAvaGer", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuAvnAcab_Click()
    EjecutaOpcionMenu "mnuAvnAcab", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuBiHorE_Click()
    EjecutaOpcionMenu "mnuBiHorE", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnubustel_Click()
    EjecutaOpcionMenu "mnubustel", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuCancDoc_Click()
    EjecutaOpcionMenu "mnuCancDoc", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuCapVen_Click()
    EjecutaOpcionMenu "mnuCapVen", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnucarcol_Click()
    EjecutaOpcionMenu "mnucarcol", Me.perfil, Me.pEmpresa
End Sub
            
Private Sub mnuCambioContrasenia_Click()
     EjecutaOpcionMenu "mPWD", Me.perfil, Me.pEmpresa
End Sub
Private Sub mnuAcercaDe_Click()
    frmAcercaDe.Show 1
End Sub

Private Sub mnuCascada_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuCieCnf_Click()
    EjecutaOpcionMenu "mnuCieCnf", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuCIPla_Click()
    EjecutaOpcionMenu "mnuCIPla", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuclaitm_Click()
    EjecutaOpcionMenu "mnuclaitm", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuclasoc_Click()
    EjecutaOpcionMenu "mnuclasoc", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuCliT_Click()
    EjecutaOpcionMenu "mnuCliT", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuComp_Click()
    EjecutaOpcionMenu "mnucomp", Me.perfil, Me.pEmpresa
End Sub

'Private Sub mnuComRep_Click()
'EjecutaOpcionMenu "mnuComRep", Me.perfil, Me.pEmpresa
'End Sub

Private Sub mnuconcep_Click()
    EjecutaOpcionMenu "mnuconcep", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuConCot_Click()
    EjecutaOpcionMenu "mnuConCot", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuConFa_Click()
    EjecutaOpcionMenu "ConsFact", Me.perfil, Me.pEmpresa
End Sub

'Private Sub mnuConFF_Click()
'EjecutaOpcionMenu "mnuConFF", Me.perfil, Me.pEmpresa
'End Sub

Private Sub mnuCot_Click()

End Sub

Private Sub mnuConLet_Click()
    EjecutaOpcionMenu "mnuConLet", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuconsal_Click()
    EjecutaOpcionMenu "mnuconsal", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuConsAut_Click()
    EjecutaOpcionMenu "mnuConsAut", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuconsumo_Click()
    EjecutaOpcionMenu "mnuconsumo", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnucosconf_Click()
    EjecutaOpcionMenu "mnucosconf", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnucosgru_Click()
    EjecutaOpcionMenu "mnucosgru", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuCosSem_Click()
    EjecutaOpcionMenu "mnuCosSem", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuCotiza_Click()
    EjecutaOpcionMenu "mnuCotiza", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuCtPrMe_Click()
    EjecutaOpcionMenu "mnuCtPrMe", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnudatec_Click()
    EjecutaOpcionMenu "mnudatec", Me.perfil, Me.pEmpresa
End Sub

'Private Sub mnuDeliv_Click()
'GeneraReportes DeliverySummary
'End Sub

Private Sub mnupocolest_Click()
    EjecutaOpcionMenu "mnupocol", Me.perfil, Me.pEmpresa
End Sub

'Private Sub mnuDeOpe_Click()
'EjecutaOpcionMenu "MnuDeOpe", Me.perfil, Me.pEmpresa
'End Sub

Private Sub mnuEfic_Click()
    EjecutaOpcionMenu "mnuEfic", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuEfiDia_Click()
    EjecutaOpcionMenu "mnuEfiDia", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuEntPen_Click()
    EjecutaOpcionMenu "mnuEntPen", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuentreg_Click()
    EjecutaOpcionMenu "mnuentreg", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuESemCS_Click()
    EjecutaOpcionMenu "mnuESemCS", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuEstClC_Click()
    EjecutaOpcionMenu "mnuEstClC", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuEstDia_Click()
    EjecutaOpcionMenu "mnuEstDia", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuestprop_Click()
    EjecutaOpcionMenu "mnuestprop", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuEstSem_Click()
    EjecutaOpcionMenu "mnuEstSem", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuetiq_Click()
    EjecutaOpcionMenu "mnuetiq", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuExcel_Click()
    Shell "C:\Archivos de programa\Microsoft Office\Office10\excel.EXE", vbNormalFocus
End Sub

'Private Sub mnuExplorer_Click()
'    Shell "explorer.exe", vbNormalFocus
'End Sub

Private Sub mnuextmar_Click()
    EjecutaOpcionMenu "mnuextmar", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnufacdet_Click()
    EjecutaOpcionMenu "mnufacdet", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuFacPro_Click()
    EjecutaOpcionMenu "mnuFacPro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuFamT_Click()
    EjecutaOpcionMenu "mnuFamT", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuFlujopd_Click()
    EjecutaOpcionMenu "mnuFlujopd", Me.perfil, Me.pEmpresa
End Sub

'Private Sub mnuforecast_Click()
'    GeneraReportes Forecast
'End Sub

Private Sub mnugalgas_Click()
    EjecutaOpcionMenu "mnugalgas", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnugamas_Click()
    EjecutaOpcionMenu "mnugamas", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuGenCoa_Click()
    EjecutaOpcionMenu "mnuGenCoa", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnugenmar_Click()
    EjecutaOpcionMenu "mnugenmar", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuGenPar_Click()
    EjecutaOpcionMenu "mnuGenPar", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuGrpSeg_Click()
    EjecutaOpcionMenu "mnuGrpSeg", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuGrupo_Click()
    EjecutaOpcionMenu "mnuGrupo", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuGrupoL_Click()
    EjecutaOpcionMenu "mnuGrupoL", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnugrupreg_Click()
    EjecutaOpcionMenu "mnugrupreg", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuGrupoT_Click()
    EjecutaOpcionMenu "mnuGrupoT", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuGruPro_Click()
    EjecutaOpcionMenu "mnuGruPro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuguiaman_Click()
    EjecutaOpcionMenu "mnuguiaman", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuguias_Click()
    EjecutaOpcionMenu "mnuguias", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuhilcru_Click()
    EjecutaOpcionMenu "mnuhilcru", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuhilten_Click()
    EjecutaOpcionMenu "mnuhilten", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuhorario_Click()
    EjecutaOpcionMenu "mnuhorario", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuHorStk_Click()
    EjecutaOpcionMenu "mnuHorStk", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuhortra_Click()
    EjecutaOpcionMenu "mnuhortra", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuImpMas_Click()
    EjecutaOpcionMenu "mnuimpmas", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuIndSem_Click()
    EjecutaOpcionMenu "mnuIndSem", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuintcol_Click()
    EjecutaOpcionMenu "mnuintcol", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnulist_Click()
    EjecutaOpcionMenu "mnulist", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuKarCtd_Click()
    EjecutaOpcionMenu "mnuKarCtd", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnukardex_Click()
    EjecutaOpcionMenu "mnukardex", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuKarMTA_Click()
    EjecutaOpcionMenu "mnuKarMTA", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuKarTer_Click()
    EjecutaOpcionMenu "mnuKarTer", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuLecEsp_Click()
    EjecutaOpcionMenu "mnuLecEsp", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuLecTic_Click()
    EjecutaOpcionMenu "mnuLecTic", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuLetra_Click()
    EjecutaOpcionMenu "mnuletra", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuLinPro_Click()
    EjecutaOpcionMenu "mnuLinPro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumantalm_Click()
    EjecutaOpcionMenu "mnumantalm", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumantp_Click()
    EjecutaOpcionMenu "mnumantp", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuMaqT_Click()
    EjecutaOpcionMenu "mnuMaqT", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuMatReq_Click()
    EjecutaOpcionMenu "mnuMatReq", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuMerRet_Click()
    EjecutaOpcionMenu "mnuMerRet", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuMosaico_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnumot_Click()
    EjecutaOpcionMenu "mnumot", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumovalm_Click()
    EjecutaOpcionMenu "mnumovalm", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuMovBan_Click()
    EjecutaOpcionMenu "mnuMovBan", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuMovCnf_Click()
    EjecutaOpcionMenu "mnuMovCnf", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumovhil_Click()
    EjecutaOpcionMenu "mnumovhil", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumovperm_Click()
    EjecutaOpcionMenu "mnumovperm", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuMovRoll_Click()
    EjecutaOpcionMenu "mnuMovRoll", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumovsal_Click()
    EjecutaOpcionMenu "mnumovsal", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumovsto_Click()
    EjecutaOpcionMenu "mnumovsto", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuMParMaq_Click()
    EjecutaOpcionMenu "mnuMParMaq", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumstock_Click()
    EjecutaOpcionMenu "mnumstock", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnumpr_Click()
    EjecutaOpcionMenu "mnumpr", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuND_Click()
    EjecutaOpcionMenu "mnuND", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuNoPrTra_Click()
    EjecutaOpcionMenu "mnuNoPrTra", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuNumDoc_Click()
    EjecutaOpcionMenu "mnuNumDoc", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuOrdAcab_Click()
    EjecutaOpcionMenu "mnuOrdAcab", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuOrdCnf_Click()
    EjecutaOpcionMenu "mnuOrdCnf", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuOrdComp_Click()
    EjecutaOpcionMenu "mnuOrdComp", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuordcort_Click()
    EjecutaOpcionMenu "mnuordcort", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuOrdPro_Click()
    EjecutaOpcionMenu "mnuOrdPro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuorig_Click()
    EjecutaOpcionMenu "mnuorig", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuOT_Click()
    EjecutaOpcionMenu "mnuOT", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuPakLis_Click()
    EjecutaOpcionMenu "mnuPakLis", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuParTelB_Click()
    EjecutaOpcionMenu "mnuParTelB", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuPartida_Click()
    EjecutaOpcionMenu "mnuPartida", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuPDCorte_Click()
    EjecutaOpcionMenu "mnuPDCorte", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuPedSit_Click()
    EjecutaOpcionMenu "mnuPedSit", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuPiezas_Click()
    EjecutaOpcionMenu "mnupiezas", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnupoco_Click()
    EjecutaOpcionMenu "Mnupoco", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuposcort_Click()
    EjecutaOpcionMenu "mnuposcort", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuprecio_Click()
    EjecutaOpcionMenu "mnuprecio", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuPrePro_Click()
    EjecutaOpcionMenu "mnuPrePro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuPrgRea_Click()
    EjecutaOpcionMenu "mnuPrgRea", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuProceso_Click()
    EjecutaOpcionMenu "mnuProceso", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuProdMen_Click()
    EjecutaOpcionMenu "mnuProdMen", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuproest_Click()
    EjecutaOpcionMenu "mnuproest", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuproform_Click()
    EjecutaOpcionMenu "mnuproform", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuProHab_Click()
    EjecutaOpcionMenu "mnuProHab", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuProPro_Click()
    EjecutaOpcionMenu "mnuProPro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuProtos_Click()
    EjecutaOpcionMenu "mnuProtos", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuProvT_Click()
    EjecutaOpcionMenu "mnuProvT", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuPrRollo_Click()
    EjecutaOpcionMenu "mnuPrRollo", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuRealGrp_Click()
    EjecutaOpcionMenu "mnuRealGrp", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuRecCon_Click()
    EjecutaOpcionMenu "mnuRecCon", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuRegCom_Click()
    EjecutaOpcionMenu "mnuRegCom", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuRegDir_Click()
    EjecutaOpcionMenu "mnuRegDir", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuRegIn_Click()

    Dim frmShowTG_PurOrd1 As frmShowTG_PurOrd
    
    Set frmShowTG_PurOrd1 = New frmShowTG_PurOrd
    Load frmShowTG_PurOrd1
    Set frmShowTG_PurOrd1.oParent = Me
    frmShowTG_PurOrd1.Show
    
End Sub

Private Sub mnuRepDRB_Click()
    EjecutaOpcionMenu "mnuRepDRB", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuRprConf_Click()
    EjecutaOpcionMenu "mnuRprConf", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnurptingc_Click()
    EjecutaOpcionMenu "mnurptingc", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnurvsreal_Click()
    EjecutaOpcionMenu "mnurvsreal", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuResAct_Click()
    EjecutaOpcionMenu "mnuResAct", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuResAsi_Click()
    EjecutaOpcionMenu "mnuResAsi", Me.perfil, Me.pEmpresa
End Sub

'Private Sub mnuResDe_Click()
'GeneraReportes TrackingReporteDetail
'End Sub

Private Sub mnureverr_Click()
    EjecutaOpcionMenu "mnureverr", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuRptIna_Click()
    EjecutaOpcionMenu "mnuRptIna", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuSalir_Click()

    End

    Unload Me
End Sub

Private Sub mnuSeman_Click()

End Sub

Private Sub mnusalnreg_Click()
    EjecutaOpcionMenu "mnusalnreg", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuSecCon_Click()
    EjecutaOpcionMenu "mnuSecCon", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuServTen_Click()
    EjecutaOpcionMenu "mnuServTen", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuSitGlob_Click()
    EjecutaOpcionMenu "mnuSitGlob", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnusolid_Click()
    EjecutaOpcionMenu "mnusolid", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuSPProv_Click()
    EjecutaOpcionMenu "mnuSPProv", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuStkCor_Click()
    EjecutaOpcionMenu "mnuStkCor", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuStkCos_Click()
    EjecutaOpcionMenu "mnuStkCos", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuStkCtd_Click()
    EjecutaOpcionMenu "mnuStkCtd", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnustkfam_Click()
    EjecutaOpcionMenu "mnustkfam", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuStkHid_Click()
    EjecutaOpcionMenu "mnuStkHid", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuStkSald_Click()
    EjecutaOpcionMenu "mnuStkSald", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuStkTed_Click()
    EjecutaOpcionMenu "mnuStkTed", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuStocks_Click()
    EjecutaOpcionMenu "mnuStocks", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnusubmar_Click()
    EjecutaOpcionMenu "mnusubmar", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutarifado_Click()
    EjecutaOpcionMenu "mnutarifa", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutelaca_Click()
    EjecutaOpcionMenu "mnutelaca", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutelcru_Click()
    EjecutaOpcionMenu "mnutelcru", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuTelSer_Click()
    EjecutaOpcionMenu "mnuTelSer", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuTelServ_Click()
    EjecutaOpcionMenu "mnuTelServ", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuTelVal_Click()
    EjecutaOpcionMenu "mnuTelVal", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutipanx_Click()
    EjecutaOpcionMenu "mnutipanx", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuTipCam_Click()
    EjecutaOpcionMenu "mnutipcam", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuTipComp_Click()
    EjecutaOpcionMenu "mnutipcomp", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutipdoc_Click()
    EjecutaOpcionMenu "mnutipdoc", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutipmov_Click()
    EjecutaOpcionMenu "mnutipmov", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutippro_Click()
    EjecutaOpcionMenu "mnutippro", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutipraya_Click()
    EjecutaOpcionMenu "mnutipraya", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutiprec_Click()
    EjecutaOpcionMenu "mnutiprec", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutit_Click()
    EjecutaOpcionMenu "mnutit", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnutrab_Click()
    EjecutaOpcionMenu "mnutrab", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuTrans_Click()
    EjecutaOpcionMenu "mnuTrans", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuTraPag_Click()
    EjecutaOpcionMenu "mnuTraPag", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuTurTra_Click()
    EjecutaOpcionMenu "mnuTurTra", Me.perfil, Me.pEmpresa
End Sub

Private Sub mnuultimos_Click()

    On Error Resume Next

    Screen.MousePointer = vbHourglass

    DoEvents

    'BorrarTablas
    'EjecutaDBF2SQL
    Dim Reg As New ADODB.Recordset

    Reg.CursorLocation = adUseClient
    Reg.Open "up_migracion", cCONNECT
    'EjecutaMigracionSQLtoDBF2
    'EjecutaMigracionSQLtoDBF2
    'BorrarTablas
    Set Reg = Nothing
    MsgBox "El proceso ha terminado", vbInformation
    Screen.MousePointer = vbDefault

    Exit Sub

hand:
    Set Reg = Nothing
    ErrorHandler Err, "mnuultimos_Click"
    Screen.MousePointer = vbDefault
End Sub

Public Sub mnuWizPO_Click()
    '    Dim frmNewWizard As frmWizard
    '    Set frmNewWizard = New frmWizard
    '    Load frmNewWizard
    '    Set frmNewWizard.oParent = Me
    '    frmNewWizard.Show
End Sub

Private Sub mProdAcab_Click()
    EjecutaOpcionMenu "mProdAcab", Me.perfil, Me.pEmpresa
End Sub

Private Sub mSegGrpGlb_Click()
    EjecutaOpcionMenu "mSegGrpGlb", Me.perfil, Me.pEmpresa
End Sub

Private Sub mStkAcab_Click()
    EjecutaOpcionMenu "mStkAcab", Me.perfil, Me.pEmpresa
End Sub

Private Sub mStkHilTel_Click()
    EjecutaOpcionMenu "mStkHilTel", Me.perfil, Me.pEmpresa
End Sub

Private Sub mStkTelCru_Click()
    EjecutaOpcionMenu "mStkTelCru", Me.perfil, Me.pEmpresa
End Sub

'Private Sub RepDelDet_Click()
'EjecutaOpcionMenu "REPDELDET", Me.perfil, Me.pEmpresa
'End Sub

'Private Sub RepTra_Click()
'EjecutaOpcionMenu "REPTRA", Me.perfil, Me.pEmpresa
'End Sub

Private Sub StaOrdComp_Click()
    EjecutaOpcionMenu "StaOrdComp", Me.perfil, Me.pEmpresa
End Sub

Private Sub TipOrdComp_Click()
    EjecutaOpcionMenu "TipOrdComp", Me.perfil, Me.pEmpresa
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next

    ' PopupMenu mnuPopmenu

    ' Select Case Button.Key
    '     Case "PRINT"
    '         Me.ActiveForm.Imprimir
    '     Case "CLOSE"
    '         Me.ActiveForm.Cerrar
    '     Case "EXIT"
    '         Unload Me
    ' End Select
End Sub

Public Property Get pUsuario() As Variant
    pUsuario = vusu
End Property

Public Property Let pUsuario(ByVal vnuevo As Variant)
    vusu = vnuevo
End Property

Public Property Get pEmpresa() As Variant
    pEmpresa = vemp
End Property

Public Property Let pEmpresa(ByVal vnuevo1 As Variant)
    vemp = vnuevo1
End Property

Public Property Get PClave() As Variant
    PClave = vpas
End Property

Public Property Let PClave(ByVal vnuevo2 As Variant)
    vpas = vnuevo2
End Property

Public Property Get perfil() As Variant
    perfil = vper
End Property

Public Property Let perfil(ByVal vnuevo3 As Variant)
    vper = vnuevo3
End Property

Private Function get_accesos3(ByVal vcod_empresa As Variant, _
                              ByVal Vcod_perfil As Variant, _
                              ByVal f As Form)

    On Error GoTo procesaerror

    'on Error Resume Next
    Dim RS1      As ADODB.Recordset

    Dim RS2      As ADODB.Recordset

    Dim sQuery   As String

    Dim j        As Integer

    Dim vCod_App As String

    Set RS1 = New ADODB.Recordset
    RS1.CursorLocation = adUseClient
    sQuery = "SELECT * FROM SEG_ADMINISTRACION WHERE COD_PERFIL='" & Vcod_perfil & "'  AND COD_EMPRESA='" & vcod_empresa & "'"
    'RS1.ActiveConnection = conn
    RS1.Open sQuery, conn.ConnectionString

    Set RS2 = New ADODB.Recordset
    RS2.CursorLocation = adUseClient

    'Opciones tipo Carpeta
    'RS2.ActiveConnection = conn
    If Not (RS1.BOF And RS1.EOF) Then

        For j = 1 To RS1.RecordCount
            vCod_App = RS1!Cod_Aplicacion
            RS2.Open "Sp_opciones2 '" & vCod_App & "','" & Vcod_perfil & "','" & vcod_empresa & "'", conn.ConnectionString

            If Not (RS2.BOF And RS2.EOF) Then
                RS2.MoveFirst

                While Not RS2.EOF

                    mnu_invisible RS2!Cod_Opcion, f
                    RS2.MoveNext

                Wend

            End If

            RS2.Close
            RS1.MoveNext
        Next j

    End If

    RS1.Close
    'Desactivar Aplicaciones no autorizadas
    sQuery = "SELECT NOM_MENU FROM SEG_APLICACION WHERE COD_APLICACION NOT IN (SELECT distinct(cod_aplicacion) FROM SEG_ADMINISTRACION WHERE COD_PERFIL='" & Vcod_perfil & "'  AND COD_EMPRESA='" & vcod_empresa & "')"
    RS1.Open sQuery

    If Not (RS1.BOF And RS1.EOF) Then

        For j = 1 To RS1.RecordCount
            mnu_invisible RS1!nom_menu, f
            RS1.MoveNext
        Next j

    End If

    Set RS1 = Nothing
    Set RS2 = Nothing

    Exit Function

procesaerror:
    ErrorHandler Err, "get_accesos3"

End Function

Private Sub mnu_invisible(ByVal sName As Variant, ByVal f As Form)

    Dim ctl As Control, mnu As Menu

    For Each ctl In f.Controls

        If TypeOf ctl Is Menu Then

            'If LTrim(RTrim(UCase(sname))) = "MNUACAB" Then Stop
            If LTrim(RTrim(UCase(sName))) = LTrim(RTrim(UCase(ctl.Name))) Then
                ctl.Visible = False

                Exit For

            End If
        End If

    Next ctl

End Sub

Private Sub mnu_OPCION(ByVal f As Form)

    'Captura los name y caption del menu y los inserta en la tabla Tmp_Opcion
    Dim ctl As Control, mnu As Menu

    For Each ctl In f.Controls

        If TypeOf ctl Is Menu Then

            xname = ctl.Name
            xcaption = ctl.Caption
            sQuery = "insert into tmp_opcion (name,caption) values ('" & xname & "','" & xcaption & "')"
            conn.Execute sQuery
            'End If
        End If

    Next ctl

End Sub

Private Function get_favoritos(ByVal vcod_empresa As Variant, _
                               ByVal Vcod_usuario As Variant, _
                               ByVal f As Form, _
                               ByVal iLanguage As String)
    Set RS1 = New ADODB.Recordset
    sQuery = "SELECT A.COD_OPCION,A.ICONO,A.DES_OPCION,A.DES_OPCION_ENG  FROM SEG_OPCIONES A,SEG_FAVORITOS B WHERE A.COD_OPCION=B.COD_OPCION AND B.COD_USUARIO='" & Vcod_usuario & "'  AND B.COD_EMPRESA='" & vcod_empresa & "'"
    RS1.ActiveConnection = conn
    RS1.CursorType = adOpenStatic
    RS1.Open sQuery

    If Not (RS1.BOF And RS1.EOF) Then

        With Toolbar1

            For j = 1 To RS1.RecordCount
                xkey = LTrim(RTrim(RS1!Cod_Opcion))
                ximg = LCase(RS1!icono)

                If iLanguage = "1" Then
                    xtip = RS1!des_opcion
                Else
                    xtip = RS1!Des_Opcion_Eng
                End If

                .Buttons.Add j, xkey, "", , ximg
                .Buttons.item(j).ToolTipText = xtip
                RS1.MoveNext
            Next j

        End With

    End If

End Function

Private Sub mnuContext_Click()

    If mnuContext.Caption = "Agregar" Then
        mnuContext.Caption = "Quitar"
    Else
        mnuContext.Caption = "Agregar"
    End If

End Sub

Private Sub Toolbar1_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuPopmenu
    End If

End Sub

Public Property Get NEmpresa() As Variant
    NEmpresa = vemp1
End Property

Public Property Let NEmpresa(ByVal vnuevo1 As Variant)
    vemp1 = vnuevo1
End Property

'Private Sub set_barra(iLanguage As String)
'Dim Pan As Panel
' For Each Panel In StatusBar1.Panels
'   If iLanguage = "2" Then
'       Panel.Text = Panel.Tag
'   End If
' Next Panel
' StatusBar1.Panels.Item(1).Text = StatusBar1.Panels.Item(1).Text & NEmpresa
' StatusBar1.Panels.Item(2).Text = StatusBar1.Panels.Item(2).Text & pUsuario
' StatusBar1.Panels.Item(4).Text = StatusBar1.Panels.Item(4).Text & ComputerName
' StatusBar1.Panels.Item(3).Text = StatusBar1.Panels.Item(3).Text & Fecha_Hora_Conexion
'End Sub

Private Sub BuildStatusBar()

    Dim tooStat As ActiveToolBars.SSTool
    
    SSActiveToolBars1.ToolBars.Add "StatusBar", ssStandard
    SSActiveToolBars1.ToolBars("StatusBar").DockedStatus = ssDockedBottom
    SSActiveToolBars1.ToolBars("StatusBar").DockFlags = ssPositionLocked
    
    Set tooStat = SSActiveToolBars1.ToolBars("StatusBar").Tools.Add("nEmpresa", ssTypeLabel)
    tooStat.Name = "EMPRESA : " & NEmpresa
    tooStat.DisplayStyle = ssDisplayTextOnlyAlways
    'tooStat.Picture =
    Set tooStat = SSActiveToolBars1.ToolBars("StatusBar").Tools.Add("nUsuario", ssTypeLabel)
    tooStat.Name = "USUARIO : " & pUsuario
    tooStat.DisplayStyle = ssDisplayTextOnlyAlways
    
    Set tooStat = SSActiveToolBars1.ToolBars("StatusBar").Tools.Add("nConexion", ssTypeLabel)
    tooStat.Name = "CONEXION : " & Fecha_Hora_Conexion
    tooStat.DisplayStyle = ssDisplayTextOnlyAlways
    
    Set tooStat = SSActiveToolBars1.ToolBars("StatusBar").Tools.Add("nEquipo", ssTypeLabel)
    tooStat.Name = "EQUIPO : " & ComputerName
    tooStat.DisplayStyle = ssDisplayTextOnlyAlways
    'ssTypeStateButton
    Set tooStat = SSActiveToolBars1.ToolBars("StatusBar").Tools.Add("nFecha", ssTypeLabel)
    tooStat.Name = Format(Date, "dd/mm/yyyy")
    tooStat.DisplayStyle = ssDisplayTextOnlyAlways
    
    SSActiveToolBars1.ToolBars("StatusBar").Tools.Add "separator", , 2
    SSActiveToolBars1.ToolBars("StatusBar").Tools.Add "separator", , 4
    SSActiveToolBars1.ToolBars("StatusBar").Tools.Add "separator", , 6
    SSActiveToolBars1.ToolBars("StatusBar").Tools.Add "separator", , 8
    
    'tooStat.State =
End Sub

Private Sub BuildMenuBarTree()

    Dim rstOpts      As ADODB.Recordset

    Dim count        As Integer

    Dim ArrayNodos() As node ' array de nodos

    'strSql = "EXEC SEL_PERMISOS_APLICACIONES '" & vper & "', '" & vemp & "'"
    
    strSql = "EXEC SEL_PERMISOS_APLICACIONES " & vbNewLine
    strSql = strSql & "@cod_empresa='" & vemp & "'" & vbNewLine
    strSql = strSql & ",@COD_PERFIL='" & vper & "'" & vbNewLine
    
    Set rstOpts = CargarRecordSetDesconectado(strSql, sconnect)
    
    If rstOpts.RecordCount = 0 Then Exit Sub
    
    Dim i As Integer

    i = 1

    With rstOpts
        count = .RecordCount

        If count > 0 Then
            .MoveFirst
            tvMenuAccesos.Nodes.Clear ' limpia el TV

            Do Until .EOF
                '        Set tooMenu = SSActiveToolBars1.ToolBars("Aplicacion").Tools.Add(!Cod_Aplicacion, ssTypeMenu)
                '       tooMenu.Name = IIf(iLanguage = 1, !Des_Aplicacion, !Des_Aplicacion_ENG)
                'tooMenu.DisplayStyle = ssDisplayTextOnlyInMenus
     
                Set treeNode = tvMenuAccesos.Nodes.Add(, tvwtext, !KEYID, IIf(iLanguage = 1, !Des_Aplicacion, !Des_Aplicacion))
    
                ' treeNode.Key = !KEYID
                '  Set ArrayNodos(nivel) = tvMenuAccesos.Nodes.Add(, , , IIf(iLanguage = 1, !Des_Aplicacion, !Des_Aplicacion_ENG), i)
                Call BuildOptionsTree(!Cod_Aplicacion)

                .MoveNext
            Loop

        End If

        .Close
    End With

    Set rstOpts = Nothing
    
    '    Set tooMenu = SSActiveToolBars1.ToolBars("Aplicacion").Tools.Add("Window", ssTypeMenu)
    '    tooMenu.Name = "&Window"
    '    SSActiveToolBars1.Tools("Window").Menu.Tools.Add "Cascada"
    '    SSActiveToolBars1.Tools("Window").Menu.Tools.Add "Mosaico"
    '
    '    Set tooMenu = SSActiveToolBars1.ToolBars("Aplicacion").Tools.Add("Exit", ssTypeButton)
    '    tooMenu.Name = "E&xit"
    '    tooMenu.DisplayStyle = ssDisplayTextOnlyAlways
    'SSActiveToolBars1.Tools("Window").Menu.Tools.Add "Exit"

End Sub

Private Sub BuildOptionsTree(strCodAplicacion As String)

    On Error GoTo Fail

    Dim rstOpts As ADODB.Recordset, tooMenu As ActiveToolBars.SSTool, sTit As String
    
    sTit = "Cargar Menu"
    
    'strSql = "SM_MUESTRA_OPCIONES_PERFIL_EMPRESA '" & vper & "', '" & vemp & "'"
    
    strSql = " EXEC SEL_PERMISOS_OPCIONES  " & vbNewLine
    strSql = strSql & "@COD_EMPRESA='" & vemp & "'" & vbNewLine
    strSql = strSql & ",@COD_PERFIL='" & vper & "'" & vbNewLine
    strSql = strSql & ",@COD_APLICACION='" & strCodAplicacion & "'" & vbNewLine

    Set rstOpts = CargarRecordSetDesconectado(strSql, sconnect)

    With rstOpts

        If .RecordCount > 0 Then
            .MoveFirst

            Do Until .EOF

                AddOpcionMenuTree rstOpts
        
                .MoveNext
            Loop

        End If

        .Close
    End With

    Set rstOpts = Nothing

    Exit Sub

Fail:
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit

End Sub

Private Sub AddOpcionMenuTree(rstOpcion As ADODB.Recordset)

    On Error GoTo Fin

    'Dim tooMenu As ActiveToolBars.SSTool
    Dim strIDPadre As String

    With rstOpcion
    
        If !nivel = 1 Then
            strIDPadre = !KEYAPPID
            '  Set treeNode = tvMenuAccesos.Nodes.Add(, , , IIf(iLanguage = 1, !Des_Aplicacion, !Des_Aplicacion_ENG))
            Set treeNode = tvMenuAccesos.Nodes.Add(strIDPadre, tvwChild, !KEYOPCID, !des_opcion)
            treeNode.Tag = !Cod_Opcion
            ' = !Cod_Aplicacion '!KEYAPPID
            '     Set tooMenu = SSActiveToolBars1.Tools(CStr(!Cod_Aplicacion)).Menu.Tools.Add(!Cod_Opcion, ssTypeButton)
        Else
            strIDPadre = !KEYOPCPADREID
            Set treeNode = tvMenuAccesos.Nodes.Add(strIDPadre, tvwChild, !KEYOPCID, !des_opcion)
            treeNode.Tag = !Cod_Opcion
            '    Set tooMenu = SSActiveToolBars1.Tools(CStr(!cod_padre))
            '     tooMenu.Type = ssTypeMenu
            '     Set tooMenu = tooMenu.Menu.Tools.Add(!Cod_Opcion, ssTypeButton)
        End If
    
        '  tooMenu.Name = IIf(iLanguage = 1, !des_opcion, !Des_Opcion_Eng)
        '  tooMenu.DisplayStyle = ssDisplayImageAndText
    End With

    Exit Sub

Fin:

    If Err.Number = 40002 Then
        MsgBox "No se encontró el Menu Superior " & rstOpcion!cod_padre & " para la opcion " & rstOpcion!Cod_Opcion & " (Nivel " & rstOpcion!nivel & ")" & " Verificar que este exista y su Nivel sea menor al de la opcion", vbCritical + vbOKOnly, "Mostrar Opcion en Menu (" & rstOpcion!Cod_Opcion & ")"
    Else
        MsgBox Err.Description, vbCritical + vbOKOnly, "Mostrar Opcion en Menu (" & rstOpcion!Cod_Opcion & ")"
    End If

End Sub

Private Sub BuildMenuBar()

    Dim rstOpts As ADODB.Recordset, tooMenu As ActiveToolBars.SSTool

    Dim count   As Integer
    
    strSql = "SM_MUESTRA_APLICACIONES_PERFIL_EMPRESA '" & vper & "', '" & vemp & "'"
    
    SSActiveToolBars1.ToolBars.Add "Aplicacion", ssMenuBar
    SSActiveToolBars1.ToolBars("Aplicacion").AllowCustomize = True
    
    Set rstOpts = CargarRecordSetDesconectado(strSql, sconnect)
    
    With rstOpts
        count = .RecordCount

        If count > 0 Then
            .MoveFirst
    
            Do Until .EOF
                Set tooMenu = SSActiveToolBars1.ToolBars("Aplicacion").Tools.Add(!Cod_Aplicacion, ssTypeMenu)
                tooMenu.Name = IIf(iLanguage = 1, !Des_Aplicacion, !Des_Aplicacion_ENG)
                'tooMenu.DisplayStyle = ssDisplayTextOnlyInMenus
                .MoveNext
            Loop

        End If

        .Close
    End With

    Set rstOpts = Nothing
    
    Set tooMenu = SSActiveToolBars1.ToolBars("Aplicacion").Tools.Add("Window", ssTypeMenu)
    tooMenu.Name = "&Window"
    SSActiveToolBars1.Tools("Window").Menu.Tools.Add "Cascada"
    SSActiveToolBars1.Tools("Window").Menu.Tools.Add "Mosaico"
    
    Set tooMenu = SSActiveToolBars1.ToolBars("Aplicacion").Tools.Add("Exit", ssTypeButton)
    tooMenu.Name = "E&xit"
    tooMenu.DisplayStyle = ssDisplayTextOnlyAlways
    'SSActiveToolBars1.Tools("Window").Menu.Tools.Add "Exit"
End Sub

Private Sub BuildOptions()

    On Error GoTo Fail

    Dim rstOpts As ADODB.Recordset, tooMenu As ActiveToolBars.SSTool, sTit As String
    
    sTit = "Cargar Menu"
    
    strSql = "SM_MUESTRA_OPCIONES_PERFIL_EMPRESA '" & vper & "', '" & vemp & "'"
    
    Set rstOpts = CargarRecordSetDesconectado(strSql, sconnect)

    With rstOpts

        If .RecordCount > 0 Then
            .MoveFirst

            Do Until .EOF
                '        If !nivel = 1 Then
                '            Set tooMenu = SSActiveToolBars1.Tools(CStr(!Cod_Aplicacion)).Menu.Tools.Add(!Cod_Opcion, ssTypeButton)
                '        Else
                '            Set tooMenu = SSActiveToolBars1.Tools(CStr(!cod_padre))
                '            tooMenu.Type = ssTypeMenu
                '            Set tooMenu = tooMenu.Menu.Tools.Add(!Cod_Opcion, ssTypeButton)
                '        End If
                '
                '        tooMenu.Name = IIf(iLanguage = 1, !des_opcion, !Des_Opcion_Eng)
                '        tooMenu.DisplayStyle = ssDisplayImageAndText
        
                AddOpcionMenu rstOpts
        
                .MoveNext
            Loop

        End If

        .Close
    End With

    Set rstOpts = Nothing

    Exit Sub

Fail:
    '    If Err.Number = 40002 Then
    '        MsgBox "No se encontró el Menu Superior " & rstOpts!cod_padre & _
    '        " para la opcion " & rstOpts!Cod_Opcion & " (Nivel " & rstOpts!nivel & ")" & _
    '        " Verificar que este exista y su Nivel sea menor al de la opcion", _
    '        vbCritical + vbOKOnly, sTit
    '    End If
    MsgBox Err.Description, vbCritical + vbOKOnly, sTit
End Sub

Private Sub AddOpcionMenu(rstOpcion As ADODB.Recordset)

    On Error GoTo Fin

    Dim tooMenu As ActiveToolBars.SSTool

    With rstOpcion

        If !nivel = 1 Then
            Set tooMenu = SSActiveToolBars1.Tools(CStr(!Cod_Aplicacion)).Menu.Tools.Add(!Cod_Opcion, ssTypeButton)
        Else
            Set tooMenu = SSActiveToolBars1.Tools(CStr(!cod_padre))
            tooMenu.Type = ssTypeMenu
            Set tooMenu = tooMenu.Menu.Tools.Add(!Cod_Opcion, ssTypeButton)
        End If
    
        tooMenu.Name = IIf(iLanguage = 1, !des_opcion, !Des_Opcion_Eng)
        tooMenu.DisplayStyle = ssDisplayImageAndText
    End With

    Exit Sub

Fin:

    If Err.Number = 40002 Then
        MsgBox "No se encontró el Menu Superior " & rstOpcion!cod_padre & " para la opcion " & rstOpcion!Cod_Opcion & " (Nivel " & rstOpcion!nivel & ")" & " Verificar que este exista y su Nivel sea menor al de la opcion", vbCritical + vbOKOnly, "Mostrar Opcion en Menu (" & rstOpcion!Cod_Opcion & ")"
    Else
        MsgBox Err.Description, vbCritical + vbOKOnly, "Mostrar Opcion en Menu (" & rstOpcion!Cod_Opcion & ")"
    End If
    
End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
    Debug.Print LCase(Trim(Tool.ID))
    Select Case LCase(Trim(Tool.ID))
        Case LCase("mnuCambioContrasenia")
            EjecutaOpcionMenu "mPWD", Me.perfil, Me.pEmpresa
        Case LCase("mnuAcercaDe")
            frmAcercaDe.Show 1
        Case "exit"
            Unload Me

        Case "cascada"
            Me.Arrange vbCascade

        Case "mosaico"
            Me.Arrange vbTileHorizontal

        Case "mnuregin"

            Dim frmShowTG_PurOrd1 As frmShowTG_PurOrd

            Set frmShowTG_PurOrd1 = New frmShowTG_PurOrd
            Load frmShowTG_PurOrd1
            Set frmShowTG_PurOrd1.oParent = Me
            frmShowTG_PurOrd1.Show
            RefreshWindowList

        Case Else

            If Tool.Type = ssTypeButton Then
                EjecutaOpcionMenu Tool.ID, Me.perfil, Me.pEmpresa
            End If
            
            

    End Select
    
    If Tool.Type = ssTypeStateButton Then

        For Each Form In MDIExtend1.ExForms

            If Tool.ID = Form.Tag Then

                Form.ZOrder 0

                Exit For

            End If

        Next Form

    End If

End Sub

'Public Sub WindowList(sNewDocID As String)
'Dim sWindowName As String, iWindow As Long
'    With SSActiveToolBars1
'        iWindow = .Tools("Window").Menu.Tools.Count + 1
'        sWindowName = iWindow - 2 & ". " & sNewDocID
'
'        .Tools("Window").Menu.Tools.Add sWindowName
'        .Tools(sWindowName).Name = sWindowName
'        .Tools(sWindowName).Type = ssTypeStateButton
'        .Tools(sWindowName).Group = "WindowList"
'        .Tools(sWindowName).GroupAllowAllUp = False
'        .Tools(sWindowName).State = ssChecked
'        .Tools(sWindowName).PictureDown = ImageList1.ListImages("Check").Picture
'        .Tools(sWindowName).Customizable = False ' Prevent tool from showing up in customizer
'
'        If iWindow = 3 Then
'            .Tools("Window").Menu.Tools.Add "separator", , 3
'        End If
'    End With
'
'End Sub

Public Sub RefreshWindowList()

    Dim Ind As Long, sIdTool As String

    With SSActiveToolBars1

        For Ind = 3 To .Tools("Ventana").Menu.Tools.count
            .Tools("Ventana").Menu.Tools.Remove 3
        Next Ind
        
        For Ind = 1 To MDIExtend1.ExForms.count
            sIdTool = Ind & ". " & MDIExtend1.ExForms(Ind).Caption
            .Tools("Ventana").Menu.Tools.Add sIdTool
            .Tools(sIdTool).Name = "&" & sIdTool
            MDIExtend1.ExForms(Ind).Tag = sIdTool
            .Tools(sIdTool).Type = ssTypeStateButton
            .Tools(sIdTool).Group = "WindowList"
            .Tools(sIdTool).GroupAllowAllUp = False
            .Tools(sIdTool).State = ssChecked
            .Tools(sIdTool).PictureDown = ImageList1.ListImages("Check").Picture
            .Tools(sIdTool).Customizable = False ' Prevent tool from showing up in customizer
        Next Ind

        If MDIExtend1.ExForms.count > 0 Then .Tools("Ventana").Menu.Tools.Add "separator", , 3
    End With

End Sub

Public Sub DropWindowList(sDocID As String)

    Dim tooDrop As SSTool

    For Each tooDrop In SSActiveToolBars1.Tools("Ventana").Menu.Tools

        If tooDrop.ID = sDocID Then
            SSActiveToolBars1.Tools("Ventana").Menu.Tools.Remove tooDrop.ID

            Exit For

        End If

    Next tooDrop

End Sub

Private Sub SeteaAnoPeriodoVigente()

    On Error GoTo errx

    Dim sSQl As String

    Dim mRs  As ADODB.Recordset

    sSQl = "CN_OBTIENE_ANO_PERIOD_VIGENTE"
    Set mRs = GetDataSet(cCONNECT, sSQl)

    If Not mRs Is Nothing Then
        Me.Tag = mRs!Ano_Vigente + mRs!Periodo_Vigente
    End If

    mRs.Close
    Set mRs = Nothing

    Exit Sub

errx:
    errores Err.Number
End Sub

Private Sub tvMenuAccesos_NodeClick(ByVal node As MSComctlLib.node)

    If node.Children = 0 Then
        EjecutaOpcionMenu node.Tag, Me.perfil, Me.pEmpresa
    End If

End Sub

Private Sub txtBuscarOpcion_GotFocus()
    '    Call focoControl(txtBuscarPOAsignada, True)
    txtBuscarOpcion.ForeColor = &H80000008

    If txtBuscarOpcion.Text = " Ingrese descripción de la opción a buscar" Then
        txtBuscarOpcion.Text = Empty
    Else
        
    End If

    Call SelectionText(txtBuscarOpcion)
    boolBuscarOpcion = True
   
End Sub

Private Sub txtBuscarOpcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        imgBuscar_Click
    End If
End Sub

Private Sub txtBuscarOpcion_LostFocus()
    ' Call focoControl(txtBuscarPOAsignada, False)
    boolBuscarEstiloCliente = False

    If Trim(txtBuscarOpcion.Text) = "" Then
        txtBuscarOpcion.Text = " Ingrese descripción de la opción a buscar"
        txtBuscarOpcion.ForeColor = &HC0C0C0
    Else
        txtBuscarOpcion.ForeColor = &H80000008
    End If
        
End Sub
