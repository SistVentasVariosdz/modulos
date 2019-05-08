VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{53A95C1B-ED4B-46C8-880A-B248CE857C32}#1.1#0"; "funcbutt.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEstadisticaVentas 
   Caption         =   "Estadistica de Ventas"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opt 
      Caption         =   "Exportacion de Prendas - Clasificacion Arancelaria / Descripción Comercial"
      Height          =   405
      Index           =   4
      Left            =   60
      TabIndex        =   9
      Top             =   1740
      Width           =   4365
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   285
      Left            =   780
      TabIndex        =   5
      Top             =   2220
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      _Version        =   393216
      Format          =   61014017
      CurrentDate     =   39504
   End
   Begin FunctionsButtons.FunctButt FunctButt1 
      Height          =   510
      Left            =   1020
      TabIndex        =   4
      Top             =   2580
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   900
      Custom          =   $"frmEstadisticaVentas.frx":0000
      Orientacion     =   0
      Style           =   0
      Language        =   1
      TypeImageList   =   0
      ControlWidth    =   1155
      ControlHeigth   =   490
      ControlSeparator=   110
   End
   Begin VB.OptionButton opt 
      Caption         =   "Listado de Proveedores de Materiales y Servicios"
      Height          =   405
      Index           =   3
      Left            =   60
      TabIndex        =   3
      Top             =   1320
      Width           =   4365
   End
   Begin VB.OptionButton opt 
      Caption         =   "Listado de materiales de Producción"
      Height          =   405
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   900
      Width           =   4365
   End
   Begin VB.OptionButton opt 
      Caption         =   "Exportación de Prendas  por Tipo de Prenda / Destino"
      Height          =   405
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   4365
   End
   Begin VB.OptionButton opt 
      Caption         =   "Producción total  de Prendas por Tipo de Prenda"
      Height          =   405
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4365
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   285
      Left            =   2820
      TabIndex        =   6
      Top             =   2220
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   61014017
      CurrentDate     =   39504
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   3720
      Top             =   2700
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Left            =   2340
      TabIndex        =   8
      Top             =   2265
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2265
      Width           =   465
   End
End
Attribute VB_Name = "frmEstadisticaVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Dia1 As Date, Dia365 As Date
Dia1 = CDate("01/01/" & Year(Date))
Dia365 = DateAdd("yyyy", 1, Dia1)
Dia365 = DateAdd("d", -1, Dia365)
dtpDesde.Value = Dia1
dtpHasta.Value = Dia365
End Sub

Private Sub FunctButt1_ActionClick(ByVal Index As Integer, ByVal ActionType As Integer, ByVal ActionName As String)
Dim strSQL As String
Select Case ActionName
Case "IMPRIMIR"
    Call Imprimir
Case "CANCELAR"
    Unload Me
End Select

End Sub
Sub Imprimir()
On Error GoTo hand
Dim oo As Object
Dim Ruta As String, iResp As Integer
    
    iResp = MsgBox("Imprimir usando Microsoft Excel?", vbQuestion + vbYesNo, "Imprimir")
    
    If opt(0).Value Then
        strSQL = "EXEC CF_SM_PRODUCCION_PRENDAS '" & dtpDesde.Value & "' ,'" & dtpHasta.Value & "'"
        Ruta = vRuta & "\rptProduccionPrendas." & IIf((iResp = vbYes), "XLT", "OTS")
    End If
    If opt(1).Value Then
        strSQL = "EXEC Ventas_Muestra_Documento_Exportacion_ENCUESTA '" & dtpDesde.Value & "' ,'" & dtpHasta.Value & "','D','','','0','','','1','','N'"
        Ruta = vRuta & "\rptExpTelaXTipPrend." & IIf((iResp = vbYes), "XLT", "OTS")
    End If
    If opt(2).Value Then
        strSQL = "EXEC CF_SM_LISTADO_MATERIALES_PRODUCCION '" & dtpDesde.Value & "' ,'" & dtpHasta.Value & "'"
        Ruta = vRuta & "\rptListadoMatProduccion." & IIf((iResp = vbYes), "XLT", "OTS")
    End If
    If opt(3).Value Then
        strSQL = "EXEC TX_SM_LISTADO_PROVEEDORES  '" & dtpDesde.Value & "' ,'" & dtpHasta.Value & "'"
        Ruta = vRuta & "\rptProveedores." & IIf((iResp = vbYes), "XLT", "OTS")
    End If
    If opt(4).Value Then
        strSQL = "EXEC Ventas_Muestra_Documento_Exportacion_ENCUESTA_CLASIFICACION_ARANCELARIA '" & dtpDesde.Value & "' ,'" & dtpHasta.Value & "','D','','','0','','','1','','N'"
        Ruta = vRuta & "\rptExpTelaEncuestaArancelaria." & IIf((iResp = vbYes), "XLT", "OTS")
    End If
    'Dim CodFabrica As String
    'CodFabrica = DevuelveCampo("select cod_razsocial from seg_empresas where cod_empresa='" & vemp & "' ", cSEGURIDAD)
    Dim strRango As String
    strRango = "Desde : " & dtpDesde.Value & "  Hasta : " & dtpHasta.Value
    
    If iResp = vbYes Then
        Set oo = CreateObject("excel.application")
        oo.Workbooks.Open Ruta
        oo.Visible = True
        oo.DisplayAlerts = False
        
        oo.Run "reporte", strSQL, cCONNECT, strRango, cSEGURIDAD, vemp
    Else
        Set oo = CreateObject("ooBusiness.Calc")
        oo.OfficeTemplateSheet = Ruta
        oo.OfficeDocumentSheet = Replace(Ruta, ".OTS", Format(Now, "YYYYMMDDHHMMSSMM") & ".ods")
        oo.MacroLibraryName = "Library1"
        oo.MacroModuleName = "Module1"
        oo.MacroName = "Reporte"
        
        oo.Run strSQL, cCONNECT, strRango, cSEGURIDAD, vemp
    End If
    Set oo = Nothing
Exit Sub
hand:
    ErrorHandler err, "GeneraReportes"
    Set oo = Nothing
End Sub

