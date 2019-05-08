VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmOrdCompItemAddS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requerimientos por Comprar"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   480
      Left            =   3135
      TabIndex        =   14
      Top             =   4500
      Width           =   1380
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detalle de Requerimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   45
      TabIndex        =   9
      Top             =   1080
      Width           =   10035
      Begin SSDataWidgets_B.SSDBGrid DGridLista 
         Height          =   2925
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Visible         =   0   'False
         Width           =   9825
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   12
         AllowColumnShrinking=   0   'False
         SelectTypeRow   =   1
         BackColorOdd    =   13434879
         RowHeight       =   423
         ExtraHeight     =   106
         Columns.Count   =   12
         Columns(0).Width=   1429
         Columns(0).Caption=   "Flag O/P"
         Columns(0).Name =   "chec"
         Columns(0).CaptionAlignment=   2
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   2
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "Cod_Fabrica"
         Columns(1).Name =   "Cod_Fabrica"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   4101
         Columns(2).Caption=   "Fábrica"
         Columns(2).Name =   "Fabrica"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         Columns(2).Style=   4
         Columns(3).Width=   2778
         Columns(3).Caption=   "O/P"
         Columns(3).Name =   "Cod_OrdPro"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(3).Locked=   -1  'True
         Columns(3).Style=   4
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
         Columns(4).Caption=   "Cod_HilTel"
         Columns(4).Name =   "Cod_HilTel"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "Des_hiltel"
         Columns(5).Name =   "Des_hiltel"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3096
         Columns(6).Caption=   "Hilo"
         Columns(6).Name =   "Hilo"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(6).Locked=   -1  'True
         Columns(6).Style=   4
         Columns(7).Width=   2672
         Columns(7).Caption=   "Cant. por Comprar"
         Columns(7).Name =   "CANTXCOMPRAR"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(7).Locked=   -1  'True
         Columns(8).Width=   3200
         Columns(8).Visible=   0   'False
         Columns(8).Caption=   "cod_destino"
         Columns(8).Name =   "cod_destino"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   3200
         Columns(9).Visible=   0   'False
         Columns(9).Caption=   "cod_estcli"
         Columns(9).Name =   "cod_estcli"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(10).Width=   3200
         Columns(10).Visible=   0   'False
         Columns(10).Caption=   "CANTIDAD"
         Columns(10).Name=   "CANTIDAD"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(11).Width=   3200
         Columns(11).Caption=   "Cod_Prov"
         Columns(11).Name=   "Cod_Prov"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         _ExtentX        =   17330
         _ExtentY        =   5159
         _StockProps     =   79
         Caption         =   "Resultados de la Busqueda"
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B.SSDBGrid DGridLista 
         Height          =   2925
         Index           =   2
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Visible         =   0   'False
         Width           =   9825
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   19
         AllowColumnShrinking=   0   'False
         SelectTypeRow   =   1
         BackColorOdd    =   13434879
         RowHeight       =   423
         ExtraHeight     =   106
         Columns.Count   =   19
         Columns(0).Width=   1429
         Columns(0).Caption=   "Flag O/P"
         Columns(0).Name =   "chec"
         Columns(0).CaptionAlignment=   2
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   2
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "Cod_Fabrica"
         Columns(1).Name =   "Cod_Fabrica"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3731
         Columns(2).Caption=   "Fábrica"
         Columns(2).Name =   "Fabrica"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         Columns(2).Style=   4
         Columns(3).Width=   2381
         Columns(3).Caption=   "O/P"
         Columns(3).Name =   "Cod_OrdPro"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(3).Locked=   -1  'True
         Columns(3).Style=   4
         Columns(4).Width=   3334
         Columns(4).Caption=   "Presentación"
         Columns(4).Name =   "Des_Present"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(4).Locked=   -1  'True
         Columns(4).Style=   4
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "Cod_Tela"
         Columns(5).Name =   "Cod_Tela"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3200
         Columns(6).Visible=   0   'False
         Columns(6).Caption=   "Des_Tela"
         Columns(6).Name =   "Des_Tela"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   3334
         Columns(7).Caption=   "Tela"
         Columns(7).Name =   "Tela"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(7).Locked=   -1  'True
         Columns(7).Style=   4
         Columns(8).Width=   3200
         Columns(8).Visible=   0   'False
         Columns(8).Caption=   "Cod_Comb"
         Columns(8).Name =   "Cod_Comb"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   3200
         Columns(9).Visible=   0   'False
         Columns(9).Caption=   "Des_Comb"
         Columns(9).Name =   "Des_Comb"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(10).Width=   2990
         Columns(10).Caption=   "Combinación"
         Columns(10).Name=   "Combinacion"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(10).Locked=   -1  'True
         Columns(10).Style=   4
         Columns(11).Width=   1535
         Columns(11).Caption=   "Talla"
         Columns(11).Name=   "Cod_Talla"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         Columns(11).Locked=   -1  'True
         Columns(11).Style=   4
         Columns(12).Width=   2699
         Columns(12).Caption=   "Cant. por Comprar"
         Columns(12).Name=   "CANTXCOMPRAR"
         Columns(12).DataField=   "Column 12"
         Columns(12).DataType=   8
         Columns(12).FieldLen=   256
         Columns(12).Locked=   -1  'True
         Columns(13).Width=   3200
         Columns(13).Visible=   0   'False
         Columns(13).Caption=   "Cod_Present"
         Columns(13).Name=   "Cod_Present"
         Columns(13).DataField=   "Column 13"
         Columns(13).DataType=   8
         Columns(13).FieldLen=   256
         Columns(14).Width=   3200
         Columns(14).Visible=   0   'False
         Columns(14).Caption=   "Cod_CompEst"
         Columns(14).Name=   "Cod_CompEst"
         Columns(14).DataField=   "Column 14"
         Columns(14).DataType=   8
         Columns(14).FieldLen=   256
         Columns(15).Width=   3200
         Columns(15).Visible=   0   'False
         Columns(15).Caption=   "cod_destino"
         Columns(15).Name=   "cod_destino"
         Columns(15).DataField=   "Column 15"
         Columns(15).DataType=   8
         Columns(15).FieldLen=   256
         Columns(16).Width=   3200
         Columns(16).Visible=   0   'False
         Columns(16).Caption=   "cod_estcli"
         Columns(16).Name=   "cod_estcli"
         Columns(16).DataField=   "Column 16"
         Columns(16).DataType=   8
         Columns(16).FieldLen=   256
         Columns(17).Width=   3200
         Columns(17).Visible=   0   'False
         Columns(17).Caption=   "CANTIDAD"
         Columns(17).Name=   "CANTIDAD"
         Columns(17).DataField=   "Column 17"
         Columns(17).DataType=   8
         Columns(17).FieldLen=   256
         Columns(18).Width=   1958
         Columns(18).Caption=   "Medida"
         Columns(18).Name=   "MEDIDA"
         Columns(18).DataField=   "Column 18"
         Columns(18).DataType=   8
         Columns(18).FieldLen=   256
         _ExtentX        =   17330
         _ExtentY        =   5159
         _StockProps     =   79
         Caption         =   "Resultados de la Busqueda"
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B.SSDBGrid DGridLista 
         Height          =   2925
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Visible         =   0   'False
         Width           =   9825
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   14
         AllowColumnShrinking=   0   'False
         SelectTypeRow   =   1
         BackColorOdd    =   13434879
         RowHeight       =   423
         ExtraHeight     =   106
         Columns.Count   =   14
         Columns(0).Width=   1429
         Columns(0).Caption=   "Flag O/P"
         Columns(0).Name =   "chec"
         Columns(0).CaptionAlignment=   2
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   2
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "Cod_Fabrica"
         Columns(1).Name =   "Cod_Fabrica"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   4101
         Columns(2).Caption=   "Fábrica"
         Columns(2).Name =   "Fabrica"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         Columns(2).Style=   4
         Columns(3).Width=   2778
         Columns(3).Caption=   "O/P"
         Columns(3).Name =   "Cod_OrdPro"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(3).Locked=   -1  'True
         Columns(3).Style=   4
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
         Columns(4).Caption=   "Cod_HilTel"
         Columns(4).Name =   "Cod_HilTel"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "Des_hiltel"
         Columns(5).Name =   "Des_hiltel"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3096
         Columns(6).Caption=   "Hilo"
         Columns(6).Name =   "Hilo"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(6).Locked=   -1  'True
         Columns(6).Style=   4
         Columns(7).Width=   3200
         Columns(7).Visible=   0   'False
         Columns(7).Caption=   "Cod_Color"
         Columns(7).Name =   "Cod_Color"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   3200
         Columns(8).Visible=   0   'False
         Columns(8).Caption=   "Des_Color"
         Columns(8).Name =   "Des_Color"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   2990
         Columns(9).Caption=   "Color"
         Columns(9).Name =   "Color"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(9).Locked=   -1  'True
         Columns(9).Style=   4
         Columns(10).Width=   2672
         Columns(10).Caption=   "Cant. por Comprar"
         Columns(10).Name=   "CANTXCOMPRAR"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(10).Locked=   -1  'True
         Columns(11).Width=   3200
         Columns(11).Visible=   0   'False
         Columns(11).Caption=   "cod_destino"
         Columns(11).Name=   "cod_destino"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         Columns(12).Width=   3200
         Columns(12).Visible=   0   'False
         Columns(12).Caption=   "cod_estcli"
         Columns(12).Name=   "cod_estcli"
         Columns(12).DataField=   "Column 12"
         Columns(12).DataType=   8
         Columns(12).FieldLen=   256
         Columns(13).Width=   3200
         Columns(13).Visible=   0   'False
         Columns(13).Caption=   "CANTIDAD"
         Columns(13).Name=   "CANTIDAD"
         Columns(13).DataField=   "Column 13"
         Columns(13).DataType=   8
         Columns(13).FieldLen=   256
         _ExtentX        =   17330
         _ExtentY        =   5159
         _StockProps     =   79
         Caption         =   "Resultados de la Busqueda"
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B.SSDBGrid DGridLista 
         Height          =   2925
         Index           =   3
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Visible         =   0   'False
         Width           =   9825
         _Version        =   196617
         DataMode        =   2
         Col.Count       =   23
         AllowColumnShrinking=   0   'False
         SelectTypeRow   =   1
         BackColorOdd    =   13434879
         RowHeight       =   423
         ExtraHeight     =   106
         Columns.Count   =   23
         Columns(0).Width=   1429
         Columns(0).Caption=   "Flag O/P"
         Columns(0).Name =   "chec"
         Columns(0).CaptionAlignment=   2
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Style=   2
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "Cod_Fabrica"
         Columns(1).Name =   "Cod_Fabrica"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3731
         Columns(2).Caption=   "Fábrica"
         Columns(2).Name =   "Fabrica"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         Columns(2).Style=   4
         Columns(3).Width=   2381
         Columns(3).Caption=   "O/P"
         Columns(3).Name =   "Cod_OrdPro"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(3).Locked=   -1  'True
         Columns(3).Style=   4
         Columns(4).Width=   3334
         Columns(4).Caption=   "Presentación"
         Columns(4).Name =   "Des_Present"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(4).Locked=   -1  'True
         Columns(4).Style=   4
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "Cod_Tela"
         Columns(5).Name =   "Cod_Tela"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3200
         Columns(6).Visible=   0   'False
         Columns(6).Caption=   "Des_Tela"
         Columns(6).Name =   "Des_Tela"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   3334
         Columns(7).Caption=   "Tela"
         Columns(7).Name =   "Tela"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(7).Locked=   -1  'True
         Columns(7).Style=   4
         Columns(8).Width=   3200
         Columns(8).Visible=   0   'False
         Columns(8).Caption=   "Cod_Comb"
         Columns(8).Name =   "Cod_Comb"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   3200
         Columns(9).Visible=   0   'False
         Columns(9).Caption=   "Des_Comb"
         Columns(9).Name =   "Des_Comb"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(10).Width=   2990
         Columns(10).Caption=   "Combinación"
         Columns(10).Name=   "Combinacion"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(10).Locked=   -1  'True
         Columns(10).Style=   4
         Columns(11).Width=   3200
         Columns(11).Visible=   0   'False
         Columns(11).Caption=   "Cod_Color"
         Columns(11).Name=   "Cod_Color"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         Columns(12).Width=   3200
         Columns(12).Visible=   0   'False
         Columns(12).Caption=   "Des_Color"
         Columns(12).Name=   "Des_Color"
         Columns(12).DataField=   "Column 12"
         Columns(12).DataType=   8
         Columns(12).FieldLen=   256
         Columns(13).Width=   3200
         Columns(13).Caption=   "Color"
         Columns(13).Name=   "Color"
         Columns(13).DataField=   "Column 13"
         Columns(13).DataType=   8
         Columns(13).FieldLen=   256
         Columns(13).Locked=   -1  'True
         Columns(13).Style=   4
         Columns(14).Width=   3200
         Columns(14).Caption=   "Receta"
         Columns(14).Name=   "Cod_Receta"
         Columns(14).DataField=   "Column 14"
         Columns(14).DataType=   8
         Columns(14).FieldLen=   256
         Columns(14).Locked=   -1  'True
         Columns(14).Style=   4
         Columns(15).Width=   1535
         Columns(15).Caption=   "Talla"
         Columns(15).Name=   "Cod_Talla"
         Columns(15).DataField=   "Column 15"
         Columns(15).DataType=   8
         Columns(15).FieldLen=   256
         Columns(15).Locked=   -1  'True
         Columns(15).Style=   4
         Columns(16).Width=   2699
         Columns(16).Caption=   "Cant. por Comprar"
         Columns(16).Name=   "CANTXCOMPRAR"
         Columns(16).DataField=   "Column 16"
         Columns(16).DataType=   8
         Columns(16).FieldLen=   256
         Columns(16).Locked=   -1  'True
         Columns(17).Width=   3200
         Columns(17).Visible=   0   'False
         Columns(17).Caption=   "Cod_Present"
         Columns(17).Name=   "Cod_Present"
         Columns(17).DataField=   "Column 17"
         Columns(17).DataType=   8
         Columns(17).FieldLen=   256
         Columns(18).Width=   3200
         Columns(18).Visible=   0   'False
         Columns(18).Caption=   "Cod_CompEst"
         Columns(18).Name=   "Cod_CompEst"
         Columns(18).DataField=   "Column 18"
         Columns(18).DataType=   8
         Columns(18).FieldLen=   256
         Columns(19).Width=   3200
         Columns(19).Visible=   0   'False
         Columns(19).Caption=   "cod_destino"
         Columns(19).Name=   "cod_destino"
         Columns(19).DataField=   "Column 19"
         Columns(19).DataType=   8
         Columns(19).FieldLen=   256
         Columns(20).Width=   3200
         Columns(20).Visible=   0   'False
         Columns(20).Caption=   "cod_estcli"
         Columns(20).Name=   "cod_estcli"
         Columns(20).DataField=   "Column 20"
         Columns(20).DataType=   8
         Columns(20).FieldLen=   256
         Columns(21).Width=   3200
         Columns(21).Visible=   0   'False
         Columns(21).Caption=   "CANTIDAD"
         Columns(21).Name=   "CANTIDAD"
         Columns(21).DataField=   "Column 21"
         Columns(21).DataType=   8
         Columns(21).FieldLen=   256
         Columns(22).Width=   1958
         Columns(22).Caption=   "Medida"
         Columns(22).Name=   "MEDIDA"
         Columns(22).DataField=   "Column 22"
         Columns(22).DataType=   8
         Columns(22).FieldLen=   256
         _ExtentX        =   17330
         _ExtentY        =   5159
         _StockProps     =   79
         Caption         =   "Resultados de la Busqueda"
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opcion de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   50
      TabIndex        =   1
      Top             =   15
      Width           =   10035
      Begin VB.CommandButton cmdBuscaColor 
         Caption         =   "..."
         Height          =   315
         Left            =   6730
         TabIndex        =   23
         Top             =   600
         Width           =   350
      End
      Begin VB.TextBox TxtColor 
         Height          =   315
         Left            =   4695
         TabIndex        =   22
         Top             =   600
         Width           =   2010
      End
      Begin VB.OptionButton OpColor 
         Height          =   255
         Left            =   3720
         TabIndex        =   20
         Top             =   650
         Width           =   255
      End
      Begin VB.OptionButton OpFam 
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   300
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscaFam 
         Caption         =   "..."
         Height          =   330
         Left            =   6735
         TabIndex        =   16
         Top             =   240
         Width           =   330
      End
      Begin VB.TextBox txtDes_Grupo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1170
         TabIndex        =   15
         Top             =   240
         Width           =   2355
      End
      Begin VB.CommandButton cmdBuscaOP 
         Caption         =   "..."
         Height          =   330
         Left            =   3195
         TabIndex        =   5
         Tag             =   "..."
         Top             =   570
         Width           =   330
      End
      Begin VB.TextBox TxtOp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1170
         MaxLength       =   20
         TabIndex        =   4
         Top             =   570
         Width           =   2010
      End
      Begin VB.TextBox TxtFamilia 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4695
         MaxLength       =   20
         TabIndex        =   3
         Top             =   240
         Width           =   2010
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   405
         Left            =   7200
         TabIndex        =   2
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label4 
         Caption         =   "Color"
         Height          =   180
         Left            =   4080
         TabIndex        =   21
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Total:"
         Height          =   255
         Left            =   8760
         TabIndex        =   18
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo"
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "O/P"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Familias"
         Height          =   195
         Index           =   1
         Left            =   4080
         TabIndex        =   6
         Top             =   330
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Cancelar"
      Height          =   480
      Left            =   5115
      TabIndex        =   0
      Top             =   4500
      Width           =   1470
   End
End
Attribute VB_Name = "frmOrdCompItemAddS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Strsql As String
Dim Rs_Lista As New ADODB.Recordset
Dim CadConn  As New ADODB.Connection

'Variables para la ejecucion del store
Public varCod_GrupoTex As String
Public varCodigo As String

Dim sTipo As String
Public Codigo, Descripcion As String
Dim opcionProv As Integer
'Definicion de variables que seran pasadas por nuestro master
Public varSer_OrdComp, varCod_OrdComp, varSec_OrdComp As String

Public varTip_Presentacion, varCod_ClaOrdComp, varCod_Proveedor As String
Public varCod_Descuento As String
Public varCod_TipRequ As Integer
Public varPorc_IGV As Double
'Variables para la ejecucion del super mega store de generacion de requerimientos
Public varAccion As String

Dim FlagU As Boolean
Dim FlagN As Boolean


Public Sub MUESTRA_GRID(opcion As Integer)
Dim Rs As ADODB.Recordset
    Select Case opcion
        Case 2, 6:
                'Strsql = "exec UP_SEL_REQUEXCOMPRARTEX 4,'" & varCod_GrupoTex & "','" & TxtOp.Text & "','" & Trim(TxtFamilia.Text) & "'"
                'opcionProv = 4
                DGridLista(3).Visible = True
        Case 3:
                'Strsql = "exec UP_SEL_REQUEXCOMPRARTEX 3,'" & varCod_GrupoTex & "','" & TxtOp.Text & "','" & Trim(TxtFamilia.Text) & "'"
                'opcionProv = 3
                DGridLista(2).Visible = True
        Case 4:
                'Strsql = "exec UP_SEL_REQUEXCOMPRARTEX 2,'" & varCod_GrupoTex & "','" & TxtOp.Text & "','" & Trim(TxtFamilia.Text) & "'"
                'opcionProv = 2
                DGridLista(1).Visible = True
        Case 5:
                'Strsql = "exec UP_SEL_REQUEXCOMPRARTEX 1,'" & varCod_GrupoTex & "','" & TxtOp.Text & "','" & Trim(TxtFamilia.Text) & "'"
                'opcionProv = 1
                DGridLista(0).Visible = True
    End Select
    
    Set Rs = New ADODB.Recordset
    Strsql = "SELECT flg_requerimiento,tip_item,tip_presentacion FROM LG_CLAORDCOMP WHERE COD_CLAORDCOMP='" & varCod_ClaOrdComp & "'"
    Rs.Open Strsql, cConnect, adOpenStatic, adLockReadOnly
    If Rs.RecordCount Then
        If Rs("flg_requerimiento") = "S" And Rs("tip_item") = "T" And Rs("tip_presentacion") = "T" Then
            OpFam_Click
        Else
            OpFam_Click
            OpColor.Enabled = False
        End If
    End If
    
    Set Rs = Nothing
    
End Sub

Public Sub CARGA_LISTA(opcion As Integer)
    On Error GoTo Cargar_DatosErr
    Dim Rs_Prov As New ADODB.Recordset
    'Dim opcionProv As Integer
    Dim i As Integer
    
    FlagN = True
    FlagU = False
    
    Strsql = "exec UP_SEL_REQUEXCOMPRARTEX '" & varSer_OrdComp & "','" & varCod_OrdComp & "','" & varCod_GrupoTex & "','" & TxtOp.Text & "','" & Trim(TxtFamilia.Text) & "','" & Trim(TxtColor.Text) & "'"
    
    Select Case opcion
        Case 2, 6:
                'Strsql = "exec UP_SEL_REQUEXCOMPRARTEX '" & varSer_OrdComp & "','" & varCod_OrdComp & "','" & varCod_GrupoTex & "','" & TxtOp.Text & "','" & Trim(TxtFamilia.Text) & "','" & Trim(TxtColor.Text) & "'"
                opcionProv = 4
        Case 3:
                'Strsql = "exec UP_SEL_REQUEXCOMPRARTEX 3,'" & varCod_GrupoTex & "','" & TxtOp.Text & "','" & Trim(TxtFamilia.Text) & "','" & Trim(TxtColor.Text) & "'"
                opcionProv = 3
        Case 4:
                'Strsql = "exec UP_SEL_REQUEXCOMPRARTEX 2,'" & varCod_GrupoTex & "','" & TxtOp.Text & "','" & Trim(TxtFamilia.Text) & "','" & Trim(TxtColor.Text) & "'"
                opcionProv = 2
        Case 5:
                'Strsql = "exec UP_SEL_REQUEXCOMPRARTEX 1,'" & varCod_GrupoTex & "','" & TxtOp.Text & "','" & Trim(TxtFamilia.Text) & "','" & Trim(TxtColor.Text) & "'"
                opcionProv = 1
    End Select
    
    'Strsql = "EXEC UP_SEL_REQUEXCOMPRARTEX " & CStr(Opcion - 1) & ",'" & Right(varCod_GrupoTex.Text, 8) & "','" & TxtOp.Text & "',''"
    
    Set Rs_Lista = Nothing
    Rs_Lista.ActiveConnection = cConnect
    Rs_Lista.CursorType = adOpenStatic
    Rs_Lista.CursorLocation = adUseClient
    Rs_Lista.LockType = adLockReadOnly
    Rs_Lista.Open Strsql
    
    Set Rs_Prov = Rs_Lista.Clone
    
    'Esto es para asignar la data al grid
        
    If Rs_Lista.RecordCount > 0 Then
        'NADA
    Else
        MsgBox "No se encontraron registros ", vbInformation, "Ordenes de Compra"
    End If
    

        Select Case opcionProv
            Case 1
                    Me.DGridLista(0).Redraw = False
                    SSDBGridSetGrid Me.DGridLista(0)
                    ADODBToSSDBGridOC Rs_Prov, DGridLista(0)
                    'RBSToSSDBGridOC Rs_Prov, DGridLista(0)
                    DGridLista(0).ActiveRowStyleSet = "RowActive"
                    DGridLista(0).SelectTypeRow = ssSelectionTypeMultiSelectRange
                    DGridLista(0).Visible = True
                    DGridLista(0).Redraw = False
                    For i = 0 To DGridLista(0).Rows
                        DGridLista(0).Bookmark = i
'                        If i >= 7 Then DGridLista(0).Scroll 0, 1
                        DGridLista(0).Columns(0).Value = 1
                        If i = DGridLista(0).Rows - 1 Then
                            FlagU = True
'                            DGridLista(0).Row = 1
                            DGridLista(0).Redraw = True
                            Exit For
                        End If
                    Next
            Case 2
                    Me.DGridLista(1).Redraw = False
                    SSDBGridSetGrid Me.DGridLista(1)
                    ADODBToSSDBGridOC Rs_Prov, DGridLista(1)
                    DGridLista(1).ActiveRowStyleSet = "RowActive"
                    DGridLista(1).SelectTypeRow = ssSelectionTypeMultiSelectRange
                    DGridLista(1).Visible = True
                    
                    DGridLista(1).Redraw = False
                    For i = 0 To DGridLista(1).Rows
                        DGridLista(1).Bookmark = i
'                        If i >= 7 Then DGridLista(1).Scroll 0, 1
                        DGridLista(1).Columns(0).Value = 1
                        If i = DGridLista(1).Rows - 1 Then
                            FlagU = True
'                            DGridLista(1).Row = 1
                            DGridLista(1).Redraw = True
                            Exit For
                        End If
                    Next
            Case 3
                    Me.DGridLista(2).Redraw = False
                    SSDBGridSetGrid Me.DGridLista(2)
                    ADODBToSSDBGridOC Rs_Prov, DGridLista(2)
                    DGridLista(2).ActiveRowStyleSet = "RowActive"
                    DGridLista(2).SelectTypeRow = ssSelectionTypeMultiSelectRange
                    DGridLista(2).Visible = True
                    
                    DGridLista(2).Redraw = False
                    For i = 0 To DGridLista(2).Rows
                        DGridLista(2).Bookmark = i
'                        If i >= 7 Then DGridLista(2).Scroll 0, 1
                        DGridLista(2).Columns(0).Value = 1
                        If i = DGridLista(2).Rows - 1 Then
                            DGridLista(2).Redraw = True
'                            DGridLista(2).Row = 1
                            FlagU = True
                            Exit For
                        End If
                    Next
            Case 4
                    Me.DGridLista(3).Redraw = False
                    SSDBGridSetGrid Me.DGridLista(3)
                    ADODBToSSDBGridOC Rs_Prov, DGridLista(3)
                    DGridLista(3).ActiveRowStyleSet = "RowActive"
                    DGridLista(3).SelectTypeRow = ssSelectionTypeMultiSelectRange
                    DGridLista(3).Visible = True
                    
                    DGridLista(3).Redraw = False
                    For i = 0 To DGridLista(3).Rows
                        DGridLista(3).Bookmark = i
'                        If i >= 7 Then DGridLista(3).Scroll 0, 1
                        DGridLista(3).Columns(0).Value = 1
                        If i = DGridLista(3).Rows - 1 Then
                            FlagU = True
'                            DGridLista(3).Row = 1
                            DGridLista(3).Redraw = True
                            Exit For
                        End If
                    Next
                    
        End Select
       
    FlagU = True
       
    Rs_Prov.Close
    Set Rs_Prov = Nothing
    
    Exit Sub
Cargar_DatosErr:
    Set Rs_Lista = Nothing
    ErrorHandler Err, "Cargar_Datos"
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo ErrorAceptar:
    

    If DGridLista(opcionProv - 1).Rows < 1 Then
        MsgBox "No existen registros para acceder a esta opcion", vbInformation, "Ordenes de Pedido"
        Exit Sub
    End If

    'Llamando a este form obtendremos el tipo de insercion
    Load frmOpcionReq
    Set frmOpcionReq.frmmaster = Me
    frmOpcionReq.Show 1
    
    If varAccion = "" Then
        Exit Sub
    End If
    
    If varAccion <> "I" Then
        Strsql = "SELECT ISNULL(MAX(Sec_OrdComp),0) FROM lg_ordcompitem WHERE Ser_OrdComp='" & Me.varSer_OrdComp & "' AND Cod_OrdComp='" & Me.varCod_OrdComp & "'"
        varSec_OrdComp = DevuelveCampo(Strsql, cConnect)
    End If
    
    Set CadConn = Nothing
    CadConn.Open cConnect
    Dim j As Integer
    FlagU = False
    'DGridLista(opcionProv - 1).Row = 0
    DGridLista(opcionProv - 1).Bookmark = 0
    For j = 0 To DGridLista(opcionProv - 1).Rows - 1
        'DGridLista(opcionProv - 1).Row = j
        'Grilla.Bookmark = j
        If Abs(IIf(Trim(DGridLista(opcionProv - 1).Columns(0).Value) = "", 0, DGridLista(opcionProv - 1).Columns(0).Value)) = 1 Then
        'If DGridLista(opcionProv - 1).Columns(0).Value = 1 Then
            Select Case varCod_TipRequ
                Case 2, 6:
                        Strsql = "exec UP_ACTUALIZA_REQ_OC  '" & _
                        varSer_OrdComp & "','" & _
                        varCod_OrdComp & "','" & _
                        varSec_OrdComp & "','" & _
                        varAccion & "','" & _
                        DGridLista(3).Columns("cod_fabrica").Text & "','" & _
                        DGridLista(3).Columns("Cod_OrdPro").Text & "','" & _
                        DGridLista(3).Columns("Cod_Present").Text & "','" & _
                        DGridLista(3).Columns("Cod_CompEst").Text & "','" & _
                        DGridLista(3).Columns("Cod_Tela").Text & "','" & _
                        DGridLista(3).Columns("Cod_Comb").Text & "','" & _
                        DGridLista(3).Columns("Cod_Color").Text & "','" & _
                        DGridLista(3).Columns("Cod_Talla").Text & "','" & _
                        DGridLista(3).Columns("cod_destino").Text & "','" & _
                        DGridLista(3).Columns("cod_estcli").Text & "'," & _
                        DGridLista(3).Columns("CANTXCOMPRAR").Text & ",''"
                        
                        'If j >= 7 Then DGridLista(3).Scroll 0, 1

                        opcionProv = 4
                Case 3:
                        
                        Strsql = "exec UP_ACTUALIZA_REQ_OC  '" & _
                        varSer_OrdComp & "','" & _
                        varCod_OrdComp & "','" & _
                        varSec_OrdComp & "','" & _
                        varAccion & "','" & _
                        DGridLista(2).Columns("cod_fabrica").Text & "','" & _
                        DGridLista(2).Columns("Cod_OrdPro").Text & "','" & _
                        DGridLista(2).Columns("Cod_Present").Text & "','" & _
                        DGridLista(2).Columns("Cod_CompEst").Text & "','" & _
                        DGridLista(2).Columns("Cod_Tela").Text & "','" & _
                        DGridLista(2).Columns("Cod_Comb").Text & "','" & _
                        "" & "','" & _
                        DGridLista(2).Columns("Cod_Talla").Text & "','" & _
                        DGridLista(2).Columns("cod_destino").Text & "','" & _
                        DGridLista(2).Columns("cod_estcli").Text & "'," & _
                        DGridLista(2).Columns("CANTXCOMPRAR").Text & ",''"
                        
                        'If j >= 7 Then DGridLista(2).Scroll 0, 1
                        
                        'Strsql = "exec UP_SEL_REQUEXCOMPRARTEX 3,'" & Right(varCod_GrupoTex.Text, 8) & "','" & TxtOp.Text & "',''"
                        opcionProv = 3
                Case 4:
                        Strsql = "exec UP_ACTUALIZA_REQ_OC  '" & _
                        varSer_OrdComp & "','" & _
                        varCod_OrdComp & "','" & _
                        varSec_OrdComp & "','" & _
                        varAccion & "','" & _
                        DGridLista(1).Columns("cod_fabrica").Text & "','" & _
                        DGridLista(1).Columns("Cod_OrdPro").Text & "','" & _
                        "" & "','" & _
                        "" & "','" & _
                        DGridLista(1).Columns("Cod_HilTel").Text & "','" & _
                        "" & "','" & _
                        DGridLista(1).Columns("Cod_Color").Text & "','" & _
                        "" & "','" & _
                        DGridLista(1).Columns("cod_destino").Text & "','" & _
                        DGridLista(1).Columns("cod_estcli").Text & "'," & _
                        DGridLista(1).Columns("CANTXCOMPRAR").Text & ",''"
                        
                        'If j >= 7 Then DGridLista(1).Scroll 0, 1
                        
                        'Strsql = "exec UP_SEL_REQUEXCOMPRARTEX 2,'" & Right(varCod_GrupoTex.Text, 8) & "','" & TxtOp.Text & "',''"
                        opcionProv = 2
                Case 5:
                
                        Strsql = "exec UP_ACTUALIZA_REQ_OC  '" & _
                        varSer_OrdComp & "','" & _
                        varCod_OrdComp & "','" & _
                        varSec_OrdComp & "','" & _
                        varAccion & "','" & _
                        DGridLista(0).Columns("cod_fabrica").Text & "','" & _
                        DGridLista(0).Columns("Cod_OrdPro").Text & "','" & _
                        "" & "','" & _
                        "" & "','" & _
                        DGridLista(0).Columns("Cod_HilTel").Text & "','" & _
                        "" & "','" & _
                        "" & "','" & _
                        "" & "','" & _
                        DGridLista(0).Columns("cod_destino").Text & "','" & _
                        DGridLista(0).Columns("cod_estcli").Text & "'," & _
                        DGridLista(0).Columns("CANTXCOMPRAR").Text & ",''"
                        
                        'If j >= 7 Then DGridLista(0).Scroll 0, 1
                        
                        'Strsql = "exec UP_SEL_REQUEXCOMPRARTEX 1,'" & Right(varCod_GrupoTex.Text, 8) & "','" & TxtOp.Text & "',''"
                        opcionProv = 1
            End Select
        
            CadConn.Execute Strsql
        End If
        
        'DGridLista(opcionProv - 1).Row = j
'        If j >= 6 Then
'            DGridLista(opcionProv - 1).Scroll 0, 1
'            DGridLista(opcionProv - 1).Row = 5
'        End If
        'DGridLista(opcionProv - 1).ROW = DGridLista(opcionProv - 1).ROW + 1
        DGridLista(opcionProv - 1).Bookmark = (j + 1)
        
    Next
    Set CadConn = Nothing
    Unload Me
    Exit Sub
ErrorAceptar:
Set CadConn = Nothing
ErrorHandler Err, "Error Aceptar"
End Sub

Private Sub cmdBuscaColor_Click()
    Set frmBusqGeneral.oParent = Me
    Strsql = "exec UP_SEL_COLORDCOMPTEX '" & varCod_GrupoTex & "','" & TxtOp.Text & "','" & Me.varSer_OrdComp & "','" & Me.varCod_OrdComp & "'"
    frmBusqGeneral.sQuery = Strsql
    frmBusqGeneral.CARGAR_DATOS
    frmBusqGeneral.Show 1
    TxtColor.Text = Codigo
    Codigo = ""

End Sub

Private Sub cmdBuscaFam_Click()

    Strsql = "exec UP_SEL_FAMORDCOMPTEX '" & Me.varSer_OrdComp & "','" & Me.varCod_OrdComp & "','" & varCod_GrupoTex & "','" & TxtOp.Text & "'"
    
'    Select Case varCod_TipRequ
'        Case 2:
'                Strsql = "exec UP_SEL_FAMORDCOMPTEX 4,'" & varCod_GrupoTex & "','" & TxtOp.Text & "'"
'        Case 3:
'                Strsql = "exec UP_SEL_FAMORDCOMPTEX 3,'" & varCod_GrupoTex & "','" & TxtOp.Text & "'"
'        Case 4:
'                Strsql = "exec UP_SEL_FAMORDCOMPTEX 2,'" & varCod_GrupoTex & "','" & TxtOp.Text & "'"
'        Case 5:
'                Strsql = "exec UP_SEL_FAMORDCOMPTEX 1,'" & varCod_GrupoTex & "','" & TxtOp.Text & "'"
'    End Select

    'Strsql = "EXEC UP_SEL_FAMORDCOMPTEX '" & varCod_GrupoTex & "','" & TxtOp.Text & "'"
    
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = Strsql
    frmBusqGeneral.CARGAR_DATOS
    frmBusqGeneral.Show 1
    TxtFamilia.Text = Codigo
    Codigo = ""

End Sub

Private Sub cmdBuscaOP_Click()
    Set frmBusqGeneral.oParent = Me
    frmBusqGeneral.sQuery = "Select cod_ordpro as Codigo, convert(char(10),fec_creacion,103) as [Fecha Creacion] from ES_ORDPRO where Cod_GrupoTex='" & varCod_GrupoTex & "' order by 1"
    frmBusqGeneral.CARGAR_DATOS
    frmBusqGeneral.Show 1
    If TxtOp <> Codigo Then
        TxtFamilia.Text = ""
    End If
    TxtOp = Codigo
End Sub

Private Sub cmdBuscar_Click()
    Call CARGA_LISTA(varCod_TipRequ)
    cmdAceptar.Enabled = True
    Call CalculaTotal(opcionProv - 1)
    FlagN = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub DGridLista_AfterUpdate(Index As Integer, RtnDispErrMsg As Integer)
    Dim ubic_cant As Integer
    Select Case Index
        Case 0:
                ubic_cant = 7
        Case 1:
                ubic_cant = 10
        Case 2:
                ubic_cant = 12
        Case 3:
                ubic_cant = 16
    End Select

    If Len(Trim(DGridLista(Index).Columns(ubic_cant).Text)) = 0 Then DGridLista(Index).Columns(ubic_cant).Text = "0"
   'If FlagN = False Then Call CalculaTotal(Index)
End Sub

Private Sub DGridLista_Change(Index As Integer)
    Dim ubic_cant As Integer
    Dim ubic_final As Integer
    Select Case Index
        Case 0:
                ubic_cant = 7
                ubic_final = 10
        Case 1:
                ubic_cant = 10
                ubic_final = 13
        Case 2:
                ubic_cant = 12
                ubic_final = 17
        Case 3:
                ubic_cant = 16
                ubic_final = 21
    End Select
    If Val(DGridLista(Index).Columns(ubic_cant).Text) > Val(DGridLista(Index).Columns(ubic_final).Text) Then
        MsgBox "El valor comprado no puede ser mayor a la disponible. Sirvase verificar", vbInformation, "Ordenes de Compra"
        DGridLista(Index).Columns(ubic_cant).Text = DGridLista(Index).Columns(ubic_final).Text
        Exit Sub
    End If
    If DGridLista(Index).Columns(0).Value = 0 Then
        txtTotal.Text = Val(txtTotal.Text) - Val(DGridLista(Index).Columns(ubic_cant).Text)
    Else
        If DGridLista(Index).Columns(0).Value = 1 Or DGridLista(Index).Columns(0).Value = -1 Then
            txtTotal.Text = Val(txtTotal.Text) + Val(DGridLista(Index).Columns(ubic_cant).Text)
        End If
    End If
    'If FlagN = False Then Call CalculaTotal(Index)
End Sub

Private Sub DGridLista_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ubic_cant As Integer
    Select Case Index
        Case 0:
                ubic_cant = 7
        Case 1:
                ubic_cant = 10
        Case 2:
                ubic_cant = 12
        Case 3:
                ubic_cant = 16
    End Select

    If DGridLista(Index).Col = ubic_cant Then
       Select Case KeyAscii
            Case 48 To 57
                    If Len(Trim(DGridLista(Index).Columns(ubic_cant).Text)) >= 5 Then KeyAscii = 0: Exit Sub
                    KeyAscii = KeyAscii
            Case 46
                    If Len(Trim(DGridLista(Index).Columns(ubic_cant).Text)) >= 5 Then KeyAscii = 0: Exit Sub
                    If InStr(1, DGridLista(Index).Columns(ubic_cant).Text, ".") > 0 Then
                        KeyAscii = 0
                    Else
                        KeyAscii = KeyAscii
                    End If
            Case 8
                    KeyAscii = KeyAscii
            Case Else
                    KeyAscii = 0
        End Select
    End If

End Sub

Private Sub Form_Load()
    FlagU = False
End Sub

Private Sub OpColor_Click()
    TxtFamilia.Enabled = False
    cmdBuscaFam.Enabled = False
    TxtFamilia.Text = ""
    TxtColor.Enabled = True
    cmdBuscaColor.Enabled = True
End Sub

Private Sub OpFam_Click()
    TxtFamilia.Enabled = True
    cmdBuscaFam.Enabled = True
    TxtColor.Enabled = False
    cmdBuscaColor.Enabled = False
    TxtColor.Text = ""
End Sub

Private Sub TxtOp_KeyPress(KeyAscii As Integer)
    Dim temp As String
    If KeyAscii = 13 Then
        TxtFamilia.Text = ""
        TxtOp = Trim(DevuelveCampo("Select dbo.uf_devuelvecodigo(5," & IIf(Trim(TxtOp) = "", 0, TxtOp) & ")", cConnect))
        If DevuelveCampo("select count(*) from ES_ORDPRO where cod_ordpro ='" & TxtOp & "' and Cod_GrupoTex='" & Right(varCod_GrupoTex, 8) & "'", cConnect) <= 0 Then
                MsgBox "Codigo no existe", vbInformation
        End If
        
    End If
End Sub

Sub CalculaTotal(Index As Integer)
Dim i As Integer
Dim vRow As Variant
Dim vTotal As Double

VB.Screen.MousePointer = 11
vTotal = 0
vRow = DGridLista(Index).Bookmark
DGridLista(Index).Redraw = False
DGridLista(Index).Bookmark = 0
DGridLista(Index).Redraw = False
FlagU = False
For i = 0 To DGridLista(Index).Rows - 1
    If DGridLista(Index).Columns(0).Value = 1 Or DGridLista(Index).Columns(0).Value = -1 Then
        vTotal = vTotal + DGridLista(Index).Columns("CANTXCOMPRAR").Value
    End If

    DGridLista(Index).Bookmark = (i + 1)
Next
DGridLista(Index).Bookmark = vRow
txtTotal.Text = vTotal
DGridLista(Index).Bookmark = 0
DGridLista(Index).Redraw = True
VB.Screen.MousePointer = 0
End Sub

