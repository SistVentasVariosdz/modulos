VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form FrmReqCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requerimientos por comprar"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   11490
   StartUpPosition =   2  'CenterScreen
   Begin SSDataWidgets_B.SSDBGrid DGridLista 
      Height          =   2925
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   11445
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   18
      AllowColumnShrinking=   0   'False
      SelectTypeRow   =   1
      BackColorOdd    =   13434879
      RowHeight       =   423
      ExtraHeight     =   106
      Columns.Count   =   18
      Columns(0).Width=   1614
      Columns(0).Caption=   "Flag  O/P"
      Columns(0).Name =   "chec"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Style=   2
      Columns(1).Width=   1614
      Columns(1).Caption=   "Serie"
      Columns(1).Name =   "serie"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2355
      Columns(2).Caption=   "Item"
      Columns(2).Name =   "item"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   3493
      Columns(3).Caption=   "Descripcion Item"
      Columns(3).Name =   "des_item"
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      Columns(4).Width=   3200
      Columns(4).Caption=   "Combinacion"
      Columns(4).Name =   "des_compest"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(5).Width=   2540
      Columns(5).Caption=   "Color"
      Columns(5).Name =   "color"
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   2037
      Columns(6).Caption=   "Talla"
      Columns(6).Name =   "talla"
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(7).Width=   2143
      Columns(7).Caption=   "Destino"
      Columns(7).Name =   "destino"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(7).Locked=   -1  'True
      Columns(8).Width=   2699
      Columns(8).Caption=   "Estilo Cliente"
      Columns(8).Name =   "cod_estilo"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   2223
      Columns(9).Caption=   "Cantidad"
      Columns(9).Name =   "cantidad"
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "cod_item"
      Columns(10).Name=   "cod_item"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   3200
      Columns(11).Visible=   0   'False
      Columns(11).Caption=   "cod_comb"
      Columns(11).Name=   "cod_comb"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "cod_color"
      Columns(12).Name=   "cod_color"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "cod_destino"
      Columns(13).Name=   "cod_destino"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   3200
      Columns(14).Visible=   0   'False
      Columns(14).Caption=   "Estilo"
      Columns(14).Name=   "estilo"
      Columns(14).CaptionAlignment=   2
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(14).Locked=   -1  'True
      Columns(15).Width=   3200
      Columns(15).Visible=   0   'False
      Columns(15).Caption=   "cod_talla"
      Columns(15).Name=   "cod_talla"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(16).Width=   3200
      Columns(16).Caption=   "Observaciones"
      Columns(16).Name=   "Observaciones"
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   8
      Columns(16).FieldLen=   256
      Columns(17).Width=   3200
      Columns(17).Caption=   "Cod_Prov"
      Columns(17).Name=   "Cod_Prov"
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   8
      Columns(17).FieldLen=   256
      Columns(17).Locked=   -1  'True
      _ExtentX        =   20188
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
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancelar"
      Height          =   525
      Left            =   6548
      TabIndex        =   2
      Top             =   3180
      Width           =   1245
   End
   Begin SSDataWidgets_B.SSDBGrid Grilla 
      Height          =   1425
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   7860
      _Version        =   196617
      DataMode        =   2
      Col.Count       =   15
      AllowColumnShrinking=   0   'False
      SelectTypeRow   =   1
      BackColorOdd    =   13434879
      RowHeight       =   423
      ExtraHeight     =   132
      Columns.Count   =   15
      Columns(0).Width=   1429
      Columns(0).Caption=   "Flag O/P"
      Columns(0).Name =   "chec"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Style=   2
      Columns(1).Width=   1429
      Columns(1).Caption=   "Serie"
      Columns(1).Name =   "Serie"
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2672
      Columns(2).Caption=   "Item"
      Columns(2).Name =   "Item"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   3200
      Columns(3).Caption=   "Descripcion Item"
      Columns(3).Name =   "des_item"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "Combinacion"
      Columns(4).Name =   "des_compest"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(4).PromptInclude=   -1  'True
      Columns(5).Width=   2064
      Columns(5).Caption=   "Color"
      Columns(5).Name =   "Color"
      Columns(5).CaptionAlignment=   2
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).Locked=   -1  'True
      Columns(6).Width=   1746
      Columns(6).Caption=   "Talla"
      Columns(6).Name =   "Talla"
      Columns(6).CaptionAlignment=   2
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).Case =   2
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(7).Width=   1720
      Columns(7).Caption=   "Destino"
      Columns(7).Name =   "Destino"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).Case =   2
      Columns(7).FieldLen=   256
      Columns(7).Locked=   -1  'True
      Columns(8).Width=   2275
      Columns(8).Caption=   "Estilo"
      Columns(8).Name =   "Estilo"
      Columns(8).CaptionAlignment=   2
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).Case =   2
      Columns(8).FieldLen=   256
      Columns(8).Locked=   -1  'True
      Columns(9).Width=   1614
      Columns(9).Caption=   "Cantidad"
      Columns(9).Name =   "cantidad"
      Columns(9).CaptionAlignment=   2
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   5
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "Cod_Item"
      Columns(10).Name=   "Cod_Item"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   3200
      Columns(11).Visible=   0   'False
      Columns(11).Caption=   "Cod_Comb"
      Columns(11).Name=   "Cod_Comb"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "Cod_Color"
      Columns(12).Name=   "Cod_Color"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "Cod_Destino"
      Columns(13).Name=   "Cod_Destino"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   3200
      Columns(14).Visible=   0   'False
      Columns(14).Caption=   "cod_estilo"
      Columns(14).Name=   "cod_estilo"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      _ExtentX        =   13864
      _ExtentY        =   2514
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   525
      Left            =   3698
      TabIndex        =   0
      Top             =   3180
      Width           =   1245
   End
End
Attribute VB_Name = "FrmReqCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cod_TipMov  As String
Public Ser_OrdComp
Public Cod_OrdComp
Public Codigo As String
Public Descripcion As String
Dim CadCon  As New ADODB.Connection
Public Sub Buscar()
On Error GoTo hand
Dim Rs As New ADODB.Recordset
Dim i As Integer

Set Rs = Nothing
Rs.CursorLocation = adUseClient
Rs.Open "UP_Lg_OrdCompItem '" & Ser_OrdComp & "','" & Cod_OrdComp & "','" & Cod_TipMov & "'", cConnect
'rs.Open "UP_Lg_OrdCompItem '" & Ser_OrdComp & "','000016','" & Cod_TipMov & "'", cCONNECT

Dim AddItemString As String
  
Grilla.FieldSeparator = vbTab

If Rs.RecordCount > 0 Then
        Rs.MoveFirst
'        Do Until rs.EOF
'            AddItemString = vbTab
'            AddItemString = AddItemString & rs("serie")
'            AddItemString = AddItemString & vbTab
'            AddItemString = AddItemString & rs("item")
'            AddItemString = AddItemString & vbTab
'            AddItemString = AddItemString & rs("des_item")
'            AddItemString = AddItemString & vbTab
'            AddItemString = AddItemString & rs("combinacion")
'            AddItemString = AddItemString & vbTab
'
'            AddItemString = AddItemString & rs("color")
'            AddItemString = AddItemString & vbTab
'            AddItemString = AddItemString & rs("talla")
'            AddItemString = AddItemString & vbTab
'
'            AddItemString = AddItemString & rs("destino")
'            AddItemString = AddItemString & vbTab
'            AddItemString = AddItemString & rs("estilo")
'            AddItemString = AddItemString & vbTab
'            AddItemString = AddItemString & rs("cantidad")
'
'
'            AddItemString = AddItemString & vbTab
'            AddItemString = AddItemString & rs("Cod_Item")
'            AddItemString = AddItemString & vbTab
'            AddItemString = AddItemString & rs("Cod_Comb")
'            AddItemString = AddItemString & vbTab
'            AddItemString = AddItemString & rs("Cod_Color")
'            AddItemString = AddItemString & vbTab
'            AddItemString = AddItemString & rs("cod_destino")
'            AddItemString = AddItemString & vbTab
'            AddItemString = AddItemString & rs("cod_estcli")
            
            'Grilla.AddItem AddItemString
'            DGridLista.AddItem AddItemString
'            rs.MoveNext
'        Loop
'        Grilla.AddItem " "
        
    Dim Rs_Prov As New ADODB.Recordset
    Set Rs_Prov = Rs.Clone
    
    Me.DGridLista.Redraw = False
    SSDBGridSetGrid Me.DGridLista
    ADODBToSSDBGridOC Rs_Prov, DGridLista
    DGridLista.ActiveRowStyleSet = "RowActive"
    DGridLista.SelectTypeRow = ssSelectionTypeMultiSelectRange
    DGridLista.Visible = True
        
        
'        For i = 0 To Grilla.Rows
'            Grilla.Bookmark = i
'            If i >= 7 Then Grilla.Scroll 0, 1
'            Grilla.Columns(0).Value = True
'            If i = Grilla.Rows - 1 Then Grilla.RemoveItem rs.RecordCount: Exit Sub
'        Next

        DGridLista.Row = 0
            For i = 0 To DGridLista.Rows
            DGridLista.Bookmark = i
                If i >= 6 Then DGridLista.Scroll 0, 1
                DGridLista.Columns(0).Value = 1
                If i = DGridLista.Rows - 1 Then Exit Sub
                'DGridLista.Row = DGridLista.Row + 1
            Next
        
Else
    MsgBox "No se encontraron registros ", vbInformation
End If
Set Rs = Nothing
Exit Sub

hand:
ErrorHandler err, "Buscar"
Set Rs = Nothing
End Sub


Public Sub SoloNumeros(ByVal pTextbox As TextBox, _
                       ByRef pKeyAscii As Integer, _
                       Optional ByVal pConDecimales As Boolean, _
                       Optional ByVal pNumDecimales As Integer, _
                       Optional ByVal pNumEnteros As Integer)
   If pNumEnteros = 0 Then pNumEnteros = 10
   If pKeyAscii = 8 Then
      If pConDecimales And pTextbox.SelStart > 0 Then
         If Mid(pTextbox, pTextbox.SelStart, 1) = "." Then
            If Len(Mid(pTextbox, 1, pTextbox.SelStart - 1)) >= pNumEnteros And Len(Mid(pTextbox, pTextbox.SelStart + 1)) > 0 Then pKeyAscii = 0
         End If
      End If
      Exit Sub
   End If
   If pKeyAscii = 46 Then
      If pConDecimales Then
         If InStr(1, pTextbox, ".") > 0 Then
            pKeyAscii = 0
         Else
            If Len(Mid(pTextbox, pTextbox.SelStart + 1)) > pNumDecimales Then pKeyAscii = 0
            If pTextbox.SelStart > 0 Then If Len(Mid(pTextbox, 1, pTextbox.SelStart - 1)) >= pNumEnteros Then pKeyAscii = 0
         End If
      Else
         pKeyAscii = 0
      End If
   Else
      If Not (pKeyAscii >= 48 And pKeyAscii <= 57) Then pKeyAscii = 0
      If pKeyAscii = 39 Or pKeyAscii = 13 Then
         pKeyAscii = 0
      End If
      
      Dim iPos As Integer
      iPos = InStr(1, pTextbox, ".")
      If iPos > 0 And pConDecimales Then _
         If Len(Mid(pTextbox, iPos)) > pNumDecimales Then _
            If InStr(pTextbox.SelStart + 1, pTextbox, ".") = 0 Then pKeyAscii = 0
            
      If pTextbox.SelStart < iPos Or iPos = 0 Then
         If pNumEnteros > 0 Then
            If InStr(pTextbox.SelStart + 1, pTextbox, ".") > 0 Then
               If Len(Mid(pTextbox, 1, InStr(pTextbox.SelStart + 1, pTextbox, ".") - 1)) >= pNumEnteros Then pKeyAscii = 0
            Else
               If Len(pTextbox) >= pNumEnteros Then pKeyAscii = 0
            End If
         End If
      End If
   End If
End Sub






Private Sub Command1_Click()
On Error GoTo hand
Set CadConn = Nothing
CadConn.Open cConnect
Dim j As Integer
For j = 0 To DGridLista.Rows - 1
    'DGridLista.Row = j
    DGridLista.Bookmark = j
    If DGridLista.Columns(0).Value = "1" Then
        CadConn.Execute "UP_ACTUALIZA_STOCKS_ITEM '" & FrmDetalleStock.Cod_Almacen & "','" & _
        FrmDetalleStock.Num_MovStk & "','" & DGridLista.Columns("cod_item").Text & "','" & _
        DGridLista.Columns("cod_comb").Text & "','" & DGridLista.Columns("cod_color").Text & "','" & DGridLista.Columns("talla").Text & "','" & _
        DGridLista.Columns("cod_destino").Text & "','" & DGridLista.Columns("cod_estilo").Text & "','',0," & _
        DGridLista.Columns("cantidad").Text & ",'I','" & DGridLista.Columns("serie").Text & "','','" & vusu & "','" & DGridLista.Columns("cod_prov").Text & "'"
    End If
Next
Set CadConn = Nothing
FrmDetalleStock.Datos "V", False
Unload Me
Exit Sub
hand:
ErrorHandler err, "Actualizando"
    Set CadConn = Nothing
    FrmDetalleStock.Datos "V", False
    Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
'Buscar Trim(Right(CmbGrupo, 8)), TxtOp, TxtFamilia
End Sub

Private Sub DGridLista_AfterUpdate(RtnDispErrMsg As Integer)
    If Len(Trim(DGridLista.Columns("cantidad").Text)) = 0 Then DGridLista.Columns("cantidad").Text = "0"
End Sub

Private Sub DGridLista_KeyPress(KeyAscii As Integer)
'46 .
'48 0
'57 9
If DGridLista.Col = 8 Then
       Select Case KeyAscii
            Case 48 To 57
                    If Len(Trim(DGridLista.Columns("cantidad").Text)) >= 9 Then KeyAscii = 0: Exit Sub
                    KeyAscii = KeyAscii
            Case 46
                If Len(Trim(DGridLista.Columns("cantidad").Text)) >= 9 Then KeyAscii = 0: Exit Sub
                If InStr(1, DGridLista.Columns("cantidad").Text, ".") > 0 Then
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
'Buscar
'cCONNECT = "Provider=sqloledb;Server=SERVIDOR;Database=lives;UID=sa;pwd=;"

End Sub




Private Sub Grilla_AfterUpdate(RtnDispErrMsg As Integer)
    If Len(Trim(Grilla.Columns(8).Text)) = 0 Then Grilla.Columns(8).Text = "0"
End Sub


Private Sub Grilla_Change()
'    If Len(Trim(Grilla.Columns(8).Text)) = 0 Then Grilla.Columns(8).Text = "0"
End Sub


Private Sub Grilla_KeyPress(KeyAscii As Integer)
'46 .
'48 0
'57 9
If Grilla.Col = 8 Then
       Select Case KeyAscii
            Case 48 To 57
                    If Len(Trim(Grilla.Columns(8).Text)) >= 9 Then KeyAscii = 0: Exit Sub
                    KeyAscii = KeyAscii
            Case 46
                If Len(Trim(Grilla.Columns(8).Text)) >= 9 Then KeyAscii = 0: Exit Sub
                If InStr(1, Grilla.Columns(8).Text, ".") > 0 Then
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




