VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECNLIB03_WINEVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'************************************************************************************************************************************************************
' AUTOR       : EDGAR I. CARDENAS NAVARRO
' FECHA       : 09/05/2013
' NOMBRE      : ECNLIB03
' DESCRIPCION : CLASE CON PROCEDIMIENTOS Y FUNCIONES VINCULADAS A CONTROLES DE FORMULARIOS Y/O EVENTOS DE LOS MISMOS INCLUIDO LOS FORMULARIOS
'************************************************************************************************************************************************************

'============================================================================================================================================================
' Ctrl_GotFocus : Resalta el valor de un control al recibir el foco
'============================================================================================================================================================
Public Sub Ctrl_GotFocus(ByRef xCtrl As Object)
    On Error Resume Next
    Dim objControl As Control
    objControl = xCtrl
    Select Case TypeName(xCtrl)
        Case "TextBox", "MaskEdBox"
            xCtrl.SelStart = 0
            xCtrl.SelLength = Len(xCtrl.Text)
        Case "Label"
        Case "ComboBox"
        Case "ListBox"
        Case "OptionButton"
        Case "CheckBox"
        Case ""
    End Select
End Sub

'============================================================================================================================================================
' TXT_KeyPressSoloNumeros : Control TXT que solo permite numeros
'============================================================================================================================================================
Public Sub TXT_KeyPressSoloNumeros(ByRef xKeyAscii As Integer)
    On Error Resume Next
    Select Case xKeyAscii
        Case 8
        Case 13
        Case Else
            Select Case Chr(xKeyAscii)
                Case "0" To "9"
                Case Else
                    xKeyAscii = 0
            End Select
    End Select
End Sub

'============================================================================================================================================================
' TXT_KeyPressSoloNumerosLetras : Control TXT que solo permite numeros y letras
'============================================================================================================================================================
Public Sub TXT_KeyPressSoloNumerosLetras(ByRef xKeyAscii As Integer)
    On Error Resume Next
    Select Case xKeyAscii
        Case 8
        Case 13
        Case Else
            Select Case Chr(xKeyAscii)
                Case "0" To "9"
                Case "a" To "z"
                Case "A" To "Z"
                Case Else
                    xKeyAscii = 0
            End Select
    End Select
End Sub

'============================================================================================================================================================
' FormateaSQL_TXT : Control RICHTEXTBOX que contiene los formatos de fuente del query analizer del IDE de SQL, Trabaja con las siguientes funciones
'                   - FormateaSQL_TXT_Simbolos
'                   - EsUnaPalabraSQL_Reservada
'                   - EsUnaPalabraSQL_TipoString
'                   - EsUnaPalabraSQL_TipoDate
' Si se quiere tener un efecto dinamico en el resultado final, recomiendo utilizarlo en el evento CHANGE del control richtextbox del GUI
'============================================================================================================================================================
Public Sub FormateaSQL_TXT(ByRef Ctrl_rtfQuerySQL As Object)
    On Error Resume Next
    If UCase(TypeName(Ctrl_rtfQuerySQL)) <> "RICHTEXTBOX" Then Exit Sub
    
    Dim sCar As String * 1
    Dim sPalabra As String
    Dim sTxT As String
    
    
    Dim i As Integer
    Dim Icon As Integer
    Dim iPosINI As Integer
    Dim iPosFIN As Integer
    Dim iPosDentroDelControl As Integer
    
    Dim dblColor As Double
    Dim dblColorAux As Double
    
    Dim blSW_EsReserv As Boolean
    Dim blSW_EsCadena As Boolean
    Dim blSW_EsTFecha As Boolean
    Dim blSW_EsParSQL As Boolean
    
    iPosDentroDelControl = Ctrl_rtfQuerySQL.SelStart
    
    sTxT = Ctrl_rtfQuerySQL.Text
    For i = 1 To Len(sTxT)
        Icon = Icon + 1
        sCar = Mid(sTxT, i, 1)
        Select Case sCar
            Case Space(1), Chr(13), Chr(10), "(", ".", "'", ","
                If sCar = "(" Then sPalabra = sPalabra & sCar
                If sCar = "'" Then sPalabra = sPalabra & sCar
                
                dblColor = &H4000&
                blSW_EsReserv = EsUnaPalabraSQL_Reservada(sPalabra, dblColorAux)
                If blSW_EsReserv = True Then dblColor = dblColorAux
                
                blSW_EsCadena = EsUnaPalabraSQL_TipoString(sPalabra, dblColorAux)
                If blSW_EsCadena = True Then dblColor = dblColorAux
                
                blSW_EsTFecha = EsUnaPalabraSQL_TipoDate(sPalabra, dblColorAux)
                If blSW_EsTFecha = True Then dblColor = dblColorAux
                
                blSW_EsParSQL = EsUnaPalabraSQL_Parametro(sPalabra)
                
                iPosFIN = i - 1
                With Ctrl_rtfQuerySQL
                    .SelStart = iPosINI - 1
                    .SelLength = iPosFIN - iPosINI + IIf(blSW_EsCadena = True, 2, 1)
                    .SelColor = dblColor
                    
                    If blSW_EsParSQL Then
                        .SelBold = True
                        Call SetSelBackColor(.hWnd, vbYellow)
                    End If
                    
                    .SelLength = 0
                    .SelStart = Len(sTxT)
                End With
                If sPalabra <> "'" Then sPalabra = Empty
                Icon = 0
                iPosINI = 0
                iPosFIN = 0
            Case Else
                If Icon = 1 Then iPosINI = i
                If sPalabra = "'" And Icon = 1 Then iPosINI = i - 1
                sPalabra = sPalabra & sCar
        End Select
    Next i
    Call FormateaSQL_TXT_Simbolos(Ctrl_rtfQuerySQL)
    If Ctrl_rtfQuerySQL.Enabled = True Then
        Ctrl_rtfQuerySQL.SelStart = iPosDentroDelControl
        Ctrl_rtfQuerySQL.SetFocus
    End If
End Sub

'============================================================================================================================================================
' FormateaSQL_TXT_Simbolos : Control RICHTEXTBOX que contiene los formatos de fuente para los simbolos de la sintaxis SQL
'============================================================================================================================================================
Private Sub FormateaSQL_TXT_Simbolos(ByRef Ctrl_rtfQuerySQL As Object)
    On Error Resume Next
    If UCase(TypeName(Ctrl_rtfQuerySQL)) <> "RICHTEXTBOX" Then Exit Sub
    Dim sPalabrasClaveSQL As String
    Dim dblColorSQL As Double
    Dim sCar As String * 1
        
    Dim i As Integer
    Dim sTxT As String
        
    sPalabrasClaveSQL = ">,<,=,*,/,\,-,+,#,,"
    dblColorSQL = &H404040
    sTxT = Ctrl_rtfQuerySQL.Text
    For i = 1 To Len(sTxT)
        sCar = Mid(sTxT, i, 1)
        If InStr(1, sPalabrasClaveSQL, sCar, vbTextCompare) > 0 Then
            With Ctrl_rtfQuerySQL
                .SelStart = i
                .SelLength = 1
                .SelColor = dblColorSQL
                .SelLength = 0
                .SelStart = Len(sTxT)
            End With
        End If
    Next i
End Sub

'============================================================================================================================================================
' EsUnaPalabraSQL_Reservada : Verifica si la palabra que recibe como parametro es una palabra reservada de SQL y devuelve el color respectivo
'============================================================================================================================================================
Private Function EsUnaPalabraSQL_Reservada(ByVal sPalabraFND As String, _
                                           Optional ByRef dblColorRET As Double, _
                                           Optional ByVal dblColorDEF As Double = &H4000&) As Boolean
    On Error Resume Next
    
    Dim sPalabrasClaveSQL(1 To 5) As String
    Dim dblColorSQL(1 To 5) As Double
    
    Dim sPalabraArray As String
    Dim sCarSeparador As String * 1
    Dim sCar As String * 1
    
    Dim i As Integer
    Dim J As Integer
    
    Dim blSW As Boolean
       
    sPalabrasClaveSQL(1) = "SELECT,FROM,INNER,LEFT,RIGHT,JOIN,WHERE,GROUP,HAVING,ORDER,BY,EXEC,EXECUTE,OUTPUT,AS,DECLARE,,"
    sPalabrasClaveSQL(2) = "GETDATE(),DATEADD(,DATEDIFF(,DATENAME(,DATEPART(,DAY(,YEAR(,MONTH(,ISDATE(,ABS(,EXP(,CHAR(,LEN(,LOWER(,UPPER(,SPACE(,REPLICATE(,STR(,REVERSE(,SUBSTRING(,UNICODE(,LEFT(,RIGHT(,RTRIM(,LTRIM(,CAST(,CONVERT(,REPLACE(,CHARINDEX(,ISNULL(,ISNUMERIC(,USERNAME(),DBNAME(),,"
    sPalabrasClaveSQL(3) = "DBO.,MAX(,MIN(,AVG(,COUNT(,,"
    sPalabrasClaveSQL(4) = "AND,OR,IN(,NOT,BETWEEN,,"
    sPalabrasClaveSQL(5) = "CHAR(,VARCHAR(,NCHAR(,NVARCHAR(,TEXT,NTEXT,XML,BIT,IMAGE,INT,INTEGER,NUMERIC(,DATETIME,SMALLDATETIME,DECIMAL,MONEY,TABLE,,"
    
    dblColorSQL(1) = vbBlue
    dblColorSQL(2) = &HFF00FF
    dblColorSQL(3) = &H8000&
    dblColorSQL(4) = &H404040
    dblColorSQL(5) = vbRed
    
    sCarSeparador = ","
    dblColorRET = dblColorDEF
    
    EsUnaPalabraSQL_Reservada = False
    For J = 1 To 5
        sPalabraArray = Empty
        blSW = False
        For i = 1 To Len(sPalabrasClaveSQL(J))
            If blSW = True Then GoTo VERIFICAR_PALABRA
RETORNO:
            sCar = Mid(sPalabrasClaveSQL(J), i, 1)
            If sCar <> sCarSeparador Then
                sPalabraArray = sPalabraArray & sCar
            Else
                blSW = True
                GoTo CORRER_UNA_VUELTA
VERIFICAR_PALABRA:
                
                If UCase(sPalabraArray) = UCase(sPalabraFND) Then
                    dblColorRET = dblColorSQL(J)
                    EsUnaPalabraSQL_Reservada = True
                    Exit Function
                End If
                blSW = False
                sPalabraArray = Empty
                GoTo RETORNO
            End If
CORRER_UNA_VUELTA:
        Next i
    Next J
End Function

'============================================================================================================================================================
' EsUnaPalabraSQL_TipoString : Verifica si la palabra que recibe como parametro es una palabra de cadena SQL y devuelve el color respectivo
'============================================================================================================================================================
Private Function EsUnaPalabraSQL_TipoString(ByVal sPalabraFND As String, _
                                            Optional ByRef dblColorRET As Double, _
                                            Optional ByVal dblColorDEF As Double = &H4000&) As Boolean
    On Error Resume Next
    
    EsUnaPalabraSQL_TipoString = False
    If Len(sPalabraFND) <= 1 Then Exit Function
    dblColorRET = dblColorDEF
    If Left(sPalabraFND, 1) = "'" And Right(sPalabraFND, 1) = "'" Then
        EsUnaPalabraSQL_TipoString = True
        dblColorRET = vbRed
    End If
End Function

'============================================================================================================================================================
' EsUnaPalabraSQL_TipoDate : Verifica si la palabra que recibe como parametro es una palabra de fecha SQL y devuelve el color respectivo
'============================================================================================================================================================
Private Function EsUnaPalabraSQL_TipoDate(ByVal sPalabraFND As String, _
                                          Optional ByRef dblColorRET As Double, _
                                          Optional ByVal dblColorDEF As Double = &H4000&) As Boolean
    On Error Resume Next
    
    EsUnaPalabraSQL_TipoDate = False
    dblColorRET = dblColorDEF
    If Left(sPalabraFND, 1) = "#" And Right(sPalabraFND, 1) = "#" Then
        EsUnaPalabraSQL_TipoDate = True
        dblColorRET = vbRed
    End If
End Function

'============================================================================================================================================================
' EsUnaPalabraSQL_Parametro : Verifica si la palabra que recibe como parametro es un parametro SQL
'============================================================================================================================================================
Private Function EsUnaPalabraSQL_Parametro(ByVal sPalabraFND As String, _
                                           Optional ByRef dblColorRET As Double) As Boolean
    On Error Resume Next
    
    EsUnaPalabraSQL_Parametro = False
    If Left(sPalabraFND, 1) = "@" And Len(sPalabraFND) > 1 Then
        EsUnaPalabraSQL_Parametro = True
        dblColorRET = &H800080
    End If
End Function

'============================================================================================================================================================
' GRILLA_IniciaScroll : Inicializa Scroll de la Grilla, puede ser MSHFLEXGRID, FLEXGRID, DATAGRID
'============================================================================================================================================================
Public Sub GRILLA_IniciaScroll(ByRef Ctrl_Grilla As Object)
    On Error Resume Next
    Select Case UCase(TypeName(Ctrl_Grilla))
        Case "MSHFLEXGRID", "DATAGRID", "MSFLEXGRID"
        Case Else
            Exit Sub
    End Select
    
    Call GRILLA_IniciaScroll(Ctrl_Grilla)
End Sub

'============================================================================================================================================================
' GRILLA_IniciaScroll : Inicializa Scroll de la Grilla, puede ser MSHFLEXGRID, FLEXGRID, DATAGRID
'============================================================================================================================================================
Public Sub GRILLA_DetieneScroll(ByRef Ctrl_Grilla As Object)
    On Error Resume Next
    Select Case UCase(TypeName(Ctrl_Grilla))
        Case "MSHFLEXGRID", "DATAGRID", "MSFLEXGRID"
        Case Else
            Exit Sub
    End Select
    Call Grilla_DetenerScroll(Ctrl_Grilla)
End Sub
