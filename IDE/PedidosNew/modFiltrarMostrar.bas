Attribute VB_Name = "modFiltrarMostrar"
Option Explicit

Public oMDIParent As Object

Public Function Filtrar(ByVal sFlag As String, _
                        ByRef oForm As Object, _
                        Optional ByRef oControl As Variant, _
                        Optional ByRef oControlDes As Variant, _
                        Optional ByVal bShowMostrar As Variant, _
                        Optional ByVal strCodClaPO As String = "") As Boolean

    On Error GoTo errores

    Dim obj          As Object

    Dim vbuff        As Variant

    Dim i            As Integer

    Dim j            As Integer

    Dim lExecFiltrar As Boolean

    Dim lFoco        As Boolean

    Dim iExiste      As Integer

    lFoco = False
    Filtrar = False
    lExecFiltrar = True
    
    If IsMissing(bShowMostrar) Then
        bShowMostrar = True
    End If
    
    If IsMissing(oControl) Then
        Set oControl = oForm.ActiveControl
    Else

        If oControl Is Nothing Then
            lExecFiltrar = False
        End If
    End If

    If lExecFiltrar Then

        Select Case sFlag

            Case Is = "ABR_CLIENTE"
                Set obj = New clsTG_Cliente
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewClientes(oControl.Text, vusu)

            Case Is = "COD_TEMCLI"
                Set obj = New clsTG_TemCli
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewTemCli(oForm.sCod_Cliente, oControl.Text)

            Case Is = "COD_TIPPRE"
                Set obj = New clsTG_PurOrd
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewTipPre(oControl.Text)

            Case "COD_COLCLIPRE", "COD_COLCLIPRE2"
                Set obj = New clsTG_ColCli
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewTG_ColCli(oForm.sCod_Cliente, oControl.Text)

            Case Is = "COD_GRUTAL"
                Set obj = New clsTG_PurOrd
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewGruTal(oControl.Text)

            Case "ABR_FABRICA", "ABR_FABRICALOT"
                Set obj = New clsTG_Fabrica
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewFabricas(oControl.Text, vusu, oForm.sCod_Cliente)

            Case "COD_DESTINO", "COD_DESTINOLOT"
                Set obj = New clsTG_Destino
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewDestinos(oControl.Text)
            
            Case "COD_DIVPRE"
                Set obj = New clsTG_LotColTal
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewDivPre(oControl.Text)

            Case "COD_ESTPRO"
                oControl.Text = StrZero1(oControl.Text, 5)
                Set obj = New clsTG_PurOrd
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewEstPropio(oControl.Text)

            Case "COD_DIVCLI"
                Set obj = New clsTG_Cliente
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewDivClientes(oForm.sCod_Cliente, oControl.Text)

            Case "COD_PAGEMB"
                Set obj = New clsTG_Cliente
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewPagEmbarque(oControl.Text)
        
            Case "COD_TIPEMB"
                Set obj = New clsTG_Cliente
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewTipEmb(oControl.Text)
        
            Case "COD_BANCO"
                Set obj = New clsTG_Cliente
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewBanco(oControl.Text)
        
            Case "COD_GRUPO"
                Set obj = New clsTG_Cliente
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewGrupo(oForm.sCod_Cliente, oControl.Text)

            Case "COD_MONEDA"
                Set obj = New clsTG_Cliente
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewMoneda(oControl.Text)
     
            Case Is = "COD_ESTCLI"
                Set obj = New clsTG_EstclIEst
                obj.ConexionString = cCONNECT
                vbuff = obj.ViewEStCliEst(oForm.sCod_Cliente, oForm.txtCod_TemCli.Text, strCodClaPO)

            Case Is = "COD_ESTPROPIO"

                '            Set obj = New clsTG_EstPropio
                '            obj.ConexionString= cCONNECT
                '            vbuff = obj.ViewEstPropio(oControl.Text)
            Case Is = "COD_ESTPROPIO_DES"

                '            Set obj = New clsTG_EstPropio
                '            obj.ConexionString= cCONNECT
                '            vbuff = obj.ViewEstPropio1(oControlDes.Text)
            Case Is = "COD_AYUDAOP"
                Set obj = New clsTG_PurOrd
                obj.ConexionString = cCONNECT
                vbuff = obj.AyudaAsignaOPS(oForm.sCod_Fabrica, oForm.sCod_Cliente, oForm.sCod_EstPro)
        End Select

        Set obj = Nothing

        If IsEmpty(vbuff) = False Then
            If UBound(vbuff, 2) + 1 = 1 Then

                Select Case sFlag

                    Case Is = "ABR_CLIENTE"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)
                        oForm.sCod_Cliente = FixNulos(vbuff(2, 0), vbstring)
                        oForm.sNivAccUsuario = FixNulos(vbuff(3, 0), vbstring)
                        oForm.dPor_ComisionCliente = FixNulos(vbuff(4, 0), vbDouble)
                        oForm.SetFormCliente (oForm.sNivAccUsuario)

                    Case Is = "COD_TEMCLI"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)

                    Case Is = "COD_COLCLIPRE"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)

                        'oControlDes.Enabled = False
                    Case Is = "COD_COLCLIPRE2"

                    Case Is = "COD_ESTPRO"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)

                    Case Is = "COD_ESTPROPIO"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)

                    Case Is = "COD_ESTPROPIO_DES"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)

                    Case Is = "COD_TIPPRE"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)

                    Case Is = "COD_GRUTAL"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)

                    Case "ABR_FABRICA", "ABR_FABRICALOT"

                        If FixNulos(vbuff(3, 0), vbstring) = "V" Then
                            Mensaje kMESSAGE_ERR_NOT_RIGHT_OPTION
                            Filtrar = False

                            Exit Function

                        End If

                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)

                        If sFlag = "ABR_FABRICA" Then
                            oForm.sCod_Fabrica = FixNulos(vbuff(2, 0), vbstring)
                        Else
                            oForm.sCod_FabricaLot = FixNulos(vbuff(2, 0), vbstring)
                        End If

                    Case "COD_DESTINO", "COD_DESTINOLOT"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)

                        If sFlag = "COD_DESTINO" Then
                            oForm.sCod_Destino = FixNulos(vbuff(0, 0), vbstring)
                        Else
                            oForm.sCod_DestinoLOT = FixNulos(vbuff(0, 0), vbstring)
                        End If

                    Case "COD_DIVCLI"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)
                    
                    Case "COD_PAGEMB"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)
                    
                    Case "COD_DIVPRE"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                    
                    Case "COD_TIPEMB"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)
                    
                    Case "COD_BANCO"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)
                    
                    Case "COD_GRUPO"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)
                    
                    Case "COD_MONEDA"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                        oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)

                    Case "COD_ESTCLI"
                        oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                 
                    Case Is = "COD_AYUDAOP"
                        
                        If oForm.Name <> "frmViewOPs" Then
                            oControl.Text = FixNulos(vbuff(0, 0), vbstring)
                            oControlDes.Text = FixNulos(vbuff(1, 0), vbstring)
                        End If
                        
                        If bShowMostrar Then
                            Filtrar = Mostrar(sFlag, oForm, oControl, oControlDes)
                        Else
                            Filtrar = False
                        End If
                                     
                End Select

                Filtrar = True

            Else

                If bShowMostrar Then
                    Filtrar = Mostrar(sFlag, oForm, oControl, oControlDes, strCodClaPO)
                Else
                    Filtrar = False
                End If
            End If

        Else

            If bShowMostrar Then

                Select Case sFlag

                    Case "COD_DESTINO", "COD_DIVCLI", "COD_PAGEMB", "COD_TIPEMB", "COD_BANCO", "COD_ESTPROPIO"
                        Filtrar = False

                    Case Else
                        Filtrar = Mostrar(sFlag, oForm, oControl, oControlDes)
                End Select

            Else
                Filtrar = False
            End If
        End If

    Else

        If bShowMostrar Then
            Filtrar = Mostrar(sFlag, oForm, oControl, oControlDes)
        Else
            Filtrar = False
        End If
    End If

    Exit Function

errores:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If

    ErrorHandler Err, Err.Description
End Function

Public Function Mostrar(ByVal sFlag As String, _
                        ByRef oForm As Object, _
                        ByRef oControl As Variant, _
                        ByRef oControlDes As Variant, _
                        Optional ByVal strCodClaPO As String = "") As Boolean

    On Error GoTo errores

    Dim vbuff           As Variant

    Dim obj             As Object

    Dim lDefaultControl As Boolean

    Dim lFoco           As Boolean

    lFoco = False
    Mostrar = False
    lDefaultControl = False

    Select Case sFlag

        Case Is = "ABR_CLIENTE"
            Set obj = New clsTG_Cliente
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewClientes("", vusu)
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal

        Case Is = "COD_ESTCLI"
            Set obj = New clsTG_EstclIEst

            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewEStCliEst(oForm.sCod_Cliente, oForm.txtCod_TemCli.Text, strCodClaPO)
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal

        Case "COD_COLCLIPRE", "COD_COLCLIPRE2"
            Set obj = New clsTG_ColCli
            obj.ConexionString = cCONNECT
            frmDatos.Buffer = obj.ViewTG_ColCli(oForm.sCod_Cliente, oControl.Text)
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal

        Case Is = "COD_ESTPROPIO_DES"
            '               Set obj = New clsTG_EstPropio
            '
            '               obj.ConexionString= cCONNECT
            '               Set frmDatos = frmDatos
            '               frmDatos.Buffer = obj.ViewEstPropio1(oControlDes.Text)
            '               frmDatos.FormatString = "Codigo|Descripción"
            '               frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            '               frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            '               frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            '               frmDatos.Height = 4000
            '               frmDatos.Width = 3000
            '               frmDatos.Top = 2000
            '               frmDatos.Left = 2500
            '               frmDatos.Show vbModal
               
        Case Is = "COD_TEMCLI"
            Set obj = New clsTG_TemCli
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewTemCli(oForm.sCod_Cliente, oControl.Text)
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal

        Case Is = "COD_TIPPRE"
            Set obj = New clsTG_PurOrd
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewTipPre(oControl.Text)
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal
               
        Case Is = "COD_GRUTAL"
            Set obj = New clsTG_PurOrd
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewGruTal(oControl.Text)
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal
                             
        Case "ABR_FABRICA", "ABR_FABRICALOT"
            Set obj = New clsTG_Fabrica
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewFabricas("", vusu, oForm.sCod_Cliente)
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal

        Case "COD_DESTINO", "COD_DESTINOLOT"
            Set obj = New clsTG_Destino
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewDestinos("")
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal
               
        Case "COD_DIVPRE"
            Set obj = New clsTG_LotColTal
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewDivPre("")
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal
               
        Case "COD_DIVCLI"
            Set obj = New clsTG_Cliente
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewDivClientes(oForm.sCod_Cliente, oControl.Text)
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal

        Case "COD_PAGEMB"
            Set obj = New clsTG_Cliente
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewPagEmbarque("")
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal
        
        Case "COD_TIPEMB"
            Set obj = New clsTG_Cliente
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewTipEmb("")
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal
        
        Case "COD_BANCO"
            Set obj = New clsTG_Cliente
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewBanco("")
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal
               
        Case "COD_GRUPO"
            Set obj = New clsTG_Cliente
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewGrupo(oForm.sCod_Cliente, "")
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal
        
        Case "COD_MONEDA"
            Set obj = New clsTG_Cliente
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewMoneda("")
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal

        Case Is = "COD_TEMCLI"
            Set obj = New clsTG_EstclIEst

            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.ViewEStCli(oForm.sCod_Cliente, oControl.Text)
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal
        
        Case Is = "COD_AYUDAOP"
            Set obj = New clsTG_PurOrd
               
            obj.ConexionString = cCONNECT
            Set frmDatos = frmDatos
            frmDatos.Buffer = obj.AyudaAsignaOPS(oForm.sCod_Fabrica, oForm.sCod_Cliente, oForm.sCod_EstPro)
            frmDatos.FormatString = "Codigo|Descripción"
            frmDatos.ColumnWidths = Array(1000, 3000, 0, 0, 0, 0, 0)
            frmDatos.ssgrdDatos.ColAlignment(0) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
            frmDatos.ssgrdDatos.ColAlignment(1) = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter

            frmDatos.Height = 4000
            frmDatos.Width = 3000
            frmDatos.Top = 2000
            frmDatos.Left = 2500
            frmDatos.Show vbModal
    
    End Select

    Set obj = Nothing

    If frmDatos.OK = True And frmDatos.DataFound Then

        Select Case sFlag

            Case Is = "ABR_CLIENTE"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)
                oForm.sCod_Cliente = FixNulos(frmDatos.TextArray(2), vbstring)
                oForm.sNivAccUsuario = FixNulos(frmDatos.TextArray(3), vbstring)
                oForm.dPor_ComisionCliente = FixNulos(frmDatos.TextArray(4), vbstring)
                oForm.SetFormCliente (oForm.sNivAccUsuario)

            Case Is = "COD_ESTPROPIO"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)

            Case Is = "COD_ESTPROPIO_DES"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)

            Case Is = "COD_COLCLIPRE"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)

                'oControlDes.Enabled = False
            Case Is = "COD_COLCLIPRE2"

            Case Is = "COD_TEMCLI"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)

            Case Is = "COD_TIPPRE"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)

            Case Is = "COD_GRUTAL"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)

            Case "ABR_FABRICA", "ABR_FABRICALOT"

                If FixNulos(frmDatos.TextArray(3), vbstring) = "V" Then
                    Unload frmDatos
                    Set frmDatos.RefObject = Nothing
                    Set frmDatos = Nothing
                    Mensaje kMESSAGE_ERR_NOT_RIGHT_OPTION
                    Mostrar = False

                    Exit Function

                End If

                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)
                    
                If sFlag = "ABR_FABRICA" Then
                    oForm.sCod_Fabrica = FixNulos(frmDatos.TextArray(2), vbstring)
                Else
                    oForm.sCod_FabricaLot = FixNulos(frmDatos.TextArray(2), vbstring)
                End If

            Case "COD_DESTINO", "COD_DESTINOLOT"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)
                    
                If sFlag = "COD_DESTINO" Then
                    oForm.sCod_Destino = FixNulos(frmDatos.TextArray(0), vbstring)
                Else
                    oForm.sCod_DestinoLOT = FixNulos(frmDatos.TextArray(0), vbstring)
                End If

            Case "COD_DIVCLI"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)

            Case "COD_DIVPRE"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                    
            Case "COD_PAGEMB"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)
                
            Case "COD_TIPEMB"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)
                
            Case "COD_BANCO"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)
                    
            Case "COD_GRUPO"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)
                
            Case "COD_MONEDA"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
                oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbstring)
                    
            Case "COD_ESTCLI"
                oControl.Text = FixNulos(frmDatos.TextArray(0), vbstring)
              
            Case Is = "COD_AYUDAOP"
                'oControl.Text = FixNulos(frmDatos.TextArray(0), vbString)
                'oControlDes.Text = FixNulos(frmDatos.TextArray(1), vbString)
                oForm.sCod_OrdPro = FixNulos(frmDatos.TextArray(0), vbstring)
                    
        End Select

        Mostrar = True
    End If

    Unload frmDatos
    Set frmDatos.RefObject = Nothing
    Set frmDatos = Nothing

    Exit Function

errores:

    If Not obj Is Nothing Then
        Set obj = Nothing
    End If

    ErrorHandler Err, Err.Description
End Function

Public Function ControlKeyDown(KeyCode As Integer, _
                               ByRef oForm As Object, _
                               ByRef oControl As Object, _
                               ByRef oControlDes As Object) As Boolean

    If KeyCode = vbKeyReturn Then
        ControlKeyDown = Filtrar(oForm.sFlag, oForm, oControl, oControlDes)
    End If

End Function

