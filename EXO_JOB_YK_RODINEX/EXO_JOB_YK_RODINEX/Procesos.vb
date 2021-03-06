Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml
Imports WinSCP

Public Class Procesos
    Private Shared log As EXO_Log.EXO_Log = Nothing
    Public Shared Function ActivaProceso(ByVal sProceso As String) As Boolean
        ActivaProceso = False
        Dim myStream As Stream = Nothing
        Dim Reader As XmlTextReader = Nothing
        Dim sHora As String = ""
        Try
            myStream = File.OpenRead(My.Application.Info.DirectoryPath.ToString & "\Connections.xml")
            Reader = New XmlTextReader(myStream)
            myStream = Nothing
            While Reader.Read
                Select Case Reader.NodeType
                    Case XmlNodeType.Element
                        Select Case Reader.Name.ToString.Trim
                            Case "ACTIVAR"
                                sHora = Reader.GetAttribute(sProceso).ToString.Trim.Trim
                        End Select
                End Select
            End While
            If DateTime.Now.ToShortTimeString = sHora Then
                ActivaProceso = True
            Else
                ActivaProceso = False
            End If
        Catch ex As Exception
            ActivaProceso = False
        Finally

        End Try
    End Function
    Public Shared Sub Articulos(ByVal sPathLog As String)
        Dim strSepardor = ";"
        Dim strArchivo As String = ""
        Dim strRuta As String = ""
        Dim strTituloCab As String = ""
        Dim strDatosCab As String = ""
        Dim strTituloLineas As String = ""
        Dim strFecha As String = ""
        Dim strSupplier_ili = "2362"
        Dim strReceiver_id = "207972"
        Dim strSQL As String = ""
        Dim dtArticulos As System.Data.DataTable = Nothing
        Dim oDBSAP As SqlConnection = Nothing
        Dim strListaPreciosV As String = ""
        Dim strRutaImagen As String = ""
        Try
            log = New EXO_Log.EXO_Log(sPathLog & "\Logs\ExpArticulos.txt", 1)
            log.escribeMensaje("|OK|" & "Hora inicio creación:  " & CDate(Now), EXO_Log.EXO_Log.Tipo.informacion)

            Conexiones.Connect_SQLServer(oDBSAP)
            log.escribeMensaje("|OK|" & "Despues de conectar", EXO_Log.EXO_Log.Tipo.informacion)

            strRuta = My.Application.Info.DirectoryPath.ToString & "\Ficheros\"
            If System.IO.Directory.Exists(strRuta) Then
            Else
                System.IO.Directory.CreateDirectory(strRuta)
            End If
            strArchivo = "Pricat.csv"


            strFecha = Format(Now.Year, "0000") & Format(Now.Month, "00") & Format(Now.Day, "00")
            'crear las dos primeras lineas de control
            strTituloCab = "HDH" & strSepardor & "DOCUMENT_ID" & strSepardor & "Date" & strSepardor & "CURRENCY" & strSepardor & "SUPPLIER_ILN" & strSepardor & "RECEIVER_ID" & strSepardor & "COUNTRY_ID" & strSepardor & "Ediwheel Version" & strSepardor & "C" & vbCrLf
            strDatosCab = "HDS" & strSepardor & "1" & strFecha & strSepardor & "EUR" & strSepardor & strSupplier_ili & strSepardor & strReceiver_id & strSepardor & "ES" & strSepardor & "B2" & strSepardor & "1" & vbCrLf


            Dim path As String = strRuta & strArchivo

            ' Create or overwrite the file.
            Dim fs As FileStream = File.Create(path)

            ' Add text to the file.
            Dim info As Byte()

            info = New System.Text.UTF8Encoding(True).GetBytes(strTituloCab)
            fs.Write(info, 0, info.Length)
            info = New System.Text.UTF8Encoding(True).GetBytes(strDatosCab)
            fs.Write(info, 0, info.Length)

            strListaPreciosV = Conexiones.GetValueDB(oDBSAP, """@EXO_OGEN1""", "U_EXO_INFV", "U_EXO_NOMV='TarifaRODINEX'")
            log.escribeMensaje("|OK|" & "Despues de lista de precios", EXO_Log.EXO_Log.Tipo.informacion)
            strRutaImagen = Conexiones.GetValueDB(oDBSAP, "OADP", "BitmapPath", "")

            strSQL = "select 'POS' POH, ROW_NUMBER() OVER(ORDER BY T1.ItemCode ASC) POS, t1.U_SEI_JANCODE EAN , t1.CardCode PROD_GRP_1, t1.ItemCode SUPPLIER_CODE, t1.ItemName DESCRIPTION_1,'' DESCRIPTION_2,'' PROD_INFO,'' WDK_ANUB, " _
            & " '' WDK_BRAND,'' WDK_BRAND_TEXT,'' BRAND,REPLACE(U_SEI_CATEGORY1,'Yokohama-Tires','Yokohama') BRAND_TEXT,'' PROD_GRP_2,''GROUP_DESCRIPTION,t1.U_SEI_WEIGHT ""WEIGHT"", T1.U_SEI_INCH RIM_INCH, " _
            & " '' PROD_CYCLE,'' THIRD_PARTY, 'N' PL, 'Y' TELM, 'Y' EDI, 'Y' ADHOC, '' PL_ID,'" & strRutaImagen & "' + COALESCE(t1.PicturName,'')  URL_1,'' URL_2, '' URL_3,'' URL_4,'' URL_5, T1.U_SEI_WIDTH WIDTH_MM, " _
            & " '' WIDTH_INCH, T1.U_SEI_ASPECT ASPECT_RATIO,'' OVL_DIAMETER, U_SEI_RAD_BIAS CONSTRUCTION_1, '' CONSTRUCTION_2,'' USAGE ,'' DEPTH,  " _
            & " T1.U_SEI_INDEXLOAD LI1,'' LI2, '' ""LI3(DWB)"" ,'' ""LI4(DWB)"", T1.U_SEI_SPEEDSY SP1, '' SP2,'' ""TL/TT"", '' FLANK, '' PR, '' RFD, '' SIZE_PREFIX, " _
            & " '' COMM_MARK, '' RIM_MM, T1.U_SEI_RUNFLAT RUN_FLAT, '' SIDEWALL, t1.U_SEI_COMMPATTERN DESIGN_1, '' DESIGN_2, '' PRODUCT_TYPE, '' VEHICLE_TYPE, '' COND_GRP , '' TAX_ID, " _
            & " '' TAX, '' SUGGESTED_PRICE, t2.Price GROSS_PRICE, '' GP_VALID_FROM,'' NET_VALUE, '' NV_VALID, '' RECYCLING_FEE, T1.U_SEI_EXTERNALN NOISE_PERFORMANCE, " _
            & " T1.U_SEI_EXTERNALNO NOISE_CLASS_TYPE, T1.U_SEI_ROLLINREGI ROLLING_RESISTANCE, T1.U_SEI_WETGRIP WET_GRIP, T1.U_SEI_TIRECATEGORY EC_VEHICLE_CLASS, " _
            & " '' EU_DIRECTIVE_NUMBER " _
            & " from oitm t1 " _
            & " LEFT OUTER JOIN ITM1 t2 on t1.ItemCode=t2.ItemCode and t2.PriceList=" & strListaPreciosV & "" _
            & " WHERE ItmsGrpCod='108' " _
            & " AND ISNUMERIC(U_SEI_JANCODE)= 1 " _
            & " AND t1.U_SEICATEGORY2 in ('PCR/VAN','TBS' ) " _
            & " and (t1.U_SEI_CATEGORY1='Yokohama-Tires' or t1.U_SEI_CATEGORY1='Alliance-Tires')" _
            & " AND (ISNULL(t1.U_EXO_Estado, '') = 'A' OR (ISNULL(t1.U_EXO_Estado, '') = 'D' " _
            & " AND ([Yokohama_prod].dbo.EXOStockB2BES(t1.ItemCode) + [Yokohama_prod].dbo.EXOStockB2BPT(t1.ItemCode)) > 0))"
            '& " AND T1.frozenFor ='N' AND T1.validFor='Y'"
            dtArticulos = New System.Data.DataTable("Articulos")
            Conexiones.FillDtDB(oDBSAP, dtArticulos, strSQL)
            If dtArticulos.Rows.Count > 0 Then
                log.escribeMensaje("|OK|" & "Obtencion de articulos", EXO_Log.EXO_Log.Tipo.informacion)
                Dim strColumnas As String = ""
                Dim intClmn As Integer = dtArticulos.Columns.Count
                For i As Integer = 0 To dtArticulos.Columns.Count - 1
                    dtArticulos.Columns(i).ColumnName.ToString()
                    strColumnas = strColumnas & dtArticulos.Columns(i).ColumnName.ToString()

                    If i = intClmn - 1 Then
                        strColumnas = strColumnas & vbCrLf
                    Else
                        strColumnas = strColumnas & strSepardor
                    End If

                Next

                info = New System.Text.UTF8Encoding(True).GetBytes(strColumnas)
                fs.Write(info, 0, info.Length)
                Dim row As DataRow
                Dim strCadena As String = ""
                For Each row In dtArticulos.Rows

                    Dim ir As Integer = 0
                    For ir = 0 To intClmn - 1 Step ir + 1
                        strCadena = strCadena & row(ir).ToString

                        If ir = intClmn - 1 Then
                            strCadena = strCadena & vbCrLf
                        Else
                            strCadena = strCadena & strSepardor
                        End If
                    Next

                    info = New System.Text.UTF8Encoding(True).GetBytes(strCadena)
                    fs.Write(info, 0, info.Length)
                    strCadena = ""
                Next
            End If

            fs.Close()
            fs = Nothing

            'este fichero le tengo que subir al ftp
            log.escribeMensaje("|OK|" & "Antes del upload", EXO_Log.EXO_Log.Tipo.informacion)
            'UploadFTP(path, "ARTICULOS")
            SubirFTP(path, "ARTICULOS", log)
            log.escribeMensaje("|OK|" & "Después del upload", EXO_Log.EXO_Log.Tipo.informacion)

        Catch ex As Exception
            Conexiones.Disconnect_SQLServer(oDBSAP)
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        End Try
        Conexiones.Disconnect_SQLServer(oDBSAP)
        log.escribeMensaje("|OK|" & "Hora fin creación articulos:  " & CDate(Now), EXO_Log.EXO_Log.Tipo.informacion)
    End Sub
    Public Shared Sub Stock(ByVal sPathLog As String)
        Dim strSepardor = ";"
        Dim strArchivo As String = ""
        Dim strRuta As String = ""
        Dim strTituloCab As String = ""
        Dim strDatosCab As String = ""
        Dim strTituloLineas As String = ""
        Dim strFecha As String = ""
        Dim strHora As String = ""
        Dim strSupplier_ili = "2362"
        Dim strReceiver_id = "207972"
        Dim strSQL As String = ""
        Dim dtStock As System.Data.DataTable = Nothing
        Dim oDBSAP As SqlConnection = Nothing

        Try
            log = New EXO_Log.EXO_Log(sPathLog & "\Logs\ExpStock.txt", 1)
            log.escribeMensaje("|OK|" & "Hora inicio creación:  " & CDate(Now), EXO_Log.EXO_Log.Tipo.informacion)

            Conexiones.Connect_SQLServer(oDBSAP)
            log.escribeMensaje("|OK|" & "Despues de conectar", EXO_Log.EXO_Log.Tipo.informacion)

            'strRuta = "C:\Desarrollo\YokohamaGrupoSoledad\Ficheros\"
            strRuta = My.Application.Info.DirectoryPath.ToString & "\Ficheros\"
            If System.IO.Directory.Exists(strRuta) Then
            Else
                System.IO.Directory.CreateDirectory(strRuta)
            End If
            strArchivo = "StockReport_.csv"


            strFecha = Format(Now.Year, "0000") & Format(Now.Month, "00") & Format(Now.Day, "00")
            strHora = Format(Now.Hour, "00") & Format(Now.Minute, "00")



            Dim path As String = strRuta & strArchivo

            ' Create or overwrite the file.
            Dim fs As FileStream = File.Create(path)

            ' Add text to the file.
            Dim info As Byte()


            ' &" sum(OnHand) - sum(IsCommited) ""AVAILABILITY"", " _
            ' &"  sum(OnHand) - sum(IsCommited)  ""QUANTITY VALUE"" " _
            strSQL = "SELECT 'POS' POH, ROW_NUMBER() OVER(ORDER BY ItemCode ASC) ""LINE ID"" , ItemCode ""SUPPLIERS ARTICLE ID"" ,ItemName ""ARTICLE TEXT"", " _
            & " U_SEI_JANCODE ""EANUCC ARTICLE ID"",  " _
            & " Case  " _
            & " When sum(OnHand) - sum(IsCommited)< 0 Then '0' " _
            & " WHEN sum(OnHand) - sum(IsCommited)> 50 THEN '>50' " _
            & " Else  cast(cast(sum(OnHand) - sum(IsCommited) as int) as varchar)   " _
            & " End   ""AVAILABILITY"",  " _
            & " Case   " _
            & " WHEN sum(OnHand) - sum(IsCommited)< 0 THEN '0' " _
            & " WHEN sum(OnHand) - sum(IsCommited)> 50 THEN '>50' " _
            & " Else cast(cast(sum(OnHand) - sum(IsCommited) as int) as varchar)  " _
            & " End   ""QUANTITY VALUE""  " _
            & " from (  " _
            & " Select t1.ItemCode, t1.ItemName,t1.U_SEI_JANCODE, t3.OnHand , t3.IsCommited  " _
            & " from oitm t1  " _
            & " inner join [Yokohama_prod].dbo.[OITW] t3 With (NOLOCK) On t1.ItemCode = t3.ItemCode and t3.WhsCode='01' " _
            & " WHERE ItmsGrpCod='108'  AND ISNUMERIC(U_SEI_JANCODE)=1 and (t1.U_SEI_CATEGORY1='Yokohama-Tires' or t1.U_SEI_CATEGORY1='Alliance-Tires') AND t1.U_SEICATEGORY2 in ('PCR/VAN','TBS' )  " _
            & " And (ISNULL(t1.U_EXO_Estado, '') = 'A' OR (ISNULL(t1.U_EXO_Estado, '') = 'D'  " _
            & " And ([Yokohama_prod].dbo.EXOStockB2BES(t1.ItemCode) + [Yokohama_prod].dbo.EXOStockB2BPT(t1.ItemCode)) > 0))  " _
            & " union all " _
            & " select t1.ItemCode,t1.ItemName, t1.U_SEI_JANCODE, t4.OnHand ,t4.IsCommited  " _
            & " from oitm t1   " _
            & " inner join [Yokohama_PT].dbo.[OITW] t4 WITH (NOLOCK) ON t1.ItemCode = t4.ItemCode and t4.WhsCode='PT01' " _
            & " WHERE ItmsGrpCod='108'  AND ISNUMERIC(U_SEI_JANCODE)=1 and (t1.U_SEI_CATEGORY1='Yokohama-Tires' or t1.U_SEI_CATEGORY1='Alliance-Tires') AND t1.U_SEICATEGORY2 in ('PCR/VAN','TBS' )  " _
            & " And (ISNULL(t1.U_EXO_Estado, '') = 'A' OR (ISNULL(t1.U_EXO_Estado, '') = 'D'  " _
            & " And ([Yokohama_prod].dbo.EXOStockB2BES(t1.ItemCode) + [Yokohama_prod].dbo.EXOStockB2BPT(t1.ItemCode)) > 0)) " _
            & " ) as DatosCompletos " _
            & " group by ItemCode,ItemName,U_SEI_JANCODE " _
            & " order by ItemCode"

            dtStock = New System.Data.DataTable("Articulos")
            Conexiones.FillDtDB(oDBSAP, dtStock, strSQL)
            If dtStock.Rows.Count > 0 Then
                log.escribeMensaje("|OK|" & "Obtencion de stock", EXO_Log.EXO_Log.Tipo.informacion)
                'crear las dos primeras lineas de control
                strTituloCab = "HDH" & strSepardor & "DOCUMENT NUMBER" & strSepardor & "VERSION" & strSepardor & "VARIANT" & strSepardor & "ISSUE DATE" & strSepardor & "ISSUE TIME" & strSepardor & "ERROR CODE" & strSepardor & "TOTAL LINES" & strSepardor & "CUSTOMER REFERENCE" & strSepardor & "BUYER PARTY ID" & strSepardor & "BUYER AGENCY CODE" & strSepardor & "ORDERING PARTY ID" & strSepardor & "ORDERING AGENCY CODE" & strSepardor & "CONSIGNEE PARTY ID" & strSepardor & "CONSIGNEE AGENCY CODE" & vbCrLf
                strDatosCab = "HDS" & strSepardor & "1" & "STOCKREPORT_B2" & strSepardor & "1" & strSepardor & "" & strSepardor & strFecha & strSepardor & strHora & strSepardor & "" & strSepardor & dtStock.Rows.Count & strSepardor & "" & strSepardor & "332845" & strSepardor & "91" & strSepardor & "" & strSepardor & "" & strSepardor & "" & strSepardor & "" & strSepardor & "" & vbCrLf

                info = New System.Text.UTF8Encoding(True).GetBytes(strTituloCab)
                fs.Write(info, 0, info.Length)
                info = New System.Text.UTF8Encoding(True).GetBytes(strDatosCab)
                fs.Write(info, 0, info.Length)

                Dim strColumnas As String = ""
                Dim intClmn As Integer = dtStock.Columns.Count
                For i As Integer = 0 To dtStock.Columns.Count - 1
                    dtStock.Columns(i).ColumnName.ToString()
                    strColumnas = strColumnas & dtStock.Columns(i).ColumnName.ToString()

                    If i = intClmn - 1 Then
                        strColumnas = strColumnas & vbCrLf
                    Else
                        strColumnas = strColumnas & strSepardor
                    End If

                Next

                info = New System.Text.UTF8Encoding(True).GetBytes(strColumnas)
                fs.Write(info, 0, info.Length)
                Dim row As DataRow
                Dim strCadena As String = ""
                For Each row In dtStock.Rows

                    Dim ir As Integer = 0
                    For ir = 0 To intClmn - 1 Step ir + 1
                        strCadena = strCadena & row(ir).ToString

                        If ir = intClmn - 1 Then
                            strCadena = strCadena & vbCrLf
                        Else
                            strCadena = strCadena & strSepardor
                        End If
                    Next

                    info = New System.Text.UTF8Encoding(True).GetBytes(strCadena)
                    fs.Write(info, 0, info.Length)
                    strCadena = ""
                Next
            End If

            fs.Close()
            fs = Nothing

            'este fichero le tengo que subir al ftp
            'SILVIA COMENTO PARA PROBAR
            log.escribeMensaje("|OK|" & "Antes del upload", EXO_Log.EXO_Log.Tipo.informacion)
            'UploadFTP(path, "STOCK")
            SubirFTP(path, "STOCK", log)
            log.escribeMensaje("|OK|" & "Después del upload", EXO_Log.EXO_Log.Tipo.informacion)

        Catch ex As Exception
            Conexiones.Disconnect_SQLServer(oDBSAP)
        End Try
        Conexiones.Disconnect_SQLServer(oDBSAP)
        log.escribeMensaje("|OK|" & "Hora fin creación stock: " & CDate(Now), EXO_Log.EXO_Log.Tipo.informacion)
    End Sub
    Public Shared Sub StockNo(ByVal sPathLog As String)
        Dim strSepardor = ";"
        Dim strArchivo As String = ""
        Dim strRuta As String = ""
        Dim strTituloCab As String = ""
        Dim strDatosCab As String = ""
        Dim strTituloLineas As String = ""
        Dim strFecha As String = ""
        Dim strHora As String = ""
        Dim strSupplier_ili = "2362"
        Dim strReceiver_id = "207972"
        Dim strSQL As String = ""
        Dim dtStock As System.Data.DataTable = Nothing
        Dim oDBSAP As SqlConnection = Nothing

        Try

            Conexiones.Connect_SQLServer(oDBSAP)

            strRuta = "C:\Desarrollo\YokohamaGrupoSoledad\Ficheros\"
            strArchivo = "StockReport_.csv"


            strFecha = Format(Now.Year, "0000") & Format(Now.Month, "00") & Format(Now.Day, "00")
            strHora = Format(Now.Hour, "00") & Format(Now.Minute, "00")



            Dim path As String = strRuta & strArchivo

            ' Create or overwrite the file.
            Dim fs As FileStream = File.Create(path)

            ' Add text to the file.
            Dim info As Byte()



            strSQL = " Select 'POS' POH, ROW_NUMBER() OVER(ORDER BY T1.ItemCode ASC) ""LINE ID"" , t1.ItemCode ""SUPPLIERS ARTICLE ID"" , t2.ItemName ""ARTICLE TEXT"", " _
            & " t2.U_SEI_JANCODE ""EANUCC ARTICLE ID""," _
            & " (sum(t1.OnHand) - sum(t1.IsCommited)) ""AVAILABILITY"" , sum(t1.OnHand) ""QUANTITY VALUE"" " _
            & " From [Yokohama_prod].dbo.OITW t1 " _
            & " INNER Join [Yokohama_prod].dbo.OITM t2 on t1.ItemCode= t2.ItemCode " _
            & "   AND (([Yokohama_prod].dbo.EXOStockB2BES(t1.ItemCode) + [Yokohama_prod].dbo.EXOStockB2BPT(t1.ItemCode)) > 0)" _
            & "  LEFT OUTER JOIN Yokohama_PT.DBO.OITW t3 on t1.ItemCode  = t3.ItemCode " _
            & "  And (([Yokohama_PT].dbo.EXOStockB2BES(t1.ItemCode) + [Yokohama_PT].dbo.EXOStockB2BPT(t1.ItemCode)) > 0) " _
            & " WHERE t2.ItmsGrpCod ='108'  AND ISNUMERIC(t2.U_SEI_JANCODE)=1 AND t2.U_SEICATEGORY2='PCR/VAN'   " _
            & "  AND t2.U_EXO_Estado ='A' " _
            & " group by t1.ItemCode, t2.U_SEI_JANCODE, t2.ItemName "
            ' & " And T2.frozenFor ='N' AND T2.validFor='Y' " _
            dtStock = New System.Data.DataTable("Stock")
            Conexiones.FillDtDB(oDBSAP, dtStock, strSQL)
            If dtStock.Rows.Count > 0 Then

                'crear las dos primeras lineas de control
                strTituloCab = "HDH" & strSepardor & "DOCUMENT NUMBER" & strSepardor & "VERSION" & strSepardor & "VARIANT" & strSepardor & "ISSUE DATE" & strSepardor & "ISSUE TIME" & strSepardor & "ERROR CODE" & strSepardor & "TOTAL LINES" & strSepardor & "CUSTOMER REFERENCE" & strSepardor & "BUYER PARTY ID" & strSepardor & "BUYER AGENCY CODE" & strSepardor & "ORDERING PARTY ID" & strSepardor & "ORDERING AGENCY CODE" & strSepardor & "CONSIGNEE PARTY ID" & strSepardor & "CONSIGNEE AGENCY CODE" & vbCrLf
                strDatosCab = "HDS" & strSepardor & "1" & "STOCKREPORT_B2" & strSepardor & "1" & strSepardor & "" & strSepardor & strFecha & strSepardor & strHora & strSepardor & "" & strSepardor & dtStock.Rows.Count & strSepardor & "" & strSepardor & "332845" & strSepardor & "91" & strSepardor & "" & strSepardor & "" & strSepardor & "" & strSepardor & "" & strSepardor & "" & vbCrLf

                info = New System.Text.UTF8Encoding(True).GetBytes(strTituloCab)
                fs.Write(info, 0, info.Length)
                info = New System.Text.UTF8Encoding(True).GetBytes(strDatosCab)
                fs.Write(info, 0, info.Length)

                Dim strColumnas As String = ""
                Dim intClmn As Integer = dtStock.Columns.Count
                For i As Integer = 0 To dtStock.Columns.Count - 1
                    dtStock.Columns(i).ColumnName.ToString()
                    strColumnas = strColumnas & dtStock.Columns(i).ColumnName.ToString()

                    If i = intClmn - 1 Then
                        strColumnas = strColumnas & vbCrLf
                    Else
                        strColumnas = strColumnas & strSepardor
                    End If

                Next

                info = New System.Text.UTF8Encoding(True).GetBytes(strColumnas)
                fs.Write(info, 0, info.Length)
                Dim row As DataRow
                Dim strCadena As String = ""
                For Each row In dtStock.Rows

                    Dim ir As Integer = 0
                    For ir = 0 To intClmn - 1 Step ir + 1
                        strCadena = strCadena & row(ir).ToString

                        If ir = intClmn - 1 Then
                            strCadena = strCadena & vbCrLf
                        Else
                            strCadena = strCadena & strSepardor
                        End If
                    Next

                    info = New System.Text.UTF8Encoding(True).GetBytes(strCadena)
                    fs.Write(info, 0, info.Length)
                    strCadena = ""
                Next
            End If

            fs.Close()
            fs = Nothing

            'este fichero le tengo que subir al ftp
            log.escribeMensaje("|OK|" & "Antes del upload", EXO_Log.EXO_Log.Tipo.informacion)
            UploadFTP(path, "STOCK")
            log.escribeMensaje("|OK|" & "Después del upload", EXO_Log.EXO_Log.Tipo.informacion)
        Catch ex As Exception
            Conexiones.Disconnect_SQLServer(oDBSAP)
        End Try
        Conexiones.Disconnect_SQLServer(oDBSAP)

    End Sub

#Region "FTP"
    Shared Sub UploadFTP(ByVal strFileNameLocal As String, strTipo As String)
        Dim strNomCorto As String = Path.GetFileName(strFileNameLocal)
        Dim dirFTP As String = Conexiones.Datos_FTP("FTP", "URL")
        Dim strUser As String = Conexiones.Datos_FTP("FTP", "Usuario")
        Dim strPass As String = Conexiones.Datos_FTP("FTP", "Password")

        Try
            Select Case strTipo
                Case "ARTICULOS"
                    dirFTP = dirFTP & "catalogo/"
                    'dirFTP = dirFTP & "stock/"
                Case "STOCK"
                    dirFTP = dirFTP & "stock/"
            End Select
            My.Computer.Network.UploadFile(strFileNameLocal, dirFTP & strNomCorto, strUser, strPass, True, 100)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Shared Sub SubirFTP(ByVal strFileNameLocal As String, strTipo As String, ByRef oLog As EXO_Log.EXO_Log)
        Dim strNomCorto As String = Path.GetFileName(strFileNameLocal)
        Dim sURLFTP As String = Conexiones.Datos_FTP("FTP", "URL")
        'Dim sPTFTP As String = Conexiones.Datos_Confi("FTP", "PUERTO")
        Dim dirFTP As String = ""
        Select Case strTipo
            Case "ARTICULOS"
                dirFTP = "/ftp/Catalogo/"
            Case "STOCK"
                dirFTP = "/ftp/Stock/"
        End Select
        Dim sUser As String = Conexiones.Datos_FTP("FTP", "Usuario")
        Dim sPass As String = Conexiones.Datos_FTP("FTP", "Password")


        Try
            Dim sessionOptions As New SessionOptions
            With sessionOptions
                .Protocol = Protocol.Sftp
                .HostName = sURLFTP
                .UserName = sUser
                .Password = sPass
                .SshHostKeyFingerprint = "ssh-ed25519 255 dbyAYPE5ICNcDQQyH5CYnF0tfBu7UgtcaH/mx5gm6Hs="
                .AddRawSettings("FSProtocol", "2")
            End With

            Using session As New Session
                ' Connect
                session.Open(sessionOptions)

                ' Upload files
                Dim transferOptions As New TransferOptions
                transferOptions.TransferMode = TransferMode.Binary

                Dim transferResult As TransferOperationResult
                transferResult = session.PutFiles(strFileNameLocal, dirFTP, False, transferOptions)

                ' Throw on any error
                transferResult.Check()

                ' Print results
                For Each transfer In transferResult.Transfers
                    oLog.escribeMensaje("Fichero subido..." & transfer.FileName, EXO_Log.EXO_Log.Tipo.informacion)
                Next
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
End Class