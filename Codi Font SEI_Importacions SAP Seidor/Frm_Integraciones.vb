Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Text
Imports System.IO
Imports System.Xml
Imports System.Net
Imports SEI_Importacions.SEI_AddOnEnum
Imports SAPbobsCOM.BoObjectTypes

Public Class Frm_Integraciones

#Region "Variables"

    Private sBD_Name As String
    Private hPath As New Hashtable

    Private bError As Boolean

#End Region

#Region "Conexiones"

    Private Function ConectarSBO(ByVal sPos As String) As Boolean

        Dim lRetCode As Long
        Dim lErrCode As Long
        Dim sErrMsg As String

        ConectarSBO = False

        Try
            lRetCode = 0
            sErrMsg = ""

            oCompany = New SAPbobsCOM.Company

            oCompany.LicenseServer = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "LS")
            oCompany.Server = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "S")
            oCompany.CompanyDB = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "BD")
            oCompany.DbUserName = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "U")
            oCompany.DbPassword = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "P")
            oCompany.UserName = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "USBO")
            oCompany.Password = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "PSBO")
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish

            Select Case IniGet(Application.StartupPath & sFicheroIni, "Parametros", "DBS")
                Case "MSSQL2005" : oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
                Case "MSSQL2008" : oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
                Case "MSSQL2012" : oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
                Case "MSSQL2014" : oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
                    'Case "MSSQL2016" : oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016
            End Select

            lRetCode = oCompany.Connect

            If lRetCode <> 0 Then
                Me.Msg_Conectar("Failed to connect.")
                oCompany.GetLastError(lErrCode, sErrMsg)
                Throw New Exception(sErrMsg)
            Else
                ConectarSBO = True
                SEI_Globals.Log(vbTab & sPos & ". Connection", "PROCESS: Made")
                Me.Msg_Conectar("DB: " & oCompany.CompanyName)
            End If

        Catch ex As Exception
            SEI_Globals.Log(vbTab & sPos & ". Connection", "ERROR: " & ex.Message)
            Throw New Exception("Error ConectarSBO: " & ex.Message)
        End Try

    End Function

#End Region

#Region "Funciones Integración"

    Private Sub ImportarArticulos()
        'Primer importem els fitxers que comencen per A-, que són els artícles que hem de modificar o crear.
        'Despres els que comencen per C-, que són els de compra.
        'I per últim els que comencen per V-, que són els de venta.

        Dim oItem As SAPbobsCOM.Items
        Dim sFolderPath As String = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "RUTAALBCOMPRAS")
        Dim sItemCode As String = ""
        Dim sCurrentLine As String
        Dim FileToMove As String
        Dim FileName As String
        Dim MoveLocation As String = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "RUTARESARTICLECOMPRA")
        Dim Dir As New DirectoryInfo(sFolderPath)
        Dim fiArr As FileInfo() = Dir.GetFiles()
        Dim fri As FileInfo
        '
        Me.Msg_Mensaje("Importando Articulos " & oCompany.CompanyName & "...")
        Try

            For Each fri In fiArr
                '
                If fri.Name.Substring(0, 1) = "A" Then
                    Dim oReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(fri.FullName)
                    oReader.TextFieldType = FileIO.FieldType.Delimited
                    oReader.Delimiters = New String() {";"}
                    '
                    Do
                        oItem = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                        sCurrentLine = oReader.ReadLine
                        sItemCode = sCurrentLine.Split(";")(0)
                        If oItem.GetByKey(sItemCode) Then 'Si existeix, actualitzem
                            oItem.ItemName = sCurrentLine.Split(";")(1)
                            If sCurrentLine.Split(";")(2) <> "" Then 'Si te idfamilia serà mateiral, sinó serà tinta
                                oItem.UserFields.Fields.Item("U_GSP_INFAMMATCOD").Value = sCurrentLine.Split(";")(2)
                                oItem.UserFields.Fields.Item("U_GSP_INSUBFAMMATCOD").Value = sCurrentLine.Split(";")(4)
                                oItem.UserFields.Fields.Item("U_GSP_INFAMCOLCOD").Value = sCurrentLine.Split(";")(6)
                                oItem.UserFields.Fields.Item("U_GSP_INSUBFAMCOLCOD").Value = sCurrentLine.Split(";")(8)
                                oItem.UserFields.Fields.Item("U_GSP_INFAMADICOD").Value = sCurrentLine.Split(";")(10)
                                oItem.UserFields.Fields.Item("U_GSP_INSUBFAMADICOD").Value = sCurrentLine.Split(";")(12)

                            Else 'És tinta

                                oItem.UserFields.Fields.Item("U_GSP_INPREFCOLOR").Value = sCurrentLine.Split(";")(14)
                                oItem.UserFields.Fields.Item("U_GSP_INFCOLORCOD").Value = sCurrentLine.Split(";")(15)
                                oItem.UserFields.Fields.Item("U_GSP_INSFCOLORCOD").Value = sCurrentLine.Split(";")(17)
                                oItem.UserFields.Fields.Item("U_GSP_INFTINTACOD").Value = sCurrentLine.Split(";")(19)
                                oItem.UserFields.Fields.Item("U_GSP_INSFTINTACOD").Value = sCurrentLine.Split(";")(21)

                            End If
                            '
                            If oItem.Update() <> 0 Then
                                SEI_Globals.Log("1.UPDATE '" & Me.sBD_Name & "'", "Error actualizando el artícluo: " & sItemCode & ". Causa: " & oCompany.GetLastErrorDescription)
                                b_Errors = True
                            End If

                        Else 'Si no existeix, el creem
                            oItem.ItemCode = sItemCode
                            oItem.ItemName = sCurrentLine.Split(";")(1)
                            oItem.InventoryItem = False
                            If sCurrentLine.Split(";")(2) <> "" Then 'Si te idfamilia serà mateiral, sinó serà tinta
                                oItem.UserFields.Fields.Item("U_GSP_INFAMMATCOD").Value = sCurrentLine.Split(";")(2)
                                oItem.UserFields.Fields.Item("U_GSP_INSUBFAMMATCOD").Value = sCurrentLine.Split(";")(4)
                                oItem.UserFields.Fields.Item("U_GSP_INFAMCOLCOD").Value = sCurrentLine.Split(";")(6)
                                oItem.UserFields.Fields.Item("U_GSP_INSUBFAMCOLCOD").Value = sCurrentLine.Split(";")(8)
                                oItem.UserFields.Fields.Item("U_GSP_INFAMADICOD").Value = sCurrentLine.Split(";")(10)
                                oItem.UserFields.Fields.Item("U_GSP_INSUBFAMADICOD").Value = sCurrentLine.Split(";")(12)

                            Else 'És tinta

                                oItem.UserFields.Fields.Item("U_GSP_INPREFCOLOR").Value = sCurrentLine.Split(";")(14)
                                oItem.UserFields.Fields.Item("U_GSP_INFCOLORCOD").Value = sCurrentLine.Split(";")(15)
                                oItem.UserFields.Fields.Item("U_GSP_INSFCOLORCOD").Value = sCurrentLine.Split(";")(17)
                                oItem.UserFields.Fields.Item("U_GSP_INFTINTACOD").Value = sCurrentLine.Split(";")(19)
                                oItem.UserFields.Fields.Item("U_GSP_INSFTINTACOD").Value = sCurrentLine.Split(";")(21)
                            End If
                            '
                            If oItem.Add() <> 0 Then
                                SEI_Globals.Log("1.ADD '" & Me.sBD_Name & "'", "Error añadiendo el artícluo: " & sItemCode & ". Causa: " & oCompany.GetLastErrorDescription)
                                b_Errors = True
                            End If
                        End If
                        '
                    Loop While Not oReader.EndOfData
                    oReader.Close()

                    'Un cop processat, movem l'arxiu per no processar-lo de nou

                    FileToMove = sFolderPath & "\" & fri.Name
                    FileName = fri.Name
                    If MoveLocation.Substring(MoveLocation.Length - 2, 1) <> "\" Then
                        MoveLocation &= "\"
                    End If
                    If System.IO.File.Exists(FileToMove) = True Then
                        System.IO.File.Move(FileToMove, MoveLocation & fri.Name.Replace(fri.Extension, "") & "_" & Date.Now.ToString("dd-MM-yyyy_hh-mm-ss") & fri.Extension)
                    End If

                End If

            Next fri
            '
        Catch ex As Exception
            SEI_Globals.Log("1.ERR. '" & Me.sBD_Name & "'", "Error preparando el artícluo: " & sItemCode & ". Causa: " & ex.Message)
            b_Errors = True
        End Try
    End Sub
    '
    Private Sub ImportarAlbaraEntradaMercaderies()
        'Primer importem els fitxers que comencen per A-, que són els artícles que hem de modificar o crear.
        'Despres els que comencen per C-, que són el de compra.
        'I per últim els que comencen per V-, que són els de venta.

        Dim oAlbara As SAPbobsCOM.Documents
        Dim sFolderPath As String = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "RUTAALBCOMPRAS")
        Dim sItemCode As String
        Dim sCurrentLine As String
        Dim FileToMove As String
        Dim MoveLocation As String = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "RUTARESALBCOMPRAS")
        Dim Dir As New DirectoryInfo(sFolderPath)
        Dim fiArr As FileInfo() = Dir.GetFiles()
        Dim fri As FileInfo
        Dim iCont As Integer = 0
        Dim sPathPDF As String = ""
        Dim sReport As String = ""
        '
        Me.Msg_Mensaje("Importando Compras " & oCompany.CompanyName & "...")

        Try

            For Each fri In fiArr
                '
                If fri.Name.Substring(0, 1) = "C" Then
                    Dim oReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(fri.FullName)
                    oReader.TextFieldType = FileIO.FieldType.Delimited
                    oReader.Delimiters = New String() {";"}
                    iCont = 0
                    '
                    oAlbara = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                    Do
                        sCurrentLine = oReader.ReadLine
                        If iCont = 0 Then 'Si és la primera linia, omplim les dades de capçalera, la resta són el detall
                            oAlbara.CardCode = sCurrentLine.Split(";")(0)
                            oAlbara.DocDate = CDate(sCurrentLine.Split(";")(1).ToString)
                            oAlbara.NumAtCard = sCurrentLine.Split(";")(2)
                        Else
                            '
                            If oAlbara.Lines.ItemCode <> "" Then
                                oAlbara.Lines.Add()
                            End If

                            sItemCode = sCurrentLine.Split(";")(0)
                            oAlbara.Lines.ItemCode = sItemCode
                            oAlbara.Lines.ItemDescription = sCurrentLine.Split(";")(1)
                            oAlbara.Lines.UserFields.Fields.Item("U_GSP_INSEMIELABOR").Value = sCurrentLine.Split(";")(2)                                                                    'Semielaborat           (Text 10)
                            If sCurrentLine.Split(";")(3).ToString <> "" Then oAlbara.Lines.UserFields.Fields.Item("U_GSP_INANCHURA").Value = CDbl(sCurrentLine.Split(";")(3).ToString)      'Amplada                (Numeric)
                            If sCurrentLine.Split(";")(4).ToString <> "" Then oAlbara.Lines.UserFields.Fields.Item("U_GSP_INESPESOR").Value = CDbl(sCurrentLine.Split(";")(4).ToString)      'Espesor                (Numeric)
                            oAlbara.Lines.UserFields.Fields.Item("U_GSP_INUDESPESOR").Value = sCurrentLine.Split(";")(5)                                                                     'UnitatEspesor          (Text 3: "Micres"; "Galgues"; "Grms/m2")
                            oAlbara.Lines.UserFields.Fields.Item("U_GSP_INFABLOTE").Value = sCurrentLine.Split(";")(6)                                                                       'LotFabricació          (Text 20)
                            oAlbara.Lines.UserFields.Fields.Item("U_GSP_INPEDIDOPROV").Value = sCurrentLine.Split(";")(7)                                                                    'ComandaProveidor       (Numeric)     
                            oAlbara.Lines.UserFields.Fields.Item("U_GSP_INPEDREL").Value = sCurrentLine.Split(";")(8)                                                                        'ComandesRelacionades   (Text 50)
                            oAlbara.Lines.Quantity = CDbl(Replace(sCurrentLine.Split(";")(9).ToString, ".", ","))
                            oAlbara.Lines.UnitPrice = CDbl(Replace(sCurrentLine.Split(";")(10).ToString, ".", ","))

                        End If
                        '
                        iCont += 1
                        '
                    Loop While Not oReader.EndOfData
                    '
                    If oAlbara.Add() <> 0 Then
                        SEI_Globals.Log("1.ADD '" & Me.sBD_Name & "'", "Error creando la Entrada de Mercancías " & fri.Name.ToString & ". Causa: " & oCompany.GetLastErrorDescription)
                        b_Errors = True
                    Else
                        'Un cop processat i creat , movem l'arxiu per no processar-lo de nou

                        oReader.Close()
                        FileToMove = sFolderPath & "\" & fri.Name
                        If MoveLocation.Substring(MoveLocation.Length - 2, 1) <> "\" Then
                            MoveLocation &= "\"
                        End If
                        If System.IO.File.Exists(FileToMove) = True Then
                            System.IO.File.Move(FileToMove, MoveLocation & fri.Name.Replace(fri.Extension, "") & "_" & Date.Now.ToString("dd-MM-yyyy_hh-mm-ss") & fri.Extension)
                        End If
                        'Creem el PDF i imprimim el document
                        Dim sDocNum As String = GetDocNum("OPDN", oCompany.GetNewObjectKey)
                        sPathPDF = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "PDFCOMPRAS") & "\" & "PedidoCompra_" & sDocNum & ".pdf"
                        sReport = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "INFORMECOMPRAS")
                        CrearPDF(sReport, oCompany.GetNewObjectKey, oPurchaseDeliveryNotes, sPathPDF, GetLangIC(oAlbara.CardCode))
                        '
                    End If
                    '
                End If
            Next fri
            '
        Catch ex As Exception
            SEI_Globals.Log("1.ERR. '" & Me.sBD_Name & "'", "Error preparando la Entrada de Mercancías " & fri.Name.ToString & ". Causa: " & ex.Message)
            b_Errors = True
        End Try
    End Sub
    '
    Private Function IsCardCodeCROPS(ByVal sCardCode As String) As Boolean
        Dim ors As SAPbobsCOM.Recordset
        Dim sSQL As String
        '
        ors = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        sSQL = "SELECT CardCode FROM OCRD WHERE QryGroup2 = 'Y' and CardCode = '" & sCardCode & "'"
        ors.DoQuery(sSQL)
        '
        If ors.RecordCount > 0 Then
            Return True
        End If
        '
        Return False
    End Function
    '
    Private Function GetDocNum(ByVal sTaula As String, ByVal sDocEntry As String) As String
        Dim ors As SAPbobsCOM.Recordset
        Dim sSQL As String
        ors = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        sSQL = "SELECT DocNum FROM " & sTaula & " WHERE DocEntry = '" & sDocEntry & "'"
        ors.DoQuery(sSQL)
        Return ors.Fields.Item("DocNum").Value
    End Function
    '
    Private Function GetLangIC(ByVal sCardCode As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim sSQL As String
        '
        oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '
        sSQL = "SELECT t1.ShortName FROM OCRD t0 inner join olng t1 on t0.langcode = t1.Code WHERE t0.CardCode = '" & sCardCode & "'"
        oRS.DoQuery(sSQL)
        '
        Return oRS.Fields.Item("ShortName").Value
    End Function
    '
    Public Sub CrearPDF(ByVal sInforme As String,
                        ByVal lDocEntry As Long,
                        ByVal iTipoDoc As SAPbobsCOM.BoObjectTypes,
                        ByRef DocPDFpath As String,
                        ByVal sLang As String,
                        Optional ByVal iCopias As Integer = 1)
        '
        Dim stEmail As New st_Email
        Dim stParametro As st_Parametro
        Dim stSubReportWhere As New st_SubReportWhere
        Dim aFormulas As ArrayList = Nothing
        '
        Dim stFormula As st_Formulas = Nothing
        Dim sLiteral As String = ""
        ' 
        ' Añadir parametros a un informe
        stSubReportWhere.aParametros = New ArrayList

        stParametro = New st_Parametro
        stParametro.FieldName = "DocKey@"
        stParametro.Value = lDocEntry.ToString
        stSubReportWhere.aParametros.Add(stParametro)
        '
        Dim sPrinter As String = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "IMPRESORA")

        SEI_Report.ImprimirPDF_EMAIL(eCrystal.EnDirectorio, sInforme, aFormulas, stSubReportWhere, True, DocPDFpath, sPrinter, iCopias)
        '
    End Sub
    '
    Private Sub Modificar_Crear_Article(ByVal sItemCode As String, ByRef sCurrentLine As String, ByVal sFileName As String, ByVal sLineNum As String)
        Dim oItem As SAPbobsCOM.Items

        oItem = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '
        If oItem.GetByKey(sItemCode) Then
            If sCurrentLine.Split(";")(0).Contains("#" & sLineNum & ".1") Then
                oItem.ItemName = sCurrentLine.Split(";")(3)
            End If

            If oItem.Update() <> 0 Then
                SEI_Globals.Log("1.ITEM_UPDATE '" & Me.sBD_Name & "'", "Error actualizando el Artículo " & sItemCode & " para la Entrega " & sFileName & ". Causa: " & oCompany.GetLastErrorDescription)
                b_Errors = True
            End If

        Else
            If sCurrentLine.Split(";")(0).Contains("#" & sLineNum & ".1") Then
                oItem.ItemCode = sItemCode
                oItem.ItemName = sCurrentLine.Split(";")(3)
                oItem.InventoryItem = False
            End If

            If oItem.Add() <> 0 Then
                SEI_Globals.Log("1.ITEM_ADD '" & Me.sBD_Name & "'", "Error creando el Artículo " & sItemCode & " para la Entrega " & sFileName & ". Causa: " & oCompany.GetLastErrorDescription)
                b_Errors = True
            End If

        End If
        '
    End Sub
    '
    Private Function ExistsOrCreate_ShipToCode(ByVal sShipToCode As String, ByVal sCardCode As String, ByVal sCurrentLine As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim sSQL As String = ""
        Dim sShipToDef As String = ""
        Dim oIC As SAPbobsCOM.BusinessPartners
        oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oIC = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        '
        sSQL = "SELECT Address FROM CRD1 WHERE CardCode = '" & sCardCode & "' and AdresType = 'S' and Address = '" & Replace(sShipToCode, "´", "''") & "'"
        '
        oRS.DoQuery(sSQL)
        If oRS.RecordCount > 0 Then
            Return oRS.Fields.Item("Address").Value
        Else
            If oIC.GetByKey(sCardCode) Then
                sShipToDef = oIC.ShipToDefault
                oIC.Addresses.Add()
                oIC.Addresses.AddressName = Replace(sShipToCode, "´", "'")
                oIC.Addresses.AddressName2 = sCurrentLine.Split(";")(9)
                oIC.Addresses.Street = sCurrentLine.Split(";")(10)
                oIC.Addresses.ZipCode = sCurrentLine.Split(";")(11)
                oIC.Addresses.City = sCurrentLine.Split(";")(12)
                oIC.Addresses.County = sCurrentLine.Split(";")(13)
                oIC.Addresses.Country = sCurrentLine.Split(";")(15)
                oIC.ShipToDefault = sShipToDef
                If oIC.Update() <> 0 Then
                    SEI_Globals.Log("1.UPDATE ShipTo '" & Me.sBD_Name & "'", "Error creando dirección de entrega " & sShipToCode & " en el cliente: " & sCardCode & ". Causa: " & oCompany.GetLastErrorDescription)
                    b_Errors = True
                End If
                Return sShipToCode
                End If
                Return ""
        End If
        '
        LiberarObjCOM(oRS)
        LiberarObjCOM(oIC)
    End Function
    '
    Private Sub ImportarVentes()
        'Primer importem els fitxers que comencen per A-, que són els artícles que hem de modificar o crear.
        'Despres els que comencen per C-, que són el de compra.
        'I per últim els que comencen per V-, que són els de venta.

        Dim oAlbara As SAPbobsCOM.Documents
        Dim oItem As SAPbobsCOM.Items
        Dim sFolderPath As String = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "RUTAALBVENTAS")
        Dim sItemCode As String
        Dim sCurrentLine As String
        Dim FileToMove As String
        Dim MoveLocation As String = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "RUTARESALBVENTAS")
        Dim Dir As New DirectoryInfo(sFolderPath)
        Dim fiArr As FileInfo() = Dir.GetFiles()
        Dim fri As FileInfo
        Dim iCont As Integer = 0
        Dim sPathPDF As String = ""
        Dim sReport As String = ""
        '
        Me.Msg_Mensaje("Importando Ventas " & oCompany.CompanyName & "...")
        Try

            For Each fri In fiArr
                '
                If fri.Name.Substring(0, 1) = "V" Then
                    Dim oReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(fri.FullName, Encoding.Default)
                    oReader.TextFieldType = FileIO.FieldType.Delimited
                    oReader.Delimiters = New String() {";"}
                    iCont = 0
                    '
                    oAlbara = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    Do
                        sCurrentLine = oReader.ReadLine
                        If iCont = 0 Then 'Si és la primera linia, omplim les dades de capçalera, la resta són el detall
                            oAlbara.CardCode = sCurrentLine.Split(";")(3)
                            oAlbara.DocDate = CDate(sCurrentLine.Split(";")(2).ToString)
                            oAlbara.NumAtCard = sCurrentLine.Split(";")(1)
                            If sCurrentLine.Split(";")(5).ToString <> "" Then oAlbara.TransportationCode = CInt(sCurrentLine.Split(";")(5).ToString)
                            oAlbara.UserFields.Fields.Item("U_GSP_INTPPORTE").Value = sCurrentLine.Split(";")(6)
                            oAlbara.UserFields.Fields.Item("U_GSP_INOBSERVPORT").Value = sCurrentLine.Split(";")(7)
                            oAlbara.UserFields.Fields.Item("U_GSP_INOBSERVCALB").Value = sCurrentLine.Split(";")(8)
                            oAlbara.UserFields.Fields.Item("U_GSP_INVALORADO").Value = sCurrentLine.Split(";")(14)
                            oAlbara.LanguageCode = RecuperarIdioma(sCurrentLine.Split(";")(16))
                            oAlbara.ShipToCode = ExistsOrCreate_ShipToCode(sCurrentLine.Split(";")(9) & "_" & sCurrentLine.Split(";")(4), oAlbara.CardCode, sCurrentLine) 'TODO: Comprovar si existeix i si no, crear-la????
                            '

                        Else
                            '
                            sItemCode = sCurrentLine.Split(";")(1)
                            '
                            '
                            If sCurrentLine.Split(";")(0).ToString.Substring(0, 1) = "#" And sCurrentLine.Split(";")(0).ToString.Contains(".1") Then
                                'Linia 1
                                If oAlbara.Lines.ItemCode <> "" Then
                                    oAlbara.Lines.Add()
                                End If
                                Modificar_Crear_Article(sItemCode, sCurrentLine, fri.FullName, sCurrentLine(1))
                                oAlbara.Lines.ItemCode = sItemCode
                                oAlbara.Lines.UserFields.Fields.Item("U_GSP_INUDMEDIDA").Value = sCurrentLine.Split(";")(2) '.Replace("?", "€") 'Encoding.ASCII.GetString(Encoding.GetEncoding("Unicode").GetBytes(sCurrentLine.Split(";")(2)))                 'Unitat mesura           (Text 10)
                                oAlbara.Lines.Quantity = CDbl(sCurrentLine.Split(";")(4).ToString.Replace(".", ","))
                                oAlbara.Lines.UnitPrice = CDbl(sCurrentLine.Split(";")(5).ToString.Replace(".", ","))

                            ElseIf sCurrentLine.Split(";")(0).ToString.Substring(0, 1) = "#" And sCurrentLine.Split(";")(0).ToString.Contains(".2") Then
                                'Linia 2
                                If sCurrentLine.Split(";")(1).ToString <> "" Then oAlbara.Lines.UserFields.Fields.Item("U_GSP_INANCHURA").Value = CDbl(sCurrentLine.Split(";")(1).ToString)      'Amplada                (Double)
                                If sCurrentLine.Split(";")(2).ToString <> "" Then oAlbara.Lines.UserFields.Fields.Item("U_GSP_INESPESOR").Value = CDbl(sCurrentLine.Split(";")(2).ToString)      'Espesor                (Double)
                                If sCurrentLine.Split(";")(3) <> "" Then oAlbara.Lines.UserFields.Fields.Item("U_GSP_INUDESPESOR").Value = GetValueUNIDADESPESOR(sCurrentLine.Split(";")(3))                 'MesuraEspesor          (Text 8)
                                oAlbara.Lines.UserFields.Fields.Item("U_GSP_INDESCMED").Value = sCurrentLine.Split(";")(4)                     'DescripcioMides        (Text 50)

                            ElseIf sCurrentLine.Split(";")(0).ToString.Substring(0, 1) = "#" And sCurrentLine.Split(";")(0).ToString.Contains(".3") Then
                                'Linia 3
                                oAlbara.Lines.UserFields.Fields.Item("U_GSP_INMARCAYLIN").Value = sCurrentLine.Split(";")(1)                   'Marcailinia            (Text 80)

                            ElseIf sCurrentLine.Split(";")(0).ToString.Substring(0, 1) = "#" And sCurrentLine.Split(";")(0).ToString.Contains(".4") Then
                                'Linia 4
                                oAlbara.Lines.UserFields.Fields.Item("U_GSP_INNUMPEDCLI").Value = sCurrentLine.Split(";")(1)                   'ComandaClient          (Text 20)     
                                oAlbara.Lines.UserFields.Fields.Item("U_GSP_INREFCLI").Value = sCurrentLine.Split(";")(2)                      'Refclient              (Text 20)

                            ElseIf sCurrentLine.Split(";")(0).ToString.Substring(0, 1) = "#" And sCurrentLine.Split(";")(0).ToString.Contains(".5") Then
                                'Linia 5
                                'NumComandaCliDeClient  (Text 20)
                                'RefClientdeClient      (Text 20)
                                If sCurrentLine.Split(";")(3).ToString <> "" Then oAlbara.Lines.UserFields.Fields.Item("U_GSP_INFABDATE").Value = CDate(sCurrentLine.Split(";")(3).ToString)     'DataFabricació         (Date)
                                oAlbara.Lines.UserFields.Fields.Item("U_GSP_INBARCODE").Value = sCurrentLine.Split(";")(4)                     'CodiBarres             (Text 15)

                            ElseIf sCurrentLine.Split(";")(0).ToString.Substring(0, 1) = "#" And sCurrentLine.Split(";")(0).ToString.Contains(".6") Then
                                'Linia 6
                                oAlbara.Lines.UserFields.Fields.Item("U_GSP_INCONTRATO").Value = sCurrentLine.Split(";")(1)                    'NumContracte           (Text 15)
                                oAlbara.Lines.UserFields.Fields.Item("U_GSP_INCALLOFF").Value = sCurrentLine.Split(";")(2)                     'NumCallOff             (Text 15)

                            ElseIf sCurrentLine.Split(";")(0).ToString.Substring(0, 1) = "#" And sCurrentLine.Split(";")(0).ToString.Contains(".7") Then
                                'Linia 7
                                If sCurrentLine.Split(";")(1).ToString <> "" Then oAlbara.Lines.UserFields.Fields.Item("U_GSP_INBOBS").Value = CInt(sCurrentLine.Split(";")(1).ToString)                            'Numbobs                (Long)
                                If sCurrentLine.Split(";")(2).ToString <> "" Then oAlbara.Lines.UserFields.Fields.Item("U_GSP_INKGBRUTO").Value = CInt(Replace(sCurrentLine.Split(";")(2).ToString, ".", ","))      'KgTotalsBruts          (Long)
                                If sCurrentLine.Split(";")(3).ToString <> "" Then oAlbara.Lines.UserFields.Fields.Item("U_GSP_INMTRSLIN").Value = CInt(Replace(sCurrentLine.Split(";")(3).ToString, ".", ","))      'MetresLineals          (Long)
                                If sCurrentLine.Split(";")(4).ToString <> "" Then oAlbara.Lines.UserFields.Fields.Item("U_GSP_INUNIDADES").Value = CDbl(sCurrentLine.Split(";")(4).ToString)                        'Unitats                (Double)
                                oAlbara.Lines.UserFields.Fields.Item("U_GSP_INFABLOTE").Value = sCurrentLine.Split(";")(5).ToString                               'LotInplacsa            (Long)
                                oAlbara.Lines.UserFields.Fields.Item("U_GSP_INTPFORMATO").Value = sCurrentLine.Split(";")(6)                                      'TipusProducte          (Text 6)

                            ElseIf sCurrentLine.Split(";")(0).ToString.Substring(0, 2) = "LV" Then
                                '->Linies de Text
                                'LV1,LV2,LV3,...
                                'Treiem la primera linia perquè sino surt l'item code de l'entrega i la primera linia de text també informa de l'itemcode i surt duplicada
                                If sCurrentLine.Split(";")(0).ToString.Substring(sCurrentLine.Split(";")(0).ToString.IndexOf(".") + 1, 1) <> "1" Then

                                    If oAlbara.SpecialLines.LineText <> "" Then
                                        oAlbara.SpecialLines.Add()
                                    End If
                                    oAlbara.SpecialLines.LineType = SAPbobsCOM.BoDocSpecialLineType.dslt_Text
                                    oAlbara.SpecialLines.LineText = sCurrentLine.Split(";")(1)
                                    oAlbara.SpecialLines.AfterLineNumber = sCurrentLine.Split(";")(0).ToString.Substring(2, sCurrentLine.Split(";")(0).ToString.IndexOf(".") - 2) - 1
                                End If
                            ElseIf sCurrentLine.Split(";")(0).ToString.Substring(0, 2) = "PA" Then 'Linies de comentaris de camps usuari capçalera (PA1,PA2...)
                                'Comentaris d'usuari capçalera
                                oAlbara.UserFields.Fields.Item("U_GSP_INFOOTER" & sCurrentLine.Split(";")(0).ToString.Substring(2, 1)).Value = sCurrentLine.Split(";")(1) 'PAx;
                            End If

                        End If
                        '
                        iCont += 1
                        '
                    Loop While Not oReader.EndOfData
                    '
                    If oAlbara.Add() <> 0 Then
                        SEI_Globals.Log("1.ADD '" & Me.sBD_Name & "'", "Error creando el Pedido de Ventas " & fri.Name.ToString & ". Causa: " & oCompany.GetLastErrorDescription)
                        b_Errors = True
                    Else
                        oReader.Close()
                        'Un cop processat i creat , movem l'arxiu per no processar-lo de nou

                        FileToMove = sFolderPath & "\" & fri.Name
                        If MoveLocation.Substring(MoveLocation.Length - 2, 1) <> "\" Then
                            MoveLocation &= "\"
                        End If
                        If System.IO.File.Exists(FileToMove) = True Then
                            System.IO.File.Move(FileToMove, MoveLocation & fri.Name.Replace(fri.Extension, "") & "_" & Date.Now.ToString("dd-MM-yyyy_hh-mm-ss") & fri.Extension)
                        End If
                        '
                        'Creem el PDF i imprimim el document
                        Dim sDocNum As String = GetDocNum("ODLN", oCompany.GetNewObjectKey)
                        sPathPDF = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "PDFVENTAS") & "\" & "EntregaVentas_" & sDocNum & ".pdf"
                        '
                        If IsCardCodeCROPS(oAlbara.CardCode) Then
                            sReport = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "INFORMEVENTASCROPS")
                        Else
                            sReport = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "INFORMEVENTAS")
                        End If
                        Dim iCopias As Integer = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "COPIAS")
                        CrearPDF(sReport, oCompany.GetNewObjectKey, oDeliveryNotes, sPathPDF, GetLangIC(oAlbara.CardCode), iCopias)
                        '
                    End If
                    '
                End If
            Next fri
            '
        Catch ex As Exception
            SEI_Globals.Log("1.ERR. '" & Me.sBD_Name & "'", "Error preparando el Pedido de Ventas " & fri.Name.ToString & ". Causa: " & ex.Message)
            b_Errors = True
        End Try
    End Sub
    '

    '
    Private Function GetValueUNIDADESPESOR(ByVal sUnidad As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim sSQL As String = ""
        oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '
        sSQL = "SELECT T0.FldValue " & vbCrLf
        sSQL &= "From UFD1 T0 " & vbCrLf
        sSQL &= "LEFT JOIN CUFD T1 On T0.TableID = T1.TableID And T0.FieldID = T1.FieldID " & vbCrLf
        sSQL &= "WHERE T1.TableID = 'DLN1' and  AliasID = 'GSP_INUDESPESOR' and FldValue like '" & sUnidad & "'"
        oRS.DoQuery(sSQL)
        '
        If oRS.RecordCount > 0 Then
            Return oRS.Fields.Item("FldValue").Value
        End If
        '
        Return ""
        '
        LiberarObjCOM(oRS)
    End Function
    '
    Private Function RecuperarIdioma(ByVal sShortName As String) As Integer
        Dim oRS As SAPbobsCOM.Recordset
        Dim ls As String
        '
        oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ls = "SELECT Code FROM OLNG WHERE ShortName = '" & sShortName & "'"
        oRS.DoQuery(ls)
        Return oRS.Fields.Item("Code").Value
        '
        LiberarObjCOM(oRS)
    End Function
    '
    Private Sub Importaciones()

        SEI_Globals.LOG_Ruta = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "LOG")
        If Not SEI_Globals.LOG_Ruta.EndsWith("\") Then SEI_Globals.LOG_Ruta &= "\"

        SEI_Globals.LOG_Fichero = "Log_" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".txt"

        Try
            Me.sBD_Name = IniGet(Application.StartupPath & sFicheroIni, "Parametros", "BD")

            SEI_Globals.Log("1. DB '" & Me.sBD_Name & "'", "PROCESS: Inicializar Importación")
            Me.Msg_Conectar("Conectando...")
            Me.ConectarSBO("1.1")

            '--------------------------------------------------------------------------------------------------
            ' Inicialización: Obtener Datos, Validaciones 
            '--------------------------------------------------------------------------------------------------
            SEI_Globals.Obtener_DatosOADM(oCompany)

            '---------------------------------------------------------------------------------------------
            ' IMPORTAR ARCHIVOS
            '---------------------------------------------------------------------------------------------
            Me.Msg_Mensaje("Importando datos " & oCompany.CompanyName & "...")
            If SEI_Globals.b_Compras Then
                Me.ImportarArticulos()
                Me.ImportarAlbaraEntradaMercaderies()
            End If
            If SEI_Globals.b_Ventas Then
                Me.ImportarVentes()
            End If
            '
            SEI_Globals.Log("1. DB '" & Me.sBD_Name & "'", "PROCESS: Finalizar Importación")

        Catch ex As Exception
            SEI_Globals.Log("ERR. DB '" & Me.sBD_Name & "'", "ERROR: Importación Fallida: " & ex.Message)
            Me.Msg_Mensaje("Error durente la importación. Revise el LOG.")
        End Try

    End Sub
    '
    Private Sub Msg_Conectar(ByVal sMsg As String)

        Try
            Me.lblConectar.Text = sMsg
            Application.DoEvents()
            Me.Refresh()

        Catch ex As Exception
            Throw New Exception("Error Msg_Conectar: " & ex.Message)
        End Try

    End Sub

    Private Sub Msg_Mensaje(ByVal sMsg As String)

        Try
            Me.lblMsg.Text = sMsg
            Me.Refresh()

        Catch ex As Exception
            Throw New Exception("Error Msg_Mensaje: " & ex.Message)
        End Try

    End Sub

#End Region

    Private Sub bIntegracions_Click(sender As System.Object, e As System.EventArgs) Handles bIntegracions.Click

        Try
            Me.Msg_Conectar("")
            Me.Msg_Mensaje("")
            If IniGet(Application.StartupPath & sFicheroIni, "Parametros", "IMPORTAR") = "0" Then
                SEI_Globals.b_Ventas = True
            ElseIf IniGet(Application.StartupPath & sFicheroIni, "Parametros", "IMPORTAR") = "1" Then
                SEI_Globals.b_Compras = True
            ElseIf IniGet(Application.StartupPath & sFicheroIni, "Parametros", "IMPORTAR") = "2" Then
                SEI_Globals.b_Ventas = True
                SEI_Globals.b_Compras = True
            End If
            '--------------------------------------------------------------------------------------------------
            ' Importaciones
            '--------------------------------------------------------------------------------------------------
            Me.Importaciones()

            '--------------------------------------------------------------------------------------------------
            ' DESCONECTAR CONEXIÓN LA BD
            '--------------------------------------------------------------------------------------------------
            If Not IsNothing(oCompany) Then oCompany.Disconnect()
            If SEI_Globals.b_Errors Then
                Me.Msg_Mensaje("Procesos finalizados con errores. Revise el LOG.")
                SEI_Globals.b_Errors = False
            Else
                Me.Msg_Mensaje("Procesos finalizados")
            End If

        Catch ex As Exception
            MsgBox("Error bIntegracions_Click: " & ex.Message)
        Finally
            LiberarObjCOM(oCompany)
        End Try

    End Sub

    Private Sub Frm_Integraciones_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        If IniGet(Application.StartupPath & sFicheroIni, "Parametros", "FormLoad") = "T" Then
            Me.bIntegracions_Click(sender, e)
            Me.Close()
        End If

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub


End Class