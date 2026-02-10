Imports System.IO
Imports System.Globalization.CultureInfo
Imports System.Text
Imports System.Reflection

Imports SAPbobsCOM.BoObjectTypes

Module SEI_Globals

    Public oCompany As SAPbobsCOM.Company
    Public b_Ventas As Boolean = False
    Public b_Compras As Boolean = False
    Public b_Errors As Boolean = False

#Region "Funciones Ficheros"

    Public Sub Copy_Directory(ByVal sRuta_Origen As String, ByVal sRuta_Destino As String)

        Try
            My.Computer.FileSystem.CopyDirectory(sRuta_Origen, sRuta_Destino, True)

        Catch ex As Exception
            Throw New Exception("Error Copy_Directory (" & sRuta_Origen & "// " & sRuta_Destino & "): " & ex.Message)
        End Try

    End Sub

    Public Sub Create_Directory(ByVal sRuta As String)

        Try
            If Not Directory.Exists(sRuta) Then Directory.CreateDirectory(sRuta)

        Catch ex As Exception
            Throw New Exception("Error Create_Directory (" & sRuta & "): " & ex.Message)
        End Try

    End Sub

    Public Sub Delete_Directory(ByVal sRuta As String)

        Try
            If Directory.Exists(sRuta) Then Directory.Delete(sRuta, True)

        Catch ex As Exception
            Throw New Exception("Error Delete_Directory (" & sRuta & "): " & ex.Message)
        End Try

    End Sub

    Public Sub Delete_File(ByVal sFichero As String)

        Try
            If File.Exists(sFichero) Then File.Delete(sFichero)

        Catch ex As Exception
            Throw New Exception("Error Delete_File (" & sFichero & "): " & ex.Message)
        End Try

    End Sub

    Public Function Move_File(ByVal sRutaOrigen As String, ByVal sRutaDestino As String, ByVal sFichero As String) As Boolean

        Move_File = False

        Try
            If File.Exists(sRutaOrigen & sFichero) Then

                Create_Directory(sRutaDestino)

                File.Move(sRutaOrigen & sFichero, sRutaDestino & sFichero)

                Move_File = True

            End If

        Catch ex As Exception
            Throw New Exception("Error Move_File (" & sFichero & "): " & ex.Message)
        End Try

    End Function

    Public Function Read_File(ByVal asFile As String) As String

        Dim str As String = ""

        Try
            If File.Exists(asFile) Then
                Dim reader As New StreamReader(asFile)
                str = reader.ReadToEnd
                reader.Close()
            End If

        Catch ex As Exception
            Throw New Exception("Error ReadFile (" & asFile & "): " & ex.Message)
        End Try

        Return str

    End Function

    Public Sub Rename_File(ByVal sRuta As String, ByVal sArchivo As String)

        Try
            If File.Exists(sRuta) Then
                ' Renombrarlo con la función renameFile  
                My.Computer.FileSystem.RenameFile(sRuta, sArchivo)
            End If

        Catch ex As Exception
            Throw New Exception("Error Rename_File: " & ex.Message)
        End Try

    End Sub

    Public Sub Shell_File(ByVal sFile As String)

        Try
            If File.Exists(sFile) Then
                Shell("Notepad.exe """ & sFile & "", AppWinStyle.NormalFocus)
            End If

        Catch ex As Exception
            Throw New Exception("Error Shell_File: " & ex.Message)
        End Try

    End Sub

    Public Sub Write_File_Line(ByVal sRuta As String, ByVal sLinea As String)

        Try
            Dim oFt As New StreamWriter(sRuta, True, System.Text.Encoding.Default)

            oFt.WriteLine(sLinea)
            oFt.Flush()
            oFt.Close()

        Catch ex As Exception
            Throw New Exception("Error Write_File_Line (" & sRuta & " - " & sLinea & "): " & ex.Message)
        End Try

    End Sub

    Public Sub Write_File_Text(ByVal sRuta As String, ByVal sText As String)

        Try
            Dim oFt As New StreamWriter(sRuta, True, System.Text.Encoding.Default)

            oFt.Write(sText)
            oFt.Flush()
            oFt.Close()

        Catch ex As Exception
            Throw New Exception("Error Write_File_Text (" & sRuta & " - " & sText & "): " & ex.Message)
        End Try

    End Sub

#End Region

#Region "Funciones Fichero .INI"

    Public Const sFicheroIni As String = "\SEI_Importacions.ini"

    ' Leer una clave de un fichero INI
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Integer, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

    Public Function IniGet(ByVal sFileName As String, ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "") As String
        '--------------------------------------------------------------------------
        ' Devuelve el valor de una clave de un fichero INI
        ' Los parámetros son:
        '   sFileName   El fichero INI
        '   sSection    La sección de la que se quiere leer
        '   sKeyName    Clave
        '   sDefault    Valor opcional que devolverá si no se encuentra la clave
        '--------------------------------------------------------------------------
        ' sSection ->   "Parametros"
        ' sKeyName ->   "U" , "I" , "P"
        '
        ' [Parametros]
        ' U = sa
        ' I = IG
        ' P =seidor.65

        Dim ret As Integer
        Dim sRetVal As String
        '
        sRetVal = New String(Chr(0), 255)
        '
        ret = GetPrivateProfileString(sSection, sKeyName, sDefault, sRetVal, Len(sRetVal), sFileName)

        If ret = 0 Then
            Return sDefault
        Else
            Return Left(sRetVal, ret)
        End If

    End Function

#End Region

#Region "Funciones LOG"

    Public LOG_Ruta As String
    Public LOG_Fichero As String

    Public Sub Log(ByVal sProceso As String, ByVal sDescripcion As String)

        Dim sLinea As String
        Dim oFitxer As StreamWriter

        SEI_Globals.Create_Directory(LOG_Ruta)

        sLinea = sProceso & vbTab & sDescripcion

        oFitxer = File.AppendText(LOG_Ruta & LOG_Fichero)
        oFitxer.WriteLine(sLinea)
        oFitxer.Flush()
        oFitxer.Close()

    End Sub

#End Region

#Region "Funciones Tipos de Datos"

    Function NullToInt(ByVal Valor As Object) As Long

        Dim lValor As Long = 0

        If IsNothing(Valor) Then
        ElseIf IsDBNull(Valor) Or Valor.ToString = "" Then
        ElseIf Not IsNumeric(Valor) Then
        Else
            lValor = CInt(Valor)
        End If

        Return lValor

    End Function

    Function NullToLong(ByVal Valor As Object) As Long

        Dim lValor As Long = 0

        If IsNothing(Valor) Then
        ElseIf IsDBNull(Valor) Or Valor.ToString = "" Then
        ElseIf Not IsNumeric(Valor) Then
        Else
            lValor = CLng(Valor)
        End If

        Return lValor

    End Function

    Function NullToDate(ByVal Valor As Object, ByVal sFormato As String) As Date

        If IsDBNull(Valor) Or Valor.ToString = "" Then
            Return Nothing
        Else
            Dim dFecha As Date = Nothing

            Select Case sFormato
                Case "yyyyMMdd" : dFecha = DateSerial(Valor.ToString.Substring(0, 4), Valor.ToString.Substring(4, 2), Valor.ToString.Substring(6, 2))
                Case "ddMMyyyy" : dFecha = DateSerial(Valor.ToString.Substring(4, 4), Valor.ToString.Substring(2, 2), Valor.ToString.Substring(0, 2))
                Case "dd/MM/yyyy" : dFecha = DateSerial(Valor.ToString.Split("/")(0), Valor.ToString.Split("/")(1), Valor.ToString.Split("/")(2))
                Case "yyyy/MM/dd" : dFecha = DateSerial(Valor.ToString.Split("/")(2), Valor.ToString.Split("/")(1), Valor.ToString.Split("/")(0))
            End Select

        End If

    End Function

#Region "Funciones Double"

    Public SEPARADOR_DECIMALES_SAP As String
    Public SEPARADOR_DECIMALES_SYS As String
    Public MONEDA_DERECHA As Boolean

    Public Sub Obtener_DatosOADM(ByVal oC As SAPbobsCOM.Company)

        Dim ls As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        Try
            SEPARADOR_DECIMALES_SAP = ","
            SEPARADOR_DECIMALES_SYS = CurrentCulture.NumberFormat.NumberDecimalSeparator
            MONEDA_DERECHA = True

            ls = "SELECT DecSep, CurOnRight = ISNULL(CurOnRight, 'N') FROM OADM"

            oRs = oC.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(ls)

            If oRs.RecordCount > 0 Then
                SEPARADOR_DECIMALES_SAP = oRs.Fields.Item("DecSep").Value
                If oRs.Fields.Item("CurOnRight").Value <> "Y" Then MONEDA_DERECHA = False
            End If

        Catch ex As Exception
            Throw New Exception("Error Obtener_DecSep: " & ex.Message)
        Finally
            LiberarObjCOM(oRs)
        End Try

    End Sub

    Function Convert_DoubleToString(ByVal Valor As Double) As String

        ' El valor que entre tiene que hacerlo solo con el separador de los decimales, el de los millares no debe de estar

        Try
            Dim sValor As String = ""

            sValor = Valor.ToString

            If sValor.Contains(",") Then
                sValor = sValor.Replace(",", SEPARADOR_DECIMALES_SAP)
            ElseIf sValor.Contains(".") Then
                sValor = sValor.Replace(".", SEPARADOR_DECIMALES_SAP)
            End If

            Return sValor

        Catch ex As Exception
            Throw New Exception("Error Convert_DoubleToString: " & ex.Message)
        End Try

    End Function

    Function Convert_DoubleToStringSQL(ByVal Valor As String) As String

        ' El valor que entre tiene que hacerlo solo con el separador de los decimales, el de los millares no debe de estar

        Try
            If Valor = "" Then Valor = "NULL"

            If Valor.Contains(",") Then
                Valor = Valor.Replace(",", ".")
            End If

            Return Valor

        Catch ex As Exception
            Throw New Exception("Error Convert_DoubleToString: " & ex.Message)
        End Try

    End Function

    Function Convert_StringToDouble(ByVal Valor As String) As Double

        ' El valor que entre tiene que hacerlo solo con el separador de los decimales, el de los millares no debe de estar

        Try
            If IsNothing(Valor) Then
                Return 0
            ElseIf IsDBNull(Valor) Or Trim(Valor.GetType.ToString) = "" Or Trim(Valor.ToString) = "" Then
                Return 0
            Else
                Dim aValor() As String = Valor.Split(SEPARADOR_DECIMALES_SAP)

                If SEPARADOR_DECIMALES_SAP = "," Then
                    Valor = aValor(0).Replace(".", "")
                Else
                    Valor = aValor(0).Replace(",", "")
                End If

                If aValor.Length = 1 Then
                Else
                    Valor = Valor & SEPARADOR_DECIMALES_SYS & aValor(1)
                End If

                Return Convert.ToDouble(Valor)

            End If

        Catch ex As Exception
            Throw New Exception("Error Convert_StringToDouble: " & ex.Message)
        End Try

    End Function

    Function Convert_StringToDouble(ByVal Valor As String, ByVal iNumDecimales As Integer) As Double

        ' El valor que entre tiene que hacerlo solo con el separador de los decimales, el de los millares no debe de estar

        Try
            If IsNothing(Valor) Then
                Return 0
            ElseIf IsDBNull(Valor) Or Trim(Valor.ToString) = "" Then
                Return 0
            Else
                Dim sValor As String = ""
                Dim sDecimales As String = ""
                Dim dValor As Double = 0

                sValor = Valor.Substring(0, Valor.Length - iNumDecimales)
                sDecimales = Right(Valor, iNumDecimales)

                sValor = sValor & SEPARADOR_DECIMALES_SYS & sDecimales

                dValor = sValor

                Return dValor

            End If

        Catch ex As Exception
            Throw New Exception("Error Convert_StringToDouble: " & ex.Message)
        End Try

    End Function

    Function GetDouble(ByVal Valor As String) As Double

        Try
            If IsNothing(Valor) Then
                Return 0
            ElseIf IsDBNull(Valor) Or Trim(Valor.GetType.ToString) = "" Or Trim(Valor.ToString) = "" Then
                Return 0
            Else
                Dim aValor() As String = Valor.Split(" ")

                If MONEDA_DERECHA Then
                    Return Convert_StringToDouble(aValor(0))
                Else
                    Return Convert_StringToDouble(aValor(1))
                End If

            End If

        Catch ex As Exception
            Throw New Exception("Error GetDouble: " & ex.Message)
        End Try

    End Function

#End Region

#Region "Funciones Text"

    Function NullToText(ByVal Valor As Object) As String

        Dim sValor As String = ""

        If IsNothing(Valor) Then
        ElseIf IsDBNull(Valor) Or Valor.ToString = "" Then
        Else
            sValor = Valor
        End If

        Return sValor

    End Function

    Function NullToText_C(ByVal Valor As Object, ByVal sCaracter As String, ByVal iLen As Integer, ByVal bRight As Boolean) As String

        ' Rellenar con caracteres

        Dim sValor As String = sCaracter

        If IsNothing(Valor) Then
        ElseIf IsDBNull(Valor) Or Valor.ToString = "" Then
        Else
            sValor = Valor
        End If

        If bRight Then
            If sValor.Length < iLen Then sValor = sValor.PadRight(iLen, sCaracter)
        Else
            If sValor.Length < iLen Then sValor = sValor.PadLeft(iLen, sCaracter)
        End If

        If sValor.Length > iLen Then sValor = sValor.Substring(0, iLen)

        Return sValor

    End Function

    Function NullToText_D(ByVal Valor As Double, ByVal iNumDecimales As Integer, ByVal sSeparador As String) As String

        ' Convertir un valor Integer/Double a Text

        Dim sValor As String = "0"
        Dim sDecimales As String = ""

        If IsNothing(Valor) Then
        Else
            sValor = Convert_DoubleToString(Valor)
        End If

        If sValor.Split(SEPARADOR_DECIMALES_SAP).Length = 2 Then
            sDecimales = sValor.Split(SEPARADOR_DECIMALES_SAP)(1)
        End If

        If sDecimales.Length < iNumDecimales Then sDecimales = sDecimales.PadRight(iNumDecimales, "0")

        sValor = sValor.Split(SEPARADOR_DECIMALES_SAP)(0)

        If sDecimales <> "" Then sValor &= sSeparador & sDecimales

        Return sValor

    End Function

    Function NullToText_F(ByVal Valor As Date, ByVal sFormato As String) As String

        Dim sValor As String = ""

        If IsNothing(Valor) Then
        ElseIf IsDBNull(Valor) Or Valor.ToString = "" Then
        Else
            Dim dFecha As Date
            dFecha = Valor
            If dFecha.ToString("yyyyMMdd") = "00010101" Or dFecha.ToString("yyyyMMdd") = "18991230" Then
            Else
                sValor = dFecha.ToString(sFormato)
            End If
        End If

        Return sValor

    End Function

    Function NullToText_SQL(ByVal Valor As Object) As String

        Dim sValor As String = ""

        If IsNothing(Valor) Then
        ElseIf IsDBNull(Valor) Or Valor.ToString = "" Then
        Else
            sValor = Valor.ToString.Replace("'", "''")
        End If

        Return sValor

    End Function

#End Region

#End Region

#Region "Funciones Varias"

    Public Sub LiberarObjCOM(ByRef oObjCOM As Object)

        'Liberar y destruir Objecto com
        If Not IsNothing(oObjCOM) Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oObjCOM)
            oObjCOM = Nothing
            GC.Collect()
        End If

    End Sub

    Public Function RecuperarValores(ByVal sCampos As String, ByVal sTabla As String, ByVal sFiltro As String, ByVal sCondicion As String, ByRef aRetorno() As String) As String

        ' sCampos       --> los campos deseados separados por coma
        ' sTabla        --> se escribe tal cual se haría para la SQL
        ' sFiltro       --> nombre del campo 1|valor del campo 1;nombre del campo 2|valor del campo 2
        ' sCondicion    --> extras como diferencias, ORDER BY, ...

        Dim ls As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        Dim aContenido() As String = Nothing
        Dim aValores() As String = Nothing
        Dim i As Integer = 0

        RecuperarValores = ""

        Try
            ls = "SELECT " & sCampos & vbCrLf
            ls &= "FROM " & sTabla & vbCrLf

            If sFiltro <> "" Then

                aContenido = sFiltro.Split(";")

                ls &= "WHERE "

                For i = 0 To aContenido.Length - 1

                    If i <> 0 Then ls &= "AND "

                    aValores = aContenido(i).Split("|")

                    ls &= aValores(0) & " = '" & aValores(1) & "'" & vbCrLf

                Next

            End If

            If sCondicion <> "" Then ls &= sCondicion

            oRs = oCompany.GetBusinessObject(BoRecordset)
            oRs.DoQuery(ls)

            If Not oRs.EoF Then

                If Not IsNothing(aRetorno) Then

                    aValores = sCampos.Split(",")

                    ReDim aRetorno(aValores.Length - 1)

                    For i = 0 To aRetorno.Length - 1
                        aRetorno(i) = SEI_Globals.NullToText(oRs.Fields.Item(i).Value)
                    Next

                Else
                    RecuperarValores = SEI_Globals.NullToText(oRs.Fields.Item(0).Value)
                End If

            End If

        Catch ex As Exception
            Throw New Exception("Error RecuperarValores: " & ex.Message)
        Finally
            LiberarObjCOM(oRs)
        End Try

    End Function

#End Region

End Module
