
Option Explicit On

Imports SEI_Importacions.SEI_AddOnEnum
Imports System.Threading
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Collections
Imports SAPbobsCOM.BoObjectTypes


Module SEI_Report

#Region "Variables"
    '
    Public Enum eOrigen
        CampoBlob
        RecursoIncustado
        Directorio
    End Enum
    '
    Private _Origen As eOrigen        ' Origen Fichero .rpt
    Private _BringToFront As Boolean  ' Poner el formulario delante 
    Private _TipoContructor As Integer
    Private _Where As String
    Private _Informe As String
    Private _Formulas As ArrayList
    Private _SubReports_Where As ArrayList
    Private _stSubReports_Where As st_SubReportWhere
    Private _InformeCrystal As CrystalDecisions.CrystalReports.Engine.ReportDocument
    Private _Copias As Integer                             ' Nº de copias
    Private _OrigenDatos As String ' "SQL" , "XML"
    Private _DsDatos As DataSet
    Private sFileRpt As String
    Private _OrigenCrystal As eCrystal
    Private _PrintPDF As Boolean
    Private _PrintPDF_File As String
    Public _stEmail As st_Email
    Private _Impresora As String
    '
#End Region

#Region "Thread"

    Private Sub ThreadReportVistaPrevia()
        '
        ' Mostramos el report por pantalla (cargando el crystal report viewer)
        ' Todo el código para crear el Report está en el evento Load del formulario de windows
        ' Tenemos que crear un thread (hilo) para poder mostrarlo independientemente de SBO
        '
        Dim myThread As New Thread(AddressOf PresentacionPreliminar)
        myThread.TrySetApartmentState(ApartmentState.STA)
        myThread.Start()
        '
    End Sub

    Private Sub ThreadReportImprimir()
        '
        ' Mostramos el report por pantalla (cargando el crystal report viewer)
        ' Todo el código para crear el Report está en el evento Load del formulario de windows
        ' Tenemos que crear un thread (hilo) para poder mostrarlo independientemente de SBO
        '
        '
        Dim myThread As New Thread(AddressOf ImprimirReport)
        myThread.TrySetApartmentState(ApartmentState.STA)
        myThread.Start()

        '
    End Sub

#End Region

#Region "Imprimir Thread"

    Public Sub ImprimirThread(ByVal eOrigen As eOrigen, ByRef sInforme As String, ByVal sWhere As String)
        '
        _Origen = eOrigen
        _TipoContructor = 1    ' Flag que indica el tipo de contructor
        _Where = sWhere
        _Informe = sInforme
        '
        ThreadReportImprimir()
        '
    End Sub

    Public Sub ImprimirThread(ByVal sInforme As String, _
                           ByVal sWhere As String, _
                           ByVal aFormulas As ArrayList)
        '
        _TipoContructor = 2    ' Flag que indica el tipo de contructor
        _Where = sWhere
        _Informe = sInforme
        _Formulas = aFormulas
        '
        ThreadReportImprimir()
        '
    End Sub

    Public Sub ImprimirThread(ByVal eOrigen As eOrigen, _
                              ByVal sInforme As String, _
                              ByVal sWhere As String, _
                              ByVal aFormulas As ArrayList, _
                              ByVal stSubReports_Where As st_SubReportWhere)
        '
        _Origen = eOrigen
        _TipoContructor = 30    ' Flag que indica el tipo de contructor
        _Informe = sInforme
        _Where = sWhere
        _Formulas = aFormulas
        _stSubReports_Where = stSubReports_Where
        '
        ThreadReportImprimir()
        '
    End Sub

    Public Sub ImprimirThread(ByVal sInforme As String, _
                           ByVal sWhere As String, _
                           ByVal aFormulas As ArrayList, _
                           ByVal aSubReports_Where As ArrayList)
        '
        _TipoContructor = 3    ' Flag que indica el tipo de contructor
        _Informe = sInforme
        _Where = sWhere
        _Formulas = aFormulas
        _SubReports_Where = aSubReports_Where
        '
        ThreadReportImprimir()
        '
    End Sub

    Public Sub ImprimirThread(ByVal sInforme As String,
                           ByVal sWhere As String,
                           ByVal aFormulas As ArrayList,
                           ByVal iCopias As Integer)
        '
        _TipoContructor = 5    ' Flag que indica el tipo de contructor
        _Where = sWhere
        _Informe = sInforme
        _Formulas = aFormulas
        _Copias = iCopias
        '
        ThreadReportImprimir()
        '
    End Sub

    Public Sub ImprimirThread(ByVal eOrigenCrystal As eCrystal,
                        ByRef sInforme As String,
                        ByVal aFormulas As ArrayList,
                        ByVal stSubReportWhere As st_SubReportWhere,
                        ByVal bPrintPDF As Boolean,
                        ByVal sPrintPDF_File As String)
        '
        _OrigenCrystal = eOrigenCrystal
        _Where = ""
        _Formulas = Nothing
        _SubReports_Where = Nothing
        '
        _TipoContructor = 50   ' Flag que indica el tipo de contructor
        _Formulas = aFormulas
        _stSubReports_Where = stSubReportWhere
        _Informe = sInforme
        _PrintPDF = bPrintPDF
        _PrintPDF_File = sPrintPDF_File
        '
        ThreadReportImprimir()
        '
    End Sub

#End Region

#Region "Imprimir"

    Public Sub Imprimir(ByRef sInforme As String, ByVal sWhere As String)
        '
        _TipoContructor = 1    ' Flag que indica el tipo de contructor
        _Where = sWhere
        _Informe = sInforme
        '
        ImprimirReport()
        '
    End Sub

    Public Sub Imprimir(ByVal sInforme As String, _
                           ByVal sWhere As String, _
                           ByVal aFormulas As ArrayList)
        '
        _TipoContructor = 2    ' Flag que indica el tipo de contructor
        _Where = sWhere
        _Informe = sInforme
        _Formulas = aFormulas
        '
        ImprimirReport()
        '
    End Sub

    Public Sub Imprimir(ByVal sInforme As String, _
                           ByVal sWhere As String, _
                           ByVal aFormulas As ArrayList, _
                           ByVal aSubReports_Where As ArrayList)
        '
        _TipoContructor = 3    ' Flag que indica el tipo de contructor
        _Informe = sInforme
        _Where = sWhere
        _Formulas = aFormulas
        _SubReports_Where = aSubReports_Where
        '
        ImprimirReport()
        '
    End Sub

    Public Sub Imprimir(ByVal sInforme As String, _
                           ByVal sWhere As String, _
                           ByVal aFormulas As ArrayList, _
                           ByVal iCopias As Integer)
        '
        _TipoContructor = 5    ' Flag que indica el tipo de contructor
        _Where = sWhere
        _Informe = sInforme
        _Formulas = aFormulas
        _Copias = iCopias
        '
        ImprimirReport()
        '
    End Sub

    Public Sub Imprimir(ByVal eOrigen As eOrigen, _
                              ByVal sInforme As String, _
                              ByVal sWhere As String, _
                              ByVal aFormulas As ArrayList, _
                              ByVal stSubReports_Where As st_SubReportWhere)
        '
        _Origen = eOrigen
        _TipoContructor = 30    ' Flag que indica el tipo de contructor
        _Informe = sInforme
        _Where = sWhere
        _Formulas = aFormulas
        _stSubReports_Where = stSubReports_Where
        '
        ImprimirReport()
        '
    End Sub

    Public Sub Imprimir(ByVal eOrigen As eOrigen, _
                           ByVal sInforme As String, _
                          ByVal stSubReports_Where As st_SubReportWhere, _
                          ByVal aFormulas As ArrayList, _
                          ByVal iCopias As Integer, _
                          ByVal sImpresora As String)
        '
        _Origen = eOrigen
        _TipoContructor = 6    ' Flag que indica el tipo de contructor
        _stSubReports_Where = stSubReports_Where
        _Informe = sInforme
        _Formulas = aFormulas
        _Copias = iCopias
        _Impresora = sImpresora
        '
        ImprimirReport()
        '
    End Sub

#End Region

#Region "Imprimir Report"

    Private Sub ImprimirReport()
        '
        'SEI_SRV_MAIL.lblmsg.Text = "Cargando documento, un momento por favor..."
        Application.DoEvents()
        '
        Dim oReport As Object = Nothing

        Select Case _Origen
            Case eOrigen.CampoBlob
                Dim oRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                sFileRpt = ObtenerRptSAP(_Informe)
                oRpt.Load(sFileRpt)
                oReport = oRpt

            Case eOrigen.Directorio
                oReport = New CrystalDecisions.CrystalReports.Engine.ReportDocument
                oReport.Load(_Informe)

            Case eOrigen.RecursoIncustado
                Dim objectType As Type = Type.GetType("SEI.SEI_Importacions." & _Informe, True)
                oReport = Activator.CreateInstance(objectType) ' Crea la instancia del objeto 
        End Select
        '
        Select Case _TipoContructor
            '
            Case 1
                Dim oReportCrystal As New SEI_ReportCrystal(oReport, _Where)
                oReportCrystal.Imprimir()
                oReportCrystal = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
                '
            Case 2
                Dim oReportCrystal As New SEI_ReportCrystal(oReport, _Where, _Formulas)
                oReportCrystal.Imprimir()
                oReportCrystal = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
                '
            Case 3
                Dim oReportCrystal As New SEI_ReportCrystal(oReport, _Where, _Formulas, _SubReports_Where)
                oReportCrystal.Imprimir()
                oReportCrystal = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
                '
            Case 5
                Dim oReportCrystal As New SEI_ReportCrystal(oReport, _Where, _Formulas, _Copias)
                oReportCrystal.Imprimir()
                oReportCrystal = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
                '
            Case 6
                Dim oReportCrystal As New SEI_ReportCrystal(oReport, _stSubReports_Where, _Formulas, _Impresora, _Copias)
                oReportCrystal.Imprimir()
                oReportCrystal = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()

        End Select
        
        'Eliminar Fichero Temporal
        If IO.File.Exists(sFileRpt) Then
            IO.File.Delete(sFileRpt)
        End If
    End Sub

#End Region

#Region "Vista Previa Thread"

    Public Sub VistaPreviaThread(ByRef sInforme As String, _
                           ByVal sWhere As String, _
                           Optional ByVal bBringToFront As Boolean = True)
        '
        _TipoContructor = 1    ' Flag que indica el tipo de contructor
        _Where = sWhere
        _Informe = sInforme
        _BringToFront = bBringToFront
        '
        ThreadReportVistaPrevia()
        '
    End Sub

    Public Sub VistaPreviaThread(ByVal sInforme As String, _
                           ByVal sWhere As String, _
                           ByVal aFormulas As ArrayList, _
                           Optional ByVal bBringToFront As Boolean = True)
        '
        _TipoContructor = 2    ' Flag que indica el tipo de contructor
        _Where = sWhere
        _Informe = sInforme
        _Formulas = aFormulas
        _BringToFront = bBringToFront
        '
        ThreadReportVistaPrevia()
        '
    End Sub

    Public Sub VistaPreviaThread(ByVal sInforme As String, _
                           ByVal sWhere As String, _
                           ByVal aFormulas As ArrayList, _
                           ByVal aSubReports_Where As ArrayList, _
                           Optional ByVal bBringToFront As Boolean = True)
        '
        _TipoContructor = 3    ' Flag que indica el tipo de contructor
        _Informe = sInforme
        _Where = sWhere
        _Formulas = aFormulas
        _SubReports_Where = aSubReports_Where
        _BringToFront = bBringToFront
        '
        ThreadReportVistaPrevia()
        '
    End Sub
    '
    Public Sub VistaPreviaThread(ByVal eOrigen As eOrigen, _
                           ByVal sInforme As String, _
                           ByVal sWhere As String, _
                           ByVal aFormulas As ArrayList, _
                           ByVal stSubReports_Where As st_SubReportWhere, _
                           Optional ByVal bBringToFront As Boolean = True)

        _Origen = eOrigen
        _TipoContructor = 30    ' Flag que indica el tipo de contructor
        _Informe = sInforme
        _Where = sWhere
        _Formulas = aFormulas
        _stSubReports_Where = stSubReports_Where
        _BringToFront = bBringToFront

        ThreadReportVistaPrevia()

    End Sub

    Public Sub VistaPreviaThread(ByVal sInforme As String, _
                       ByVal sOrigenDatos As String, _
                       ByVal DSdatos As DataSet, _
                       Optional ByVal bBringToFront As Boolean = True)
        '
        ' Origen de Datos XML
        '
        _TipoContructor = 40    ' Flag que indica el tipo de contructor
        _Informe = sInforme
        _OrigenDatos = sOrigenDatos
        _DsDatos = DSdatos
        _BringToFront = bBringToFront
        '
        ThreadReportVistaPrevia()
        '
    End Sub

    Public Sub VistaPreviaThread(ByVal sInforme As String, _
                       ByVal sOrigenDatos As String, _
                       ByRef DSdatos As DataSet, _
                       ByVal aFormulas As ArrayList, _
                       Optional ByVal bBringToFront As Boolean = True)
        '
        ' Origen de Datos XML
        '
        _TipoContructor = 41    ' Flag que indica el tipo de contructor
        _Informe = sInforme
        _OrigenDatos = sOrigenDatos
        _DsDatos = DSdatos
        _Formulas = aFormulas
        _BringToFront = bBringToFront
        '
        ThreadReportVistaPrevia()
        '
    End Sub

#End Region

#Region "Vista Previa"

    Public Sub VistaPrevia(ByRef sInforme As String, _
                           ByVal sWhere As String, _
                           Optional ByVal bBringToFront As Boolean = True)
        '
        _TipoContructor = 1    ' Flag que indica el tipo de contructor
        _Where = sWhere
        _Informe = sInforme
        _BringToFront = bBringToFront
        '
        PresentacionPreliminar()
        '
    End Sub

    Public Sub VistaPrevia(ByVal sInforme As String, _
                           ByVal sWhere As String, _
                           ByVal aFormulas As ArrayList, _
                           Optional ByVal bBringToFront As Boolean = True)
        '
        _TipoContructor = 2    ' Flag que indica el tipo de contructor
        _Where = sWhere
        _Informe = sInforme
        _Formulas = aFormulas
        _BringToFront = bBringToFront
        '
        PresentacionPreliminar()
        '
    End Sub

    Public Sub VistaPrevia(ByVal sInforme As String, _
                           ByVal sWhere As String, _
                           ByVal aFormulas As ArrayList, _
                           ByVal aSubReports_Where As ArrayList, _
                           Optional ByVal bBringToFront As Boolean = True)
        '
        _TipoContructor = 3    ' Flag que indica el tipo de contructor
        _Informe = sInforme
        _Where = sWhere
        _Formulas = aFormulas
        _SubReports_Where = aSubReports_Where
        _BringToFront = bBringToFront
        '
        PresentacionPreliminar()
        '
    End Sub

#End Region

#Region "Presentacion Preliminar"

    Private Sub PresentacionPreliminar()
        '
        'SEI_SRV_MAIL.lblmsg.Text = "Previsualización en curso, un momento por favor..."
        Application.DoEvents()
        '
        Dim oReport As Object = Nothing

        Select Case _Origen
            Case eOrigen.CampoBlob
                Dim oRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                sFileRpt = ObtenerRptSAP(_Informe)
                oRpt.Load(sFileRpt)
                oReport = oRpt

            Case eOrigen.Directorio

            Case eOrigen.RecursoIncustado
                Dim objectType As Type = Type.GetType("SEI.SEI_Importacions." & _Informe, True)
                oReport = Activator.CreateInstance(objectType) ' Crea la instancia del objeto 
        End Select
        '
        Select Case _TipoContructor
            '
            Case 1
                Dim oReportCrystal As New SEI_ReportCrystal(oReport, _Where, _BringToFront)
                oReportCrystal.VistaPrevia()
                oReportCrystal = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
                '
            Case 2
                Dim oReportCrystal As New SEI_ReportCrystal(oReport, _Where, _Formulas, _BringToFront)
                oReportCrystal.VistaPrevia()
                oReportCrystal = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
                '
            Case 3
                Dim oReportCrystal As New SEI_ReportCrystal(oReport, _Where, _Formulas, _SubReports_Where, _BringToFront)
                oReportCrystal.VistaPrevia()
                oReportCrystal = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
                '
            Case 40
                Dim oReportCrystal As New SEI_ReportCrystal(oReport, _OrigenDatos, _DsDatos, _BringToFront)
                oReportCrystal.VistaPrevia()
                oReportCrystal = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()

            Case 41
                Dim oReportCrystal As New SEI_ReportCrystal(oReport, _OrigenDatos, _DsDatos, _Formulas, _BringToFront)
                oReportCrystal.VistaPrevia()
                oReportCrystal = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()

        End Select
        '
        'Eliminar Fichero Temporal
        If IO.File.Exists(sFileRpt) Then
            IO.File.Delete(sFileRpt)
        End If
    End Sub

    Public Function ObtenerRptSAP(ByVal sValue As String, Optional ByVal sCampName As String = "DocCode") As String
        '
        ' Por defecto se buscará por DocName
        '
        Dim oCompanyService As SAPbobsCOM.CompanyService = Nothing
        Dim oBlobParams As SAPbobsCOM.BlobParams = Nothing
        Dim blobNewFilePath As String
        Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment = Nothing
        '
        ObtenerRptSAP = ""
        '
        Try
             'get company service
            oCompanyService = oCompany.GetCompanyService

            ' Specify the table and blob field 
            oBlobParams = oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams)
            oBlobParams.Table = "RDOC"
            oBlobParams.Field = "Template"

            '// Specify the file name to which to write the blob 
            blobNewFilePath = Application.StartupPath & "\" & oCompany.UserSignature & Now.ToString("yyyyMMddhhmmss") & ".rpt"
            oBlobParams.FileName = blobNewFilePath

            '// Specify the key field and value of the row from which to get the blob 
            oKeySegment = oBlobParams.BlobTableKeySegments.Add()
            oKeySegment.Name = sCampName   ' DocCode - DocName
            oKeySegment.Value = sValue     ' Sample ->"RCRI0001"

            '// Save the blob to the file 
            oCompanyService.SaveBlobToFile(oBlobParams)
            '
            ObtenerRptSAP = blobNewFilePath

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            LiberarObjCOM(oCompanyService)
            LiberarObjCOM(oBlobParams)
            LiberarObjCOM(oKeySegment)
        End Try

    End Function

#End Region

#Region "Imprimir PDF - FAX"
    Public Sub ImprimirPDF_EMAIL(ByVal eOrigenCrystal As eCrystal, _
                            ByRef sInforme As String, _
                            ByVal sWhere As String, _
                            ByVal bPrintPDF As Boolean, _
                            ByVal sPrintPDF_File As String, _
                            ByVal stEmail As st_Email)

        _OrigenCrystal = eOrigenCrystal
        _TipoContructor = 6    ' Flag que indica el tipo de contructor
        _Where = sWhere
        _Informe = sInforme
        _PrintPDF = bPrintPDF
        _PrintPDF_File = sPrintPDF_File
        _stEmail = stEmail

        ImprimirReportPDF_EMAIL()

    End Sub
    '
    Public Sub ImprimirPDF_EMAIL(ByVal eOrigenCrystal As eCrystal, _
                        ByRef sInforme As String, _
                        ByVal stSubReportWhere As st_SubReportWhere, _
                        ByVal bPrintPDF As Boolean, _
                        ByVal sPrintPDF_File As String, _
                        ByVal stEmail As st_Email)

        _OrigenCrystal = eOrigenCrystal
        _Where = ""
        _Formulas = Nothing
        _SubReports_Where = Nothing
        '
        _TipoContructor = 7    ' Flag que indica el tipo de contructor
        _stSubReports_Where = stSubReportWhere
        _Informe = sInforme
        _PrintPDF = bPrintPDF
        _PrintPDF_File = sPrintPDF_File
        _stEmail = stEmail

        ImprimirReportPDF_EMAIL()


    End Sub
    '
    Public Sub ImprimirPDF_EMAIL(ByVal eOrigenCrystal As eCrystal,
                        ByRef sInforme As String,
                        ByVal aFormulas As ArrayList,
                        ByVal stSubReportWhere As st_SubReportWhere,
                        ByVal bPrintPDF As Boolean,
                        ByVal sPrintPDF_File As String,
                        ByVal sPrinter As String,
                        Optional ByVal iCopias As Integer = 1)

        _OrigenCrystal = eOrigenCrystal
        _Where = ""
        _Formulas = Nothing
        _SubReports_Where = Nothing
        '
        _TipoContructor = 8    ' Flag que indica el tipo de contructor
        _Formulas = aFormulas
        _stSubReports_Where = stSubReportWhere
        _Informe = sInforme
        _PrintPDF = bPrintPDF
        _PrintPDF_File = sPrintPDF_File
        _Impresora = sPrinter
        _Copias = iCopias

        ImprimirReportPDF_EMAIL()


    End Sub
    '
    Private Sub ImprimirReportPDF_EMAIL()
        '
        Dim oReport As Object = Nothing
        Dim sPath As String = Application.StartupPath & "\Reports"
        Dim ReportPath As String = sPath & "\" & _Informe & ".rpt"
        '
        Application.DoEvents()
        '
        Select Case _OrigenCrystal
            Case eCrystal.Incrustado
                Dim objectType As Type = Type.GetType("SEI_Importacions." & _Informe, True)
                'Dim oReport As Object = Activator.CreateInstance(objectType) ' Crea la instancia del objeto 
                oReport = Activator.CreateInstance(objectType) ' Crea la instancia del objeto 

            Case eCrystal.EnDirectorio
                oReport = New CrystalDecisions.CrystalReports.Engine.ReportDocument
                oReport.Load(ReportPath)

            Case eCrystal.CampoBlob
                Dim oRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                sFileRpt = ObtenerRptSAP(_Informe)
                oRpt.Load(sFileRpt)
                oReport = oRpt

        End Select
        '
        Select Case _TipoContructor
            '
            Case 6
                '
                Dim oReportCrystal As New SEI_ReportCrystal(oReport, _Where, _PrintPDF, _PrintPDF_File)
                oReportCrystal.Imprimir()
                'Enviar_Email(oReportCrystal.PrintPDF_FilePath)
                'If File.Exists(oReportCrystal.PrintPDF_FilePath) Then
                '    File.Delete(oReportCrystal.PrintPDF_FilePath)
                'End If

            Case 7
                '
                Dim oReportCrystal As New SEI_ReportCrystal(oReport, _stSubReports_Where, _PrintPDF, _PrintPDF_File)
                oReportCrystal.Imprimir()

            Case 8
                '
                Dim oReportCrystal As New SEI_ReportCrystal(oReport, _Formulas, _stSubReports_Where, _PrintPDF, _PrintPDF_File, _Impresora, _Copias)
                oReportCrystal.Imprimir()

        End Select
    End Sub
    '
#End Region

End Module
