Option Explicit On
'
Imports SEI_Importacions.SEI_AddOnEnum
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Text
Imports System.Collections
Imports System.Threading
Imports System.Data
Imports System.IO

Public Class SEI_ReportCrystal

    Protected _BringToFront As Boolean  ' Poner el formulario delante 
    Protected _Where As String
    Protected _Informe As String
    Protected _Formulas As ArrayList
    Protected _SubReports_Where As ArrayList
    Protected _stSubReports_Where As st_SubReportWhere
    Protected _InformeCrystal As CrystalDecisions.CrystalReports.Engine.ReportDocument
    Protected _TipoConstructor As Integer
    Protected _Copias As Integer
    Protected _OrigenDatos As String ' "SQL" , "XML"
    Protected _DsDatos As DataSet
    Protected _NomInforme As String
    Protected _Impresora As String
    Protected _Printer As String
    Protected _PrintPDF As Boolean
    Protected _PrintPDF_File As String
    Protected _PrintPDF_FilePath As String


    ' Obtener Ruta Fichero PDF
    Public Property PrintPDF_FilePath() As String
        Get
            PrintPDF_FilePath = _PrintPDF_FilePath
        End Get
        Set(ByVal value As String)
            _PrintPDF_FilePath = value
        End Set
    End Property
    ' Nombre del Fichero
    Public Property PrintPDF_File() As String
        Get
            PrintPDF_File = _PrintPDF_File
        End Get
        Set(ByVal value As String)
            _PrintPDF_File = value
        End Set
    End Property
#Region "Constructor"

    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                   ByVal sWhere As String, _
                   Optional ByVal bBringToFront As Boolean = True)
        '
        _Where = sWhere
        _InformeCrystal = oInforme
        _TipoConstructor = 1          ' Flag para saber como se ha instanciado el objeto
        _BringToFront = bBringToFront ' Poner el formulario delante 
        '
    End Sub

    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                   ByVal sWhere As String, _
                   ByVal aFormulas As ArrayList, _
                   Optional ByVal bBringToFront As Boolean = True)
        '
        _Where = sWhere
        _InformeCrystal = oInforme
        _Formulas = aFormulas
        _TipoConstructor = 2  ' Flag para saber como se ha instanciado el objeto
        _BringToFront = bBringToFront ' Poner el formulario delante 
        '
    End Sub

    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
               ByVal sWhere As String, _
               ByVal aFormulas As ArrayList, _
               ByVal aSubReports_Where As ArrayList, _
               Optional ByVal bBringToFront As Boolean = True)
        '
        _Where = sWhere
        _InformeCrystal = oInforme
        _Formulas = aFormulas
        _SubReports_Where = aSubReports_Where
        _TipoConstructor = 3  ' Flag para saber como se ha instanciado el objeto
        _BringToFront = bBringToFront ' Poner el formulario delante 
        '
    End Sub
    '
    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                   ByVal sWhere As String, _
                   ByVal aFormulas As ArrayList, _
                   ByVal iCopias As Integer)
        '
        _Where = sWhere
        _InformeCrystal = oInforme
        _Formulas = aFormulas
        _TipoConstructor = 5  ' Flag para saber como se ha instanciado el objeto
        _Copias = iCopias     ' Nº de Copias
        '
    End Sub
    '
    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                   ByVal sWhere As String, _
                   ByVal bPrintPDF As Boolean, _
                   ByVal sPrintPDF_File As String)

        _Where = sWhere
        _InformeCrystal = oInforme
        _TipoConstructor = 6  ' Flag para saber como se ha instanciado el objeto
        _PrintPDF = bPrintPDF ' Exportar a PDF
        _PrintPDF_File = sPrintPDF_File     ' Nombre de fichero a exportar a PDF

    End Sub
    '
    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                   ByVal stSubReportWhere As st_SubReportWhere, _
                   ByVal bPrintPDF As Boolean, _
                   ByVal sPrintPDF_File As String)

        _stSubReports_Where = stSubReportWhere
        _InformeCrystal = oInforme
        _TipoConstructor = 7  ' Flag para saber como se ha instanciado el objeto
        _PrintPDF = bPrintPDF ' Exportar a PDF
        _PrintPDF_File = sPrintPDF_File     ' Nombre de fichero a exportar a PDF

    End Sub
    '
    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument,
                    ByVal aFormulas As ArrayList,
                   ByVal stSubReportWhere As st_SubReportWhere,
                   ByVal bPrintPDF As Boolean,
                   ByVal sPrintPDF_File As String,
                   ByVal sImpresora As String,
                   ByVal iCopias As Integer)

        _Formulas = aFormulas
        _stSubReports_Where = stSubReportWhere
        _InformeCrystal = oInforme
        _TipoConstructor = 8  ' Flag para saber como se ha instanciado el objeto
        _PrintPDF = bPrintPDF ' Exportar a PDF
        _PrintPDF_File = sPrintPDF_File     ' Nombre de fichero a exportar a PDF
        _Impresora = sImpresora
        _Copias = iCopias
    End Sub

    '
    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
               ByVal sWhere As String, _
               ByVal aFormulas As ArrayList, _
               ByVal stSubReports_Where As st_SubReportWhere, _
               Optional ByVal bBringToFront As Boolean = True)

        _Where = sWhere
        _InformeCrystal = oInforme
        _Formulas = aFormulas
        _stSubReports_Where = stSubReports_Where
        _TipoConstructor = 30  ' Flag para saber como se ha instanciado el objeto
        _BringToFront = bBringToFront ' Poner el formulario delante 

    End Sub

    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
               ByVal sOrigenDatos As String, _
               ByVal oDsDatos As DataSet, _
               Optional ByVal bBringToFront As Boolean = True)
        '
        _InformeCrystal = oInforme
        _OrigenDatos = sOrigenDatos
        _DsDatos = oDsDatos
        _TipoConstructor = 40  ' Flag para saber como se ha instanciado el objeto
        _BringToFront = bBringToFront ' Poner el formulario delante 
        '
    End Sub

    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
              ByVal sOrigenDatos As String, _
              ByVal oDsDatos As DataSet, _
              ByVal aFormulas As ArrayList, _
              Optional ByVal bBringToFront As Boolean = True)
        '
        _InformeCrystal = oInforme
        _OrigenDatos = sOrigenDatos
        _DsDatos = oDsDatos
        _TipoConstructor = 41  ' Flag para saber como se ha instanciado el objeto
        _Formulas = aFormulas
        _BringToFront = bBringToFront ' Poner el formulario delante 
        '
    End Sub

    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                   ByVal stSubReports_Where As st_SubReportWhere, _
                   ByVal aFormulas As ArrayList, _
                   ByVal sImpresora As String, _
                   ByVal iCopias As Integer, _
                   Optional ByVal NomInforme As String = "")

        _NomInforme = NomInforme
        _stSubReports_Where = stSubReports_Where
        _InformeCrystal = oInforme
        _Formulas = aFormulas
        _TipoConstructor = 9  ' Flag para saber como se ha instanciado el objeto
        _Copias = iCopias     ' Nº de Copias
        _Impresora = sImpresora

        '
    End Sub

#End Region

#Region "Funciones"

    Private Sub LoadReport()
        '  
        Dim stFormula As st_Formulas
        Dim stSubReportWhere As st_SubReportWhere
        Dim stParametro As st_Parametro
        Dim sPath As String = Application.StartupPath
        Dim ReportPath As String = sPath & "\Reports\" & Me._Informe
        '
        Dim myLogin As New CrystalDecisions.Shared.TableLogOnInfo
        Dim myTable As CrystalDecisions.CrystalReports.Engine.Table
        Dim oReport As CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim crFormulas As FormulaFieldDefinitions
        '
        oReport = _InformeCrystal
        '
        'Dim oReport As new CrystalDecisions.CrystalReports.Engine.ReportDocument
        'oReport.Load(ReportPath)
        '
        myLogin.ConnectionInfo.ServerName = oCompany.Server
        myLogin.ConnectionInfo.UserID = IniGet(sPath & "\SEI_Importacions.ini", "Parametros", "U")     ' Usuario
        myLogin.ConnectionInfo.Password = IniGet(sPath & "\SEI_Importacions.ini", "Parametros", "P")   ' Password
        myLogin.ConnectionInfo.DatabaseName = oCompany.CompanyDB
        '
        '-----------------------------------------------------------------------------------
        ' Conexion Tablas
        '-----------------------------------------------------------------------------------
        For Each myTable In oReport.Database.Tables
            myTable.ApplyLogOnInfo(myLogin)
            'myTable.Location = "@SEI"
            'objReport.Database.Tables("MyTable").SetDataSource(objDataSet.Tables("MyTable"))    
        Next
        '-----------------------------------------------------------------------------------
        ' Formulas
        '-----------------------------------------------------------------------------------
        If Not IsNothing(_Formulas) Then
            crFormulas = oReport.DataDefinition.FormulaFields
            '
            For Each stFormula In _Formulas
                crFormulas.Item(stFormula.Nombre).Text = "'" & stFormula.Valor.ToString & "'"
            Next
        End If
        '
        '-----------------------------------------------------------------------------------
        ' SubReports
        '-----------------------------------------------------------------------------------
        'Crystal Report's report document object
        'Dim objReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        'objReport.VerifyDatabase()

        'Sub report object of crystal report.
        'Dim mySubReportObject As CrystalDecisions.CrystalReports.Engine.SubreportObject

        'Sub report document of crystal report.
        Dim mySubRepDoc As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        '
        For Each mySubRepDoc In oReport.Subreports
            For Each myTable In mySubRepDoc.Database.Tables
                myTable.ApplyLogOnInfo(myLogin)
            Next
            '
            ' Where Subinformes
            '
            If Not IsNothing(_SubReports_Where) Then
                '
                For Each stSubReportWhere In _SubReports_Where
                    If stSubReportWhere.NombreReport = mySubRepDoc.Name Then
                        '-----------------------------------------------------------------------------------
                        ' Formulas
                        '-----------------------------------------------------------------------------------
                        If Not IsNothing(stSubReportWhere.aFormulas) Then
                            crFormulas = mySubRepDoc.DataDefinition.FormulaFields
                            '
                            For Each stFormula In stSubReportWhere.aFormulas
                                crFormulas.Item(stFormula.Nombre).Text = stFormula.Valor.ToString
                            Next
                        End If

                        'mySubRepDoc.DataDefinition.RecordSelectionFormula = stSubReportWhere.ValorWhere
                    End If
                Next
            End If
        Next
        '-----------------------------------------------------------------------------------
        '
        If Not IsNothing(_stSubReports_Where) Then

            If Not IsNothing(_stSubReports_Where.aParametros) Then
                For Each stParametro In _stSubReports_Where.aParametros
                    oReport.SetParameterValue(stParametro.FieldName, stParametro.Value)
                Next
            End If

        End If
        '   
        If Not IsNothing(Me._Where) Then
            If Me._Where <> "" Then
                oReport.DataDefinition.RecordSelectionFormula = Me._Where
            End If
        End If

        If Me._PrintPDF = True Then
            Me.PrintPDF_FilePath = Me._PrintPDF_File
            oReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Me.PrintPDF_FilePath)
        End If
        If NullToText(_Impresora).Trim <> "" Then
            ' ''-->> Si _Impresora = Predeterminada, no informem i agafa la per defecte, si "" no imprimim i sino, la impressora informada
            If _Impresora <> "Predeterminada" Then
                oReport.PrintOptions.PrinterName = _Impresora
            End If
            '
            If Me._Copias <> 0 Then
                oReport.PrintToPrinter(Me._Copias, True, 0, 0)
            Else
                oReport.PrintToPrinter(1, True, 0, 0)
            End If
        End If
        '
        oReport.Close()
        '
    End Sub

    Public Sub Imprimir()
        '
        Dim sError As String = ""
        '

        Application.DoEvents()
        '
        ' Variable para controlar el estado de la impresión
        Dim bPrintOk As Boolean = True
        '
        Try
            '
            LoadReport()
            '
        Catch ex As Exception
            '
            ' Si salta cualquier excepción en el proceso de impresión, ponemos la variable a falso
            bPrintOk = False
            sError = ex.Message
            'Me._ParentAddon.SBO_Application.MessageBox(ex.Message)
            '
        End Try
        '
        ' Mostramos al usuario el resultado de la impresión de la oferta
        '
        If bPrintOk Then
            'SEI_SRV_MAIL.lblmsg.Text = "Documento impreso correctamente"
            Application.DoEvents()
        Else
            'SEI_SRV_MAIL.lblmsg.Text = "Ha ocurrido un error al imprimir el documento. Causa:" & sError
            Application.DoEvents()
        End If
        '
    End Sub

    Public Sub VistaPrevia()
        '
        'SEI_SRV_MAIL.lblmsg.Text = ("Previsualización en curso, un momento por favor...", _
        '                    SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        ''
        '' Variable para controlar el estado de la impresión
        'Dim bPrintOk As Boolean = True
        'Dim oFormVisor As SEI_VisorCrystal
        ''
        'Try
        '    '
        '    Select Case Me._TipoConstructor
        '        '
        '        Case Is = 1
        '            oFormVisor = New SEI_VisorCrystal(Me._ParentAddon, Me._InformeCrystal, Me._Where)
        '            If Me._BringToFront Then
        '                Dim oWindowsSbo As SEI_WindowsSbo
        '                oWindowsSbo = New SEI_WindowsSbo(Me._ParentAddon)
        '                oFormVisor.ShowDialog(oWindowsSbo)
        '            Else
        '                oFormVisor.ShowDialog()
        '                oFormVisor.Dispose()
        '                oFormVisor = Nothing
        '                GC.Collect()
        '                GC.WaitForPendingFinalizers()
        '            End If
        '            '
        '        Case Is = 2
        '            oFormVisor = New SEI_VisorCrystal(Me._ParentAddon, Me._InformeCrystal, Me._Where, Me._Formulas)
        '            If Me._BringToFront Then
        '                Dim oWindowsSbo As SEI_WindowsSbo
        '                oWindowsSbo = New SEI_WindowsSbo(Me._ParentAddon)
        '                oFormVisor.ShowDialog(oWindowsSbo)
        '                oFormVisor = Nothing
        '                GC.Collect()
        '                GC.WaitForPendingFinalizers()
        '            Else
        '                oFormVisor.ShowDialog()
        '                oFormVisor.Dispose()
        '                oFormVisor = Nothing
        '                GC.Collect()
        '                GC.WaitForPendingFinalizers()
        '            End If
        '            '
        '        Case Is = 3
        '            oFormVisor = New SEI_VisorCrystal(Me._ParentAddon, Me._InformeCrystal, Me._Where, Me._Formulas, Me._SubReports_Where)
        '            If Me._BringToFront Then
        '                Dim oWindowsSbo As SEI_WindowsSbo
        '                oWindowsSbo = New SEI_WindowsSbo(Me._ParentAddon)
        '                oFormVisor.ShowDialog(oWindowsSbo)
        '                oFormVisor.Dispose()
        '                oFormVisor = Nothing
        '                GC.Collect()
        '                GC.WaitForPendingFinalizers()
        '            Else
        '                oFormVisor.ShowDialog()
        '                oFormVisor.Dispose()
        '                oFormVisor = Nothing
        '                GC.Collect()
        '                GC.WaitForPendingFinalizers()
        '            End If
        '            '
        '        Case Is = 30
        '            oFormVisor = New SEI_VisorCrystal(Me._ParentAddon, Me._InformeCrystal, Me._Where, Me._Formulas, Me._stSubReports_Where)
        '            If Me._BringToFront Then
        '                Dim oWindowsSbo As SEI_WindowsSbo
        '                oWindowsSbo = New SEI_WindowsSbo(Me._ParentAddon)
        '                oFormVisor.ShowDialog(oWindowsSbo)
        '                Me._InformeCrystal.Close()
        '            Else
        '                oFormVisor.ShowDialog()
        '                Me._InformeCrystal.Close()
        '            End If

        '        Case Is = 40
        '            oFormVisor = New SEI_VisorCrystal(Me._ParentAddon, Me._InformeCrystal, Me._OrigenDatos, Me._DsDatos)
        '            If Me._BringToFront Then
        '                Dim oWindowsSbo As SEI_WindowsSbo
        '                oWindowsSbo = New SEI_WindowsSbo(Me._ParentAddon)
        '                oFormVisor.ShowDialog(oWindowsSbo)
        '                oFormVisor.Dispose()
        '                oFormVisor = Nothing
        '                GC.Collect()
        '                GC.WaitForPendingFinalizers()
        '            Else
        '                oFormVisor.ShowDialog()
        '                oFormVisor.Dispose()
        '                oFormVisor = Nothing
        '                GC.Collect()
        '                GC.WaitForPendingFinalizers()
        '            End If

        '        Case Is = 41
        '            oFormVisor = New SEI_VisorCrystal(Me._ParentAddon, Me._InformeCrystal, Me._OrigenDatos, Me._DsDatos, Me._Formulas)
        '            If Me._BringToFront Then
        '                Dim oWindowsSbo As SEI_WindowsSbo
        '                oWindowsSbo = New SEI_WindowsSbo(Me._ParentAddon)
        '                oFormVisor.ShowDialog(oWindowsSbo)
        '                oFormVisor.Dispose()
        '                oFormVisor = Nothing
        '                GC.Collect()
        '                GC.WaitForPendingFinalizers()
        '            Else
        '                oFormVisor.ShowDialog()
        '                oFormVisor.Dispose()
        '                oFormVisor = Nothing
        '                GC.Collect()
        '                GC.WaitForPendingFinalizers()
        '            End If
        '            '
        '    End Select
        '    '
        'Catch ex As Exception
        '    '
        '    ' Si salta cualquier excepción en el proceso de impresión, ponemos la variable a falso
        '    bPrintOk = False
        '    Me._ParentAddon.SBO_Application.MessageBox(ex.Message)
        '    '
        'End Try
        ''
        '' Mostramos al usuario el resultado de la impresión de la oferta
        ''
        'If bPrintOk Then
        '    SEI_SRV_MAIL.lblmsg.Text = ("Documento visualizado correctamente", _
        '                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        'Else
        '    SEI_SRV_MAIL.lblmsg.Text = ("Ha ocurrido un error al visualizar el documento", _
        '        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        'End If
        ''
    End Sub

#End Region

End Class
