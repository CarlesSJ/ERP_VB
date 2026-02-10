Option Explicit On
'
Imports SEI_Importacions.SEI_AddOnEnum
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Text
Imports System.Collections
Imports System.Data

Public Class SEI_VisorCrystal
    '
    Protected _BringToFront As Boolean  ' Poner el formulario delante 
    Protected _Where As String
    Protected _Informe As String
    Protected _Formulas As ArrayList
    Protected _SubReports_Where As ArrayList
    Protected _stSubReports_Where As st_SubReportWhere
    Protected _InformeCrystal As CrystalDecisions.CrystalReports.Engine.ReportDocument
    Protected _OrigenDatos As String
    Protected _DsDatos As dataset
    '
#Region "Constructor"
    '
    Public Sub New()
        '
        '
        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()
        '
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        '
    End Sub

    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal sWhere As String)
        '
        _Where = sWhere
        _InformeCrystal = oInforme
        '
        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()
        '
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        '
    End Sub

    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
                  ByVal sWhere As String, _
                  ByVal aFormulas As ArrayList)
        '
        _Where = sWhere
        _InformeCrystal = oInforme
        _Formulas = aFormulas
        '
        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()
        '
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        '
    End Sub

    Public Sub New(ByVal sInforme As String, _
                   ByVal sWhere As String, _
                   ByVal aFormulas As ArrayList, _
                   ByVal aSubReports_Where As ArrayList)
        '
        _Where = sWhere
        _Informe = sInforme
        _Formulas = aFormulas
        _SubReports_Where = aSubReports_Where
        '
        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()
        '
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        '
    End Sub

    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
               ByVal sWhere As String, _
               ByVal aFormulas As ArrayList, _
               ByVal aSubReports_Where As ArrayList)
        '
        _Where = sWhere
        _InformeCrystal = oInforme
        _Formulas = aFormulas
        _SubReports_Where = aSubReports_Where
        '
        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()
        '
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        '
    End Sub
    '
    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
               ByVal sWhere As String, _
               ByVal aFormulas As ArrayList, _
               ByVal stSubReports_Where As st_SubReportWhere)

        _Where = sWhere
        _InformeCrystal = oInforme
        _Formulas = aFormulas
        _stSubReports_Where = stSubReports_Where

        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub
    '
    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
               ByVal sOrigenDatos As String, _
               ByVal oDsDatos As DataSet)
        '
        _InformeCrystal = oInforme
        _OrigenDatos = sOrigenDatos
        _DsDatos = oDsDatos
        '
        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()
        '
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        '
    End Sub

    Public Sub New(ByRef oInforme As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
              ByVal sOrigenDatos As String, _
              ByVal oDsDatos As DataSet, _
              ByVal aFormulas As ArrayList)
        '
        _InformeCrystal = oInforme
        _OrigenDatos = sOrigenDatos
        _DsDatos = oDsDatos
        _Formulas = aFormulas
        '
        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()
        '
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        '
    End Sub


#End Region

#Region "Funciones"

    Private Sub ConfigureCrystalReports()
        '
        '  
        Select Case Me._OrigenDatos
            Case ""
                ConfigureCrystalReports_SQL()
            Case "SQL"
                ConfigureCrystalReports_SQL()
            Case "XML"
                ConfigureCrystalReports_XML()
        End Select
        '
    End Sub
    '
    Private Sub ConfigureCrystalReports_SQL()
        Dim stFormula As st_Formulas
        Dim stSubReportWhere As st_SubReportWhere
        Dim stParametro As st_Parametro
        Dim sPath As String = Application.StartupPath
        '
        Dim myLogin As New CrystalDecisions.Shared.TableLogOnInfo
        Dim myTable As CrystalDecisions.CrystalReports.Engine.Table
        Dim oReport As CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim crFormulas As FormulaFieldDefinitions
        '
        oReport = _InformeCrystal '
        '
        myLogin.ConnectionInfo.ServerName = oCompany.Server
        myLogin.ConnectionInfo.UserID = IniGet(sPath & "\S_SEI_ATLANTIS_MAIL.ini", "Parametros", "U")    ' Usuario
        myLogin.ConnectionInfo.Password = IniGet(sPath & "\S_SEI_ATLANTIS_MAIL.ini", "Parametros", "P")  ' Password
        myLogin.ConnectionInfo.DatabaseName = oCompany.CompanyDB
        myLogin.ConnectionInfo.Type = ConnectionInfoType.CRQE
        '
        '-----------------------------------------------------------------------------------
        ' Conexion Tablas
        '-----------------------------------------------------------------------------------
        For Each myTable In oReport.Database.Tables
            myTable.ApplyLogOnInfo(myLogin)
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

        Else
            oReport.DataDefinition.RecordSelectionFormula = Me._Where
        End If
        '
        CrystalReportViewer.ShowExportButton = True
        CrystalReportViewer.ReportSource = oReport
        '
    End Sub
    '
    Private Sub ConfigureCrystalReports_XML()
        '
        Dim oReport As CrystalDecisions.CrystalReports.Engine.ReportDocument

        oReport = _InformeCrystal
        oReport.SetDataSource(Me._DsDatos)

        CrystalReportViewer.ShowExportButton = True
        CrystalReportViewer.ReportSource = oReport
        '
    End Sub
    '
    Protected Overrides Sub Finalize()
        '
        MyBase.Finalize()
        '
    End Sub

    Private Sub CrystalReportViewer_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles CrystalReportViewer.KeyPress
        '
        MessageBox.Show(("Form.KeyPress: '" + _
            e.KeyChar.ToString() + "' pressed."))
        '
    End Sub

    Private Sub CrystalReportViewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer.Load
        '
        ConfigureCrystalReports()
        Me.WindowState = FormWindowState.Maximized
        '
    End Sub

    Public Sub Mostrar()
        '
        'SEI_Importacions.lblmsg.Text = "Previsualización en curso, un momento por favor..."
        Application.DoEvents()

        ' Mostramos el formulario
        Me.ShowDialog()
        '
    End Sub

    Private Sub SEI_VisorCrystal_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        '
        MessageBox.Show(("Form.KeyPress: '" + _
            e.KeyChar.ToString() + "' pressed."))

        'MsgBox("Informacio", MsgBoxStyle.Information, "tecla")
        '
        'Me._ParentAddon.SBO_Application.Forms.Item("").Select()
        '
    End Sub

#End Region

End Class