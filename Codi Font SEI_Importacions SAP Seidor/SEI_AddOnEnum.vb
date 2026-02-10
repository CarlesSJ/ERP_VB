Imports System.Collections
'
Public Class SEI_AddOnEnum
#Region "Estructura VisorCrystal Formulas"

    Public Enum enSBO_LoadFormTypes
        XmlFile = 0       ' Formularios .srf (xml) 
        LogicOnly = 1     ' Formulario de Sap 
        GuiByCode = 3
    End Enum

    Public Structure st_Formulas

        Public Nombre As String
        Public Valor As String

    End Structure

    Public Structure st_SubReportWhere

        Public NombreReport As String
        Public ValorWhere As String
        Public aParametros As ArrayList
        Public aFormulas As ArrayList
        Public aParametrosSubReports As ArrayList

    End Structure


    Public Structure st_Parametro

        Public FieldName As String
        Public ParameterRange As CrystalDecisions.Shared.ParameterRangeValue
        Public Value As String

    End Structure

    Public Enum eCrystal
        Incrustado = 0
        EnDirectorio = 1
        CampoBlob = 2
    End Enum

    ' En el proceso de facturación Electronica
    Public Enum eEnviarFE
        EnDirectorio = 0
        Mail = 1
        Impresora = 2
    End Enum

#End Region

End Class

