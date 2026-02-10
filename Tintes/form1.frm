VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox comanda 
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   315
      Width           =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function mirarsihihaCingularReal(vnumtreball As Double, vordremodificacio As Double) As Boolean
   Dim vurl As String
   Dim generarfitxer_pdf As String
   generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(vnumtreball, "00000") + "\pdf" + Format(vnumtreball, "00000") + "-" + Format(vordremodificacio, "000") + "_CR.pdf"
   If existeix(generarfitxer_pdf) Then
      mirarsihihaCingularReal = True
   End If
   
   
End Function

Sub preparaelPDF(vnomfitxerpdf As String, vrotacio As Double, vMirall As String)
  If Not existeix(vnomfitxerpdf) Then Exit Sub
  vMirall = UCase(vMirall) 'vMirall si es V es vertical H es horitzotal
  If existeix("c:\temp\pdfimpresio.gif") Then Kill "c:\temp\pdfimpresio.gif"
  ConvertirFormats vnomfitxerpdf, "c:\temp\pdfimpresio.gif", 50
  If Not existeix("c:\temp\pdfimpresio.gif") Then GoTo fi
  If vMirall = "H" Then InvertirHVImatge "c:\temp\pdfimpresio.gif", "c:\temp\pdfimpresio.gif"
  If vMirall = "V" Then InvertirHVImatge "c:\temp\pdfimpresio.gif", "c:\temp\pdfimpresio.gif", True
  If vrotacio > 0 Then RotarImatge "c:\temp\pdfimpresio.gif", "c:\temp\pdfimpresio.gif", vrotacio
fi:
End Sub

Function comandavalida(numc As Double, msg As String, Optional nocomprovarllista As Boolean) As Boolean

End Function
