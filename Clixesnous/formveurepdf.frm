VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form formveurepdf 
   Caption         =   "Visor PDF"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7005
      _cx             =   5080
      _cy             =   5080
   End
End
Attribute VB_Name = "formveurepdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
  ajustartamanyform
End Sub
Sub ajustartamanyform()
  AcroPDF1.width = Me.width - 100
  AcroPDF1.Height = Me.Height - 100
  AcroPDF1.Top = 100
  AcroPDF1.Left = 100
End Sub
