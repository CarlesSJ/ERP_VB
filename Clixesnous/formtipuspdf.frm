VERSION 5.00
Begin VB.Form formtipuspdf 
   BackColor       =   &H00EAD9CE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipus de PDF"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12420
   Icon            =   "formtipuspdf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   12420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Height          =   315
      Left            =   660
      Picture         =   "formtipuspdf.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Enviar document de Revisió del pdf prèvi."
      Top             =   30
      Width           =   315
   End
   Begin VB.CommandButton eliminar_PRSC 
      Height          =   360
      Left            =   4515
      Picture         =   "formtipuspdf.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Eliminacio del PDF Normal"
      Top             =   1080
      Width           =   360
   End
   Begin VB.CommandButton botopdfprevisc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pdf Prèvi Sep. Colors"
      Height          =   1095
      Left            =   2895
      Picture         =   "formtipuspdf.frx":0E1E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   345
      Width           =   1620
   End
   Begin VB.CommandButton botopdfprevi 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pdf Prèvi"
      Height          =   1095
      Left            =   630
      Picture         =   "formtipuspdf.frx":3890
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   360
      Width           =   1620
   End
   Begin VB.CommandButton eliminar_pr 
      Height          =   360
      Left            =   2250
      Picture         =   "formtipuspdf.frx":50BA
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Eliminacio del PDF Normal"
      Top             =   1095
      Width           =   360
   End
   Begin VB.CommandButton Command1 
      Height          =   330
      Left            =   5295
      Picture         =   "formtipuspdf.frx":5644
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Copiar pdf de la versió anterior"
      Top             =   30
      Width           =   315
   End
   Begin VB.CommandButton eliminar_cingular 
      Height          =   360
      Left            =   11085
      Picture         =   "formtipuspdf.frx":5BCE
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Eliminacio del PDF Cingular Real2"
      Top             =   1080
      Width           =   360
   End
   Begin VB.CommandButton eliminar_sep 
      Height          =   360
      Left            =   9015
      Picture         =   "formtipuspdf.frx":6158
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminacio del PDF Separació de colors."
      Top             =   1095
      Width           =   360
   End
   Begin VB.CommandButton eliminar_pdf 
      Height          =   360
      Left            =   6885
      Picture         =   "formtipuspdf.frx":66E2
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminacio del PDF Normal"
      Top             =   1095
      Width           =   360
   End
   Begin VB.CommandButton botopdfcingular 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cingular Real2"
      Height          =   1095
      Left            =   9465
      Picture         =   "formtipuspdf.frx":6C6C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   345
      Width           =   1620
   End
   Begin VB.CommandButton botopdfcapes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pdf Sep. Colors"
      Height          =   1095
      Left            =   7395
      Picture         =   "formtipuspdf.frx":9EB6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1620
   End
   Begin VB.CommandButton botopdfnormal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pdf Fotogravador"
      Height          =   1095
      Left            =   5265
      Picture         =   "formtipuspdf.frx":C928
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1620
   End
   Begin VB.Image imatgecapesprevioff 
      Height          =   900
      Left            =   2445
      Picture         =   "formtipuspdf.frx":E152
      Top             =   15
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imatgeprevioff 
      Height          =   765
      Left            =   0
      Picture         =   "formtipuspdf.frx":E6B9
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imatgecingularoff 
      Height          =   960
      Left            =   6915
      Picture         =   "formtipuspdf.frx":EDD7
      Top             =   15
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imatgenormaloff 
      Height          =   765
      Left            =   45
      Picture         =   "formtipuspdf.frx":10421
      Top             =   885
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imatgecapesoff 
      Height          =   900
      Left            =   4815
      Picture         =   "formtipuspdf.frx":11C4B
      Top             =   30
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "formtipuspdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub botopdfcapes_Click()
    formtipuspdf.tag = "SC"
    formtipuspdf.Hide
End Sub

Private Sub botopdfcingular_Click()
 formtipuspdf.tag = "CR"
    formtipuspdf.Hide
End Sub

Private Sub botopdfnormal_Click()
    formtipuspdf.tag = "N"
    formtipuspdf.Hide
End Sub

Private Sub botopdfprevi_Click()
 formtipuspdf.tag = "PR"
    formtipuspdf.Hide
End Sub

Private Sub botopdfprevisc_Click()
    formtipuspdf.tag = "PRSC"
    formtipuspdf.Hide
End Sub

Private Sub Command1_Click()
   formtipuspdf.tag = "N"
   formtipuspdf.Hide
End Sub

Private Sub Command2_Click()
  formclixes.enviar_revisio_previ
  
End Sub


Private Sub Command3_Click()
 
End Sub

Private Sub eliminar_cingular_Click()
   treure_vincle_pdf ("cingular")
End Sub

Private Sub eliminar_pdf_Click()
    treure_vincle_pdf ("pdf")
End Sub

Private Sub eliminar_pr_Click()
   treure_vincle_pdf ("previ")
End Sub

Private Sub eliminar_PRSC_Click()
 treure_vincle_pdf ("previSC")
End Sub

Private Sub eliminar_sep_Click()
    treure_vincle_pdf ("separacio")
End Sub
Sub treure_vincle_pdf(v As String)
    formtipuspdf.tag = "Borrar_" + v
    formtipuspdf.Hide
End Sub
Private Sub Form_Activate()
    botopdfnormal.Enabled = True
    botopdfcapes.Enabled = True
    botopdfcingular.Enabled = True
    botopdfprevi.Enabled = True
    botopdfprevisc.Enabled = True
    DoEvents
    If Not existeix(formclixes.rutapdftreball(, False)) Then botopdfnormal.Picture = imatgenormaloff.Picture
    If Not existeix(formclixes.rutapdftreball(, True)) Then botopdfcapes.Picture = imatgecapesoff.Picture
    If Not existeix(formclixes.rutapdftreball(, True, True)) Then botopdfcingular.Picture = imatgecingularoff.Picture
    If Not existeix(formclixes.rutapdftreball(, , , True)) Then botopdfprevi.Picture = imatgeprevioff.Picture
    If Not existeix(formclixes.rutapdftreball(, , , , True)) Then botopdfprevisc.Picture = imatgecapesprevioff.Picture
    DoEvents
    If modificartintes Then
     eliminar_pr.visible = False
     eliminar_PRSC.visible = False
     eliminar_cingular.visible = False
     eliminar_pdf.visible = False
     eliminar_sep.visible = False
    End If
    
End Sub

