VERSION 5.00
Begin VB.Form avis 
   Caption         =   "Avis"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   Icon            =   "avis.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1380
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H006BEBB1&
      Caption         =   "D'acord"
      Height          =   420
      Left            =   555
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   900
      Width           =   4905
   End
   Begin VB.Label missatge 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   135
      Width           =   45
   End
End
Attribute VB_Name = "avis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Activate()
   avis.Height = missatge.Height + 1500
   avis.width = missatge.width + 1000
   Command1.Top = avis.Height - 1000
   Command1.width = missatge.width
   
   If UCase(App.EXEName) <> "MANTENIMENT TINTES" Then
    avis.Top = (Form1.Height / 2 - avis.Height / 2)
    avis.Left = (Form1.width / 2 - avis.width / 2)
   End If
End Sub

