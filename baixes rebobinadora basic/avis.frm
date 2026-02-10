VERSION 5.00
Begin VB.Form avis 
   Caption         =   "Avis"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   1380
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Vale"
      Height          =   360
      Left            =   2205
      TabIndex        =   1
      Top             =   945
      Width           =   1590
   End
   Begin VB.Label missatge 
      Height          =   720
      Left            =   360
      TabIndex        =   0
      Top             =   135
      Width           =   4830
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
