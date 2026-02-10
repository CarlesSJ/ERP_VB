VERSION 5.00
Begin VB.Form avis2 
   Caption         =   "Avis"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   ControlBox      =   0   'False
   Icon            =   "avis2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botoxinxeta 
      Height          =   360
      Left            =   3825
      Picture         =   "avis2.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Evitar que es tanqui automàticament la finestra."
      Top             =   945
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox ctxtmissatge 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   60
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5550
      Top             =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Vale"
      Height          =   360
      Left            =   2205
      TabIndex        =   1
      Top             =   945
      Width           =   1590
   End
   Begin VB.Label missatge 
      Alignment       =   2  'Center
      Height          =   720
      Left            =   360
      TabIndex        =   0
      Top             =   135
      Width           =   5625
   End
End
Attribute VB_Name = "avis2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub botoxinxeta_Click()
   If botoxinxeta.BackColor = &H80FFFF Then
      botoxinxeta.BackColor = &H8000000F
       Else: botoxinxeta.BackColor = &H80FFFF
   End If
End Sub

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Timer1_Timer()
   While avis.caption = "Duplicant..."
     avis.SetFocus
     DoEvents
  Wend
End Sub
