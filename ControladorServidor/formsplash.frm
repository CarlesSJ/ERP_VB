VERSION 5.00
Begin VB.Form formsplash 
   BackColor       =   &H005C31DD&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9420
   ClipControls    =   0   'False
   Icon            =   "formsplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   270
      Top             =   195
   End
   Begin VB.Label eterror 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5910
      Left            =   450
      TabIndex        =   0
      Top             =   1185
      Width           =   8550
   End
End
Attribute VB_Name = "formsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -- Api SetForegroundWindow Para traer la ventana al frente
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub Timer1_Timer()
   Call SetForegroundWindow(Me.hwnd)
   
End Sub
