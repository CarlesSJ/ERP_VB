VERSION 5.00
Begin VB.Form formactualitzant 
   BorderStyle     =   0  'None
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   6450
      Left            =   -150
      TabIndex        =   0
      Top             =   -195
      Width           =   9330
      Begin VB.Timer Timer1 
         Interval        =   400
         Left            =   330
         Top             =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No et moguis d'on ets mentres actualitzes"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   48
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4440
         Left            =   210
         TabIndex        =   2
         Top             =   1575
         Width           =   8955
      End
      Begin VB.Image Image1 
         Height          =   3585
         Left            =   1005
         Picture         =   "formactualitzant.frx":0000
         Top             =   1995
         Width           =   6915
      End
      Begin VB.Label etconnectant 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actualitzant..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1245
         Left            =   465
         TabIndex        =   1
         Top             =   240
         Width           =   8415
      End
   End
End
Attribute VB_Name = "formactualitzant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vnoprimerpla As Boolean
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub etconnectant_DblClick()
vnoprimerpla = Not vnoprimerpla
End Sub

Private Sub Frame1_DblClick()
  vnoprimerpla = Not vnoprimerpla
End Sub

Private Sub Image1_DblClick()
vnoprimerpla = Not vnoprimerpla
End Sub

Private Sub Timer1_Timer()
  Image1.Left = 1100 + (Int((100 * Rnd) + 1))
  Image1.Top = 2000 + (Int((100 * Rnd) + 1))
  DoEvents
  Label1.Visible = Not Label1.Visible
  If Not vnoprimerpla Then SetForegroundWindow formactualitzant.hwnd
End Sub
