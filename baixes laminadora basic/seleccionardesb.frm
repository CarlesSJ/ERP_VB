VERSION 5.00
Begin VB.Form seleccionardesb 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Triar Desbobinador"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2025
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Desb 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1005
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   75
      Width           =   795
   End
   Begin VB.CommandButton desb1 
      BackColor       =   &H0000FF00&
      Caption         =   "Desb 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   795
   End
End
Attribute VB_Name = "seleccionardesb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
desb = 2
   Unload seleccionardesb
End Sub

Private Sub desb1_Click()
   desb = 1
   Unload seleccionardesb
End Sub
