VERSION 5.00
Begin VB.Form formmsgbox 
   ClientHeight    =   4515
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8205
   Icon            =   "formmsgbox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "D'Acord"
      Height          =   780
      Left            =   3060
      TabIndex        =   0
      Top             =   3630
      Width           =   1620
   End
   Begin VB.Label etiqueta 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   105
      TabIndex        =   1
      Top             =   120
      Width           =   7980
   End
End
Attribute VB_Name = "formmsgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Unload formmsgbox
End Sub
