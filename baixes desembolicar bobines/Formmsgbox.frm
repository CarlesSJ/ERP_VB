VERSION 5.00
Begin VB.Form Formmsgbox 
   Caption         =   "Missatge"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   2565
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Acceptar"
      Height          =   465
      Left            =   3165
      TabIndex        =   1
      Top             =   1935
      Width           =   2520
   End
   Begin VB.Label etmissatge 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005C31DD&
      Height          =   1380
      Left            =   285
      TabIndex        =   0
      Top             =   315
      Width           =   9120
   End
End
Attribute VB_Name = "Formmsgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
