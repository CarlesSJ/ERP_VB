VERSION 5.00
Begin VB.Form formesperant 
   BackColor       =   &H00EAD9CE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Esperant..."
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10905
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bnoacabat 
      BackColor       =   &H005C31DD&
      Caption         =   "No Acabat - No Muntat"
      Height          =   1455
      Left            =   90
      Picture         =   "formesperant.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3735
      Width           =   1935
   End
   Begin VB.CommandButton bfet 
      BackColor       =   &H0025EFAD&
      Caption         =   "Fet - Muntat"
      Height          =   1455
      Left            =   8820
      Picture         =   "formesperant.frx":08A5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3765
      Width           =   1935
   End
   Begin VB.Label etadhesiu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00ED823A&
      Height          =   1140
      Left            =   2235
      TabIndex        =   4
      Top             =   4080
      Width           =   6435
   End
   Begin VB.Label etcolor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1170
      Left            =   600
      TabIndex        =   1
      Top             =   990
      Width           =   10440
   End
   Begin VB.Label etesperant 
      BackStyle       =   0  'Transparent
      Caption         =   "Esperant foto del micropunt muntat..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00ED823A&
      Height          =   1170
      Left            =   600
      TabIndex        =   0
      Top             =   255
      Width           =   10440
   End
   Begin VB.Image cfotomicropunt 
      Height          =   1815
      Left            =   3810
      Picture         =   "formesperant.frx":10DE
      Stretch         =   -1  'True
      Top             =   2070
      Width           =   2040
   End
End
Attribute VB_Name = "formesperant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bfet_Click()
  Form1.btinterfet_Click cadbl(bfet.tag) - 1
  Unload formesperant
End Sub

Private Sub bnoacabat_Click()
  Unload formesperant
End Sub

