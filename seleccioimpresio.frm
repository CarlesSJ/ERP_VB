VERSION 5.00
Begin VB.Form seleccioimpresio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selecció Impresió"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   Icon            =   "seleccioimpresio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "D'acord"
      Height          =   390
      Left            =   1140
      TabIndex        =   0
      Top             =   2205
      Width           =   1770
   End
   Begin VB.Frame Frame1 
      Height          =   1860
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   3495
      Begin VB.CheckBox Checkimps 
         Caption         =   "[Comanda] + Imps"
         Height          =   210
         Left            =   1425
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.OptionButton imprimir 
         Caption         =   "Comandes marcades per impresores."
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   5
         Top             =   1470
         Width           =   3240
      End
      Begin VB.OptionButton imprimir 
         Caption         =   "Comanda actual per pantalla"
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   4
         Top             =   960
         Width           =   2550
      End
      Begin VB.OptionButton imprimir 
         Caption         =   "Comandes"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   210
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton imprimir 
         Caption         =   "Comanda actual per impresora"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   585
         Width           =   2565
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000A&
         Height          =   1080
         Left            =   135
         Top             =   180
         Width           =   2940
      End
   End
End
Attribute VB_Name = "seleccioimpresio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Option1_Click()

End Sub

Private Sub Command1_Click()
  Me.Hide
  Me.Tag = "1"
End Sub

Private Sub Form_Activate()
  Me.Tag = "0"
End Sub

