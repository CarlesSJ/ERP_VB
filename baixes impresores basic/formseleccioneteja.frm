VERSION 5.00
Begin VB.Form formseleccioneteja 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4020
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Acceptar"
      Height          =   630
      Left            =   2940
      Picture         =   "formseleccioneteja.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   765
      Width           =   900
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Borrar"
      Height          =   630
      Left            =   2940
      Picture         =   "formseleccioneteja.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Intensa"
      Height          =   630
      Left            =   2010
      Picture         =   "formseleccioneteja.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Normal"
      Height          =   630
      Left            =   1065
      Picture         =   "formseleccioneteja.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H006BEBB1&
      Caption         =   "Ràpida"
      Height          =   630
      Left            =   105
      Picture         =   "formseleccioneteja.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   45
      Width           =   900
   End
   Begin VB.Label etNeteges 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   210
      TabIndex        =   3
      Top             =   795
      Width           =   2700
   End
End
Attribute VB_Name = "formseleccioneteja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   etNeteges.tag = etNeteges.tag + "R"
   etNeteges = etNeteges + IIf(etNeteges <> "", "+", "") + "R"
End Sub

Private Sub Command2_Click()
   etNeteges.tag = etNeteges.tag + "N"
   etNeteges = etNeteges + IIf(etNeteges <> "", "+", "") + "N"
End Sub

Private Sub Command3_Click()
   etNeteges.tag = etNeteges.tag + "I"
   etNeteges = etNeteges + IIf(etNeteges <> "", "+", "") + "I"
End Sub

Private Sub Command4_Click()
   etNeteges.tag = ""
   etNeteges = ""
End Sub

Private Sub Command5_Click()
  formseleccioneteja.Hide
End Sub

