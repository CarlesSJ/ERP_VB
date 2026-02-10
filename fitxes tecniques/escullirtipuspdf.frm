VERSION 5.00
Begin VB.Form formescullirpdf 
   BackColor       =   &H00EEE4D7&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5430
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00ED823A&
      Caption         =   "Pdf Seguretat"
      Height          =   960
      Left            =   3915
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   30
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F3B378&
      Caption         =   "Pdf Fitxa Tècnica"
      Height          =   960
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   30
      Width           =   1485
   End
   Begin VB.CommandButton botopdfconformitat 
      BackColor       =   &H00DBBDAA&
      Caption         =   "Pdf Conformitat"
      Height          =   960
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   30
      Width           =   1485
   End
End
Attribute VB_Name = "formescullirpdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub botopdfconformitat_Click()
  formescullirpdf.Tag = "conformitat"
  Me.Hide
End Sub

Private Sub Command1_Click()
 formescullirpdf.Tag = "fitxatecnica"
   Me.Hide
End Sub

Private Sub Command2_Click()
  formescullirpdf.Tag = "seguretat"
  Me.Hide
End Sub

Private Sub Command3_Click()
   Me.Hide
End Sub
