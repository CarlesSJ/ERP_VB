VERSION 5.00
Begin VB.Form Formplantillarevisio 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   11340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   Picture         =   "formPlantillaRevisio.frx":0000
   ScaleHeight     =   11340
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Check1"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   2100
      TabIndex        =   3
      Top             =   1260
      Width           =   195
   End
   Begin VB.Label etmarcailinia 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1290
      TabIndex        =   2
      Top             =   495
      Width           =   3690
   End
   Begin VB.Label ettreball 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6525
      TabIndex        =   1
      Top             =   300
      Width           =   4110
   End
   Begin VB.Label etclient 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   780
      TabIndex        =   0
      Top             =   285
      Width           =   4110
   End
End
Attribute VB_Name = "Formplantillarevisio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
