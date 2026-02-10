VERSION 5.00
Begin VB.Form missatgevist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Missatge Previsualitzacio de Comanda"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "missatgevist.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   375
      Top             =   135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Obrir aquest informe tarda uns segons i si ja l'has vist prem aquest botó."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   690
      TabIndex        =   0
      Top             =   675
      Width           =   3165
   End
End
Attribute VB_Name = "missatgevist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   escriure_ini "Baixes", "imprimircomanda", "0", "comandes.ini"
   Unload missatgevist
   assignardecimalipunt
End Sub

Private Sub Timer1_Timer()
  On Error Resume Next
  AppActivate "Imprimint comanda"
End Sub
