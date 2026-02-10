VERSION 5.00
Begin VB.Form formimprimintetiqueta 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6600
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8325
      Top             =   3630
   End
   Begin VB.Label etnumetiqueta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1860
      Left            =   255
      TabIndex        =   1
      Top             =   3810
      Width           =   9165
   End
   Begin VB.Label etmissatge 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Imprimint l'etiqueta per la llauna"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   300
      TabIndex        =   0
      Top             =   825
      Width           =   9165
   End
End
Attribute VB_Name = "formimprimintetiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Sub Timer1_Timer()
    Static contador As Byte
    contador = contador + 1
    SetForegroundWindow formimprimintetiqueta.hwnd
    If contador > 9 Then
       contador = 0
       On Error Resume Next
       formimprimintetiqueta.Hide
       Unload formimprimintetiqueta
    End If
End Sub
