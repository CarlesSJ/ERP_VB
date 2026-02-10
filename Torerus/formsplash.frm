VERSION 5.00
Begin VB.Form formsplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "formsplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   150
      Top             =   300
   End
   Begin VB.Label etiqueta 
      Caption         =   "Carregant..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1140
      TabIndex        =   0
      Top             =   2565
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   3300
      Left            =   675
      Picture         =   "formsplash.frx":048A
      Top             =   -210
      Width           =   3300
   End
End
Attribute VB_Name = "formsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   If App.PrevInstance Then MsgBox "El programa ja està obert." + Chr(10) + "SI NO VEUS EL PROGRAMA PROVA DE REINICIAR LA TABLET.", vbCritical, "Atenció": End
   Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Static cont As Byte
    Static cont2 As Byte
    If cont = 0 Then FormTorerus.Show
    etiqueta = "Carregant" + String(cont, ".")
    cont = cont + 1
    cont2 = cont2 + 1
    On Error Resume Next
     formsplash.SetFocus
    If cont = 5 Then cont = 1
    If cont2 = 10 Then Unload formsplash
End Sub
