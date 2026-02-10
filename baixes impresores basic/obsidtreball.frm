VERSION 5.00
Begin VB.Form obsidtreball 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Observació Id Treball"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton bnotepad 
      BackColor       =   &H00FFC0C0&
      Height          =   330
      Left            =   4635
      Picture         =   "obsidtreball.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   75
      Width           =   315
   End
   Begin VB.TextBox obsid 
      Height          =   3030
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   15
      Width           =   4590
   End
End
Attribute VB_Name = "obsidtreball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bnotepad_Click()
    Clipboard.Clear
    Clipboard.SetText obsid.text
    
    Shell "notepad.exe", vbNormalFocus
    'Send the keys CTRL+V To Notepad (i.e the window that has focus)
    Sendkeys "^V"
    wait 2
    Clipboard.Clear
End Sub

Private Sub Form_Activate()
 r = ""
 obsidtreball.obsid.SetFocus
 If Me.Left = 0 Then
     Me.Left = (Screen.width / 2) - (Me.width / 2)
     Me.Top = (Screen.Height / 2) - (Me.Height / 2)
 End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 110 Then KeyCode = 188
End Sub

Private Sub Form_Load()
 'obsidtreball.obsid.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
  r = obsid.text
End Sub
