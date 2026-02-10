VERSION 5.00
Begin VB.Form teclattactil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   9435
   Visible         =   0   'False
   Begin VB.Label lletraenter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "   Enter"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   8130
      TabIndex        =   3
      Top             =   1215
      Width           =   1035
   End
   Begin VB.Label lletrabs 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<---"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8115
      TabIndex        =   2
      Top             =   735
      Width           =   1035
   End
   Begin VB.Label lletraesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ESC"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8175
      TabIndex        =   1
      Top             =   90
      Width           =   1035
   End
   Begin VB.Image teclaenter 
      Height          =   1305
      Left            =   8070
      Picture         =   "teclattactil.frx":0000
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   1245
   End
   Begin VB.Image teclabs 
      Height          =   615
      Left            =   8070
      Picture         =   "teclattactil.frx":03DE
      Stretch         =   -1  'True
      Top             =   675
      Width           =   1245
   End
   Begin VB.Image teclaesc 
      Height          =   615
      Left            =   8085
      Picture         =   "teclattactil.frx":07BC
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1245
   End
   Begin VB.Label lletra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   150
      Width           =   675
   End
   Begin VB.Image tecles 
      Height          =   615
      Index           =   0
      Left            =   270
      Picture         =   "teclattactil.frx":0B9A
      Top             =   105
      Width           =   645
   End
End
Attribute VB_Name = "teclattactil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOZORDER = &H4
    Const SWP_NOREDRAW = &H8
    Const SWP_NOACTIVATE = &H10
    Const SWP_FRAMECHANGED = &H20
    Const SWP_SHOWWINDOW = &H40
    Const SWP_HIDEWINDOW = &H80
    Const SWP_NOCOPYBITS = &H100
    Const SWP_NOOWNERZORDER = &H200
    Const SWP_DRAWFRAME = SWP_FRAMECHANGED
    Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
    Const HWND_TOP = 0
    Const HWND_BOTTOM = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
Private Sub Form_Activate()
  Dim pini As Double
  Dim qwertymin As String
  Dim posicio As Double
  qwertymin = "qwertyuiopasdfghjklñzxcvbnm,.-"
  qwertymay = "QWERTYUIOPASDFGHJKLÑZXCVBNM;:_"
  If r = "numeric" Then r = "": qwertymay = "1234567890.QWERTYUIOPASDFGHJKLÑZXCVBNM;:_"
  If Me.Tag <> "" Then Exit Sub
  Me.Tag = "1"
  Me.Visible = True
  Me.Top = 10000
  
  posicio = campcontrol.Top
  If campcontrol.Container.Name <> "form1" Then
     posicio = campcontrol.Container.Top + campcontrol.Top
  End If
  'MsgBox posicio
  'If posicio > 4500 Then
  '   i = SetWindowPos(hwnd, HWND_TOPMOST, posicio * 100, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
  '    Else: i = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
  'End If
  
  pini = tecles(0).Left
  pini = pini - tecles(0).Width
t = tecles(0).Top
 l = tecles(0).Left
 lletra(0).Caption = Mid(qwertymay, 1, 1)
 For i = 1 To 40
  If i = 11 Or i = 21 Or i = 31 Then
     t = t + tecles(0).Height + 10
     l = pini + ((tecles(0).Width / 2) * (i / 10))
       Else: l = tecles(i - 1).Left
   End If
   'faig la tecla
   Load tecles(i)
   tecles(i).Top = t
   tecles(i).Width = tecles(0).Width
   tecles(i).Height = tecles(0).Height
   tecles(i).Left = l + tecles(0).Width
   tecles(i).Visible = True
   'faig la lletra
   Load lletra(i)
   lletra(i).Top = t + 45
   lletra(i).Width = lletra(0).Width
   tecles(i).Height = lletra(0).Height
   lletra(i).Left = l + lletra(0).Width
   lletra(i).Caption = Mid(qwertymay, i + 1, 1)
    lletra(i).ZOrder 0
   lletra(i).Visible = True

    
  
 Next i
  Me.Width = tecles(i - 1).Left + tecles(i - 1).Width + 2000
  Me.Height = tecles(i - 1).Top + tecles(i - 1).Height + 1500
  Me.Caption = "TECLAT QWERTY"
 Load tecles(i)
 tecles(i).Top = tecles(i - 1).Top + tecles(i - 1).Height
 tecles(i).Left = tecles(0).Width + (tecles(i).Width * 1.5)
 tecles(i).Width = tecles(i).Width * 7
 tecles(i).Stretch = True
 tecles(i).Visible = True
  Load lletra(i)
 lletra(i).AutoSize = False
 lletra(i).Top = tecles(i).Top + 45
 lletra(i).Left = tecles(i).Left
 lletra(i).Width = tecles(i).Width
 lletra(i).Caption = " "
 lletra(i).ZOrder 0
 lletra(i).Visible = True
 Me.Visible = True

If posicio > 4500 Then
       teclattactil.Top = Form1.Top + 50
     Else
     teclattactil.Top = Form1.Top + 4500
  End If
  teclattactil.Left = Form1.Left + 100

End Sub


Private Sub Label1_Click()

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  tecles(Index).Left = tecles(Index).Left + 30
  tecles(Index).Top = tecles(Index).Top + 30
  lletra(Index).Left = lletra(Index).Left + 30
  lletra(Index).Top = lletra(Index).Top + 30
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  tecles(Index).Left = tecles(Index).Left + 30
  tecles(Index).Top = tecles(Index).Top + 30
  lletra(Index).Left = lletra(Index).Left + 30
  lletra(Index).Top = lletra(Index).Top + 30
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
End Sub

Private Sub lletra_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  lletra(Index).Tag = "1"
  tecles(Index).Left = tecles(Index).Left + 30
  tecles(Index).Top = tecles(Index).Top + 30
  lletra(Index).Left = lletra(Index).Left + 30
  lletra(Index).Top = lletra(Index).Top + 30

 ' For j = 1 To 900000
 '   j = j + 1
 '   j = j - 1
 ' Next j
 
  
End Sub



Private Sub lletra_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If lletra(Index).Tag = "1" Then
 tecles(Index).Left = tecles(Index).Left - 30
  tecles(Index).Top = tecles(Index).Top - 30
  lletra(Index).Left = lletra(Index).Left - 30
  lletra(Index).Top = lletra(Index).Top - 30
  On Error Resume Next
  campcontrol.Text = campcontrol.Text + lletra(Index).Caption
  lletra(Index).Tag = ""
End If
End Sub

Private Sub lletrabs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lletrabs.Tag = "1"
lletrabs.Left = lletrabs.Left + 30
  lletrabs.Top = lletrabs.Top + 30
  teclabs.Left = teclabs.Left + 30
 teclabs.Top = teclabs.Top + 30
End Sub

Private Sub lletrabs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lletrabs.Tag = "1" Then
  teclabs.Left = teclabs.Left - 30
  teclabs.Top = teclabs.Top - 30
  lletrabs.Left = lletrabs.Left - 30
  lletrabs.Top = lletrabs.Top - 30
  lletrabs.Tag = ""
  If Len(campcontrol.Text) > 0 Then campcontrol.Text = Mid(campcontrol.Text, 1, Len(campcontrol.Text) - 1)
End If
End Sub

Private Sub lletraenter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lletraenter.Tag = "1"
lletraenter.Left = lletraenter.Left + 30
  lletraenter.Top = lletraenter.Top + 30
  teclaenter.Left = teclaenter.Left + 30
 teclaenter.Top = teclaenter.Top + 30
End Sub

Private Sub lletraenter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lletraenter.Tag = "1" Then
  teclaenter.Left = teclaenter.Left - 30
  teclaenter.Top = teclaenter.Top - 30
  lletraenter.Left = lletraenter.Left - 30
  lletraenter.Top = lletraenter.Top - 30
  lletraenter.Tag = ""
  Unload teclattactil
  SendKeys "{TAB}"
End If
End Sub

Private Sub lletraesc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lletraesc.Tag = "1"
lletraesc.Left = lletraesc.Left + 30
  lletraesc.Top = lletraesc.Top + 30
  teclaesc.Left = teclaesc.Left + 30
 teclaesc.Top = teclaesc.Top + 30
End Sub

Private Sub lletraesc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If lletraesc.Tag = "1" Then
  teclaesc.Left = teclaesc.Left - 30
  teclaesc.Top = teclaesc.Top - 30
  lletraesc.Left = lletraesc.Left - 30
  lletraesc.Top = lletraesc.Top - 30
  On Error Resume Next
  campcontrol.Text = ""
  lletraesc.Tag = ""
End If
End Sub

Private Sub tecles_Click(Index As Integer)
'campcontrol.Text = campcontrol.Text + Trim(Index)
End Sub
