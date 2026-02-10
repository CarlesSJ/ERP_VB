VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formseleccio 
   Caption         =   "c"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "formseleccio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6060
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "formseleccio.frx":0442
      Height          =   5310
      Left            =   90
      OleObjectBlob   =   "formseleccio.frx":0452
      TabIndex        =   5
      Top             =   645
      Width           =   4500
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   75
      TabIndex        =   0
      Top             =   -60
      Width           =   4515
      Begin VB.CommandButton Command3 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2835
         Picture         =   "formseleccio.frx":0E04
         TabIndex        =   4
         Top             =   165
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   75
         TabIndex        =   3
         Top             =   195
         Width           =   2715
      End
      Begin VB.CommandButton sortirs 
         Height          =   390
         Left            =   4050
         Picture         =   "formseleccio.frx":118E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Alta  Registres"
         Top             =   165
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Height          =   390
         Left            =   3600
         Picture         =   "formseleccio.frx":1690
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   165
         Width           =   390
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1275
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1965
   End
End
Attribute VB_Name = "formseleccio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub refrescar()
 Dim tipusdato As Byte
 Dim grandoto As Integer
 Dim espais As Byte

 Data1.Refresh
 DBGrid2.Refresh
 DBGrid2.ReBind
 DBGrid2.AllowUpdate = False
 On Error GoTo fi
 For i = 0 To 50
   tipusdato = Data1.Recordset.Fields(DBGrid2.Columns(i).DataField).Type
   grandato = Data1.Recordset.Fields(DBGrid2.Columns(i).DataField).Size
   If grandato < 5 Then grandato = 5
   DBGrid2.Columns(i).Width = grandato * 115
   DBGrid2.Columns(i).Caption = UCase(DBGrid2.Columns(i).Caption)
    
 Next i
fi:
If formseleccio.Tag = "1" Then DBGrid2.Columns(0).Width = 0
End Sub

Private Sub Command1_Click()
  acceptar
End Sub

Private Sub Command2_Click()
 
End Sub

Private Sub Command3_Click()
  If Text1.Tag = "1" Then
   Data1.Recordset.FindFirst (DBGrid2.Columns(DBGrid2.Col).DataField + " like '*" + Text1.Text + "*'")
   Text1.Tag = ""
    Else: Data1.Recordset.FindNext (DBGrid2.Columns(DBGrid2.Col).DataField + " like '*" + Text1.Text + "*'"): Text1.Tag = ""
  End If
  'dbgrid2.Visible = True
  DBGrid2.SetFocus
End Sub

Private Sub Data1_Reposition()
 If DBGrid2.Tag = "" Then DBGrid2.Tag = Data1.RecordSource

End Sub

Private Sub dbgrid2_HeadClick(ByVal ColIndex As Integer)
  Data1.RecordSource = DBGrid2.Tag + " order by " + DBGrid2.Columns(ColIndex).DataField
  refrescar
End Sub

Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
 If KeyCode = 38 And DBGrid2.Row = 0 Then Text1.SetFocus
End Sub

Private Sub DBGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then acceptar
End Sub

Private Sub Form_Activate()
DBGrid2.Col = 1
DBGrid2.SetFocus
DoEvents
Text1.SetFocus
Data1.RecordSource = Data1.RecordSource + " "
If InStr(1, Data1.RecordSource, " from ") <> 0 Then
   formseleccio.Caption = "Busqueda de " + UCase(Mid(Data1.RecordSource, InStr(1, Data1.RecordSource, "from ") + 5, InStr(1, Mid(Data1.RecordSource, InStr(1, Data1.RecordSource, "from ") + 5), " ")))
 Else: formseleccio.Caption = "Busqueda de " + UCase(Data1.RecordSource)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 13 Or KeyCode = 112) And Not DBGrid2.AllowUpdate And Screen.ActiveControl.Name = "DBGrid2" Then acceptar
If (KeyCode = 13 Or KeyCode = 112) And Screen.ActiveForm.Name = "formseleccio" And Screen.ActiveControl.Name = "Text1" Then
   Command3_Click
End If
If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
End Sub

Private Sub Form_Load()
  centerscreen Me
  
End Sub

Private Sub Form_Resize()
On Error Resume Next
DBGrid2.Width = formseleccio.Width - 200
Frame1.Width = formseleccio.Width - 400
DBGrid2.Height = formseleccio.Height - DBGrid2.Top - 700
End Sub

Private Sub Form_Unload(Cancel As Integer)
seleccioret = 0
End Sub


Sub acceptar()
  seleccioret = 1
  Me.Hide
End Sub

Private Sub sortirs_Click()
 Unload Me
End Sub

Private Sub Text1_Change()
Text1.Tag = "1"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then DBGrid2.SetFocus
  'If KeyCode = 13 Then Command3_Click
End Sub

