VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formseleccio 
   Caption         =   "c"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "formseleccio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6060
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botofiltre 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   675
      Picture         =   "formseleccio.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   690
      Width           =   345
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "formseleccio.frx":0B14
      Height          =   5310
      Left            =   90
      OleObjectBlob   =   "formseleccio.frx":0B24
      TabIndex        =   4
      Top             =   645
      Width           =   4500
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   75
      TabIndex        =   0
      Top             =   -60
      Width           =   4515
      Begin VB.CommandButton Command2 
         Caption         =   "Cap"
         Height          =   375
         Left            =   3210
         TabIndex        =   5
         Top             =   180
         Width           =   405
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   75
         TabIndex        =   3
         Top             =   195
         Width           =   3090
      End
      Begin VB.CommandButton sortirs 
         Height          =   375
         Left            =   4050
         Picture         =   "formseleccio.frx":14F7
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Alta  Registres"
         Top             =   180
         Width           =   405
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   3630
         Picture         =   "formseleccio.frx":1A81
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   405
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
 'On Error GoTo errorr
 Data1.Refresh
 'If Data1.Recordset.EOF Then GoTo fi
 DBGrid2.Refresh
 DBGrid2.ReBind
 DBGrid2.AllowUpdate = False
 On Error GoTo fi
 
 For i = 0 To 50
   tipusdato = Data1.Recordset.Fields(DBGrid2.Columns(i).DataField).Type
   grandato = Data1.Recordset.Fields(DBGrid2.Columns(i).DataField).Size
   If grandato < 5 Then grandato = 5
   If Not formseleccio.visible Then DBGrid2.Columns(i).Width = grandato * 115
   DBGrid2.Columns(i).Caption = UCase(DBGrid2.Columns(i).Caption)
    
 Next i
fi:
If formseleccio.Tag = "1" Then DBGrid2.Columns(0).Width = 0
Exit Sub
errorr:

  Data1.RecordSource = UCase(formseleccio.Tag)
  Resume
End Sub

Private Sub Command1_Click()
  acceptar
End Sub

Private Sub Command2_Click()
seleccioret = 9
  Me.Hide
End Sub

Sub filtrarresultats()
 Dim colu As Byte
 colu = DBGrid2.col
 If sortirs.Tag <> "filtre" Then
  If Text1.Tag = "1" Then
   Data1.Recordset.FindFirst (DBGrid2.Columns(DBGrid2.col).DataField + " like '*" + Text1.Text + "*'")
   Text1.Tag = ""
    Else: Data1.Recordset.FindNext (DBGrid2.Columns(DBGrid2.col).DataField + " like '*" + Text1.Text + "*'"): Text1.Tag = ""
  End If
   Else
      Data1.RecordSource = possarfiltre
      'MsgBox Data1.RecordSource
      Data1.Refresh
      refrescar
   End If
  DBGrid2.visible = True
  'DBGrid2.SetFocus
  DBGrid2.col = colu

End Sub
Function possarfiltre()
   Dim va As String
   Dim res As String
   Dim andowhere As String
   Dim vcol As Integer
   va = formseleccio.Tag
   vcol = cadbl(botofiltre.Tag)
   andowhere = " where "
   If InStr(1, LCase(va), "where") Then andowhere = " and "
   res = va
   If InStr(1, LCase(va), "order") Then res = Mid(va, 1, InStr(1, va, "order") - 1)
   res = res + andowhere + (crearfiltredesdeltexte)
   'If InStr(1, va, "order") Then res = res + " " + Mid(va, InStr(1, va, "order"))
   If cadbl(botofiltre.Tag) > 0 Then res = res + " order by " + DBGrid2.Columns(vcol).DataField
   possarfiltre = res
End Function
Function crearfiltredesdeltexte() As String
  Dim camp As String
  Dim vseleccio As String
  Dim vvalor As String
  Dim vmid As String
  Dim v As Integer
  v = cadbl(botofiltre.Tag)
  camp = DBGrid2.Columns(v).DataField
  vvalor = Text1.Text
  While InStr(1, vvalor, ",") And Len(vvalor) > 1
     vmid = Mid(vvalor, 1, InStr(1, vvalor, ",") - 1)
     vvalor = Mid(vvalor, InStr(1, vvalor, ",") + 1)
     vseleccio = vseleccio + IIf(vseleccio <> "", " and ", "") + camp + " like '*" + vmid + "*'"
  Wend
  vseleccio = vseleccio + IIf(vseleccio <> "", " and ", "") + camp + " like '*" + vvalor + "*'"
  crearfiltredesdeltexte = vseleccio
End Function
Private Sub Data1_Reposition()
 If DBGrid2.Tag = "" Then DBGrid2.Tag = Data1.RecordSource

End Sub

Private Sub DBGrid2_DblClick()
acceptar
End Sub

Private Sub DBGrid2_HeadClick(ByVal ColIndex As Integer)
  Dim instsql As String
  colocar_botofiltre ColIndex
  instsql = DBGrid2.Tag
  If InStr(1, instsql, "order by") > 0 Then instsql = Mid(instsql, 1, InStr(1, instsql, "order by") - 1)
  If InStr(1, instsql, "group by") = 0 Then instsql = instsql + " order by [" + DBGrid2.Columns(ColIndex).DataField + "]"
  Data1.RecordSource = instsql
  refrescar
  DBGrid2.col = ColIndex
  Text1.SetFocus
  filtrarresultats
End Sub
Sub colocar_botofiltre(col As Integer)
   botofiltre.Left = DBGrid2.Columns(col).Left + DBGrid2.Columns(col).Width - botofiltre.Width
   botofiltre.Tag = col
End Sub
Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
 If KeyCode = 38 And DBGrid2.row = 0 Then Text1.SetFocus
End Sub

Private Sub DBGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then acceptar
End Sub

Private Sub Form_Activate()
If DBGrid2.Columns.Count > 1 Then
   DBGrid2.col = 1
   colocar_botofiltre DBGrid2.col
End If
DBGrid2.SetFocus
DoEvents
Text1.SetFocus
Data1.RecordSource = Data1.RecordSource + " "
If InStr(1, Data1.RecordSource, " from ") <> 0 Then
   formseleccio.Caption = "Busqueda de " + UCase(Mid(Data1.RecordSource, InStr(1, Data1.RecordSource, "from ") + 5, InStr(1, Mid(Data1.RecordSource, InStr(1, Data1.RecordSource, "from ") + 5), " ")))
 Else: formseleccio.Caption = "Busqueda de " + UCase(Data1.RecordSource)
End If
If cadbl(Text1.Tag) > 0 Then DBGrid2.col = cadbl(Text1.Tag): Text1.Tag = ""
If Me.Tag = "" Then Me.Tag = Data1.RecordSource
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 13 Or KeyCode = 112) And Not DBGrid2.AllowUpdate And Screen.ActiveControl.Name = "DBGrid2" Then acceptar
If (KeyCode = 13 Or KeyCode = 112) And Screen.ActiveForm.Name = "formseleccio" And Screen.ActiveControl.Name = "Text1" Then
   filtrarresultats
End If
If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
End Sub

Private Sub Form_Load()
  centerscreen Me
  Frame1.Tag = Data1.RecordSource
 
End Sub

Private Sub Form_Resize()
On Error Resume Next
DBGrid2.Width = formseleccio.Width - 300
Frame1.Width = formseleccio.Width - 300
DBGrid2.Height = formseleccio.Height - DBGrid2.Top - 700
sortirs.Left = Frame1.Width - sortirs.Width - 75
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
 filtrarresultats
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then DBGrid2.SetFocus

  'If KeyCode = 13 Then Command3_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub
