VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formseleccio 
   Caption         =   "Escullir"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   ControlBox      =   0   'False
   Icon            =   "formseleccio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6060
   ScaleWidth      =   6630
   StartUpPosition =   1  'CenterOwner
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
      Begin VB.CommandButton alta 
         Height          =   390
         Left            =   2115
         Picture         =   "formseleccio.frx":0E25
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   180
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton CommandXLS 
         Height          =   390
         Left            =   3960
         Picture         =   "formseleccio.frx":13AF
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   300
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton Command3 
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
         Picture         =   "formseleccio.frx":1939
         Style           =   1  'Graphical
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
         Picture         =   "formseleccio.frx":1EC3
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Sortir"
         Top             =   180
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Height          =   390
         Left            =   3600
         Picture         =   "formseleccio.frx":244D
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
 Dim vamplades(20) As Double
 Dim i As Double
 For f = 0 To DBGrid2.Columns.Count - 1
    If DBGrid2.Columns(f).Visible Then
        vamplades(f) = DBGrid2.Columns(f).Width
         Else: vamplades(f) = 0
    End If
 Next f
'  Clipboard.Clear
' Clipboard.SetText Data1.RecordSource
 Data1.Refresh
 
 'If Data1.Recordset.EOF Then GoTo fi
 DBGrid2.Refresh

 DBGrid2.ReBind
 DBGrid2.AllowUpdate = False
 On Error GoTo fi
 For f = 0 To DBGrid2.Columns.Count - 1
   tipusdato = Data1.Recordset.Fields(DBGrid2.Columns(f).DataField).Type
   grandato = Data1.Recordset.Fields(DBGrid2.Columns(f).DataField).Size
   If grandato < 5 Then grandato = 5
   DBGrid2.Columns(f).Width = grandato * 115
   DBGrid2.Columns(f).Caption = UCase(DBGrid2.Columns(f).Caption)
    
 Next f
 If Int(vamplades(0)) <> 1544 Then
    For f = 0 To DBGrid2.Columns.Count - 1
       If vamplades(f) > 100 Then
            DBGrid2.Columns(f).Width = vamplades(f)
            DBGrid2.Columns(f).Visible = True
            Else: DBGrid2.Columns(f).Visible = False
       End If
    Next f
 End If
fi:
If formseleccio.Tag = "1" Then DBGrid2.Columns(0).Width = 0
End Sub

Private Sub alta_Click()
seleccioret = 2
  Me.Hide
End Sub

Private Sub Command1_Click()
  acceptar
End Sub

Private Sub Command2_Click()
 
End Sub

Private Sub Command3_Click()
 Dim colu As Byte
 Dim vamples(200) As Variant
 colu = DBGrid2.Col
 If InStr(1, UCase(Data1.RecordSource), "GROUP BY ") > 0 Then Exit Sub
 If Command3.Tag <> "filtre" Then
  If Text1.Tag = "1" Then
   Data1.Recordset.FindFirst (DBGrid2.Columns(DBGrid2.Col).DataField + " like '*" + Text1.Text + "*'")
   Text1.Tag = ""
    Else: Data1.Recordset.FindNext (DBGrid2.Columns(DBGrid2.Col).DataField + " like '*" + Text1.Text + "*'"): Text1.Tag = ""
  End If
   Else
      For i = 0 To DBGrid2.Columns.Count - 1: vamples(i) = DBGrid2.Columns(i).Width: Next i
      Data1.RecordSource = possarfiltre
      Data1.Refresh
      For i = 0 To DBGrid2.Columns.Count - 1: DBGrid2.Columns(i).Width = vamples(i): Next i
'      refrescar
   End If
  DBGrid2.Visible = True
  DBGrid2.SetFocus
  DBGrid2.Col = colu
End Sub
Function possarfiltre()
   Dim va As String
   Dim res As String
   Dim andowhere As String
   va = formseleccio.Tag
   andowhere = " where "
   If InStr(1, va, "where") Then andowhere = " and "
   res = va
   If InStr(1, va, "order") Then res = Mid(va, 1, InStr(1, va, "order") - 1)
   res = res + andowhere + (DBGrid2.Columns(DBGrid2.Col).DataField + " like '*" + Text1.Text + "*'")
  ' If InStr(1, va, "order") Then res = Mid(va, InStr(1, va, "order"))
   possarfiltre = res
End Function

Private Sub CommandXLS_Click()
  Dim vnomfitxer As String
  Dim i As Long
  Dim vlinia As String
  vnomfitxer = "c:\temp\llistat_seleccio.csv"
  If Data1.Recordset.EOF Then Exit Sub
  If existeix(vnomfitxer) Then Kill vnomfitxer
  Open vnomfitxer For Output As #1
   Data1.Recordset.MoveFirst
     'capçalera
   For i = 0 To Data1.Recordset.Fields.Count - 1
       vlinia = vlinia + UCase(Data1.Recordset.Fields(i).Name) + ";"
   Next i
   Print #1, vlinia
      'linies
   While Not Data1.Recordset.EOF
     vlinia = ""
     For i = 0 To Data1.Recordset.Fields.Count - 1
       vlinia = vlinia + atrim(Data1.Recordset.Fields(i).Value) + ";"
     Next i
     Print #1, vlinia
     Data1.Recordset.MoveNext
   Wend
  Close #1
    'si existeix obro el fitxer
  If existeix(vnomfitxer) Then obrir_document vnomfitxer
End Sub

Private Sub Data1_Reposition()
 If DBGrid2.Tag = "" Then DBGrid2.Tag = Data1.RecordSource

End Sub

Private Sub DBGrid2_DblClick()
acceptar
End Sub

Private Sub DBGrid2_HeadClick(ByVal ColIndex As Integer)
If InStr(1, UCase(DBGrid2.Tag), " GROUP BY") Then Exit Sub
  If InStr(1, DBGrid2.Tag, " order by") > 0 Then DBGrid2.Tag = Mid(DBGrid2.Tag, 1, InStr(1, DBGrid2.Tag, " order by"))
  Data1.RecordSource = DBGrid2.Tag + " order by " + DBGrid2.Columns(ColIndex).DataField
 
  refrescar
  DBGrid2.Col = ColIndex
End Sub

Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
 If KeyCode = 38 And DBGrid2.Row = 0 Then Text1.SetFocus
End Sub

Private Sub DBGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then acceptar
End Sub

Private Sub Form_Activate()
 
If DBGrid2.Columns.Count > 1 Then DBGrid2.Col = 1
DBGrid2.SetFocus
DoEvents
If Text1.Visible Then Text1.SetFocus
Data1.RecordSource = Data1.RecordSource + " "
If InStr(1, Data1.RecordSource, " from ") <> 0 Then
   formseleccio.Caption = "Busqueda de " + UCase(Mid(Data1.RecordSource, InStr(1, Data1.RecordSource, "from ") + 5, InStr(1, Mid(Data1.RecordSource, InStr(1, Data1.RecordSource, "from ") + 5), " ")))
 Else: formseleccio.Caption = "Busqueda de " + UCase(Data1.RecordSource)
End If
If cadbl(Text1.Tag) > 0 Or Text1.Tag = "0" Then DBGrid2.Col = cadbl(Text1.Tag): Text1.Tag = ""
If Me.Tag = "" Then Me.Tag = Data1.RecordSource
End Sub

Private Sub Form_GotFocus()
Form_Resize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
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
DBGrid2.Width = formseleccio.Width - 300
Frame1.Width = formseleccio.Width - 300
DBGrid2.Height = formseleccio.Height - DBGrid2.Top - 700
sortirs.Left = Frame1.Width - sortirs.Width - 75
CommandXLS.Left = sortirs.Left - CommandXLS.Width - 75
CommandXLS.Top = sortirs.Top
If formseleccio.alta.Visible Then
        formseleccio.alta.Left = formseleccio.sortirs.Left - (formseleccio.alta.Width + 10)
        formseleccio.alta.Top = formseleccio.sortirs.Top
End If
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

Private Sub Text1_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub
