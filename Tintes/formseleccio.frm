VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formseleccio 
   Caption         =   "c"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   Icon            =   "formseleccio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6060
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botofiltre 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1815
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
      Width           =   5085
      Begin VB.CommandButton bborrar 
         Height          =   375
         Left            =   4020
         Picture         =   "formseleccio.frx":14F7
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Borrar registre."
         Top             =   195
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2835
         Picture         =   "formseleccio.frx":15E3
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   345
      End
      Begin VB.CommandButton bimprimir 
         Height          =   375
         Left            =   4185
         Picture         =   "formseleccio.frx":1B6D
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cap"
         Height          =   375
         Left            =   3195
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
         Width           =   2730
      End
      Begin VB.CommandButton sortirs 
         Height          =   375
         Left            =   4605
         Picture         =   "formseleccio.frx":20F7
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Alta  Registres"
         Top             =   180
         Width           =   405
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   3615
         Picture         =   "formseleccio.frx":2681
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   405
      End
      Begin VB.Label cmissatge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3660
         TabIndex        =   7
         Top             =   285
         Width           =   75
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
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Type camp

    caption As String
    width As Double
    visible As Boolean
    tag As String

End Type

' Declaramos el vector
Dim valorsreixa(0 To 40) As camp
Sub refrescar()
 Dim tipusdato As Byte
 Dim grandoto As Integer
 Dim espais As Byte
 On Error GoTo errorr
 guardarvalorsreixa
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
   If Not formseleccio.visible Then DBGrid2.Columns(i).width = grandato * 115
   DBGrid2.Columns(i).caption = UCase(DBGrid2.Columns(i).caption)
    
 Next i
fi:
If formseleccio.tag = "1" Then DBGrid2.Columns(0).width = 0
carregarvalorsreixa
Exit Sub
errorr:
 ' Clipboard.Clear
 ' Clipboard.SetText Data1.RecordSource
  'MsgBox err.Description
  Data1.RecordSource = UCase(formseleccio.tag)
  MsgBox "ERROR", vbCritical, "Error": Exit Sub
  Resume
End Sub
Sub guardarvalorsreixa()
  If Me.tag = "" Then Exit Sub
  For i = 0 To DBGrid2.Columns.Count - 1
    valorsreixa(i).caption = DBGrid2.Columns(i).caption
    valorsreixa(i).width = DBGrid2.Columns(i).width
    valorsreixa(i).visible = DBGrid2.Columns(i).visible
  Next i
End Sub
Sub carregarvalorsreixa()
  If Me.tag = "" Then Exit Sub
  For i = 0 To DBGrid2.Columns.Count - 1
    DBGrid2.Columns(i).caption = valorsreixa(i).caption
    DBGrid2.Columns(i).width = valorsreixa(i).width
    DBGrid2.Columns(i).visible = valorsreixa(i).visible
  Next i
End Sub

Private Sub bborrar_Click()
  ColocarEnTop formseleccio, False
  seleccioret = 8
  Me.Hide
End Sub

Private Sub bimprimir_Click()
seleccioret = 9
Me.Hide
End Sub

Private Sub Command1_Click()
  ColocarEnTop formseleccio, False
  seleccioret = 1
  Me.Hide
End Sub

Private Sub Command2_Click()
 ColocarEnTop formseleccio, False
 seleccioret = 9
  Me.Hide
End Sub

Sub filtrarresultats()
 Dim colu As Byte
 On Error Resume Next
 colu = DBGrid2.col
 If sortirs.tag <> "filtre" Then
  If Text1.tag = "1" Then
   Data1.Recordset.FindFirst (DBGrid2.Columns(DBGrid2.col).DataField + " like '*" + Text1.Text + "*'")
   Text1.tag = ""
    Else: Data1.Recordset.FindNext (DBGrid2.Columns(DBGrid2.col).DataField + " like '*" + Text1.Text + "*'"): Text1.tag = ""
  End If
   Else
      Data1.RecordSource = possarfiltre
      'MsgBox Data1.RecordSource
      'Data1.Refresh
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
   va = formseleccio.tag
   vcol = cadbl(botofiltre.tag)
   andowhere = " where "
   If InStr(1, LCase(va), "where") Then andowhere = " and "
   res = va
   If InStr(1, LCase(va), "order") Then res = Mid(va, 1, InStr(1, va, "order") - 1)
   res = res + andowhere + (crearfiltredesdeltexte)
   'If InStr(1, va, "order") Then res = res + " " + Mid(va, InStr(1, va, "order"))
   If cadbl(botofiltre.tag) > 0 Then res = res + " order by " + DBGrid2.Columns(vcol).DataField
   possarfiltre = res
End Function
Function crearfiltredesdeltexte() As String
  Dim camp As String
  Dim vseleccio As String
  Dim vvalor As String
  Dim vmid As String
  Dim v As Integer
  v = cadbl(botofiltre.tag)
  camp = DBGrid2.Columns(v).DataField
  vvalor = Text1.Text
  While InStr(1, vvalor, " ") And Len(vvalor) > 1
     vmid = Mid(vvalor, 1, InStr(1, vvalor, " ") - 1)
     vvalor = Mid(vvalor, InStr(1, vvalor, " ") + 1)
     vseleccio = vseleccio + IIf(vseleccio <> "", " and ", "") + camp + " like '*" + vmid + "*'"
  Wend
  vseleccio = vseleccio + IIf(vseleccio <> "", " and ", "") + camp + " like '*" + vvalor + "*'"
  crearfiltredesdeltexte = vseleccio
End Function

Private Sub Command3_Click()
  filtrarresultats
End Sub

Private Sub Data1_Reposition()
 If DBGrid2.tag = "" Then DBGrid2.tag = Data1.RecordSource

End Sub

Private Sub DBGrid2_DblClick()
acceptar
End Sub

Private Sub DBGrid2_HeadClick(ByVal ColIndex As Integer)
    dbgrid1_clickcapçalera ColIndex
End Sub
Sub dbgrid1_clickcapçalera(ColIndex As Integer)
  Dim instsql As String
  colocar_botofiltre ColIndex
  instsql = DBGrid2.tag
 ' If InStr(1, UCase(instsql), "GROUP BY") > 0 Then GoTo cont
  If InStr(1, instsql, "order by") > 0 Then instsql = Mid(instsql, 1, InStr(1, instsql, "order by") - 1)
  instsql = substituir(instsql, ";", "")
  ' MsgBox instsql
  Data1.RecordSource = instsql + " order by " + atrim(ColIndex + 1) '+ DBGrid2.Columns(ColIndex).DataField
  
  refrescar
cont:
  DBGrid2.col = ColIndex
  If Text1.visible Then Text1.SetFocus
  filtrarresultats
End Sub
Sub colocar_botofiltre(col As Integer)
   botofiltre.Left = DBGrid2.Columns(col).Left + DBGrid2.Columns(col).width - botofiltre.width
   botofiltre.tag = col
End Sub
Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
 If KeyCode = 38 And DBGrid2.Row = 0 Then Text1.SetFocus
End Sub

Private Sub DBGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then acceptar
End Sub

Private Sub DBGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'   If Button = 2 Then MsgBox DBGrid2.Columns(DBGrid2.col).width
End Sub

Private Sub Form_Activate()
  ColocarEnTop formseleccio, True
'If DBGrid2.Columns.Count > 0 And cadbl(cmissatge.tag) <= DBGrid2.Columns.Count Then
'   DBGrid2.col = IIf(cmissatge.tag <> "", cadbl(cmissatge.tag), IIf(DBGrid2.Columns.Count > 1, 1, 0))
'   colocar_botofiltre DBGrid2.col
'End If
If DBGrid2.Columns.Count > 1 Then
   If cadbl(botofiltre.tag) = 0 Then DBGrid2.col = 1 Else DBGrid2.col = cadbl(botofiltre.tag)
   colocar_botofiltre DBGrid2.col
End If
DBGrid2.SetFocus
DoEvents
If formseleccio.visible Then Text1.SetFocus
Data1.RecordSource = Data1.RecordSource + " "
If Me.caption = "c" Then
 If InStr(1, Data1.RecordSource, " from ") <> 0 Then
    formseleccio.caption = "Busqueda de " + UCase(Mid(Data1.RecordSource, InStr(1, Data1.RecordSource, "from ") + 5, InStr(1, Mid(Data1.RecordSource, InStr(1, Data1.RecordSource, "from ") + 5), " ")))
  Else: formseleccio.caption = "Busqueda de " + UCase(Data1.RecordSource)
 End If
End If
If cadbl(Text1.tag) > 0 Then DBGrid2.col = cadbl(Text1.tag): Text1.tag = ""
If Me.tag = "" Then Me.tag = Data1.RecordSource
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = 13 Or KeyCode = 112) Then acceptar
'If (KeyCode = 13 Or KeyCode = 112) And Not DBGrid2.AllowUpdate And Screen.ActiveControl.Name = "DBGrid2" Then acceptar
'If (KeyCode = 13 Or KeyCode = 112) And Screen.ActiveForm.Name = "formseleccio" And Screen.ActiveControl.Name = "Text1" Then
'   filtrarresultats
'End If
If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
End Sub

Private Sub Form_Load()
  centerscreen Me
  Frame1.tag = Data1.RecordSource
 
End Sub

Private Sub Form_Resize()
On Error Resume Next
DBGrid2.width = formseleccio.width - 300
Frame1.width = formseleccio.width - 300
DBGrid2.Height = formseleccio.Height - DBGrid2.Top - 700
sortirs.Left = Frame1.width - sortirs.width - 75
bimprimir.Left = Frame1.width - sortirs.width - 75 - bimprimir.width

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ColocarEnTop formseleccio, False
  seleccioret = 0
End Sub


Sub acceptar()
  seleccioret = 1
  Me.Hide
End Sub

Private Sub sortirs_Click()
 ColocarEnTop formseleccio, False
 Unload Me
End Sub

Private Sub Text1_Change()
 Text1.tag = "1"
 filtrarresultats
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then DBGrid2.SetFocus
  
  'If KeyCode = 13 Then Command3_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub
Public Function ColocarEnTop(vform As Form, ByVal fColocarEnTop As Boolean) As Boolean
On Error Resume Next
Dim f As Boolean
'Si la función falla devuelve \"False\"
f = (SetWindowPos(vform.hwnd, IIf(fColocarEnTop = True, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, flags) <> 0)
fEstaEnTop = (fColocarEnTop And (f = True))
ColocarEnTop = f
End Function
