VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formaccessoris 
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   Icon            =   "formaccessoris.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6195
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox opcions 
      Height          =   315
      ItemData        =   "formaccessoris.frx":0442
      Left            =   390
      List            =   "formaccessoris.frx":044F
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4125
      Top             =   825
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   75
      TabIndex        =   0
      Top             =   -75
      Width           =   8925
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "formaccessoris.frx":0470
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   975
         Picture         =   "formaccessoris.frx":08A2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   525
         Picture         =   "formaccessoris.frx":0BB4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Consulta Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   8460
         Picture         =   "formaccessoris.frx":0F02
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Alta  Registres"
         Top             =   135
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Height          =   390
         Left            =   8010
         Picture         =   "formaccessoris.frx":1404
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   135
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label estattaula 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   6
         Top             =   150
         Width           =   1515
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
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "formaccessoris.frx":1716
      Height          =   5475
      Left            =   105
      Negotiate       =   -1  'True
      OleObjectBlob   =   "formaccessoris.frx":1726
      TabIndex        =   10
      Top             =   555
      Width           =   8880
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   105
      TabIndex        =   8
      Top             =   6075
      Width           =   8820
   End
   Begin VB.Label autonum 
      Height          =   135
      Left            =   0
      TabIndex        =   7
      Top             =   1335
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "formaccessoris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub refrescar()
 Dim tipusdato As Byte
 Dim grandoto As Integer
 Dim espais As Byte
 Dim factor As Integer
 Data1.Refresh
 DBGrid1.Refresh
 DBGrid1.ReBind
 DBGrid1.AllowUpdate = False
  centerscreen Me
  If Me.Tag <> "" Then
     factor = cadbl(Me.Tag)
    Else: factor = 130
  End If
  
 On Error GoTo cont
 For i = 0 To 50
   tipusdato = Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Type
   grandato = Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Size
   
   If grandato < 5 Then grandato = 5
   DBGrid1.Columns(i).Width = grandato * (IIf(tipusdato <> 10, factor, 130))
   DBGrid1.Columns(i).Caption = UCase(DBGrid1.Columns(i).Caption)
 Next i
cont:
On Error GoTo fi
 If DBGrid1.Columns.Count > 5 Then
  DBGrid1.Columns(5).Visible = False
  DBGrid1.Columns(2).Width = DBGrid1.Columns(2).Width / 2
  DBGrid1.Columns(1).Width = DBGrid1.Columns(1).Width + 130
  DBGrid1.Columns(1).Button = True
 End If
fi:

If autonum.Caption <> "" Then
     DBGrid1.Columns(0).Width = 0
     Set dbtmp = OpenDatabase(Data1.DatabaseName)
  End If
  
End Sub

Private Sub alta_Click()
  alta_registre
End Sub

Private Sub Command1_Click()
'  acceptar
 
End Sub


Private Sub Data1_Reposition()
 If DBGrid1.Tag = "" Then DBGrid1.Tag = Data1.RecordSource

End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
  Data1.RecordSource = DBGrid1.Tag + " order by " + DBGrid1.Columns(ColIndex).DataField
  refrescar
End Sub
Private Sub DBGrid1_ButtonClick(ByVal ColIndex As Integer)
  opcions.Visible = True
  opcions.Text = "T Truquel"
  'opcions.Left = DBGrid1.Columns(ColIndex).Left + 30
  opcions.Top = DBGrid1.RowTop(DBGrid1.row) + DBGrid1.Top
  opcions.SetFocus
  SendKeys ("%{DOWN}")
  'While Screen.ActiveControl.Name = "opcions"
  '  DoEvents
  'Wend
  'If Data1.Recordset.EditMode > 0 Then DBGrid1.Col = 1: DoEvents: DBGrid1.SetFocus
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 And DBGrid1.col = 1 Then DBGrid1_ButtonClick 1
  If KeyCode = 113 And DBGrid1.col = 4 And DBGrid1.AllowUpdate Then
     DBGrid1.Text = triar_mesura
  End If
End Sub
Function triar_mesura() As String
  Load formseleccio
   formseleccio.Caption = "Selecciona un Valor"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from mesureslineals"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   triar_mesura = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
End Function

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
 Dim t As Integer
  t = tamany_camp(Data1.Recordset.Fields(DBGrid1.Columns(DBGrid1.col).DataField))
  If Len(DBGrid1.Text) >= t And KeyAscii > 47 Then KeyAscii = 0
  If DBGrid1.col = 4 Then KeyAscii = 0
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err

If DBGrid1.Bookmark <> LastRow Then
 If Data1.Recordset.EditMode = 0 Then
     DBGrid1.AllowUpdate = False
 End If
End If

status = Data1.Recordset.Fields(DBGrid1.Columns(DBGrid1.col).DataField).ValidationText
If Data1.Recordset.EditMode = 1 And DBGrid1.col = 1 Then status = status + "(NO ES POT MODIFICAR NOMES BORRAR)"
Exit Sub
err:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
 If KeyAscii = 13 Then SendKeys "{TAB}": KeyCode = 0
If KeyAscii > 50 Then KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub Form_Load()
  centerscreen Me
  DBGrid1.ReBind
  refrescar
End Sub

Private Sub Form_Resize()
   DBGrid1.Width = formaccessoris.Width - DBGrid1.Left - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
seleccioret = 0
End Sub

Private Sub Label1_Click()

End Sub

Private Sub modificar_Click()
  If Not Data1.Recordset.EOF Then
   DBGrid1.Enabled = True
   Data1.Recordset.Edit
   DBGrid1.AllowUpdate = True
     DBGrid1.SetFocus
   'DBGrid1.Col = 1
 
  End If
End Sub

Private Sub opcions_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then guardarvalorcombo
End Sub

Private Sub opcions_LostFocus()
  guardarvalorcombo
  'opcions.Visible = False
  'DBGrid1.Col = 2
  'DoEvents
  'DBGrid1.SetFocus

End Sub

Sub guardarvalorcombo()
If Data1.Recordset.EditMode > 0 Then DBGrid1.Columns(1).Text = Mid(opcions.List(opcions.ListIndex), 1, 1)
opcions.Visible = False
DBGrid1.SetFocus
DBGrid1.col = 0

DoEvents
End Sub
Private Sub sortir_Click()
 Unload Me
End Sub

Sub acceptar()
  seleccioret = 1
  Me.Hide
End Sub

Private Sub Text1_Change()
Text1.Tag = "1"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 65 And Not DBGrid1.AllowUpdate Then alta_registre: KeyCode = 0
'If KeyCode = 69 Then buscar_registre
If KeyCode = 27 Then cancelar_registre
If KeyCode = 112 Then gravar_registre
'If KeyCode = 13 Then SendKeys "{TAB}": KeyCode = 0


End Sub


Sub alta_registre()
 If DBGrid1.AllowUpdate = False Then
      DBGrid1.Enabled = True
      Data1.Recordset.AddNew
      DoEvents
      DBGrid1.col = 0
      If autonum.Caption <> "" Then
        'busco el mes gran i el poso a codi +1
        Set rsttmp = dbtmp.OpenRecordset("select max(codi) as [grancodi] from " + autonum.Caption)
        If Not rsttmp.EOF Then
          DBGrid1.Columns(0).Text = atrim(cadbl(rsttmp!grancodi) + 1)
              Else: DBGrid1.Columns(0).Text = "1"
        End If
        DBGrid1.col = 1
      End If
      DBGrid1.AllowUpdate = True
      
      DBGrid1.SetFocus
      DBGrid1_ButtonClick 1
 End If
End Sub
Sub gravar_registre()
On Error GoTo err
 If Data1.Recordset.EditMode > 0 Then
      Data1.Recordset.Update
      DBGrid1.AllowUpdate = False
      Data1.Recordset.Bookmark = Data1.Recordset.LastModified
 End If

Exit Sub
err:
  MsgBox "Hi ha hagut un error al gravar les dades. Potser algun valor està duplicat a la clau. Torna-ho a provar", vbCritical, "Atenció"
End Sub
Sub cancelar_registre()
  If Data1.Recordset.EditMode > 0 Then
   Data1.Recordset.CancelUpdate
     Else: Unload Me
  End If
  DBGrid1.AllowUpdate = False
  opcions.Visible = False
End Sub

Private Sub eliminar_Click()
 On Error GoTo err
  If MsgBox("Segur que vols Eliminar?", vbYesNo + vbCritical, "Atenció") = 6 Then
    Data1.Recordset.Delete
    Data1.Recordset.MoveNext
    If Data1.Recordset.EOF Then Data1.Recordset.MovePrevious
  End If
 Exit Sub
err:
  MsgBox "No s'ha pogut eliminar possiblement perque tingui registres relacionats. O bé no hi ha res per eliminar."
End Sub

Private Sub Timer1_Timer()
 estattaula.Caption = textestattaula(Data1.EditMode)
  If estattaula.ForeColor <> QBColor(0) Then
     estattaula.ForeColor = QBColor(0)
    Else: estattaula.ForeColor = QBColor(14)
  End If
End Sub
