VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formalta 
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "formalta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4125
      Top             =   825
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "formalta.frx":0442
      Height          =   5340
      Left            =   75
      Negotiate       =   -1  'True
      OleObjectBlob   =   "formalta.frx":0452
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   4515
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   75
      TabIndex        =   0
      Top             =   -75
      Width           =   4515
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "formalta.frx":0E00
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   975
         Picture         =   "formalta.frx":1232
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   525
         Picture         =   "formalta.frx":1544
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Consulta Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   4050
         Picture         =   "formalta.frx":1892
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Height          =   390
         Left            =   3600
         Picture         =   "formalta.frx":1D94
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
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
         TabIndex        =   7
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
End
Attribute VB_Name = "formalta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub refrescar()
 Dim tipusdato As Byte
 Dim grandoto As Integer
 Dim espais As Byte
 Data1.Refresh
 DBGrid1.Refresh
 DBGrid1.ReBind
 DBGrid1.AllowUpdate = False
 On Error GoTo fi
 For i = 0 To 50
   tipusdato = Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Type
   grandato = Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Size
   If grandato < 5 Then grandato = 5
   DBGrid1.Columns(i).Width = grandato * 115
   DBGrid1.Columns(i).Caption = UCase(DBGrid1.Columns(i).Caption)
 Next i
fi:
End Sub

Private Sub alta_Click()
  alta_registre
End Sub

Private Sub Command1_Click()
  acceptar
End Sub


Private Sub Data1_Reposition()
 If DBGrid1.Tag = "" Then DBGrid1.Tag = Data1.RecordSource

End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
  Data1.RecordSource = DBGrid1.Tag + " order by " + DBGrid1.Columns(ColIndex).DataField
  refrescar
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{TAB}": KeyCode = 0
 If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
End Sub

Private Sub Form_Load()
  centerscreen Me
  DBGrid1.ReBind
End Sub

Private Sub Form_Unload(Cancel As Integer)
seleccioret = 0
End Sub

Private Sub Label1_Click()

End Sub

Private Sub modificar_Click()
   DBGrid1.Enabled = True
   Data1.Recordset.Edit
   DBGrid1.AllowUpdate = True
   DBGrid1.Col = 0
   DBGrid1.SetFocus
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
      'busco el mes gran i el poso a codi +1
      'If Not buscant Then
      '  Set rsttmp = dbtmp.OpenRecordset("select max(codi) as [grancodi] from clients")
      '  If Not rsttmp.EOF Then
      '    Text1 = atrim(cadbl(rsttmp!grancodi) + 1)
      '   Else: Text1 = "1"
      '  End If
      'End If
        DBGrid1.AllowUpdate = True
      DBGrid1.Col = 0
      DBGrid1.SetFocus
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

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
  If Len(DBGrid1.Text) >= Data1.Recordset.Fields(DBGrid1.Columns(DBGrid1.Col).DataField).Size And KeyAscii > 55 Then KeyAscii = 0
End Sub

