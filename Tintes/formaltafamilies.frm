VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formaltafamilies 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "formaltafamilies.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton canvicolor 
      Caption         =   "Canviar el color d'etiqueta de la subfamilia"
      Height          =   315
      Left            =   1380
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5490
      Width           =   3615
   End
   Begin VB.Data subfamilies 
      Caption         =   "subfamilies"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2610
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5430
      Visible         =   0   'False
      Width           =   2325
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "formaltafamilies.frx":0442
      Height          =   2430
      Left            =   165
      OleObjectBlob   =   "formaltafamilies.frx":0458
      TabIndex        =   11
      Top             =   5820
      Width           =   6810
   End
   Begin VB.TextBox colsbloc 
      Height          =   285
      Left            =   15
      TabIndex        =   10
      Top             =   450
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4125
      Top             =   825
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "formaltafamilies.frx":132F
      Height          =   4875
      Left            =   105
      Negotiate       =   -1  'True
      OleObjectBlob   =   "formaltafamilies.frx":133F
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   6870
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   135
      TabIndex        =   0
      Top             =   -30
      Width           =   6870
      Begin VB.CommandButton alta 
         Height          =   390
         Left            =   75
         Picture         =   "formaltafamilies.frx":1D0E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Alta  Registres"
         Top             =   135
         Width           =   390
      End
      Begin VB.CommandButton eliminar 
         Height          =   390
         Left            =   855
         Picture         =   "formaltafamilies.frx":2298
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Eliminacio Registres"
         Top             =   135
         Width           =   390
      End
      Begin VB.CommandButton modificar 
         Height          =   390
         Left            =   465
         Picture         =   "formaltafamilies.frx":2822
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Consulta Registres"
         Top             =   135
         Width           =   390
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   6420
         Picture         =   "formaltafamilies.frx":2DAC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Alta  Registres"
         Top             =   120
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Height          =   390
         Left            =   1245
         Picture         =   "formaltafamilies.frx":3336
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   135
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
      Height          =   345
      Left            =   1470
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5445
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label Label1 
      Caption         =   "Subfamilies"
      Height          =   255
      Left            =   135
      TabIndex        =   12
      Top             =   5535
      Width           =   2070
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   135
      TabIndex        =   9
      Top             =   5985
      Width           =   4470
   End
   Begin VB.Label autonum 
      Height          =   135
      Left            =   0
      TabIndex        =   8
      Top             =   1335
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "formaltafamilies"
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
 DBGrid2.Columns(0).Locked = True
  centerscreen Me
  If Me.tag <> "" Then
     factor = cadbl(Me.tag)
    Else: factor = 130
  End If
  
 On Error GoTo fi
 For i = 0 To 50
   tipusdato = Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Type
   grandato = Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Size
   
   If grandato < 5 Then grandato = 5
   DBGrid1.Columns(i).width = grandato * (IIf(tipusdato <> 10, factor, 130))
   DBGrid1.Columns(i).caption = UCase(DBGrid1.Columns(i).caption)
 Next i
  
         
  
fi:

If autonum.caption <> "" Then
     DBGrid1.Columns(0).width = 0
     Set dbtmp = OpenDatabase(Data1.DatabaseName)
  End If
  
End Sub

Private Sub alta_Click()
  alta_registre
End Sub

Private Sub canvicolor_Click()
  Dim colorescullit As String
  If subfamilies.Recordset.EOF Then Exit Sub
  Load formseleccio
  formseleccio.caption = "Selecciona un Color"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from colorsetiquetes order by nomcolor"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(1).width = 0
  formseleccio.Show 1
  If seleccioret = 1 Then
   colorescullit = atrim(formseleccio.Data1.Recordset!nomcolor)
  End If
  Unload formseleccio
  If colorescullit <> atrim(subfamilies.Recordset!color) Then
      If subfamilies.Recordset.EditMode = 0 Then subfamilies.Recordset.Edit
      subfamilies.Recordset!color = colorescullit
      subfamilies.Recordset.Update
  End If
End Sub

Private Sub Command1_Click()
  acceptar
End Sub


Private Sub Data1_Reposition()
 If DBGrid1.tag = "" Then DBGrid1.tag = Data1.RecordSource
  If Not Data1.Recordset.EOF And Data1.tag <> "" Then
    subfamilies.RecordSource = "select * from " + Data1.tag + " where codifam=" + atrim(cadbl(Data1.Recordset!codi))
    subfamilies.Refresh
  End If
End Sub

Private Sub DBGrid1_ButtonClick(ByVal ColIndex As Integer)
   If ColIndex = 2 Then
        Load formseleccio
        formseleccio.caption = "Selecciona un Valor"
        formseleccio.Data1.DatabaseName = camitintes
        formseleccio.Data1.RecordSource = "select * from tipusfamilies order by nom"
        formseleccio.refrescar
        formseleccio.Show 1
        If seleccioret = 1 Then
         DBGrid1.Text = atrim(formseleccio.Data1.Recordset!Alias)
        End If
        Unload formseleccio
   End If
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
  Data1.RecordSource = DBGrid1.tag + " order by " + DBGrid1.Columns(ColIndex).DataField
  refrescar
End Sub


Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'If DBGrid1.AllowUpdate = False Then MsgBox "Primer has d'editar o afegir": Exit Sub
'If KeyCode = 113 Then
'  If LCase(Data1.Recordset.Fields(DBGrid1.Col).Name) Like "*familia*" Then
'    r = triar_familia
'    DoEvents
'    If r <> DBGrid1.Text Then DBGrid1.Text = "": DBGrid1.Text = r
'    DoEvents
'    formaltarep.SetFocus
'    DBGrid1.SetFocus
'  End If
'End If
End Sub
Function triar_familia() As Integer
  Load formseleccio
  nomf = LCase(formaltarep.caption)
  formseleccio.caption = "Selecciona un Valor"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  If InStr(1, nomf, "materials") > 0 Then
   If InStr(1, nomf, "famili") > 0 Then
    formseleccio.Data1.RecordSource = "select * from familiesmaterials"
   End If
  End If
  If InStr(1, nomf, "colorants") > 0 Then
    If InStr(1, nomf, "famili") > 0 Then
     formseleccio.Data1.RecordSource = "select * from familiescolorants"
    End If
  End If
  
  If InStr(1, nomf, "aditius") > 0 Then
    If InStr(1, nomf, "famili") > 0 Then
     formseleccio.Data1.RecordSource = "select * from familiesaditius"
    End If
  End If
  If formseleccio.Data1.RecordSource = "" Then triar_familia = 0: Exit Function
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   triar_familia = atrim(formseleccio.Data1.Recordset!codi)
  End If
  Unload formseleccio
End Function
Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
 Dim t As Integer
  t = tamany_camp(Data1.Recordset.Fields(DBGrid1.Columns(DBGrid1.col).DataField))
  If Len(DBGrid1.Text) >= t And KeyAscii > 47 Then KeyAscii = 0
 If atrim(colsbloc) <> "" And InStr(1, DBGrid1.col, colsbloc) Then KeyAscii = 0
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err
If Data1.Recordset.Fields(DBGrid1.col).Name = "familia" Then
   status = "PREM F2 PER SEL.LECCIONA FAMILIA"
 Else: status = ""
End If
If DBGrid1.Bookmark <> LastRow Then
 If Data1.Recordset.EditMode = 0 Then
     DBGrid1.AllowUpdate = False
 End If
End If


If Data1.Recordset.Fields(DBGrid1.Columns(DBGrid1.col).DataField).ValidationText <> "" Then status = Data1.Recordset.Fields(DBGrid1.Columns(DBGrid1.col).DataField).ValidationText

Exit Sub
err:
End Sub

Private Sub DBGrid2_BeforeDelete(Cancel As Integer)
   If InputBox("Borrar aquesta subfamilia implica canvis de relació amb els materials," + Chr(13) + Chr(10) + "ESCRIU [Eliminar] PER COMFIRMAR L'ELIMINACIÓ", "Atenció") <> "Eliminar" Then
       Cancel = True
   End If
   formaltafamilies.SetFocus
End Sub

Private Sub DBGrid2_DblClick()
  If DBGrid2.col = 3 Then
     canvicolor_Click
  End If
End Sub

Private Sub DBGrid2_HeadClick(ByVal ColIndex As Integer)
  subfamilies.RecordSource = "select * from " + Data1.tag + " where codifam=" + atrim(cadbl(Data1.Recordset!codi)) + " order by " + DBGrid2.Columns(ColIndex).DataField
  subfamilies.Refresh
End Sub

Private Sub DBGrid2_OnAddNew()
  Dim rst As Recordset
  Dim gran As Long
  If Data1.Recordset.EOF Or cadbl(Data1.Recordset!codi) = 0 Then Exit Sub
  Set rst = dbtintes.OpenRecordset("select max(codi) as gran from " + Data1.tag)
  If Not rst.EOF Then
   gran = cadbl(rst!gran)
   DBGrid2.Columns(0).Text = gran + 1
   DBGrid2.Columns(1).Text = atrim(cadbl(Data1.Recordset!codi))
  End If
  Set rst = Nothing
End Sub

Private Sub Form_Activate()
Frame1.width = DBGrid1.width
sortir.Left = Frame1.width - sortir.width - 75

End Sub

Private Sub Form_Click()
  DBGrid1.col = DBGrid1.Columns.Count - 1
  While Not DBGrid1.Columns(2).visible
     formaltarep.width = formaltarep.width + 100
  Wend
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
 If KeyAscii = 13 Then SendKeys "{TAB}": KeyCode = 0
If KeyAscii > 50 Then KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub Form_Load()
  centerscreen Me
  DBGrid1.ReBind
  
End Sub

Private Sub Form_Resize()
DBGrid1.width = formaltafamilies.width - 285
Frame1.width = DBGrid1.width
sortir.Left = Frame1.width - sortir.width - 75

End Sub

Private Sub Form_Unload(Cancel As Integer)
seleccioret = 0
End Sub

Private Sub modificar_Click()
  If Not Data1.Recordset.EOF Then
   DBGrid1.Enabled = True
   Data1.Recordset.Edit
   DBGrid1.AllowUpdate = True
   DBGrid1.col = 0
   DBGrid1.SetFocus
  End If
End Sub

Private Sub sortir_Click()
 Unload Me
End Sub

Sub acceptar()
  gravar_registre
End Sub

Private Sub Text1_Change()
Text1.tag = "1"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'If KeyCode = 65 And Not DBGrid1.AllowUpdate Then alta_registre: KeyCode = 0
'If KeyCode = 69 Then buscar_registre
If KeyCode = 27 Then cancelar_registre
If KeyCode = 112 Then gravar_registre
'If KeyCode = 13 Then SendKeys "{TAB}": KeyCode = 0
' If KeyCode = 113 And DBGrid1.AllowUpdate Then
'    If colsbloc <> "" And InStr(1, atrim(DBGrid1.Col), colsbloc) Then
'     DBGrid1.Text = triar_mesura
'    End If
' End If
End Sub
Function triar_mesura() As String
  Load formseleccio
  formseleccio.caption = "Selecciona un Valor"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from mesureslineals"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   triar_mesura = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
End Function

Sub alta_registre()
 If DBGrid1.AllowUpdate = False Then
      DBGrid1.Enabled = True
      Data1.Recordset.AddNew
      DoEvents
      DBGrid1.col = 0
      'If autonum.Caption <> "" Then
         'busco el mes gran i el poso a codi +1
        Set rsttmp = dbtintes.OpenRecordset(DBGrid1.tag + " order by codi")
        If Not rsttmp.EOF Then
          rsttmp.MoveLast
          DBGrid1.Columns(0).Text = atrim(cadbl(rsttmp!codi) + 1)
              Else: DBGrid1.Columns(0).Text = "1"
        End If
        DBGrid1.col = 1
      'End If
      DBGrid1.AllowUpdate = True
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
  On Error Resume Next
  Unload subbusqueda
End Sub

Private Sub eliminar_Click()
 On Error GoTo err
  If MsgBox("Segur que vols Eliminar aquesta familia i subfamilies vinculades?", vbYesNo + vbCritical, "Atenció") = 6 Then
   If InputBox("Has d'escriure [Eliminar] per eliminar aquesta familia i subfamilies.", "Atenció") = "Eliminar" Then
    subfamilies.Refresh
    While Not subfamilies.Recordset.EOF
     subfamilies.Recordset.Delete
     subfamilies.Recordset.MoveNext
    Wend
    Data1.Recordset.Delete
    Data1.Recordset.MoveNext
    If Data1.Recordset.EOF Then Data1.Recordset.MovePrevious
   End If
  End If
 Exit Sub
err:
  MsgBox "No s'ha pogut eliminar possiblement perque tingui registres relacionats. O bé no hi ha res per eliminar."
End Sub

Private Sub subfamilies_Reposition()
   Dim rstcolor As Recordset
   Dim color As Long
   color = 15
   If Not subfamilies.Recordset.EOF And Not subfamilies.Recordset.BOF Then
      
      Set rstcolor = dbtintes.OpenRecordset("select * from colorsetiquetes where nomcolor='" + atrim(subfamilies.Recordset!color) + "'")
      If Not rstcolor.EOF Then color = rstcolor!codicolor
      
   End If
   canvicolor.BackColor = QBColor(color)
End Sub

Private Sub Timer1_Timer()
 estattaula.caption = textestattaula(Data1.EditMode)
  If estattaula.ForeColor <> QBColor(0) Then
     estattaula.ForeColor = QBColor(0)
    Else: estattaula.ForeColor = QBColor(14)
  End If
End Sub
