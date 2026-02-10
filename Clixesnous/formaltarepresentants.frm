VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formaltarep 
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   Icon            =   "formaltarepresentants.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
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
      Bindings        =   "formaltarepresentants.frx":0442
      Height          =   5340
      Left            =   105
      Negotiate       =   -1  'True
      OleObjectBlob   =   "formaltarepresentants.frx":0452
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   75
      TabIndex        =   0
      Top             =   -45
      Width           =   4710
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "formaltarepresentants.frx":0E21
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   975
         Picture         =   "formaltarepresentants.frx":13AB
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   525
         Picture         =   "formaltarepresentants.frx":1935
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Consulta Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   4050
         Picture         =   "formaltarepresentants.frx":1EBF
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Height          =   390
         Left            =   3600
         Picture         =   "formaltarepresentants.frx":2449
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
         Left            =   2415
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
   Begin VB.Label etcoditintaalternativa 
      Caption         =   "Label1"
      Height          =   105
      Left            =   4905
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   75
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
Attribute VB_Name = "formaltarep"
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

Private Sub Command1_Click()
  acceptar
End Sub


Private Sub Data1_Reposition()
 If DBGrid1.tag = "" Then DBGrid1.tag = Data1.RecordSource

End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
  Data1.RecordSource = DBGrid1.tag + " order by " + DBGrid1.Columns(ColIndex).DataField
  refrescar
End Sub


Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If DBGrid1.AllowUpdate = False Then MsgBox "Primer has d'editar o afegir": Exit Sub
If KeyCode = 113 Then
  If LCase(Data1.Recordset.Fields(DBGrid1.col).Name) Like "*familia*" Then
    r = triar_familia
    DoEvents
    If r <> DBGrid1.Text Then DBGrid1.Text = "": DBGrid1.Text = r
    DoEvents
    formaltarep.SetFocus
    DBGrid1.SetFocus
  End If
End If
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

Private Sub Form_Activate()
Frame1.width = DBGrid1.width
sortir.Left = Frame1.width - sortir.width - 75

End Sub

Private Sub Form_Click()
  'DBGrid1.Col = DBGrid1.Columns.Count - 1
  'While Not DBGrid1.Columns(1).Visible
  '   formaltarep.Width = formaltarep.Width + 100
  'Wend
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
DBGrid1.width = formaltarep.width - 285
Frame1.width = DBGrid1.width
sortir.Left = Frame1.width - sortir.width - 75

End Sub

Private Sub Form_Unload(Cancel As Integer)
seleccioret = 0
End Sub

Private Sub Label1_Click()

End Sub

Private Sub modificar_Click()
  If Not Data1.Recordset.EOF Then
   If InStr(1, LCase(Data1.RecordSource), " tintes_alternatives ") > 0 Then MsgBox "NO POTS EDITAR, ESBORRA I TORNA A AFEGIR EL QUE CORRESPONGUI.", vbCritical, "ATENCIÓ": Exit Sub
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
  seleccioret = 1
  Me.Hide
End Sub

Private Sub Text1_Change()
Text1.tag = "1"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 65 And Not DBGrid1.AllowUpdate Then alta_registre: KeyCode = 0
'If KeyCode = 69 Then buscar_registre
If KeyCode = 27 Then cancelar_registre
If KeyCode = 112 Then gravar_registre
'If KeyCode = 13 Then SendKeys "{TAB}": KeyCode = 0
 If KeyCode = 113 And DBGrid1.AllowUpdate Then
    If colsbloc <> "" And InStr(1, atrim(DBGrid1.col), colsbloc) Then
     DBGrid1.Text = triar_mesura
    End If
 End If
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
      If autonum.caption <> "" Then
        'busco el mes gran i el poso a codi +1
        Set rsttmp = dbtmp.OpenRecordset("select max(codi) as [grancodi] from " + autonum.caption)
        If Not rsttmp.EOF Then
          DBGrid1.Columns(0).Text = atrim(cadbl(rsttmp!grancodi) + 1)
              Else: DBGrid1.Columns(0).Text = "1"
        End If
        DBGrid1.col = 1
      End If
      DBGrid1.AllowUpdate = True
      If InStr(1, LCase(Data1.RecordSource), " tintes_alternatives ") > 0 Then
           If Data1.Recordset.RecordCount > 2 Then
                 MsgBox "No es pot entrar mes de 3 tintes alternatives.", vbCritical, "Error"
                 Data1.Recordset.CancelUpdate
                 Exit Sub
           End If
           Data1.Recordset!id_tinter = cadbl(formaltarep.tag)
           DBGrid1.AllowUpdate = False
           
           If etcoditintaalternativa <> "" Then triartinta_semblant etcoditintaalternativa
           triartinta
           If DBGrid1.Columns(1) = "" Or tintajaentrada(DBGrid1.Columns(1)) Then Data1.Recordset.CancelUpdate: Exit Sub
           Data1.Recordset.Update
           DBGrid1.Refresh
      End If
      DBGrid1.SetFocus
 End If
End Sub
Sub triartinta_semblant(vcoditinta As String)
  Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim caigudes As Double
  Dim vsql2 As String
  vsql2 = "SELECT distinct Tintes_alternatives.coditinta FROM Tintes_alternatives INNER JOIN Tintesclixesnous ON Tintes_alternatives.id_tinter =Tintesclixesnous.id_tinter WHERE (((Tintesclixesnous.coditinta)='" + atrim(vcoditinta) + "'))"
  sql = "SELECT  codi,descripcio,referenciacolor,nominterndelbido as Bidó from tintes_tot where codi in(" + vsql2 + ")"
  Clipboard.Clear
  Clipboard.SetText sql
  were = " order by descripcio"
  Load formseleccio
  formseleccio.Data1.DatabaseName = rutadelfitxer(camiclixes) + "tintes.mdb"
  formseleccio.Data1.RecordSource = sql + were
  formseleccio.width = 14000
  formseleccio.sortirs.tag = "filtre"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 500
  formseleccio.colocar_botofiltre 0
  If formseleccio.Data1.Recordset.EOF Then GoTo fi
  formseleccio.Show 1
  If seleccioret = 1 Then
    DBGrid1.Columns(2) = atrim(formseleccio.Data1.Recordset!descripcio)
    DBGrid1.Columns(1) = atrim(formseleccio.Data1.Recordset!codi)
  End If
  If seleccioret = 9 Then
    DBGrid1.Columns(2) = ""
    DBGrid1.Columns(1) = ""
  End If
 '  Data1.Recordset!client = Text2.Text
 '  nomclient.Caption = atrim(formseleccio.Data1.Recordset!nom)
  
 ' End If
fi:
  Unload formseleccio
End Sub

Function tintajaentrada(vcoditinta As String) As Boolean
  Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from tintes_alternatives where id_tinter=" + atrim(cadbl(formaltarep.tag)))
   rst.FindFirst "coditinta='" + atrim(vcoditinta) + "'"
   If Not rst.NoMatch Then tintajaentrada = True
   If tintajaentrada = False Then
     Set rst = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(cadbl(formaltarep.tag)))
     rst.FindFirst "coditinta='" + atrim(vcoditinta) + "'"
     If Not rst.NoMatch Then tintajaentrada = True
   End If
   If tintajaentrada Then MsgBox "Aquesta tinta ja està entrada en aquest tinter.", vbCritical, "Error"
  Set rst = Nothing
End Function

Sub triartinta()
  Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim caigudes As Double
  
  sql = "SELECT  codi,descripcio,referenciacolor,nominterndelbido as Bidó from tintes_tot "
  were = " order by descripcio"
  Load formseleccio
  formseleccio.Data1.DatabaseName = rutadelfitxer(camiclixes) + "tintes.mdb"
  formseleccio.Data1.RecordSource = sql + were
  formseleccio.width = 14000
  formseleccio.sortirs.tag = "filtre"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 500
  formseleccio.colocar_botofiltre 0
  formseleccio.Show 1
  If seleccioret = 1 Then
    DBGrid1.Columns(2) = atrim(formseleccio.Data1.Recordset!descripcio)
    DBGrid1.Columns(1) = atrim(formseleccio.Data1.Recordset!codi)
  End If
  If seleccioret = 9 Then
    DBGrid1.Columns(2) = ""
    DBGrid1.Columns(1) = ""
  End If
 '  Data1.Recordset!client = Text2.Text
 '  nomclient.Caption = atrim(formseleccio.Data1.Recordset!nom)
  
 ' End If
  Unload formseleccio
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
 estattaula.caption = textestattaula(Data1.EditMode)
  If estattaula.ForeColor <> QBColor(0) Then
     estattaula.ForeColor = QBColor(0)
    Else: estattaula.ForeColor = QBColor(14)
  End If
End Sub
