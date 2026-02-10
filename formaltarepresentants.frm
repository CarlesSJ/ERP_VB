VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formaltarep 
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   Icon            =   "formaltarepresentants.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox opcions 
      Height          =   315
      ItemData        =   "formaltarepresentants.frx":0442
      Left            =   1875
      List            =   "formaltarepresentants.frx":0458
      TabIndex        =   11
      Top             =   1890
      Visible         =   0   'False
      Width           =   1605
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
      Bindings        =   "formaltarepresentants.frx":0486
      Height          =   5340
      Left            =   120
      Negotiate       =   -1  'True
      OleObjectBlob   =   "formaltarepresentants.frx":0496
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
      Width           =   4950
      Begin VB.CommandButton bactiva 
         Caption         =   "Activa/Inactiva"
         Height          =   315
         Left            =   2070
         TabIndex        =   13
         Top             =   195
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "formaltarepresentants.frx":0E65
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Alta  Registres"
         Top             =   165
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   990
         Picture         =   "formaltarepresentants.frx":13EF
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Eliminacio Registres"
         Top             =   165
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   525
         Picture         =   "formaltarepresentants.frx":1979
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Consulta Registres"
         Top             =   165
         Width           =   420
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   4440
         Picture         =   "formaltarepresentants.frx":1F03
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Alta  Registres"
         Top             =   165
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Height          =   390
         Left            =   3990
         Picture         =   "formaltarepresentants.frx":248D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   165
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label lvariable 
         Caption         =   "Label1"
         Height          =   240
         Left            =   3195
         TabIndex        =   12
         Top             =   225
         Visible         =   0   'False
         Width           =   330
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
   Begin VB.Label status 
      BackColor       =   &H0000FFFF&
      Height          =   225
      Left            =   60
      TabIndex        =   9
      Top             =   5985
      Width           =   7170
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
  If Me.Tag <> "" Then
     factor = cadbl(Me.Tag)
    Else: factor = 130
  End If
  
 On Error GoTo fi
 For i = 0 To 50
   tipusdato = Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Type
   grandato = Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Size
   
   If grandato < 5 Then grandato = 5
   DBGrid1.Columns(i).Width = grandato * (IIf(tipusdato <> 10, factor, 130))
   DBGrid1.Columns(i).Caption = UCase(DBGrid1.Columns(i).Caption)
 Next i
  
         
  
fi:

If autonum.Caption <> "" Then
     DBGrid1.Columns(0).Width = 0
     Set dbtmp = OpenDatabase(Data1.DatabaseName)
  End If
  
End Sub

Private Sub alta_Click()
   If formaltarep.Tag = "tipuspaperfrontal" Then
        MsgBox "El format d'entrada de la descripció es el següent: " + Chr(10) + " ex: 1 Paper frontal en un costat. [un numero (1-9) espai i la descripció]", vbInformation, "Error"
   End If
   
  alta_registre
End Sub

Private Sub bactiva_Click()
  If Not Data1.Recordset.EOF Then
     Data1.Recordset.Edit
     Data1.Recordset!estat = IIf(atrim(Data1.Recordset!estat) = "ACTIVA", "INACTIVA", "ACTIVA")
     Data1.Recordset.Update
     Data1.Recordset.Move 0
  End If
End Sub

Private Sub Command1_Click()
  acceptar
End Sub


Private Sub Data1_Reposition()
 If DBGrid1.Tag = "" Then DBGrid1.Tag = Data1.RecordSource

End Sub

Private Sub DBGrid1_BeforeUpdate(Cancel As Integer)
   If formaltarep.alta.Tag = "codifabricacio" Then
     'If Len(DBGrid1.Columns(2)) < 11 Or Len(DBGrid1.Columns(2)) > 11 Then
     '   MsgBox "El camp de codicomptable ha de ser de 11 digits", vbInformation, "Atenció"
     '   Cancel = 1
     '   DBGrid1.col = 2
     '   DBGrid1.SetFocus
     'End If
     If MsgBox("Vols que aquest codi comptable sigui el predeterminat d'aquest client?", vbInformation + vbYesNo, "Atenció") = vbYes Then
          Data1.Recordset!predeterminat = 1
          Data1.Database.Execute "update clients_codiscomptables set predeterminat=false where codifabricacio=" + atrim(formaltarep.alta.HelpContextID)
     End If
   End If
   If formaltarep.Tag = "tipuspaperfrontal" Then
     If cadbl(Mid(DBGrid1.Columns(1), 1, 2)) > 9 Or cadbl(Mid(DBGrid1.Columns(1), 1, 2)) < 1 Then
        MsgBox "El format d'entrada de la descripció es el següent: " + Chr(10) + " ex: 1 Paper frontal en un costat. [un numero (1-9) espai i la descripció]", vbCritical, "Error"
        Cancel = 1
     End If
   End If
End Sub

Private Sub DBGrid1_ButtonClick(ByVal ColIndex As Integer)
  triarcolor
End Sub

Private Sub DBGrid1_DblClick()
    triarcolor
End Sub
Sub triarcolor()
  If formaltarep.Caption = "Manteniment de Adhesius" And DBGrid1.Columns(DBGrid1.col).Caption = "COLOR" Then
    If Data1.Recordset.EditMode > 0 Then
        opcions.Visible = True
        'opcions.Left = DBGrid1.Columns(ColIndex).Left + 30
        opcions.Top = DBGrid1.RowTop(DBGrid1.row) + DBGrid1.Top
        opcions.Left = DBGrid1.Columns(DBGrid1.col).Left + DBGrid1.Left
        'opcions.Left = DBGrid1.RowTop(DBGrid1.Row) + DBGrid1.Left
        opcions.SetFocus
        SendKeys ("%{DOWN}")
    End If
  End If
  If formaltarep.Tag = "codifabricacio" And UCase(DBGrid1.Columns(DBGrid1.col).DataField) = "MONEDA" Then
    opcions.Clear
    opcions.AddItem "Euros"
    opcions.AddItem "Dolars"
    If Data1.Recordset.EditMode > 0 Then
        opcions.Visible = True
        'opcions.Left = DBGrid1.Columns(ColIndex).Left + 30
        opcions.Top = DBGrid1.RowTop(DBGrid1.row) + DBGrid1.Top
        opcions.Left = DBGrid1.Columns(DBGrid1.col).Left + DBGrid1.Left
        'opcions.Left = DBGrid1.RowTop(DBGrid1.Row) + DBGrid1.Left
        opcions.SetFocus
        SendKeys ("%{DOWN}")
    End If
  End If
  
End Sub
Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
  Data1.RecordSource = DBGrid1.Tag + " order by " + DBGrid1.Columns(ColIndex).DataField
  refrescar
End Sub


Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim v As String
 Dim m As String
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
  If Data1.Recordset.Fields(DBGrid1.col).Name = "codicomptable" Then
    v = triar_codicomptable(m)
    DoEvents
    If v <> DBGrid1.Text Then DBGrid1.Text = "": DBGrid1.Text = v
    If m <> "" Then DBGrid1.Columns("Moneda") = m
    DBGrid1.col = 3: DBGrid1.Text = r
    DoEvents
    formaltarep.SetFocus
    DBGrid1.SetFocus
  End If
  If Data1.Recordset.Fields(DBGrid1.col).Name = "predeterminat" Then
    If Data1.Recordset.EditMode = 0 Then MsgBox "Primer has d'estar editant.": Exit Sub
    If MsgBox("Vols passar aquest codi comptable com a predeterminat?", vbInformation + vbYesNo, "Atenció") = vbNo Then Exit Sub
    Data1.Database.Execute "update clients_codiscomptables set predeterminat=false where codifabricacio=" + atrim(formaltarep.alta.HelpContextID)
    Data1.Recordset!predeterminat = True
    Data1.Recordset.Update
    DBGrid1.Refresh
    DBGrid1.col = 2
    DoEvents
    formaltarep.SetFocus
    DBGrid1.SetFocus
  End If
End If
If Data1.Recordset.Fields(DBGrid1.col).Name = "codicomptable" Then KeyCode = 0
End Sub
Function triar_codicomptable(m As String) As Double
  Load formseleccio
  nomf = LCase(formaltarep.Caption)
  formseleccio.Caption = "Selecciona un Valor"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select codisap,nomclient,moneda from clients_codissap"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 1500
  formseleccio.DBGrid2.Columns(1).Width = 3500
  formseleccio.DBGrid2.Columns(2).Width = 500
  'formseleccio.Command3.tag = "filtre"
  formseleccio.Show 1
  If seleccioret = 1 Then
   triar_codicomptable = atrim(formseleccio.Data1.Recordset!codisap)
   r = atrim(formseleccio.Data1.Recordset!nomclient)
   m = atrim(formseleccio.Data1.Recordset!moneda)
  End If
  Unload formseleccio
End Function
Function triar_familia() As Integer
  Load formseleccio
  nomf = LCase(formaltarep.Caption)
  formseleccio.Caption = "Selecciona un Valor"
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
 If Data1.Recordset.Fields(DBGrid1.col).Name = "codicomptable" Then KeyAscii = 0
 If Data1.Recordset.Fields(DBGrid1.col).Name = "predeterminat" Then KeyAscii = 0
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err
status = ""
If formaltarep.alta.Tag = "codifabricacio" And Data1.Recordset.Fields(DBGrid1.col).Name = "codicomptable" Then
   status = "PREM F2 PER SEL.LECCIONAR EL CODICOMPTABLE"
     Else: If InStr(1, status, " CODICOMPTABLE") > 0 Then status = ""
End If
If formaltarep.alta.Tag = "codifabricacio" And Data1.Recordset.Fields(DBGrid1.col).Name = "predeterminat" Then
   status = "PREM F2 PER SEL.LECCIONAR COM A CODI COMPTABLE PREDETERMINAT"
     Else: If InStr(1, status, " PREDETERMINAT") > 0 Then status = ""
End If
If Data1.Recordset.Fields(DBGrid1.col).Name = "maquinescompatibles" Then
   status = "ESCRIU LES SOLDADORES COMPATIBLES SEPARAT PER COMES: 2,4,6,7"
   status.Visible = True
End If
If formaltarep.alta.Tag = "codifabricacio" And UCase(Data1.Recordset.Fields(DBGrid1.col).Name) = "MONEDA" Then
  opcions.Visible = False
   opcions.Clear
   opcions.AddItem "Euros"
   opcions.AddItem "Dolars"
   opcions.Text = atrim(Data1.Recordset.Fields(DBGrid1.col))
   opcions.Width = DBGrid1.Columns(DBGrid1.col).Width
   opcions.Top = DBGrid1.RowTop(DBGrid1.row) + DBGrid1.Top
   opcions.Left = DBGrid1.Columns(DBGrid1.col).Left + DBGrid1.Left
   opcions.Visible = True
     Else: opcions.Visible = False
End If
'If formaltarep.alta.Tag = "codifabricacio" And UCase(Data1.Recordset.Fields(DBGrid1.col).Name) = "PREDETERMINAT" Then
'   opcions.Clear
'   opcions.AddItem "Predeterminada"
'   opcions.AddItem ""
'   opcions.Visible = True
'   opcions.Text = IIf(atrim(Data1.Recordset.Fields(DBGrid1.col)) = "1", "Predeterminada", "")
'   opcions.Width = 1000
'   opcions.Top = DBGrid1.RowTop(DBGrid1.row) + DBGrid1.Top
'   opcions.Left = DBGrid1.Columns(DBGrid1.col).Left + DBGrid1.Left
'     Else: opcions.Visible = False
'End If

If Data1.Recordset.Fields(DBGrid1.col).Name = "familia" Then
   status = "PREM F2 PER SEL.LECCIONA FAMILIA"
 Else: If InStr(1, status, " FAMILIA") > 0 Then status = ""
End If
If DBGrid1.Bookmark <> LastRow Then
 If Data1.Recordset.EditMode = 0 Then
     DBGrid1.AllowUpdate = False
 End If
End If
If formaltarep.autonum = "tubbase" And UCase(Data1.Recordset.Fields(DBGrid1.col).Name) = "TIPUSMATERIAL" Then
   status = "ESCRIU C PER CARTRO I P PER PVC"
End If

If Data1.Recordset.Fields(DBGrid1.Columns(DBGrid1.col).DataField).ValidationText <> "" Then status = Data1.Recordset.Fields(DBGrid1.Columns(DBGrid1.col).DataField).ValidationText
If status = "" Then status.BackColor = formaltarep.BackColor Else status.BackColor = QBColor(14)
Exit Sub
err:
End Sub

Private Sub Form_Activate()
  Dim i As Integer
  Frame1.Width = DBGrid1.Width
  sortir.Left = Frame1.Width - sortir.Width - 75
  If alta.Tag = "codifabricacio" Then
        
        For i = 0 To DBGrid1.Columns.Count - 1
           If LCase(DBGrid1.Columns(i).Caption) = "predeterminat" Then
               DBGrid1.Columns(i).Caption = "Pred."
           End If
        Next i
  End If
End Sub

Private Sub Form_Click()
  'DBGrid1.Col = DBGrid1.Columns.Count - 1
  'While Not DBGrid1.Columns(2).Visible
 '    formaltarep.Width = formaltarep.Width + 100
 ' Wend
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
 If KeyAscii = 13 Then SendKeys "{TAB}": KeyCode = 0
If KeyAscii > 50 Then KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub Form_Load()
  lvariable = ""
  centerscreen Me
  DBGrid1.ReBind
  
End Sub

Private Sub Form_Resize()
If formaltarep.Width - 385 > 0 Then DBGrid1.Width = formaltarep.Width - 385
Frame1.Width = DBGrid1.Width
If Frame1.Width - sortir.Width - 75 > 0 Then sortir.Left = Frame1.Width - sortir.Width - 75
If formaltarep.Height - DBGrid1.Top - 800 > 0 Then
    DBGrid1.Height = formaltarep.Height - DBGrid1.Top - 800
    status.Top = DBGrid1.Top + DBGrid1.Height + 50
    status.Width = formaltarep.Width
    status.Left = 100
End If
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
   DBGrid1.col = 0
   DBGrid1.SetFocus
  End If
End Sub

Private Sub opcions_Change()
  guardarvalorcombo
End Sub
Sub guardarvalorcombo()
If opcions.Visible = False Then Exit Sub
If Data1.Recordset.EditMode > 0 Then DBGrid1.Text = opcions.Text
opcions.Visible = False
DBGrid1.col = 2
DoEvents
DBGrid1.SetFocus
SendKeys "{TAB}"

End Sub

Private Sub opcions_Click()
  guardarvalorcombo
End Sub

Private Sub opcions_LostFocus()
   opcions.Visible = False
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
 If KeyCode = 113 And DBGrid1.AllowUpdate Then
    If colsbloc <> "" And InStr(1, atrim(DBGrid1.col), colsbloc) Then
     DBGrid1.Text = triar_mesura
    End If
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
      If alta.Tag <> "" Then
         Data1.Recordset.Fields(alta.Tag) = IIf(lvariable = "", atrim(alta.HelpContextID), lvariable)
      End If

      If DBGrid1.Columns(0).Visible Then
         DBGrid1.col = 0
          Else: DBGrid1.col = 1
      End If
      DBGrid1.SetFocus
      If formaltarep.Tag = "justificants merma" Then altadejustificantmerma
  End If
End Sub
Sub altadejustificantmerma()
    Dim vnumjustificant As String
    Dim vdata As String
    Dim vcolor As String
    Dim vkgjustificant As Double
    vdata = InputBox("Escriu la data del justificant.", "Data")
    vnumjustificant = cadbl(InputBox("Escriu el número del justificant. NOMES NUMEROS", "Nº Justificant"))
    vkgjustificant = cadbl(InputBox("Escriu els KG material del justificant.", "KG Justificant"))
    If vkgjustificant = 0 Or vdata = "" Or vnumjustificant = "" Then MsgBox "S'ha de posar un valor de KG.", vbCritical, "Error": Data1.Recordset.CancelUpdate: DBGrid1.EditActive = False: Exit Sub
    vcolor = UCase(InputBox("Escriu el color del rebuig. [V]Verd,[M]Vermell,[B]Blau"))
    Select Case vcolor
       Case "V"
          vcolor = "DESPVERD01"
          Case "M"
          vcolor = "DESPVERMELL01"
          Case "B"
          vcolor = "DESPBLAU01"
           Case Else
             vcolor = ""
    End Select
    If vcolor = "" Then MsgBox "No s'ha escullit cap color correcte.", vbCritical, "Error": Data1.Recordset.CancelUpdate: DBGrid1.EditActive = False
    Data1.Recordset!datafactura = vdata
    Data1.Recordset!numerofactura = vnumjustificant
    Data1.Recordset!tipus = vcolor
    Data1.Recordset!kgfactura = vkgjustificant
    'DBGrid1.Columns(0).text = vdata
    'DBGrid1.Columns(1).text = vnumjustificant
    'DBGrid1.Columns(2).text = vcolor
   ' DBGrid1.Columns(3).text = vkgjustificant
    demanar_proveidorrecilatge
    gravar_registre
    Data1.Recordset.FindFirst "numerofactura=" + atrim(vnumjustificant) + ""
End Sub
Sub demanar_proveidorrecilatge()
   formaltarep.Data1.RecordSource = "select nomproveidor,nifproveidor from facturesSAPreciclatge order by nomproveidor"
  Load formseleccio
  formseleccio.Caption = "Selecciona un Proveïdor"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select distinCt nomproveidor,nifproveidor from facturesSAPreciclatge order by nomproveidor"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 3500
  formseleccio.DBGrid2.Columns(1).Width = 1500
  'formseleccio.Command3.tag = "filtre"
  formseleccio.Show 1
  If seleccioret = 1 Then
    'DBGrid1.Columns(4).text = atrim(formseleccio.Data1.Recordset!nomproveidor)
    'DBGrid1.Columns(5).text = atrim(formseleccio.Data1.Recordset!nifproveidor)
    Data1.Recordset!nomproveidor = atrim(formseleccio.Data1.Recordset!nomproveidor)
    Data1.Recordset!nifproveidor = atrim(formseleccio.Data1.Recordset!nifproveidor)
  End If
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
  Dim rst As Recordset
  Dim vcodi As Double
 On Error GoTo err
  If MsgBox("Segur que vols Eliminar?", vbYesNo + vbCritical, "Atenció") = 6 Then
    If formaltarep.alta.Tag = "codifabricacio" Then
      If Data1.Recordset!predeterminat Then MsgBox "Abans d'eliminar un codi comptable predeterminat hauries de canviar el codi comptable PREDETERMINAT.", vbCritical, "PREDETERMINAT": Exit Sub
    End If
    If InStr(1, LCase(Data1.RecordSource), "tractamentcares") > 0 Then
      vcodi = cadbl(Data1.Recordset!codi)
      Set rst = dbtmp.OpenRecordset("select * from materials where codidescmatcara1=" + atrim(vcodi) + " or codidescmatcara2=" + atrim(vcodi))
      If Not rst.EOF Then MsgBox "Aquest codi s'utilitza a la taula de materials no pots eliminar-lo.", vbCritical, "Error": GoTo fi
    End If
    Data1.Recordset.Delete
    Data1.Recordset.MoveNext
    If Data1.Recordset.EOF Then Data1.Recordset.MovePrevious
  End If
fi:
  Set rst = Nothing
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
