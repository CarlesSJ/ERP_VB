VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formaltafamilies 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "formaltafamilies.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   60
      TabIndex        =   34
      Top             =   5475
      Width           =   5550
      Begin VB.ComboBox combocolors 
         DataField       =   "color"
         DataSource      =   "subfamilies"
         Height          =   315
         ItemData        =   "formaltafamilies.frx":0442
         Left            =   1455
         List            =   "formaltafamilies.frx":0458
         TabIndex        =   36
         Top             =   165
         Width           =   3270
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "formaltafamilies.frx":0486
         Height          =   2415
         Left            =   90
         OleObjectBlob   =   "formaltafamilies.frx":049C
         TabIndex        =   35
         Top             =   660
         Width           =   5370
      End
      Begin VB.Label etpostit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Height          =   195
         Left            =   5415
         TabIndex        =   38
         Top             =   465
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label7 
         Caption         =   "Color SubFamilia:"
         Height          =   240
         Left            =   75
         TabIndex        =   37
         Top             =   195
         Width           =   1215
      End
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
      Bindings        =   "formaltafamilies.frx":138F
      Height          =   4875
      Left            =   105
      Negotiate       =   -1  'True
      OleObjectBlob   =   "formaltafamilies.frx":139F
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   5460
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   135
      TabIndex        =   0
      Top             =   -30
      Width           =   4905
      Begin VB.CommandButton Command2 
         BackColor       =   &H008080FF&
         Caption         =   "Tolerancia Cº "
         Height          =   360
         Left            =   3105
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Tolerancia temperatures per material"
         Top             =   165
         Width           =   1275
      End
      Begin VB.CommandButton alta 
         Height          =   390
         Left            =   75
         Picture         =   "formaltafamilies.frx":1D6E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Alta  Registres"
         Top             =   135
         Width           =   390
      End
      Begin VB.CommandButton eliminar 
         Height          =   390
         Left            =   855
         Picture         =   "formaltafamilies.frx":22F8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Eliminacio Registres"
         Top             =   135
         Width           =   390
      End
      Begin VB.CommandButton modificar 
         Height          =   390
         Left            =   465
         Picture         =   "formaltafamilies.frx":2882
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Consulta Registres"
         Top             =   135
         Width           =   390
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   4395
         Picture         =   "formaltafamilies.frx":2E0C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Height          =   390
         Left            =   1245
         Picture         =   "formaltafamilies.frx":3396
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
         Left            =   1665
         TabIndex        =   7
         Top             =   180
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
   Begin VB.Frame ftolerancies 
      Caption         =   "Tolerancies de Temperatures"
      Height          =   2835
      Left            =   45
      TabIndex        =   12
      Top             =   5505
      Visible         =   0   'False
      Width           =   5040
      Begin VB.Data datatolerancies 
         Caption         =   "datatolerancies"
         Connect         =   "Access"
         DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\COMANDES.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2115
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "materialstoleranciestemp"
         Top             =   315
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.CommandButton Command6 
         Height          =   390
         Left            =   1245
         Picture         =   "formaltafamilies.frx":3920
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton Command5 
         Height          =   390
         Left            =   465
         Picture         =   "formaltafamilies.frx":3EAA
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Consulta Registres"
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton Command4 
         Height          =   390
         Left            =   855
         Picture         =   "formaltafamilies.frx":4434
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   225
         Width           =   390
      End
      Begin VB.CommandButton Command3 
         Height          =   390
         Left            =   75
         Picture         =   "formaltafamilies.frx":49BE
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   225
         Width           =   390
      End
      Begin VB.Frame fdadestolerancies 
         Enabled         =   0   'False
         Height          =   885
         Left            =   90
         TabIndex        =   14
         Top             =   585
         Width           =   4875
         Begin VB.Frame Frame5 
            Caption         =   "Tolerancia Túnel"
            Height          =   645
            Left            =   3210
            TabIndex        =   27
            Top             =   120
            Width           =   1455
            Begin VB.TextBox ttde 
               DataField       =   "toleranciatunelde"
               DataSource      =   "datatolerancies"
               Height          =   285
               Left            =   330
               TabIndex        =   22
               Top             =   240
               Width           =   420
            End
            Begin VB.TextBox tta 
               DataField       =   "toleranciatunela"
               DataSource      =   "datatolerancies"
               Height          =   285
               Left            =   930
               TabIndex        =   23
               Top             =   240
               Width           =   420
            End
            Begin VB.Label Label6 
               Caption         =   "de"
               Height          =   255
               Left            =   75
               TabIndex        =   29
               Top             =   285
               Width           =   210
            End
            Begin VB.Label Label5 
               Caption         =   "a"
               Height          =   195
               Left            =   765
               TabIndex        =   28
               Top             =   270
               Width           =   270
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Tolerancia Tinter"
            Height          =   645
            Left            =   1740
            TabIndex        =   24
            Top             =   120
            Width           =   1455
            Begin VB.TextBox tde 
               DataField       =   "toleranciatinterde"
               DataSource      =   "datatolerancies"
               Height          =   285
               Left            =   315
               TabIndex        =   20
               Top             =   240
               Width           =   420
            End
            Begin VB.TextBox ta 
               DataField       =   "toleranciatintera"
               DataSource      =   "datatolerancies"
               Height          =   285
               Left            =   930
               TabIndex        =   21
               Top             =   240
               Width           =   420
            End
            Begin VB.Label Label4 
               Caption         =   "de"
               Height          =   255
               Left            =   75
               TabIndex        =   26
               Top             =   285
               Width           =   210
            End
            Begin VB.Label Label3 
               Caption         =   "a"
               Height          =   195
               Left            =   765
               TabIndex        =   25
               Top             =   270
               Width           =   270
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Micres"
            Height          =   645
            Left            =   270
            TabIndex        =   15
            Top             =   120
            Width           =   1455
            Begin VB.TextBox ma 
               DataField       =   "micresa"
               DataSource      =   "datatolerancies"
               Height          =   285
               Left            =   930
               TabIndex        =   19
               Top             =   240
               Width           =   420
            End
            Begin VB.TextBox mde 
               DataField       =   "micresde"
               DataSource      =   "datatolerancies"
               Height          =   285
               Left            =   315
               TabIndex        =   18
               Top             =   240
               Width           =   420
            End
            Begin VB.Label Label2 
               Caption         =   "a"
               Height          =   195
               Left            =   765
               TabIndex        =   17
               Top             =   270
               Width           =   270
            End
            Begin VB.Label Label1 
               Caption         =   "de"
               Height          =   255
               Left            =   75
               TabIndex        =   16
               Top             =   285
               Width           =   210
            End
         End
      End
      Begin MSDBGrid.DBGrid reixatolerancies 
         Bindings        =   "formaltafamilies.frx":4F48
         Height          =   1245
         Left            =   75
         OleObjectBlob   =   "formaltafamilies.frx":4F62
         TabIndex        =   13
         Top             =   1530
         Width           =   4875
      End
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
  alta_registre
End Sub

Private Sub combocolors_Click()
   If LCase(Data1.Tag) <> "subfamiliesmaterials" Then MsgBox "Aqui no pots triar color": Exit Sub
   If subfamilies.Recordset.EditMode = 0 And Not subfamilies.Recordset.EOF Then subfamilies.Recordset.Edit
   subfamilies.Recordset.Update
   If Not subfamilies.Recordset.EOF Then subfamilies.Recordset.Move 0
End Sub

Private Sub combocolors_GotFocus()
'  If Data1.Recordset.EditMode = 0 Then DBGrid1.SetFocus
End Sub

Private Sub combocolors_LostFocus()
   possarcolors (combocolors)
End Sub

Private Sub Command1_Click()
  acceptar
End Sub


Private Sub Command2_Click()
  ftolerancies.Visible = Not ftolerancies.Visible
  ftolerancies.ZOrder 0
End Sub

Private Sub Command3_Click()
  If datatolerancies.Recordset.EditMode = 0 Then
     fdadestolerancies.Enabled = True
     datatolerancies.Recordset.AddNew
     datatolerancies.Recordset!codifammaterial = Data1.Recordset!codi
     mde.SetFocus
       Else: MsgBox "Ja estàs editant... primer guarda"
  End If
End Sub

Private Sub Command4_Click()
   If datatolerancies.Recordset.EOF Then Exit Sub
   If datatolerancies.Recordset.EditMode > 0 Then MsgBox "No pots eliminar mentres edites el registre.": Exit Sub
   If MsgBox("Segur que vols eliminar aquesta tolerancia de temperatures?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
       datatolerancies.Recordset.Delete
       datatolerancies.Refresh
   End If
End Sub

Private Sub Command5_Click()
  If datatolerancies.Recordset.EOF Then Exit Sub
  If datatolerancies.Recordset.EditMode = 0 Then
     fdadestolerancies.Enabled = True
     datatolerancies.Recordset.Edit
     mde.SetFocus
       Else: MsgBox "Ja estàs editant... primer guarda"
  End If
End Sub

Private Sub Command6_Click()
  If datatolerancies.Recordset.EditMode = 0 Then MsgBox "No estàs editant...": Exit Sub
   fdadestolerancies.Enabled = False
   datatolerancies.Recordset.Update
   datatolerancies.Refresh
End Sub

Private Sub Data1_Reposition()
 If DBGrid1.Tag = "" Then DBGrid1.Tag = Data1.RecordSource
  If Not Data1.Recordset.EOF And Data1.Tag <> "" Then
    Data1.Database.Execute "delete * from " + Data1.Tag + " where codi=null"
    If LCase(Data1.Tag) = "subfamiliesmaterials" Then
       subfamilies.RecordSource = "select * from " + Data1.Tag + " where codifam=" + atrim(cadbl(Data1.Recordset!codi))
       combocolors.Visible = True
      Else:
         subfamilies.RecordSource = "select codi,codifam,descripcio,'S/C' as color from " + Data1.Tag + " where codifam=" + atrim(cadbl(Data1.Recordset!codi))
         combocolors.Visible = False
    End If
    subfamilies.Refresh
    
    datatolerancies.RecordSource = "select * from materialstoleranciestemp where codifammaterial=" + atrim(cadbl(Data1.Recordset!codi)) + " order by micresde"
    datatolerancies.Refresh
  End If
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
  Data1.RecordSource = DBGrid1.Tag + " order by " + DBGrid1.Columns(ColIndex).DataField
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
  If DBGrid2.Columns(DBGrid2.col).Caption = "Compat." Then
       v = IIf(MsgBox("Vols utilitzar aquesta familia com a COMPATIBLE a la reserva de material?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes, "S", "N")
       DBGrid2 = v
       DBGrid2.EditActive = False
  End If
End Sub

Private Sub DBGrid2_HeadClick(ByVal ColIndex As Integer)
  subfamilies.RecordSource = "select * from " + Data1.Tag + " where codifam=" + atrim(cadbl(Data1.Recordset!codi)) + " order by " + DBGrid2.Columns(ColIndex).DataField
  subfamilies.Refresh
End Sub

Private Sub DBGrid2_LostFocus()
etpostit.Visible = False
End Sub

Private Sub DBGrid2_OnAddNew()
  Dim rst As Recordset
  Dim gran As Long
  If Data1.Recordset.EOF Or cadbl(Data1.Recordset!codi) = 0 Then MsgBox "En el codi 0 no pots afegir subfamilies": Exit Sub
  Set rst = dbtmp.OpenRecordset("select max(codi) as gran from " + Data1.Tag)
  If Not rst.EOF Then
   gran = cadbl(rst!gran)
   DBGrid2.Columns(0).Text = gran + 1
   DBGrid2.Columns(1).Text = atrim(cadbl(Data1.Recordset!codi))
  End If
  Set rst = Nothing
End Sub

Private Sub DBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   possarcolors (combocolors)
   If DBGrid2.Columns(DBGrid2.col).Caption = "Compat." Then
         etpostit.Visible = True: etpostit = "Fer dos clics per canviar el valor."
          Else: etpostit.Visible = False
          
   End If
End Sub

Private Sub Form_Activate()
Frame1.Width = DBGrid1.Width
sortir.Left = Frame1.Width - sortir.Width - 75

End Sub
Function possarcolors(color As String) As String
  Dim codicolor As Double
  codicolor = QBColor(15)
  Select Case color
    Case "VERD"
       codicolor = QBColor(10)
    Case "TARONJA"
       codicolor = &H62B1F2
    Case "BLAU"
       codicolor = QBColor(9)
    Case "ROSA"
       codicolor = &HC78DFA
    Case "GROC"
       codicolor = QBColor(6)
    Case "VERMELL"
       codicolor = QBColor(12)
    Case "BLANC"
       codicolor = QBColor(15)
    Case Else
          codicolor = QBColor(15)
  End Select
  combocolors.BackColor = codicolor
End Function

Private Sub Form_Click()
  DBGrid1.col = DBGrid1.Columns.Count - 1
  While Not DBGrid1.Columns(2).Visible
     formaltarep.Width = formaltarep.Width + 100
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
DBGrid1.Width = formaltarep.Width - 285
Frame1.Width = DBGrid1.Width
sortir.Left = Frame1.Width - sortir.Width - 75

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
  seleccioret = 1
  Me.Hide
End Sub

Private Sub Text1_Change()
Text1.Tag = "1"
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
      'If autonum.Caption <> "" Then
         'busco el mes gran i el poso a codi +1
        Set rsttmp = dbtmp.OpenRecordset(DBGrid1.Tag + " order by codi")
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
   If tearticlesrelacionats(Data1.Recordset!codi) Then MsgBox "Aquesta familia te articles relacionats, no pots eliminar-la.", vbCritical, "Error": GoTo fi
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
fi:
  
 Exit Sub
err:
  MsgBox "No s'ha pogut eliminar possiblement perque tingui registres relacionats. O bé no hi ha res per eliminar."
End Sub
Function tearticlesrelacionats(vcodifam As Double) As Boolean
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select top 1 * from materials where familia=" + atrim(vcodifam))
  If Not rst.EOF Then tearticlesrelacionats = True
  Set rst = Nothing
End Function

Private Sub Timer1_Timer()
 estattaula.Caption = textestattaula(Data1.EditMode)
  If estattaula.ForeColor <> QBColor(0) Then
     estattaula.ForeColor = QBColor(0)
    Else: estattaula.ForeColor = QBColor(14)
  End If
End Sub
