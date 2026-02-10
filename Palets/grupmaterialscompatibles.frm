VERSION 5.00
Begin VB.Form grupmaterialscompatibles 
   Caption         =   "Grups de materials compatibles"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5685
   Icon            =   "grupmaterialscompatibles.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   60
      Top             =   600
   End
   Begin VB.CommandButton Command2 
      Height          =   360
      Left            =   255
      Picture         =   "grupmaterialscompatibles.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Alta  Registres"
      Top             =   3135
      Width           =   420
   End
   Begin VB.CommandButton sortir 
      Height          =   390
      Left            =   5250
      Picture         =   "grupmaterialscompatibles.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Sortir"
      Top             =   150
      Width           =   390
   End
   Begin VB.Data grupdpalets 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   2910
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   165
      Width           =   2205
   End
   Begin VB.Frame Grups 
      Height          =   4545
      Left            =   135
      TabIndex        =   0
      Top             =   630
      Width           =   5415
      Begin VB.CommandButton Command3 
         Height          =   360
         Left            =   555
         Picture         =   "grupmaterialscompatibles.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Eliminacio Registres"
         Top             =   2505
         Width           =   420
      End
      Begin VB.TextBox cnomdelgrup 
         DataField       =   "nomdelgrup"
         DataSource      =   "grupdpalets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   12
         Top             =   120
         Width           =   3495
      End
      Begin VB.ListBox llistasubmaterials 
         Height          =   1230
         Left            =   90
         TabIndex        =   10
         Top             =   3150
         Width           =   5205
      End
      Begin VB.CommandButton consultar 
         Height          =   360
         Left            =   45
         Picture         =   "grupmaterialscompatibles.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Buscar Registres"
         Top             =   795
         Width           =   420
      End
      Begin VB.Label etcodimaterial 
         BackStyle       =   0  'Transparent
         Caption         =   "__________________________________"
         DataField       =   "codimaterialprincipal"
         DataSource      =   "grupdpalets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   30
         TabIndex        =   14
         Top             =   450
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nom del Grup:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   13
         Top             =   180
         Width           =   1890
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Subfamilies de material compatibles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   165
         TabIndex        =   11
         Top             =   2895
         Width           =   5055
      End
      Begin VB.Label etfamilies 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PEBD - ANTIDESLIZANTE"
         DataField       =   "descripciofamilies"
         DataSource      =   "grupdpalets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   105
         TabIndex        =   9
         Top             =   1365
         Width           =   5055
      End
      Begin VB.Label etmatprincipal 
         BackStyle       =   0  'Transparent
         Caption         =   "__________________________________"
         DataField       =   "descripciomaterialprincipal"
         DataSource      =   "grupdpalets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   615
         TabIndex        =   7
         Top             =   765
         Width           =   4710
      End
   End
   Begin VB.Frame Frame1 
      Height          =   570
      Left            =   60
      TabIndex        =   2
      Top             =   15
      Width           =   5625
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   30
         Picture         =   "grupmaterialscompatibles.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Alta  Registres"
         Top             =   135
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   900
         Picture         =   "grupmaterialscompatibles.frx":213C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Eliminacio Registres"
         Top             =   135
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   465
         Picture         =   "grupmaterialscompatibles.frx":26C6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Edicio del  Registres"
         Top             =   135
         Width           =   420
      End
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   1380
         Picture         =   "grupmaterialscompatibles.frx":2C50
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Acceptar canvis"
         Top             =   150
         Width           =   420
      End
      Begin VB.Label etestat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1860
         TabIndex        =   16
         Top             =   240
         Width           =   1485
      End
   End
End
Attribute VB_Name = "grupmaterialscompatibles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
 Dim rstpalets As Recordset
  Dim elgran As Double
  'If palets.Recordset.EOF Then Exit Sub
  If grupdpalets.Recordset.EditMode > 0 Then MsgBox "Estas editant. Primer finalitza l'edicio.": Exit Sub
  activarframes True
  grupdpalets.Recordset.AddNew
  activarframes True
  cnomdelgrup.SetFocus
  'txtFields(5).SetFocus
End Sub

Private Sub Command1_Click()
  If grupdpalets.Recordset.EditMode > 0 Then
     grupdpalets.Recordset.Update
     grupdpalets.Recordset.Bookmark = grupdpalets.Recordset.LastModified
     activarframes False
  End If
End Sub

Private Sub Command2_Click()
      If grupdpalets.Recordset.EditMode > 0 Then MsgBox "Primer guarda els canvis del material principal.", vbExclamation, "Atenció": Exit Sub
      If cadbl(etcodimaterial) = 0 Then MsgBox "Primer has d'escullir un material principal.", vbCritical, "Error": Exit Sub
      Load formseleccio
      formseleccio.Caption = "Escull familia compatible"
      formseleccio.Data1.DatabaseName = cami
      ordre = " order by proveidor,descripcio"
      formseleccio.Data1.RecordSource = "SELECT codi,descripcio from subfamiliesmaterials where codifam in (select familia from materials where codi=" + atrim(cadbl(etcodimaterial)) + ") order by descripcio"
      formseleccio.refrescar
      'formseleccio.Width = formseleccio.Width + ((formseleccio.Width / 100) * 20)
      formseleccio.DBGrid2.Columns(0).Width = 0
      formseleccio.DBGrid2.Columns(1).Width = 2500
      formseleccio.Command2.Tag = "1"
      formseleccio.Show 1
      If seleccioret = 1 Then
        ' Clipboard.Clear
        ' Clipboard.SetText "insert into grupsmaterialscompatibles_linies (idgrupsdematerialscompatibles,nomsubfamilia,codisubfamilia) value (" + atrim(grupdpalets.Recordset!numerodegrup) + ",'" + atrim(formseleccio.Data1.Recordset!descripcio) + "'," + atrim(formseleccio.Data1.Recordset!codi) + ")"
         dbtmp.Execute "insert into grupsmaterialscompatibles_linies (idgrupsdematerialscompatibles,nomsubfamilia,codisubfamilia) values (" + atrim(grupdpalets.Recordset!numerodegrup) + ",'" + atrim(formseleccio.Data1.Recordset!descripcio) + "'," + atrim(formseleccio.Data1.Recordset!codi) + ")"
         actualitzar_vinculats
         generar_sql_subfamilies
      End If
      Unload formseleccio
End Sub
Sub generar_sql_subfamilies()
   Dim rst As Recordset
   Dim vsql As String
   Set rst = dbtmp.OpenRecordset("select * from grupsmaterialscompatibles_linies where idgrupsdematerialscompatibles=" + atrim(grupdpalets.Recordset!numerodegrup))
   While Not rst.EOF
     vsql = vsql + " or materials.subfamilia=" + atrim(rst!codisubfamilia)
     rst.MoveNext
   Wend
   dbtmp.Execute "update grupsmaterialscompatibles set sqlsubfamilies='" + vsql + "' where numerodegrup=" + atrim(grupdpalets.Recordset!numerodegrup)
   Set rst = Nothing
End Sub


Private Sub Command3_Click()
    eliminar_subfamilia
End Sub

Private Sub consultar_Click()
    Dim vsql As String
    If llistasubmaterials.ListCount > 0 Then MsgBox "Per poder canviar de material has de borrar primer totes les subfamilies perquè sino podria haver una barreja de materials.", vbCritical, "ERROR": Exit Sub

      Load formseleccio
      formseleccio.Caption = "Escull Material Principal"
      formseleccio.Data1.DatabaseName = cami
      ordre = " order by proveidor,descripcio"
      formseleccio.Data1.RecordSource = "SELECT materials.codi as [Codi], materials.descripcio as [Descripcio], materials.refproducte as [RefProducte], proveidors.nom as [Proveidor] FROM materials LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE (((materials.codi)>499)) order by descripcio"
      formseleccio.refrescar
      formseleccio.Width = formseleccio.Width + ((formseleccio.Width / 100) * 20)
      formseleccio.DBGrid2.Columns(0).Width = 500
      formseleccio.DBGrid2.Columns(1).Width = 2500
      formseleccio.DBGrid2.Columns(2).Width = 1000
      formseleccio.DBGrid2.Columns(3).Width = 1500
      formseleccio.Command2.Tag = "2"
      formseleccio.Show 1
      If seleccioret = 1 Then
         etcodimaterial = atrim(formseleccio.Data1.Recordset!codi)
         etmatprincipal = atrim(formseleccio.Data1.Recordset!descripcio)
         etfamilies = generar_subfamilies(formseleccio.Data1.Recordset!codi, vsql)
         grupdpalets.Recordset!sqlprincipal = vsql
      End If
      Unload formseleccio
End Sub
Function generar_subfamilies(vcodi As Double, vsql As String) As String
   Dim rst As Recordset
   
   vsql = "SELECT familiesmaterials.descripcio as descfam, subfamiliesmaterials.descripcio as dessubfam, familiescolorants.descripcio as desccol, subfamiliescolorants.descripcio as descsubcol, familiesaditius.descripcio as descadi, subfamiliesaditius.descripcio as descsubad, materials.codi, materials.descripcio , materials.proveidor, proveidors.nom, materials.refproducte, materials.grmcm3, materials.grmm2, familiesmaterials.codi AS codimat, subfamiliesmaterials.codi AS codisubmat, familiescolorants.codi AS codicol, subfamiliescolorants.codi AS codisubcol, familiesaditius.codi AS codiadi, subfamiliesaditius.codi AS codisubadi "
   vsql = vsql + " FROM proveidors INNER JOIN ((((((materials INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi) INNER JOIN familiescolorants ON materials.familiacol = familiescolorants.codi) INNER JOIN subfamiliescolorants ON materials.subfamiliacol = subfamiliescolorants.codi) INNER JOIN familiesaditius ON materials.familiaad = familiesaditius.codi) INNER JOIN subfamiliesaditius ON materials.subfamiliaad = subfamiliesaditius.codi) ON proveidors.codi = materials.proveidor "
   vsql = vsql + " Where (((materials.codi) =" + atrim(vcodi) + ")) "
   'Clipboard.Clear
   'Clipboard.SetText vsql
   Set rst = dbtmpb.OpenRecordset(vsql)
   If Not rst.EOF Then
       generar_subfamilies = atrim(rst!descfam) + " - " + atrim(rst!dessubfam) + Chr(13) + Chr(10)
       generar_subfamilies = generar_subfamilies + atrim(rst!desccol) + " - " + atrim(rst!descsubcol) + Chr(13) + Chr(10)
       generar_subfamilies = generar_subfamilies + atrim(rst!descadi) + " - " + atrim(rst!descsubad) + Chr(13) + Chr(10)
       vsql = "materials.familia =" + atrim(rst!codimat) + " and materials.familiacol=" + atrim(rst!codicol) + " and materials.subfamiliacol=" + atrim(rst!codisubcol) + " and materials.familiaad=" + atrim(rst!codiadi) + " and materials.subfamiliaad=" + atrim(rst!codisubadi) + " and (materials.subfamilia=" + atrim(rst!codisubmat)
        Else: vsql = ""
   End If
   Set rst = Nothing
End Function
Private Sub eliminar_Click()
  If grupdpalets.Recordset.EOF Then MsgBox "No hi ha registres": Exit Sub
  If hihacomandesambaquestgrup(grupdpalets.Recordset!numerodegrup) Then MsgBox "Hi ha comandes que ja es van imprimir amb aquest Grup de compatibles no pots eliminar-lo sense perdre la relació.", vbCritical, "Error": GoTo fi
  If llistasubmaterials.ListCount > 0 Then MsgBox "Hi ha subfamilies relacionades, primer hauries de eliminar-les.", vbInformation, "Atenció": Exit Sub
  If MsgBox("Segur que vols borrar aquest grup de palets?", vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
     'If Not grupdpalets.Recordset.EOF Then
      grupdpalets.Recordset.Delete
      grupdpalets.Refresh
     'End If
  End If
fi:
  activarframes False
End Sub
Function hihacomandesambaquestgrup(vnumgrup As Double) As Boolean
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select codigrupmaterialcompatible from comandes_extres where codigrupmaterialcompatible=" + atrim(vnumgrup))
  If Not rst.EOF Then
    hihacomandesambaquestgrup = True
  End If
  Set rst = Nothing
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Command1_Click
  If KeyCode = 27 Then
   If grupdpalets.Recordset.EditMode > 0 Then
     grupdpalets.Recordset.CancelUpdate
     activarframes False
   End If
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
   grupdpalets.RecordSource = "grupsmaterialscompatibles"
   grupdpalets.DatabaseName = Form1.palets.DatabaseName
   grupdpalets.Refresh
   If Not grupdpalets.Recordset.EOF Then grupdpalets.Recordset.MoveLast: grupdpalets.Recordset.MoveFirst
End Sub
Sub activarframes(estat As Boolean)
  Grups.Enabled = estat
 
End Sub
Private Sub grupdpalets_Reposition()
grupdpalets.Caption = "Grups " + atrim(grupdpalets.Recordset.AbsolutePosition + 1) + " / " + atrim(grupdpalets.Recordset.RecordCount)
 If grupdpalets.Recordset.EditMode <> 3 Then activarframes False
 actualitzar_vinculats
End Sub
Sub actualitzar_vinculats()
 Dim rst As Recordset
 If grupdpalets.Recordset.EOF Then Exit Sub
 llistasubmaterials.Clear
 Set rst = dbtmp.OpenRecordset("select * from grupsmaterialscompatibles_linies where idgrupsdematerialscompatibles=" + atrim(cadbl(grupdpalets.Recordset!numerodegrup)))
 While Not rst.EOF
   llistasubmaterials.AddItem rst!nomsubfamilia
   rst.MoveNext
 Wend
 Set rst = Nothing
End Sub

Private Sub lblLabels_Click(Index As Integer)

End Sub

Private Sub List1_Click()

End Sub

Private Sub llistasubmaterials_DblClick()
'   eliminar_subfamilia
End Sub
Sub eliminar_subfamilia()
  If llistasubmaterials.ListIndex = -1 Then MsgBox "Escull una subfamilia primer.", vbCritical, "Error": Exit Sub
If MsgBox("Segur que vols treure aquesta subfamilia de la selecció?" + Chr(10) + llistasubmaterials.Text, vbExclamation + vbDefaultButton2 + vbYesNo, "Treure familia") = vbNo Then Exit Sub
   dbtmp.Execute "delete* from grupsmaterialscompatibles_linies where nomsubfamilia='" + atrim(llistasubmaterials.Text) + "' and idgrupsdematerialscompatibles=" + atrim(grupdpalets.Recordset!numerodegrup)
   actualitzar_vinculats
   generar_sql_subfamilies
End Sub

Private Sub modificar_Click()
  If grupdpalets.Recordset.EOF Then Exit Sub
    activarframes True
    grupdpalets.Recordset.Edit
    cnomdelgrup.SetFocus
End Sub

Private Sub seccio_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub seccio_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub sortir_Click()
  If grupdpalets.Recordset.EditMode > 0 Then grupdpalets.Recordset.Update
  Unload grupmaterialscompatibles
End Sub




Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
  etestat = IIf(grupdpalets.Recordset.EditMode > 0, "Editant...", "")
End Sub
