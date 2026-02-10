VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form comandespendents 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Comandes amb compres pendents "
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13740
   ControlBox      =   0   'False
   Icon            =   "comandespendents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   13740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Reservar"
      Height          =   555
      Left            =   1290
      Picture         =   "comandespendents.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   32
      Tag             =   "comprar"
      ToolTipText     =   "Mirar les compres fetes pendents de rebre."
      Top             =   60
      Width           =   960
   End
   Begin VB.ListBox postit 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   2370
      Left            =   8070
      TabIndex        =   16
      Top             =   1275
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   810
      Left            =   3390
      TabIndex        =   6
      Top             =   -75
      Width           =   8520
      Begin VB.CommandButton buscarxrcomanda 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Buscar per comanda"
         Height          =   270
         Left            =   1725
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   135
         Width           =   2415
      End
      Begin VB.ComboBox cmaterial 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   540
         TabIndex        =   7
         Top             =   420
         Width           =   4065
      End
      Begin VB.TextBox cespesor 
         Height          =   285
         Left            =   4740
         TabIndex        =   11
         Top             =   420
         Width           =   510
      End
      Begin VB.TextBox ctubolam 
         Height          =   300
         Left            =   90
         TabIndex        =   8
         Top             =   420
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Caption         =   "Kg Comprats/ Kg Estoc  Kg Precomanda"
         Height          =   585
         Left            =   5340
         TabIndex        =   13
         Top             =   105
         Width           =   3135
         Begin VB.TextBox ckilosprecomanda 
            BackColor       =   &H00FFC0C0&
            Height          =   300
            Left            =   1965
            TabIndex        =   17
            Top             =   225
            Width           =   840
         End
         Begin VB.TextBox ckiloslliures 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   1050
            TabIndex        =   15
            Top             =   225
            Width           =   840
         End
         Begin VB.TextBox ckiloscomprats 
            BackColor       =   &H00C0C0FF&
            Height          =   300
            Left            =   150
            TabIndex        =   14
            Top             =   210
            Width           =   840
         End
      End
      Begin VB.Label mesuraespesor 
         BackStyle       =   0  'Transparent
         Caption         =   "M i c r e s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   4740
         TabIndex        =   31
         Top             =   285
         Width           =   585
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Espesor"
         Height          =   240
         Left            =   4710
         TabIndex        =   12
         Top             =   90
         Width           =   780
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
         Height          =   240
         Left            =   840
         TabIndex        =   10
         Top             =   195
         Width           =   2865
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "T/L"
         Height          =   255
         Left            =   135
         TabIndex        =   9
         Top             =   195
         Width           =   420
      End
   End
   Begin VB.Frame fxrbuscar 
      BackColor       =   &H00CCCCFF&
      Caption         =   "Buscar per:"
      Height          =   615
      Left            =   225
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   13410
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Buscar per comanda"
         Height          =   270
         Left            =   3315
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   225
         Width           =   2415
      End
      Begin VB.CheckBox bxrentregat 
         BackColor       =   &H00CCCCFF&
         Caption         =   "Entregat"
         Height          =   210
         Left            =   11655
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox bxrpendent 
         BackColor       =   &H00CCCCFF&
         Caption         =   "Pendent"
         Height          =   210
         Left            =   11655
         TabIndex        =   28
         Top             =   135
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CommandButton consultar 
         Height          =   375
         Left            =   12750
         Picture         =   "comandespendents.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   """ materialrebut=false """
         ToolTipText     =   "Buscar Registres"
         Top             =   165
         Width           =   585
      End
      Begin VB.TextBox bxrgrmm2 
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   6480
         TabIndex        =   26
         Tag             =   "Grm/m2"
         Text            =   "Grm/m2"
         Top             =   195
         Width           =   690
      End
      Begin VB.TextBox bxrfamilies 
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   8010
         TabIndex        =   25
         Tag             =   "Families"
         Text            =   "Families"
         Top             =   195
         Width           =   3540
      End
      Begin VB.TextBox bxrmicres 
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   5895
         TabIndex        =   24
         Tag             =   "Micres"
         Text            =   "Micres"
         Top             =   195
         Width           =   600
      End
      Begin VB.TextBox bxrample 
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   2610
         TabIndex        =   23
         Tag             =   "Ample"
         Text            =   "Ample"
         Top             =   195
         Width           =   585
      End
      Begin VB.TextBox bxrdataentrega 
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   1575
         TabIndex        =   22
         Tag             =   "Data Entrega"
         Text            =   "Data Entrega"
         Top             =   195
         Width           =   1050
      End
      Begin VB.TextBox bxrcomanda 
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   420
         TabIndex        =   21
         Tag             =   "Comanda"
         Text            =   "Comanda"
         Top             =   195
         Width           =   1185
      End
   End
   Begin VB.Data mirarcompres 
      Caption         =   "mirarcompres"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   15
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   810
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid reixamirar 
      Bindings        =   "comandespendents.frx":109E
      Height          =   5970
      Left            =   390
      OleObjectBlob   =   "comandespendents.frx":10B5
      TabIndex        =   18
      Top             =   1740
      Visible         =   0   'False
      Width           =   13350
   End
   Begin VB.ComboBox comboopcions 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6090
      TabIndex        =   5
      Top             =   1365
      Visible         =   0   'False
      Width           =   7425
   End
   Begin VB.CommandButton compraromirar 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Comprades"
      Height          =   555
      Left            =   2370
      Picture         =   "comandespendents.frx":3047
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "comprar"
      ToolTipText     =   "Mirar les compres fetes pendents de rebre."
      Top             =   60
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Comprar"
      Height          =   540
      Left            =   225
      Picture         =   "comandespendents.frx":35D1
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Comprar per les comandes sel.leccionades"
      Top             =   75
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   11955
      Picture         =   "comandespendents.frx":3B5B
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Refrescar"
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton sortir 
      Height          =   495
      Left            =   12975
      Picture         =   "comandespendents.frx":40E5
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Sortir"
      Top             =   60
      Width           =   585
   End
   Begin VB.Data comprescomandes 
      Caption         =   "comprescomandes"
      Connect         =   "Access"
      DatabaseName    =   "C:\temp\comprestmp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5595
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "comprescomandes"
      Top             =   105
      Visible         =   0   'False
      Width           =   1950
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "comandespendents.frx":466F
      Height          =   6510
      Left            =   240
      OleObjectBlob   =   "comandespendents.frx":4689
      TabIndex        =   0
      Top             =   735
      Width           =   13365
   End
End
Attribute VB_Name = "comandespendents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codimaterialacomprar As Double


Sub possarpostitcompres(CoLL As String)
   Dim comprats As Double
   Dim were As String
   Dim rstf As Recordset
   Dim lliures As Double
   Dim rstk As Recordset
   Dim rstll As Recordset
   If cmaterial.ListIndex = -1 Then Exit Sub
   Set rstf = dbconsulta.OpenRecordset("select * from familiescomprescomandes where id=" + atrim(cmaterial.ItemData(cmaterial.ListIndex)))
   If rstf.EOF Then ckiloscomprats = "0": ckiloslliures = "0": Exit Sub
   were = "codimaterial=" + atrim(rstf!material) + " and semielaborat='" + atrim(rstf!migelaborat) + "' and micres=" + passaradecimalpunt(atrim(rstf!micres)) + " and capcalera.materialrebut=False "
   Set rstk = dbtmpb.OpenRecordset("SELECT capcalera.numcomanda, liniescompra.codimaterial, liniescompra.semielaborat, liniescompra.micres, liniescompra.grmm2, liniescompra.quantitatkg,liniescompra.idliniacompra FROM capcalera INNER JOIN liniescompra ON capcalera.id = liniescompra.idcompra where " + were + ";")
   postit.Clear
   While Not rstk.EOF
      If CoLL = "C" Then
         postit.AddItem atrim(rstk!numcomanda) + Chr(9) + atrim(rstk!quantitatkg) + " Kg"
        Else
          Set rstll = dbtmpb.OpenRecordset("select comandavisual,kgcompra from comandesxlinia where comandavisual='ESTOC' AND idliniacompra=" + atrim(cadbl(rstk!idliniacompra)))
          If Not rstll.EOF Then postit.AddItem atrim(rstll!numcomanda) + Chr(9) + atrim(rstk!kgcompra) + " Kg"
      End If
      rstk.MoveNext
   Wend
   Set rstf = Nothing
   Set rstk = Nothing
End Sub

Private Sub buscarxrcomanda_Click()
  Dim comanda As Double
  Dim rstf As Recordset
  Dim i As Integer
  Dim pendentoentregat As String
  Dim llistaids As String
  Dim rstm As Recordset
  comanda = cadbl(InputBox("Entra la comanda que vols buscar.", "Buscar per comanda"))
  If comanda > 0 Then
      Set rstf = comprescomandes.Database.OpenRecordset("select * from comprescomandes where comanda=" + atrim(comanda))
      If Not rstf.EOF Then
       'criteri = " material = " + atrim(cadbl(rstf!material))
        Set rstm = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstf!material)))
       criteri = " descripcio='" + descripciomaterial(rstm) + "' "
       criteri = criteri + " and migelaborat='" + atrim(atrim(rstf!semielaborat)) + "'"
       criteri = criteri + " and micres=" + passaradecimalpunt(cadbl(rstf!espesor))
       Set rstf = comprescomandes.Database.OpenRecordset("select * from familiescomprescomandes where " + criteri)
       llistaids = ""
       While Not rstf.EOF
          llistaids = llistaids + "#" + atrim(rstf!id) + "$"
          rstf.MoveNext
       Wend
       If llistaids <> "" Then
          For i = 0 To cmaterial.ListCount - 1
            If InStr(1, llistaids, "#" + atrim(cmaterial.ItemData(i)) + "$") > 0 Then
               cmaterial.ListIndex = i
               cmaterial_Click
               comprescomandes.Recordset.FindFirst "comanda=" + atrim(comanda)
            End If
          Next i
       End If
       Else:
         Set rstf = comandescompra.capcalera.Database.OpenRecordset("SELECT capcalera.numcomanda,capcalera.nomprov,liniescompra.kgentregats FROM capcalera RIGHT JOIN (liniescompra RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra) ON capcalera.id = liniescompra.idcompra WHERE (((comandesxlinia.numcomanda)=" + atrim(comanda) + "));")
         If Not rstf.EOF Then
            pendentoentregat = IIf(cadbl(rstf!kgentregats) > 0, "MATERIAL ENTREGAT !!!!!!!!!!!!!!!!!!!!", "MATERIAL PENDENT D'ENTREGAR")
            MsgBox "Aquesta comanda ja està comprada a la comanda " + atrim(rstf!numcomanda) + Chr(10) + pendentoentregat + Chr(10) + atrim(rstf!nomprov)
           Else
            MsgBox "No he trobat aquesta comanda a la llista de pendents de comprar"
            comprescomandes.RecordSource = "select * from comprescomandes where COMANDA=1"
            comprescomandes.Refresh
         End If
         Exit Sub
      End If
  End If
End Sub

Private Sub bxfamilies_Change()
 
End Sub

Private Sub bxfamilies_GotFocus()

End Sub

Private Sub bxmicres_Change()
 
End Sub

Private Sub bxmicres_GotFocus()

End Sub

Private Sub bxrample_GotFocus()
  bxrcontrolagafafocus
End Sub

Private Sub bxrample_LostFocus()
 If bxrample.Text = "" Then
       bxrample.Text = bxrample.tag
       bxrample.ForeColor = &H808080
   End If
End Sub

Private Sub bxrcomanda_GotFocus()
   
   bxrcontrolagafafocus
End Sub

Private Sub bxrcomanda_LostFocus()
   If bxrcomanda.Text = "" Then
       bxrcomanda.Text = bxrcomanda.tag
       bxrcomanda.ForeColor = &H808080
   End If
End Sub

Private Sub bxrdataentrega_GotFocus()
   
 bxrcontrolagafafocus
End Sub
Sub bxrcontrolagafafocus()
  Dim cntrl As Control
  Set cntrl = Screen.ActiveControl
  If cntrl.Text <> "" Then
     If cntrl.Text = cntrl.tag Then cntrl.Text = ""
     cntrl.ForeColor = QBColor(0)
     
   Else:
       
       cntrl.Text = cntrl.tag
       cntrl.ForeColor = &H808080
  End If
End Sub
Private Sub bxrdataentrega_LostFocus()
   If Not IsDate(bxrdataentrega) Then
       bxrdataentrega.Text = bxrdataentrega.tag
       bxrdataentrega.ForeColor = &H808080
   End If
End Sub

Private Sub bxrfamilies_GotFocus()
bxrcontrolagafafocus
End Sub

Private Sub bxrfamilies_LostFocus()
If bxrfamilies.Text = "" Then
       bxrfamilies.Text = bxrfamilies.tag
       bxrfamilies.ForeColor = &H808080
   End If
End Sub

Private Sub bxrgrmm2_GotFocus()
bxrcontrolagafafocus
End Sub

Private Sub bxrgrmm2_LostFocus()
If bxrgrmm2.Text = "" Then
       bxrgrmm2.Text = bxrgrmm2.tag
       bxrgrmm2.ForeColor = &H808080
   End If
End Sub

Private Sub bxrmicres_GotFocus()
bxrcontrolagafafocus
End Sub

Private Sub bxrmicres_LostFocus()
 If bxrmicres.Text = "" Then
       bxrmicres.Text = bxrmicres.tag
       bxrmicres.ForeColor = &H808080
   End If
End Sub

Private Sub ckiloscomprats_DblClick()
  Dim x As Double
  Dim Y As Double
  If cmaterial.ListIndex = -1 Then Exit Sub
  If postit.visible Then postit.visible = False: Exit Sub
  x = ckiloscomprats.Left + Frame2.Left + Frame1.Left
  Y = ckiloscomprats.Top + Frame2.Top + Frame1.Top
  postitkgcompratsolliures "C"
'  possarpostitcompres "C"
  
  postit.visible = True
  postit.Left = x
  postit.Top = Y + ckiloscomprats.Height
End Sub
Sub postitkgcompratsolliures(CoLL As String)
   Dim comprats As Double
   Dim were As String
   Dim rstf As Recordset
   Dim lliures As Double
   Dim rstk As Recordset
   Dim rstll As Recordset
   Dim rstc As Recordset
   Dim data As String
   Dim eop As String
   postit.Clear
   Set rstf = dbconsulta.OpenRecordset("select * from familiescomprescomandes where id=" + atrim(cmaterial.ItemData(cmaterial.ListIndex)))
   If rstf.EOF Then ckiloscomprats = "0": ckiloslliures = "0": Exit Sub
   were = "codimaterial=" + atrim(rstf!material) + " and semielaborat='" + atrim(rstf!migelaborat) + "' and micres=" + passaradecimalpunt(atrim(rstf!micres)) + " and capcalera.materialrebut=False "
   Set rstk = dbtmpb.OpenRecordset("SELECT capcalera.numcomanda,capcalera.dataentrega, liniescompra.codimaterial, liniescompra.semielaborat, liniescompra.micres, liniescompra.grmm2, liniescompra.quantitatkg,liniescompra.idliniacompra FROM capcalera INNER JOIN liniescompra ON capcalera.id = liniescompra.idcompra where " + were + ";")
   While Not rstk.EOF
      comprats = comprats + cadbl(rstk!quantitatkg)
      If CoLL = "LL" Or CoLL = "P" Then
        If CoLL = "LL" Then
           eop = "ESTOC"
          Else: eop = "PRECOMANDA"
        End If
        Set rstll = dbtmpb.OpenRecordset("select comandavisual,kgcompra from comandesxlinia where comandavisual='" + eop + "' AND idliniacompra=" + atrim(cadbl(rstk!idliniacompra)))
        If Not rstll.EOF Then
         data = Format(rstk!dataentrega, "dd/mm/yy")
         If atrim(data) = "" Then data = "00/00/00"
         postit.AddItem data + " -> " + atrim(rstll!kgcompra) + " Kg"
         postit.ItemData(postit.NewIndex) = cadbl(rstk!idliniacompra)
        End If
         Else
           data = Format(rstk!dataentrega, "dd/mm/yy")
           If atrim(data) = "" Then data = "00/00/00"
           postit.AddItem data + " -> " + atrim(rstk!quantitatkg) + " Kg"
           postit.ItemData(postit.NewIndex) = cadbl(rstk!idliniacompra)
      End If
      rstk.MoveNext
   Wend
   Set rstf = Nothing
   Set rstk = Nothing

End Sub
Private Sub ckiloscomprats_LostFocus()
  If Screen.ActiveControl.Name <> "postit" Then postit.visible = False
End Sub

Private Sub ckiloslliures_DblClick()
 Dim x As Double
  Dim Y As Double
  If cmaterial.ListIndex = -1 Then Exit Sub
  If postit.visible Then postit.visible = False: Exit Sub
  x = ckiloslliures.Left + Frame2.Left + Frame1.Left
  Y = ckiloslliures.Top + Frame2.Top + Frame1.Top
  postitkgcompratsolliures "LL"
'  possarpostitcompres "LL"
  postit.visible = True
  postit.Left = x
  postit.Top = Y + ckiloslliures.Height
End Sub

Private Sub ckiloslliures_LostFocus()
  If Screen.ActiveControl.Name <> "postit" Then postit.visible = False
End Sub

Private Sub ckilosprecomanda_DblClick()
Dim x As Double
  Dim Y As Double
  If cmaterial.ListIndex = -1 Then Exit Sub
  If postit.visible Then postit.visible = False: Exit Sub
  x = ckilosprecomanda.Left + Frame2.Left + Frame1.Left
  Y = ckilosprecomanda.Top + Frame2.Top + Frame1.Top
  postitkgcompratsolliures "P"
'  possarpostitcompres "C"
  
  postit.visible = True
  postit.Left = x
  postit.Top = Y + ckilosprecomanda.Height
End Sub

Private Sub cmaterial_Change()
  cmaterial.width = 4100
End Sub

Private Sub cmaterial_Click()
  Dim rstf As Recordset
  Dim rstmat As Recordset
  Dim criteri As String
  Dim subconsulta As String
  cmaterial.width = 4100
  If cmaterial.ListCount = 0 Then Exit Sub
  If cmaterial.ListIndex <> -1 Then
      ckiloslliures.tag = ""
      Set rstf = dbconsulta.OpenRecordset("select * from familiescomprescomandes where id=" + atrim(cmaterial.ItemData(cmaterial.ListIndex)))
      If rstf.EOF Then Exit Sub
      Set rstmat = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(rstf!material))
      If rstmat.EOF Then Exit Sub
      'cmaterial.Text = rstf!descripcio
      ctubolam.Text = rstf!migelaborat
      cespesor.Text = IIf(rstf!micres < 0, rstf!micres * -1, rstf!micres)
      If Not rstf.EOF Then
      With rstmat
       subconsulta = " familiamat = " + atrim(cadbl(!familia)) + " And subfamiliamat = " + atrim(cadbl(!subfamilia))
       subconsulta = subconsulta + ajsinum(" and familiacol=", !familiacol) + ajsinum(" and subfamiliacol=", !subfamiliacol)
       subconsulta = subconsulta + ajsinum(" and familiaad=", !familiaad) + ajsinum(" and subfamiliaad=", !subfamiliaad)
      End With
      
       'criteri = " material = " + atrim(cadbl(rstf!material))
       criteri = subconsulta
       criteri = criteri + " and semielaborat='" + atrim(atrim(rstf!migelaborat)) + "'"
       'criteri = criteri + " and espesor=" + atrim(IIf(cadbl(rstf!micres) > 0, passaradecimalpunt(rstf!micres), passaradecimalpunt(rstf!micres) * -1))
       criteri = criteri + " and espesor=" + atrim(passaradecimalpunt(rstf!micres))
       mesuraespesor = IIf(cadbl(rstf!micres) > 0, "M i c r e s", "G r m/m 2")
       comprescomandes.RecordSource = "select * from comprescomandes where " + criteri + "  order by cont"
'       MsgBox comprescomandes.RecordSource
       comprescomandes.Refresh
'       possarkgdisponibles
      End If
  End If
  Set rstf = Nothing
  totalskgcompratsilliures
End Sub
Sub possarkgdisponibles()
   Dim mtrsdisponibles As Double
   Dim mtrsreservats As Double
   Dim pes1mtrs As Double
   If comprescomandes.Recordset.EOF Then Exit Sub
   comprescomandes.Recordset.MoveFirst
   While Not comprescomandes.Recordset.EOF
      metresdisponiblesireservats comprescomandes.Recordset, mtrsdisponibles, mtrsreservats, pes1mtrs
      comprescomandes.Recordset.Edit
      comprescomandes.Recordset!mtrsdisponibles = mtrsdisponibles
      comprescomandes.Recordset!kgdisponibles = Int(mtrsdisponibles * pes1mtrs)
      comprescomandes.Recordset.Update
      comprescomandes.Recordset.MoveNext
   Wend
End Sub
Function ajsinum(txt As String, v As Variant) As String 'ajuntar si es numeric
   If Not IsNumeric(v) Then Exit Function
   ajsinum = txt + atrim(cadbl(v))
End Function
Sub totalskgcompratsilliures()
   Dim comprats As Double
   Dim were As String
   Dim rstf As Recordset
   Dim lliures As Double
   Dim precomanda As Double
   Dim rstk As Recordset
   Dim rstll As Recordset
   Dim rstm As Recordset
   Dim whereespesor As String
   If cmaterial.ListIndex = -1 Then Exit Sub
   ckiloscomprats = "0": ckiloslliures = "0"
   Set rstf = dbconsulta.OpenRecordset("select * from familiescomprescomandes where id=" + atrim(cmaterial.ItemData(cmaterial.ListIndex)))
   If rstf.EOF Then Exit Sub
   Set rstm = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(rstf!material))
   If rstm.EOF Then Exit Sub
   If rstf!micres > 0 Then
      whereespesor = "micres = " + passaradecimalpunt(atrim(rstf!micres))
     Else: whereespesor = "grmm2=" + passaradecimalpunt(atrim(rstf!micres * -1))
   End If
   were = " (familia=" + atrim(cadbl(rstm!familia)) + " and subfamilia=" + atrim(cadbl(rstm!subfamilia)) + ajsinum(" and familiacol=", rstm!familiacol) + ajsinum(" and subfamiliacol=", rstm!subfamiliacol)
   were = were + ajsinum(" and familiaad=", rstm!familiaad) + ajsinum(" and subfamiliaad=", rstm!subfamiliaad) + ") "
   were = were + " and semielaborat='" + atrim(rstf!migelaborat) + "' and " + whereespesor  'capcalera.materialrebut=False "
   Set rstk = dbtmpb.OpenRecordset("SELECT capcalera.numcomanda, liniescompra.codimaterial, liniescompra.semielaborat, liniescompra.micres, liniescompra.grmm2, liniescompra.quantitatkg,liniescompra.idliniacompra FROM capcalera INNER JOIN liniescompra ON capcalera.id = liniescompra.idcompra where liniescompra.kgentregats=0 and " + were + ";")
   'MsgBox "SELECT capcalera.numcomanda, liniescompra.codimaterial, liniescompra.semielaborat, liniescompra.micres, liniescompra.grmm2, liniescompra.quantitatkg,liniescompra.idliniacompra FROM capcalera INNER JOIN liniescompra ON capcalera.id = liniescompra.idcompra where " + were + ";"
   'Shell "c:\windows\system32\cmd.exe /c echo " + "SELECT capcalera.numcomanda, liniescompra.codimaterial, liniescompra.semielaborat, liniescompra.micres, liniescompra.grmm2, liniescompra.quantitatkg,liniescompra.idliniacompra FROM capcalera INNER JOIN liniescompra ON capcalera.id = liniescompra.idcompra where " + were + ";" + " > c:\eco.txt"
   
   While Not rstk.EOF
      comprats = comprats + cadbl(rstk!quantitatkg)
      Set rstll = dbtmpb.OpenRecordset("select comandavisual,kgcompra from comandesxlinia where comandavisual='ESTOC' AND idliniacompra=" + atrim(cadbl(rstk!idliniacompra)))
      If Not rstll.EOF Then lliures = lliures + cadbl(rstll!kgcompra)
      Set rstll = dbtmpb.OpenRecordset("select comandavisual,kgcompra from comandesxlinia where comandavisual='PRECOMANDA' AND idliniacompra=" + atrim(cadbl(rstk!idliniacompra)))
      If Not rstll.EOF Then precomanda = precomanda + cadbl(rstll!kgcompra)
      rstk.MoveNext
   Wend
   ckiloscomprats = Format(comprats, "#,##0")
   ckiloslliures = Format(lliures, "#,##0")
   ckilosprecomanda = Format(precomanda, "#,##0")
   Set rstf = Nothing
   Set rstk = Nothing
   Set rstll = Nothing
End Sub
Private Sub cmaterial_DropDown()
  If cmaterial.ListCount = 0 Then possarfamiliesialtresicombomaterials
  cmaterial.width = 7000
End Sub

Private Sub cmaterial_LostFocus()
   cmaterial.width = 4100
End Sub

Private Sub comboopcions_Click()
  
  acceptaropcio
  comboopcions.visible = False
End Sub
Sub acceptaropcio()
   Dim numc As Double
   Dim numid As Long
   Dim rstl As Recordset
   Dim rstlc As Recordset
   If MsgBox("Segur que vols afegir la comanda " + atrim(comprescomandes.Recordset!comanda) + Chr(10) + Chr(13) + " a la linia " + comboopcions + " pendent d'assignar compra?", vbInformation + vbYesNo, "Atenció") = vbYes Then
     numid = comboopcions.ItemData(comboopcions.ListIndex)
     numc = cadbl(comprescomandes.Recordset!comanda)
     Set rstl = dbtmpb.OpenRecordset("select * from comandesxlinia  where id=" + atrim(numid))
     If Not rstl.EOF Then
      If comprescomandes.Recordset!kgpendents > rstl!kgcompra Then
         If MsgBox("Compte ja que n'hi ha menys dels que tens pendents." + Chr(10) + "Vols fer-ho igualment?", vbCritical + vbYesNo, "Atenció") = vbNo Then Exit Sub
      End If
      If rstl!kgcompra - comprescomandes.Recordset!kgpendents > 10 Then
         'clonar el registre
         Set rstlc = rstl.Clone
         rstlc.AddNew
           rstlc!idliniacompra = rstl!idliniacompra
           rstlc!numcomanda = rstl!numcomanda
           rstlc!comandavisual = rstl!comandavisual
           rstlc!kgcompra = Int(rstl!kgcompra) - Int(comprescomandes.Recordset!kgpendents)
          ' rstlc!kgpendents = Int(rstl!kgpendents) - Int(comprescomandes.Recordset!kgpendents)
         rstlc.Update
      End If
      rstl.Edit
      rstl!numcomanda = atrim(numc)
      rstl!comandavisual = atrim(numc)
      rstl!kgcompra = Int(comprescomandes.Recordset!kgpendents)
      'If rstl!kgpendents > 0 Then rstl!kgpendents = rstl!kgcompra
      rstl.Update
      
      comprescomandes.Recordset.Edit
      comprescomandes.Recordset!perlinkar = "Linkat"
      comprescomandes.Recordset.Update
     End If
   End If
   Set rstlc = Nothing
   Set rstl = Nothing
End Sub
Private Sub comboopcions_LostFocus()
   comboopcions.visible = False
End Sub

Private Sub Command1_Click()
  Dim rstpendents As Recordset
  Dim cont As Byte
  Dim ord As String
  Dim vsql As String
  
  ratoli "espera"
  If compraromirar.tag = "comprar" Then
   dbconsulta.Execute "delete * from comprescomandes"
   dbconsulta.Execute "INSERT INTO comprescomandes ( comanda,client,mtrscomanda,pesx1000,material,ample,plegat,solapa,obert,tractat,microperforat,semielaborat,espesor,mesuraesp,texteimp ) SELECT DISTINCT comandes.comanda, comandes.client,comandes.cantitatex, comandes.pes1000mtrs,materialex,ampleesq,plegatesq,solapa,oberturaex,tractatex,micropex,tubolam,espessor,mesuraesp,texteimpressio FROM comandes INNER JOIN comandesxlinia ON comandes.comanda <> comandesxlinia.numcomanda WHERE comandes.comanda>151000 and ( comandes.proximaseccio='E' and comandes.cantitatex>0  and isdate(dataactivacio));"
   
   'vsql = "SELECT DISTINCT comandes.comanda, comandes.client, comandes.cantitatex, comandes.pes1000mtrs, comandes.materialex, comandes.ampleesq, comandes.plegatesq, comandes.solapa, comandes.oberturaex, comandes.tractatex, comandes.micropex, comandes.tubolam, comandes.espessor, comandes.mesuraesp, comandes.texteimpressio FROM comandes LEFT JOIN comandesxlinia ON comandes.comanda = comandesxlinia.numcomanda WHERE (((comandes.cantitatex)>0) AND ((comandesxlinia.idliniacompra) Is Null) AND ((comandes.proximaseccio)='E') AND ((IsDate([dataactivacio]))<>False))"
   'dbconsulta.Execute "INSERT INTO comprescomandes ( comanda,client,mtrscomanda,pesx1000,material,ample,plegat,solapa,obert,tractat,microperforat,semielaborat,espesor,mesuraesp,texteimp ) " + vsql
   
   Set rstpendents = dbconsulta.OpenRecordset("comprescomandes")
   DoEvents
   While Not rstpendents.EOF
'     cont = cont + 1
     actualitzaregistre rstpendents
     rstpendents.MoveNext
 '    If cont = 100 Then cont = 0: DoEvents
   Wend
   dbconsulta.Execute "update comprescomandes set microperforat='Si' where microperforat='1'"
   dbconsulta.Execute "update comprescomandes set microperforat='No' where microperforat='0'"
   dbconsulta.Execute "delete * from comprescomandes where kgpendents<1"
   ordenarregistres
   possarfamiliesialtresicombomaterials
     Else
       mirarcompres.RecordSource = "select * from comprescomandes where comanda=1"
       mirarcompres.Refresh
       reixamirar.Refresh
       borrataulamirarcompres
       comandescompra.capcalera.Database.Execute "SELECT capcalera.numcomanda, capcalera.materialrebut,capcalera.dataentrega,liniescompra.*, '' AS nomfamilies,'' as microp into mirarcompres IN '" + fitxertemp + "' FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra " + IIf(consultar.tag <> "", " where ", "") + consultar.tag
       possarnomsfamilies
       dbconsulta.Execute "update mirarcompres set microp='Si' where microperforat=1"
       dbconsulta.Execute "update mirarcompres set microp='No' where microperforat=0"
       ord = reixamirar.tag
       If ord = "" Then ord = "nomfamilies,micres"
       consultar.tag = consultar.tag + " " + fxrbuscar.tag
       mirarcompres.RecordSource = "select * from mirarcompres " + IIf(consultar.tag <> "", " where ", "") + consultar.tag + " order by " + ord
       mirarcompres.Refresh
       If mirarcompres.Recordset.EOF Then MsgBox "No hi ha cap resultat."
  End If
  ratoli "normal"
  comprescomandes.RecordSource = "select * from comprescomandes order by cont"
  comprescomandes.Refresh
  
End Sub
Sub possarnomsfamilies()
  Dim rstn As Recordset
  Dim rstm As Recordset
  Set rstn = dbconsulta.OpenRecordset("mirarcompres")
  While Not rstn.EOF
     Set rstm = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstn!codimaterial)))
     If Not rstm.EOF Then
'      dbconsulta.Execute "update mirarcompres set nomfamilies='" + descripciomaterial(rstm) + "' where numcomanda=" + atrim(cadbl(rstn!numcomanda))
     rstn.Edit
       rstn!nomfamilies = descripciomaterial(rstm)
      rstn.Update
     End If
     rstn.MoveNext
  Wend
End Sub
Sub borrataulamirarcompres()
   On Error Resume Next
   dbconsulta.Execute "drop table mirarcompres"
End Sub
Sub crearnouregistreblanc(rsto As Recordset, cont As Long)
   Dim bk As Long
   
    bk = rsto!id
   rsto.AddNew
   rsto!cont = cont
   
   rsto.Update
   rsto.FindFirst "id=" + atrim(bk)
End Sub
Sub ordenarregistres()
  Dim rsto As Recordset
  Dim cont As Long
  Dim Control As String
  cont = 1
  Set rsto = dbconsulta.OpenRecordset("select *,trim(semielaborat)+trim(familiamat)+trim(subfamiliamat)+trim(familiacol)+trim(subfamiliacol)+trim(familiaad)+trim(subfamiliaad)+semielaborat+trim(espesor) as control from comprescomandes order by trim(semielaborat)+trim(familiamat)+trim(subfamiliamat)+trim(familiacol)+trim(subfamiliacol)+trim(familiaad)+trim(subfamiliaad)+semielaborat+trim(espesor), ample")
  While Not rsto.EOF
   If cadbl(rsto!cont) > 0 Then GoTo proxim
   If Trim(Control) <> Trim(rsto!Control) And cont > 1 Then
      Control = Trim(rsto!Control)
      crearnouregistreblanc rsto, cont
      cont = cont + 1
'       Else
'         emplenarcomboopcions rsto
'         If comboopcions.ListCount > 0 Then
'           If rsto.EditMode = 0 Then rsto.Edit
'           rsto!perlinkar = "???????????"
'         End If
   End If
   If rsto.EditMode = 0 Then rsto.Edit
   
   rsto!cont = cont
   rsto.Update
   cont = cont + 1
proxim:
   rsto.MoveNext
  Wend
  
End Sub
Sub metresdisponiblesireservats(rstp As Recordset, mtrsd As Double, mtrsr As Double, pes1mtrs As Double)
   Dim rstr As Recordset
   Dim rstm As Recordset
   Dim were As String
   Dim microp As String
   
   microp = IIf(atrim(rstp!microperforat) = "S", "", "Not")
   were = " (Bobines.Disponible > 0 And Palets.Disponible  And Palets.Ample = " + passaradecimalpunt(cadbl(rstp!ample)) + " And Palets.solapa = " + passaradecimalpunt(cadbl(rstp!solapa)) + " And Palets.Plegat = " + atrim(cadbl(rstp!plegat)) + " And Palets.carestractat = '" + atrim(rstp!tractat) + "' And Palets.obert = '" + atrim(rstp!obert) + "' And " + microp + " Palets.microperforat And Palets.semielaborat = '" + atrim((rstp!semielaborat)) + "' And Palets.micres = " + passaradecimalpunt(cadbl(rstp!espesor)) + ")"
   Set rstr = dbstocks.OpenRecordset("SELECT Palets.codimatprognou, Bobines.disponible as mtrsdisponibles, Bobines.Mts, Bobines.kilos, Palets.Ample, Palets.Plegat, Palets.Solapa, Palets.carestractat, Palets.obert, Palets.microperforat, Palets.semielaborat, Palets.micres, Palets.Disponible FROM Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet WHERE " + were + ";")
   
   If Not rstr.EOF Then
      If cadbl(rstr!mts) > 0 Then pes1mtrs = cadbl(rstr!kilos) / cadbl(rstr!mts)
   End If
   While Not rstr.EOF
     Set rstm = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(rstr!codimatprognou))
     If sonlesmateixesfamilies(rstp, rstm) Then mtrsd = mtrsd + cadbl(rstr!mtrsdisponibles)
     rstr.MoveNext
   Wend
   
End Sub
Function sonlesmateixesfamilies(rstp As Recordset, rstm As Recordset) As Boolean
   If cadbl(rstp!familiamat) = cadbl(rstm!familia) Then
             If cadbl(rstp!subfamiliamat) = cadbl(rstm!subfamilia) Then
               If cadbl(rstp!familiacol) = cadbl(rstm!familiacol) Then
                 If cadbl(rstp!subfamiliacol) = cadbl(rstm!subfamiliacol) Then
                   If cadbl(rstp!familiaad) = cadbl(rstm!familiaad) Then
                     If cadbl(rstp!subfamiliaad) = cadbl(rstm!subfamiliaad) Then
                          sonlesmateixesfamilies = True
                     End If
                   End If
                 End If
               End If
             End If
          End If
End Function
Sub actualitzaregistre(rstp As Recordset)
   Dim rstt As Recordset
   Dim rstm As Recordset
   Dim rstmat As Recordset
   Dim rstcli As Recordset
   Dim mtrsdisponibles As Double
   Dim mtrsreservats As Double
   Dim rstextres As Recordset
   
   rstp.Edit
   rstp!espesor = micresmaterial(CByte(cadbl(rstp!mesuraesp)), rstp!espesor, rstp!semielaborat)
   Set rstmat = dbtmp.OpenRecordset("select descripcio from materials where codi=" + atrim(rstp!material))
   Set rstextres = dbtmp.OpenRecordset("select codigrupmaterialcompatible from comandes_extres where comanda=" + atrim(cadbl(rstp!comanda)))
   Set rstm = dbtmp.OpenRecordset("SELECT materials.codi, familiesmaterials.codi, subfamiliesmaterials.codi, familiescolorants.codi, subfamiliescolorants.codi, familiesaditius.codi, subfamiliesaditius.codi FROM (subfamiliescolorants RIGHT JOIN (subfamiliesaditius RIGHT JOIN (familiesmaterials RIGHT JOIN (familiescolorants RIGHT JOIN (familiesaditius RIGHT JOIN materials ON familiesaditius.codi = materials.familiaad) ON familiescolorants.codi = materials.familiacol) ON familiesmaterials.codi = materials.familia) ON subfamiliesaditius.codi = materials.subfamiliaad) ON subfamiliescolorants.codi = materials.subfamiliacol) LEFT JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi WHERE (((materials.codi)=" + atrim(cadbl(rstp!material)) + "));")
   Set rstt = dbtmpb.OpenRecordset("select sum(kgcompra) as totalkg from comandesxlinia where numcomanda=" + atrim(cadbl(rstp!comanda)))
   Set rstcli = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(rstp!client))
   If Not rstextres.EOF Then rstp!compatible = nomgrupcompatible(cadbl(rstextres!codigrupmaterialcompatible))
   If Not rstt.EOF Then
     rstp!kgcomprats = cadbl(rstt!totalkg)
     If cadbl(rstp!pesx1000) > 0 Then rstp!mtrscomprats = (cadbl(rstp!kgcomprats) * 1000) / cadbl(rstp!pesx1000)
   End If
   Set rstt = dbstocks.OpenRecordset("select sum(metres) as totalmtrs from percomandaoclient where numcomanda=" + atrim(cadbl(rstp!comanda)))
   If Not rstt.EOF Then
      If rstt!totalmtrs > 0 Then rstp!mtrsassignats = cadbl(rstt!totalmtrs): rstp!aor = "R"
   End If
   Set rstt = dbstocks.OpenRecordset("select sum(metres) as totalmtrs from parcials where comanda='" + atrim(cadbl(rstp!comanda)) + "'")
   If Not rstt.EOF Then
      If rstt!totalmtrs > 0 Then rstp!mtrsassignats = cadbl(rstt!totalmtrs): rstp!aor = "A"
   End If
   rstp!mtrspendents = Redondejar(cadbl(rstp!mtrscomanda) - cadbl(rstp!mtrscomprats) - cadbl(rstp!mtrsassignats), 0)
   rstp!kgpendents = Redondejar(cadbl((rstp!mtrspendents / 1000) * cadbl(rstp!pesx1000), 0), 0)
   rstp!familiamat = cadbl(rstm![familiesmaterials.codi])
   rstp!subfamiliamat = cadbl(rstm![subfamiliesmaterials.codi])
   rstp!familiacol = cadbl(rstm![familiescolorants.codi])
   rstp!subfamiliacol = cadbl(rstm![subfamiliescolorants.codi])
   rstp!familiaad = cadbl(rstm![familiesaditius.codi])
   rstp!subfamiliaad = cadbl(rstm![subfamiliesaditius.codi])
   If Not rstmat.EOF Then rstp!nommaterial = atrim(rstmat!descripcio)
   If Not rstcli.EOF Then rstp!client = atrim(rstcli!nom)
   rstp.Update
   
   Set rstt = Nothing
End Sub

Function hihaproveidors() As Boolean
  Dim subconsutal As String
  Dim rsts As Recordset
  subconsulta = crear_subconsulta_deproveidors
  Set rsts = dbtmp.OpenRecordset(subconsulta)
  If rsts.EOF Then
     hihaproveidors = False
    Else: hihaproveidors = True
  End If
End Function

Function capseleccionada() As Boolean
   Dim rst As Recordset
   Dim bk As String
   'If reixa.Columns("seleccionat") = "S" Then bk = reixa.Columns("comanda")
   bk = reixa.Columns("comanda")
   comprescomandes.UpdateRecord
   comprescomandes.Refresh
   
   Set rst = comprescomandes.Database.OpenRecordset("select comanda,seleccionat from comprescomandes where seleccionat='S'")
   If Not rst.EOF Then
       bk = rst!comanda
       capseleccionada = False
         Else:
         capseleccionada = True
   End If
   comprescomandes.Recordset.FindFirst "comanda=" + atrim(cadbl(bk))
   Set rst = Nothing
End Function

Private Sub Command2_Click()
  If reixa.visible = False Then compraromirar_Click: Exit Sub
  'If reixa.SelBookmarks.Count <= 0 Then MsgBox "No hi ha cap comanda sel.leccionada": Exit Sub
  If capseleccionada Then MsgBox "No hi ha cap comanda seleccionada per comprar": Exit Sub
  If comandescompra.comprovarsilacomandajashacomprat(comprescomandes.Recordset!comanda) Then MsgBox "Aquesta comanda ja està comprada.": Exit Sub
  If Not hihaproveidors Then
     MsgBox "No hi ha proveidors pels materials de les comandes sel.leccionades", vbCritical, "Atenció"
     GoTo fi
  End If
  vcomprantmaterialcompatible = False
  If atrim(comprescomandes.Recordset!compatible) <> "" Then vcomprantmaterialcompatible = True
  comandescompra.novacomanda
  triar_proveidor_seleccio
  comandescompra.refrescar_proveidor
  If comandescompra.proveidor = "" Then
     comandescompra.capcalera.Recordset.CancelUpdate
     MsgBox "No hi ha proveidor sel.leccionat..." + Chr(10) + Chr(13) + " S'ha cancelat la comanda", vbCritical, "Cancelat"
     GoTo fi
  End If
  
   If comandescompra.capcalera.Recordset.EditMode > 0 Then comandescompra.capcalera.Recordset.Update
   comandescompra.capcalera.Recordset.Bookmark = comandescompra.capcalera.Recordset.LastModified

  demanar_valorsdelalinia
  'Or diamextlinia = 0 Or mandrillinia = 0
  If preulinia = 0 Or amplelinia = 0 Or Not IsDate(dataentrega) Or codimaterialacomprar = 0 Then
    If comandescompra.capcalera.Recordset.EditMode > 0 Then comandescompra.capcalera.Recordset.CancelUpdate
    MsgBox "Falten parametres de linia per fer la compra d'aquest material", vbCritical, "Atenció"
    GoTo fi
  End If
  comandescompra.capcalera.Recordset.Edit
  comandescompra.capcalera.Recordset!dataentrega = dataentrega
  comandescompra.capcalera.Recordset.Update
  crear_liniadecompra
  afegircomandesalalinia
'  wait 3
  comandescompra.borrarlesliniesdedescripcio False
  If comandescompra.liniescompra.Recordset.EditMode = 0 Then comandescompra.liniescompra.Recordset.Edit
  comandescompra.oklinia
  comandescompra.calcular_totals_comanda comandescompra.capcalera.Recordset!numcomanda
  comandescompra.capcalera.Recordset.Bookmark = comandescompra.capcalera.Recordset.LastModified
fi:
End Sub
Function demanarmaterialacomprar() As Double
  Dim rstmat As Recordset
  Dim familiesmat As String
  Dim subconsulta As String
  Dim vmaterialcompatible As String
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  vmaterialcompatible = subconsultamaterialcompatible
  With comprescomandes.Recordset
  If vmaterialcompatible = "" Then
         subconsulta = " familia = " + atrim(cadbl(!familiamat)) + " And subfamilia = " + atrim(cadbl(!subfamiliamat))
         subconsulta = subconsulta + " and familiacol=" + atrim(cadbl(!familiacol)) + " and subfamiliacol=" + atrim(cadbl(!subfamiliacol))
         subconsulta = subconsulta + " and familiaad=" + atrim(cadbl(!familiaad)) + " and subfamiliaad=" + atrim(cadbl(!subfamiliaad))
       Else: subconsulta = substituirtot(vmaterialcompatible, "where", "")
  End If
  
  If materialexacte Then subconsulta = " codi=" + atrim(cadbl(!material))
  End With
  formseleccio.Data1.DatabaseName = cami
  
  'Set formseleccio.Data1.Recordset = dbtmp.OpenRecordset("select * from materials where codi>499 and proveidor=" + atrim(cadbl(comandescompra.capcalera.Recordset!codiproveidor)) + " and " + subconsulta + " order by descripcio")
  
  formseleccio.Data1.RecordSource = "select * from materials where codi>499 and proveidor=" + atrim(cadbl(comandescompra.capcalera.Recordset!codiproveidor)) + " and " + subconsulta + " order by descripcio"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   demanarmaterialacomprar = atrim(cadbl(formseleccio.Data1.Recordset!codi))
     Else: demanarmaterialacomprar = 0
  End If
  Unload formseleccio
End Function
Sub demanar_valorsdelalinia()
  Dim data As String
   codimaterialacomprar = demanarmaterialacomprar
demanarample:
   amplelinia = cadbl(reixa.Columns("ample"))
   amplelinia = passaradecimal(InputBox("Entra l'amplada del material que vols comprar", "Dades compra", IIf(amplelinia > 0, amplelinia, "")))
   If amplelinia > 300 Then MsgBox "Aquest ample de material es massa gran, torna a entrar-lo", vbCritical + vbOKOnly, "Atenció": GoTo demanarample
   preulinia = passaradecimal(InputBox("Entra el PREU del material que vols comprar", "Dades compra"))
   diamextlinia = passaradecimal(InputBox("Entra el diametre exterior de la bobina que vols comprar" + Chr(10) + Chr(13) + "    ULL!!!! AMB CENTIMETRES. Ex: 80 ", "Dades compra", 80))
   mandrillinia = passaradecimal(InputBox("Entra el mandril de la bobina que vols comprar" + Chr(10) + Chr(13) + "    ULL!!!! AMB CENTIMETRES. Ex: 15.2 ", "Dades compra", 15.2))
   data = "01/01/01"
   While DateDiff("s", Now, CVDate(data)) < 1
     data = InputBox("Entra la data prevista d'entrega del material" + Chr(10) + "Ex: dd/mm/yy", "Dades compra")
     If Not IsDate(data) Then Exit Sub
     If DateDiff("s", Now, CVDate(data)) < 1 Then
       MsgBox "Aquesta data d'entrega no es valida", vbCritical, "Error"
         Else: dataentrega = data
     End If
   Wend
End Sub
Sub afegircomandesalalinia()
 Dim i As Integer
 Dim numc As Double
 comprescomandes.Recordset.MoveFirst
 While Not comprescomandes.Recordset.EOF
  If comprescomandes.Recordset!seleccionat = "S" Then
   numc = cadbl(comprescomandes.Recordset!comanda)
   kgcompralinia = cadbl(comprescomandes.Recordset!kgpendents)
   If comandescompra.comparasielmaterialcorrespon(atrim(numc), comandescompra.liniescompra.Recordset!idliniacompra) = 1 Then
      comandescompra.afegircomandaalinia atrim(numc)
       Else: MsgBox "La comanda " + atrim(numc) + " no coincideix el material amb la compra feta.", vbCritical + vbOKOnly, "Atenció"
   End If
   comprescomandes.Recordset.Edit
   comprescomandes.Recordset!seleccionat = "F"
   comprescomandes.Recordset.Update
  End If
  comprescomandes.Recordset.MoveNext
  
 Wend
 comandescompra.actualitzar_valors_comanda
 comandescompra.sumar_kilos
  
End Sub
Sub crear_liniadecompra()
  comandescompra.liniescompra.Recordset.AddNew
  comandescompra.fdescmat.Enabled = True
  comandescompra.liniescompra.Recordset!idcompra = comandescompra.capcalera.Recordset!id
  comandescompra.liniescompra.Recordset!tipusmaterialcomprat = "M"
  possarmaterial
  possarcaracteristiques
  formselecciotipuscompra.tag = "M"
  comandescompra.oklinia
End Sub
Sub possarcaracteristiques()
  With comprescomandes.Recordset
   comandescompra.itl = atrim(!semielaborat)
   comandescompra.icares = atrim(!tractat)
   comandescompra.iobert = atrim(!obert)
   comandescompra.iplegat = cadbl(!plegat)
   comandescompra.isolapa = cadbl(!solapa)
   comandescompra.imicrop = IIf(atrim(!microperforat) = "S", 1, 0)
   comandescompra.iample = amplelinia
   comandescompra.iespesor = IIf(cadbl(!espesor) < 0, cadbl(!espesor) * -1, cadbl(!espesor))
   comandescompra.diamext = diamextlinia
   comandescompra.mandril = mandrillinia
   comandescompra.preu = preulinia
 End With
End Sub
Sub possarmaterial()
  Dim rstm As Recordset
  Set rstm = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(codimaterialacomprar)))
  If Not rstm.EOF Then
   comandescompra.liniescompra.Recordset!codimaterial = atrim(cadbl(rstm!codi))
   comandescompra.liniescompra.Recordset!nommaterial = atrim(rstm!descripcio)
   comandescompra.liniescompra.Recordset!grmm2 = cadbl(rstm!grmm2)
   comandescompra.possarmicresogrmm2
   comandescompra.combomaterial.Text = comandescompra.liniescompra.Recordset!nommaterial
   comandescompra.possar_families comandescompra.liniescompra.Recordset!codimaterial, rstm
  End If
End Sub
Function crear_subconsulta_deproveidors() As String
  Dim subconsulta As String
  Dim rstc As Recordset
  With comprescomandes.Recordset
  If atrim(comprescomandes.Recordset!compatible) = "" Then
   subconsulta = "select proveidor from materials where familia=" + atrim(cadbl(!familiamat)) + " and subfamilia=" + atrim(cadbl(!subfamiliamat))
   subconsulta = subconsulta + " and familiacol=" + atrim(cadbl(!familiacol)) + " and subfamiliacol=" + atrim(cadbl(!subfamiliacol))
   subconsulta = subconsulta + " and familiaad=" + atrim(cadbl(!familiaad)) + " and subfamiliaad=" + atrim(cadbl(!subfamiliaad))
     Else
       subconsulta = "select proveidor from materials " + subconsultamaterialcompatible
'       MsgBox subconsulta
  End If
  If materialexacte Then subconsulta = "select proveidor from materials WHERE codi=" + atrim(cadbl(!material))
  End With
  crear_subconsulta_deproveidors = subconsulta
  Set rstc = Nothing
End Function
Function subconsultamaterialcompatible() As String
   Dim rstc As Recordset
   Dim vfamiliescompatibles As String
   Set rstc = dbtmp.OpenRecordset("select codigrupmaterialcompatible from comandes_extres where comanda=" + atrim(atrim(comprescomandes.Recordset!comanda)))
   If Not rstc.EOF Then
            vfamiliescompatibles = familiescompatibles(rstc!codigrupmaterialcompatible)
            If vfamiliescompatibles <> "" Then
               subconsultamaterialcompatible = " where " + vfamiliescompatibles
            End If
   End If
   Set rstc = Nothing
End Function
Sub triar_proveidor_seleccio()
  Dim subconsulta As String
  If comandescompra.capcalera.Recordset.EditMode = 0 Then Exit Sub
  Load formseleccio
  subconsulta = crear_subconsulta_deproveidors
 
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  'Set formseleccio.Data1.Recordset = dbtmp.OpenRecordset("select * from proveidors where codi in (" + subconsulta + ")")
  formseleccio.Data1.RecordSource = "select * from proveidors where codi in (" + subconsulta + ")"
  formseleccio.refrescar
  'MsgBox "select * from proveidors where codi in (" + subconsulta + ")"
  Clipboard.Clear
  Clipboard.SetText "select * from proveidors where codi in (" + subconsulta + ")"
  formseleccio.Show 1
  If seleccioret = 1 Then
   comandescompra.capcalera.Recordset!codiproveidor = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   comandescompra.capcalera.Recordset!nomprov = atrim(formseleccio.Data1.Recordset!nom)
   comandescompra.proveidor = atrim(formseleccio.Data1.Recordset!nom)
   comandescompra.capcalera.Recordset!codiproveidorcomercial = comandescompra.triar_proveidor_comercial(comandescompra.capcalera.Recordset!codiproveidor)
  End If
  Unload formseleccio
End Sub

Private Sub Command3_Click()
  'Dim numc As String
  'comprescomandes.Refresh
  'numc = InputBox("Entra el numero de comanda a buscar", "ATenció")
  'comprescomandes.Recordset.FindFirst "comanda=" + atrim(cadbl(numc))
    Dim rstc As Recordset
    Dim numc As Double
    numc = cadbl(InputBox("Entra el numero de comanda de que busques:", "Buscar per numero de comanda"))
    consultar.tag = ""
    If numc > 0 Then
        Set rstc = comandescompra.capcalera.Database.OpenRecordset("SELECT capcalera.numcomanda FROM capcalera RIGHT JOIN (liniescompra RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra) ON capcalera.id = liniescompra.idcompra WHERE (((comandesxlinia.numcomanda)=" + atrim(numc) + "));")
        If Not rstc.EOF Then consultar.tag = " numcomanda in ("
        r = ""
        If rstc.EOF Then MsgBox "No s'ha trobat aquesta comanda.": Exit Sub
        While Not rstc.EOF
         consultar.tag = consultar.tag + r + atrim(cadbl(rstc!numcomanda))
         r = ","
         rstc.MoveNext
        Wend
        If consultar.tag <> "" Then consultar.tag = consultar.tag + ")"
        Command1_Click
        
    End If

End Sub

Private Sub Command4_Click()
   Dim ample As Double
   Dim f As Double
   If reixa.Columns("seleccionat") = "S" Then bk = reixa.Columns("comanda")
   comprescomandes.UpdateRecord
   comprescomandes.Refresh
   Set rst = dbconsulta.OpenRecordset("select comanda,seleccionat,ample,mtrspendents from comprescomandes where seleccionat='S'")
   dbstocks.Execute "delete * from pendentsdereservar"
   If rst.EOF Then Exit Sub
   ample = rst!ample
   While Not rst.EOF
     If ample <> rst!ample Then MsgBox "Hi ha una amplada diferent en la sel.leccio de comandes.", vbCritical, "Atenció": Exit Sub
     rst.MoveNext
   Wend
   rst.MoveFirst
   While Not rst.EOF
      dbstocks.Execute "insert into pendentsdereservar (comanda,reservar,metres,entrat) values (" + atrim(rst!comanda) + ",false," + atrim(rst!mtrspendents) + ",false)"
      rst.MoveNext
   Wend
   Command4.tag = "reserva"
   
   f = Shell(rutadelfitxer(llegir_ini("General", "rutaprogbaixes", "comandes.ini")) + "palets.exe comandes.ini comprant", vbNormalFocus)
   wait 2
   MsgBox "Reservant..." + Chr(10) + " Fes click per continuar."
   canviarles_s_per_p_alreservar
End Sub
Sub canviarles_s_per_p_alreservar()
  Dim rstp As Recordset
  comprescomandes.Refresh
      While Not comprescomandes.Recordset.EOF
         If comprescomandes.Recordset!seleccionat = "S" Then
            Set rstp = dbstocks.OpenRecordset("select * from percomandaoclient where numcomanda=" + atrim(comprescomandes.Recordset!comanda))
            If Not rstp.EOF Then
              comprescomandes.Recordset.Edit
              comprescomandes.Recordset!seleccionat = "F"
              comprescomandes.Recordset.Update
            End If
         End If
         comprescomandes.Recordset.MoveNext
      Wend
   Set rstp = Nothing
   comprescomandes.Refresh
End Sub
Private Sub compraromirar_Click()
  If compraromirar.tag = "comprar" Then
      reixa.visible = False
      reixamirar.visible = True
      fxrbuscar.visible = reixamirar.visible
      compraromirar.tag = "mirar"
      Frame1.Enabled = False
'      Command1_Click
      consultar_Click
        Else:
           compraromirar.tag = "comprar"
           reixamirar.visible = False
           fxrbuscar.visible = reixamirar.visible
           Frame1.Enabled = True
           reixa.visible = True
  End If

End Sub
Function micresmaterial(codimesuralineal As Byte, espesor As Double, tubolam As String) As String
  Dim rstmesural As Recordset
  Dim descripcio As String
 ' Dim r As String
  Set rstmesural = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(codimesuralineal))
  If rstmesural.EOF Then
     micresmaterial = 0: Exit Function
  End If
  descripcio = rstmesural!descripcio
  r = espesor
  If descripcio = "GALGUES" Then
            If tubolam = "T" Then
                 r = cadbl(espesor / 4, 1)
                  Else: r = cadbl(espesor / 2, 1)
            End If
  End If
  'If InStr(1, descripcio, "GR/") > 0 Then
  '  micresmaterial = espesor * -1
  'End If
  descripcio = IIf(descripcio = "MICRES", "Mic", descripcio)
  descripcio = IIf(descripcio = "GALGUES", "Mic", descripcio)
  If InStr(1, descripcio, "GR/") > 0 Then
     descripcio = "GR/MT2"
     r = cadbl(r) * -1
  End If
     
  micresmaterial = r
  r = descripcio
End Function

Private Sub consultar_Click()
  Dim were As String
  fxrbuscar.tag = ""
  If bxrcomanda <> bxrcomanda.tag Then
     were = were + IIf(were = "", "", " and ")
     were = were + " numcomanda=" + atrim(cadbl(bxrcomanda))
  End If
  If bxrdataentrega <> bxrdataentrega.tag Then
     were = were + IIf(were = "", "", " and ")
     were = were + " dataentrega=#" + atrim(bxrdataentrega) + "#"
  End If
  If bxrample <> bxrample.tag Then
     were = were + IIf(were = "", "", " and ")
     were = were + " ample=" + atrim(cadbl(bxrample))
  End If
  If bxrmicres <> bxrmicres.tag Then
     were = were + IIf(were = "", "", " and ")
     were = were + " micres=" + atrim(cadbl(bxrmicres))
  End If
  If bxrgrmm2 <> bxrgrmm2.tag Then
     were = were + IIf(were = "", "", " and ")
     were = were + " grmm2=" + atrim(cadbl(bxrgrmm2))
  End If
  
  
  If bxrpendent.Value = 1 And bxrentregat.Value = 0 Then
     were = were + IIf(were = "", "", " and ")
    ' were = were + " materialrebut=false "
     were = were + " totentregat=false "
  End If
  
  
  If bxrpendent.Value = 0 And bxrentregat.Value = 1 Then
     were = were + IIf(were = "", "", " and ")
    ' were = were + " materialrebut=true "
     were = were + " totentregat=true "
  End If
    
  If bxrfamilies <> bxrfamilies.tag Then
     fxrbuscar.tag = IIf(were = "", "", " and ")
     fxrbuscar.tag = fxrbuscar.tag + " nomfamilies like '*" + treure_apostruf(bxrfamilies) + "*'"
     
  End If
  consultar.tag = were
  Command1_Click
End Sub

Private Sub Form_Load()
   
    comprescomandes.DatabaseName = fitxertemp
    Set dbconsulta = DBEngine.OpenDatabase(fitxertemp)
    mirarcompres.DatabaseName = fitxertemp
    comprescomandes.RecordSource = "select * from comprescomandes where id=-1 order by cont"
    comprescomandes.Refresh
    comprescomandes.Database.Execute "update comprescomandes set seleccionat='' "
     reixamirar.Top = 1290 'reixa.Top
     reixamirar.Left = reixa.Left
     reixamirar.width = reixa.width
     'reixamirar.Height = reixa.Height - (reixa.Top - 1170)
End Sub
Function familiescompatibles(vcodi As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from grupsmaterialscompatibles where numerodegrup=" + atrim(vcodi))
   If Not rst.EOF Then familiescompatibles = atrim(rst!sqlprincipal) + " " + atrim(rst!sqlsubfamilies) + ")"
   familiescompatibles = substituirtot(familiescompatibles, "materials.", "")
   Set rst = Nothing
'   MsgBox familiescompatibles
End Function
Sub emplenarcomboopcions(rstt As Recordset)
   Dim rstsel As Recordset
   Dim were As String
   Dim rstc As Recordset
   comboopcions.Clear
   Set rstc = dbtmp.OpenRecordset("SELECT comandes_extres.codigrupmaterialcompatible,comandes.comanda, comandes_extres.materialexacte, comandes.materialex FROM comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda where comandes.comanda = " + atrim(rstt!comanda))
   With rstt
   'were = " liniescompra.kgentregats=0 and semielaborat='" + atrim((!semielaborat)) + "' and familia=" + atrim(cadbl(!familiamat)) + " and subfamilia=" + atrim(cadbl(!subfamiliamat)) + ajsinum(" and familiacol=", !familiacol) + ajsinum(" and subfamiliacol=", !subfamiliacol) + ajsinum(" and familiaad=", !familiaad) + ajsinum(" and subfamiliaad=", !subfamiliaad) + " And (micres=" + passaradecimalpunt(cadbl(!espesor)) + " or grmm2=" + passaradecimalpunt(cadbl(!espesor) * -1) + ")"
   were = " (not liniescompra.totentregat) and semielaborat='" + atrim((!semielaborat)) + "' and familia=" + atrim(cadbl(!familiamat)) + " and subfamilia=" + atrim(cadbl(!subfamiliamat)) + ajsinum(" and familiacol=", !familiacol) + ajsinum(" and subfamiliacol=", !subfamiliacol) + ajsinum(" and familiaad=", !familiaad) + ajsinum(" and subfamiliaad=", !subfamiliaad) + " And (micres=" + passaradecimalpunt(cadbl(!espesor)) + " or grmm2=" + passaradecimalpunt(cadbl(!espesor) * -1) + ")"
   If Not rstc.EOF Then
        If cadbl(rstc!codigrupmaterialcompatible) > 0 Then
            vfamiliescompatibles = familiescompatibles(rstc!codigrupmaterialcompatible)
            were = " (not liniescompra.totentregat) and semielaborat='" + atrim((!semielaborat)) + "' and " + vfamiliescompatibles + " And (micres=" + passaradecimalpunt(cadbl(!espesor)) + " or grmm2=" + passaradecimalpunt(cadbl(!espesor) * -1) + ")"
        End If
   End If
   If Not rstc.EOF Then If rstc!materialexacte Then were = were + " and  (liniescompra.codimaterial=" + atrim(rstc!materialex) + ") "
   End With
'   Set rstsel = dbtmpb.OpenRecordset("SELECT comandesxlinia.comandavisual ,comandesxlinia.id as numeroid, comandesxlinia.kgcompra as kg,comandesxlinia.descripcio as descripcio,liniescompra.* FROM liniescompra RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra WHERE ((comandesxlinia.comandavisual='PRECOMANDA') and " + were + ");")
   Set rstsel = dbtmpb.OpenRecordset("SELECT capcalera.dataentrega,capcalera.numcomanda, capcalera.nomprov, comandesxlinia.comandavisual, comandesxlinia.Id AS numeroid, comandesxlinia.kgcompra AS kg, comandesxlinia.descripcio AS descripcio, liniescompra.* FROM capcalera RIGHT JOIN (liniescompra RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra) ON capcalera.id = liniescompra.idcompra WHERE  ((comandesxlinia.comandavisual='PRECOMANDA') and " + were + ") order by capcalera.dataentrega;")
   
   While Not rstsel.EOF
      comboopcions.AddItem "P " + justifica(rstsel!numcomanda, 6) + " " + justifica(Format(rstsel!dataentrega, "dd/mm/yy"), 8) + " " + justifica(Format(rstsel!kg, "#,##0"), 7) + " Kg " + justifica(atrim(rstsel!ample), 6) + " cm" + " -> " + justifica(Mid(atrim(rstsel!nomprov), 1, 10), 10) + "|" + atrim(rstsel!descripcio)
      comboopcions.ItemData(comboopcions.NewIndex) = cadbl(rstsel![numeroid])
      rstsel.MoveNext
   Wend
   
   Set rstsel = dbtmpb.OpenRecordset("SELECT capcalera.dataentrega,capcalera.numcomanda, capcalera.nomprov, comandesxlinia.comandavisual, comandesxlinia.Id AS numeroid, comandesxlinia.kgcompra AS kg, comandesxlinia.descripcio AS descripcio, liniescompra.* FROM capcalera RIGHT JOIN (liniescompra RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra) ON capcalera.id = liniescompra.idcompra WHERE ((comandesxlinia.comandavisual='ESTOC') and " + were + ") order by capcalera.dataentrega;")
   While Not rstsel.EOF
      comboopcions.AddItem "E " + justifica(rstsel!numcomanda, 6) + " " + justifica(Format(rstsel!dataentrega, "dd/mm/yy"), 8) + " " + justifica(Format(rstsel!kg, "#,##0"), 7) + " Kg " + justifica(atrim(rstsel!ample), 6) + " cm" + " -> " + justifica(Mid(atrim(rstsel!nomprov), 1, 10), 10) + "|" + atrim(rstsel!descripcio)
      comboopcions.ItemData(comboopcions.NewIndex) = cadbl(rstsel![numeroid])
      rstsel.MoveNext
   Wend
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Shift = 2 Then
    comprescomandes.RecordSource = "select * from comprescomandes order by cont"
    comprescomandes.Refresh
  End If
End Sub

Private Sub postit_DblClick()
  Dim rstc As Recordset
  Dim numll As Double
  Dim numc As Double
  If postit.ListIndex = -1 Then Exit Sub
  numll = postit.ItemData(postit.ListIndex)
  Set rstc = comandescompra.capcalera.Database.OpenRecordset("SELECT capcalera.numcomanda, liniescompra.idliniacompra FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE (((liniescompra.idliniacompra)=" + atrim(numll) + "));")
  If Not rstc.EOF Then
   numc = cadbl(rstc!numcomanda)
   comandescompra.capcalera.RecordSource = "capcalera"
   comandescompra.capcalera.Recordset.FindFirst "numcomanda=" + atrim(numc)
   DoEvents
   comandescompra.liniescompra.Recordset.FindFirst "idliniacompra=" + atrim(numll)
   DoEvents
   comandespendents.Hide
  End If
  Set rstc = Nothing
End Sub

Private Sub postit_GotFocus()
  postit.visible = True
End Sub

Private Sub postit_LostFocus()
postit.visible = False
End Sub
Sub saltaralacomanda()
    Dim comanda As String
    Dim rstc As Recordset
    comanda = atrim(cadbl(reixa.Columns("comanda")))
    Set rstc = comandescompra.capcalera.Database.OpenRecordset("SELECT comandesxlinia.numcomanda, capcalera.numcomanda as comanda FROM capcalera RIGHT JOIN (liniescompra RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra) ON capcalera.id = liniescompra.idcompra WHERE (((comandesxlinia.numcomanda)=" + comanda + "));")
    comanda = "0"
    If Not rstc.EOF Then comanda = rstc!comanda
    comandescompra.capcalera.RecordSource = "capcalera"
    comandescompra.capcalera.Refresh
    comandescompra.capcalera.Recordset.FindFirst "numcomanda=" + comanda
    Set rstc = Nothing
    sortir_Click
End Sub
Private Sub reixa_ButtonClick(ByVal ColIndex As Integer)
  Dim vdescfam As String
  Dim vcodifam As Double
  
' If comprescomandes.Recordset.EditMode > 0 Then
  If cmaterial = "" Then MsgBox "Primer escull el material corresponent": Exit Sub
  If ColIndex = 13 Then
     If comandescompra.comprovarsilacomandajashacomprat(comprescomandes.Recordset!comanda) Then
        If MsgBox("Aquesta comanda ja està comprada no pots linkar-la" + Chr(10) + Chr(13) + " Et carrego la comanda on es va comprar?", vbInformation + vbYesNo, "Atenció") = vbYes Then
         saltaralacomanda
        End If
        Exit Sub
     End If
     If comprescomandes.Recordset!perlinkar = "Linkat" Then
        If MsgBox("Aquesta comanda ja està linkada." + Chr(10) + Chr(13) + " Et carrego la comanda on es va comprar? ", vbInformation, "Atenció") = vbYes Then
          saltaralacomanda
        End If
        Exit Sub
     End If
     emplenarcomboopcions comprescomandes.Recordset
     comboopcions.visible = True
     comboopcions.Top = reixa.RowTop(reixa.Row) + reixa.Top
     comboopcions.Left = reixa.Columns(ColIndex).Left + reixa.Left
     'comboopcions.Width = reixa.Columns(ColIndex).Width
     comboopcions.SetFocus
     SendKeys ("%{DOWN}")
    'End If
  End If
  If ColIndex = 14 Then
      escullir_familiacompatible vcodifam, vdescfam
      If vcodifam > 0 Then
           dbtmp.Execute "update comandes_extres set codigrupmaterialcompatible=" + atrim(vcodifam) + " where comanda=" + atrim(cadbl(reixa.Columns("comanda")))
           comprescomandes.Recordset.Edit
           comprescomandes.Recordset!compatible = vdescfam
           comprescomandes.Recordset.Update
           comprescomandes.Recordset.Move 0
            Else
              If vcodifam < 0 Then
                If MsgBox("Vols treure el grup compatible d'aquesta comanda?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
                    dbtmp.Execute "update comandes_extres set codigrupmaterialcompatible=0 where comanda=" + atrim(cadbl(reixa.Columns("comanda")))
                    comprescomandes.Recordset.Edit
                    comprescomandes.Recordset!compatible = ""
                    comprescomandes.Recordset.Update
                    comprescomandes.Recordset.Move 0
                End If
              End If
      End If
  End If
End Sub
Function nomgrupcompatible(vcodi As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select nomdelgrup from grupsmaterialscompatibles where numerodegrup=" + atrim(vcodi))
   If Not rst.EOF Then nomgrupcompatible = atrim(rst!nomdelgrup)
   Set rst = Nothing
End Function


Sub escullir_familiacompatible(vcodifam As Double, vdescfam As String)
  Dim rstmat As Recordset
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "palets.mdb"
  formseleccio.Data1.RecordSource = "select numerodegrup,nomdelgrup as [Descripció] from grupsmaterialscompatibles order by nomdelgrup"
  formseleccio.refrescar
  formseleccio.width = 5500
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 4000
  formseleccio.Show 1
  If seleccioret = 1 Then
   vcodifam = cadbl(formseleccio.Data1.Recordset!numerodegrup)
   vdescfam = atrim(formseleccio.Data1.Recordset.Fields("Descripció"))
  End If
  If seleccioret = 9 Then
    vcodifam = -1
  End If
  Unload formseleccio
End Sub
Private Sub reixa_DblClick()
  Dim rstc As Recordset
  Dim numc As Double
  If cmaterial = "" Then MsgBox "Primer escull el material corresponent": Exit Sub
If reixa.Columns(reixa.col).DataField = "seleccionat" Then
   If reixa.Text = "" Then
      If (cadbl(ckilosprecomanda) > 0 Or cadbl(ckiloslliures) > 0) And ckiloslliures.tag = "" Then
         MsgBox "Hi ha kilos lliures o de precomanda per aquest material", vbInformation + vbOKOnly, "Atenció"
         ckiloslliures.tag = "1"
      End If
      If mirarsiesmaterialexacte(cadbl(reixa.Columns("Comanda"))) Then Exit Sub
      
      reixa.Columns("Seleccionat") = "S"
      
     Else: If reixa.Text = "S" Then reixa.Text = ""
   End If
   reixa.EditActive = False
End If
If reixa.Columns(reixa.col).DataField = "kgcomprats" Then
    reixa.col = 0
    numc = cadbl(reixa.Text)
    Set rstc = comandescompra.capcalera.Database.OpenRecordset("SELECT capcalera.numcomanda FROM capcalera RIGHT JOIN (liniescompra RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra) ON capcalera.id = liniescompra.idcompra WHERE (((comandesxlinia.numcomanda)=" + atrim(numc) + "));")
    If Not rstc.EOF Then
      If MsgBox("La comanda de compra es la " + atrim(rstc!numcomanda) + Chr(10) + "VOLS CARREGAR-LA?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
         comandescompra.capcalera.RecordSource = "select * from capcalera where numcomanda=" + atrim(rstc!numcomanda)
         comandescompra.capcalera.Refresh
         comandespendents.Hide
      End If
    End If
    Set rstc = Nothing
End If
End Sub
Function mirarsiesmaterialexacte(numc As Double) As Boolean
   Dim rstc As Recordset
   materialexacte = False
   Set rstc = dbtmp.OpenRecordset("select materialexacte from comandes_extres where comanda=" + atrim(numc))
   If Not rstc.EOF Then materialexacte = cabool(rstc!materialexacte)
   Set rstc = Nothing
   If materialexacte Then
       If Not capseleccionada Then
         MsgBox "Aquest material s'ha de comprar per separat ja que el client vol exactament aquest codi de material." + Chr(10) + "TREU TOTES LES COMANDES SELECCIONADES I TORNA A MARCAR AQUESTA", vbInformation, "Atenció"
         mirarsiesmaterialexacte = True
       End If
   End If
End Function
Private Sub reixa_KeyDown(KeyCode As Integer, Shift As Integer)
If UCase(Chr$(KeyCode)) = "S" Then reixa.col = 12: reixa_DblClick
End Sub

Private Sub reixamirar_DblClick()
   comandescompra.capcalera.RecordSource = "capcalera"
   comandescompra.capcalera.Recordset.FindFirst "numcomanda=" + atrim(cadbl(reixamirar.Columns("numcomanda")))
End Sub

Private Sub reixamirar_HeadClick(ByVal ColIndex As Integer)
   Dim direccio As String
   reixamirar.ClearSelCols
   ratoli "espera"
   direccio = " ASC"
   If reixamirar.tag = reixamirar.Columns(ColIndex).DataField + " ASC" Then
      direccio = " DESC"
   End If
   reixamirar.tag = reixamirar.Columns(ColIndex).DataField + direccio
   consultar_Click
    'If Not mirarcompres.Recordset.EOF Then reixamirar.Row = 0: reixamirar.SetFocus
    
   ratoli "normal"
End Sub

Private Sub selfamilies_Click()
   
  
  
   'formseleccio.refrescar
   'formseleccio.DBGrid2.Columns("id").Width = 0
   'formseleccio.DBGrid2.Columns("migelaborat").Width = 200
   'formseleccio.DBGrid2.Columns("material").Width = 800
   'formseleccio.DBGrid2.Columns("descripcio").Width = 2800
   'formseleccio.DBGrid2.Columns("micres").Width = 500
   'formseleccio.DBGrid2.Columns("metresdisponibles").Width = 600
   
  ' formseleccio.Show 1
   
End Sub
Sub possarfamiliesialtresicombomaterials()
  Dim rstf As Recordset
  Dim descmat As String
  Dim mespesor As String
  Dim ultespesor As Double
  Dim ulttol As String
  Dim espesor As Double
  ratoli "espera"
  cmaterial.Clear
  filtarfamiliesipossarmetresdisponibles
  Set rstf = dbconsulta.OpenRecordset("select * from familiescomprescomandes order by migelaborat,descripcio,micres")
  While Not rstf.EOF
    If atrim(rstf!descripcio) = descmat And ulttol = atrim(rstf!migelaborat) And cadbl(rstf!micres) = ultespesor Then GoTo proxim
    mespesor = IIf(cadbl(rstf!micres) > 0, "", " Gr")
    espesor = IIf(cadbl(rstf!micres) < 0, cadbl(rstf!micres) * -1, cadbl(rstf!micres))
    cmaterial.AddItem atrim(rstf!migelaborat) + " " + justifica(rstf!descripcio, 50) + "_" + justifica(Format(espesor, "0.0"), 7) + mespesor
    cmaterial.ItemData(cmaterial.NewIndex) = rstf!id
    descmat = atrim(rstf!descripcio)
    ultespesor = cadbl(rstf!micres)
    ulttol = atrim(rstf!migelaborat)
proxim:
    rstf.MoveNext
  Wend
  Set rstf = Nothing
  comandespendents.SetFocus
  cmaterial.SetFocus
  If cmaterial.ListCount > 0 Then SendKeys "%{DOWN}"
  ratoli "normal"
End Sub
Function justifica(valor As Variant, pos As Byte) As String
    Dim v As String
    v = atrim(valor)
    v = Mid(v, 1, pos)
    If cadbl(valor) > 0 Then
       justifica = String(pos - Len(v), " ") + v
      Else: justifica = v + String(pos - Len(v), " ")
    End If
End Function
Sub filtarfamiliesipossarmetresdisponibles()
   Dim rstf As Recordset
   Dim valors As String
   Dim rstm As Recordset
   Dim mtrsdis As Double
   Dim descmat As String
   dbconsulta.Execute "delete * from familiescomprescomandes"
   Set rstf = dbconsulta.OpenRecordset("SELECT distinct semielaborat, material, espesor From comprescomandes GROUP BY semielaborat, material, espesor;")
   
   
   While Not rstf.EOF
'     mtrsdis = metresdisponibles(rstf)
     Set rstm = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstf!material)))
     descmat = descripciomaterial(rstm)
     valors = "('" + atrim(rstf!semielaborat) + "'," + atrim(cadbl(rstf!material)) + ",'" + descmat + "'," + passaradecimalpunt(cadbl(rstf!espesor)) + "," + atrim(cadbl(mtrsdis)) + ")"
     dbconsulta.Execute "insert into familiescomprescomandes (migelaborat,material,descripcio,micres,metresdisponibles) values " + valors
     rstf.MoveNext
   Wend
   
   
   
End Sub
Function metresdisponibles(rstf As Recordset) As Double
  Dim rstb As Recordset
  Dim were As String
  were = " palets.micres=" + atrim(cadbl(rstf!espesor)) + " and palets.semielaborat='" + atrim(rstf!semielaborat) + "' and palets.codimatprognou=" + atrim(cadbl(rstf!material))
  Set rstb = dbstocks.OpenRecordset("SELECT Sum(Bobines.disponible) AS totaldisponible, Palets.codimatprognou, Palets.micres, Palets.semielaborat FROM Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet Where (((Palets.disponible) = True) And ((Bobines.disponible) > 0)) GROUP BY Palets.codimatprognou, Palets.micres, Palets.semielaborat HAVING " + were + ";")
  If Not rstb.EOF Then
    metresdisponibles = rstb!totaldisponible
   Else: metresdisponibles = 0
  End If
  
End Function
Private Sub sortir_Click()
   comandespendents.Hide
   comandescompra.SetFocus
End Sub
Function descripciomaterial(rstmat As Recordset) As String
  Dim desc As String
  Dim rstfam As Recordset
  Set rstfam = dbtmp.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(cadbl(rstmat!familia)))
  If Not rstfam.EOF Then desc = desc + atrim(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from subfamiliesmaterials where codi=" + atrim(cadbl(rstmat!subfamilia)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from familiescolorants where codi=" + atrim(cadbl(rstmat!familiacol)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from subfamiliescolorants where codi=" + atrim(cadbl(rstmat!subfamiliacol)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from familiesaditius where codi=" + atrim(cadbl(rstmat!familiaad)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from subfamiliesaditius where codi=" + atrim(cadbl(rstmat!subfamiliaad)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  descripciomaterial = desc
End Function

Function af(v As Variant) As String
  v = atrim(v)
  If Len(v) > 1 Then
     v = " + " + v
    Else: v = ""
  End If
  af = v
End Function
