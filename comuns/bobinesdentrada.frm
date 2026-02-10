VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form bobinesdentrada 
   Caption         =   "Bobines d'Entrada"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   4335
   Begin VB.CommandButton mantenimentbob 
      Height          =   390
      Left            =   3435
      Picture         =   "bobinesdentrada.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Manteniment de la bobina anònima."
      Top             =   975
      Width           =   390
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   3480
      Top             =   1305
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton imprimirparcial 
      Height          =   390
      Left            =   3465
      Picture         =   "bobinesdentrada.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir l'etiqueta de bobina d'entrada parcial."
      Top             =   525
      Width           =   390
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   390
      Left            =   3450
      Picture         =   "bobinesdentrada.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   75
      Width           =   390
   End
   Begin VB.TextBox multiseleccio 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Text            =   "1"
      Top             =   315
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   4395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   7752
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   16761024
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton desbloquejat 
      Height          =   390
      Left            =   3420
      Picture         =   "bobinesdentrada.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Bloqueja l'edició de la reixa"
      Top             =   3555
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton bloquejat 
      Height          =   390
      Left            =   3420
      Picture         =   "bobinesdentrada.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Permet l'edició del camp de utilitzat."
      Top             =   3990
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Frame frameescanerbobina 
      Height          =   585
      Left            =   30
      TabIndex        =   7
      Top             =   4290
      Width           =   4110
      Begin VB.CommandButton alta 
         Height          =   420
         Left            =   3615
         Picture         =   "bobinesdentrada.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Escriure el palet manualment"
         Top             =   120
         Width           =   450
      End
      Begin VB.TextBox ccodidebarres 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   915
         MaxLength       =   10
         TabIndex        =   8
         Top             =   120
         Width           =   1950
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Bobina:"
         Height          =   240
         Left            =   75
         TabIndex        =   9
         Top             =   195
         Width           =   1050
      End
      Begin VB.Image Image1 
         Height          =   435
         Left            =   2955
         Picture         =   "bobinesdentrada.frx":213C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   450
      End
   End
   Begin VB.Image nocheck 
      Height          =   180
      Left            =   480
      Picture         =   "bobinesdentrada.frx":2B3F
      Stretch         =   -1  'True
      Top             =   630
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image check 
      Height          =   180
      Left            =   285
      Picture         =   "bobinesdentrada.frx":2D31
      Stretch         =   -1  'True
      Top             =   660
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "bobinesdentrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vcancelartecles As Boolean
Dim iniconfigreixa As String


Function esrestu(ByVal palet As Double, ByVal bobina As Integer) As Boolean
  Dim rstparcial As Recordset
  Dim mtsdisponibles As Double
  Dim mts As Double
  esretu = False
  actualitzar_metres_disponibles palet, bobina
  Set rstparcial = dbstocks.OpenRecordset("select disponible,mts from bobines where idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)))
  If Not rstparcial.EOF Then mtsdisponibles = cadbl(rstparcial!disponible): mts = cadbl(rstparcial!mts)
  Set rstparcial = dbstocks.OpenRecordset("select metres,utilitzada  from parcials where not utilitzada and idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)))
  If Not rstparcial.EOF Then
   rstparcial.MoveLast
   If mtsdisponibles <= 0 And rstparcial.RecordCount = 1 And rstparcial!metres < mts Then esrestu = True
     Else: If mtsdisponibles < mts Then esrestu = True
  End If
End Function
Function esparcial(ByVal palet As Double, ByVal bobina As Integer) As Boolean
  Dim rstparcial As Recordset
  Dim mtsdisponibles As Double
  Dim mts As Double
  esparcial = False
  actualitzar_metres_disponibles palet, bobina
  Set rstparcial = dbstocks.OpenRecordset("select disponible,mts from bobines where idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)))
  If Not rstparcial.EOF Then mtsdisponibles = cadbl(rstparcial!disponible): mts = cadbl(rstparcial!mts)
  Set rstparcial = dbstocks.OpenRecordset("select metres,utilitzada  from parcials where not utilitzada and idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)))
  If Not rstparcial.EOF Then
    rstparcial.MoveLast
    If rstparcial.RecordCount > 1 Then esparcial = True
    If rstparcial.RecordCount = 1 And mtsdisponibles > 0 Then esparcial = True
  End If
End Function



Private Sub alta_Click()
  If Now > CVDate("19/02/2021") Then
     MsgBox "Avís" + Chr(10) + "Si entres la bobina manualment s'enviarà un e-mail a oficina informant.", vbOKOnly + vbInformation, "Atenció"
     Form1.botoensenyarpacking.Tag = "afegidamanualment"
  End If
  Command1.Enabled = True
  ccodidebarres.MaxLength = 9
  ccodidebarres.SetFocus
  
End Sub

Private Sub bloquejat_Click()
  desbloquejat.Visible = True
  bloquejat.Visible = False
  MsgBox "Atenció aquest canvis s'han de fer quan s'està segur que la bobina està acavada.", vbCritical + vbOKOnly, "Atenció"
End Sub

Sub comprobarnumerobobina(Optional vmarcar As Boolean)
seleccionar_un 0
 If siexisteixlabobinaalallista(ccodidebarres) Then
     ccodidebarres.BackColor = QBColor(10)
     seleccionar_un 1 'selecciona la bobina
'     Command1_Click 'accepta la bobina
      Else: ccodidebarres.BackColor = QBColor(12)
  End If
  If ccodidebarres.Text = "" Then ccodidebarres.BackColor = QBColor(15)
End Sub
Function buscarlabobinaalareixa(vpalet As Double, vbob As Double) As Boolean
   Dim vi As Integer

   For vi = 1 To reixa.Rows - 1
      If cadbl(reixa.TextMatrix(vi, 1)) = atrim(vpalet) And cadbl(reixa.TextMatrix(vi, 2)) = atrim(vbob) Then
         reixa.row = vi
         reixa.col = 0
         buscarlabobinaalareixa = True: GoTo sortir
      End If
   Next vi
sortir:
   
End Function
Function siexisteixlabobinaalallista(vcodi As String) As Boolean
   Dim vpalet As Double
   Dim vbob As Double
   Dim vcont As Double
   vcodi = atrim(vcodi)
   'vcodi = "50442-1"
   convertirScanambPaletiBobina2 vcodi, vpalet, vbob
   If vbob > 0 Then
       If buscarlabobinaalareixa(vpalet, vbob) Then siexisteixlabobinaalallista = True
   End If
End Function


Private Sub ccodidebarres_KeyDown(KeyCode As Integer, Shift As Integer)
  Static vhoraultimaentrada As Date
  vcancelartecles = False
  If KeyCode = 13 Then
     vcancelartecles = True
     comprobarnumerobobina True
     Command1_Click 'accepta la bobina si fa enter
     Exit Sub
  End If
  If Len(ccodidebarres) < 1 And ccodidebarres.MaxLength > 9 Then vhoraultimaentrada = Now
 'If vhoraultimaentrada = 0 Then vhoraultimaentrada = Now
If ccodidebarres.MaxLength > 9 And DateDiff("s", vhoraultimaentrada, Now) >= 1 Then
   ccodidebarres = ""
   vcancelartecles = True
   If Len(ccodidebarres) > 0 Then KeyCode = 0
 End If
 DoEvents
 ' vhoraultimaentrada = Now
 
 comprobarnumerobobina
End Sub

Private Sub ccodidebarres_KeyPress(KeyAscii As Integer)
 'Static vhoraultimaentrada As Date
' If vhoraultimaentrada = 0 Then vhoraultimaentrada = Now
'If ccodidebarres.MaxLength > 9 And DateDiff("s", vhoraultimaentrada, Now) > 1 Then
'   ccodidebarres = ""
'   If Len(ccodidebarres) > 0 Then KeyAscii = 0
' End If
' If Len(ccodidebarres) > 0 And ccodidebarres.MaxLength > 9 Then vhoraultimaentrada = Now
' comprobarnumerobobina
If vcancelartecles Then KeyAscii = 0
End Sub

Private Sub codidebarres_Change()

End Sub

Private Sub Command1_Click()
  'ccodidebarres_Change
  bobinesdentrada.Tag = "acceptar"
  bobinesdentrada.Hide
End Sub

Private Sub Command2_Click()

End Sub

Private Sub desbloquejat_Click()
desbloquejat.Visible = False
  bloquejat.Visible = True
End Sub

Private Sub Form_Activate()
  If InStr(1, UCase(App.EXEName), "SOLDADORES") > 0 Then Command1.Enabled = True
  If Form1.botoensenyarpacking.Tag = "afegidamanualment" Then
   Command1.Enabled = True
   ccodidebarres.MaxLength = 9
  End If
  ccodidebarres.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 110 Then KeyCode = 188
End Sub

Private Sub Form_Resize()
 If bobinesdentrada.Width < 1000 Then Exit Sub
 If bobinesdentrada.Height < 1000 Then Exit Sub
 reixa.Width = bobinesdentrada.Width - 700
 reixa.Height = bobinesdentrada.Height - reixa.Top - 1000
 Command1.Left = bobinesdentrada.Width - Command1.Width - 300
 bloquejat.Left = Command1.Left
 bloquejat.Top = bobinesdentrada.Height - bloquejat.Height - 600
 desbloquejat.Left = bloquejat.Left
 desbloquejat.Top = bloquejat.Top
 frameescanerbobina.Top = bobinesdentrada.Height - frameescanerbobina.Height - 550
 imprimirparcial.Left = Command1.Left 'bobinesdentrada.Width - imprimirparcial.Width - 150
 mantenimentbob.Left = Command1.Left 'imprimirparcial.Left
End Sub

Private Sub Image1_Click()
  ccodidebarres.SetFocus
End Sub

Private Sub Image1_DblClick()
 ccodidebarres.SetFocus
End Sub

Private Sub imprimirparcial_Click()
  Dim numpalet As Double
  Dim numbobina As Double
  numpalet = cadbl(reixa.TextMatrix(reixa.row, bobinesdentrada.columnadelcamp("idpalet")))
  numbobina = cadbl(reixa.TextMatrix(reixa.row, bobinesdentrada.columnadelcamp("idbobina")))
  If numpalet > 0 And numbobina > 0 Then
    imprimir_bobinaparcial numpalet, numbobina, , 1
  End If
End Sub
Sub borrartmps()
  On Error Resume Next
  Kill "c:\temp\~imp*.*"
  On Error GoTo 0
End Sub
Function fernomtaulatmp() As String
  borrartmps
  fernomtaulatmp = "c:\temp\~imp" + format(Now, "ddhhnnss") + ".mdb"
  DBEngine.CreateDatabase fernomtaulatmp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
End Function
Sub crearindextaulatmp(fitxer As String)
   Dim dbt As Database
   Set dbt = OpenDatabase(fitxer)
   dbt.Execute "create index p on palets (idpalet) with PRIMARY"
   dbt.Execute "create index codimatprognou on palets (codimatprognou)"
   dbt.Execute "create index a on parcials (idpalet,idbobina)"
   
  ' dbt.Execute "create index [C7174766-2C27-4C22-BEB0-3B55358F1C2B] on bobines (idpalet)"
   dbt.Execute "create index PrimaryKey on bobines (idpalet,idbobina) with PRIMARY"
   dbt.Close
   Set dbt = Nothing
End Sub

Sub imprimir_bobinaparcial(numpalet As Double, numbobina As Double, Optional noferload = False, Optional copies As Byte)
 Dim nomtaulatmp As String
 Dim esstoc As Boolean
 Dim rstestoc As Recordset
 Dim numestoc As Double
 Dim rstbob As Recordset
 netejarreport llistat
 nomtaulatmp = fernomtaulatmp
 If numpalet > 120000 Then MsgBox "Nomes es pot imprimir full de parcial del material anónim.": Exit Sub
 If cadbl(copies) = 0 Then copies = 1
 dbstocks.Execute "select * into palets IN '" + nomtaulatmp + "' from palets where idpalet=" + atrim(numpalet)
 dbstocks.Execute "select * into bobines IN '" + nomtaulatmp + "' from bobines where idpalet=" + atrim(numpalet) + " and idbobina=" + atrim(numbobina)
 dbstocks.Execute "select * into parcials IN '" + nomtaulatmp + "' from parcials where idpalet=" + atrim(numpalet) + " and idbobina=" + atrim(numbobina)
 crearindextaulatmp nomtaulatmp
 Set rstestoc = dbstocks.OpenRecordset("select * from parcials where idpalet=" + atrim(numpalet) + " and idbobina=" + atrim(numbobina) + " and (cdbl(comanda)<3000 and cdbl(comanda)>2000)")
 Set rstbob = dbstocks.OpenRecordset("select * from bobines where idpalet=" + atrim(numpalet) + " and idbobina=" + atrim(numbobina))
 If rstbob.EOF Then Exit Sub
 If Not rstestoc.EOF Then esestoc = True: numestoc = atrim(rstestoc!comanda)
 Set rstestoc = Nothing
If noferload Then bobinesdentrada.Tag = "noload"
'llistat.DiscardSavedData = True
llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "etbobparcial.rpt"
'llistat.ReportFileName = "C:\ETBOBPARCIAL.RPT"
'llistat.Destination = crptToWindow
llistat.Destination = crptToPrinter
For i = 0 To 10
 llistat.DataFiles(i) = llegir_ini("General", "cami", "comandes.ini")
Next i
llistat.DataFiles(0) = nomtaulatmp 'llegir_ini("General", "ruta_stocks", "comandes.ini")
llistat.DataFiles(1) = nomtaulatmp
llistat.DataFiles(9) = nomtaulatmp
DoEvents
'wait (4)
 
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 llistat.Formulas(0) = "diameterebobina='Ø Bob: " + calcular_diametre(numpalet, numbobina, cadbl(rstbob!tamanycanutu)) + " cm Canuto " + atrim(cadbl(rstbob!tamanycanutu)) + " cm'"
 llistat.Formulas(1) = "metresdisponiblesreals=" + atrim(calcular_mtrsdispreals(numpalet, numbobina))
 llistat.Formulas(2) = "nomproveidor='" + nomproveidor(numpalet) + "'"
 llistat.Formulas(3) = "diameterebobina15=''"
 llistat.Formulas(4) = "esstoc='" + IIf(esestoc, "S", "") + "'"
 llistat.Formulas(5) = "numestoc=" + atrim(numestoc)
 llistat.CopiesToPrinter = copies
 'llistat.SelectionFormula = "{palets.idpalet}=" + atrim(numpalet) + " and {bobines.idbobina}=" + atrim(numbobina)
 llistat.ReplaceSelectionFormula "{palets.idpalet}=" + atrim(numpalet) + " and {bobines.idbobina}=" + atrim(numbobina)
 llistat.Action = 1
 netejarreport llistat
 'Unload llistat
End Sub
Function nomproveidor(numpalet As Double) As String
   Dim rstmaterial As Recordset
   Dim rstpalet As Recordset
   Dim dbcomandes As Database
   Dim rstpro As Recordset
   If InStr(1, dbtmpb.Name, "comandes.mdb") > 0 Then
      Set dbcomandes = dbtmpb
     Else: Set dbcomandes = dbtmp
   End If
   Set rstpalet = dbstocks.OpenRecordset("select codimatprognou from palets where idpalet=" + atrim(numpalet))
   If Not rstpalet.EOF Then
       Set rstmaterial = dbcomandes.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstpalet!codimatprognou)))
       If Not rstmaterial.EOF Then
         Set rstpro = dbcomandes.OpenRecordset("select nom from proveidors where codi=" + atrim(cadbl(rstmaterial!proveidor)))
         If Not rstpro.EOF Then nomproveidor = atrim(rstpro!nom)
       End If
   End If
   Set rstmaterial = Nothing
   Set rstpalet = Nothing
   Set dbcomandes = Nothing
   Set rstpro = Nothing
End Function
Function calcular_mtrsdispreals(palet As Double, bobina As Double, Optional nocomptar100 As Boolean) As Double
   Dim rstb As Recordset
   Dim rstp As Recordset
   Dim total As Double
   'dbstocks.Execute "delete * from parcials where metres=0 and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina)
   Set rstb = dbstocks.OpenRecordset("select mts from bobines where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
   Set rstp = dbstocks.OpenRecordset("select sum(metres) as tmetres from parcials where utilitzada and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + IIf(nocomptar100, " and (comanda<>'100')", ""))
   If Not rstb.EOF Then total = cadbl(rstb!mts)
   If Not rstp.EOF And Not rstb.EOF Then
      total = cadbl(rstb!mts) - cadbl(rstp!tmetres)
   End If
   calcular_mtrsdispreals = total
   Set rstb = Nothing
   Set rstp = Nothing
End Function
Function calcular_diametre(palet As Double, bobina As Double, Optional canutu As Double) As String
  Dim rstp As Recordset
  Dim rstb As Recordset
  Dim metres As Double
  Dim micres As Double
  
  Dim diametre As Double
  Dim pi As Double
  Set rstp = dbstocks.OpenRecordset("select micres,codimatprognou from palets where idpalet=" + atrim(palet))
  If cadbl(canutu) = 0 Then
       Set rstb = dbstocks.OpenRecordset("select tamanycanutu from bobines where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
       If rstb.EOF Then GoTo fi
       canutu = cadbl(rstb!tamanycanutu)
       If canutu < 5 Or canutu > 20 Then canutu = 7.6  'si el canutu te una mida possiblement incorrecte posso 7.6
  End If
  If canutu < 10 Then canutu = canutu + 2 'afegeixo l'amplada del cartrò del canutu
  If canutu >= 10 Then canutu = canutu + 2.8 'afegeixo l'amplada del cartrò del canutu
  metres = cadbl(calcular_mtrsdispreals(palet, bobina))
  If metres <= 0 Then calcular_diametre = 0: GoTo fi
  Set rstb = dbstocks.OpenRecordset("select * from materials where codi=" + atrim(rstp!codimatprognou))
  If Not rstp.EOF Then
    pi = 4 * Atn(1)
    canutu = (canutu / 2) / 100
    micres = cadbl(rstp!micres)
    If cadbl(rstb!micresdelsgrm2) > 0 Then micres = cadbl(rstb!micresdelsgrm2)
    'If micres < 0 Then
    '   micres = (micres * -1)
    '   micres = micres / 1.2
    'End If
    micres = (micres * 0.0001) / 100
    diametre = Sqr(((metres * micres) / pi) + (canutu * canutu)) * 200
    calcular_diametre = Redondejar(diametre, 0)
    If cadbl(calcular_diametre) < 9 Then calcular_diametre = "0"
  End If
fi:
  Set rstp = Nothing
  Set rstb = Nothing
End Function

Private Sub mantenimentbob_Click()
  Dim numpalet As Double
  Dim numbobina As Double
  Dim taula As String
  Dim numc As Double
  numpalet = cadbl(reixa.TextMatrix(reixa.row, bobinesdentrada.columnadelcamp("idpalet")))
  numbobina = cadbl(reixa.TextMatrix(reixa.row, bobinesdentrada.columnadelcamp("idbobina")))
  taula = atrim(reixa.TextMatrix(reixa.row, bobinesdentrada.columnadelcamp("TAULA")))
  If numpalet > 0 And numbobina > 0 And taula = "parcials" Then
    numc = ncomanda
    If seccioanterior(numc) <> "E" Then numc = ncomanda2
    estatdelabobina numpalet, numbobina, 0, numc
  End If
End Sub

Private Sub palet_Change()

End Sub

Private Sub palet_KeyPress(KeyAscii As Integer)

End Sub

Private Sub reixa_Click()
  seleccionar
End Sub
Sub seleccionar_un(vmarcar As Byte)
  Dim vr As Integer
  Dim vc As Integer
  Dim i As Integer
  vr = reixa.row
  vc = reixa.col
  For i = 1 To reixa.Rows - 1
    reixa.row = i
    reixa.RowSel = i
    reixa.ColSel = 0
    reixa.col = 0
    Set reixa.CellPicture = nocheck.Picture
    reixa.TextMatrix(reixa.row, 0) = "0"
  Next i
  reixa.row = vr
  reixa.col = 0
  reixa.RowSel = vr
  reixa.ColSel = 0
  If Not vmarcar Then
       Set reixa.CellPicture = nocheck.Picture
       reixa.TextMatrix(reixa.row, 0) = "0"
  End If
  If vmarcar Then
       Set reixa.CellPicture = check.Picture
       reixa.TextMatrix(reixa.row, reixa.col) = "1"
  End If
  reixa.row = vr
  reixa.col = vc
End Sub

Sub seleccionar()
 If rstconsulta.Fields(reixa.ColData(reixa.col)).Type = 1 And rstconsulta.Fields(reixa.ColData(reixa.col)).Name <> "utilitzada" Then
  reixa.RowSel = reixa.row
  reixa.ColSel = reixa.col
  If reixa.CellPicture = check.Picture Then
       Set reixa.CellPicture = nocheck.Picture
       reixa.TextMatrix(reixa.row, reixa.col) = "0"
      Else:
       If ok_comprovar_multiseleccio Then
         Set reixa.CellPicture = check.Picture: reixa.TextMatrix(reixa.row, reixa.col) = "1"
       End If

  End If
 End If
 'If Not bloquejat.Visible And rstconsulta.Fields(reixa.ColData(reixa.col)).Type = 1 And rstconsulta.Fields(reixa.ColData(reixa.col)).Name = "utilitzada" Then
 ' reixa.RowSel = reixa.row
 ' reixa.ColSel = reixa.col
 '  If reixa.CellPicture = check.Picture Then
 '      Set reixa.CellPicture = nocheck.Picture
 '      reixa.TextMatrix(reixa.row, reixa.col) = "0"
 '      carregar_bobinesdentrada "marcarutilitzada", , reixa.TextMatrix(reixa.row, columnadelcamp("idpalet")), reixa.TextMatrix(reixa.row, columnadelcamp("idbobina")), ncomanda, False, ncomanda2
'      Else:
'         Set reixa.CellPicture = check.Picture
'         reixa.TextMatrix(reixa.row, reixa.col) = "1"
'         carregar_bobinesdentrada "marcarutilitzada", , reixa.TextMatrix(reixa.row, columnadelcamp("idpalet")), reixa.TextMatrix(reixa.row, columnadelcamp("idbobina")), ncomanda, True, ncomanda2
 '  End If
 'End If
End Sub
Function ok_comprovar_multiseleccio()
   Dim cont As Integer
   
   For i = 1 To reixa.Rows - 1
      If reixa.TextMatrix(i, reixa.col) = "1" Then
        cont = cont + 1
      End If
   Next i
   If cont < cadbl(multiseleccio) Then
      ok_comprovar_multiseleccio = True
     Else: ok_comprovar_multiseleccio = False
   End If
End Function

Private Sub reixa_DblClick()
   If Command1.Enabled = False Then Exit Sub
   If reixa.TextMatrix(0, reixa.col) = "UTILITZADA" And InStr(1, reixa.TextMatrix(reixa.row, 6), "bobines") > 0 Then
      If MsgBox("Segur que vols passar aquesta bobina a NO UTILITZADA?", vbInformation + vbYesNo + vbDefaultButton2, "ATENCIÓ") = vbYes Then
         dbbaixes.Execute "update  " + atrim(reixa.TextMatrix(reixa.row, 6)) + " set utilitzadaabaixa=false where id=" + atrim(cadbl(reixa.TextMatrix(reixa.row, 7)))
      End If
      Unload bobinesdentrada
   End If
End Sub

Private Sub reixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 2 Then canviarnomcapcalera: Exit Sub
End Sub
Sub canviarnomcapcalera()
    r = InputBox("Entra el nom que vols a la capçalera", "Canvi de nom de columna")
    If atrim(r) <> "" Then
          escriure_ini "NomsReixa", UCase(rstconsulta.Fields(reixa.ColData(reixa.col)).Name) + "-nom", r, iniconfigreixa
          reixa.TextMatrix(0, reixa.col) = r
    End If
End Sub


Sub actualitzar_metres_disponibles(ByVal palet As Double, ByVal bobina As Double)
  Dim rstparcial As Recordset
  Dim total As Double
  total = 0
  
  'Set rstparcial = dbstocks.OpenRecordset("select sum(metres) as total from parcials where cdbl(comanda)>10000  and idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)))
  Set rstparcial = dbstocks.OpenRecordset("select sum(metres) as total from parcials where idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)))
  If Not rstparcial.EOF Then total = cadbl(rstparcial!total)
  dbstocks.Execute "update bobines set disponible=mts-" + atrim(total) + " where idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina))
  Set rstparcial = Nothing
End Sub





Private Sub Form_Load()
  carregardades
End Sub
Sub carregardades()
 If nohiharegistres(rstconsulta) Then Exit Sub
 iniconfigreixa = "reixabaixesimpresora.ini"
 reixa.Rows = 1
 configurar_reixa
 poblar_reixa
 bobinesdentrada.Tag = ""
End Sub
Function columnadelcamp(nom As String) As Integer
  Dim i As Byte
  nom = UCase(nom)
  If reixa.Cols < 3 Then columnadelcamp = 0: Exit Function
  For i = 0 To reixa.Cols - 1
    If UCase(rstconsulta.Fields(reixa.ColData(i)).Name) = nom Then columnadelcamp = i
  Next i

End Function
Function nohiharegistres(rstc As Recordset) As Boolean
  Dim rec As Double
  On Error GoTo sortir
  reg = rstc.RecordCount
  nohiharegistres = False
  Exit Function
sortir:
  nohiharegistres = True
  
End Function

Sub configurar_reixa()
  Dim col As Integer
  Dim i As Integer
  'reixa.Clear
  If nohiharegistres(rstconsulta) Then Exit Sub
  col = 0
  reixa.Rows = 2
  reixa.Cols = rstconsulta.Fields.Count
  reixa.FixedRows = 1
  reixa.FixedCols = 0
  For i = 0 To rstconsulta.Fields.Count - 1
     reixa.col = col
     reixa.ColData(i) = i
     reixa.TextMatrix(0, col) = UCase(rstconsulta.Fields(i).Name)
     reixa.ColWidth(col) = IIf(rstconsulta.Fields(i).Size > 25, 25, rstconsulta.Fields(i).Size) * 200
     If LCase(rstconsulta.Fields(i).Name) = "idb" Then reixa.ColWidth(col) = 0
     col = col + 1
  Next i
  carregar_amples_reixa
End Sub
Sub carregar_amples_reixa()
 Dim ample As String
 If existeix("c:\windows\" + iniconfigreixa) Then
  For j = 0 To reixa.Cols - 1
   ample = llegir_ini("AmplesReixa", UCase(reixa.TextMatrix(0, j)), iniconfigreixa)
   
   If ample <> "{[}]" Then reixa.ColWidth(j) = cadbl(ample)
   r = llegir_ini("NomsReixa", UCase(rstconsulta.Fields(reixa.ColData(j)).Name) + "-nom", iniconfigreixa)
   If r <> "{[}]" Then
      reixa.TextMatrix(0, j) = r
   End If
 Next j
 If cadbl(llegir_ini("Amplesformulari", "ample", iniconfigreixa)) > 1000 Then
  
  bobinesdentrada.Width = cadbl(llegir_ini("Amplesformulari", "ample", iniconfigreixa))
  bobinesdentrada.Height = cadbl(llegir_ini("Amplesformulari", "alt", iniconfigreixa))
  bobinesdentrada.Top = cadbl(llegir_ini("Posicioformulari", "top", iniconfigreixa))
  bobinesdentrada.Left = cadbl(llegir_ini("Posicioformulari", "left", iniconfigreixa))
    Else
      bobinesdentrada.Top = (Screen.Height / 2) - (bobinesdentrada.Height / 2)
      bobinesdentrada.Left = (Screen.Width / 2) - (bobinesdentrada.Width / 2)
 End If
 Form_Resize
End If
End Sub
Sub guardar_amples_reixa()
If iniconfigreixa <> "" Then
  For j = 0 To reixa.Cols - 1
   escriure_ini "AmplesReixa", UCase(rstconsulta.Fields(reixa.ColData(j)).Name), atrim(reixa.ColWidth(j)), iniconfigreixa
 Next j
End If
escriure_ini "Amplesformulari", "ample", atrim(bobinesdentrada.Width), iniconfigreixa
escriure_ini "Amplesformulari", "alt", atrim(bobinesdentrada.Height), iniconfigreixa
escriure_ini "Posicioformulari", "left", atrim(bobinesdentrada.Left), iniconfigreixa
escriure_ini "Posicioformulari", "top", atrim(bobinesdentrada.Top), iniconfigreixa
End Sub

Sub poblar_reixa()
  Dim row As Integer
  Dim col As Integer
  Dim vample As Double
  Dim vampleant As Double
  Dim valor As String
  reixa.Rows = 2
  reixa.FillStyle = flexFillRepeat
  reixa.BackColor = QBColor(15)

  row = 1
  If Not rstconsulta.EOF Then
     rstconsulta.MoveFirst
    Else: Exit Sub
  End If
  While Not rstconsulta.EOF
    'posso els checks si el camp es check i si no el valor corresponent a cada camp
    For col = 0 To rstconsulta.Fields.Count - 1
      reixa.col = col
      reixa.row = row
     If rstconsulta!tipus = "O" Then reixa.CellBackColor = QBColor(14) 'bobina jumbo
     If rstconsulta!tipus = "R" Then reixa.CellBackColor = QBColor(10) 'restu
     If rstconsulta!tipus = "P" Then reixa.CellBackColor = QBColor(13) 'parcial
     If rstconsulta!tipus = "Z" Then reixa.CellBackColor = QBColor(2) 'bobina acavada
     If rstconsulta!tipus = "I" Or rstconsulta!tipus = "L" Then reixa.CellBackColor = QBColor(7)
     If rstconsulta.Fields(col).Type = 1 Then
        
        reixa.TextMatrix(row, col) = IIf(rstconsulta.Fields(col), "1", "0")
        Set reixa.CellPicture = IIf(reixa.TextMatrix(row, col) = "0", nocheck.Picture, check.Picture)
        reixa.CellForeColor = reixa.CellBackColor
         Else:
           'valor = formatreixa(rstconsulta.Fields(col))
           If rstconsulta.Fields(col).Name = "idb" Then
                valor = cadbl(rstconsulta.Fields(col))
                 Else:
                   If rstconsulta.Fields(col).Type = 10 Then
                      valor = atrim(rstconsulta.Fields(col))
                        Else: valor = cadbl(rstconsulta.Fields(col))
                   End If
           End If
           reixa.TextMatrix(row, col) = valor
           'reixa.col = col
           'reixa.row = row
           If cadbl(reixa.TextMatrix(row, col)) < 0 Then
              reixa.CellForeColor = QBColor(12)
             Else: reixa.CellForeColor = QBColor(0)
           End If
     End If
    Next col
    'incremento la fila
    reixa.Rows = row + 2
    row = row + 1
    reixa.row = row
    rstconsulta.MoveNext

  Wend
  reixa.Rows = row
End Sub
Function formatreixa(ByVal valor As Variant) As String
   If cadbl(valor) <> 0 Then
         If (cadbl(valor) - Int(cadbl(valor))) <> 0 Then
            valor = Redondejar(cadbl(valor), 1)
           Else: valor = Redondejar(CDbl(valor), 0)
         End If
   End If
   If IsNull(valor) Then valor = ""
   formatreixa = valor
End Function

Private Sub Form_Unload(Cancel As Integer)
 If bobinesdentrada.Visible Then
   guardar_amples_reixa
 End If
End Sub


Sub convertirScanambPaletiBobina2(vcodi As String, vpalet As Double, vbob As Double)
   Dim vcont As Double
   vcodi = atrim(vcodi)
   While vcont < Len(vcodi)
      If Not IsNumeric(Mid(vcodi, vcont + 1, 1)) Then
        vpalet = cadbl(Mid(vcodi, 1, vcont))
        If Len(vcodi) >= vcont + 2 Then vbob = cadbl(Mid(vcodi, vcont + 2))
        GoTo sortir
      End If
      vcont = vcont + 1
   Wend
sortir:
End Sub

