VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formconsultaestats 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Consulta de l'estat dels treballs"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15270
   ControlBox      =   0   'False
   Icon            =   "Formconsultaestats.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   15180
   ScaleWidth      =   24960
   StartUpPosition =   2  'CenterScreen
   Begin Proyecto1.MouseWheel MouseWheel1 
      Left            =   5325
      Top             =   345
      _ExtentX        =   926
      _ExtentY        =   318
   End
   Begin VB.CommandButton bposicioordre 
      Height          =   315
      Left            =   1695
      Picture         =   "Formconsultaestats.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Eliminar totes les linies"
      Top             =   1020
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton bordre 
      Height          =   315
      Left            =   0
      Picture         =   "Formconsultaestats.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar totes les linies"
      Top             =   15
      Width           =   300
   End
   Begin VB.Frame Fbotons 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   13305
      TabIndex        =   4
      Top             =   -15
      Width           =   1935
      Begin VB.CommandButton sortir 
         Height          =   480
         Left            =   1260
         Picture         =   "Formconsultaestats.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Sortir"
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton exportaraxls 
         BackColor       =   &H00F0F0F0&
         Height          =   480
         Left            =   30
         Picture         =   "Formconsultaestats.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exportar a Excel la sel.lecció"
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Height          =   480
         Left            =   645
         Picture         =   "Formconsultaestats.frx":1F46
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir sel.lecció"
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.ComboBox filtre 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Index           =   0
      Left            =   75
      TabIndex        =   2
      Top             =   615
      Width           =   555
   End
   Begin VB.CommandButton treurefiltre 
      Height          =   285
      Left            =   0
      Picture         =   "Formconsultaestats.frx":24D0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Eliminar totes les linies"
      Top             =   330
      Width           =   300
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   6465
      Left            =   15
      TabIndex        =   0
      Top             =   930
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   11404
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorSel    =   16756318
      ForeColorSel    =   16711680
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   3
   End
   Begin VB.Label etregistres 
      Height          =   165
      Left            =   75
      TabIndex        =   11
      Top             =   7755
      Width           =   15150
   End
   Begin VB.Label etordre 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   345
      TabIndex        =   10
      Top             =   60
      Width           =   3885
   End
   Begin VB.Label etmsgajuda 
      BackColor       =   &H0080FFFF&
      Height          =   270
      Left            =   300
      TabIndex        =   3
      Top             =   330
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "Formconsultaestats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fitxertmpestats As String
Dim dbplanificacio As Database
Dim whereultimfiltre As String



Sub configreixa(Optional nocarregaramples As Boolean)
  Dim rst As Recordset
  Dim col As Integer
  Dim enes As Byte
  'reixa.LeftCol = 0
  If reixa.Rows > 1 Then reixa.TopRow = 1
  Set rst = dbconsulta.OpenRecordset("select * from consultaestats")
  col = 0
  enes = 0
  reixa.Cols = rst.Fields.Count
  For i = 0 To rst.Fields.Count - 1
       reixa.ColAlignment(col) = 2
       reixa.TextMatrix(0, col) = campsestat(i + 1, 3)
       If Not nocarregaramples Then colocarfiltre col, i + 1
       
       col = col + 1
  Next i
   
     
  reixa.Cols = reixa.Cols - (enes + 1)
  carregar_amples_reixa
  'reixa.Row = 0
  'For i = 0 To reixa.Cols - 1
  '  reixa.col = i
  '  reixa.ColSel = i
  '  reixa.CellBackColor = QBColor(8)
  'Next i
End Sub
Sub carregar_amples_reixa()
 Dim ample As String
 Dim X As Long
 Dim j As Integer
 If iniconfigreixa <> "" Then ' existeix("c:\windows\" + iniconfigreixa) Then
 
  X = reixa.Left + 35
  For j = 0 To reixa.Cols - 1
   ample = llegir_ini("AmplesReixa", UCase(reixa.TextMatrix(0, j)), iniconfigreixa)
   If ample <> "{[}]" Then
    reixa.ColWidth(j) = cadbl(ample)
    If X < reixa.width Then
     filtre(j).Left = X
     filtre(j).width = cadbl(ample)
     filtre(j).visible = True
     filtre(j).ForeColor = &H808080
      Else: If filtre.Count < j - 1 Then filtre(j).visible = False
    End If
    X = X + cadbl(ample)
   End If
 Next j
End If

End Sub

Function ordredelataula() As String
  If bordre.tag = "" Then
     ordredelataula = " order by cvdate(dataobertura) desc"
    Else: ordredelataula = " order by " + bordre.tag
  End If
End Function
Sub poblarlareixa(Optional were As String)
  Dim i As Byte
  Dim fila As Integer
  Dim col As Byte
  Dim rst As Recordset
  Dim apuntxrimprimir As Double
  Dim tenimmaterial As Boolean
  Dim tenimclixes As Boolean
  Dim textetaula As String
  ratoli "espera"
  etregistres = ""
  reixa.visible = False
  reixa.Clear
  reixa.BackColor = QBColor(15)
  configreixa IIf(were <> "", True, False)
  reixa.Rows = 1
  Set rst = dbconsulta.OpenRecordset("select * from consultaestats " + IIf(were <> "", " where " + were, "") + ordredelataula)
  If rst.EOF Then GoTo fi
  fila = 0
  reixa.tag = "poblant"
  While Not rst.EOF
   fila = fila + 1
   reixa.Rows = fila + 1
   For i = 0 To rst.Fields.Count - 1
     If campsestat(i + 1, 1) <> "" Then
      reixa.TextMatrix(fila, i) = IIf(IsNull(rst.Fields(campsestat(i + 1, 1))), "", rst.Fields(campsestat(i + 1, 1)))
      If reixa.TextMatrix(fila, i) = "0:00:00" Then reixa.TextMatrix(fila, i) = ""
      posarelcolordelcamp fila, i, campsestat(i + 1, 1)
     End If
   Next i
   rst.MoveNext
  Wend
  etregistres.caption = atrim(rst.RecordCount) + " Treballs"
fi:
  Set rst = Nothing
  reixa.tag = ""
  reixa.visible = True
  ratoli "normal"
End Sub
Sub posarelcolordelcamp(fila As Integer, columna As Byte, vcamp As String)
  reixa.col = columna
  reixa.row = fila
  If vcamp = "observacions" Then
     reixa.CellBackColor = &HEAD9CE
      Else: reixa.CellBackColor = &H80000005
  End If
End Sub
Sub colocarfiltre(col As Integer, i As Long)
  If filtre.Count <= col Then Load filtre(col)
  filtre(col).Text = campsestat(i, 3)
  filtre(col).tag = i
'  Load filtre(col + 1)
End Sub


Private Sub bordre_Click()
 etmsgajuda = "Prem sobre la columna que vols ordenar."
 etmsgajuda.width = 3000
 etmsgajuda.Left = treurefiltre.Left + treurefiltre.width + 100
 etmsgajuda.visible = True
 bordre.BackColor = QBColor(14)
 reixa.BackColorFixed = QBColor(14)
End Sub

Private Sub Command3_Click()
   imprimirseleccio False
End Sub
Sub imprimirseleccio(vexportar As Boolean)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistatconsultaestatclixes" + IIf(vexportar, "exportació", "") + ".rpt", 1)
  oreport.Database.Tables.Item(1).SetDataSource dbconsulta.OpenRecordset("select * from consultaestats " + IIf(whereultimfiltre <> "", " where " + whereultimfiltre, "") + ordredelataula), 3
  '"c:\temp\consultaestatstmp.mdb"
  oreport.DiscardSavedData
  If vexportar Then
   oreport.ExportOptions.DiskFileName = "c:\temp\consultaestattreballs.xls"
   oreport.ExportOptions.PDFExportAllPages = True
   oreport.ExportOptions.FormatType = crEFTExcel97 ' crEFTExcel80Tabular
   oreport.ExportOptions.DestinationType = crEDTDiskFile
   oreport.Export False
   obrir_document "c:\temp\consultaestattreballs.xls"
   GoTo fi
  End If
  oreport.PageEngine.ValueFormatOptions = crIncludeFieldValues
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
    Else
      oreport.PrintOut False, 1
  End If
fi:

End Sub

Private Sub exportaraxls_Click()
    'imprimirseleccio True
    generar_xls
End Sub
Function seleccionats(iniciofi As String) As Double
    If reixa.row > reixa.RowSel Then
      If iniciofi = "inici" Then seleccionats = reixa.RowSel
      If iniciofi = "fi" Then seleccionats = reixa.row
        Else
          If iniciofi = "inici" Then seleccionats = reixa.row
          If iniciofi = "fi" Then seleccionats = reixa.RowSel
    End If
End Function
Sub generar_xls()
   Dim i As Byte
   Dim rst As Recordset
   Dim linia As String
   
   Set rst = dbconsulta.OpenRecordset("select * from consultaestats " + IIf(whereultimfiltre <> "", " where " + whereultimfiltre, "") + ordredelataula)
   If rst.EOF Then MsgBox "No hi ha dades per exportar", vbCritical, "Error": Exit Sub
   Open "c:\temp\consultaestattreballs.csv" For Output As #1
   If Not rst.EOF Then
    For i = 0 To rst.Fields.Count - 1
      linia = linia + IIf(linia = "", "", ";") + atrim(campsestat(i + 1, 3))
    Next i
    Print #1, linia
   End If
   While Not rst.EOF
    linia = ""
    If (seleccionats("fi") - seleccionats("inici")) > 0 Then
       If Not tocaexportar(rst.Fields("Treball"), seleccionats("inici"), seleccionats("fi")) Then GoTo proxim
    End If
    For i = 1 To rst.Fields.Count - 1
      linia = linia + IIf(linia = "", "", ";") + """" + IIf(rst.Fields(i).Name = "codibarres", "Nº: ", "") + atrim(rst.Fields(i)) + """"
    Next i
    Print #1, linia
proxim:
    rst.MoveNext
   Wend
   Close #1
   wait 2
   obrir_document "c:\temp\consultaestattreballs.csv"
      
End Sub
Function tocaexportar(treball As String, inici As Double, fi As Double) As Boolean
   Dim i As Long
   tocaexportar = False
   For i = inici To fi
      If reixa.TextMatrix(i, 0) = treball Then tocaexportar = True: GoTo fi
   Next i
fi:
End Function
Private Sub filtre_DropDown(Index As Integer)
    carregar_combo_filtre Index
   
End Sub

Private Sub filtre_GotFocus(Index As Integer)
bxrcontrolagafafocus Index
  ultimfiltre = Index
  If filtre(Index).width < 500 Then filtre(Index).HelpContextID = filtre(Index).width: filtre(Index).width = 1000
End Sub
Sub bxrcontrolagafafocus(i As Integer)
  Dim cntrl As Control
  Set cntrl = Screen.ActiveControl
  If cntrl.Text <> "" Then
     If cntrl.Text = campsestat(cadbl(filtre(i).tag), 3) Then cntrl.Text = ""
     cntrl.ForeColor = QBColor(0)
     
   Else:
       
       cntrl.Text = campsestat(cadbl(filtre(i).tag), 3)
       cntrl.ForeColor = &H808080
  End If
End Sub


Private Sub filtre_LostFocus(Index As Integer)
  Dim noufiltre As String
  If Index = 998 Then whereultimfiltre = "": Exit Sub
  noufiltre = crearfiltre
  If filtre(ultimfiltre).Text = "" Then
    filtre(ultimfiltre).Text = campsestat(cadbl(filtre(ultimfiltre).tag), 3)
    filtre(ultimfiltre).ForeColor = &H808080
    If filtre(ultimfiltre).HelpContextID <> 0 Then filtre(ultimfiltre).width = filtre(ultimfiltre).HelpContextID
  End If
  If noufiltre <> whereultimfiltre Or Index = 999 Then
     If noufiltre <> "" Then poblarlareixa noufiltre
  End If
  If Index = 999 And noufiltre = "" Then
     poblarlareixa
  End If
  ratoli "normal"
  reixa.visible = True
  whereultimfiltre = noufiltre
  'Me.caption = whereultimfiltre
  possaretiquetaajuda
  'Command3.tag = noufiltre ' el guardo pel llistat
  
End Sub
Sub possaretiquetaajuda()
   Dim i As Byte
   etmsgajuda.visible = False
   For i = 0 To filtre.Count - 1
    If InStr(1, filtre(i), ",") > 0 Then
      etmsgajuda.caption = "Una coma busca dos valors"
      etmsgajuda.width = filtre(i).width
      If etmsgajuda.width < 2000 Then etmsgajuda.width = 2000
      etmsgajuda.Left = filtre(i).Left
      etmsgajuda.visible = True
      GoTo fi
    End If
   Next i
fi:
End Sub
Function possarweres(ByVal camp As String, condicio As String, ByVal filtre As String) As String
  Dim re As String
'camps(j, 1) + " LIKE '*" + treure_apostruf(filtre(i)) + "*'"
  filtre = filtre + ","
  If camp = "nomclient" And cadbl(Mid(filtre, 1, InStr(1, filtre, ",") - 1)) > 0 Then camp = "codiclient"
  While InStr(1, filtre, ",") > 0 And filtre <> ""
    If camp <> "codiclient" Then
       
       If camp = "estatclixe" Then
         re = IIf(re <> "", re + " or ", "") + camp + " = '" + Mid(filtre, 1, InStr(1, filtre, ",") - 1) + "'"
           Else: re = IIf(re <> "", re + " or ", "") + camp + " like '*" + Mid(filtre, 1, InStr(1, filtre, ",") - 1) + "*'"
       End If
      Else: re = IIf(re <> "", re + " or ", "") + camp + " =" + atrim(cadbl(Mid(filtre, 1, InStr(1, filtre, ",") - 1))) + ""
    End If
    filtre = Mid(filtre, InStr(1, filtre, ",") + 1)
  Wend
  If re <> "" Then re = "(" + re + ")"
  possarweres = re
End Function
Function crearwere(i As Integer) As String
   Dim w As String
   Dim j As Integer
   If filtre(i) = "" Then Exit Function
   j = cadbl(filtre(i).tag)
   If campsestat(j, 2) = "date" Then
      If IsDate(filtre(i)) Then
         crearwere = campsestat(j, 1) + "=#" + Format(filtre(i), "mm/dd/yy") + "# "
      End If
      Exit Function
   End If
   If InStr(1, campsestat(j, 2), "string") > 0 Then
       crearwere = possarweres(campsestat(j, 1), "LIKE", treure_apostruf(filtre(i)))
       Exit Function
   End If
   crearwere = campsestat(j, 1) + "=" + passaradecimalpunt(atrim(cadbl(filtre(i))))
   

End Function
Function crearfiltre() As String
  Dim i As Integer
  Dim were As String
  Dim w As String
  For i = 0 To filtre.Count - 1
    If filtre(i).Text <> campsestat(cadbl(filtre(i).tag), 3) And campsestat(cadbl(filtre(i).tag), 1) <> "comanda" Then
      w = crearwere(i)
      If were = "" Then
         were = w
        Else: If w <> "" Then were = were + " and " + w
      End If
    End If
  Next i
  crearfiltre = were
End Function


Sub carregar_combo_filtre(Index As Integer)
   Dim rst As Recordset
   Set rst = dbconsulta.OpenRecordset("select distinct " + atrim(campsestat(Index + 1, 1)) + " as valor from consultaestats " + IIf(whereultimfiltre <> "", " where " + whereultimfiltre, "") + " order by " + atrim(campsestat(Index + 1, 1)) + " asc")
   filtre(Index).Clear
   While Not rst.EOF
      If atrim(rst!valor) <> "" Then filtre(Index).AddItem rst!valor
      rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Private Sub Form_Load()
   iniconfigreixa = "c:\windows\clixesconsultaestats.ini"
   fitxertmpestats = "c:\temp\consultaestatstmp.mdb"
   carregartamanyform
   crearfitxertemp
   carregardadesfitxertemporal
   
   poblarlareixa
End Sub
Sub borrarlataula()
   dbconsulta.Execute "delete * from consultaestats"
End Sub
Sub carregardadesfitxertemporal()
   Dim rst As Recordset
   Dim rstnou As Recordset
   Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificacio.mdb", , True)
   borrarlataula
   Set rst = dbclixes.OpenRecordset("SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio, Modificacions.dataobertura, Modificacions.fotograbador, Clixes.nomclienttemporal, [marca]+' - '+[linia] AS marcailinia, Modificacions.pdfvalid, Clixes.codidebarres, Clixes_modifi.data_prevista, Clixes_modifi.data_fi, Clixes_modifi.descripcioestat FROM Clixes INNER JOIN (Clixes_modifi INNER JOIN Modificacions ON (Clixes_modifi.ordremodificacio = Modificacions.ordre) AND (Clixes_modifi.id_treball = Modificacions.id_treball)) ON Clixes.id_treball = Modificacions.id_treball  WHERE (((Clixes_modifi.data_fi) Is Null) and Clixes_modifi.descripcioestat <>'-') order by clixes_modifi.id_treball DESC;")
   Set rstnou = dbconsulta.OpenRecordset("select * from consultaestats")
   While Not rst.EOF
    If rst!descripcioestat <> "-" Then
     copiarregistreatemporal rst, rstnou
    End If
    rst.MoveNext
   Wend
  ' wait 2
   Set rst = Nothing
   Set rstnou = Nothing
   Set dbplanificacio = Nothing
End Sub
Function canviarlacoma(ByVal n As String) As String
   While InStr(n, ",")
     n = Mid(n, 1, InStr(1, n, ",") - 1) + "¸" + Mid(n, InStr(1, n, ",") + 1)
   Wend
   If n = "{[}]" Then n = ""
   canviarlacoma = n
End Function
Sub copiarregistreatemporal(rst As Recordset, rstnou As Recordset)
   rstnou.AddNew
   rstnou!treball = atrim(rst!id_treball) + "/" + atrim(rst!ordremodificacio)
   rstnou!dataobertura = rst!dataobertura
   rstnou!codibarres = rst!codidebarres
   rstnou!client = canviarlacoma(atrim(rst!nomclienttemporal))
   rstnou!fotogravador = nomfotogravador(cadbl(rst!fotograbador))
   rstnou!marcalinia = rst!marcailinia
   rstnou!pdf = IIf(rst!pdfvalid, "Si", "No")
   rstnou!estatclixe = rst!descripcioestat
   rstnou!dataprevista = rst!data_prevista
   rstnou!comandes = buscarcomandes(rst!id_treball, rst!ordremodificacio)
   rstnou!dataentrega = buscardataentrega(IIf(InStr(1, rstnou!comandes, ",") = 0, rstnou!comandes, Mid(rstnou!comandes, 1, InStr(1, rstnou!comandes, ","))))
   rstnou!observacions = carregaobservacio(rstnou!treball)
   rstnou.Update
End Sub
Function carregaobservacio(treball As String) As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select observacions from consultaestats where treball='" + atrim(treball) + "'")
   If Not rst.EOF Then
       carregaobservacio = atrim(rst!observacions)
   End If
End Function
Function buscardataentrega(numc As String) As Date
   Dim rst As Recordset
   Set rst = dbplanificacio.OpenRecordset("select data1 from planificaciototes where comanda=" + atrim(cadbl(numc)))
   If Not rst.EOF Then If Not IsNull(rst!Data1) Then buscardataentrega = Format(rst!Data1, "dd/mm/yy")
   Set rst = Nothing
   If buscardataentrega = "0:00:00" Then buscardataentrega = Empty
End Function
Function buscarcomandes(id_treball As Double, ordremodificacio As Double) As String
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select comanda from comandes where proximaseccio='E' and numtreball=" + atrim(id_treball) + " and numordremodificacio=" + atrim(ordremodificacio))
   While Not rst.EOF
      buscarcomandes = buscarcomandes + IIf(buscarcomandes = "", "", ", ") + atrim(rst!comanda)
      rst.MoveNext
   Wend
   If buscarcomandes = "" Then buscarcomandes = "-"
   Set rst = Nothing
End Function
Function nomfotogravador(codi As Long) As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select nomfotogravador from fotogravadors where codi=" + atrim(codi))
   If Not rst.EOF Then nomfotogravador = rst!nomfotogravador
End Function
Sub crearfitxertemp()
     
    If Not existeix(fitxertmpestats) Then
       crearfitxertemporal
    End If
   Set dbconsulta = DBEngine.OpenDatabase(fitxertmpestats)
   carregarllistadecampstemporals
   creartaula
   
End Sub
Sub crearfitxertemporal()
    borrartemps
    If Not existeix(fitxertmpestats) Then
       DBEngine.CreateDatabase fitxertmpestats, dbLangGeneral
    End If
    
    
End Sub
Sub borrartemps()
   On Error Resume Next
    Kill fitxertmpestats
End Sub
Sub carregarllistadecampstemporals()
  Dim i As Byte
  i = 1
  campsestat(i, 1) = "Treball": campsestat(i, 2) = "string": campsestat(i, 3) = "Treball/v": i = i + 1
  campsestat(i, 1) = "dataobertura": campsestat(i, 2) = "date": campsestat(i, 3) = "Data Ob:": i = i + 1
  campsestat(i, 1) = "comandes": campsestat(i, 2) = "string": campsestat(i, 3) = "Comandes": i = i + 1
  campsestat(i, 1) = "codibarres": campsestat(i, 2) = "string": campsestat(i, 3) = "Codi Barres": i = i + 1
  campsestat(i, 1) = "fotogravador": campsestat(i, 2) = "string": campsestat(i, 3) = "Nom Fotograbador": i = i + 1
  campsestat(i, 1) = "client": campsestat(i, 2) = "string": campsestat(i, 3) = "Nom Client": i = i + 1
  campsestat(i, 1) = "marcalinia": campsestat(i, 2) = "string": campsestat(i, 3) = "Marca i Linia": i = i + 1
  campsestat(i, 1) = "pdf": campsestat(i, 2) = "string": campsestat(i, 3) = "Pdf": i = i + 1
  campsestat(i, 1) = "estatclixe": campsestat(i, 2) = "string": campsestat(i, 3) = "Estat del Clixé": i = i + 1
  campsestat(i, 1) = "dataentrega": campsestat(i, 2) = "date": campsestat(i, 3) = "Data planificació": i = i + 1
  campsestat(i, 1) = "dataprevista": campsestat(i, 2) = "date": campsestat(i, 3) = "Data fotogravador": i = i + 1
  campsestat(i, 1) = "observacions": campsestat(i, 2) = "string": campsestat(i, 3) = "Observacions": i = i + 1
  campsestat(i, 1) = "": campsestat(i, 2) = "": campsestat(i, 3) = "": i = i + 1
  
End Sub
Sub creartaula()
  Dim i As Integer
 On Error GoTo jaexisteix
  dbconsulta.Execute ("create table consultaestats (id counter)")
  On Error GoTo 0
  dbconsulta.Execute "CREATE INDEX principal ON consultaestats ([id]) witH PRIMARY;"


  For i = 1 To 100
    If campsestat(i, 1) <> "" Then
       dbconsulta.Execute ("alter table consultaestats add column " + campsestat(i, 1) + " " + campsestat(i, 2))
        Else: i = 1000
    End If
  Next i
  
  SetAllowZeroLength dbconsulta
  Exit Sub
jaexisteix:
  dbconsulta.Execute "drop table consultaestats"
  Resume
End Sub

Function SetAllowZeroLength(db As Database)
    Dim i As Integer, j As Integer
    Dim td As TableDef, fld As Field

    
    'The following line prevents the code from stopping if you do not
    'have permissions to modify particular tables, such as system
    'tables.
    On Error Resume Next
    For i = 0 To db.TableDefs.Count - 1
       Set td = db(i)
       For j = 0 To td.Fields.Count - 1
          Set fld = td(j)
          If (fld.Type = 10) And Not _
            fld.AllowZeroLength Then
             fld.AllowZeroLength = True
          End If
       Next j
    Next i
    
End Function

Sub guardar_amples_reixa()
Dim j As Integer
If iniconfigreixa <> "" Then
  For j = 0 To reixa.Cols - 1
   escriure_ini "AmplesReixa", UCase(reixa.TextMatrix(0, j)), atrim(reixa.ColWidth(j)), iniconfigreixa
 Next j
End If
End Sub

Private Sub Form_Resize()
   If Formconsultaestats.Height - reixa.Top - 800 < 1 Then Exit Sub
   reixa.width = Formconsultaestats.width - 300
   reixa.Height = Formconsultaestats.Height - reixa.Top - 800
   Fbotons.Left = Formconsultaestats.width - Fbotons.width - 300
   etregistres.Top = reixa.Height + reixa.Top
   If Formconsultaestats.tag <> "canvianttamany" Then
    escriure_ini "TamanyForm", "ample", atrim(Formconsultaestats.width), iniconfigreixa
    escriure_ini "TamanyForm", "alt", atrim(Formconsultaestats.Height), iniconfigreixa
   End If
End Sub
Sub carregartamanyform()
  If cadbl(llegir_ini("TamanyForm", "ample", iniconfigreixa)) > 0 Then
   Formconsultaestats.tag = "canvianttamany"
   Formconsultaestats.width = llegir_ini("TamanyForm", "ample", iniconfigreixa)
   Formconsultaestats.Height = llegir_ini("TamanyForm", "alt", iniconfigreixa)
   Formconsultaestats.tag = ""
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
guardar_amples_reixa
End Sub

Private Sub MouseWheel1_WheelMove(bDown As Boolean)
  Dim v As Byte
  v = 3
  If reixa.Rows < 2 Then Exit Sub
  If bDown Then
     If reixa.TopRow + v < reixa.Rows Then
        reixa.TopRow = reixa.TopRow + v
       Else: reixa.TopRow = reixa.Rows - 1
     End If
    Else:
        If reixa.TopRow - v > 1 Then
           reixa.TopRow = reixa.TopRow - v
          Else: reixa.TopRow = 1
        End If
  End If
  
End Sub

Private Sub reixa_Click()
  If bordre.BackColor = QBColor(14) Then
      vordre = campsestat(reixa.col + 1, 1)
      If vordre = "dataobertura" Then vordre = "cvdate(dataobertura)"
      If InStr(1, bordre.tag, vordre) > 0 Then
          If InStr(1, bordre.tag, "ASC") > 0 Then
                bordre.tag = " DESC"
              Else: bordre.tag = " ASC"
          End If
           Else
              bordre.tag = " ASC"
      End If
      etordre = campsestat(reixa.col + 1, 3) + " " + bordre.tag
      bordre.tag = vordre + bordre.tag
      etmsgajuda.visible = False
      bordre.BackColor = treurefiltre.BackColor
      reixa.BackColorFixed = treurefiltre.BackColor
      poblarlareixa whereultimfiltre
  End If
End Sub

Private Sub reixa_DblClick()
   If campsestat(reixa.col + 1, 1) = "observacions" Then
     resp = InputBox("Entra la observació que vulguis apuntar." + Chr(10) + "Si vols borrar-la fes un espai en blanc", "Entrada observació", reixa.Text)
     If resp <> "" Then
        reixa.Text = treure_apostruf(resp)
        guardarobservacio reixa.TextMatrix(reixa.row, 0), reixa.Text
        dbconsulta.Execute "update consultaestats set observacions='" + reixa.Text + "' where treball='" + reixa.TextMatrix(reixa.row, 0) + "'"
     End If
   End If
   If campsestat(reixa.col + 1, 1) = "Treball" Then
    If InStr(1, reixa.Text, "/") > 0 Then
        FormClixes.clixes.Recordset.FindFirst "id_treball=" + atrim(cadbl(Mid(reixa.Text, 1, InStr(1, reixa.Text, "/") - 1)))
        FormClixes.SetFocus
     End If
   End If
End Sub
Sub guardarobservacio(treball As String, valornou As String)
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from consultaestats where treball='" + atrim(treball) + "'")
   If Not rst.EOF Then
       If valornou = "" Then
          rst.Delete
          GoTo fi
           Else: rst.Edit
       End If
         Else
           rst.AddNew
           rst!treball = treball
   End If
   rst!observacions = valornou
   rst.Update
fi:
   Set rst = Nothing
End Sub

Private Sub reixa_LostFocus()
    guardar_amples_reixa
End Sub
Sub borrarelfiltre()
   configreixa
   poblarlareixa
   filtre_LostFocus 998
End Sub

Private Sub sortir_Click()
  Unload Formconsultaestats
End Sub

Private Sub treurefiltre_Click()
 borrarelfiltre
End Sub
