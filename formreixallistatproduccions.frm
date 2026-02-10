VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formreixallistatproduccions 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reixa consulta produccions"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15270
   ControlBox      =   0   'False
   Icon            =   "formreixallistatproduccions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox cultimaref 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Només ultima Referència"
      Height          =   240
      Left            =   2070
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.CommandButton bposicioordre 
      Height          =   315
      Left            =   1695
      Picture         =   "formreixallistatproduccions.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar totes les linies"
      Top             =   1020
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton bordre 
      Height          =   315
      Left            =   0
      Picture         =   "formreixallistatproduccions.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Eliminar totes les linies"
      Top             =   15
      Width           =   300
   End
   Begin VB.Frame Fbotons 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   13830
      TabIndex        =   4
      Top             =   -15
      Width           =   1410
      Begin VB.CommandButton sortir 
         Height          =   480
         Left            =   690
         Picture         =   "formreixallistatproduccions.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Sortir"
         Top             =   120
         Width           =   645
      End
      Begin VB.CommandButton exportaraxls 
         BackColor       =   &H00F0F0F0&
         Height          =   480
         Left            =   45
         Picture         =   "formreixallistatproduccions.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exportar a Excel la sel.lecció"
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Height          =   480
         Left            =   1845
         Picture         =   "formreixallistatproduccions.frx":1F46
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir sel.lecció"
         Top             =   120
         Visible         =   0   'False
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
      Picture         =   "formreixallistatproduccions.frx":24D0
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
      Height          =   270
      Left            =   75
      TabIndex        =   10
      Top             =   7650
      Width           =   15150
   End
   Begin VB.Label etordre 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   345
      TabIndex        =   9
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
Attribute VB_Name = "Formreixallistatproduccions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim fitxertmpestats As String
Dim dbplanificacio As Database
Dim rsttemporal As Recordset
Dim whereultimfiltre As String



Sub configreixa(Optional nocarregaramples As Boolean)
  Dim rst As Recordset
  Dim col As Integer
  Dim enes As Byte
  'reixa.LeftCol = 0
  If reixa.Rows > 1 Then reixa.TopRow = 1
  Set rst = dbconsulta.OpenRecordset("select * from llistatprodu")
  col = 0
  enes = 0
  reixa.Cols = rst.Fields.Count + 1
  For i = 0 To rst.Fields.Count - 1
       reixa.ColAlignment(col) = 2
       reixa.TextMatrix(0, col) = UCase(rst.Fields(i).Name) 'campsestat(i + 1, 3)
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
    If X < reixa.Width Then
     filtre(j).Left = X
     filtre(j).Width = cadbl(ample)
     filtre(j).Visible = True
     filtre(j).ForeColor = &H808080
      Else: If filtre.Count < j - 1 Then filtre(j).Visible = False
    End If
    X = X + cadbl(ample)
   End If
 Next j
End If

End Sub
Sub carregarnomvectordecamps(rst As Recordset)
   Dim j As Integer
   For j = 0 To rst.Fields.Count - 1
      campsestat(j, 1) = rst.Fields(i).Name: campsestat(j, 2) = tipusdecamp(rst.Fields(j).Type): campsestat(j, 3) = UCase(rst.Fields(j).Name)
   Next j
End Sub
Function tipusdecamp(v As Double) As String
   If v = dbDate Then tipusdecamp = "date"
   If v = dbText Then tipusdecamp = "string"
   If v = dbByte Then tipusdecamp = "byte"
   If v = dbDouble Then tipusdecamp = "double"
   If v = dbLong Then tipusdecamp = "long"
End Function
Function ordredelataula() As String
  If bordre.Tag = "" Then
     ordredelataula = " order by maq,dia"
    Else: ordredelataula = " order by " + bordre.Tag
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
  Dim vordre As String
  Dim vultimacodi As String
  ratoli "espera"
  etregistres = ""
  reixa.Visible = False
  reixa.Clear
  reixa.BackColor = QBColor(15)
  reixa.Rows = 1
  If cultimaref.Value = 1 Then
       vordre = "order by maq,dia"
        Else: vordre = ordredelataula
  End If
  Set rst = dbconsulta.OpenRecordset("select * from llistatprodu " + IIf(were <> "", " where " + were, "") + vordre)
  If rst.EOF Then GoTo fi
  Set rsttemporal = rst.Clone
  carregarnomvectordecamps rsttemporal
  configreixa IIf(were <> "", True, False)
  fila = 0
  reixa.Tag = "poblant"
  While Not rst.EOF
   If cultimaref.Value = 1 Then If rst!comanda = vultimcodi Then GoTo proxim
   fila = fila + 1
   reixa.Rows = fila + 1
   For i = 0 To rst.Fields.Count - 1
     If rst.Fields(i) <> "" Then
      reixa.TextMatrix(fila, i) = IIf(IsNull(rst.Fields(i)), "", rst.Fields(i))
      If reixa.TextMatrix(fila, i) = "0:00:00" Then reixa.TextMatrix(fila, i) = ""
      posarelcolordelcamp fila, i, rst.Fields(i)
     End If
   Next i
proxim:
   vultimcodi = rst!comanda
   rst.MoveNext
  Wend
  etregistres.Caption = "Registres: " + atrim(rst.RecordCount) + IIf(cultimaref.Value = 1, " -> Filtrats: " + atrim(reixa.Rows - 1), "")
fi:
  Set rst = Nothing
  reixa.Tag = ""
  reixa.Visible = True
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
  filtre(col).Text = UCase(rsttemporal.Fields(col).Name)
  filtre(col).Tag = i
'  Load filtre(col + 1)
End Sub


Private Sub bordre_Click()
 etmsgajuda = "Prem sobre la columna que vols ordenar."
 etmsgajuda.Width = 3000
 etmsgajuda.Left = treurefiltre.Left + treurefiltre.Width + 100
 etmsgajuda.Visible = True
 bordre.BackColor = QBColor(14)
 reixa.BackColorFixed = QBColor(14)
End Sub

Private Sub Command3_Click()
   imprimirseleccio False
End Sub
Sub imprimirseleccio(vexportar As Boolean)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.report
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

Private Sub cultimaref_Click()
   borrarelfiltre
   poblarlareixa
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
Function esalareixa(vnumc As Double) As Boolean
   Dim vtrobat As Boolean
   Dim vcont As Double
   While Not vtrobat And vcont < reixa.Rows
      If cadbl(reixa.TextMatrix(vcont, 4)) = vnumc Then vtrobat = True
      vcont = vcont + 1
   Wend
   esalareixa = vtrobat
End Function
Sub generar_xls()
   Dim i As Byte
   Dim rst As Recordset
   Dim linia As String
   Dim vprimerclient As Boolean
   If subbusqueda.checkunperun.Value = 1 Then vprimerclient = True
   
   Set rst = dbconsulta.OpenRecordset("select * from consultaestats " + IIf(whereultimfiltre <> "", " where " + whereultimfiltre, "") + ordredelataula)
   If rst.EOF Then MsgBox "No hi ha dades per exportar", vbCritical, "Error": Exit Sub
   Open "c:\temp\consultarefinplacsa.csv" For Output As #1
   If Not rst.EOF Then
    If vprimerclient Then linia = "Client;DireccioEnviament"
    For i = 0 To rst.Fields.Count - 1 ' IIf(vprimerclient, 3, 1)
       If Not (vprimerclient And rst.Fields(i).Name = "direnvio" And rst.Fields(i).Name = "nomclient") Then
         linia = linia + IIf(linia = "", "", ";") + atrim(campsestat(i + 1, 3))
       End If
    Next i
    Print #1, linia
   End If
   While Not rst.EOF
    linia = ""
   ' If (seleccionats("fi") - seleccionats("inici")) > 0 Then
   '    If Not tocaexportar(rst.Fields("numtreball"), seleccionats("inici"), seleccionats("fi")) Then GoTo proxim
   ' End If
   ' If Not esalareixa(cadbl(rst!numcomanda)) Then GoTo proxim
   ' If vprimerclient Then linia = atrim(rst!nomclient) + ";" + atrim(rst!direnvio)
    For i = 1 To rst.Fields.Count - 1 '- IIf(vprimerclient, 3, 1)
      If Not (vprimerclient And rst.Fields(i).Name = "direnvio" And rst.Fields(i).Name = "nomclient") Then
        linia = linia + IIf(linia = "", "", ";") + """" + IIf(rst.Fields(i).Name = "codibarres", "Nº: ", "") + atrim(rst.Fields(i)) + """"
      End If
    Next i
    Print #1, linia
proxim:
    rst.MoveNext
   Wend
   Close #1
   wait 2
   obrir_document "c:\temp\consultarefinplacsa.csv"
      
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
  If filtre(Index).Width < 500 Then filtre(Index).HelpContextID = filtre(Index).Width: filtre(Index).Width = 1000
End Sub
Sub bxrcontrolagafafocus(i As Integer)
  Dim cntrl As Control
  Set cntrl = Screen.ActiveControl
  If cntrl.Text <> "" Then
     If cntrl.Text = campsestat(cadbl(filtre(i).Tag), 3) Then cntrl.Text = ""
     cntrl.ForeColor = QBColor(0)
     
   Else:
       
       cntrl.Text = campsestat(cadbl(filtre(i).Tag), 3)
       cntrl.ForeColor = &H808080
  End If
End Sub


Private Sub filtre_LostFocus(Index As Integer)
  Dim noufiltre As String
  If Index = 998 Then whereultimfiltre = "": Exit Sub
  noufiltre = crearfiltre
  If filtre(ultimfiltre).Text = "" Then
    filtre(ultimfiltre).Text = campsestat(cadbl(filtre(ultimfiltre).Tag), 3)
    filtre(ultimfiltre).ForeColor = &H808080
    If filtre(ultimfiltre).HelpContextID <> 0 Then filtre(ultimfiltre).Width = filtre(ultimfiltre).HelpContextID
  End If
  If noufiltre <> whereultimfiltre Or Index = 999 Then
     If noufiltre <> "" Then poblarlareixa noufiltre
  End If
  If Index = 999 And noufiltre = "" Then
     poblarlareixa
  End If
  ratoli "normal"
  reixa.Visible = True
  whereultimfiltre = noufiltre
  'Me.caption = whereultimfiltre
  possaretiquetaajuda
  'Command3.tag = noufiltre ' el guardo pel llistat
  
End Sub
Sub possaretiquetaajuda()
   Dim i As Byte
   etmsgajuda.Visible = False
   For i = 0 To filtre.Count - 1
    If InStr(1, filtre(i), ",") > 0 Then
      etmsgajuda.Caption = "Una coma busca dos valors"
      etmsgajuda.Width = filtre(i).Width
      If etmsgajuda.Width < 2000 Then etmsgajuda.Width = 2000
      etmsgajuda.Left = filtre(i).Left
      etmsgajuda.Visible = True
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
  While (InStr(1, filtre, ",") > 0) And filtre <> ""
    'si hi ha coma
    If InStr(1, filtre, ",") > 0 Then
        If camp <> "codiclient" Then
           If camp = "estatclixe" Then
             re = IIf(re <> "", re + " or ", "") + camp + " = '" + Mid(filtre, 1, InStr(1, filtre, ",") - 1) + "'"
               Else: re = IIf(re <> "", re + " or ", "") + camp + " like '*" + Mid(filtre, 1, InStr(1, filtre, ",") - 1) + "*'"
           End If
          Else: re = IIf(re <> "", re + " or ", "") + camp + " =" + atrim(cadbl(Mid(filtre, 1, InStr(1, filtre, ",") - 1))) + ""
        End If
        filtre = Mid(filtre, InStr(1, filtre, ",") + 1)
        GoTo proxima
End If
    
    'si hi ha punticoma
   ' If InStr(1, Mid(filtre, 1, Len(filtre) - 1), ";") > 0 Then
   '     If camp <> "codiclient" Then
   '        If camp = "estatclixe" Then
   '          re = IIf(re <> "", re + " and ", "") + camp + " = '" + Mid(filtre, 1, InStr(1, filtre, ";") - 1) + "'"
   '            Else: re = IIf(re <> "", re + " and ", "") + camp + " like '*" + Mid(filtre, 1, InStr(1, filtre, ";") - 1) + "*'"
    '       End If
    '      Else: re = IIf(re <> "", re + " and ", "") + camp + " =" + atrim(cadbl(Mid(filtre, 1, InStr(1, filtre, ";") - 1))) + ""
    '    End If
    '    filtre = Mid(filtre, InStr(1, filtre, ";") + 1)
    '    GoTo proxima
    'End If
    
proxima:
    
  Wend
  If re <> "" Then re = "(" + re + ")"
  possarweres = re
End Function
Function crearwere(i As Integer) As String
   Dim w As String
   Dim j As Integer
   If filtre(i) = "" Then Exit Function
   j = cadbl(filtre(i).Tag)
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
    If filtre(i).Text <> campsestat(cadbl(filtre(i).Tag), 3) And campsestat(cadbl(filtre(i).Tag), 1) <> "comanda" Then
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
   Dim rst As Recordset '   + atrim(campsestat(Index + 1, 1))    atrim(campsestat(Index + 1, 1)) +
   Set rst = dbconsulta.OpenRecordset("select distinct " + atrim(campsestat(Index + 1, 1)) + " as valor from llistatprodu " + IIf(whereultimfiltre <> "", " where " + whereultimfiltre, "") + " order by " + atrim(campsestat(Index + 1, 1)) + " asc")
   filtre(Index).Clear
   While Not rst.EOF
      If atrim(rst!valor) <> "" Then filtre(Index).AddItem rst!valor
      rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Private Sub Form_Load()
   Dim vnopoblar As Boolean
   If subbusqueda.etestatusllistat.Tag = "-" Then Exit Sub
   If r = "nopoblar" Then vnopoblar = True: r = ""
   iniconfigreixa = "c:\windows\reixaproduccionsinplacsa.ini"
   fitxertmpestats = "c:\temporal.mdb"
   Me.Caption = "Consulta referencies d'Inplacsa...(Carregant)"
   DoEvents
   If subbusqueda.etestatusllistat.Tag = "parant" Then GoTo fi
   carregartamanyform
   Set dbconsulta = DBEngine.OpenDatabase(fitxertmpestats)
'   crearfitxertemp
  ' carregardadesfitxertemporal
   If Not vnopoblar Then poblarlareixa
   Me.Caption = "Produccions inplacsa"
   If subbusqueda.etestatusllistat.Tag = "parant" Then GoTo fi
   Exit Sub
fi:
   Unload Me
End Sub
Sub borrarlataula()
   dbconsulta.Execute "delete * from consultaestats"
End Sub
Sub carregardadesfitxertemporal()
   Dim rst As Recordset
   Dim rstnou As Recordset
   Dim dblocal As Database
   borrarlataula
   ratoli "espera"
   Workspaces(0).BeginTrans
   Set dblocal = OpenDatabase(cami)
   Set dblocal = dbtmp
   Set rst = rstconsulta
   Set rstnou = dbconsulta.OpenRecordset("select * from consultaestats")
   If rst.EOF Then MsgBox "No hi ha dades.": Exit Sub
   rst.MoveLast
   rst.MoveFirst
   While Not rst.EOF
     subbusqueda.etestatusllistat = "Carregant la reixa...  " + atrim(rst.AbsolutePosition) + "/" + atrim(rst.RecordCount): DoEvents
     copiarregistreatemporal rst, rstnou, dblocal
    rst.MoveNext
    If subbusqueda.etestatusllistat.Tag = "parant" Then GoTo fi
   Wend
  ratoli "normal"
  Workspaces(0).CommitTrans
fi:
   Set rst = Nothing
   Set rstnou = Nothing
   Set dbplanificacio = Nothing
   Set dblocal = Nothing
End Sub
Function canviarlacoma(ByVal n As String) As String
   While InStr(n, ",")
     n = Mid(n, 1, InStr(1, n, ",") - 1) + "¸" + Mid(n, InStr(1, n, ",") + 1)
   Wend
   If n = "{[}]" Then n = ""
   canviarlacoma = n
End Function
Function buscarnomdelclient(vcodiclient As Integer) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(vcodiclient), dbOpenSnapshot, dbReadOnly)
  If Not rst.EOF Then
     buscarnomdelclient = atrim(rst!nom)
  End If
End Function
Function buscarnomclientfact(vnumc As Double) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select codicomptable from comandes_extres where comanda=" + atrim(vnumc), dbOpenSnapshot, dbReadOnly)
  If rst.EOF Then Exit Function
  Set rst = dbtmp.OpenRecordset("select nomclient from clients_codissap where codiSAP=" + atrim(cadbl(rst!codicomptable)), dbOpenSnapshot, dbReadOnly)
  If Not rst.EOF Then
     buscarnomclientfact = atrim(rst!nomclient)
  End If
End Function
Function buscarnomdirenvio(vdirenvio As Integer) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from clients_envios where id=" + atrim(vdirenvio), dbOpenSnapshot, dbReadOnly)
  If Not rst.EOF Then
     buscarnomdirenvio = atrim(rst!nome) + "(" + atrim(rst!poblacioe) + ")"
  End If
End Function
Function buscarlacomandacorrecte(vnumc As Double) As Double
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select refinplacsa from comandes_extres where comanda=" + atrim(vnumc) + "", dbOpenSnapshot, dbReadOnly)
   
   If Not rst.EOF Then Set rst = dbtmp.OpenRecordset("SELECT comandes.comanda, comandes.datacomanda,comandes_extres.refinplacsa, comandes.numordremodificacio FROM comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda WHERE (((comandes_extres.refinplacsa)='" + atrim(rst!refinplacsa) + "')AND ((comandes.producte)<>'PC' And (comandes.producte)<>'PC2' And (comandes.producte)<>'PCP')) ORDER BY comandes.datacomanda DESC , comandes.numordremodificacio DESC;", dbOpenSnapshot, dbReadOnly)
   If Not rst.EOF Then buscarlacomandacorrecte = rst!comanda
   Set rst = Nothing
End Function
Function referenciavella(vrefcli As String, vrefclialt As String) As String
   Dim vref As String
   Dim vvref As String
   Dim i As Byte
   vref = vrefcli + " | " + vrefclialt + " | "
   i = 1
   vvref = ""
   While atrim(Mid(vref, i, InStr(i, vref, "|"))) <> ""
      vvref = atrim(Mid(vref, i, (InStr(i, vref, "|") - (i))))
      i = InStr(i, vref, "|") + 1
      If Mid(vvref, 1, 1) = "0" Then referenciavella = vvref
   Wend
End Function
Function referenciasap(vrefcli As String, vrefclialt As String) As String
   Dim vref As String
   Dim vvref As String
   Dim i As Byte
   While InStr(1, vrefclialt, "/")
     vrefclialt = substituir(vrefclialt, "/", "|")
   Wend
    
   vref = vrefcli + " | " + vrefclialt + " | "
   i = 1
   vvref = ""
   While atrim(Mid(vref, i, InStr(i, vref, "|"))) <> ""
      vvref = atrim(Mid(vref, i, (InStr(i, vref, "|") - (i))))
      i = InStr(i, vref, "|") + 1
      If Mid(vvref, 1, 1) <> "0" And atrim(Mid(vvref, 1, 1)) <> "" And Len(vvref) < 7 Then referenciasap = vvref
   Wend
End Function

Sub copiarregistreatemporal(rst As Recordset, rstnou As Recordset, dblocal As Database)
   Dim rstc As Recordset
   Dim rstc2 As Recordset
   Dim rstcextres As Recordset
   Dim vpe As Double
   Dim vme As Double
   Dim vke As Double
   Dim vsql As String
   Dim vnumcomanda As Double
  
   If rst!q > 1 Then
      vnumcomanda = buscarlacomandacorrecte(rst!maxcomanda)  'haig de buscar la comanda que tingui la versio de treball mes alta
       Else: vnumcomanda = rst!maxcomanda
   End If
   Set rstc = dblocal.OpenRecordset("SELECT comandes.*, productes.ruta as laruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi where comanda = " + atrim(vnumcomanda), dbOpenSnapshot, dbReadOnly)
   If rstc.EOF Then Exit Sub
   Set rstcextres = dblocal.OpenRecordset("select * from comandes_extres where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
   If rstcextres.EOF Then Exit Sub
   If cadbl(rst!maxcomanda) = 0 Then Exit Sub
   ' refinplacsa, first(producte) as Pr,first(refclient) as Ref_, count(*) as Q,Max(datacomanda) AS maxdata, Max(comanda) AS maxcomanda
   rstnou.AddNew
   rstnou!nomclient = atrim(rstc!client) + " - " + buscarnomdelclient(rstc!client)
   rstnou!direnvio = buscarnomdirenvio(cadbl(rstc!direnvio))
   rstnou!datacomanda = rstc!datacomanda
   rstnou!numcomandes = cadbl(rst!q)
   rstnou!refinplacsa = atrim(rst!refinplacsa)
   If rstnou!refinplacsa = "" Then
       rstnou!datacomanda = Now
       rstnou!refinplacsa = "Sense Referència"
       GoTo cont
        Else
          If cadbl(Mid(rstnou!refinplacsa, 1, 2)) > 0 Then
                rstnou!vref1 = cadbl(Mid(rstnou!refinplacsa, 1, 2))
                rstnou!vref2 = Mid(rstnou!refinplacsa, 3)
              Else
                rstnou!vref1 = 1
                rstnou!vref2 = atrim(rstnou!refinplacsa)
          End If
   End If
   If InStr(1, rstnou!nomclient, "CROP´S") > 0 Then
      rstnou!refclient = referenciasap(atrim(rstc!refclient), atrim(rstc!refclialt))
      rstnou!refclientvella = referenciavella(atrim(rstc!refclient), atrim(rstc!refclialt))
        Else: rstnou!refclient = atrim(rstc!refclient)
   End If
   rstnou!stopped = mirarsiestaanuladalareferencia(atrim(rstnou!refclient), atrim(rstnou!refclientvella), rstc!client)
   rstnou!numcomanda = cadbl(rstc!comanda)
   rstnou!producte = atrim(rstc!producte)
   rstnou!peskg = cadbl(rstc!rebkilos)
   rstnou!ampleext = atrim(cadbl(rstc!ampleesq) * 10) + IIf(cadbl(rstc!plegatesq) > 0, "/" + atrim(cadbl(rstc!plegatesq) * 10), "")
   If InStr(1, rstc!laruta, "I") Then
     passardadesdeltreball rstnou, cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio)
   End If
   If InStr(1, rstc!laruta, "R") Then
      rstnou!amplereb = (cadbl(rstc!amplereb) * 10)
      rstnou!simulteneitatreb = cadbl(rstc!simulteneitatreb)
   End If
   If InStr(1, rstc!laruta, "S") Then
      rstnou!amplesol = atrim(cadbl(rstc!amplesol) * 10) + IIf(cadbl(rstc!ampleplegsol) > 0, "/" + atrim(cadbl(rstc!ampleplegsol) * 10), "")
      rstnou!longitud = atrim(cadbl(rstc!longitudsol) * 10) + IIf(cadbl(rstc!fuellebasesol) > 0, "/" + atrim(cadbl(rstc!fuellebasesol) * 10), "")
      rstnou!solapa = cadbl(rstc!solapasol) * 10
      rstnou!tipussoldadura = atrim(rstc!tipusoldadura)
   End If
   rstnou!clientfacturacio = buscarnomclientfact(rstc!comanda)
   rstnou!quantitatteorica = cadbl(rstc!tubbaseext)
   rstnou!unitatteorica = buscarmesuraunitat(cadbl(rstc!mesuraquantdemanada))
   rstnou!numtreball = cadbl(rstc!numtreball)
   If rst!q = 1 Then
    rstnou!metresentregats = cadbl(rstcextres!metresentregats) 'calcularmetresentregats(rstc!comanda)
    rstnou!numpecesentregades = cadbl(rstcextres!numpecesentregades) 'calcularpecesentregades(rstc!comanda, rstnou)
    rstnou!kilosentregats = cadbl(rstcextres!kilosentregats) 'calcularkilosentregats(rstc!comanda, calcularpesxrpeça_consulta(rstc), rstnou)
    rstnou!pvp_unitat = cadbl(rstc!pvp)
    rstnou!unitatpvp = buscarmesurapvp(cadbl(rstc!mesurapvp))
    rstnou!pvp = cadbl(rstcextres!pvptotal) ' buscartotalcomanda(rstc!comanda, rstnou)
      Else
        Set rstc2 = dblocal.OpenRecordset("SELECT comandesmesextres.*, productes.ruta as laruta FROM comandesmesextres INNER JOIN productes ON comandesmesextres.producte = productes.codi " + subbusqueda.Tag + " and refinplacsa='" + atrim(rst!refinplacsa) + "'")
        vme = 0: vpe = 0: vke = 0: vpvp = 0
        While Not rstc2.EOF
          '  MsgBox rstc2!comanda
            rstnou!metresentregats = cadbl(rstc2!metresentregats) 'calcularmetresentregats(rstc2!comanda)
            rstnou!numpecesentregades = cadbl(rstc2!numpecesentregades) 'calcularpecesentregades(rstc2!comanda, rstnou)
            rstnou!kilosentregats = cadbl(rstc2!kilosentregats) 'calcularkilosentregats(rstc2!comanda, calcularpesxrpeça_consulta(rstc2), rstnou)
            rstnou!pvp_unitat = 0
            rstnou!unitatpvp = buscarmesurapvp(cadbl(rstc2!mesurapvp))
            rstnou!pvp = cadbl(rstc2!pvptotal) 'buscartotalcomanda(rstc2!comanda, rstnou)
            rstnou!unitatpvp = ""
            rstc2.MoveNext
            vme = vme + Redondejar(cadbl(rstnou!metresentregats), 0)
            vpe = vpe + Redondejar(cadbl(rstnou!numpecesentregades), 0)
            vke = vke + Redondejar(cadbl(rstnou!kilosentregats), 0)
            vpvp = vpvp + Redondejar(cadbl(rstnou!pvp), 2)
        Wend
        rstnou!metresentregats = vme
        rstnou!numpecesentregades = vpe
        rstnou!kilosentregats = vke
        rstnou!pvp = vpvp
   End If

  ' vsql = "update comandes_extres set metresentregats=" + atrim(Redondejar(cadbl(rstnou!metresentregats), 0))
  ' vsql = vsql + ", numpecesentregades=" + passaradecimalpunt(Redondejar(cadbl(rstnou!numpecesentregades), 0))
  ' vsql = vsql + ", kilosentregats=" + passaradecimalpunt(cadbl(rstnou!kilosentregats))
  ' vsql = vsql + ",pvptotal=" + passaradecimalpunt(cadbl(rstnou!pvp))
  ' vsql = vsql + " where comanda=" + atrim(rstc!comanda)
  ' dbtmp.Execute vsql
   possarespesorimaterial rstnou, rstc!comanda, rstc!linkcomanda1, rstc!linkcomanda2
cont:
   rstnou.Update
End Sub
Function mirarsiestaanuladalareferencia(vrefnova As String, vrefvella As String, vnumclient As Double) As String
   Dim rst As Recordset
   mirarsiestaanuladalareferencia = False
   Set rst = dbtmp.OpenRecordset("select * from refclient_stopped where numclient=" + atrim(vnumclient) + " and (refclient='" + atrim(vrefnova) + "' or refclient='" + atrim(vrefvella) + "')", dbOpenSnapshot, dbReadOnly)
   If Not rst.EOF Then mirarsiestaanuladalareferencia = True
End Function
Sub vella_copiarregistreatemporal(rst As Recordset, rstnou As Recordset)
   Dim rstc As Recordset
   Dim rstc2 As Recordset
   Dim vpe As Double
   Dim vme As Double
   Dim vke As Double
   Dim vsql As String
   Set rstc = dbtmp.OpenRecordset("SELECT comandesmesextres.*, productes.ruta as laruta FROM comandesmesextres INNER JOIN productes ON comandesmesextres.producte = productes.codi where comanda = " + atrim(rst!maxcomanda), dbOpenSnapshot, dbReadOnly)
   If rstc.EOF Then Exit Sub
   If cadbl(rst!maxcomanda) = 0 Then Exit Sub
   If rstc!producte = "PC" Or rstc!producte = "PC2" Or rstc!producte = "PCP" Or rstc!producte = "PC3" Or rstc!pvptotal > 0 Or (rstc!proximaseccio <> "T" And rstc!proximaseccio <> "P") Then Exit Sub
   ' refinplacsa, first(producte) as Pr,first(refclient) as Ref_, count(*) as Q,Max(datacomanda) AS maxdata, Max(comanda) AS maxcomanda
   rstnou.AddNew
   rstnou!nomclient = atrim(rstc!client) + " - " + buscarnomdelclient(rstc!client)
   rstnou!direnvio = buscarnomdirenvio(cadbl(rstc!direnvio))
   rstnou!datacomanda = rst!maxdata
   rstnou!numcomandes = cadbl(rst!q)
   rstnou!refinplacsa = atrim(rst!refinplacsa)
  ' If rstnou!refinplacsa = "" Then
  '     rstnou!datacomanda = Now
  '     rstnou!refinplacsa = "Sense Referència"
  '     GoTo cont
  '      Else
          If cadbl(Mid(rstnou!refinplacsa, 1, 2)) > 0 Then
                rstnou!vref1 = cadbl(Mid(rstnou!refinplacsa, 1, 2))
                rstnou!vref2 = Mid(rstnou!refinplacsa, 3)
              Else
                rstnou!vref1 = 1
                rstnou!vref2 = atrim(rstnou!refinplacsa)
          End If
   'End If
   rstnou!refclient = atrim(rst!ref_)
   rstnou!numcomanda = cadbl(rst!maxcomanda)
   rstnou!producte = atrim(rst!pr)
   rstnou!peskg = cadbl(rstc!rebkilos)
   rstnou!ampleext = atrim(cadbl(rstc!ampleesq) * 10) + IIf(cadbl(rstc!plegatesq) > 0, "/" + atrim(cadbl(rstc!plegatesq) * 10), "")
   If InStr(1, rstc!laruta, "I") Then
     passardadesdeltreball rstnou, cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio)
   End If
   If InStr(1, rstc!laruta, "R") Then
      rstnou!amplereb = (cadbl(rstc!amplereb) * 10)
      rstnou!simulteneitatreb = cadbl(rstc!simulteneitatreb)
   End If
   If InStr(1, rstc!laruta, "S") Then
      rstnou!amplesol = atrim(cadbl(rstc!amplesol) * 10) + IIf(cadbl(rstc!ampleplegsol) > 0, "/" + atrim(cadbl(rstc!ampleplegsol) * 10), "")
      rstnou!longitud = atrim(cadbl(rstc!longitudsol) * 10) + IIf(cadbl(rstc!fuellebasesol) > 0, "/" + atrim(cadbl(rstc!fuellebasesol) * 10), "")
      rstnou!solapa = cadbl(rstc!solapasol) * 10
      rstnou!tipussoldadura = atrim(rstc!tipusoldadura)
   End If
   rstnou!clientfacturacio = buscarnomclientfact(rstc!comanda)
   rstnou!quantitatteorica = cadbl(rstc!tubbaseext)
   rstnou!unitatteorica = buscarmesuraunitat(cadbl(rstc!mesuraquantdemanada))
   rstnou!numtreball = cadbl(rstc!numtreball)
   If rst!q = 1 Then
    rstnou!metresentregats = calcularmetresentregats(rstc!comanda)
    rstnou!numpecesentregades = calcularpecesentregades(rstc!comanda, rstnou)
    rstnou!kilosentregats = calcularkilosentregats(rstc!comanda, calcularpesxrpeça_consulta(rstc), rstnou)
    rstnou!pvp_unitat = cadbl(rstc!pvp)
    rstnou!unitatpvp = buscarmesurapvp(cadbl(rstc!mesurapvp))
    rstnou!pvp = buscartotalcomanda(rstc!comanda, rstnou)
      Else
        Set rstc2 = dbtmp.OpenRecordset("SELECT comandesmesextres.*, productes.ruta as laruta FROM comandesmesextres INNER JOIN productes ON comandesmesextres.producte = productes.codi " + subbusqueda.Tag + " and refinplacsa='" + atrim(rst!refinplacsa) + "'")
        vme = 0: vpe = 0: vke = 0: vpvp = 0
        While Not rstc2.EOF
            rstnou!metresentregats = calcularmetresentregats(rstc2!comanda)
            rstnou!numpecesentregades = calcularpecesentregades(rstc2!comanda, rstnou)
            rstnou!kilosentregats = calcularkilosentregats(rstc2!comanda, calcularpesxrpeça_consulta(rstc2), rstnou)
            rstnou!pvp_unitat = 0
            rstnou!unitatpvp = buscarmesurapvp(cadbl(rstc2!mesurapvp))
            rstnou!pvp = buscartotalcomanda(rstc2!comanda, rstnou)
            rstnou!unitatpvp = ""
            rstc2.MoveNext
            vme = vme + Redondejar(cadbl(rstnou!metresentregats), 0)
            vpe = vpe + Redondejar(cadbl(rstnou!numpecesentregades), 0)
            vke = vke + Redondejar(cadbl(rstnou!kilosentregats), 0)
            vpvp = vpvp + Redondejar(cadbl(rstnou!pvp), 2)
        Wend
        rstnou!metresentregats = vme
        rstnou!numpecesentregades = vpe
        rstnou!kilosentregats = vke
        rstnou!pvp = vpvp
   End If

   vsql = "update comandes_extres set metresentregats=" + atrim(Redondejar(cadbl(rstnou!metresentregats), 0))
   vsql = vsql + ", numpecesentregades=" + passaradecimalpunt(Redondejar(cadbl(rstnou!numpecesentregades), 0))
   vsql = vsql + ", kilosentregats=" + passaradecimalpunt(cadbl(rstnou!kilosentregats))
   vsql = vsql + ",pvptotal=" + passaradecimalpunt(cadbl(rstnou!pvp))
   vsql = vsql + " where comanda=" + atrim(rstc!comanda)
   'MsgBox vsql
   dbtmp.Execute vsql
   possarespesorimaterial rstnou, rstc!comanda, rstc!linkcomanda1, rstc!linkcomanda2
cont:
   rstnou.Update
End Sub
Function buscartotalcomanda(vnumc As Double, rstnou As Recordset) As Double
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("SELECT comandes.pvp,comandes.comanda,comandes.ampleesq, mesures.unitatinterna, Clients_envios.packinglistalbara, Clients_envios.pesnetbrut,clients_envios.albaraarrodonirkg as arrodonirkg FROM (comandes INNER JOIN Clients_envios ON comandes.direnvio = Clients_envios.id) INNER JOIN mesures ON comandes.mesurapvp = mesures.codi WHERE (((comandes.comanda)=" + atrim(vnumc) + "));")
   If rst.EOF Then Exit Function
   With rstnou
   triarelvalordepenguentdelaunitat = 0
   Select Case rstnou!unitatpvp
     Case "€/1000U"
       buscartotalcomanda = Redondejar(!numpecesentregades / 1000, 3)
     Case "€/U"
       buscartotalcomanda = cadbl(!numpecesentregades)
     Case "€/B"
       'buscartotalcomanda = !numbobs
     Case "€/K"
       If Not rst!pesnetbrut Then
            buscartotalcomanda = Redondejar(!kilosentregats, 1)
             Else: buscartotalcomanda = Redondejar(!kilosentregats, 1)
       End If
     Case "€/M"
       buscartotalcomanda = !metresentregats
     Case "€/KM"
       buscartotalcomanda = Redondejar(!metresentregats / 1000, 2)
     Case "€/FIX"
       buscartotalcomanda = 1
     Case "€/M2"
       buscartotalcomanda = Redondejar(metresentregats * (cadbl(rst!ampleesq) / 1000), 2)
   End Select
   End With
   If rst!unitatinterna = "€/K" And rst!arrodonirkg Then buscartotalcomanda = Redondejar(buscartotalcomanda, 0)
   buscartotalcomanda = Redondejar(buscartotalcomanda * cadbl(rst!pvp), 2)
   Set rst = Nothing
End Function


Function buscarmesuraunitat(vunitat As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from mesureslineals where codi=" + atrim(vunitat))
   If Not rst.EOF Then buscarmesuraunitat = atrim(rst!descripcio)
End Function
Function buscarmesurapvp(vunitatpvp As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from mesures where codi=" + atrim(vunitatpvp))
   If Not rst.EOF Then buscarmesurapvp = atrim(rst!unitatinterna)
End Function
Function calcularpesxrpeça_consulta(rst As Recordset) As Double
    Dim pesgrmcm2 As Double
    Dim rst2 As Recordset
    Set rst2 = dbtmp.OpenRecordset("select solpesgrmcm2 from comandes_extres where comanda=" + atrim(rst!comanda))
    If rst2.EOF Then Exit Function
    If cadbl(rst!cantitatsol) = 0 Then Exit Function
    pesgrmcm2 = cadbl(rst2!solpesgrmcm2)
    calcularpesxrpeça = pesgrmcm2 * ((cadbl(rst!amplesol) + cadbl(rst!solapasol)) * cadbl(rst!longitudsol))
    calcularpesxrpeça = calcularpesxrpeça * IIf(rst!migelaboratsol = "L", 1, 2)
    Set rst2 = Nothing
    'calcularpesxrpeça = cadbl(rst!cantitatsol) * calcularpesxrpeça
End Function
Function calcularkilosentregats(vnumc As Double, vpespeça As Double, rstnou As Recordset) As Double
   Dim rst As Recordset
   'Set rst = dbbaixes.OpenRecordset("SELECT rebobinadores.comanda, bobinesreb.* FROM rebobinadores INNER JOIN bobinesreb ON rebobinadores.Id = bobinesreb.Id where comanda = " + atrim(vnumc))
   'While Not rst.EOF
   '   calcularkilosentregats = calcularkilosentregats + cadbl(rst!kilos)
   '   rst.MoveNext
   'Wend
   Set rst = dbbaixes.OpenRecordset("SELECT tmetres,tkilos from rebobinadorestot where comanda = " + atrim(vnumc))
   If Not rst.EOF Then calcularkilosentregats = cadbl(rst!tkilos)
   If calcularkilosentregats = 0 Then calcularkilosentregats = Redondejar(vpespeça * cadbl(rstnou!numpecesentregades), 2)
   Set rst = Nothing
End Function
Function calcularpecesentregades(vnumc As Double, rstnou As Recordset) As Double
  Dim rst As Recordset
   'Set rst = dbbaixes.OpenRecordset("SELECT soldadores.comanda, bobinessol.* FROM soldadores INNER JOIN bobinessol ON soldadores.Id = bobinessol.Id where comanda = " + atrim(vnumc))
   'While Not rst.EOF
   '   calcularpecesentregades = calcularpecesentregades + cadbl(rst!unitatsxsac)
   '   rst.MoveNext
   'Wend
   Set rst = dbbaixes.OpenRecordset("SELECT tunitats from soldadorestot where comanda = " + atrim(vnumc))
   If Not rst.EOF Then calcularpecesentregades = cadbl(rst!tunitats)
   If calcularpecesentregades = 0 And cadbl(rstnou!desarrollimp) > 0 Then calcularpecesentregades = Redondejar(cadbl(rstnou!metresentregats) / (rstnou!desarrollimp / 1000), 0)
   Set rst = Nothing
End Function
Function calcularmetresentregats(vnumc As Double) As Double
   Dim rst As Recordset
   'Set rst = dbbaixes.OpenRecordset("SELECT rebobinadores.comanda, bobinesreb.* FROM rebobinadores INNER JOIN bobinesreb ON rebobinadores.Id = bobinesreb.Id where comanda = " + atrim(vnumc))
   'While Not rst.EOF
   '   calcularmetresentregats = calcularmetresentregats + cadbl(rst!metres)
   '   rst.MoveNext
   'Wend
   Set rst = dbbaixes.OpenRecordset("SELECT tmetres,tkilos from rebobinadorestot where comanda = " + atrim(vnumc))
   If Not rst.EOF Then calcularmetresentregats = cadbl(rst!tmetres)
   Set rst = Nothing
End Function
Function descripciomaterialconcatenat(rstmat As Recordset) As String
   Dim c As String
   c = atrim(rstmat![familiesmaterials.descripcio])
   c = c + IIf(rstmat![subfamiliesmaterials.descripcio] <> "", "+" + atrim(rstmat![subfamiliesmaterials.descripcio]), "")
   c = c + IIf(atrim(rstmat![familiescolorants.descripcio]) <> "", "+" + atrim(rstmat![familiescolorants.descripcio]), "")
   descripciomaterialconcatenat = c
End Function
Sub possarespesorimaterial(rstnou As Recordset, numc1 As Double, numc2 As Double, numc3 As Double)
    Dim rstmat1 As Recordset
  Dim rstmat2 As Recordset
  Dim rstmat3 As Recordset
  Dim espesormat1 As Double
  Dim espesormat2 As Double
  Dim espesormat3 As Double
  Dim descripciomat As String
  Dim tipusfilm As String
  Dim codimat As String
  Dim rstcomandes As Recordset
  Set rstcomandes = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(numc1) + " or comanda=" + atrim(numc2) + " or comanda=" + atrim(numc3), dbOpenSnapshot, dbReadOnly)
  If rstcomandes.EOF Then Exit Sub
  rstcomandes.FindFirst "comanda=" + atrim(numc1)
  codimat = IIf(Not rstcomandes.NoMatch, cadbl(rstcomandes!materialex), 0)
  Set rstmat1 = dbtmp.OpenRecordset("SELECT familiesmaterials.descripcio, familiescolorants.descripcio, subfamiliesmaterials.descripcio FROM ((familiescolorants INNER JOIN materials ON familiescolorants.codi = materials.familiacol) INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi WHERE (((materials.codi)=" + atrim(codimat) + "));", dbOpenSnapshot, dbReadOnly)
  rstcomandes.FindFirst "comanda=" + atrim(numc2)
  codimat = IIf(Not rstcomandes.NoMatch, cadbl(rstcomandes!materialex), 0)
  Set rstmat2 = dbtmp.OpenRecordset("SELECT familiesmaterials.descripcio, familiescolorants.descripcio, subfamiliesmaterials.descripcio FROM ((familiescolorants INNER JOIN materials ON familiescolorants.codi = materials.familiacol) INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi WHERE (((materials.codi)=" + atrim(codimat) + "));", dbOpenSnapshot, dbReadOnly)
  rstcomandes.FindFirst "comanda=" + atrim(numc3)
  codimat = IIf(Not rstcomandes.NoMatch, cadbl(rstcomandes!materialex), 0)
  Set rstmat3 = dbtmp.OpenRecordset("SELECT familiesmaterials.descripcio, familiescolorants.descripcio, subfamiliesmaterials.descripcio FROM ((familiescolorants INNER JOIN materials ON familiescolorants.codi = materials.familiacol) INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi WHERE (((materials.codi)=" + atrim(codimat) + "));", dbOpenSnapshot, dbReadOnly)
  If Not rstmat1.EOF Then
     rstcomandes.FindFirst "comanda=" + atrim(numc1)
     If Not rstcomandes.NoMatch Then
        descripciomat = descripciomaterialconcatenat(rstmat1)  'atrim(rstmat1![familiesmaterials.descripcio]), atrim(rstmat1![familiescolorants.descripcio]))rstmat1![subfamiliesmaterials.descripcio]
        espesormat1 = formcomandes.micresmaterial(cadbl(rstcomandes!mesuraesp), cadbl(rstcomandes!espessor), atrim(rstcomandes!tubolam))
     End If
  End If
  If Not rstmat2.EOF Then
     rstcomandes.FindFirst "comanda=" + atrim(numc2)
     If Not rstcomandes.NoMatch Then
        descripciomat = descripciomat + "/" + descripciomaterialconcatenat(rstmat2)
        espesormat2 = formcomandes.micresmaterial(cadbl(rstcomandes!mesuraesp), cadbl(rstcomandes!espessor), atrim(rstcomandes!tubolam))
     End If
  End If
  If Not rstmat3.EOF Then
     rstcomandes.FindFirst "comanda=" + atrim(numc3)
     If Not rstcomandes.NoMatch Then
        descripciomat = descripciomat + " // " + descripciomaterialconcatenat(rstmat3)
        espesormat3 = formcomandes.micresmaterial(cadbl(rstcomandes!mesuraesp), cadbl(rstcomandes!espessor), atrim(rstcomandes!tubolam))
     End If
  End If
  rstnou!micres = atrim(espesormat1) + IIf(cadbl(espesormat2) <> 0, "+" + atrim(espesormat2), "") + IIf(cadbl(espesormat3) <> 0, "+" + atrim(espesormat3), "")
  rstnou!descfamiliamat = descripciomat
  Set rstmat1 = Nothing
  Set rstmat2 = Nothing
  Set rstmat3 = Nothing
  Set rstcomandes = Nothing
End Sub

Sub passardadesdeltreball(rstnou As Recordset, numtreball As Double, ordre As Double)
   Dim rstclixes As Recordset
   If numtreball < 1 Then Exit Sub
   If ordre = 0 Then ordre = 1
   Set rstclixes = dbclixes.OpenRecordset("SELECT marca,linia, descripcioquantitatlinia, tinters, desarroll FROM Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball where clixes.id_treball = " + atrim(numtreball) + " And ordre = " + atrim(ordre), dbOpenSnapshot, dbReadOnly)
   If rstclixes.EOF Then Exit Sub
   rstnou!texteimpresio = atrim(rstclixes!marca) + " - " + atrim(rstclixes!linia) + " #" + atrim(rstclixes!descripcioquantitatlinia)
   rstnou!tintes = cadbl(rstclixes!tinters)
   rstnou!desarrollimp = cadbl(rstclixes!desarroll)
   Set rstclixes = Nothing
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
       DBEngine.CreateDatabase fitxertmpestats, dbLangGeneral, DatabaseTypeEnum.dbVersion30
    End If
    
    
End Sub
Sub borrartemps()
   On Error Resume Next
    Kill fitxertmpestats
End Sub
Sub carregarllistadecampstemporals()
  Dim i As Byte
  i = 1
  campsestat(i, 1) = "datacomanda": campsestat(i, 2) = "date": campsestat(i, 3) = "Data_Com": i = i + 1
  campsestat(i, 1) = "numcomandes": campsestat(i, 2) = "long": campsestat(i, 3) = "Nº_Com": i = i + 1
  campsestat(i, 1) = "refinplacsa": campsestat(i, 2) = "string": campsestat(i, 3) = "Ref_Inplacsa": i = i + 1
  campsestat(i, 1) = "refclient": campsestat(i, 2) = "string": campsestat(i, 3) = "Ref_Client": i = i + 1
  campsestat(i, 1) = "refclientvella": campsestat(i, 2) = "string": campsestat(i, 3) = "Ref_Client_vella": i = i + 1
  campsestat(i, 1) = "stopped": campsestat(i, 2) = "string": campsestat(i, 3) = "Stopped": i = i + 1
  campsestat(i, 1) = "numcomanda": campsestat(i, 2) = "double": campsestat(i, 3) = "Nº_Comanda": i = i + 1
  campsestat(i, 1) = "producte": campsestat(i, 2) = "string": campsestat(i, 3) = "Producte": i = i + 1
  campsestat(i, 1) = "texteimpresio": campsestat(i, 2) = "string": campsestat(i, 3) = "Marca_i_Linia": i = i + 1
  campsestat(i, 1) = "ampleext": campsestat(i, 2) = "string": campsestat(i, 3) = "EAmple": i = i + 1
  campsestat(i, 1) = "amplereb": campsestat(i, 2) = "double": campsestat(i, 3) = "RAmple": i = i + 1
  campsestat(i, 1) = "peskg": campsestat(i, 2) = "double": campsestat(i, 3) = "Kg_entregats": i = i + 1
  campsestat(i, 1) = "desarrollimp": campsestat(i, 2) = "long": campsestat(i, 3) = "Desarroll_Imp": i = i + 1
  campsestat(i, 1) = "tintes": campsestat(i, 2) = "byte": campsestat(i, 3) = "Tintes": i = i + 1
  campsestat(i, 1) = "simulteneitatreb": campsestat(i, 2) = "byte": campsestat(i, 3) = "Bandes_Reb": i = i + 1
  campsestat(i, 1) = "amplesol": campsestat(i, 2) = "string": campsestat(i, 3) = "SAmple": i = i + 1
  campsestat(i, 1) = "longitud": campsestat(i, 2) = "string": campsestat(i, 3) = "SLongitud": i = i + 1
  campsestat(i, 1) = "solapa": campsestat(i, 2) = "string": campsestat(i, 3) = "Solapa_Sol": i = i + 1
  campsestat(i, 1) = "tipussoldadura": campsestat(i, 2) = "string": campsestat(i, 3) = "Tipus_Soldadura": i = i + 1
  campsestat(i, 1) = "micres": campsestat(i, 2) = "string": campsestat(i, 3) = "Espesor": i = i + 1
  campsestat(i, 1) = "descfamiliamat": campsestat(i, 2) = "string": campsestat(i, 3) = "Desc_Families": i = i + 1
  campsestat(i, 1) = "vref1": campsestat(i, 2) = "byte": campsestat(i, 3) = "Vref1": i = i + 1
  campsestat(i, 1) = "vref2": campsestat(i, 2) = "string": campsestat(i, 3) = "Vref2": i = i + 1
  campsestat(i, 1) = "nomclient": campsestat(i, 2) = "string": campsestat(i, 3) = "NomClient": i = i + 1
  campsestat(i, 1) = "direnvio": campsestat(i, 2) = "string": campsestat(i, 3) = "DireccioEnviament": i = i + 1
  campsestat(i, 1) = "clientfacturacio": campsestat(i, 2) = "string": campsestat(i, 3) = "NomClientFacturació": i = i + 1
  campsestat(i, 1) = "quantitatteorica": campsestat(i, 2) = "double": campsestat(i, 3) = "Quantitat_Teòrica": i = i + 1
  campsestat(i, 1) = "unitatteorica": campsestat(i, 2) = "string": campsestat(i, 3) = "Unitat_Teòrica": i = i + 1
  campsestat(i, 1) = "numtreball": campsestat(i, 2) = "double": campsestat(i, 3) = "NºTreball": i = i + 1
  campsestat(i, 1) = "numpecesentregades": campsestat(i, 2) = "double": campsestat(i, 3) = "Peces_Ent": i = i + 1
  campsestat(i, 1) = "metresentregats": campsestat(i, 2) = "double": campsestat(i, 3) = "Mts_Ent": i = i + 1
  campsestat(i, 1) = "kilosentregats": campsestat(i, 2) = "double": campsestat(i, 3) = "Kg_Ent": i = i + 1
  campsestat(i, 1) = "pvp_unitat": campsestat(i, 2) = "double": campsestat(i, 3) = "Pvp_unitat": i = i + 1
  campsestat(i, 1) = "unitatpvp": campsestat(i, 2) = "string": campsestat(i, 3) = "Unitat_Pvp": i = i + 1
  campsestat(i, 1) = "PVP": campsestat(i, 2) = "double": campsestat(i, 3) = "Pvp": i = i + 1
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
   escriure_ini "AmplesReixa", UCase(reixa.TextMatrix(0, j)), atrim(Redondejar(reixa.ColWidth(j), 0)), iniconfigreixa
 Next j
End If
End Sub

Private Sub Form_Resize()
   If Formreixallistatproduccions.Height - reixa.Top - 800 < 1 Then Exit Sub
   reixa.Width = Formreixallistatproduccions.Width - 300
   reixa.Height = Formreixallistatproduccions.Height - reixa.Top - 800
   Fbotons.Left = Formreixallistatproduccions.Width - Fbotons.Width - 300
   etregistres.Top = reixa.Height + reixa.Top
   If Formreixallistatproduccions.Tag <> "canvianttamany" Then
    escriure_ini "TamanyForm", "ample", atrim(Formreixallistatproduccions.Width), iniconfigreixa
    escriure_ini "TamanyForm", "alt", atrim(Formreixallistatproduccions.Height), iniconfigreixa
   End If
End Sub
Sub carregartamanyform()
  If cadbl(llegir_ini("TamanyForm", "ample", iniconfigreixa)) > 0 Then
   Formreixallistatproduccions.Tag = "canvianttamany"
   Formreixallistatproduccions.Width = llegir_ini("TamanyForm", "ample", iniconfigreixa)
   Formreixallistatproduccions.Height = llegir_ini("TamanyForm", "alt", iniconfigreixa)
   Formreixallistatproduccions.Tag = ""
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
      vordre = reixa.TextMatrix(0, reixa.col + 1) 'campsestat(reixa.col + 1, 1)
     ' If vordre = "dataobertura" Then vordre = "cvdate(dataobertura)"
      If InStr(1, bordre.Tag, vordre) > 0 Then
          If InStr(1, bordre.Tag, "ASC") > 0 Then
                bordre.Tag = " DESC"
              Else: bordre.Tag = " ASC"
          End If
           Else
              bordre.Tag = " ASC"
      End If
      etordre = reixa.TextMatrix(0, reixa.col + 1) + " " + bordre.Tag
      bordre.Tag = vordre + bordre.Tag
      etmsgajuda.Visible = False
      bordre.BackColor = treurefiltre.BackColor
      reixa.BackColorFixed = treurefiltre.BackColor
      poblarlareixa whereultimfiltre
  End If
End Sub

Private Sub reixa_DblClick()
'  rstconsulta.FindFirst "maxcomanda=" + atrim(cadbl(reixa.TextMatrix(reixa.row, 4)))
 ' Formreixallistatproduccions.Hide
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
'  If Me.Caption = "Consulta referencies d'Inplacsa...(Carregant)" Then
'     On Error Resume Next
'     Unload Me
'       Else: formreixallistatproduccions.Hide
 ' End If
 Me.Caption = "Tancar"
 Formreixallistatproduccions.Hide
End Sub

Private Sub treurefiltre_Click()
 borrarelfiltre
End Sub
