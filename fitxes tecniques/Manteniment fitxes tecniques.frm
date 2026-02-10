VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form form_mantenimentFT 
   Caption         =   "Manteniment de Fitxes Tècniques"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15570
   Icon            =   "Manteniment fitxes tecniques.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   15570
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton exportaraxls 
      BackColor       =   &H00F0F0F0&
      Height          =   480
      Left            =   14895
      Picture         =   "Manteniment fitxes tecniques.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Exportar a Excel la sel.lecció"
      Top             =   90
      Width           =   615
   End
   Begin VB.ListBox listdown 
      BackColor       =   &H0080FFFF&
      Height          =   840
      Left            =   7395
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton botodown 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      MaskColor       =   &H00DADAFE&
      TabIndex        =   7
      Top             =   375
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CommandButton bperfamilia 
      Caption         =   "Agrupar Xr Familia"
      Height          =   510
      Left            =   1845
      TabIndex        =   6
      Top             =   105
      Width           =   1650
   End
   Begin VB.CommandButton bperarticle 
      Caption         =   "Agrupar Xr Article"
      Height          =   510
      Left            =   135
      TabIndex        =   4
      Top             =   105
      Width           =   1650
   End
   Begin VB.CommandButton Command56 
      Height          =   270
      Left            =   30
      Picture         =   "Manteniment fitxes tecniques.frx":0EA8
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Eliminar totes les linies"
      Top             =   945
      Width           =   240
   End
   Begin VB.TextBox filtre 
      BackColor       =   &H00FFC0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   930
      Width           =   630
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      DragIcon        =   "Manteniment fitxes tecniques.frx":1432
      Height          =   5700
      Left            =   285
      TabIndex        =   1
      Top             =   1200
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   10054
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   14737632
      BackColorSel    =   15971192
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
      FormatString    =   ""
      OLEDropMode     =   1
   End
   Begin VB.Label etagrupacio 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   75
      TabIndex        =   5
      Top             =   660
      Width           =   1500
   End
   Begin VB.Label ettotalregistres 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   6915
      Width           =   3495
   End
   Begin VB.Menu m_opcions_reixa 
      Caption         =   "opcions_reixa"
      Visible         =   0   'False
      Begin VB.Menu m_art_sense_FT 
         Caption         =   "Article sense fitxa tècnica"
         Begin VB.Menu m_segurno 
            Caption         =   "Segur?  No"
         End
         Begin VB.Menu m_segur_si 
            Caption         =   "Segur? Si"
         End
      End
      Begin VB.Menu m_veurearticlessenseFT 
         Caption         =   "Veure articles sense FT"
      End
   End
   Begin VB.Menu m_opcions_reixa_eliminar 
      Caption         =   "opcions_reixa_eliminar"
      Visible         =   0   'False
      Begin VB.Menu meliminarpdf 
         Caption         =   "Eliminar Pdf relacionat"
      End
   End
End
Attribute VB_Name = "form_mantenimentFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim dbtintes As Database
 Dim ultimweres As String
 Dim werereixa As String
 Dim vnumcolseleccionada As Double
 Dim ruta_documentacio_ft As String
 Dim vnumcolumnaclickreixa As Integer

Private Sub exportaraxls_Click()
   generar_xls
End Sub

Private Sub listdown_Click()
   Dim rst As Recordset
   Dim vcodiart As String
   Dim vdesc As String
   Dim vidioma As String
   vcodiart = reixa.TextMatrix(reixa.Row, 0)
   If cadbl(vcodiart) = 0 Then Exit Sub
   'Set rst = dbcomandes.OpenRecordset("select * from fitxestecniques where codiarticle=" + vcodiart)
   vidioma = Mid(listdown, 1, 2)
   vdesc = Mid(listdown.Text, 4)
   listdown.Visible = False
   If vidioma <> "" Then
     vdesc = UCase(InputBox("Escriu el nom comercial per l'idioma " + vidioma, "Nom comercial", vdesc))
     vdesc = treure_apostruf(vdesc)
     If vdesc <> "" Then dbcomandes.Execute "update fitxestecniques set nomcomercial_" + vidioma + "='" + atrim(vdesc) + "' where codiarticle=" + vcodiart
     
   End If
   Set rst = Nothing
End Sub

Private Sub listdown_LostFocus()
  listdown.Visible = False
End Sub

 Private Sub m_veurearticlessenseFT_Click()

  Load formseleccio
  formseleccio.Caption = "Selecciona el material que vols tornar visible"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codi,descripcio from materials where materials.noteFT order by descripcio"
  formseleccio.refrescar
  formseleccio.Command3.Tag = "filtre"
  'formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Width = 4500
  formseleccio.DBGrid2.Columns(0).Width = 500
  formseleccio.Show 1
  If seleccioret = 1 Then
   If cadbl(formseleccio.Data1.Recordset!codi) <> 0 Then
     If MsgBox("Segur que vols activar la FT d'aquest material?" + Chr(10) + atrim(UCase(formseleccio.Data1.Recordset!descripcio)), vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
       dbcomandes.Execute "update  materials set noteFT=false where codi=" + atrim(formseleccio.Data1.Recordset!codi)
       If etagrupacio.Tag = "F" Then bperfamilia_Click
       If etagrupacio.Tag = "A" Then bperarticle_Click
     End If
   End If
  End If
  Unload formseleccio
  
 End Sub
 Private Sub m_segur_si_click()
   Dim vcodiarticle As String
   vcodiarticle = reixa.TextMatrix(reixa.Row, 0)
   If cadbl(vcodiarticle) = 0 Then Exit Sub
   dbcomandes.Execute "update  materials set noteFT=true where codi=" + vcodiarticle
   
   If etagrupacio.Tag = "F" Then bperfamilia_Click
   If etagrupacio.Tag = "A" Then bperarticle_Click
 End Sub
 Private Sub meliminarpdf_click()
   Dim vcampescullit As String
   Dim vcodimaterial As String
   Dim vnomfitxerdesti As String
   vcampescullit = IIf(InStr(1, UCase(reixa.TextMatrix(0, vnumcolumnaclickreixa)), "CONFORMITAT"), "CONFORMITAT", vcampescullit)
   vcampescullit = IIf(InStr(1, UCase(reixa.TextMatrix(0, vnumcolumnaclickreixa)), "SEGURETAT"), "SEGURETAT", vcampescullit)
   vcampescullit = IIf(InStr(1, UCase(reixa.TextMatrix(0, vnumcolumnaclickreixa)), "FT"), "FT", vcampescullit)
   
   If MsgBox("Segur que vols desvincular aquest PDF de " + vcampescullit + "?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   vcodimaterial = reixa.TextMatrix(reixa.Row, numcol("Codi"))
   vnomfitxerdesti = ruta_documentacio_ft + "\" + vcodimaterial + "\" + vcampescullit + "_" + vcodimaterial + "*.pdf"
   Kill vnomfitxerdesti
   If vcampescullit = "FT" Then
     dbcomandes.Execute "update fitxestecniques set master=false,pdf_fitxatecnica=false,datavigenciafitxatecnica=null where codiarticle=" + vcodimaterial
     reixa.TextMatrix(reixa.Row, numcol("Pdf FT")) = "N"
     reixa.TextMatrix(reixa.Row, numcol("Vigència FT")) = ""
     reixa.TextMatrix(reixa.Row, numcol("Master")) = "N"
   End If
   If vcampescullit = "SEGURETAT" Then
     dbcomandes.Execute "update fitxestecniques set pdf_fitxaseguretat=false,datavigenciafitxaseguretat=null where codiarticle=" + vcodimaterial
     reixa.TextMatrix(reixa.Row, numcol("Pdf Seguretat")) = "N"
     reixa.TextMatrix(reixa.Row, numcol("Vigència Seguretat")) = ""
   End If
   If vcampescullit = "CONFORMITAT" Then
     dbcomandes.Execute "update fitxestecniques set pdf_conformitat=false,datavigenciaconformitat=null where codiarticle=" + vcodimaterial
     reixa.TextMatrix(reixa.Row, numcol("Pdf Conformitat")) = "N"
     reixa.TextMatrix(reixa.Row, numcol("Vigència conformitat")) = ""
   End If
   
 End Sub
Sub netejar_reixa(vArtFam As String)
   Dim rst As Recordset
   Dim i As Byte
   Dim col As Byte
   Set rst = dbcomandes.OpenRecordset("select * from fitxestecniques")
   reixa.Rows = 1
   reixa.Cols = 1
   col = 0
   For i = 0 To rst.Fields.Count - 1
     If valorpropietat(rst.Fields(i), "Caption") <> "" Then
     ' If vArtFam = "A" Then If rst.Fields(i).Name = "nomfamilia" Then GoTo cont
     ' If vArtFam = "F" Then If rst.Fields(i).Name = "nomarticle" Then GoTo cont
      reixa.Cols = col + 1
      reixa.col = col
      reixa.Text = valorpropietat(rst.Fields(i), "Caption")
      If filtre.Count <= col Then Load filtre(col)
      If Screen.ActiveControl.Name <> "filtre" Then
       filtre(col).DataField = rst.Fields(i).Name
       filtre(col).Text = valorpropietat(rst.Fields(i), "Caption")
      End If
      col = col + 1
cont:
     End If
   Next i
   
End Sub
Function valorpropietat(rst As Field, v As String) As String
  Dim i As Byte
  For i = 0 To rst.Properties.Count - 1
      If rst.Properties(i).Name = v Then valorpropietat = rst.Properties(i)
  Next i
End Function
Sub poblar_reixa(vArtFam As String)
   Dim rst As Recordset
   Dim fila As Integer
   Dim i As Byte
   Dim col As Integer
   reixa.Redraw = False
   netejar_reixa etagrupacio.Tag
   'ettotalcomandes.Caption = ""
   Set rst = dbcomandes.OpenRecordset("select * from fitxestecniques " + IIf(werereixa <> "", " where " + werereixa, " order by nomcomercial_es"))
   If rst.EOF Then GoTo fi
   rst.MoveLast
   rst.MoveFirst
   fila = 1
   While Not rst.EOF
      col = 0
      If vArtFam = "A" And cadbl(rst!codiarticle) = 0 Then GoTo contreg
      If vArtFam = "F" And (Not rst!master And cadbl(rst!codiarticle) <> 0) Then GoTo contreg
      reixa.Rows = fila + 1
      For i = 0 To rst.Fields.Count - 1
       ' If vArtFam = "A" Then If rst.Fields(i).Name = "nomfamilia" Then GoTo cont
       ' If vArtFam = "F" Then If rst.Fields(i).Name = "nomarticle" Then GoTo cont
        If rst.Fields(i).Name = "nomcomercial_es" Then
          If atrim(rst.Fields("nomcomercial_es")) = "" Or atrim(rst.Fields("nomcomercial_fr")) = "" Or atrim(rst.Fields("nomcomercial_en")) = "" Then
               reixa.col = col
               reixa.Row = fila
               reixa.CellBackColor = &HDADAFE   'vermell clarisim
        End If
        End If
        If valorpropietat(rst.Fields(i), "Caption") <> "" Then
            possar_el_valor_alareixa fila, col, rst.Fields(i), rst
            col = col + 1
        End If
        
cont:
      Next i
      fila = fila + 1
contreg:
      rst.MoveNext
   Wend
fi:
  ettotalregistres.Caption = "Registres: " + atrim(reixa.Rows - 1)
   reixa.Redraw = True
   Set rst = Nothing
End Sub

Sub possar_el_valor_alareixa(fila As Integer, col As Integer, vcamp As Field, vrst As Recordset)
  Dim v As String
  Dim vcolor As Double
  If vcamp.Type = 1 Then v = IIf(vcamp.Value, "S", "N")
  If vcamp.Type = 4 Or vcamp.Type = 7 Then v = cadbl(vcamp.Value)
  If vcamp.Type = 10 Then v = atrim(vcamp.Value)
  If v = "" Then v = atrim(vcamp.Value)
  reixa.TextMatrix(fila, col) = v
  'If vcamp.Name = "numtreball" Then If cadbl(vcamp.Value) < 0 Then vcolor = QBColor(12)
  'If vcamp.Name = "numtreball" Then If nhihandosiguals(cadbl(v), vrst.Fields("versiotreball")) Then vcolor = QBColor(14)
  'If vcamp.Name = "numtreball" Then If vrst.Fields("estatclixe") = "POLIMERS O CLIXES" Then vcolor = &HF3B378  'blau 'GoTo fi
  'If vcamp.Name = "numtreball" Then
  '  If atrim(vrst!maquina) = "FS" Then vcolor = QBColor(5)
  'End If
  'If vcamp.Name = "metres" Then If cadbl(vcamp.Value) = 0 Then vcolor = QBColor(12)
  'If vcamp.Name = "gestionat" Then
  '    If vcamp.Value = "S" Then vcolor = &HC0FFC0
  '    If vcamp.Value = "C" Then vcolor = QBColor(12)
  '    If vcamp.Value = "M" Then vcolor = QBColor(14)
  '    If vcamp.Value = "P" Then vcolor = &H80C0FF
  '    If vcamp.Value = "N" Then vcolor = 0
  'End If
  If vcolor > 0 Then
     reixa.col = col
     reixa.Row = fila
     reixa.CellBackColor = vcolor
  End If
End Sub
Sub guardar_amples_reixa()
Dim j As Integer
If iniconfigreixa <> "" Then
  For j = 0 To reixa.Cols - 1
   escriure_ini "AmplesReixaFitxestecniques", UCase(reixa.TextMatrix(0, j)), atrim(reixa.ColWidth(j)), iniconfigreixa
 Next j
End If
End Sub
Sub carregar_amples_reixa()
 Dim ample As String
 Dim X As Long
 Dim j As Integer
 If iniconfigreixa <> "" Then ' existeix("c:\windows\" + iniconfigreixa) Then
  X = reixa.Left + 35
  For j = 0 To reixa.Cols - 1
   ample = llegir_ini("AmplesReixaFitxestecniques", UCase(reixa.TextMatrix(0, j)), iniconfigreixa)
   If ample = "{[}]" Then ample = 1000
   reixa.ColWidth(j) = cadbl(ample)
    If X < reixa.Width Then
     filtre(j).Left = X
     filtre(j).Width = cadbl(ample)
     filtre(j).Visible = True
     filtre(j).ForeColor = &H808080
      Else: If filtre.Count < j - 1 Then filtre(j).Visible = False
    End If
    X = X + cadbl(ample)
 Next j
End If
filtre(0).Width = filtre(0).Width - 50
filtre(0).Left = filtre(0).Left + 50
End Sub


Private Sub botodown_Click()
   posarvalorsalallista
End Sub
Sub posarvalorsalallista()
    Dim rst As Recordset
  Dim vcodiart As String
  vcodiart = reixa.TextMatrix(reixa.Row, 0)
  If cadbl(vcodiart) = 0 Then Exit Sub

  listdown.Clear
  Set rst = dbcomandes.OpenRecordset("select * from fitxestecniques where codiarticle=" + vcodiart)
  If Not rst.EOF Then
     listdown.AddItem "    "
     listdown.AddItem "ES: " + atrim(rst!nomcomercial_es)
     listdown.AddItem "FR: " + atrim(rst!nomcomercial_fr)
     listdown.AddItem "EN: " + atrim(rst!nomcomercial_en)
  End If
  listdown.Visible = True
  listdown.SetFocus
End Sub

Private Sub bperarticle_Click()
   ratoli "espera"
   actualitzardadesarticlesFT
  'refrescar_dades_comandesactives
  etagrupacio = "Agrupat per Article"
  etagrupacio.Tag = "A"
  netejar_reixa etagrupacio.Tag
  poblar_reixa etagrupacio.Tag
  carregar_amples_reixa
  ordenarlareixa 1, 1
  ratoli "normal"
End Sub

Sub bxrcontrolagafafocus(i As Integer)
  Dim cntrl As Control
  Set cntrl = Screen.ActiveControl
  If cntrl.Text <> "" Then
     If cntrl.Text = reixa.TextMatrix(0, i) Then cntrl.Text = ""
     cntrl.ForeColor = QBColor(0)
   Else:
       cntrl.Text = reixa.TextMatrix(0, i)
       cntrl.ForeColor = &H808080
  End If
End Sub

Private Sub bperfamilia_Click()
  ratoli "espera"
  actualitzardadesarticlesFT
  'refrescar_dades_comandesactives
  etagrupacio = "Agrupat per Familia"
  etagrupacio.Tag = "F"
  netejar_reixa etagrupacio.Tag
  poblar_reixa etagrupacio.Tag
  carregar_amples_reixa
  ordenarlareixa 0, 2

  ratoli "normal"
End Sub
Sub ordenarlareixa(vnumcolumna As Integer, vordre As Byte)
  Static vultimacolumna
  reixa.col = vnumcolumna 'Desde que columna iniciar
  reixa.ColSel = vnumcolumna 'Hasta que columna terminar
  If reixa.Rows < 2 Then Exit Sub
  reixa.Row = 1 'Primer renglon del MsFlex a sortear
  reixa.RowSel = reixa.Rows - 1 'Ultimo renglon del msflex a sortear
   'vordre es 1 per ascendent i 2 per descendent
  If vordre = 0 Then
     vordre = 1
     If vultimacolumna = vnumcolumna Then
       vordre = 2: vultimacolumna = 9999
        Else: vultimacolumna = vnumcolumna
     End If
  End If
  reixa.Sort = vordre ' flexSortStringNoCaseAscending 'metodo de sorteo deseado
  reixa.Row = 1
  reixa.RowSel = 1
  
End Sub

Private Sub Command56_Click()
  werereixa = ""
   poblar_reixa etagrupacio.Tag
   carregar_amples_reixa
End Sub

Private Sub filtre_GotFocus(Index As Integer)
  bxrcontrolagafafocus Index
End Sub

Private Sub filtre_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then KeyCode = 0: filtre(IIf(Index = 0, 1, 0)).SetFocus
End Sub

Private Sub filtre_LostFocus(Index As Integer)
  ultimweres = werereixa
  werereixa = crearfiltre
  If filtre(Index).Text = "" Then
    filtre(Index).Text = reixa.TextMatrix(0, Index)
    filtre(Index).ForeColor = &H808080
  End If
  If ultimweres <> werereixa Then
   poblar_reixa etagrupacio.Tag
   carregar_amples_reixa
  End If
End Sub
Function crearfiltre() As String
  Dim i As Integer
  Dim were As String
  Dim w As String
  For i = 0 To filtre.Count - 1
    If filtre(i).Text <> reixa.TextMatrix(0, i) Then
      w = crearwere(i)
      If were = "" Then
         were = w
        Else: If w <> "" Then were = were + " and " + w
      End If
    End If
  Next i
  crearfiltre = were
End Function
Function crearwere(i As Integer) As String
   Dim w As String
   Dim j As Integer
   Dim rst As Recordset
   Dim vcamp As String
   If filtre(i) = "" Then Exit Function
   Set rst = dbcomandes.OpenRecordset("select * from fitxestecniques")
   vcamp = filtre(i).DataField
   If rst.Fields(vcamp).Type = 8 Then
      If IsDate(filtre(i)) Then
         crearwere = vcamp + "=#" + Format(filtre(i), "mm/dd/yy") + "# "
      End If
      GoTo fi
   End If
   If rst.Fields(vcamp).Type = 1 Then
         crearwere = vcamp + "=" + IIf(UCase(filtre(i)) = "S", "True", "False")
      GoTo fi
   End If
   If rst.Fields(vcamp).Type = 10 Then
       crearwere = possarweres(vcamp, "LIKE", treure_apostruf(filtre(i)))
       GoTo fi
   End If
   If InStr(1, filtre(i), ",") > 0 Then
       crearwere = vcamp + " in (" + atrim(filtre(i)) + ")"
     Else:
        If Not (Mid(filtre(i), 1, 1) = "<" Or Mid(filtre(i), 1, 1) = ">" Or Mid(filtre(i), 1, 1) = "=") Then
           crearwere = vcamp + "=" + passaradecimalpunt(atrim(cadbl(filtre(i))))
             Else: crearwere = vcamp + passaradecimalpunt(atrim(filtre(i)))
        End If
   End If
   
   
fi:
   Set rst = Nothing
End Function
Function possarweres(ByVal camp As String, condicio As String, ByVal filtre As String) As String
  Dim re As String
'camps(j, 1) + " LIKE '*" + treure_apostruf(filtre(i)) + "*'"
  filtre = filtre + ","
  'If camp = "nomclient" And cadbl(Mid(filtre, 1, InStr(1, filtre, ",") - 1)) > 0 Then camp = "codiclient"
  While InStr(1, filtre, ",") > 0 And filtre <> ""
    If camp <> "codiclient" Then
       re = IIf(re <> "", re + " or ", "") + camp + " like '*" + Mid(filtre, 1, InStr(1, filtre, ",") - 1) + "*'"
      Else: re = IIf(re <> "", re + " or ", "") + camp + " =" + atrim(cadbl(Mid(filtre, 1, InStr(1, filtre, ",") - 1))) + ""
    End If
    filtre = Mid(filtre, InStr(1, filtre, ",") + 1)
  Wend
  If re <> "" Then re = "(" + re + ")"
  possarweres = re
End Function

Sub carregartamanyform()
  If cadbl(llegir_ini("TamanyFormFT", "ample", iniconfigreixa)) > 0 Then
   form_mantenimentFT.Tag = "canvianttamany"
   form_mantenimentFT.Width = llegir_ini("TamanyFormFT", "ample", iniconfigreixa)
   form_mantenimentFT.Height = llegir_ini("TamanyFormFT", "alt", iniconfigreixa)
   form_mantenimentFT.Tag = ""
  End If

End Sub
Sub generar_xls()
   Dim i As Integer
   Dim j As Integer
   Dim rst As Recordset
   Dim linia As String
   Dim vprimerclient As Boolean
   
   Open "c:\temp\exportar-fitxestecniques.csv" For Output As #1
   i = 0
   While i < reixa.Rows
    linia = ""
    For j = 0 To reixa.Cols - 1
         linia = linia + IIf(linia = "", "", ";") + atrim(reixa.TextMatrix(i, j))
    Next j
    Print #1, linia
    i = i + 1
   Wend
   Close #1
   wait 2
   obrir_document "c:\temp\exportar-fitxestecniques.csv"
      
End Sub

Private Sub Form_Load()
  fitxerini = "comandes.ini"
  iniconfigreixa = "reixamantenimentFT.ini"
  cami = llegir_ini("General", "cami", fitxerini)
  ruta_documentacio_ft = llegir_ini("ruta", "ruta_documentacio_fitxestecniques", rutadelfitxer(cami) + "valorsprograma.ini")
  
  Set dbcomandes = DBEngine.OpenDatabase(cami)
  Set dbtintes = DBEngine.OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
  carregartamanyform
  
  
 ' bperarticle_Click
End Sub
Function rutadelfitxer(cam As String) As String
   Dim c As Byte
   c = 0
   While InStr(c + 1, cam, "\") <> 0
    c = InStr(c + 1, cam, "\")
   Wend
   If c = 0 Then c = Len(cam)
   rutadelfitxer = Mid(cam, 1, c)
End Function
Sub actualitzardadesarticlesFT()
  Dim rsta As Recordset
  Dim rst As Recordset
  Dim vsql As String
  Dim vsql2 As String
  'seleccion i afegeixo els materials
  vsql = "SELECT materials.codi, materials.descripcio, familiesmaterials.descripcio AS f1, subfamiliesmaterials.descripcio AS sf1, familiescolorants.descripcio AS f2, subfamiliescolorants.descripcio AS sf2, familiesaditius.descripcio AS f3, subfamiliesaditius.descripcio AS sf3"
  vsql = vsql + " FROM (((((materials INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN familiescolorants ON materials.familiacol = familiescolorants.codi) INNER JOIN familiesaditius ON materials.familiaad = familiesaditius.codi) INNER JOIN subfamiliesaditius ON materials.subfamiliaad = subfamiliesaditius.codi) INNER JOIN subfamiliescolorants ON materials.subfamiliacol = subfamiliescolorants.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi "
 
  Set rst = dbcomandes.OpenRecordset("select * from fitxestecniques")
  Set rsta = dbcomandes.OpenRecordset(vsql + " where materials.codi not in (select codiarticle from fitxestecniques) and materials.codi>499 and materials.noteFT=false ")
  While Not rsta.EOF
    rst.AddNew
    rst!tipusmaterial = "M"
    rst!codiarticle = rsta!codi
    rst!nomarticle = atrim(rsta!descripcio)
    rst!nomfamilia = atrim(rsta!f1) + "-" + atrim(rsta!sf1) + "-" + atrim(rsta!f2) + "-" + atrim(rsta!sf2) + "-" + atrim(rsta!f3) + "-" + atrim(rsta!sf3)
    rst.Update
    rsta.MoveNext
  Wend
  
  'selcciono i afegeixo les tintes
  vsql = "SELECT tintes_tot.codi, First(tintes_tot.descripcio) AS pdescripcio, First(tintes_tot.refproveidor) AS prefproveidor, First(tintes_tot.descripciofam) AS p11, First(tintes_tot.descripciosubfam) AS p12, First(tintes_tot.descripciofamcol) AS p21, First(tintes_tot.descripciosubfamcol) AS p22"
  vsql = vsql + " From tintes_tot Where (((tintes_tot.nomproveidor) <> 'INPLACSA') and cdbl(codi) not in (select codiarticle from fitxestecniques)) GROUP BY tintes_tot.codi;"
 
  Set rst = dbcomandes.OpenRecordset("select * from fitxestecniques")
  Set rsta = dbtintes.OpenRecordset(vsql)
  While Not rsta.EOF
    rst.AddNew
    rst!tipusmaterial = "T"
    rst!codiarticle = rsta!codi
    rst!master = True
    rst!nomarticle = atrim(rsta!pdescripcio)
    rst!nomfamilia = atrim(rsta!p11) + "-" + atrim(rsta!p12) + "-" + atrim(rsta!p21) + "-" + atrim(rsta!p22)
    rst.Update
    rsta.MoveNext
  Wend
  
  vsql2 = "SELECT nomfamilia From fitxestecniques "
  vsql2 = vsql2 + " WHERE (((fitxestecniques.nomfamilia) Not In (select nomfamilia from fitxestecniques where master=true)))"
  vsql2 = vsql2 + " GROUP BY fitxestecniques.nomfamilia;"
  
  vsql = "SELECT fitxestecniques.nomfamilia, Count(fitxestecniques.id) AS CuentaDeid From fitxestecniques "
  vsql = vsql + " WHERE (((fitxestecniques.nomfamilia) Not In (select nomfamilia from fitxestecniques where master=true)))"
  vsql = vsql + " GROUP BY fitxestecniques.nomfamilia;"
  dbcomandes.Execute "Delete * from fitxestecniques where (codiarticle=0 or codiarticle=null) and nomfamilia not in (" + vsql2 + ")"
  dbcomandes.Execute "delete * from fitxestecniques where codiarticle not in (select codi from materials where noteFT=false) and tipusmaterial='M'"
  dbtintes.Execute "delete * from fitxestecniques where codiarticle not in (select cdbl(codi) from tintes) and tipusmaterial='T'"
  Set rsta = dbcomandes.OpenRecordset(vsql)
  While Not rsta.EOF
    rst.FindFirst "codiarticle=0 and nomfamilia='" + atrim(rsta!nomfamilia) + "'"
    If rst.NoMatch Then
      rst.AddNew
      rst!tipusmaterial = "M"
      rst!nomfamilia = atrim(rsta!nomfamilia)
      rst.Update
    End If
    rsta.MoveNext
  Wend
  Set rst = Nothing
  Set rsta = Nothing
End Sub
Private Sub Form_Resize()
If form_mantenimentFT.Height - reixa.Top - 800 < 1 Then Exit Sub
   reixa.Width = form_mantenimentFT.Width - 800
   reixa.Height = form_mantenimentFT.Height - reixa.Top - 800
   exportaraxls.Left = form_mantenimentFT.Width - exportaraxls.Width - 300
   'Fbotons.Left = form_mantenimentFT.Width - Fbotons.Width - 300
   ettotalregistres.Top = reixa.Height + reixa.Top
   If form_mantenimentFT.Tag <> "canvianttamany" Then
    escriure_ini "TamanyFormFT", "ample", atrim(form_mantenimentFT.Width), iniconfigreixa
    escriure_ini "TamanyFormFT", "alt", atrim(form_mantenimentFT.Height), iniconfigreixa
   End If
End Sub

Private Sub reixa_Click()
 vnumcolseleccionada = reixa.col
 If reixa.TextMatrix(0, vnumcolseleccionada) = "Nom Comercial" Then
   possarelcombodelnom
    Else: botodown.Visible = False
 End If
 reixa.col = 0
 reixa.ColSel = reixa.Cols - 1
End Sub
Sub possarelcombodelnom()
 botodown.Left = reixa.ColPos(reixa.col) + reixa.Left + reixa.ColWidth(reixa.col) - botodown.Width + 20
 botodown.Height = reixa.RowHeight(reixa.Row)
 botodown.Top = reixa.RowPos(reixa.Row) + reixa.Top + 20
 botodown.Visible = True
 listdown.Left = reixa.ColPos(reixa.col) + reixa.Left + 20
 listdown.Top = reixa.RowPos(reixa.Row) + reixa.RowHeight(reixa.Row) + reixa.Top + 20
 listdown.Width = reixa.ColWidth(reixa.col)
 
 
End Sub
Private Sub reixa_LostFocus()
  guardar_amples_reixa
End Sub

Private Sub reixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim vcodiart As Double
  vnumcolumnaclickreixa = numcolumnaonestaelpunter(X)
  vcodiart = cadbl(reixa.TextMatrix(reixa.Row, numcol("CodiArt")))
  If Button = 2 And vnumcolumnaclickreixa < 5 Then
     'Me.m_opcions_reixa.WindowList = True
     If vcodiart < 2000 And vcodiart > 0 Then
       Me.PopupMenu m_opcions_reixa
     End If
  End If
  
  If Button = 2 And vnumcolumnaclickreixa > 4 Then
     'Me.m_opcions_reixa.WindowList = True
     If vcodiart > 0 Then Me.PopupMenu m_opcions_reixa_eliminar
  End If
  
  If Y > 0 And Y < (reixa.CellHeight) Then
      ordenarlareixa reixa.col, 0
  End If
End Sub
Function numcolumnaonestaelpunter(X As Single) As String
   Dim i As Byte
   Dim n As Double
   For i = 0 To reixa.Cols - 1
     If X > reixa.ColPos(i) Then n = i ' IIf(i = 0, 0, i - 1)
   Next i
   numcolumnaonestaelpunter = n
End Function
Function numcol(nom As String) As Byte
  Dim i As Integer
   numcol = 0
   For i = 0 To reixa.Cols - 1
     If reixa.TextMatrix(0, i) = nom Then numcol = i
   Next i
   
End Function
Private Sub reixa_OLEDragDrop(Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim vnomfitxer As String
  Dim vdata As String
  Dim vcodimaterial As String
  Dim vnommaterial As String
  Dim vnomfitxerdesti As String
  Unload formescullirpdf
  vnomfitxer = Data.Files(1)
  If Not existeix(vnomfitxer) Then Exit Sub
  formescullirpdf.Show 1
  If formescullirpdf.Tag = "" Then GoTo fi
  vnommaterial = reixa.TextMatrix(reixa.Row, numcol("Article"))
  vnommaterial = treuresimbols(vnommaterial)
  vcodimaterial = reixa.TextMatrix(reixa.Row, numcol("Codi"))
  
  
  If formescullirpdf.Tag = "seguretat" Then
     vnomfitxerdesti = ruta_documentacio_ft + "\" + vcodimaterial + "\SEGURETAT_" + vcodimaterial + " " + vnommaterial + ".pdf"
     If existeix(vnomfitxerdesti) Then
        If MsgBox("Aquest fitxer ja existeix, vols sobreescriure'l?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
     End If
     vdata = InputBox("Entra la data de vigència de la SEGURETAT:", "Data vigència")
     If Not IsDate(vdata) Then MsgBox "Aquesta data no es vàlida.", vbCritical, "Error": Exit Sub
     If Not existeix(ruta_documentacio_ft + "\" + vcodimaterial) Then MkDir ruta_documentacio_ft + "\" + vcodimaterial
     Copiar_Fitxer vnomfitxer, vnomfitxerdesti
     dbcomandes.Execute "update fitxestecniques set datavigenciafitxaseguretat=#" + Format(vdata, "mm/dd/yy") + "#, pdf_fitxaseguretat=true where codiarticle=" + vcodimaterial
     reixa.TextMatrix(reixa.Row, numcol("Pdf Seguretat")) = "S"
     reixa.TextMatrix(reixa.Row, numcol("Vigència Seguretat")) = Format(vdata, "dd/mm/yyyy")
  End If
  If formescullirpdf.Tag = "conformitat" Then
     vnomfitxerdesti = ruta_documentacio_ft + "\" + vcodimaterial + "\CONFORMITAT_" + vcodimaterial + " " + vnommaterial + ".pdf"
     If existeix(vnomfitxerdesti) Then
        If MsgBox("Aquest fitxer ja existeix, vols sobreescriure'l?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
     End If
     vdata = InputBox("Entra la data de vigència de la CONFORMITAT:", "Data vigència")
     If Not IsDate(vdata) Then MsgBox "Aquesta data no es vàlida.", vbCritical, "Error": Exit Sub
     If Not existeix(ruta_documentacio_ft + "\" + vcodimaterial) Then MkDir ruta_documentacio_ft + "\" + vcodimaterial
     Copiar_Fitxer vnomfitxer, vnomfitxerdesti
     dbcomandes.Execute "update fitxestecniques set pdf_conformitat=true,datavigenciaconformitat=#" + Format(vdata, "mm/dd/yy") + "# where codiarticle=" + vcodimaterial
     reixa.TextMatrix(reixa.Row, numcol("Pdf Conformitat")) = "S"
     reixa.TextMatrix(reixa.Row, numcol("Vigència conformitat")) = Format(vdata, "dd/mm/yyyy")
  End If
  If formescullirpdf.Tag = "fitxatecnica" Then
     vnomfitxerdesti = ruta_documentacio_ft + "\" + vcodimaterial + "\FT_" + vcodimaterial + " " + vnommaterial + ".pdf"
     If existeix(vnomfitxerdesti) Then
        If MsgBox("Aquest fitxer ja existeix, vols sobreescriure'l?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
     End If
     vdata = InputBox("Entra la data de vigència de la FITXA TÈCNICA:", "Data vigència")
     If Not IsDate(vdata) Then MsgBox "Aquesta data no es vàlida.", vbCritical, "Error": Exit Sub
     If Not existeix(ruta_documentacio_ft + "\" + vcodimaterial) Then MkDir ruta_documentacio_ft + "\" + vcodimaterial
     Copiar_Fitxer vnomfitxer, vnomfitxerdesti
     dbcomandes.Execute "update fitxestecniques set master=true,pdf_fitxatecnica=true,datavigenciafitxatecnica=#" + Format(vdata, "mm/dd/yy") + "# where codiarticle=" + vcodimaterial
     reixa.TextMatrix(reixa.Row, numcol("Pdf FT")) = "S"
     reixa.TextMatrix(reixa.Row, numcol("Vigència FT")) = Format(vdata, "dd/mm/yyyy")
     reixa.TextMatrix(reixa.Row, numcol("Master")) = "S"
  End If
  
  
fi:
  Unload formescullirpdf
End Sub

Private Sub reixa_Scroll()
  botodown.Visible = False
  listdown.Visible = False
End Sub
