VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form fbuscar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscador de Clixes"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15060
   ControlBox      =   0   'False
   Icon            =   "formbuscar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Resultat de la Busqueda"
      Height          =   4350
      Left            =   90
      TabIndex        =   1
      Top             =   2280
      Width           =   14790
      Begin VB.ListBox postit 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   2370
         Left            =   2685
         TabIndex        =   15
         Top             =   735
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid reixa 
         Height          =   3990
         Left            =   135
         TabIndex        =   2
         Top             =   210
         Width           =   14490
         _ExtentX        =   25559
         _ExtentY        =   7038
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dades a Buscar"
      Height          =   2205
      Left            =   105
      TabIndex        =   0
      Top             =   45
      Width           =   14760
      Begin VB.CommandButton Command2 
         Height          =   540
         Left            =   13020
         Picture         =   "formbuscar.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Buscar Registres i el resultat possar-lo a Excel."
         Top             =   960
         Width           =   1545
      End
      Begin VB.CheckBox cnomesreprint 
         Caption         =   "Només amb reprint"
         Height          =   240
         Left            =   11115
         TabIndex        =   27
         Top             =   570
         Width           =   1845
      End
      Begin VB.ComboBox Combofotogravador 
         DataSource      =   "clixes"
         Height          =   315
         Left            =   1305
         TabIndex        =   25
         Top             =   1305
         Width           =   4980
      End
      Begin VB.TextBox canilox 
         Height          =   315
         Left            =   11610
         TabIndex        =   24
         Top             =   1290
         Width           =   735
      End
      Begin VB.ComboBox combotinta 
         DataSource      =   "clixes"
         Height          =   315
         Left            =   7410
         TabIndex        =   21
         Top             =   1275
         Width           =   3645
      End
      Begin VB.CommandButton Command1 
         Height          =   330
         Left            =   10680
         Picture         =   "formbuscar.frx":0A83
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Consultar comandes amb aquest arxiu."
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox arxiu 
         DataField       =   "refclient"
         DataSource      =   "datavinculats"
         Height          =   285
         Left            =   9435
         TabIndex        =   18
         Top             =   510
         Width           =   1215
      End
      Begin VB.CheckBox edicioimps 
         Caption         =   "Activar l'edició dels IMPs als post-it de  nom de client"
         Height          =   240
         Left            =   4155
         TabIndex        =   17
         Top             =   1890
         Width           =   5055
      End
      Begin VB.CheckBox agrupantpertreball 
         Caption         =   "Agrupant per treball i modificació actual"
         Height          =   240
         Left            =   105
         TabIndex        =   16
         Top             =   1875
         Width           =   3900
      End
      Begin VB.ComboBox nomclient 
         DataField       =   "nomclient"
         DataSource      =   "datavinculats"
         Height          =   315
         Left            =   690
         TabIndex        =   12
         Top             =   495
         Width           =   3165
      End
      Begin VB.TextBox refclient 
         DataField       =   "refclient"
         DataSource      =   "datavinculats"
         Height          =   285
         Left            =   7425
         TabIndex        =   11
         Top             =   510
         Width           =   1980
      End
      Begin VB.TextBox codidebarres 
         DataField       =   "codidebarres"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   3945
         TabIndex        =   7
         Top             =   510
         Width           =   2325
      End
      Begin VB.ComboBox marcaproducte 
         DataSource      =   "clixes"
         Height          =   315
         Left            =   660
         TabIndex        =   6
         Top             =   900
         Width           =   5655
      End
      Begin VB.ComboBox liniaproducte 
         DataSource      =   "clixes"
         Height          =   315
         Left            =   7425
         TabIndex        =   5
         Top             =   870
         Width           =   5400
      End
      Begin VB.CommandButton sortir 
         Height          =   540
         Left            =   13020
         Picture         =   "formbuscar.frx":100D
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Sortir"
         Top             =   1485
         Width           =   1545
      End
      Begin VB.CommandButton consultar 
         Height          =   540
         Left            =   13020
         Picture         =   "formbuscar.frx":1597
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Buscar Registres"
         Top             =   420
         Width           =   1545
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fotogravador:"
         Height          =   285
         Left            =   135
         TabIndex        =   26
         Top             =   1365
         Width           =   1005
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Anilox:"
         Height          =   285
         Left            =   11115
         TabIndex        =   23
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tinta:"
         Height          =   285
         Left            =   6900
         TabIndex        =   22
         Top             =   1335
         Width           =   525
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Arxiu"
         Height          =   270
         Left            =   9825
         TabIndex        =   19
         Top             =   285
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Nom del Client"
         Height          =   285
         Left            =   1530
         TabIndex        =   14
         Top             =   270
         Width           =   1380
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia client"
         Height          =   270
         Left            =   7560
         TabIndex        =   13
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Codi de Barres"
         Height          =   270
         Left            =   4530
         TabIndex        =   10
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         Height          =   300
         Left            =   135
         TabIndex        =   9
         Top             =   915
         Width           =   675
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Linia:"
         Height          =   285
         Left            =   6900
         TabIndex        =   8
         Top             =   930
         Width           =   525
      End
   End
End
Attribute VB_Name = "fbuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstcomandesactives As Recordset

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
Sub triartinta()
  Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim caigudes As Double
  
  sql = "SELECT  codi,descripcio,referenciacolor from tintes_tot "
  were = " order by descripcio"
  Load formseleccio
  formseleccio.Data1.DatabaseName = rutadelfitxer(camiclixes) + "tintes.mdb"
  formseleccio.Data1.RecordSource = sql + were
  formseleccio.width = 14000
  formseleccio.sortirs.tag = "filtre"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
    combotinta = atrim(formseleccio.Data1.Recordset!descripcio)
    combotinta.tag = atrim(formseleccio.Data1.Recordset!codi)
  End If
  If seleccioret = 9 Then
    combotinta = ""
    combotinta.tag = ""
  End If
 '  Data1.Recordset!client = Text2.Text
 '  nomclient.Caption = atrim(formseleccio.Data1.Recordset!nom)
  
 ' End If
  Unload formseleccio
End Sub


Private Sub Combo1_DropDown()
    
End Sub

Private Sub Combofotogravador_DropDown()
    Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "select codi,nomfotogravador from fotogravadors where actiu"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.sortirs.tag = "filtre"
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           Combofotogravador = formseleccio.DBGrid2.Columns("nomfotogravador")
           Combofotogravador.tag = atrim(formseleccio.DBGrid2.Columns("CODI"))
        End If
   End If
   If seleccioret = 9 Then
        Combofotogravador.tag = ""
        Combofotogravador = ""
   End If
   Unload formseleccio
End Sub

Private Sub combotinta_DropDown()
   triartinta
End Sub

Private Sub Command1_Click()
   Dim comanda As String
   If arxiu = "" Then Exit Sub
   Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select comanda,datacomanda,client,numtreball,arxiu from comandes where arxiu='" + atrim(arxiu) + "'" + IIf(cadbl(nomclient.tag) > 0, "and client=" + atrim(nomclient.tag), "")
   
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar

   'formseleccio.DBGrid2.Columns(0).Width = 0
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           comanda = formseleccio.DBGrid2.Columns("comanda")
        End If
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   If comanda <> "" Then formclixes.cridarcomandes cadbl(comanda)
End Sub

Private Sub Check1_Click()

End Sub

Private Sub Command2_Click()


  Dim row As Integer
  Dim col As Integer
  Dim vnomfitxer As String
  ratoli "espera"
  ferlaconsulta
  vnomfitxer = "c:\temp\consultabuscadordeclixes.csv"
  row = 1
  col = 0
  reixa.Rows = 1
  If Not rstconsulta.EOF Then
     rstconsulta.MoveFirst
      Else: MsgBox "No hi ha cap resultat amb aquesta consulta.", vbExclamation, "Consulta": GoTo fi
  End If
'Borro el CSV anterior i dona error si no pot borrarlo
  If existeix(vnomfitxer) Then
     If Not borrar_fitxer(vnomfitxer) Then MsgBox "No es pot generar el Excel mira que no estigui obert.", vbCritical, "Error": GoTo fi
  End If
' obro el fitxer CSV
  Open "c:\temp\consultabuscadordeclixes.csv" For Output As #1
' posso la capcalera del CSV
  For col = 0 To rstconsulta.Fields.Count - 1
       linia = linia + IIf(linia = "", "", ";") + UCase(rstconsulta.Fields(col).Name)
  Next col
  Print #1, linia
'passo totes les dades al CSV
  While Not rstconsulta.EOF
    linia = ""
    If triarelsdelatintaconcreta(rstconsulta) Then
'      reixa.Rows = row + 1
      For col = 0 To rstconsulta.Fields.Count - 1
       linia = linia + IIf(linia = "", "", ";") + atrim(contingutcasella(col))
      Next col
      Print #1, linia
      row = row + 1
    End If
    rstconsulta.MoveNext
  Wend
  Close #1
'si existeix el CSV l'obro
  If existeix(vnomfitxer) Then obrir_document vnomfitxer
  
fi:
  
  ratoli "normal"
End Sub
Function borrar_fitxer(vfitxer As String) As Boolean
   On Error GoTo errors
   borrar_fitxer = True
   Kill vfitxer
   Exit Function
errors:
    borrar_fitxer = False
End Function

Private Sub consultar_Click()
  ratoli "espera"
  ferlaconsulta
  configreixa
  poblarreixa
  carregar_amples_reixa
  ratoli "normal"
End Sub
Sub ferlaconsulta()
   Dim selagrupant As String
   Dim sel As String
   Dim were As String
   Dim from As String
   Dim groupby As String
   groupby = " GROUP BY Clixes.id_treball"
   selagrupant = "SELECT First(Modificacions.id_modificacio) AS id_modificacio "
   sel = "SELECT Clixes.id_treball AS treball, Modificacions.ordre, Clixes.nomclienttemporal, Clixes.arxiu, Clientsvinculats.codiclient, clients.nom AS nom_client, Clixes.codidebarres, Clixes.marca, Clixes.linia, clientsvinculats.direnvio,Clientsvinculats.refclient, Clientsvinculats.refclientalternatives, Modificacions.datapdf,MODIFICACIONS.pdfvalid, Clixes.estatclixe, '' AS [lots actius], Modificacions.amplelamina, Modificacions.desarroll, Modificacions.tinters,modificacions.reimpres, false as noensenyar "
   
   'from = " from Linies INNER JOIN (Marques INNER JOIN (clients INNER JOIN (Clientsvinculats INNER JOIN (Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball) ON (Clientsvinculats.id_treball = Modificacions.id_treball) AND (Clientsvinculats.ordremodificacio = Modificacions.ordre)) ON clients.codi = Clientsvinculats.codiclient) ON Marques.id_marca = Clixes.id_marca) ON (Marques.id_marca = Linies.id_marca) AND (Linies.id_linia = Clixes.id_linia)"
   from = "   FROM clients RIGHT JOIN (Clientsvinculats RIGHT JOIN (Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball) ON (Clientsvinculats.id_treball = Modificacions.id_treball) AND (Clientsvinculats.ordremodificacio = Modificacions.ordre)) ON clients.codi = Clientsvinculats.codiclient "
   were = crearwereconsulta
   If were = "" Then
      If cadbl(combotinta.tag) > 0 Then
         were = " clients.codi>0 "
         MsgBox "Consulta feta sobre tots els treballs i potser que trigui una mica." + Chr(10) + "Fes <ACCEPTAR> per començar.", vbInformation, "Atenció"
           Else: were = "clixes.id_treball=-1"
      End If
   End If
 '  Clipboard.Clear
 '  Clipboard.SetText sel + from + " where " + were
   If were <> "" Then
      If agrupantpertreball.Value = 0 Then
        Set rstconsulta = dbconsulta.OpenRecordset(sel + from + " where " + were)
         Else:
           Set rstconsulta = dbconsulta.OpenRecordset(sel + from + " where modificacions.id_modificacio in (" + selagrupant + from + " where " + were + groupby + ")")
      End If
      If Not rstconsulta.EOF Then Set rstcomandesactives = dbcomandes.OpenRecordset("SELECT comanda,numtreball from comandes WHERE (((comandes.proximaseccio)<>'T' And (comandes.proximaseccio)<>'P' And (comandes.proximaseccio)<>'V')) order by numtreball Desc")
   End If
End Sub
Function triarelsdelatintaconcreta(rstconsulta As Recordset) As Boolean
   Dim rsttinta As Recordset
   Dim vconsultatinta As String
   Dim vconsultaanilox As String
   If cadbl(combotinta.tag) = 0 And cadbl(canilox) = 0 Then triarelsdelatintaconcreta = True: GoTo fi
   If cadbl(combotinta.tag) > 0 Then vconsultatinta = " coditinta='" + atrim(cadbl(combotinta.tag)) + "'"
   If cadbl(canilox) > 0 Then vconsultaanilox = IIf(vconsultatinta <> "", " and ", "") + " anilox=" + atrim(cadbl(canilox)) + ""
   Set rsttinta = dbclixes.OpenRecordset("select * from tintes where " + vconsultatinta + vconsultaanilox)
   If rsttinta.EOF Then GoTo fi
   rsttinta.FindFirst "id_treball=" + atrim(rstconsulta!treball) + " and ordremodificacio=" + atrim(rstconsulta!ordre)
   If Not rsttinta.NoMatch Then triarelsdelatintaconcreta = True
fi:
   Set rsttinta = Nothing
End Function
Function crearwereconsulta()
   Dim w As String
   If atrim(refclient) <> "" Then w = "(Clientsvinculats.refclient Like '*" + atrim(refclient) + "*' OR Clientsvinculats.refclientalternatives Like '*" + atrim(refclient) + "*')"
   If atrim(codidebarres) <> "" Then w = w + IIf(w <> "", " And ", "") + " clixes.codidebarres like'*" + atrim(codidebarres) + "*' "
   If cadbl(nomclient.tag) > 0 Then w = w + IIf(w <> "", " And ", "") + " (clientsvinculats.codiclient=" + atrim(cadbl(nomclient.tag)) + " or clixes.codiclienttemporal=" + atrim(cadbl(nomclient.tag)) + ")"
   If marcaproducte <> "" Then w = w + IIf(w <> "", " And ", "") + " clixes.marca like '*" + atrim(marcaproducte) + "*'"
   If liniaproducte <> "" Then w = w + IIf(w <> "", " And ", "") + " clixes.linia like '*" + atrim(liniaproducte) + "*'"
   If arxiu <> "" Then w = w + IIf(w <> "", " And ", "") + " clixes.arxiu = '" + atrim(arxiu) + "'"
   If cnomesreprint.Value = 1 Then w = w + IIf(w <> "", " And ", "") + " modificacions.reimpres = True "
   If cadbl(Combofotogravador.tag) > 0 Then w = w + IIf(w <> "", " And ", "") + " modificacions.fotograbador=" + atrim(cadbl(Combofotogravador.tag))
   'If cadbl(marcaproducte.Tag) = "-1" Then w = w + IIf(w <> "", " And ", "") + " marques.marca like '*" + atrim(marcaproducte) + "*'"
   'If cadbl(liniaproducte.Tag) = "-1" Then w = w + IIf(w <> "", " And ", "") + " linies.linia like '*" + atrim(liniaproducte) + "*'"
   crearwereconsulta = w
End Function
Sub carregar_amples_reixa()
 Dim ample As String
 Dim j As Integer
 If iniconfigreixa <> "" Then ' existeix("c:\windows\" + iniconfigreixa) Then
  For j = 0 To reixa.Cols - 1
   If Mid(rstconsulta.Fields(j).Name, 1, 3) <> "id_" Then
    ample = llegir_ini("AmplesReixa", UCase(reixa.TextMatrix(0, j)), iniconfigreixa)
    If ample <> "{[}]" Then
     reixa.ColWidth(j) = cadbl(ample)
    End If
   End If
 Next j
End If
End Sub
Sub guardar_amples_reixa()
Dim j As Integer
If iniconfigreixa <> "" Then
  For j = 0 To reixa.Cols - 1
   escriure_ini "AmplesReixa", UCase(reixa.TextMatrix(0, j)), atrim(reixa.ColWidth(j)), iniconfigreixa
 Next j
End If
End Sub
Private Sub Form_Load()
  Set dbconsulta = DBEngine.OpenDatabase(camiclixes)
  iniconfigreixa = "c:\windows\configreixabuscarclixes.ini"
End Sub

Sub configreixa()
  Dim rst As Recordset
  Dim col As Long
  Dim enes As Byte
  reixa.LeftCol = 0
  
  Set rst = rstconsulta
  If reixa.Rows > 1 Then reixa.TopRow = 1
  col = 0
  enes = 0
  reixa.Cols = rst.Fields.Count
  For i = 0 To rst.Fields.Count - 1
       reixa.ColAlignment(i) = 2
       reixa.TextMatrix(0, i) = UCase(rst.Fields(i).Name)
       If Mid(rstconsulta.Fields(i).Name, 1, 3) = "id_" Then
         reixa.ColWidth(i) = 1
       End If
       If rstconsulta.Fields(i).Name = "nomclienttemporal" Then
         reixa.ColWidth(i) = 1
       End If
       
  Next i
 
End Sub
Sub poblarreixa()
  Dim row As Integer
  Dim col As Integer
  row = 1
  col = 0
  reixa.Rows = 1
  If Not rstconsulta.EOF Then rstconsulta.MoveFirst
  While Not rstconsulta.EOF
    If triarelsdelatintaconcreta(rstconsulta) Then
      reixa.Rows = row + 1
      For col = 0 To rstconsulta.Fields.Count - 1
       reixa.TextMatrix(row, col) = atrim(contingutcasella(col))
      Next col
      row = row + 1
    End If
    rstconsulta.MoveNext
      
  Wend
  
End Sub
Function contingutcasella(col As Integer) As String
    Dim v As String
    Dim rstc As Recordset
    Dim cont As Integer
cont = 0
   v = atrim(rstconsulta.Fields(col))
   If rstconsulta.Fields(col).Name = "codidebarres" Then
     Set rstc = dbconsulta.OpenRecordset("SELECT Count(Clixes.id_treball) AS cidtreball, Clixes.codidebarres from clixes GROUP BY Clixes.codidebarres HAVING (((Clixes.codidebarres)='" + v + "'))")
     If Not rstc.EOF Then v = IIf(rstc!cidtreball > 1, "(" + atrim(rstc!cidtreball) + ") ", "") + v
   End If
   If rstconsulta.Fields(col).Name = "nom_client" Then
     Set rstc = dbconsulta.OpenRecordset("SELECT Clientsvinculats.id_treball, Count(Clientsvinculats.codiclient) AS cclients from Clientsvinculats GROUP BY Clientsvinculats.id_treball" + IIf(agrupantpertreball.Value = 0, " ,clientsvinculats.ordremodificacio=" + atrim(rstconsulta!ordre), "") + " HAVING Clientsvinculats.id_treball=" + atrim(cadbl(rstconsulta!treball)) + IIf(agrupantpertreball.Value = 0, " and clientsvinculats.ordremodificacio=" + atrim(cadbl(rstconsulta!ordre)), ""))
     If Not rstc.EOF Then
        v = IIf(rstc!cclients > 1, "(" + atrim(rstc!cclients) + ") ", "") + v
         Else: v = atrim(rstconsulta!nomclienttemporal)
     End If
   End If
   If rstconsulta.Fields(col).Name = "lots actius" Then
     rstcomandesactives.MoveFirst
     While Not rstcomandesactives.EOF And rstconsulta!treball < rstcomandesactives!numtreball
       If rstcomandesactives!numtreball = rstconsulta!treball Then
        v = v + " " + atrim(rstcomandesactives!comanda)
        cont = cont + 1
       End If
       rstcomandesactives.MoveNext
     Wend
     If v = "" Then v = "----------"
     v = IIf(cont > 1, "(" + atrim(cont) + ") ", "") + v
   End If
   
   
   contingutcasella = v
End Function



Private Sub liniaproducte_DropDown()
 
   If marcaproducte = "" Then MsgBox "Primer has d'escollir una marca.", vbCritical, "Atenció": Exit Sub
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "select distinct linia from clixes where marca='" + marcaproducte + "' order by linia"
   
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar

   'formseleccio.DBGrid2.Columns(0).Width = 0
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           liniaproducte = formseleccio.DBGrid2.Columns("linia")
        End If
   End If
   If seleccioret = 9 Then
           liniaproducte = ""
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub liniaproducte_KeyDown(KeyCode As Integer, Shift As Integer)

   'If marcaproducte.Tag <> "-1" Then msgbox "No pots buscar SEMBLANTS per linia de producte si la marca també busques SEMBLANTSliniaproducte.Tag = "": liniaproducte = ""
End Sub

Private Sub marcaproducte_DropDown()
   Dim subconsulta As String
   Dim client As String
   If cadbl(nomclient.tag) > 0 Then
     client = nomclient.tag
     subconsulta = "SELECT distinct marca FROM Clixes INNER JOIN Clientsvinculats ON Clixes.id_treball = Clientsvinculats.id_treball WHERE (Clientsvinculats.codiclient=" + client + ")"
   End If
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   If subconsulta <> "" Then
       formseleccio.Data1.RecordSource = subconsulta
      Else
       formseleccio.Data1.RecordSource = "select distinct marca from clixes order by marca"
   End If
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   'formseleccio.DBGrid2.Columns(0).Width = 0
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           marcaproducte = formseleccio.DBGrid2.Columns("marca")
           
        End If
   End If
    If seleccioret = 9 Then
        marcaproducte = ""
        
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub marcaproducte_KeyDown(KeyCode As Integer, Shift As Integer)
   marcaproducte.tag = "-1"
   'If liniaproducte.Tag = "-1" Then liniaproducte.Tag = "": liniaproducte = ""
End Sub

Private Sub nomclient_DropDown()
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select * from clients"
  formseleccio.refrescar
  formseleccio.width = 13000
  formseleccio.Show 1
  
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           nomclient = formseleccio.DBGrid2.Columns("nom")
           nomclient.tag = cadbl(formseleccio.DBGrid2.Columns("codi"))
        End If
   End If
    If seleccioret = 9 Then
        nomclient = ""
        nomclient.tag = "0"
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   SendKeys "{TAB}"
   'codimuntadora.SetFocus
End Sub
Function lordremesgran(treball As Long) As Long
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select max(ordre) as gran from modificacions group by id_treball=" + atrim(treball))
   If Not rst.EOF Then
        lordremesgran = cadbl(rst!gran)
          Else: lordremesgran = 0
   End If
End Function
Private Sub postit_DblClick()
    Dim generarfitxer_imp As String
    Dim nidtreball As Long
    Dim ordrem As Long
    Dim codiclient As Long
    Dim direnvio As String
    Dim idclientvinculat As Double
    Dim rst As Recordset
    If reixa.TextMatrix(0, reixa.col) <> "NOM_CLIENT" Then Exit Sub
    idclientvinculat = postit.ItemData(postit.ListIndex)
    nidtreball = reixa.TextMatrix(reixa.row, numcol("TREBALL"))
    'ordrem = lordremesgran(nidtreball)
    Set rst = dbclixes.OpenRecordset("select codiclient,direnvio,nomclient,ordremodificacio from clientsvinculats where id=" + atrim(idclientvinculat) + " order by nomclient")
    ordrem = cadbl(rst!ordremodificacio)
    If rst.EOF Then Exit Sub
    codiclient = cadbl(rst!codiclient)
    direnvio = cadbl(rst!direnvio)
    If edicioimps.Value = 1 Then
       While ordrem > 0
        generarfitxer_imp = ruta_documentacio_clixes + "\" + Format(nidtreball, "00000") + "\IMP" + Format(nidtreball, "00000") + "-" + Format(ordrem, "000") + "-" + Format(codiclient, "000000") + "_" + atrim(direnvio) + ".doc"
        'MsgBox generarfitxer_imp
        If existeix(generarfitxer_imp) Then
           obrir_document generarfitxer_imp
           Exit Sub
        End If
        ordrem = ordrem - 1
       Wend
       If ordrem = 0 Then MsgBox "No he trobat cap IMP per aquest client", vbCritical, "No trobat"
    End If
    postit.visible = False
End Sub
Function numcol(nom As String) As Byte
   numcol = 0
   For i = 0 To reixa.Cols - 1
     If reixa.TextMatrix(0, i) = nom Then numcol = i
   Next i
   
End Function

Private Sub reixa_Click()
'   rstconsulta.FindFirst "id_reball=" + atrim(cadbl(reixa.TextMatrix(reixa.row, 0)))
  possarpostit
End Sub
Sub possarpostit()
   carregarpostit
   posicionarelpostit
End Sub
Sub posicionarelpostit()
   If postit.ListCount = 0 Then postit.visible = False: Exit Sub
   postit.visible = True
   postit.Left = reixa.CellLeft + reixa.Left
   postit.Top = reixa.CellTop + reixa.CellHeight + reixa.Top
   postit.width = reixa.CellWidth
   
End Sub

Sub carregarpostit()
   Dim rst As Recordset
   Dim codidbarres As String
   postit.Clear
   If reixa.TextMatrix(0, reixa.col) = "CODIDEBARRES" Then
      codidbarres = atrim(Mid(reixa.Text, InStr(1, reixa.Text, ")") + 1))
      Set rst = dbclixes.OpenRecordset("select id_Treball from clixes where codidebarres='" + atrim(codidbarres) + "'")
      While Not rst.EOF
        postit.AddItem rst!id_treball
        rst.MoveNext
      Wend
      If postit.ListCount < 2 Then postit.Clear
   End If
   
   If reixa.TextMatrix(0, reixa.col) = "NOM_CLIENT" Then
      
      Set rst = dbclixes.OpenRecordset("select id,nomclient,direnvio,ordremodificacio from clientsvinculats where id_treball=" + atrim(reixa.TextMatrix(reixa.row, 0)) + IIf(agrupantpertreball.Value = 0, " and ordremodificacio=" + atrim(reixa.TextMatrix(reixa.row, 1)), "") + " order by ordremodificacio DESC,nomclient")
      While Not rst.EOF
        postit.AddItem "v" + atrim(rst!ordremodificacio) + " " + atrim(rst!direnvio) + " - " + rst!nomclient
        postit.ItemData(postit.NewIndex) = rst!ID
        rst.MoveNext
      Wend
   End If
   
   If reixa.TextMatrix(0, reixa.col) = "LOTS ACTIUS" Then
      
      rstcomandesactives.MoveFirst
      While Not rstcomandesactives.EOF And cadbl(reixa.TextMatrix(reixa.row, 0)) < rstcomandesactives!numtreball
       If rstcomandesactives!numtreball = cadbl(reixa.TextMatrix(reixa.row, 0)) Then
        v = v + " " + atrim(rstcomandesactives!comanda)
        cont = cont + 1
       End If
       rstcomandesactives.MoveNext
     Wend
   End If
End Sub

Private Sub reixa_DblClick()
   formclixes.clixes.Recordset.FindFirst "id_treball=" + atrim(cadbl(reixa.TextMatrix(reixa.row, numcol("TREBALL"))))
   fbuscar.Hide
End Sub

Private Sub reixa_RowColChange()
   postit.visible = False
End Sub

Private Sub sortir_Click()
  guardar_amples_reixa
 ' Unload fbuscar
 fbuscar.Hide
End Sub

Private Sub Text1_Change()

End Sub

