VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formllistatreferencies 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Llistat de referències per client"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9195
   Icon            =   "formllistatreferencies.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox capartirdata 
      Height          =   345
      Left            =   1335
      TabIndex        =   10
      Top             =   870
      Width           =   1485
   End
   Begin VB.Data datatarifa 
      Caption         =   "datatarifa"
      Connect         =   "Access"
      DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\COMANDES.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3870
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tarifapressupost"
      Top             =   1035
      Visible         =   0   'False
      Width           =   1740
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "formllistatreferencies.frx":058A
      Height          =   3615
      Left            =   2895
      OleObjectBlob   =   "formllistatreferencies.frx":059F
      TabIndex        =   9
      Top             =   855
      Width           =   6255
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   2385
      Picture         =   "formllistatreferencies.frx":12E0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Borrar fitxer sel.leccionat"
      Top             =   4920
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Preview del  Llistat"
      Height          =   435
      Left            =   5115
      TabIndex        =   7
      ToolTipText     =   "Només visualtzació no es gravarà"
      Top             =   4620
      Width           =   1740
   End
   Begin VB.FileListBox llistafitxers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   135
      Pattern         =   "Llist_Press_*.csv"
      TabIndex        =   5
      Top             =   2640
      Width           =   2670
   End
   Begin VB.TextBox crefpressupost 
      Height          =   345
      Left            =   1335
      TabIndex        =   3
      Top             =   480
      Width           =   1485
   End
   Begin VB.ComboBox cgrupclient 
      Height          =   315
      Left            =   1335
      TabIndex        =   2
      Top             =   105
      Width           =   7425
   End
   Begin VB.CommandButton bgenerarllistat 
      Caption         =   "Generar Llistat"
      Height          =   435
      Left            =   7050
      TabIndex        =   0
      ToolTipText     =   "El llistat es genererà i es gravarà"
      Top             =   4620
      Width           =   1740
   End
   Begin VB.Label Label4 
      Caption         =   "A partir data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   11
      Top             =   915
      Width           =   1140
   End
   Begin VB.Label Label3 
      Caption         =   "Doble clic per veure el fitxer"
      Height          =   285
      Left            =   135
      TabIndex        =   6
      Top             =   4920
      Width           =   2205
   End
   Begin VB.Label Label2 
      Caption         =   "Ref. Press:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   4
      Top             =   525
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Grup/Client:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Width           =   1140
   End
End
Attribute VB_Name = "formllistatreferencies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bgenerarllistat_Click()
  If MsgBox("Aquesta opció genera el llistat i el guarda." + Chr(10) + "Vols fer-ho?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   generarllistatreferenciesperclient True
   llistafitxers.Refresh
End Sub

Function extreusetmana(vdata As Date) As Byte
  extreusetmana = DatePart("ww", vdata, vbMonday, vbFirstFourDays)
End Function
Function buscadataentrega(numc As Double) As Date
   Dim rsttmp As Recordset
   
   Set rsttmp = dbplanificacio.OpenRecordset("select * from planificaciototes where comanda=" + atrim(numc))
   If Not rsttmp.EOF Then
        If Not IsNull(rsttmp!Data2) Then buscadataentrega = atrim(rsttmp!Data2)
   End If
   'Set dbplanificacio = Nothing
   Set rsttmp = Nothing
End Function
Sub possarmaterialsicolors(rsttemporal As Recordset, numc1 As Double, numc2 As Double, numc3 As Double)
  Dim rstmat1 As Recordset
  Dim rstmat2 As Recordset
  Dim rstmat3 As Recordset
  Dim espesormat1 As Double
  Dim descripciomat As String
  Dim tipusfilm As String
  Dim codimat As String
  Dim rstcomandes As Recordset
  Set rstcomandes = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(numc1) + " or comanda=" + atrim(numc2) + " or comanda=" + atrim(numc3))
  If rstcomandes.EOF Then Exit Sub
  rstcomandes.FindFirst "comanda=" + atrim(numc1)
  codimat = IIf(Not rstcomandes.NoMatch, cadbl(rstcomandes!materialex), 0)
  Set rstmat1 = dbtmp.OpenRecordset("SELECT familiesmaterials.descripcio, familiescolorants.descripcio, subfamiliesmaterials.descripcio FROM ((familiescolorants INNER JOIN materials ON familiescolorants.codi = materials.familiacol) INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi WHERE (((materials.codi)=" + atrim(codimat) + "));")
  rstcomandes.FindFirst "comanda=" + atrim(numc2)
  codimat = IIf(Not rstcomandes.NoMatch, cadbl(rstcomandes!materialex), 0)
  Set rstmat2 = dbtmp.OpenRecordset("SELECT familiesmaterials.descripcio, familiescolorants.descripcio, subfamiliesmaterials.descripcio FROM ((familiescolorants INNER JOIN materials ON familiescolorants.codi = materials.familiacol) INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi WHERE (((materials.codi)=" + atrim(codimat) + "));")
  rstcomandes.FindFirst "comanda=" + atrim(numc3)
  codimat = IIf(Not rstcomandes.NoMatch, cadbl(rstcomandes!materialex), 0)
  Set rstmat3 = dbtmp.OpenRecordset("SELECT familiesmaterials.descripcio, familiescolorants.descripcio, subfamiliesmaterials.descripcio FROM ((familiescolorants INNER JOIN materials ON familiescolorants.codi = materials.familiacol) INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi WHERE (((materials.codi)=" + atrim(codimat) + "));")
  If Not rstmat1.EOF Then
     rstcomandes.FindFirst "comanda=" + atrim(numc1)
     If Not rstcomandes.NoMatch Then
        descripciomat = atrim(micresmaterial(cadbl(rstcomandes!mesuraesp), cadbl(rstcomandes!espessor), atrim(rstcomandes!tubolam))) + " " + nomdescripciomat(atrim(rstmat1![familiesmaterials.descripcio]), atrim(rstmat1![familiescolorants.descripcio]))
        tipusfilm = atrim(rstmat1![familiescolorants.descripcio]) + IIf(InStr(1, rstmat1![subfamiliesmaterials.descripcio], "MATE "), " MATE", "")
        If descripciomat = "PE" And rstmat2.EOF And rstmat3.EOF Then descripciomat = "PE PHOTO"
        espesormat1 = micresmaterial(cadbl(rstcomandes!mesuraesp), cadbl(rstcomandes!espessor), atrim(rstcomandes!tubolam))
     End If
  End If
  If Not rstmat2.EOF Then
     rstcomandes.FindFirst "comanda=" + atrim(numc2)
     If Not rstcomandes.NoMatch Then
        descripciomat = descripciomat + " + " + atrim(micresmaterial(cadbl(rstcomandes!mesuraesp), cadbl(rstcomandes!espessor), atrim(rstcomandes!tubolam))) + " " + nomdescripciomat(atrim(rstmat2![familiesmaterials.descripcio]), atrim(rstmat1![familiescolorants.descripcio]))
        tipusfilm = tipusfilm + "/" + atrim(rstmat2![familiescolorants.descripcio]) + IIf(InStr(1, rstmat2![subfamiliesmaterials.descripcio], "MATE "), " MATE", "")
     End If
  End If
  If Not rstmat3.EOF Then
     rstcomandes.FindFirst "comanda=" + atrim(numc3)
     If Not rstcomandes.NoMatch Then
        descripciomat = descripciomat + " + " + atrim(micresmaterial(cadbl(rstcomandes!mesuraesp), cadbl(rstcomandes!espessor), atrim(rstcomandes!tubolam))) + " " + nomdescripciomat(atrim(rstmat3![familiesmaterials.descripcio]), atrim(rstmat1![familiescolorants.descripcio]))
        tipusfilm = tipusfilm + "/" + atrim(rstmat3![familiescolorants.descripcio]) + IIf(InStr(1, rstmat3![subfamiliesmaterials.descripcio], "MATE "), " MATE", "")
     End If
  End If
  rsttemporal!tipusfilm = evaluartipusfilm(tipusfilm)
  rsttemporal!nomestructura = nomdelestructura(rstmat1, rstmat2, rstmat3, espesormat1)
  rsttemporal!descripciomat = descripciomat
  Set rstmat1 = Nothing
  Set rstmat2 = Nothing
  Set rstmat3 = Nothing
  Set rstcomandes = Nothing
End Sub
Function nomdelestructura(rstmat1 As Recordset, rstmat2 As Recordset, rstmat3 As Recordset, espesormat1 As Double) As String
  If Not rstmat1.EOF Then
     nomdelestructura = atrim(nomdescripciomat(atrim(rstmat1![familiesmaterials.descripcio]), atrim(rstmat1![familiescolorants.descripcio]), True))
     If nomdelestructura = "PET" And espesormat1 = 8 Then nomdelestructura = "PET8"
  End If
  If Not rstmat2.EOF Then nomdelestructura = nomdelestructura + "+" + atrim(nomdescripciomat(atrim(rstmat2![familiesmaterials.descripcio]), atrim(rstmat2![familiescolorants.descripcio]), True))
  If Not rstmat3.EOF Then nomdelestructura = nomdelestructura + "+" + atrim(nomdescripciomat(atrim(rstmat3![familiesmaterials.descripcio]), atrim(rstmat3![familiescolorants.descripcio]), True))
End Function
Function evaluartipusfilm(tipusfilm As String) As String
  If InStr(1, tipusfilm, "BLANCO") Then evaluartipusfilm = "White"
  If InStr(1, tipusfilm, "MATE") Then evaluartipusfilm = IIf(evaluartipusfilm = "White", "Matt White", "Matt")
  If InStr(1, tipusfilm, "METAL") Then evaluartipusfilm = "Glossy"
  If InStr(1, tipusfilm, "TRANSP") Then evaluartipusfilm = IIf(evaluartipusfilm = "", "Transp", evaluartipusfilm)
End Function
Function nomdescripciomat(nommat As String, colormat As String, Optional nomestructura As Boolean) As String
  Dim vnom As String
  nommat = nommat + " "
  vnom = Mid(nommat, 1, InStr(1, nommat, " "))
  If atrim(vnom) = "PEBD" Then vnom = "PE" + IIf(InStr(1, atrim(nommat), "RIGID") > 0, " RIGID", "")
  vnom = atrim(vnom) + IIf(InStr(1, colormat, "METAL"), " MET", "")
  nomdescripciomat = vnom
  If nomestructura Then
     nomdescripciomat = nomdescripciomat + " "
     nomdescripciomat = atrim(Mid(nomdescripciomat, 1, InStr(1, nomdescripciomat, " ")))
     nomdescripciomat = nomdescripciomat + IIf(InStr(1, vnom, " MET") > 0, "MET", "")
  End If
End Function
Sub ensenyar_comandes_orfes(rst As Recordset)
   Dim msg As String
   While Not rst.EOF
      msg = msg + Trim(rst!comanda) + Chr(9)
      rst.MoveNext
   Wend
   If msg <> "" Then MsgBox "Les seguents comandes no tenen pressupost assignat:" + Chr(10) + msg, vbCritical, "Atenció"
End Sub
Sub generarllistatreferenciesperclient(gravarlo As Boolean)
  Dim dbtemporal As Database
  Dim taulatemp As String
  Dim rsttemporal As Recordset
  Dim nomfitxercsv As String
  Dim rst As Recordset
  If crefpressupost = "" Then MsgBox "No hi ha referencia de pressupost.", vbCritical, "Atenció": Exit Sub
  If Not IsDate(capartirdata) Then MsgBox "La data d'inici de llistat no es correcte", vbCritical, "Atenció": Exit Sub
  taulatemp = "c:\temp\temporal.mdb"
  ratoli "espera"
 ' Me.Caption = "Processant... "
  If Not existeix(taulatemp) Then DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
  On Error Resume Next
  Set dbtemporal = OpenDatabase(taulatemp)
  dbtemporal.Execute ("drop table relaciocomandes")
  dbtemporal.Execute ("create table relaciocomandes (nomestructura string,codipressupost byte, setmana string(10),id counter,datacomanda date,dataentrega date ,confirm string(1),comandaclient string,numcomanda double,repeticio string(3),tipusfilm string,descripciomat string,marca string,linia string,quantlinia string,colors byte,micromacro string(2),reimpres string(1),refclient string,codibarres string,fabricaentregada string,kgentregats double,euros double)")
  On Error GoTo 0
  Set rsttemporal = dbtemporal.OpenRecordset("select * from relaciocomandes")
 
  'cgrupclient.Tag = "7035"
  Set rst = dbtmp.OpenRecordset("Select *  FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (numtreball>0 and ((InStr(1,[ruta],'I'))>0) and client IN (" + cgrupclient.Tag + ") and (trim(numpressupost)='' or numpressupost=null)) and datacomanda>#" + Format(capartirdata, "mm/dd/yy") + "#")
  ensenyar_comandes_orfes rst
  Set rst = dbtmp.OpenRecordset("Select *  FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (numtreball>0 and ((InStr(1,[ruta],'I'))>0)and client IN (" + cgrupclient.Tag + ") and trim(numpressupost)='" + crefpressupost + "') and datacomanda>#" + Format(capartirdata, "mm/dd/yy") + "#")
  Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificacio.mdb")
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  While Not rst.EOF
     If cadbl(rst!numtreball) > 0 Then
      rsttemporal.AddNew
      If Not IsNull(rst!datamaterial) Then rsttemporal!setmana = "week " + atrim(extreusetmana(rst!datamaterial))
      rsttemporal!dataentrega = buscadataentrega(rst!comanda)
      rsttemporal!repeticio = IIf(atrim(rst!impressio) = "R", "Yes", "No")
      possarmaterialsicolors rsttemporal, rst!comanda, cadbl(rst!linkcomanda1), cadbl(rst!linkcomanda2)
      rsttemporal!refclient = atrim(rst!refclient) + " "
      possarliniamarcaiquanlinia cadbl(rst!numtreball), rst!numordremodificacio, rsttemporal
      rsttemporal!numcomanda = rst!comanda
      rsttemporal!datacomanda = rst!datacomanda
      rsttemporal!confirm = "X"
      rsttemporal!fabricaentregada = poblaciodirenvio(rst!direnvio)
      rsttemporal!kgentregats = cadbl(rst!tubbaseext)
      rsttemporal!comandaclient = atrim(rst!comandaclient)
      rsttemporal!codipressupost = buscarcodipressupost(rsttemporal!nomestructura)
      rsttemporal!micromacro = IIf(atrim(rst!microperforat) = "S" Or atrim(rst!microperforatsol) = "S", "Mi", "")
      If atrim(rst!rebmacroperforat) = "S" Then rsttemporal!micromacro = "Ma"
      rsttemporal.Update
      rst.MoveNext
     End If
  Wend
  ratoli "normal"
  Set rsttemporal = dbtemporal.OpenRecordset("select * from relaciocomandes")
  If rsttemporal.EOF Then Exit Sub
  nomfitxercsv = IIf(gravarlo, nomdelllistat(crefpressupost), "c:\temp\Llistatreferenciestemp.csv")
  passardemdbacsv dbtemporal, nomfitxercsv
  If existeix(nomfitxercsv) Then Shell "c:\windows\system32\cmd.exe /c start " + nomfitxercsv
  Set dbclixes = Nothing
  Set dbplanificacio = Nothing
End Sub
Function micresmaterial(codimesuralineal As Byte, espesor As Double, tubolam As String) As String
  Dim rstmesural As Recordset
  Dim descripcio As String
 ' Dim r As String
  Set rstmesural = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(codimesuralineal))
  If rstmesural.EOF Then Exit Function
  descripcio = rstmesural!descripcio
  r = espesor
  If descripcio = "GALGUES" Then
            If tubolam = "T" Then
                 r = espesor / 4
                  Else: r = espesor / 2
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

Function buscarcodipressupost(descmat As String) As Byte
    buscarcodipressupost = 0
    datatarifa.Recordset.FindFirst "estructura like '*" + atrim(descmat) + "*'"
    If Not datatarifa.Recordset.NoMatch Then buscarcodipressupost = cadbl(datatarifa.Recordset!codipressupost)
End Function
Sub possarcapcaleracsv(linia As String, rst As Recordset, rstenvios As Recordset)
  Dim i As Byte
  For i = 0 To rst.Fields.Count - 1
    If atrim(rst.Fields(i).Name) <> "euros" And atrim(rst.Fields(i).Name) <> "fabricaentregada" And atrim(rst.Fields(i).Name) <> "kgentregats" Then
        linia = linia + IIf(linia <> "", ";", "Nº ") + UCase(atrim(rst.Fields(i).Name))
    End If
  Next i
  rstenvios.MoveFirst
  While Not rstenvios.EOF
    linia = linia + IIf(linia <> "", ";", "") + UCase(atrim(rstenvios!fabricaentregada))
    rstenvios.MoveNext
  Wend
End Sub
Sub passardemdbacsv(dbtemporal As Database, nomdelllistat As String)
  Dim rst As Recordset
  Dim rstenvios As Recordset
  Dim linia As String
  Set rstenvios = dbtemporal.OpenRecordset("select distinct fabricaentregada from relaciocomandes order by fabricaentregada")
  Set rst = dbtemporal.OpenRecordset("select * from relaciocomandes order by id")
  Open nomdelllistat For Output As #1
  If Not rst.EOF Then
     possarcapcaleracsv linia, rst, rstenvios
     Print #1, linia
  End If
  While Not rst.EOF
     linia = ""
     For i = 0 To rst.Fields.Count - 1
       If rst.Fields(i).Name <> "kgentregats" Then
        If rst.Fields(i).Name <> "fabricaentregada" Then
         linia = linia + IIf(linia <> "", ";", "") + atrim(rst.Fields(i))
           Else: possarelskgentregatsalafabrica linia, rst, rstenvios
        End If
       End If
     Next i
     Print #1, linia
     rst.MoveNext
  Wend
  Close 1
End Sub
Sub possarelskgentregatsalafabrica(linia As String, rst As Recordset, rstenvios As Recordset)
  rstenvios.MoveFirst
  While Not rstenvios.EOF
     linia = linia + ";"
     If rst!fabricaentregada = rstenvios!fabricaentregada Then linia = linia + atrim(rst!kgentregats)
     rstenvios.MoveNext
  Wend
End Sub
Function nomdelllistat(vref As String) As String
  Dim d As String
  Dim c As String
  nomdelllistat = rutadelfitxer(cami) + "Llistats_referencies\"
  d = Dir(nomdelllistat + vreg + "*.csv")
  While d <> ""
    If InStr(1, d, vref) > 0 Then c = Mid(d, InStr(1, d, "-") + 1, 3)
    d = Dir
  Wend
  nomdelllistat = nomdelllistat + vref + "-" + Format(cadbl(c) + 1, "00") + ".csv"
End Function
Function poblaciodirenvio(direnvio As Long) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select poblacioe from clients_envios where id=" + atrim(direnvio))
   If Not rst.EOF Then poblaciodirenvio = UCase(rst!poblacioe)
   Set rst = Nothing
End Function
Sub possarliniamarcaiquanlinia(treball As Double, versio As Integer, rsttemporal As Recordset)
   Dim rst As Recordset
   Dim rstv As Recordset
   Dim marca As String
   Dim linia As String
   Dim quant As String
   Dim codibarres As String
   Dim colors As Byte
   
   Set rst = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(treball))
   Set rstv = dbclixes.OpenRecordset("select tinters,reimpres from modificacions where id_treball=" + atrim(treball) + " and ordre=" + atrim(versio))
   If rst.EOF Then Exit Sub
   marca = atrim(rst!marca)
   linia = atrim(rst!linia)
   quant = atrim(rst!descripcioquantitatlinia)
   codibarres = atrim(rst!codidebarres) + " "
   If Not rstv.EOF Then colors = rstv!tinters
   Set rst = Nothing
   'Set dbclixes = Nothing
   rsttemporal!marca = marca
   rsttemporal!linia = linia
   rsttemporal!quantlinia = quant
   rsttemporal!codibarres = codibarres
   rsttemporal!colors = colors
   rsttemporal!reimpres = IIf(rstv!reimpres, "S", "N")
End Sub

Private Sub capartirdata_Change()
    escriure_ini "LlistesReferencies", "apartirdata", capartirdata, fitxerini
End Sub

Private Sub cgrupclient_Click()
  If cgrupclient.Text = "  << Nou Grup >>" Then nougrup: Exit Sub
  If cgrupclient.Text = "   << Eliminar Grup >>" Then eliminargrup cgrupclient.ToolTipText: Exit Sub
  cgrupclient.Tag = Mid(cgrupclient, InStr(1, cgrupclient, ":") + 1)
  cgrupclient.ToolTipText = UCase(Mid(cgrupclient, 1, InStr(1, cgrupclient, ":")))
  datatarifa.RecordSource = "select * from tarifapressupost where gruptarifa='" + cgrupclient.ToolTipText + "' order by codipressupost"
  datatarifa.Refresh
End Sub
Sub eliminargrup(nomdelgrup)
  Dim i As Byte
  Dim t As String
   If MsgBox("Segur que vols eliminar el grup " + nomdelgrup, vbInformation + vbYesNo, "Eliminar grup") Then
       For i = 1 To 100
         t = llegir_ini("LlistesReferencies", "Ref" + atrim(i), fitxerini)
         If UCase(Mid(t, 1, InStr(1, t, ":"))) = nomdelgrup Then
            escriure_ini "LlistesReferencies", "Ref" + atrim(i), "", fitxerini
            GoTo fi
         End If
       Next i
   End If
fi:
End Sub
Sub nougrup()
   Dim n As String
   Dim t As String
   Dim c As String
   Dim i As Byte
   n = InputBox("Entra el nom com vols que es digui el nou grup" + Chr(10) + " Ex: Grup Ardo")
   If atrim(n) = "" Then Exit Sub
   c = InputBox("Entra els codis de client separats per comes" + Chr(10) + " Ex: 1234,4321,5678,8765")
   If atrim(c) = "" Then Exit Sub
   t = " "
   i = 1
   While t <> "{[}]" And t <> ""
     t = llegir_ini("LlistesReferencies", "Ref" + atrim(i), fitxerini)
     If t <> "{[}]" And t <> "" Then i = i + 1
   Wend
   escriure_ini "LlistesReferencies", "Ref" + atrim(i), atrim(n) + ": " + atrim(c), fitxerini
End Sub
Private Sub cgrupclient_DropDown()
   carregar_llistes_alcombo
End Sub
Sub carregar_llistes_alcombo()
   Dim t As String
   t = cgrupclient
   cgrupclient.Clear
   cgrupclient = t
   For i = 1 To 100
     t = llegir_ini("LlistesReferencies", "Ref" + atrim(i), fitxerini)
     If t <> "{[}]" Then
       cgrupclient.AddItem t
     End If
   Next i
   cgrupclient.AddItem "  << Nou Grup >>"
   If cgrupclient <> "" Then cgrupclient.AddItem "   << Eliminar Grup >>"
End Sub

Private Sub Command1_Click()
generarllistatreferenciesperclient False
End Sub

Private Sub Command2_Click()
  Dim nomfitxer As String
  datatarifa.RecordSource = "select * from tarifapressupost where gruptarifa='_'"
  datatarifa.Refresh
  If llistafitxers.ListIndex = -1 Then MsgBox "Primer has d'escullir el fitxer que vols borrar", vbCritical, "Error": Exit Sub
  nomfitxer = rutadelfitxer(cami) + "Llistats_referencies\" + llistafitxers
  If MsgBox("Segur que vols borrar el fitxer " + llistafitxers + "?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
  Kill nomfitxer
  wait 2
  llistafitxers.Refresh
End Sub

Private Sub crefpressupost_Change()
   llistafitxers.Pattern = crefpressupost + "*.csv"
   llistafitxers.Refresh
   escriure_ini "LlistesReferencies", "ultimaref", crefpressupost, fitxerini
End Sub

Private Sub DBGrid1_Click()

End Sub

Private Sub DBGrid1_OnAddNew()

End Sub

Private Sub Form_Load()
   llistafitxers.path = rutadelfitxer(cami) + "Llistats_referencies"
   cgrupclient = llegir_ini("LlistesReferencies", "Ref1", fitxerini)
   cgrupclient_Click
   crefpressupost = llegir_ini("LlistesReferencies", "ultimaref", fitxerini)
   capartirdata = llegir_ini("LlistesReferencies", "apartirdata", fitxerini)
   If crefpressupost = "{[}]" Then crefpressupost = ""
   If capartirdata = "{[}]" Then capartirdata = Format(Now, "dd/mm/yy")
End Sub

Private Sub llistafitxers_DblClick()
 Dim nomfitxer As String
  nomfitxer = rutadelfitxer(cami) + "Llistats_referencies\" + llistafitxers
  If existeix(nomfitxer) Then Shell "c:\windows\system32\cmd.exe /c " + nomfitxer
End Sub

Private Sub reixa_OnAddNew()
  If cgrupclient.ToolTipText = "" Then datatarifa.Recordset.CancelUpdate: Exit Sub
  datatarifa.Recordset!gruptarifa = cgrupclient.ToolTipText
End Sub
