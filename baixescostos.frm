VERSION 5.00
Begin VB.Form baixescostos 
   Caption         =   "Baixes Costos"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5235
   Icon            =   "baixescostos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   585
      Left            =   30
      TabIndex        =   16
      Top             =   3465
      Width           =   5160
      Begin VB.CommandButton parar 
         BackColor       =   &H008080FF&
         Caption         =   "Parar"
         Height          =   300
         Left            =   4530
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   180
         Width           =   555
      End
      Begin VB.Label processant 
         Height          =   390
         Left            =   105
         TabIndex        =   17
         Top             =   105
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtrar per comanda"
      Height          =   1320
      Left            =   3435
      TabIndex        =   13
      Top             =   405
      Width           =   1665
      Begin VB.TextBox numcomanda 
         Height          =   300
         Left            =   255
         TabIndex        =   14
         Top             =   705
         Width           =   1170
      End
      Begin VB.Label Label6 
         Caption         =   "Comanda:"
         Height          =   330
         Left            =   480
         TabIndex        =   15
         Top             =   405
         Width           =   840
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Llistar"
      Height          =   495
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2940
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtre Multiple"
      Height          =   2490
      Left            =   195
      TabIndex        =   0
      Top             =   390
      Width           =   3150
      Begin VB.TextBox codiproducte 
         Height          =   300
         Left            =   765
         TabIndex        =   10
         Top             =   1950
         Width           =   450
      End
      Begin VB.ComboBox impresora 
         Height          =   315
         ItemData        =   "baixescostos.frx":058A
         Left            =   855
         List            =   "baixescostos.frx":0597
         TabIndex        =   9
         Text            =   "Totes"
         Top             =   1110
         Width           =   2085
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   600
         TabIndex        =   6
         Top             =   1560
         Width           =   2340
      End
      Begin VB.TextBox datafi 
         Height          =   345
         Left            =   1680
         TabIndex        =   3
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox datainici 
         Height          =   345
         Left            =   255
         TabIndex        =   1
         Top             =   570
         Width           =   1170
      End
      Begin VB.Label descripcioproducte 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1260
         TabIndex        =   12
         Top             =   1935
         Width           =   1845
      End
      Begin VB.Label Label5 
         Caption         =   "Producte:"
         Height          =   330
         Left            =   60
         TabIndex        =   11
         Top             =   2010
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Impresora:"
         Height          =   330
         Left            =   45
         TabIndex        =   8
         Top             =   1155
         Width           =   810
      End
      Begin VB.Label Label3 
         Caption         =   "Client:"
         Height          =   330
         Left            =   60
         TabIndex        =   7
         Top             =   1620
         Width           =   510
      End
      Begin VB.Label Label2 
         Caption         =   "Data Fi"
         Height          =   225
         Left            =   1890
         TabIndex        =   4
         Top             =   285
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inici"
         Height          =   225
         Left            =   465
         TabIndex        =   2
         Top             =   285
         Width           =   1035
      End
   End
End
Attribute VB_Name = "baixescostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub codiproducte_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then triarproducte
End Sub

Private Sub Command1_Click()
  Dim bdtemporal As String
  Dim numc As Double
  Dim rstcom As Recordset
  Dim rstfiltre As Recordset
  Dim afegir As Boolean
  Dim afegides As Double
  Set dbconsulta = Nothing
  bdtemporal = "c:\baixescostos.mdb"
  If Not eliminar_basededades(bdtemporal) Then MsgBox "No es pot obrir la base de dades, si està oberta tanqueu-la sisplau." + Chr(10) + Chr(13) + bdtemporal: Exit Sub
  DBEngine.CreateDatabase bdtemporal, dbLangGeneral, DatabaseTypeEnum.dbVersion30
  Set dbconsulta = OpenDatabase(bdtemporal)
  Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", fitxerini))
  crear_taules_tmp
  obrestocks
  Set rstconsulta = dbconsulta.OpenRecordset("costos")
  afegides = 0
  If cadbl(numcomanda) > 0 Then
    numc = cadbl(numcomanda)
    Set rsttmp = dbtmp.OpenRecordset("select * from comandes where proximaseccio='T' and comanda=" + atrim(numc))
    If Not rsttmp.EOF Then afegir_comanda cadbl(numc): afegides = afegides + 1
      Else
        filtrarregistres rstfiltre
        
        While Not rstfiltre.EOF And parar.Tag <> "1"
         afegir = False
         numc = cadbl(rstfiltre!comanda)
         processant = "Processant la comanda: " + atrim(numc) + "   ---> " + atrim(rstfiltre.AbsolutePosition + 1) + " de " + atrim(rstfiltre.RecordCount) + " Comandes (sense filtrar)."
         DoEvents
         afegir = False
         Set rsttmp = dbtmp.OpenRecordset("select producte,client,proximaseccio from comandes where  comanda=" + atrim(rstfiltre!comanda))
         If Not rsttmp.EOF Then
           If rsttmp!proximaseccio = "T" Then
            If codiproducte <> "" And Text3.Tag <> "" Then If rsttmp!producte = codiproducte And Text3.Tag = rsttmp!client Then afegir = True
            If codiproducte <> "" And Text3.Tag = "" Then If rsttmp!producte = codiproducte Then afegir = True
            If codiproducte = "" And Text3.Tag <> "" Then If Text3.Tag = rsttmp!client Then afegir = True
            If codiproducte = "" And Text3.Tag = "" Then afegir = True
           End If
         End If
         If afegir Then afegir_comanda cadbl(numc): afegides = afegides + 1
         rstfiltre.MoveNext
        Wend
        parar.Tag = ""
        
  End If
  dbstocks.Close
  dbconsulta.Close
  Set dbstocks = Nothing
  Set rsttmp = Nothing
  Set rstconsulta = Nothing
  Set dbconsulta = Nothing
  If afegides > 0 Then
           If MsgBox("Procés acavat... s'han afegit " + atrim(afegides) + " comandes." + Chr(10) + Chr(13) + "Vols obrir el fitxer creat?", vbYesNo, "Analisis de costos") = vbYes Then
                 DoEvents
                 obrir_document Chr$(34) + bdtemporal + Chr$(34)
                 DoEvents
           End If
        End If
  processant = "Procés acavat... " + atrim(afegides) + " comandes afegides." + Chr(10) + Chr(13) + " Fitxer creat a " + bdtemporal
End Sub
Sub filtrarregistres(rstfiltre As Recordset)
   Dim aon As String
   Dim filtredates As String
   Dim filtremaquina As String
   processant = "Creant la consulta de registres..."
   DoEvents
   Set rstfiltre = dbtmp.OpenRecordset("select * from comandes where comanda=-1")
   If Not IsDate(datainici) Or Not IsDate(datafi) Then MsgBox "Les dates no son correctes": Exit Sub
   filtredates = " (first(impressores.datainici)>=#" + Format(datainici, "mm/dd/yy") + "# and first(impressores.datainici)<=#" + Format(datafi, "mm/dd/yy") + "#) "
   If impresora.ListIndex > 0 Then
      filtremaquina = "and first(impressores.numeromaquina= " + atrim(impresora.ItemData(impresora.ListIndex)) + ")"
     Else:    filtremaquina = " and first(impressores.numeromaquina)"
   End If
   'If cadbl(Text3.Tag) > 0 Then aon = aon + " and client=" + atrim(Text3.Tag)
   'If codiproducte <> "" Then aon = aon + " and producte='" + codiproducte + "'"
   'aon = aon + " and producte<>'PC' and producte<>'PC2' and proximaseccio='T' "
   Set rstfiltre = dbbaixes.OpenRecordset("SELECT DISTINCT impressores.comanda, first(impressores.datainici), first(impressores.numeromaquina) From impressores GROUP BY impressores.comanda, impressores.numeromaquina HAVING ((" + filtredates + ") " + filtremaquina + ");")

   'Set rstfiltre = dbtmp.OpenRecordset("select * from comandes where " + aon + " order by comanda")
   If Not rstfiltre.EOF Then
      rstfiltre.MoveLast: rstfiltre.MoveFirst
      'MsgBox "Registres trobats: " + atrim(rstfiltre.RecordCount)
   End If
   processant = "Consulta creada..."
   DoEvents
End Sub
Sub sumar_totals_bobines_entregades(numc As Double, ventregatm As Double, ventregatk As Double)
  Dim vpendentm As Double

  Dim vpendentk As Double
  Dim rsttmpt As Recordset
  Set rsttmpt = dbbaixes.OpenRecordset("select metresisacs,data,kilosiunitats from bobinesent where comanda=" + atrim(numc))
 ' rsttmpt.MoveLast
  While Not rsttmpt.EOF
    
     If rsttmpt!Data <> "" Then
       ventregatm = ventregatm + cadbl(rsttmpt!metresisacs)
       ventregatk = ventregatk + cadbl(rsttmpt!kilosiunitats)
      Else:
         vpendentm = vpendentm + cadbl(rsttmpt!metresisacs)
         vpendentk = vpendentk + cadbl(rsttmpt!kilosiunitats)
    End If
    rsttmpt.MoveNext
  Wend
  If vpendentm > 0 Then ventregatm = 0: ventregatk = 0
 ' entregatm = Format(ventregatm, "#,##0.00")
 ' pendentm = Format(vpendentm, "#,##0.00")
 ' entregatk = Format(ventregatk, "#,##0.00")
 ' pendentk = Format(vpendentk, "#,##0.00")
  
  Set rsttmpt = Nothing
End Sub

Sub afegir_comanda(numc As Double)
   Dim rstimp As Recordset
   rstconsulta.AddNew
   dadescapcalera numc
   dadesimpresores numc
   dadeslaminadores numc
   dadesrebobinadores numc
   dadesentrega numc
   rstconsulta.Update
End Sub

Sub dadesentrega(numc As Double)
   Dim rstent As Recordset
   Dim ventregatm As Double
   Dim ventregatk As Double
   Dim total As Double
   Dim rstbobreb As Recordset
   Dim unitats As Double
   Dim preumaterial As Double
   sumar_totals_bobines_entregades numc, ventregatm, ventregatk
   If ventregatm = 0 Then Exit Sub
   rstconsulta!entrega_kilosentregats = ventregatk
   rstconsulta!entrega_metresentregats = ventregatm
   Set rstent = dbtmp.OpenRecordset("SELECT comanda, pvp, descripcio FROM comandes INNER JOIN mesures ON comandes.mesurapvp = mesures.codi WHERE (((comandes.comanda)=" + atrim(numc) + "));")
   
   If Not rstent.EOF Then
     rstconsulta!entrega_unitatfacturada = rstent!descripcio: rstconsulta!entrega_eurosxunitat = rstent!pvp
     If InStr(1, rstent!descripcio, "KGR") > 0 Then total = ventregatk * cadbl(rstent!pvp)
     If InStr(1, rstent!descripcio, "KM") > 0 Then total = (ventregatk / 1000) * cadbl(rstent!pvp)
     If InStr(1, rstent!descripcio, "LINEAL") > 0 Then total = ventregatm * cadbl(rstent!pvp)
     If InStr(1, rstent!descripcio, "QUADRAT") > 0 Then
        Set rstbobreb = dbbaixes.OpenRecordset("SELECT ample FROM rebobinadores INNER JOIN bobinesreb ON rebobinadores.Id = bobinesreb.controlid WHERE (rebobinadores.comanda=" + atrim(numc) + ") AND (rebobinadores.tipus='F');")
        total = ventregatm * (cadbl(rstbobreb!ample) / 100) * cadbl(rstent!pvp)
     End If
     If InStr(1, rstent!descripcio, "UNITAT") > 0 Then
        Set rstbobreb = dbtmp.OpenRecordset("select dessarroll from comandes where comanda=" + atrim(numc))
        If Not rstbobreb.EOF Then
         If cadbl(rstbobreb!dessarroll) > 0 Then unitats = Fix(ventregatm / (cadbl(rstbobreb!dessarroll) / 1000))
         total = unitats * cadbl(rstent!pvp)
        End If
     End If
     rstconsulta!entrega_pvpcomanda = total
   End If
   With rstconsulta
   preumaterials = (cadbl(!kilos_lot1) * cadbl(!preumaterial_lot1)) + (cadbl(!kilos_lot2) * cadbl(!preumaterial_lot2)) + (cadbl(!kilos_lot3) * cadbl(!preumaterial_lot3))
   total = preumaterials + cadbl(!impresora_cost) + cadbl(!impresora_costtintes) + cadbl(!impresora_costdisolvents)
   total = total + cadbl(!laminadora1_cost) + cadbl(!laminadora1_costadhesius) + cadbl(!laminadora2_cost) + cadbl(!laminadora2_costadhesius)
   total = total + cadbl(!rebobinadora_cost)
   End With
   rstconsulta!entrega_costcomanda = total
   Set rstbobreb = Nothing
   Set rstent = Nothing
End Sub


Sub dadesrebobinadores(numc As Double)
   Dim rstreb As Recordset
   Set rstreb = dbbaixes.OpenRecordset("SELECT comanda,numeromaquina, Sum(totalmetres) AS sumatotalmetres,  Sum(totalhores) AS sumatotalhores From rebobinadores where (tipus = 'C' ) GROUP BY comanda,numeromaquina HAVING (((comanda)=" + atrim(numc) + "));")
   If Not rstreb.EOF Then rstconsulta!rebobinadora_tempscanvi = rstreb!sumatotalhores: rstconsulta!rebobinadora_numrebobinadora = rstreb!numeromaquina
   Set rstreb = dbbaixes.OpenRecordset("SELECT comanda,  Sum(totalmetres) AS sumatotalmetres, Avg(metresminut) AS avgmtrsminut, Sum(totalhores) AS sumatotalhores From rebobinadores where (((tipus) = 'F')) GROUP BY comanda HAVING (((comanda)=" + atrim(numc) + "));")
   If Not rstreb.EOF Then rstconsulta!rebobinadora_tempsfunc = rstreb!sumatotalhores: rstconsulta!rebobinadora_mtrsmin = rstreb!avgmtrsminut
   rstconsulta!rebobinadora_cost = 35 * (rstconsulta!rebobinadora_tempsfunc + rstconsulta!rebobinadora_tempscanvi)
End Sub
Sub dadeslaminadores(numc As Double)
   Dim rstlam As Recordset
   Dim rstcom As Recordset
   Dim comandalam2 As Double
   Dim amplebob As Double
   Dim n As Double
   Set rstcom = dbtmp.OpenRecordset("select linkcomanda2 from comandes where comanda=" + atrim(numc))
   If Not rstcom.EOF Then comandalam2 = cadbl(rstcom!linkcomanda2)
   Set rstcom = Nothing
   Set rstlam = dbbaixes.OpenRecordset("SELECT palet FROM bobinesentlam INNER JOIN (laminadores INNER JOIN bobineslam ON laminadores.Id = bobineslam.controlid) ON bobinesentlam.id = bobineslam.Id WHERE (((laminadores.comanda)=" + atrim(numc) + ") AND ((bobinesentlam.paletobobina)='p' Or (bobinesentlam.paletobobina)='P'));")
   If Not rstlam.EOF Then
     Set rststocks = dbstocks.OpenRecordset("select ample from palets where idpalet=" + atrim(rstlam!palet))
     If Not rststocks.EOF Then amplebob = (rststocks!ample / 1000)
   End If
   Set rstlam = dbbaixes.OpenRecordset("SELECT comanda, Sum(totalmetres) AS sumatotalmetres, Avg(mtrsminut) AS avgmtrsminut, Sum(totalhores) AS sumatotalhores From laminadores where (tipus = 'C' ) GROUP BY comanda HAVING (((comanda)=" + atrim(numc) + "));")
   If Not rstlam.EOF Then rstconsulta!laminadora1_tempscanvi = rstlam!sumatotalhores
   Set rstlam = dbbaixes.OpenRecordset("SELECT comanda,  Sum(totalmetres) AS sumatotalmetres, Avg(mtrsminut) AS avgmtrsminut, Sum(totalhores) AS sumatotalhores From laminadores where (((tipus) = 'F')) GROUP BY comanda HAVING (((comanda)=" + atrim(numc) + "));")
   If Not rstlam.EOF Then
      rstconsulta!laminadora1_tempsfunc = rstlam!sumatotalhores: rstconsulta!laminadora1_mtrsmin = rstlam!avgmtrsminut
      rstconsulta!metres_lot2 = rstlam!sumatotalmetres
      rstconsulta!kilos_lot2 = rstlam!sumatotalmetres * (rstconsulta!lot2_factorpesxmil * amplebob)
   End If
   Set rstlam = dbbaixes.OpenRecordset("select (kg1+kg2) as kiloscola from laminadoresadhesius  where comanda=" + atrim(numc))
   If Not rstlam.EOF Then
     rstconsulta!laminadora1_totaladhesius = rstlam!kiloscola
     rstconsulta!laminadora1_costadhesius = cadbl(rstconsulta!laminadora1_totaladhesius) * 3.44
   End If
   rstconsulta!laminadora1_cost = 35 * (rstconsulta!laminadora1_tempsfunc + rstconsulta!laminadora1_tempscanvi)
   
   
   'segon proces de laminació
   If comandalam2 > 0 Then
    
    Set rstlam = dbbaixes.OpenRecordset("SELECT palet FROM bobinesentlam INNER JOIN (laminadores INNER JOIN bobineslam ON laminadores.Id = bobineslam.controlid) ON bobinesentlam.id = bobineslam.Id WHERE (((laminadores.comanda)=" + atrim(comandalam2) + ") AND ((bobinesentlam.paletobobina)='p' Or (bobinesentlam.paletobobina)='P'));")
    If Not rstlam.EOF Then
      Set rststocks = dbstocks.OpenRecordset("select ample from palets where idpalet=" + atrim(rstlam!palet))
      If Not rststocks.EOF Then amplebob = (rststocks!ample / 1000)
    End If
   
    Set rstlam = dbbaixes.OpenRecordset("SELECT comanda,  Sum(totalmetres) AS sumatotalmetres, Avg(mtrsminut) AS avgmtrsminut, Sum(totalhores) AS sumatotalhores From laminadores where (tipus = 'C' ) GROUP BY comanda HAVING (((comanda)=" + atrim(comandalam2) + "));")
    If Not rstlam.EOF Then rstconsulta!laminadora2_tempscanvi = rstlam!sumatotalhores
    Set rstlam = dbbaixes.OpenRecordset("SELECT comanda,  Sum(totalmetres) AS sumatotalmetres, Avg(mtrsminut) AS avgmtrsminut, Sum(totalhores) AS sumatotalhores From laminadores where (tipus = 'F') GROUP BY comanda HAVING (((comanda)=" + atrim(comandalam2) + "));")
    If Not rstlam.EOF Then
      rstconsulta!laminadora2_tempsfunc = rstlam!sumatotalhores: rstconsulta!laminadora2_mtrsmin = rstlam!avgmtrsminut
      rstconsulta!metres_lot3 = rstlam!sumatotalmetres
      rstconsulta!kilos_lot3 = rstlam!sumatotalmetres * (rstconsulta!lot3_factorpesxmil * amplebob)
    End If
    Set rstlam = dbbaixes.OpenRecordset("select (kg1+kg2) as kiloscola from laminadoresadhesius  where comanda=" + atrim(numc))
    
    If Not rstlam.EOF Then rstconsulta!laminadora1_totaladhesius = rstlam!kiloscola
    rstconsulta!laminadora2_costadhesius = cadbl(rstconsulta!laminadora2_totaladhesius) * 3.44
    rstconsulta!laminadora2_cost = 35 * (rstconsulta!laminadora2_tempsfunc + rstconsulta!laminadora2_tempscanvi)
   End If
End Sub
Sub dadesimpresores(numc As Double)
   Dim rstimp As Recordset
   Dim amplebob As Double
   Dim n As Double
   
   Set rstimp = dbbaixes.OpenRecordset("SELECT palet FROM bobinesentimp INNER JOIN (impressores INNER JOIN bobinesimp ON impressores.Id = bobinesimp.controlid) ON bobinesentimp.id = bobinesimp.Id WHERE (((impressores.comanda)=" + atrim(numc) + ") );")
    If Not rstimp.EOF Then
      Set rststocks = dbstocks.OpenRecordset("select ample from palets where idpalet=" + atrim(rstimp!palet))
      If Not rststocks.EOF Then amplebob = (rststocks!ample / 1000)
    End If
   
   Set rstimp = dbbaixes.OpenRecordset("SELECT comanda,numeromaquina, Sum(mtrsprova) AS totalmtrsprova, Sum(totalmetres) AS sumatotalmetres, Avg(mtrsminut) AS avgmtrsminut, Sum(totalhores) AS sumatotalhores From impressores where (((tipus) = 'A' Or (tipus) = 'M')) GROUP BY comanda,numeromaquina HAVING (((comanda)=" + atrim(numc) + "));")
   If Not rstimp.EOF Then rstconsulta!impresora_mtrsprova = rstimp!totalmtrsprova: rstconsulta!impresora_tempscanvi = rstimp!sumatotalhores: rstconsulta!impresora_numimpresora = cadbl(rstimp!numeromaquina)
   Set rstimp = dbbaixes.OpenRecordset("SELECT comanda,numeromaquina, Sum(mtrsprova) AS totalmtrsprova, Sum(totalmetres) AS sumatotalmetres, Avg(mtrsminut) AS avgmtrsminut, Sum(totalhores) AS sumatotalhores From impressores where (((tipus) = 'F')) GROUP BY comanda,numeromaquina HAVING (((comanda)=" + atrim(numc) + "));")
   If Not rstimp.EOF Then
      rstconsulta!impresora_tempsfunc = rstimp!sumatotalhores
      rstconsulta!impresora_mtrsmin = rstimp!avgmtrsminut
      rstconsulta!impresora_metresfunc = rstimp!sumatotalmetres
      rstconsulta!metres_lot1 = rstconsulta!impresora_metresfunc + rstconsulta!impresora_mtrsprova
      rstconsulta!kilos_lot1 = rstconsulta!metres_lot1 * (rstconsulta!lot1_factorpesxmil * amplebob)
   End If
   Set rstimp = dbbaixes.OpenRecordset("select kg1,kg2,kg3,kg4,kg5,kg6,kg7,kg8,kg9,kg10  from impresorespantones where comanda=" + atrim(numc))
   If Not rstimp.EOF Then
     With rstimp
     rstconsulta!impresora_kgtintes = cadbl(!kg1) + cadbl(!kg2) + cadbl(!kg3) + cadbl(!kg4) + cadbl(!kg5) + cadbl(!kg6) + cadbl(!kg7) + cadbl(!kg8)
     rstconsulta!impresora_kgdisolvents = cadbl(!kg9) + cadbl(!kg10)
     End With
   End If
   rstconsulta!impresora_costtintes = cadbl(rstconsulta!impresora_kgtintes) * 4.27 + rstconsulta!impresora_kgdisolvents * 2.25
   rstconsulta!impresora_costdisolvents = cadbl(rstconsulta!impresora_kgdisolvents) * 2.25
   n = cadbl(rstconsulta!impresora_numimpresora)
   rstconsulta!impresora_cost = IIf(n = 5, 95, 155) * (rstconsulta!impresora_tempsfunc + rstconsulta!impresora_tempscanvi)
    Set rstimp = Nothing
End Sub
Sub obrestocks(Optional noobrirbd As Boolean)
 Dim camistocks As String
camistocks = llegir_ini("General", "ruta_stocksmdb", "comandes.ini")
If camistocks = "{[}]" Then camistocks = "\\Ser2\documentos\Stock Reclamaciones\Estoc inplacsa.mdb"
'If Not existeix(camistocks) Then camistocks = "\\serverprodu\dades\progcomandes\dades\copiaestocinplacsa.mdb"
If Not noobrirbd Then Set dbstocks = OpenDatabase(camistocks)
End Sub
Sub dadescapcalera(numc As Double)
  Dim rstcom1 As Recordset
  Dim rstcom2 As Recordset
  Dim rstcom3 As Recordset
  Dim rstbobent As Recordset
  Dim rstcli As Recordset
  Dim rstmat As Recordset
  
  Set rstcom1 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
  Set rstcom2 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(rstcom1!linkcomanda1)))
  Set rstcom3 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(rstcom1!linkcomanda2)))
  Set rstcli = dbtmp.OpenRecordset("select * from clients where codi=" + atrim(cadbl(rstcom1!client)))
  Set rstbobent = dbbaixes.OpenRecordset("SELECT impressores.comanda, impressores.tipus, bobinesentimp.palet, bobinesentimp.bobina FROM bobinesentimp INNER JOIN (impressores INNER JOIN bobinesimp ON impressores.Id = bobinesimp.controlid) ON bobinesentimp.id = bobinesimp.Id where (((impressores.comanda) = " + atrim(numc) + ") And ((impressores.tipus) = 'F')) ORDER BY impressores.Id;")
  
  rstconsulta!lot1 = rstcom1!comanda
  rstconsulta!lot2 = rstcom1!linkcomanda1
  rstconsulta!lot3 = rstcom1!linkcomanda2
  If cadbl(rstcom1!ampleesq) > 0 Then rstconsulta!lot1_factorpesxmil = (rstcom1!pes1000mtrs) / cadbl(rstcom1!ampleesq)
  If rstcom2!comanda > 0 And cadbl(rstcom2!ampleesq) > 0 Then rstconsulta!lot2_factorpesxmil = (rstcom2!pes1000mtrs) / cadbl(rstcom2!ampleesq)
  If rstcom3!comanda > 0 And cadbl(rstcom3!ampleesq) > 0 Then rstconsulta!lot3_factorpesxmil = (rstcom3!pes1000mtrs) / cadbl(rstcom3!ampleesq)
  If Not rstcli.EOF Then rstconsulta!codiclient = cadbl(rstcli!codi): rstconsulta!nomclient = atrim(rstcli!nom)
  rstconsulta!producte = rstcom1!producte
  If Not rstbobent.EOF Then
    Set rstmat = dbstocks.OpenRecordset("SELECT Palets.Idpalet, Palets.Ample, Productes.Nomprod, Families.Nomfam FROM (Families INNER JOIN Productes ON Families.Idfam = Productes.Idfam) INNER JOIN Palets ON Productes.Idprod = Palets.Idprod WHERE (((Palets.Idpalet)=" + atrim(cadbl(rstbobent!palet)) + "));")
    If Not rstmat.EOF Then rstconsulta!material_lot1 = atrim(rstmat!Nomprod): rstconsulta!nomfamilia_lot1 = atrim(rstmat!Nomfam): rstconsulta!ample_lot1 = cadbl(rstmat!ample)
  End If
  If rstcom2!comanda > 0 Then
     Set rstbobent = dbbaixes.OpenRecordset("SELECT impressores.comanda, impressores.tipus, bobinesentimp.palet, bobinesentimp.bobina FROM bobinesentimp INNER JOIN (impressores INNER JOIN bobinesimp ON impressores.Id = bobinesimp.controlid) ON bobinesentimp.id = bobinesimp.Id where (((impressores.comanda) = " + atrim(cadbl(rstcom1!linkcomanda1)) + ") And ((impressores.tipus) = 'F')) ORDER BY impressores.Id;")
     If Not rstbobent.EOF Then
       Set rstmat = dbstocks.OpenRecordset("SELECT Palets.Idpalet, Palets.Ample, Productes.Nomprod, Families.Nomfam FROM (Families INNER JOIN Productes ON Families.Idfam = Productes.Idfam) INNER JOIN Palets ON Productes.Idprod = Palets.Idprod WHERE (((Palets.Idpalet)=" + atrim(cadbl(rstbobent!palet)) + "));")
       If Not rstmat.EOF Then rstconsulta!material_lot2 = atrim(rstmat!Nomprod): rstconsulta!nomfamilia_lot2 = atrim(rstmat!Nomfam): rstconsulta!ample_lot2 = cadbl(rstmat!ample)
     End If
  End If
  
  If rstcom3!comanda > 0 Then
    Set rstbobent = dbbaixes.OpenRecordset("SELECT impressores.comanda, impressores.tipus, bobinesentimp.palet, bobinesentimp.bobina FROM bobinesentimp INNER JOIN (impressores INNER JOIN bobinesimp ON impressores.Id = bobinesimp.controlid) ON bobinesentimp.id = bobinesimp.Id where (((impressores.comanda) = " + atrim(cadbl(rstcom1!linkcomanda2)) + ") And ((impressores.tipus) = 'F')) ORDER BY impressores.Id;")
    If Not rstbobent.EOF Then
      Set rstmat = dbstocks.OpenRecordset("SELECT Palets.Idpalet, Palets.Ample, Productes.Nomprod, Families.Nomfam FROM (Families INNER JOIN Productes ON Families.Idfam = Productes.Idfam) INNER JOIN Palets ON Productes.Idprod = Palets.Idprod WHERE (((Palets.Idpalet)=" + atrim(cadbl(rstbobent!palet)) + "));")
      If Not rstmat.EOF Then rstconsulta!material_lot3 = atrim(rstmat!Nomprod): rstconsulta!nomfamilia_lot3 = cadbl(rstmat!Nomfam): rstconsulta!ample_lot3 = cadbl(rstmat!ample)
    End If
  End If
  Set rstbobent = Nothing
  
  Set rstcom1 = Nothing
  Set rstcom2 = Nothing
  Set rstcom3 = Nothing
End Sub
Function eliminar_basededades(nombd As String) As Boolean
  On Error GoTo erro
  If existeix(nombd) Then Kill nombd
  eliminar_basededades = True
  Exit Function
erro:
   eliminar_basededades = False
End Function
Sub crear_taules_tmp()
  Dim camps(100, 2) As String
  processant = "Creant les taules i obrint les base de dades necessaries..."
  DoEvents
  taula_tmp = "costos"
  On Error Resume Next
   dbconsulta.Execute "drop table " + taula_tmp
  On Error GoTo 0
  i = 1
  camps(i, 1) = "lot1": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "lot2": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "lot3": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "lot1_factorpesxmil": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "lot2_factorpesxmil": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "lot3_factorpesxmil": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "codiclient": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "nomclient": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "producte": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "material_lot1": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "ample_lot1": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "nomfamilia_lot1": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "material_lot2": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "ample_lot2": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "nomfamilia_lot2": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "material_lot3": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "ample_lot3": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "nomfamilia_lot3": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "metres_lot1": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "kilos_lot1": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "preumaterial_lot1": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "metres_lot2": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "kilos_lot2": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "preumaterial_lot2": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "metres_lot3": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "kilos_lot3": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "preumaterial_lot3": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "impresora_numimpresora": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "impresora_tempscanvi": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "impresora_tempsfunc": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "impresora_metresfunc": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "impresora_kgtintes": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "impresora_kgdisolvents": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "impresora_mtrsprova": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "impresora_mtrsmin": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "impresora_cost": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "impresora_costtintes": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "impresora_costdisolvents": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "laminadora1_tempscanvi": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "laminadora1_tempsfunc": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "laminadora1_cost": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "laminadora1_totaladhesius": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "laminadora1_costadhesius": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "laminadora1_mtrsmin": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "laminadora2_tempscanvi": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "laminadora2_tempsfunc": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "laminadora2_cost": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "laminadora2_totaladhesius": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "laminadora2_costadhesius": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "laminadora2_mtrsmin": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "rebobinadora_numrebobinadora": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "rebobinadora_tempsfunc": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "rebobinadora_tempscanvi": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "rebobinadora_mtrsmin": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "rebobinadora_cost": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "entrega_kilosentregats": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "entrega_metresentregats": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "entrega_eurosxunitat": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "entrega_unitatfacturada": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "entrega_pvpcomanda": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "entrega_costcomanda": camps(i, 2) = "double": i = i + 1
  
  
  dbconsulta.Execute ("create table " + taula_tmp + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbconsulta.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
  
End Sub


Private Sub Command2_Click()
  parar.Tag = "1"
End Sub

Private Sub datafi_LostFocus()
If IsDate(datafi) Then
   datafi = Format(datafi, "dd/mm/yy")
  End If
End Sub

Private Sub datainici_LostFocus()
  If IsDate(datainici) Then
   datainici = Format(datainici, "dd/mm/yy")
  End If
End Sub

Private Sub parar_Click()
  parar.Tag = "1"
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triarclient
End Sub
Sub triarclient()
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select * from clients"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text3.Tag = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   Text3.Text = atrim(formseleccio.Data1.Recordset!nom)
  End If
  Unload formseleccio
  
End Sub
Sub triarproducte()
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codi,descripcio,ruta from productes"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   codiproducte = atrim(formseleccio.Data1.Recordset!codi)
   descripcioproducte = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub
Private Sub Text4_Change()

End Sub
