VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ordremuntatge 
   Caption         =   "Ordre Muntatge"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   5520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Ordremuntatge.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport llistat 
      Left            =   4980
      Top             =   4695
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   360
      Left            =   4995
      TabIndex        =   6
      Top             =   6810
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton insertar 
      Enabled         =   0   'False
      Height          =   465
      Left            =   4800
      Picture         =   "Ordremuntatge.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Insertar comandes"
      Top             =   1920
      Width           =   645
   End
   Begin VB.CommandButton borrar 
      Enabled         =   0   'False
      Height          =   465
      Left            =   4800
      Picture         =   "Ordremuntatge.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Borrar comanda"
      Top             =   2385
      Width           =   645
   End
   Begin VB.CommandButton sortir 
      Height          =   465
      Left            =   4800
      Picture         =   "Ordremuntatge.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Sortir"
      Top             =   2865
      Width           =   645
   End
   Begin VB.CommandButton Command1 
      Height          =   1035
      Left            =   4800
      Picture         =   "Ordremuntatge.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Afegir comandes"
      Top             =   75
      Width           =   645
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "Ordremuntatge.frx":1BB2
      Height          =   7575
      Left            =   45
      OleObjectBlob   =   "Ordremuntatge.frx":1BCA
      TabIndex        =   0
      Top             =   45
      Width           =   4695
   End
   Begin VB.Data ordrecomandes 
      Caption         =   "ordrecomandes"
      Connect         =   "Access"
      DatabaseName    =   "W:\progcomandes\dades\baixes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   4470
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "muntadora_ordremuntatge"
      Top             =   3765
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.CommandButton afegir 
      Height          =   465
      Left            =   4800
      Picture         =   "Ordremuntatge.frx":240F
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Afegir comandes"
      Top             =   1440
      Width           =   645
   End
   Begin VB.Label ethoresmuntades 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   7
      Top             =   7695
      Width           =   7290
   End
   Begin VB.Menu mactualitzacioclixes 
      Caption         =   "Actualització de Clixes"
      Begin VB.Menu mllistatsiconsultes 
         Caption         =   "Llistats i consultes"
         Begin VB.Menu mclixesvells 
            Caption         =   "Llistat de clixes amb última data d'impresió"
         End
         Begin VB.Menu mllistatquantsperbossa 
            Caption         =   "Llistat de quantitats de bosses per arxiu"
         End
         Begin VB.Menu mconsultesliniaimarca 
            Caption         =   "Consultes Linia i Marca i referències"
         End
      End
      Begin VB.Menu mdadesdelsclixes 
         Caption         =   "Dades dels clixes"
      End
      Begin VB.Menu mcanviubicacio 
         Caption         =   "Canvi XL de la bossa"
      End
      Begin VB.Menu mcanvisaniloxidensitats 
         Caption         =   "Canvis Anilox i Densitats dels operaris"
      End
   End
   Begin VB.Menu mestocadhesiu 
      Caption         =   "Estoc Adhesiu"
   End
   Begin VB.Menu mmenuimpresores 
      Caption         =   "menuimpresores"
      Visible         =   0   'False
      Begin VB.Menu mF2 
         Caption         =   "F2"
      End
      Begin VB.Menu mfw 
         Caption         =   "FW"
      End
   End
   Begin VB.Menu mpdfbaixes 
      Caption         =   "Veure baixes anteriors"
   End
End
Attribute VB_Name = "ordremuntatge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub afegir_comandes_maquina(vnommaquina As String)
   Dim comandes As String
   Dim comanda As Double
   Dim vtoteslescomandes As String
   Dim numtreball As Double
   Dim vhemafegituna As Boolean
   
   Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificaciooperaris.mdb")
   comandes = InputBox("Entra les comandes separades per / o una de sola", "Entrada comandes")
   comandes = comandes + "/"
   vtoteslescomandes = comandes
   While Len(comandes) > 0
      If InStr(1, comandes, "/") = 0 Then comandes = ""
      comanda = cadbl(Mid(comandes, 1, InStr(1, comandes, "/") - 1))
      comandes = Mid(comandes, InStr(1, comandes, "/") + 1)
      If comandavalida(comanda, , numtreball) Then vhemafegituna = True: afegircomanda comanda, vnommaquina, False, True
      'imprimiretiquetabossaclixes numtreball, llistat, True
      Form1.comprovarsihihaunaltrereferenciaperimprimir cadbl(comanda), numtreball
   Wend
   If InStr(1, LCase(nomordinador), "muntadora") > 0 And vhemafegituna Then
     'sha afegit desde maquina avisar a tintes
      enviarmailatintes "Comanda/s afegides per l'operari: " + vtoteslescomandes
   End If
   possarvalorscomandavisual
   ordrecomandes.Refresh
   Set dbplanificacio = Nothing
End Sub
Sub enviarmailatintes(vcos As String)
   Dim rutamdb As String
   Dim dbavisos As Database
   Dim rsta As Recordset
   Dim vemail As String
   vcos = treure_apostruf(vcos)
   vemail = "tintesinplacsa@gmail.com; expedicions@inplacsa.com"
  ' vemail = "miquel.inplacsa@gmail.com"
   rutamdb = rutadelfitxer(cami) + "avisosincidencies.mdb"
   Set dbavisos = DBEngine.OpenDatabase(rutamdb)
   dbavisos.Execute ("insert into envios_mails (data,destinatari,assumpte,cos) values (now,'" + vemail + "','Comanda afegida a muntadora pels Operaris','" + vcos + "')")
   dbavisos.Close
   Set dbavisos = Nothing
End Sub
Sub enviar_email_oficina_vindraclient(vnumc As String)
   Dim rutamdb As String
   Dim dbavisos As Database
   Dim rsta As Recordset
   Dim vemail As String
   Dim vcos As String
   Set rsta = dbcomandes.OpenRecordset("select * from comandesmesextres where comanda=" + atrim(vnumc))
   If rsta.EOF Then Exit Sub
   vcos = "La comanda " + atrim(vnumc) + " s,ha afegit a muntadora i ha de venir el client a verificar-la."
   vcos = vcos + Chr(10) + Chr(13) + "Client: " + atrim(rsta!nomclient) + Chr(10) + Chr(13) + atrim(rsta!marcailinia)
   vcos = treure_apostruf(vcos)
   vemail = "avisBAT"
  ' vemail = "miquel.inplacsa@gmail.com"
   rutamdb = rutadelfitxer(cami) + "avisosincidencies.mdb"
   Set dbavisos = DBEngine.OpenDatabase(rutamdb)
   dbavisos.Execute ("insert into envios_mails (data,destinatari,assumpte,cos) values (now,'" + vemail + "','Comanda a muntadora HA DE VENIR EL CLIENT A DONAR OK.','" + vcos + "')")
   dbavisos.Close
   Set dbavisos = Nothing
End Sub
Sub afegircomanda(numc As Double, vnommaquina As String, Optional insertar As Boolean, Optional primer As Boolean)
 Dim posicio As Double
  Dim rst As Recordset
  Dim vresp As String
  Set rst = dbbaixes.OpenRecordset("select estatgestio from comandesrevisadesatintes where comanda=" + atrim(numc))
  If Not rst.EOF Then
     If atrim(rst!estatgestio) = "C" Then
          If UCase(InputBox("Aquesta comanda te una compra de tinta relacionada vols afegir-la igualment?" + Chr(10) + "ESCRIU [SI] o [NO]", "Compra relacionada")) = "NO" Then
               GoTo fi
          End If
     End If
  End If
  Set rst = dbcomandes.OpenRecordset("select clientvindraarevisarimpresio from comandes_extres where comanda=" + atrim(numc))
  If Not rst.EOF Then
    If rst!clientvindraarevisarimpresio Then
       vresp = UCase(InputBox("ALERTA!!!" + Chr(10) + "AQUESTA COMANDA NECESSITA OK DEL CLIENT." + Chr(10) + "Escriu SI per afegir-la igualment.", "VINDRÀ EL CLIENT."))
       If vresp = "SI" Then
          enviar_email_oficina_vindraclient atrim(numc)
         Else: Exit Sub
       End If
    End If
  End If
  If primer Then posicio = 50: GoTo afegir
  If ordrecomandes.Recordset.EOF Then
    posicio = 90
     Else
        If Not insertar Then
          ordrecomandes.Recordset.MoveLast
        End If
        posicio = ordrecomandes.Recordset!ordre + 100
        If insertar Then
          posicio = posicio - 50
           Else: posicio = posicio + 100
        End If
        
  End If
afegir:
  ordrecomandes.Recordset.AddNew
  ordrecomandes.Recordset!comanda = numc
  ordrecomandes.Recordset!ordre = posicio
  ordrecomandes.Recordset!nummaquina = vnommaquina
  ordrecomandes.Recordset.Update
  afegircomandaaplanificacioimpresoresoperaris numc, vnommaquina
  'un cop afegida la trec de reclamades a l'oficina per l'encarregat de impresores
  dbbaixes.Execute "delete * from planificacio_reclamades where numcomanda=" + atrim(numc)
  
  reordenarcomandes
fi:
  Set rst = Nothing
End Sub
Sub afegircomandaaplanificacioimpresoresoperaris(numc As Double, vnommaquina As String)
  Dim rstplanificacio As Recordset
  Dim vnummaq As Integer
  vnummaq = IIf(vnommaquina = "FW", 7, 9)
  Set rstplanificacio = dbplanificacio.OpenRecordset("select * from planificacioimp where comanda=" + atrim(numc)) ' + " and maquina=" + atrim(vnummaq))
  If rstplanificacio.EOF Then
      dbplanificacio.Execute "insert into planificacioimp (comanda,ordre,maquina) values (" + atrim(numc) + ",998," + atrim(vnummaq) + ")"
     Else
         rstplanificacio.Edit
         If rstplanificacio!ordre = 999 Then rstplanificacio!ordre = 998
         rstplanificacio!maquina = vnummaq
         rstplanificacio.Update
  End If
  Set rstplanificacio = Nothing
End Sub
Sub reordenarcomandes()
   Dim posicio As Double
   ordrecomandes.Refresh
   posicio = 100
   While Not ordrecomandes.Recordset.EOF
      If ordrecomandes.Recordset!ordre <> posicio Then
         ordrecomandes.Recordset.Edit
         ordrecomandes.Recordset!ordre = posicio
         ordrecomandes.Recordset.Update
      End If
      posicio = posicio + 100
      ordrecomandes.Recordset.MoveNext
   Wend
   ordrecomandes.Refresh
End Sub

Private Sub afegir_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   'PopupMenu mmenuimpresores
   MsgBox "Si vols afegir una comanda ho has de fer desde l'ordinador d'IMPRESORES a l'ordre d'impresió podras afegir-la.", vbInformation, "Atenció"
End Sub

Private Sub borrar_Click()
  If Not ordrecomandes.Recordset.EOF And Not ordrecomandes.Recordset.BOF Then
     If MsgBox("Vols borrar aquest ordre de comandes de la comanda " + atrim(ordrecomandes.Recordset!comanda), vbInformation + vbYesNo, "Eliminar ordre") = vbYes Then
        ordrecomandes.Recordset.Delete
        ordrecomandes.Refresh
     End If
  End If
End Sub

Private Sub Command1_Click()
   If Not ordrecomandes.Recordset.EOF And Form1.ordredelescomandes.tag = "" Then Form1.numcomanda = atrim(ordrecomandes.Recordset!comanda)
     
   Unload ordremuntatge
End Sub

Sub baixadeclixes()
   Dim numtreball As Double
   Dim amaquina As Boolean
   If llegir_ini("Baixes", "programaamaquina", "comandes.ini") = 1 Then amaquina = True
   numtreball = cadbl(InputBox("Entra el numero de treball a modificar.", "Actualització del treball"))
   If numtreball > 0 Then
    'If amaquina Then If Not existeix("c:\ordprog.ini") Then Shell "\\serverprodu\dades\progcomandes\aplicacio\dccmd.exe -width=1024 -height=768"
    ShellAndWait "\\serverprodu\dades\progcomandes\aplicacio\clixesnous.exe comandes.ini " + atrim(numtreball) + " baixaclixes", vbNormalFocus
    'If amaquina Then If Not existeix("c:\ordprog.ini") Then Shell "\\serverprodu\dades\progcomandes\aplicacio\dccmd.exe -width=800 -height=600"
    Form1.imprimiretbossatreball numtreball, True
   End If
   
End Sub

Private Sub Command2_Click()
  Dim rstnum As Recordset
  Dim rst As Recordset
  Set rst = dbbaixes.OpenRecordset("select * from impresores_aniloxos ")
  While Not rst.EOF
    If rst!okcanvi < 2 Then
            If rst!observacions_comanda <> "" Or rst!anilox_comanda <> rst!anilox_original Or caadbl(rst!densitat_comanda) <> caadbl(rst!densitat_original) Or atrim(rst!tinta_original) <> atrim(rst!tinta_comanda) Or atrim(IIf(rst!coditinta_original = "0", "", rst!coditinta_original)) <> atrim(IIf(rst!coditinta_comanda = "0", "", rst!coditinta_comanda)) Or rst!ordretinter <> rst!ordretinter_original Then
                 rst.Edit
                 rst!okcanvi = 1
                 rst.Update
                Else:
                  rst.Edit
                  rst!okcanvi = 0
                  rst.Update
            End If
     End If
   
    rst.MoveNext
  Wend
End Sub
Function caadbl(valor As Variant) As Double
  If IsNull(valor) Then caadbl = 0
  If IsNumeric(valor) Then caadbl = valor
End Function

Private Sub Form_Load()
  possarvalorscomandavisual
  ordrecomandes.DatabaseName = Form1.datamuntadora.DatabaseName
  dbbaixes.Execute "UPDATE muntadoratot INNER JOIN muntadora_ordremuntatge ON muntadoratot.comanda = muntadora_ordremuntatge.comanda SET muntadora_ordremuntatge.muntada = [muntadoratot].[acabada];"
  dbbaixes.Execute "UPDATE impresores_ordreimpresio RIGHT JOIN muntadora_ordremuntatge ON impresores_ordreimpresio.comanda = muntadora_ordremuntatge.comanda SET muntadora_ordremuntatge.ordre = [impresores_ordreimpresio].[ordre];"
  dbbaixes.Execute "UPDATE muntadora_ordremuntatge LEFT JOIN impresores_ordreimpresio ON muntadora_ordremuntatge.comanda = impresores_ordreimpresio.comanda SET muntadora_ordremuntatge.comandavisual = Trim([impresores_ordreimpresio].[comanda]) & ' [' & Format([impresores_ordreimpresio].[dataprogramada],'dd/mm hh:nn')&']-'+trim([muntadora_ordremuntatge].[nummaquina]) WHERE (((impresores_ordreimpresio.dataprogramada) Is Not Null));"
  ordrecomandes.RecordSource = "select * from muntadora_ordremuntatge where not muntada order by nummaquina,ordre"
  
End Sub
Sub possarvalorscomandavisual()
   Dim rst As Recordset
   Set rst = dbbaixes.OpenRecordset("select * from muntadora_ordremuntatge order by ordre")
   While Not rst.EOF
      dbbaixes.Execute "update muntadora_ordremuntatge set comandavisual='" + mirarsialtracomandapendent(rst!comanda) + "-" + atrim(rst!nummaquina) + "' where comanda=" + atrim(rst!comanda)
      rst.MoveNext
   Wend
   calcular_horesmuntades
End Sub
Sub calcular_horesmuntades()
   Dim rst As Recordset
   Dim vsql As String
   Dim vtotalhores As Double
   Dim vmetresminut As Double
   
   ethoresmuntades = ""
   vsql = "SELECT comandes.cantitatex AS metres, impresores_ordreimpresio.nommaquina,impresores_ordreimpresio.metresminutultimcop "
   vsql = vsql + " FROM comandes RIGHT JOIN impresores_ordreimpresio ON comandes.comanda = impresores_ordreimpresio.comanda"
   vsql = vsql + " where (comandes.comanda In (select comanda from muntadoratot where acabada)) and maquina=9 "
   vsql = vsql + " ORDER BY impresores_ordreimpresio.nommaquina;"
calcular:
   Set rst = dbbaixes.OpenRecordset(vsql)
   vtotalhores = 0
   If Not rst.EOF Then vnommaquina = IIf(atrim(rst!nommaquina) = "", "F2óFW", atrim(rst!nommaquina))
   While Not rst.EOF
     vmetresminut = 200
     If cadbl(rst!metresminutultimcop) > 0 Then vmetresminut = cadbl(rst!metresminutultimcop)
     'converteixo les comandes a minuts comptat 200metres per minut i faix 1,5h de canvi per comanda
     vtotalhores = vtotalhores + (cadbl(rst!metres) / vmetresminut) + 90
     
     rst.MoveNext
   Wend
   vtotalhores = Redondejar(vtotalhores / 60, 0)
   ethoresmuntades = ethoresmuntades + vnommaquina + ": " + atrim(vtotalhores) + "H. "
   If InStr(1, vsql, "maquina=7") = 0 Then
        vsql = "SELECT comandes.cantitatex AS metres, impresores_ordreimpresio.nommaquina,impresores_ordreimpresio.metresminutultimcop "
        vsql = vsql + " FROM comandes RIGHT JOIN impresores_ordreimpresio ON comandes.comanda = impresores_ordreimpresio.comanda"
        vsql = vsql + " where (comandes.comanda In (select comanda from muntadoratot where acabada)) and maquina=7 "
        vsql = vsql + " ORDER BY impresores_ordreimpresio.nommaquina;"
        GoTo calcular
   End If
   ethoresmuntades = ethoresmuntades
End Sub
Sub calcular_horesmuntades_NOVALID()
   Dim rst As Recordset
   Dim vsql As String
   Dim vtotalhores As Double
   ethoresmuntades = ""
   vsql = "SELECT Sum(comandes.cantitatex) AS Tmetres, impresores_ordreimpresio.nommaquina, Count(comandes.comanda) AS Tcomandes"
   vsql = vsql + " FROM comandes RIGHT JOIN impresores_ordreimpresio ON comandes.comanda = impresores_ordreimpresio.comanda"
   vsql = vsql + " WHERE (((comandes.comanda) In (select comanda from muntadoratot where acabada)) AND ((comandes.proximaseccio)='I'))"
   vsql = vsql + " GROUP BY impresores_ordreimpresio.nommaquina;"

   Set rst = dbbaixes.OpenRecordset(vsql)
   While Not rst.EOF
     'converteixo les comandes a minuts comptat 200metres per minut i faix 1,5h de canvi per comanda
     vtotalhores = (cadbl(rst!tmetres) / 200) + (90 * cadbl(rst!Tcomandes))
     vtotalhores = Redondejar(vtotalhores / 60, 0)
     vnommaquina = IIf(atrim(rst!nommaquina) = "", "F2óFW", atrim(rst!nommaquina))
     ethoresmuntades = ethoresmuntades + vnommaquina + ": " + atrim(vtotalhores) + "H. "
     rst.MoveNext
   Wend
   ethoresmuntades = "Hores muntades: " + ethoresmuntades
End Sub
Function mirarsialtracomandapendent(numc As Double) As String

   Dim rstc As Recordset
   Dim c As String
   Set rstc = dbcomandes.OpenRecordset("select comanda from comandes where proximaseccio='E' and numtreball in (select numtreball from comandes where comanda=" + atrim(numc) + ")")
   c = ""
   While Not rstc.EOF
     If cadbl(rstc!comanda) <> numc Then c = c + atrim(rstc!comanda) + " "
     rstc.MoveNext
   Wend
   mirarsialtracomandapendent = atrim(numc)
   If c <> "" Then mirarsialtracomandapendent = "** " + atrim(numc) + " **"

End Function
Private Sub insertar_Click()
   Dim comandes As String
   Dim comanda As Double
   Dim posicioactual As Double
   If Not ordrecomandes.Recordset.EOF Then
      posicioactual = ordrecomandes.Recordset!comanda
     Else: MsgBox "Primer has de colocar-te a la comanda on vols insertar la serie": Exit Sub
   End If
   comandes = InputBox("Entra les comandes separades per / o una de sola", "Entrada comandes")
   comandes = comandes + "/"
   While Len(comandes) > 0
      If InStr(1, comandes, "/") = 0 Then comandes = ""
      comanda = cadbl(Mid(comandes, 1, InStr(1, comandes, "/") - 1))
      comandes = Mid(comandes, InStr(1, comandes, "/") + 1)
      If comandavalida(comanda) Then
        afegircomanda comanda, "F2", True
        ordrecomandes.Recordset.FindFirst "comanda=" + atrim(comanda)
      End If
   Wend
   possarvalorscomandavisual
   ordrecomandes.Refresh
End Sub
Function posicioenlaruta(numc As Double) As String
  Dim rstp As Recordset
  Dim rstpr As Recordset
  Dim laruta As String
   
  'If InStr(1, "VPT", seccioactual) = 0 Then Exit Function
  Set rstp = dbbaixes.OpenRecordset("SELECT comandes.comanda,comandes.proximaseccio,comandes.producte, rebobinadorestot.acavada as acavadar, laminadorestot.acavada as acavadal, impressorestot.acavada as acavadai FROM ((comandes LEFT JOIN rebobinadorestot ON comandes.comanda = rebobinadorestot.comanda) LEFT JOIN laminadorestot ON comandes.comanda = laminadorestot.comanda) LEFT JOIN impressorestot ON comandes.comanda = impressorestot.comanda WHERE (((comandes.comanda)=" + atrim(numc) + "));")
  If Not rstp.EOF Then
     Set rstpr = dbcomandes.OpenRecordset("select ruta from productes where codi='" + atrim(rstp!producte) + "'")
     If rstpr.EOF Then Exit Function
     laruta = atrim(rstpr!ruta)
     If InStr(1, laruta, "R") > 0 And cadblnull_1(rstp!acavadar) = 0 Then posicioenlaruta = "R"
     If InStr(1, laruta, "L") > 0 And cadblnull_1(rstp!acavadal) = 0 Then posicioenlaruta = "L"
     If InStr(1, laruta, "I") > 0 And cadblnull_1(rstp!acavadai) = 0 Then posicioenlaruta = "I"
  End If
  If posicioenlaruta = "" Or atrim(rstp!proximaseccio) = "E" Then posicioenlaruta = rstp!proximaseccio
  
  Set rstp = Nothing
  Set rstpr = Nothing
End Function
Function cadblnull_1(acabada As Variant) As Double
   If IsNull(acabada) Then cadblnull_1 = -1: Exit Function
   cadblnull_1 = cadbl(acabada)
End Function
Function comandavalida(numc As Double, Optional nocomprovarllista As Boolean, Optional ByRef numtreball As Double, Optional ByRef numordremodificacio) As Boolean
   Dim rst As Recordset
   Dim rst_extres As Recordset
   comandavalida = False
   If numc = 0 Then Exit Function
   If Not nocomprovarllista Then
     Set rst = dbbaixes.OpenRecordset("select * from muntadora_ordremuntatge where comanda=" + atrim(numc))
     If Not rst.EOF Then MsgBox "La comanda " + atrim(numc) + " ja està a la llista.": GoTo fi
   End If
   Set rst = dbcomandes.OpenRecordset("SELECT comandes.dataactivacio,comandes.comanda, productes.ruta, comandes.proximaseccio,comandes.impressio,comandes.numtreball,comandes.numordremodificacio FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(numc) + "));")
   Set rst_extres = dbcomandes.OpenRecordset("select * from comandes_extres where comanda=" + atrim(numc), , ReadOnly)
   If Not rst_extres.EOF Then
      If cadbl(rst_extres!passaraimpresores) = 0 Then
         MsgBox "Atenció comanda en StandBy de PLANIFICACIÓ.", vbCritical, "STANDBY"
         comandavalida = False
         GoTo fi
      End If
   End If
   If Not rst.EOF Then
       If rst!dataactivacio = Null Then MsgBox "Aquesta comanda està desactivada no es pot muntar.", vbCritical, "Error": comandavalida = False: GoTo fi
       numtreball = cadbl(rst!numtreball)
       numordremodificacio = cadbl(rst!numordremodificacio)
       proximaseccio = posicioenlaruta(numc)
       If proximaseccio = "I" And InStr(1, rst!ruta, "I") > 0 Then
             comandavalida = True
               Else
                 If InStr(1, rst!ruta, "I") = 0 Then MsgBox "La comanda " + atrim(numc) + " no te seccio d'impresores"
                 If proximaseccio <> "I" Then
                   MsgBox "La comanda " + atrim(numc) + " no està apunt per imprimir."
                 End If
                 
       End If
       If rst!impressio = "F" Then
          MsgBox "A la comanda " + atrim(numc) + " li Falta Autoritzar.", vbCritical, "Atenció"
          comandavalida = False
       End If
        Else: MsgBox "La comanda " + atrim(numc) + " no existeix."
   End If
   If comandavalida Then
      If Not tepackinglist(cadbl(numc)) Then
         MsgBox "Aquesta comanda encara no te material assignat.", vbCritical, "Atenció"
         comandavalida = False
      End If
   End If
   If comandavalida Then
     If Not Form1.clixesentratsafabrica(cadbl(numc)) Then
       comandavalida = False
       MsgBox "La comanda " + atrim(numc) + " no te els CLIXES ENTRATS a disseny. No es poden utilitzar.", vbCritical, "Atenció"
     End If
   End If
fi:
   Set rst_extres = Nothing
   Set rst = Nothing
End Function
Function tepackinglist(numc As Double) As Boolean
   Dim rstt As Recordset
   tepackinglist = False
   Set rstt = dbstocks.OpenRecordset("select * from parcials where  comanda='" + atrim(numc) + "'")
   If Not rstt.EOF Then tepackinglist = True
   Set rstt = dbcomandes.OpenRecordset("select assignarstock from comandes_extres where comanda=" + atrim(numc))
   If Not rstt.EOF Then
      If rstt!assignarstock Then tepackinglist = True
   End If
End Function



Private Sub mcanvisaniloxidensitats_Click()
   Dim numc As String
   Dim vtreball As Double
   Dim vmodificacio As Double
  ' numc = "151646"
  ' escriure_ini "baixes", "imprimircomandanomesimp", "S", "comandes.ini"
  ' escriure_ini "baixes", "imprimircomanda", numc, "comandes.ini"
  ' Shell "\\serverprodu\dades\progcomandes\aplicacio\comandes.exe comandes.ini imprimir"
  numc = 1
  While numc <> 0
    numc = llista_aniloxospendentsok
    Unload formaniloxos
    If numc <> 0 Then
        Load formaniloxos
        formaniloxos.tag = atrim(numc)
        formaniloxos.fbotonsok.tag = "activats"
        formaniloxos.tag = numc
        formaniloxos.Show 1
        
         'comprovo si la comanda afectada s'han de fer canvis a la secció d'impresores
        Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
        comprovar_diferencies_amb_la_comanda cadbl(numc)
        Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
        If rstc.EOF Then
          vtreball = cadbl(rstc!numtreball): vmodificacio = cadbl(rstc!numordremodificacio)
          Set rstc = dbcomandes.OpenRecordset("select * from comandes where numtreball=" + atrim(vtreball) + " and numordremodificacio=" + atrim(vmodificacio) + " and (proximaseccio='I' or proximaseccio='E')")
          While Not rstc.EOF
            comprovar_diferencies_amb_la_comanda cadbl(rstc!comanda)
          Wend
        End If
    End If
  Wend
  
End Sub
Sub comprovar_diferencies_amb_la_comanda(numc As Double)
  If mirardiferenciescomandaitreball(cadbl(numc)) Then
           imprimirdiferenciescomandaitreball cadbl(numc)
           If UCase(InputBoxEx("La comanda " + atrim(numc) + " te diferencies amb aquest treball." + Chr(10) + "VOLS ACTUALITZAR LA COMANDA AMB LES DADES D'AQUEST TREBALL? (Escriu [si] per fer-ho.)", "Atenció")) = "SI" Then
               posardiferenciesacomandadeltreball cadbl(numc)
           End If
  End If
End Sub
Function llista_aniloxospendentsok() As String
  Dim instsql As String
  instsql = "SELECT impresores_aniloxos.comanda AS Num_Comanda, First(clients.nom) AS Nom_Client, First(comandes.marcailinia) AS Marca_Linia FROM impresores_aniloxos LEFT JOIN (comandes LEFT JOIN clients ON comandes.client = clients.codi) ON impresores_aniloxos.comanda = comandes.comanda Where (((impresores_aniloxos.okcanvi) = 1)) GROUP BY impresores_aniloxos.comanda HAVING (((impresores_aniloxos.comanda)>0)) order by First(clients.nom)"


  Load formseleccionou
  
  formseleccionou.Data1.DatabaseName = cami
  formseleccionou.Data1.RecordSource = instsql
  
  
  'formseleccio.Width = 7000
  formseleccionou.refrescar
  formseleccionou.DBGrid2.Columns(0).width = 1000
  formseleccionou.DBGrid2.Columns(1).width = 2500
  formseleccionou.DBGrid2.Columns(2).width = 5000
  formseleccionou.width = 10000
  formseleccionou.caption = "Escullir Comanda"
  formseleccionou.botofiltre.tag = "0"
  formseleccionou.Show 1
  llista_aniloxospendentsok = "0"
  If seleccioret = 1 Then
   llista_aniloxospendentsok = atrim(cadbl(formseleccionou.Data1.Recordset!num_comanda))
  End If
  Unload formseleccionou
   ratoli "normal"
End Function

Private Sub mcanviubicacio_Click()
   Dim numtreball As Double
   Dim rst As Recordset
   Dim vnouxl As String
   numtreball = cadbl(InputBox("Entra el numero de treball a modificar.", "Actualització del treball"))
   
   If numtreball > 0 Then
       Set rst = dbclixes.OpenRecordset("Select linia,marca,arxiu from clixes where id_treball=" + atrim(numtreball))
       If Not rst.EOF Then
         vnouxl = InputBox("Entra el nou XL pel treball " + atrim(numtreball) + Chr(10) + atrim(rst!linia) + " - " + atrim(rst!marca), "Nou XL", atrim(rst!arxiu))
         If InStr(1, vnouxl, "-") = 0 Then MsgBox "Aquest arxiu no porta '-'", vbCritical, "Error": Exit Sub
         dbclixes.Execute "update clixes set arxiu='" + atrim(vnouxl) + "' where id_treball=" + atrim(numtreball)
         Form1.imprimiretbossatreball numtreball, True
         'imprimiretiquetabossaclixes numtreball, llistat, True
       End If
   End If
End Sub

Private Sub mclixesvells_Click()
   llistatclixesvells
End Sub
Sub llistatclixesvells()
   Dim numtreball As Double
   Dim amaquina As Boolean
   If llegir_ini("Baixes", "programaamaquina", "comandes.ini") = 1 Then amaquina = True
   If amaquina Then If Not existeix("c:\ordprog.ini") Then Shell "\\serverprodu\dades\progcomandes\aplicacio\dccmd.exe -width=1024 -height=768"
   ShellAndWait "\\serverprodu\dades\progcomandes\aplicacio\clixesnous.exe comandes.ini  1 llistatclixesvells", vbNormalFocus
   If amaquina Then If Not existeix("c:\ordprog.ini") Then Shell "\\serverprodu\dades\progcomandes\aplicacio\dccmd.exe -width=800 -height=600"
   
End Sub
Function borrartaula(db As Database, nomtaula As String) As Boolean
   borrartaula = True
   On Error GoTo err
   db.Execute "drop table " + nomtaula
   Exit Function
err:
   borrartaula = False
End Function
Sub carregartoteslesreferencies()
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim ultimtreball As Double
   Dim ref As String
   Dim rstct As Recordset
   Set rst = dbclixes.OpenRecordset("select * from tmp_consultaliniaimarca")
   Set rstct = dbclixes.OpenRecordset("select distinct refclient,id_Treball   from clientsvinculats  where refclient<>'' order by id_treball")
   While Not rstct.EOF
     Set rstc = dbclixes.OpenRecordset("select distinct refclient  from clientsvinculats where id_treball=" + atrim(rstct!id_treball))
     ref = ""
     ultimtreball = rstct!id_treball
     While Not rstc.EOF
       ref = ref + IIf(ref <> "", " ¦ ", "") + atrim(rstc!refclient)
       rstc.MoveNext
     Wend
     dbclixes.Execute "update  tmp_consultaliniaimarca set prefclient='" + atrim(ref) + "' where id_treball=" + atrim(rstct!id_treball)
     While rstct!id_treball = ultimtreball
       rstct.MoveNext
       If rstct.EOF Then GoTo cont
     Wend
cont:
   Wend
   Set rst = Nothing
   Set rstc = Nothing
   Set rstct = Nothing
End Sub
Private Sub mconsultesliniaimarca_Click()
   
  Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim vtreball As Double
  Dim vordre As Double
  Dim caigudes As Double
  Dim fseleccio As Form
  
  Set fseleccio = formseleccionou
 ' MsgBox "Aquesta consulta tarda uns segons a generar-se..." + Chr(10) + "Sigues pacient.", vbOKOnly, "Consulta"
  ratoli "espera"
  If Not borrartaula(dbclixes, "tmp_consultaliniaimarca") Then MsgBox "Error borrant la taula, surt del programa i torna entrar.", vbCritical, "Atenció": GoTo fi
  sql = "SELECT First(Clixes.marca) AS marca, First(Clixes.linia) AS linia, First(Clientsvinculats.refclient) AS Prefclient, Clixes.id_treball, First(Clixes.arxiu) AS Prxiu INTO tmp_consultaliniaimarca FROM Clixes LEFT JOIN Clientsvinculats ON Clixes.id_treball = Clientsvinculats.id_treball where databaixaclixe=null GROUP BY Clixes.id_treball;"
  dbclixes.Execute sql
  carregartoteslesreferencies
  were = " order by linia desc"
  seleccioret = 1
  Load fseleccio
fseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "clixesnous.mdb"
fseleccio.Data1.RecordSource = "select * from tmp_consultaliniaimarca " + were
fseleccio.width = 11500
fseleccio.DBGrid2.col = 1
fseleccio.refrescar
  While seleccioret = 1
        
        fseleccio.DBGrid2.Columns(0).width = 3315
        fseleccio.DBGrid2.Columns(1).width = 3284
        fseleccio.DBGrid2.Columns(2).width = 1604
        fseleccio.DBGrid2.Columns(3).width = 794
        fseleccio.DBGrid2.Columns(4).width = 780
        'fseleccio.DBGrid2.Columns(5).width = 780
        fseleccio.sortirs.tag = "filtre"
        ratoli "normal"
        fseleccio.Show 1
        If seleccioret = 1 Then
          vtreball = atrim(fseleccio.Data1.Recordset!id_treball)
          Set rst = dbclixes.OpenRecordset("select ordre from modificacions where id_treball=" + atrim(vtreball) + " order by ordre desc")
          If Not rst.EOF Then vordre = atrim(rst!ordre)
          If vtreball > 0 And vordre > 0 Then obrir_pdf_treball vtreball, vordre
          
          Set rst = Nothing
        End If
  Wend
  Unload fseleccio
fi:
   ratoli "normal"
End Sub

Private Sub mdadesdelsclixes_Click()
   baixadeclixes
End Sub

Private Sub mestocadhesiu_Click()
  estocadhesiu.Show 1
End Sub
Private Sub mf2_Click()
  afegir_comandes_maquina "F2"
End Sub

Private Sub mfw_Click()
  afegir_comandes_maquina "FW"
End Sub

Private Sub mllistatquantsperbossa_Click()
     Dim resp As String
     Dim oapp As CRAXDDRT.Application
     Dim oreport As CRAXDDRT.Report
     
     Set oapp = New CRAXDDRT.Application
     Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistatquantitatsbossesperarxiu.rpt", 1)
     oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "clixesnous.mdb"
     'oreport.RecordSelectionFormula = "mid({Clixes.ubicacio},1,5)<>'Palet' and {Clixes.arxiu}<>'' and isnull({Clixes.databaixaclixe}) and {clixes.estatclixe}<>'RETORNEM CLIXES'"
     oreport.RecordSelectionFormula = "{@arxiusenseXL}>0 and (trim({Clixes.arxiu})<>'' and isnull({Clixes.databaixaclixe}) and {clixes.estatclixe}<>'RETORNEM CLIXES')"
     oreport.DiscardSavedData
     'oreport.FormulaFields.GetItemByName("rangdates").Text = "'Data inici: " + Format(datainici, "dd/mm/yyyy") + " i Data fi: " + Format(datafi, "dd/mm/yyyy") + "'"
     Load veurereport
     veurereport.CRViewer.ReportSource = oreport
     veurereport.CRViewer.DisplayGroupTree = False
     veurereport.CRViewer.ViewReport
     veurereport.WindowState = 2
     veurereport.Show 1

End Sub

Private Sub mpdfbaixes_Click()

  Dim vtreball As String
  Dim vcomanda As Double
  vtreball = InputBox("Escriu el numero de treball que vols veure o el numero de comanda.", "Veure baixa de treball o comanda")
  If cadbl(vtreball) > 100000 Then
       vcomanda = vtreball
        Else: vcomanda = buscarultimacomanda(vtreball)
  End If
  If vcomanda > 0 Then ensenya_PDF_baixes vcomanda
End Sub
Function buscarultimacomanda(vtreball As String) As Double
   Dim vordre As Integer
   Dim rst As Recordset
   If cadbl(vtreball) = 0 Then Exit Function
   
   vtreball = cadbl(vtreball)
   Set rst = dbcomandes.OpenRecordset("select comanda from comandes where numtreball=" + atrim(vtreball) + " and (proximaseccio<>'E' and proximaseccio<>'I') order by comanda desc")
   If Not rst.EOF Then buscarultimacomanda = cadbl(rst!comanda)
End Function

Sub ensenya_PDF_baixes(vcomanda As Double)
 Dim carpetadesti As String
 carpetadesti = llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini")
 carpetadesti = carpetadesti + "\Les_" + atrim(atrim(Int(cadbl(vcomanda) / 1000)) + "000") + "\" + atrim(vcomanda) + "\" + atrim(vcomanda) + "_BaixaMuntadora.pdf"
 obrir_document carpetadesti
End Sub

Private Sub sortir_Click()
Unload ordremuntatge
End Sub
