VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form formencarregat 
   Caption         =   "Opcions d'Encarregat"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   4425
   ClipControls    =   0   'False
   Icon            =   "formencarregat.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Comandes amb les bobines a sala LAMINADORA"
      Height          =   2130
      Left            =   240
      TabIndex        =   15
      Top             =   2220
      Width           =   4110
      Begin VB.CommandButton Command6 
         Height          =   1035
         Left            =   2340
         Picture         =   "formencarregat.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Imprimir llistat de comandes pendents."
         Top             =   930
         Width           =   705
      End
      Begin VB.CommandButton Command5 
         Caption         =   "NO Fet"
         Height          =   510
         Left            =   1185
         Picture         =   "formencarregat.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Canutus no tallats"
         Top             =   1455
         Width           =   1140
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Fet"
         Height          =   510
         Left            =   1200
         Picture         =   "formencarregat.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Canutus tallats"
         Top             =   930
         Width           =   1125
      End
      Begin VB.TextBox ccomandalam 
         BackColor       =   &H00EEE4D7&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1275
         TabIndex        =   16
         Top             =   345
         Width           =   1800
      End
      Begin VB.Label etfetbobines 
         BackStyle       =   0  'Transparent
         Caption         =   "NO PLAN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3105
         TabIndex        =   21
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Comanda:"
         Height          =   285
         Left            =   75
         TabIndex        =   17
         Top             =   420
         Width           =   1290
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2145
      Left            =   4395
      TabIndex        =   14
      Top             =   285
      Visible         =   0   'False
      Width           =   4410
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   4035
         Top             =   180
      End
      Begin VB.Image Image1 
         Height          =   1785
         Left            =   285
         Picture         =   "formencarregat.frx":1628
         Stretch         =   -1  'True
         Top             =   180
         Width           =   3825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Llista producció Lam"
      Height          =   1845
      Left            =   4830
      TabIndex        =   10
      Top             =   45
      Width           =   1830
      Begin VB.CommandButton bafegircomanda 
         Height          =   315
         Left            =   135
         Picture         =   "formencarregat.frx":47CA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Afegir comanda a la llista"
         Top             =   285
         Width           =   360
      End
      Begin VB.CommandButton beliminarcomanda 
         Height          =   330
         Left            =   1350
         Picture         =   "formencarregat.frx":4D54
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Eliminar comanda a la llista"
         Top             =   270
         Width           =   330
      End
      Begin VB.ListBox llistaproduccio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   90
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   -30
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame frameactdes 
      Caption         =   "Comandes amb canutus ja tallats."
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   -30
      Width           =   4110
      Begin VB.CommandButton Command3 
         Caption         =   "Agafar Std"
         Height          =   510
         Left            =   105
         Picture         =   "formencarregat.frx":52DE
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Canutus tallats"
         Top             =   1485
         Width           =   1125
      End
      Begin VB.CommandButton Command2 
         Height          =   1035
         Left            =   1245
         Picture         =   "formencarregat.frx":5868
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir llistat de tallats."
         Top             =   960
         Width           =   705
      End
      Begin VB.CommandButton Command1 
         Height          =   510
         Left            =   3270
         Picture         =   "formencarregat.frx":5DF2
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir llistat no tallats"
         Top             =   975
         Width           =   705
      End
      Begin VB.TextBox lotperactivar 
         BackColor       =   &H00EEE4D7&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1380
         TabIndex        =   3
         Top             =   315
         Width           =   1800
      End
      Begin VB.CommandButton Command7 
         Height          =   510
         Left            =   105
         Picture         =   "formencarregat.frx":637C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Canutus tallats"
         Top             =   975
         Width           =   1125
      End
      Begin VB.CommandButton Command8 
         Height          =   510
         Left            =   2205
         Picture         =   "formencarregat.frx":6906
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Canutus no tallats"
         Top             =   975
         Width           =   1065
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Comanda:"
         Height          =   285
         Left            =   150
         TabIndex        =   6
         Top             =   405
         Width           =   1290
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Canutus Tallats"
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   525
         TabIndex        =   5
         Top             =   765
         Width           =   1230
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Canutus No Tallats"
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   2505
         TabIndex        =   4
         Top             =   765
         Width           =   1470
      End
   End
   Begin VB.Menu mopcions 
      Caption         =   "Opcions"
      Begin VB.Menu mcanutusstandard 
         Caption         =   "Canutus Estandard"
      End
      Begin VB.Menu mdonarllaunadebaixa 
         Caption         =   "Donar llauna de baixa"
      End
   End
End
Attribute VB_Name = "formencarregat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bafegircomanda_Click()
  Dim vnumc As String
  Dim rstc As Recordset
  vnumc = InputBox("Entra el numero de comanda que vols afegir", "Afegir comanda a la llista")
  If cadbl(vnumc) > 0 Then
     Set rstc = dbbaixes.OpenRecordset("Select proximaseccio from comandes where comanda=" + atrim(vnumc))
     If Not rstc.EOF Then
        If rstc!proximaseccio = "I" Or rstc!proximaseccio = "L" Then
           dbbaixes.Execute "insert into laminadora_ordremuntatge (comanda) values (" + atrim(cadbl(vnumc)) + ")"
             Else: MsgBox "Aquesta comanda esta a la secció " + atrim(rstc!proximaseccio) + " no està apunt per laminar."
        End If
         Else: MsgBox "Aquesta comanda no existeix", vbCritical, "Error"
     End If
  End If
  carregar_llista_muntatge
End Sub

Private Sub beliminarcomanda_Click()
  Dim vnumc As String
  vnumc = InputBox("Entra el numero de comanda que vols ELIMINAR de la llista", "ELIMINAR comanda a la llista")
  If cadbl(vnumc) > 0 Then
     dbtmpb.Execute "DELETE * from laminadora_ordremuntatge where comanda=" + atrim(cadbl(vnumc))
     MsgBox "Comanda " + atrim(vnumc) + " borrada."
     carregar_llista_muntatge
  End If
End Sub

Private Sub ccomandalam_Change()
  mirar_estat_bobines
End Sub
Function mirar_estat_bobines()
   etfetbobines = ""
   If Not estaplanificada(cadbl(ccomandalam)) Then etfetbobines = "NO PLAN": GoTo fi
   If Len(ccomandalam.Tag) >= 6 Then
      mirarestatFET cadbl(ccomandalam.Tag)
   Else: etfetbobines = ""
   End If
fi:
   

End Function
Function estaplanificada(vnumc As Long) As Boolean
   Dim vsql As String
   Dim rst As Recordset
   Dim vcomanda As Long
   vcomanda = vnumc
   Set rst = dbtmp.OpenRecordset("select linkcomanda1,linkcomanda2,comanda,producte from comandes where comanda=" + atrim(vcomanda))
   If rst.EOF Then Exit Function
  ' If rst!producte = "PC" Then
  '   vcomanda = rst!linkcomanda1
  '     Else: vcomanda = rst!linkcomanda2
  ' End If
   vsql = "SELECT planificaciolam_operaris.comanda, comandes.proximaseccio, planificaciolam_operaris.ordre, *"
   vsql = vsql + " FROM planificaciolam_operaris INNER JOIN comandes ON planificaciolam_operaris.comanda = comandes.comanda "
   vsql = vsql + " Where comandes.comanda=" + atrim(vcomanda) + " and (((comandes.proximaseccio) = 'I' Or (comandes.proximaseccio) = 'L') And ((planificaciolam_operaris.ordre) > 0 And (planificaciolam_operaris.ordre) < 998)) ORDER BY planificaciolam_operaris.ordre;"
   
   Set rst = dbtmpb.OpenRecordset(vsql)
   If rst.EOF Then
      estaplanificada = False
       Else: estaplanificada = True
   End If
   ccomandalam.Tag = vcomanda
   Set rst = Nothing
End Function
Sub mirarestatFET(vcomanda As Long)
    Dim rst As Recordset
    Set rst = dbtmpb.OpenRecordset("select * from laminadora_bobinesasala where comanda=" + atrim(vcomanda))
    If rst.EOF Then
        etfetbobines = "Pndt"
          Else: etfetbobines = "Fet"
    End If
    Set rst = Nothing
End Sub

Private Sub Command1_Click()
  llistat_prepararcanutus True
End Sub

Private Sub Command2_Click()
   llistat_prepararcanutus False
End Sub

Private Sub Command3_Click()

    If lotperactivar = "" Then MsgBox "Has de possar un lot per afegir-lo", vbCritical, "Error": Exit Sub
   dbtmpb.Execute "delete * from canutusjatallats where comanda=" + atrim(cadbl(lotperactivar))
   dbtmpb.Execute "insert into canutusjatallats (comanda,agafarstd) values (" + atrim(lotperactivar) + ",True)"
   MsgBox "Comanda marcada com a canutus Tallats agafar Standards.", vbInformation, "Info"
End Sub

Private Sub command4_click()
   If ccomandalam = "" Then MsgBox "Has de possar una comanda per afegir-lo", vbCritical, "Error": Exit Sub
   If etfetbobines <> "Pndt" Then MsgBox "Aquesta comanda no està apunt.", vbCritical, "Error": GoTo fi
   dbtmpb.Execute "delete * from laminadora_bobinesasala where comanda=" + atrim(cadbl(ccomandalam.Tag))
   dbtmpb.Execute "insert into laminadora_bobinesasala (comanda) values (" + atrim(cadbl(ccomandalam.Tag)) + ")"
   MsgBox "Comanda marcada com a bobines a SALA (LAM).", vbInformation, "Info"
fi:
   mirar_estat_bobines
End Sub

Private Sub Command5_Click()
   dbtmpb.Execute "delete * from laminadora_bobinesasala where comanda=" + atrim(cadbl(ccomandalam.Tag))
   MsgBox "Comanda marcada com pendent de portar a SALA.", vbInformation, "Info"
   mirar_estat_bobines
End Sub

Private Sub Command6_Click()
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  emplenar_taula_bobines
  wait 2
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistat_bobinesLAM_PNDT_SALA.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "BAIXES.mdb"
  'oreport.Database.Tables.Item(2).Location = rutadelfitxer(cami) + "BAIXES.mdb"
  
  oreport.DiscardSavedData
'  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
 '   Else
 '     oreport.DisplayProgressDialog = False
 '     oreport.PrintOut False, 1
 ' End If
  
End Sub
Sub emplenar_taula_bobines()
  Dim rst As Recordset
  Dim rstllistat As Recordset
  Dim rst2 As Recordset
  Dim rstopcions As Recordset
  Dim rstextres As Recordset
  Dim rstbobinesasala As Recordset
  
  Dim vnumc2 As Double
  Dim vnumc3 As Double
  Dim vnumc As Double
  Dim vcomandap As Double
  Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
  
  dbtmpb.Execute "delete * from tmp_llistatdesembolicarlaminadora"
  
  Set rst = dbtmpb.OpenRecordset("select comanda,ordre,maquina from planificaciolam_operaris where ordre>0 and ordre<998 ")
  Set rstbobinesasala = dbtmpb.OpenRecordset("select * from laminadora_bobinesasala")
  If rst.EOF Then GoTo fi
  Set rstllistat = dbtmpb.OpenRecordset("select * from tmp_llistatdesembolicarlaminadora")
  
  While Not rst.EOF
    vnumc = cadbl(rst!comanda)
    Set rst2 = dbtmpb.OpenRecordset("SELECT comandes.proximaseccio, comandes.linkcomanda1, comandes.linkcomanda2,comandes.cantitatex, productes.ruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi Where comanda = " + atrim(vnumc))
    If rst2.EOF Then GoTo proxima
    vnumc2 = cadbl(rst2!linkcomanda1)
    vnumc3 = cadbl(rst2!linkcomanda2)
    vmetres = rst2!cantitatex
    rstbobinesasala.FindFirst "comanda=" + atrim(rst!comanda)
    If Not rstbobinesasala.NoMatch Then GoTo proxima
    If (rst2!proximaseccio <> "I" And rst2!proximaseccio <> "L") Then GoTo proxima
    vcomandap = vnumc
    
bucle:
    If vnumc > 0 Then
     If Not hihaimpresora(vnumc) Then
      Set rst2 = dbtmpb.OpenRecordset("SELECT Palets.codimatprognou, Parcials.*, Bobines.Sit FROM Bobines INNER JOIN (Palets INNER JOIN Parcials ON Palets.Idpalet = Parcials.idpalet) ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet) where comanda='" + atrim(vnumc) + "'")
      'es estoc?
      
      Set rstopcions = dbstocks.OpenRecordset("select * from opcionsdajust where comanda=" + atrim(vnumc))
      If Not rstopcions.EOF Then
             Set rsttmp = dbstocks.OpenRecordset("SELECT Palets.codimatprognou, Parcials.*, Bobines.Sit FROM Bobines INNER JOIN (Palets INNER JOIN Parcials ON Palets.Idpalet = Parcials.idpalet) ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet) where palets.idpalet in (select paletexemple from grupsdepalets where numerogrup=" + atrim(rstopcions!grupdestoc) + ")")
             If Not rsttmp.EOF Then Set rstmat = dbtmp.OpenRecordset("select * from [llistat materials] where codi=" + atrim(rsttmp!codimatprognou))
             rstllistat.AddNew
             rstllistat!comanda = vnumc
             rstllistat!comandaprincipal = vcomandap
             rstllistat!situacio = "E" + atrim(cadbl(rstopcions!grupdestoc)) + " - " + Format(vmetres, "#,##0") + "mtrs"
             rstllistat!ordre = rst!ordre
             rstllistat!palet = rsttmp!idpalet
             rstllistat!maquina = cadbl(rst!maquina)
             rstllistat!tipusbobina = "E"
             If Not rsttmp.EOF Then rstllistat!nommaterial = atrim(rstmat![familiesmaterials.descripcio]) + "-" + atrim(rstmat![subfamiliesmaterials.descripcio]) + "-" + atrim(rstmat![familiescolorants.descripcio]) + "-" + atrim(rstmat![subfamiliescolorants.descripcio])
             rstllistat.Update
             GoTo proxima
      End If

      'si entre en el while es que no es estoc i hi ha bobines
      While Not rst2.EOF
        Set rstmat = dbtmp.OpenRecordset("select * from [llistat materials] where codi=" + atrim(rst2!codimatprognou))
        rstllistat.AddNew
        rstllistat!comandaprincipal = vcomandap
        rstllistat!comanda = vnumc
        rstllistat!palet = rst2!idpalet
        rstllistat!bobina = rst2!idbobina
        rstllistat!ordre = rst!ordre
        rstllistat!maquina = cadbl(rst!maquina)
        rstllistat!situacio = atrim(rst2!sit)
           '30/5/24 En miralles ha demanat que surti sempre la SITUACIÓ no només quan sigui diferent de LAM
        'rstllistat!situacio = IIf(atrim(rst2!sit) <> "LAM", atrim(rst2!sit), "")
        rstllistat!nommaterial = atrim(rstmat![familiesmaterials.descripcio]) + "-" + atrim(rstmat![subfamiliesmaterials.descripcio]) + "-" + atrim(rstmat![familiescolorants.descripcio]) + "-" + atrim(rstmat![subfamiliescolorants.descripcio])
        rstllistat!tipusbobina = IIf(bobinesdentrada.esrestu(rst2!idpalet, rst2!idbobina), "R", "")
        If rstllistat!tipusbobina = "" Then rstllistat!tipusbobina = IIf(bobinesdentrada.esparcial(rst2!idpalet, rst2!idbobina), "P", "")
        If rstllistat!tipusbobina = "" Then rstllistat!tipusbobina = "Q"
        rstllistat.Update
        rst2.MoveNext
      Wend
     End If
    End If
    If vnumc2 > 0 Then vnumc = vnumc2: vnumc2 = 0: GoTo bucle
    If vnumc3 > 0 Then vnumc = vnumc3: vnumc3 = 0: GoTo bucle
proxima:
    rst.MoveNext
  Wend
fi:
  Set rst2 = Nothing
  Set rst = Nothing
  Set rstllistat = Nothing
End Sub
Function hihaimpresora(vnumc) As Boolean
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("SELECT comandes.proximaseccio, comandes.linkcomanda1, comandes.linkcomanda2, productes.ruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi Where comanda = " + atrim(vnumc))
  If Not rst.EOF Then
      If InStr(1, rst!ruta, "I") > 0 Then hihaimpresora = True
  End If
  Set rst = Nothing
End Function
Private Sub Command7_Click()
If lotperactivar = "" Then MsgBox "Has de possar un lot per afegir-lo", vbCritical, "Error": Exit Sub
   dbtmpb.Execute "delete * from canutusjatallats where comanda=" + atrim(cadbl(lotperactivar))
   dbtmpb.Execute "insert into canutusjatallats (comanda) values (" + atrim(lotperactivar) + ")"
   MsgBox "Comanda marcada com a canutus tallats.", vbInformation, "Info"
End Sub

Private Sub Command8_Click()
   dbtmpb.Execute "delete * from canutusjatallats where comanda=" + atrim(cadbl(lotperactivar))
   MsgBox "Comanda marcada com a canutus NO tallats.", vbInformation, "Info"
End Sub

 Sub llistat_prepararcanutus(vtallats As Boolean)
   Dim sql As String
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim rst3 As Recordset
   Dim vcomandaanterior As Double
   ratoli "espera"
   DoEvents
   borrartaulatmp_canutuspertallar
   formencarregat.Caption = "Seleccionant registres..."
   wait 2
   sql = "insert INTO tmp_canutuspertallar SELECT comandes.* FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.proximaseccio)<>'E' And (comandes.proximaseccio)<>'V' And (comandes.proximaseccio)<>'P' And (comandes.proximaseccio)<>'T' And (comandes.proximaseccio)<>'I' And (comandes.proximaseccio)<>'S') AND ((InStr(1,[productes].[ruta],'R'))>0)) OR (((comandes.proximaseccio)='I') AND ((InStr(1,[productes].[ruta],'R'))>0) AND ((comandes.comanda) In (select comanda from muntadora_ordremuntatge))) OR (((comandes.proximaseccio)='I') AND ((InStr(1,[productes].[ruta],'R'))>0) AND ((comandes.comanda) In (select comanda from muntadoratot where acabada)));"
   dbtmpb.Execute sql
   Set rst = dbtmp.OpenRecordset("select * from comandes where proximaseccio<>'T' and dataactivacio<>null")
   While Not rst.EOF
      Set rst3 = dbtmp.OpenRecordset("select tipusmaterialcanutureb from comandes_extres where comanda=" + atrim(rst!comanda))
      Set rst2 = dbtmp.OpenRecordset("select InStr(1,[productes].[ruta],'R') as tereb from productes where codi='" + atrim(rst!producte) + "'")
      If Not rst2.EOF Then
        If rst2!tereb > 0 And atrim(rst3!tipusmaterialcanutureb) = "P" Then
           dbtmpb.Execute "insert into tmp_canutuspertallar select * from comandes where comanda=" + atrim(rst!comanda)
        End If
      End If
      rst.MoveNext
   Wend
   'On Error GoTo fi
   dbtmpb.Execute "delete * from  canutusjatallats where comanda in (SELECT comandes.comanda FROM comandes RIGHT JOIN canutusjatallats ON comandes.comanda = canutusjatallats.comanda WHERE (((comandes.proximaseccio)='T')))"
   'dbtmpb.Execute "delete * from tmp_canutuspertallar where comanda in (SELECT DISTINCTROW First(tmp_canutuspertallar.comanda) AS comanda2 From tmp_canutuspertallar GROUP BY tmp_canutuspertallar.comanda HAVING (((Count(tmp_canutuspertallar.comanda))>1));)"
   formencarregat.Caption = "Preparant el llistat..."
   wait 2
   
   'dbtmpb.Execute "update tmp_canutuspertallar set seccioactual='*' where comanda in (select comanda from canutusjatallats)"
   If vtallats Then
       dbtmpb.Execute "delete * from  tmp_canutuspertallar where comanda in (select comanda from canutusjatallats)"
         Else: dbtmpb.Execute "delete * from  tmp_canutuspertallar where comanda not in (select comanda from canutusjatallats)"
   End If
   dbtmpb.Execute "delete * From tmp_canutuspertallar WHERE (((tmp_canutuspertallar.tubbase) Is Null)) OR (((tmp_canutuspertallar.tubbase)=0))"
   dbtmpb.Execute "delete * from tmp_canutuspertallar as t1 where amplereb in (select ample_Canutu from canutusestandard where mida_canutu=t1.tubbase)"
   Set rst = dbtmpb.OpenRecordset("select * from tmp_canutuspertallar order by comanda")
   dbtmpb.Execute "update tmp_canutuspertallar set seccioactual=''" 'borro tot el continut de seccioactual per mes avall utilitzarlo per passar el material del canutu
   While Not rst.EOF
     If vcomandaanterior = rst!comanda Then
       rst.Delete
       GoTo cont
     End If
     Set rst2 = dbtmp.OpenRecordset("select tipusmaterialcanutureb from comandes_Extres where comanda=" + atrim(rst!comanda))
     rst.Edit
     'aquest dos camps son de la taula temporal no de la PRINCIPAL'aprofito el camp seccioactual per guardar el tipus de canutu PVC o Cartró
     rst!seccioactual = atrim(rst2!tipusmaterialcanutureb)
     rst!rebobinadora = calcularcanutosnecessaris(rst)
     
     rst.Update
     vcomandaanterior = rst!comanda
cont:
     rst.MoveNext
   Wend
   formencarregat.Caption = "Llençant el llistat... "
   wait 4
   DoEvents
   'On Error Resume Next
   'Set rst = dbtmp.OpenRecordset(sql)
   'Form1.Caption = "Imprimint la bobina...."
   
   'llistat de 7
   'wait 2
    llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistatcanutusenfabricacio.rpt"
    llistat.Destination = crptToWindow
    'llistat.Destination = crptToPrinter
    llistat.CopiesToPrinter = 1
    llistat.DataFiles(0) = cami
   
    DoEvents
    'wait (2)
    For i = 1 To 10
      llistat.Formulas(i) = ""
    Next i
    'llistat.DiscardSavedData = True
     If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
     DoEvents
    llistat.Formulas(0) = "titolsinforme='" + "Llistat de canutus necessaris en fabricació." + "'"
    If Not vtallats Then
       llistat.Destination = crptToWindow
       llistat.WindowState = crptMaximized
       llistat.Formulas(0) = "titolsinforme='" + "Llistat de canutus que ja han estat tallats." + "'"
    End If
    'Form1.Caption = ""
    ratoli "normal"
    DoEvents
    llistat.Action = 1
fi:
   Set rst = Nothing
   Set rst2 = Nothing
   Set rst3 = Nothing
 End Sub
 Function calcularcanutosnecessaris(rst As Recordset) As Integer
  Dim vmetresdelabobina As Double
  Dim vmicres As Double
  
  vmetresdelabobina = cadbl(rst!mtrslinbob)
  'si no hi ha metres per la bobina calculo sobre diametre de 50cm
  If vmetresdelabobina = 0 Then
    vmicres = espesordelmaterial(rst)
    vmetresdelabobina = calcular_diametre(50, cadbl(rst!tubbase), vmicres)
  End If
  
  If vmetresdelabobina > 0 Then
   If ((cadbl(rst!rebmtrs) / vmetresdelabobina) - Int((cadbl(rst!rebmtrs) / vmetresdelabobina))) * 10 > 0 Then
      calcularcanutosnecessaris = Redondejar(((cadbl(rst!rebmtrs) / vmetresdelabobina) + 1) / cadbl(rst!simulteneitatreb), 0) * cadbl(rst!simulteneitatreb)
     Else: calcularcanutosnecessaris = Redondejar((cadbl(rst!rebmtrs) / vmetresdelabobina) / cadbl(rst!simulteneitatreb), 0) * cadbl(rst!simulteneitatreb)
   End If
     Else
  End If

 End Function
 
 Function espesordelmaterial(rstc As Recordset) As Double
  Dim rst As Recordset
  Dim vespesor As Double
  Dim vmesura As String
  Set rst = dbtmp.OpenRecordset("select comanda,mesuraesp,espessor,tubolam from comandes where comanda=" + atrim(rstc!comanda) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2))
  While Not rst.EOF
    If rst!comanda > 0 Then
      vespesor = vespesor + micresmaterial(cadbl(rst!mesuraesp), rst!espessor, rst!tubolam)
    End If
    rst.MoveNext
  Wend
  espesordelmaterial = vespesor
End Function
Function micresmaterial(codimesuralineal As Byte, espesor As Double, tubolam As String) As Double
  Dim rstmesural As Recordset
  Set rstmesural = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(codimesuralineal))
  r = ""
  If rstmesural.EOF Then Exit Function
  r = espesor
  If rstmesural!descripcio = "GALGUES" Then
            If tubolam = "T" Then
                 r = Redondejar(espesor / 4, 0)
                  Else: r = Redondejar(espesor / 2, 0)
            End If
  End If
  If InStr(1, rstmesural!descripcio, "GR/") > 0 Then
    r = espesor * -1
  End If
  micresmaterial = r
End Function
 
 Function calcular_diametre(diametreext As Double, canutu As Double, micres As Double) As Double
    Dim metres As Double
    diametreext = diametreext * 10 ' paso a milimetres
    canutu = canutu * 10 'paso a milimetres
    'calcul
    metres = ((diametreext * diametreext) - (canutu * canutu) / micres) * 0.746
    calcular_diametre = Redondejar(metres, 0)
 End Function
Sub borrartaulatmp_canutuspertallar()
  ' On Error Resume Next
  '  dbtmpb.Execute "drop table tmp_canutuspertallar"
  ' On Error GoTo 0
  dbtmpb.Execute "delete * from tmp_canutuspertallar"
End Sub

Private Sub Form_Activate()
 '   If Now < "23/01/2022" Then
 '      Frame2.Visible = True: DoEvents
 '      Frame2.Left = 0
 '      Frame2.Top = -60
 '   End If
End Sub

Private Sub Form_Load()
  'carregar_llista_muntatge True
  
  
  
End Sub

Sub carregar_llista_muntatge(Optional vcomprovarestatdelallista As Boolean)
   Dim rst As Recordset
   Dim rstc As Recordset
   Set rst = dbbaixes.OpenRecordset("select * from laminadora_ordremuntatge order by comanda")
   llistaproduccio.Clear
   While Not rst.EOF
     If vcomprovarestatdelallista Then
         Set rstc = dbtmp.OpenRecordset("select proximaseccio from comandes where comanda=" + atrim(cadbl(rst!comanda)))
         If Not rstc.EOF Then
            If InStr(1, "RSVPT", atrim(rstc!proximaseccio) + "  ") > 0 Then
               dbtmpb.Execute "DELETE * from laminadora_ordremuntatge where comanda=" + atrim(cadbl(rst!comanda))
               GoTo proxima
            End If
         End If
     End If
     llistaproduccio.AddItem atrim(cadbl(rst!comanda))
proxima:
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Private Sub mcanutusstandard_Click()
  Load formaltarep
  formaltarep.Caption = "Manteniment Tubos Estandad"
  'formaccessoris.autonum = "accessoris"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from canutusestandard order by ample_canutu"
  formaltarep.refrescar
  'formaltarep.DBGrid1.Columns(1).Caption = "T/A/C"
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(1).Caption = "Ample"
  formaltarep.DBGrid1.Columns(0).Caption = "Canutu"
  formaltarep.Show 1
End Sub

Private Sub mdonarllaunadebaixa_Click()
  Dim vnumllauna As String
  vnumllauna = InputBox("Escaneja el numero de llauna.", "Donar de baixa una llauna")
  donardebaixalallauna vnumllauna
End Sub

Sub donardebaixalallauna(vnumllauna As String)
  Dim rst As Recordset
  Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
  Set rst = dbtintes.OpenRecordset("select * from llaunes where numllauna='" + atrim(vnumllauna) + "'")
  If rst.EOF Then MsgBox "Aquesta llauna no existeix.", vbCritical, "Error": Exit Sub
  ferelretorndetinta vnumllauna, 0, True
  Set rst = dbtintes.OpenRecordset("select * from dadesllaunestotes where numllauna='" + atrim(vnumllauna) + "'")
  If Not rst.EOF Then
     If cadbl(rst!capacitatactual < 1) Then
       MsgBox "Contenidor/llauna Nº: " + vnumllauna + " ja l'he donada de baixa.", vbInformation, "Atenció"
     End If
  End If
  Set rst = Nothing
End Sub

Private Sub Timer1_Timer()
  Frame2.Visible = False
End Sub
