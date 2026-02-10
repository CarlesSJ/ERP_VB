VERSION 5.00
Begin VB.Form formpackinglist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Organitzar el PackingList"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7635
   Icon            =   "packinglist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   5490
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   7395
      Begin VB.CheckBox checkperpantalla 
         Caption         =   "Per pantalla"
         Height          =   210
         Left            =   5640
         TabIndex        =   14
         Top             =   5250
         Width           =   1635
      End
      Begin VB.CheckBox cdetallbobines 
         Caption         =   "Detall de Bobines"
         Height          =   210
         Left            =   5565
         TabIndex        =   12
         Top             =   405
         Width           =   1635
      End
      Begin VB.Frame Frame3 
         Caption         =   "Bases Automàtic"
         Height          =   870
         Left            =   5550
         TabIndex        =   10
         Top             =   630
         Width           =   1635
         Begin VB.CommandButton Command3 
            BackColor       =   &H0080FF80&
            Caption         =   "Auto-Assignar Bases"
            Height          =   510
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   1290
         End
      End
      Begin VB.CommandButton bimprimirpackinglist 
         Height          =   375
         Left            =   5790
         Picture         =   "packinglist.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Imprimir Packinglist"
         Top             =   4860
         Width           =   1290
      End
      Begin VB.Frame Frame1 
         Caption         =   "Bases Manualment"
         Height          =   1725
         Left            =   5580
         TabIndex        =   2
         Top             =   1500
         Width           =   1590
         Begin VB.CommandButton Command2 
            Caption         =   "Ajuntar Palets"
            Height          =   435
            Left            =   165
            TabIndex        =   7
            Top             =   1200
            Width           =   1290
         End
         Begin VB.CommandButton bmenysbase 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   945
            TabIndex        =   6
            Top             =   435
            Width           =   285
         End
         Begin VB.CommandButton bmesbase 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   945
            TabIndex        =   5
            Top             =   210
            Width           =   285
         End
         Begin VB.TextBox cnumerodebase 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   390
            TabIndex        =   4
            Text            =   "1"
            Top             =   225
            Width           =   555
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Assignar Base"
            Height          =   435
            Left            =   165
            TabIndex        =   3
            Top             =   675
            Width           =   1290
         End
      End
      Begin VB.ListBox llista 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4845
         Left            =   210
         MultiSelect     =   1  'Simple
         TabIndex        =   1
         Top             =   450
         Width           =   5235
      End
      Begin VB.Label etavis 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2490
         TabIndex        =   13
         Top             =   120
         Width           =   4200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Llista de Comandes-Palets"
         Height          =   300
         Left            =   345
         TabIndex        =   8
         Top             =   195
         Width           =   2115
      End
   End
End
Attribute VB_Name = "formpackinglist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bimprimirpackinglist_Click()
  If llista.tag <> "Relaciofeta" Then relacionarpaletsambbasesautomaticament: wait (1)
  If packinglist_generar_fitxer_temporal(cadbl(formvendes.cnumalbara)) Then
     imprimir_packinglist cadbl(formvendes.cnumalbara)
  End If
  
End Sub
Sub inizialitzaretiquetes(vector)
  Dim i As Byte
  i = 1
  vector(i, 1) = "BOBS": vector(i, 2) = "Reels": vector(i, 3) = "Bobines": vector(i, 4) = "Bobinas": vector(i, 5) = "Bobines": i = i + 1
  vector(i, 1) = "CAIXES/SACS": vector(i, 2) = "Boxes/Bags": vector(i, 3) = "Caisses/Sacs": vector(i, 4) = "Cajas/Bolsas": vector(i, 5) = "Caixes/Bosses": i = i + 1
  vector(i, 1) = "FECHAENVIO": vector(i, 2) = "Shipping date": vector(i, 3) = "Date de livraison": vector(i, 4) = "Fecha de envio": vector(i, 5) = "Data d´enviament": i = i + 1
  vector(i, 1) = "KILOS": vector(i, 2) = "Weight": vector(i, 3) = "Poids ": vector(i, 4) = "Peso": vector(i, 5) = "Pes": i = i + 1
  vector(i, 1) = "MTRS": vector(i, 2) = "Meters": vector(i, 3) = "Mètres": vector(i, 4) = "Metros": vector(i, 5) = "Metres": i = i + 1
  vector(i, 1) = "NPALET": vector(i, 2) = "Palet Nº": vector(i, 3) = "Palet Nº": vector(i, 4) = "NºPalet": vector(i, 5) = "NºPalet": i = i + 1
  vector(i, 1) = "PEDIDO": vector(i, 2) = " Order Nº": vector(i, 3) = "Commande Nº": vector(i, 4) = "Pedido Nº": vector(i, 5) = "Comanda Nº": i = i + 1
  vector(i, 1) = "PEDIDOCLI": vector(i, 2) = "Delivery Order No.": vector(i, 3) = "Bon de livraison Nº": vector(i, 4) = "Orden de entrega Nº": vector(i, 5) = "Ordre d'entrega": i = i + 1
  vector(i, 1) = "REFERENCIA": vector(i, 2) = "Ref.": vector(i, 3) = "Réf.": vector(i, 4) = "Ref": vector(i, 5) = "Ref": i = i + 1
  vector(i, 1) = "TBASES": vector(i, 2) = "T.Bases": vector(i, 3) = "T.Bases": vector(i, 4) = "T.Bases": vector(i, 5) = "T.Bases": i = i + 1
  vector(i, 1) = "TBOBS": vector(i, 2) = "T.Reels": vector(i, 3) = "T.Bobines": vector(i, 4) = "T.Bobinas": vector(i, 5) = "T.Bobinas": i = i + 1
  vector(i, 1) = "TCAIXES": vector(i, 2) = "T.Boxes": vector(i, 3) = "T.Caisses": vector(i, 4) = "T.Cajas": vector(i, 5) = "T.Caixes": i = i + 1
  vector(i, 1) = "TILOS": vector(i, 2) = "T.Weight": vector(i, 3) = "T.Poids ": vector(i, 4) = "T.Peso": vector(i, 5) = "T.Pes": i = i + 1
  vector(i, 1) = "TMETROS": vector(i, 2) = "T.Meters": vector(i, 3) = "T.Mètres": vector(i, 4) = "T.Metros": vector(i, 5) = "T.Metres": i = i + 1
  vector(i, 1) = "TOTAL": vector(i, 2) = "Total": vector(i, 3) = "Total": vector(i, 4) = "Total": vector(i, 5) = "Total": i = i + 1
  vector(i, 1) = "TPALETS": vector(i, 2) = "T.Palets": vector(i, 3) = "T.Palets": vector(i, 4) = "T.Palets": vector(i, 5) = "T.Palets": i = i + 1
  vector(i, 1) = "TUNIDADES": vector(i, 2) = "T.Pcs": vector(i, 3) = "T.Pcs": vector(i, 4) = "T.Pzs": vector(i, 5) = "T.Pcs": i = i + 1
  vector(i, 1) = "UNIDADES": vector(i, 2) = "Pcs": vector(i, 3) = "Pcs": vector(i, 4) = "Pzs": vector(i, 5) = "Pcs": i = i + 1
  vector(i, 1) = "KILOSNETOS": vector(i, 2) = "Net Weight": vector(i, 3) = "Poids net": vector(i, 4) = "Peso neto": vector(i, 5) = "Pes net": i = i + 1
  vector(i, 1) = "-": i = i + 1
  
End Sub


Sub possar_etiquetesidioma(oreport As CRAXDDRT.Report, idiomaclient As String)
  Dim i As Byte
  Dim colidioma As Byte
  Dim vector(100, 5)
  inizialitzaretiquetes vector
  colidioma = 2
  If idiomaclient = "ES" Then colidioma = 4
  If idiomaclient = "FR" Then colidioma = 3
  i = 1
  While vector(i, 1) <> "-"
     oreport.FormulaFields.GetItemByName("e" + LCase(vector(i, 1))).Text = "'" + treure_apostruf(vector(i, colidioma)) + "'"
     i = i + 1
  Wend
  
End Sub
Sub llistatpackinglist_possar_dades_formules(oreport As CRAXDDRT.Report, numc As Double)
   Dim rst As Recordset
   
   Set rst = dbcomandes.OpenRecordset("SELECT Clients_envios.nome, Clients_envios.domicilie, Clients_envios.codipostale, Clients_envios.poblacioe, Clients_envios.provinciae FROM Clients_envios  where id=" + atrim(cadbl(formvendes.datacapcalera.Recordset!id_direnvio)))
   If Not rst.EOF Then
     oreport.FormulaFields.GetItemByName("nomclient").Text = "'" + atrim(rst!nome) + "'"
     oreport.FormulaFields.GetItemByName("direccioclient").Text = "'" + atrim(rst!domicilie) + "'"
     oreport.FormulaFields.GetItemByName("poblacio").Text = "'" + atrim(rst!codipostale) + "-" + atrim(rst!poblacioe) + "'"
     oreport.FormulaFields.GetItemByName("provincia").Text = "'" + atrim(rst!provinciae) + "'"
     oreport.FormulaFields.GetItemByName("dataenviament").Text = "'" + atrim(formvendes.datacapcalera.Recordset!dataalbara) + "'"
     oreport.FormulaFields.GetItemByName("detallbobina").Text = "'" + atrim(cdetallbobines.Value) + "'"
     'oreport.FormulaFields.GetItemByName("dataenviament").Text = "'0'"
     oreport.FormulaFields.GetItemByName("GTotalpalets").Text = "'" + buscareltotalpackinglist("P") + "'"
     oreport.FormulaFields.GetItemByName("GTotalCaixes").Text = "'" + buscareltotalpackinglist("C") + "'"
     oreport.FormulaFields.GetItemByName("GTotalBobines").Text = "'" + buscareltotalpackinglist("B") + "'"
   End If
   Set rst = Nothing
End Sub
Function buscareltotalpackinglist(v As String) As String
   Dim dbtmp As Database
   Dim rsttmp As Recordset
   Set dbtmp = DBEngine.OpenDatabase("c:\temp\~llistatpacking.mdb")
   If v = "P" Then
      Set rsttmp = dbtmp.OpenRecordset("select distinct numpaletentrega,first(comandapaletentrega) from bobinesent group by comandapaletentrega,numpaletentrega")
      If Not rsttmp.EOF Then
        rsttmp.MoveLast
        buscareltotalpackinglist = atrim(rsttmp.RecordCount)
      End If
   End If
   If v = "C" Then
      Set rsttmp = dbtmp.OpenRecordset("select * from bobinesent where tipusproducte='Caixes'")
      If Not rsttmp.EOF Then
        rsttmp.MoveLast
        buscareltotalpackinglist = atrim(rsttmp.RecordCount)
      End If
   End If
   If v = "B" Then
      Set rsttmp = dbtmp.OpenRecordset("select * from bobinesent where tipusproducte='Bobines'")
      If Not rsttmp.EOF Then
        rsttmp.MoveLast
        buscareltotalpackinglist = atrim(rsttmp.RecordCount)
      End If
   End If
   
   Set rsttmp = Nothing
   Set dbtmp = Nothing
End Function
Sub imprimir_packinglist(numc As Double)
 
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "packinglistexpedicionsstd.rpt", 1)
 ' oreport.SQLQueryString = ""
'  oreport.RecordSelectionFormula = "{capcaleraalbara.numalbara}=" + atrim(datacapcalera.Recordset!numalbara)
'  oreport.FormulaFields.GetItemByName("nomdirenvio").Text = "'" + treure_apostruf(etinfodelclient.tag) + "'"
'  oreport.SQLQueryString = ""
 ' oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "vendes.mdb"
 ' oreport.Database.Tables.Item(2).Location = rutadelfitxer(cami) + "vendes.mdb"
  possar_etiquetesidioma oreport, idiomaclient
  llistatpackinglist_possar_dades_formules oreport, numc
  oreport.DiscardSavedData
  oreport.VerifyOnEveryPrint = False
  
  
  
  If existeix("c:\ordprog.ini") Or checkperpantalla.Value = 1 Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
    Else
      oreport.PrintOut False, 1
  End If
  

End Sub
Function packinglist_generar_fitxer_temporal(numalb As Double) As Boolean
  Dim vnomfitxer As String
  Dim dbtmp As Database
  Dim rstc As Recordset
  Dim rsttmp As Recordset
  Dim rsttmp2 As Recordset
  Dim vsumapeces As Double
  vnomfitxer = "c:\temp\~llistatpacking.mdb"
  If Not espoteliminar(vnomfitxer) Then MsgBox "No es pot generar el llistat, mira que no estigui obert el llistat o la impresió pendent", vbCritical, "Error": Exit Function
  Set dbtmp = DBEngine.OpenDatabase(vnomfitxer)
  dbbaixes.Execute "select * into bobinesent IN '" + vnomfitxer + "' from bobinesent where numalbara=" + atrim(numalb)
  dbtmp.Execute "alter table bobinesent  add column refclient text"
  dbtmp.Execute "alter table bobinesent  add column comandaclient text"
  dbtmp.Execute "alter table bobinesent  add column tipusproducte text"
  dbtmp.Execute "alter table bobinesent  add column peces double"
  dbtmp.Execute "alter table bobinesent  add column marcailinia text"
  Set rsttmp = dbtmp.OpenRecordset("select distinct comanda from bobinesent")
  While Not rsttmp.EOF
     Set rstc = dbcomandes.OpenRecordset("select refclient,comandaclient,marcailinia from comandes where comanda=" + atrim(cadbl(rsttmp!comanda)))
     If Not rstc.EOF Then dbtmp.Execute "update bobinesent set marcailinia='" + atrim(rstc!marcailinia) + "',refclient='" + atrim(rstc!refclient) + "',comandaclient='" + atrim(rstc!comandaclient) + "' where comanda=" + atrim(rsttmp!comanda)
     rsttmp.MoveNext
  Wend
  Set rsttmp = formvendes.datacapcalera.Database.OpenRecordset("select * from liniesalbara where numalbara=" + atrim(numalb))
  While Not rsttmp.EOF
     Set rsttmp2 = dbtmp.OpenRecordset("select * from bobinesent where comanda=" + atrim(cadbl(rsttmp!lotinplacsa)))
     vsumapeces = 0
     While Not rsttmp2.EOF
        rsttmp2.Edit
        If atrim(rsttmp2!seccio) = "R" And cadbl(rsttmp!metreslineals) > 0 Then
            rsttmp2!tipusproducte = "Bobines"
            rsttmp2!peces = Redondejar((cadbl(rsttmp!unitats) / cadbl(rsttmp!metreslineals)) * cadbl(rsttmp2!metresisacs), 0)
            If (vsumapeces + cadbl(rsttmp2!peces)) > cadbl(rsttmp!unitats) Then rsttmp2!peces = cadbl(rsttmp!unitats) - vsumapeces
            vsumapeces = vsumapeces + cadbl(rsttmp2!peces)
        End If
        If atrim(rsttmp2!seccio) = "S" Then
            rsttmp2!tipusproducte = atrim(rsttmp!tipusproducte)
            rsttmp2!peces = cadbl(rsttmp2!metresisacs)
            rsttmp2!metresisacs = 0
        End If
        rsttmp2.Update
        rsttmp2.MoveNext
     Wend
     rsttmp.MoveNext
  Wend
  packinglist_generar_fitxer_temporal = True
  Set rstc = Nothing
  Set rsttmp = Nothing
  Set dbtmp = Nothing
End Function
Function espoteliminar(vnomfitxer As String) As Boolean
   espoteliminar = True
   On Error GoTo fi
   If existeix(vnomfitxer) Then Kill vnomfitxer
   DBEngine.CreateDatabase vnomfitxer, dbLangSpanish
   Exit Function
fi:
   espoteliminar = False
End Function

Private Sub bmenysbase_Click()
    cnumerodebase = cadbl(cnumerodebase) - 1
    If cnumerodebase < 0 Then cnumerodebase = "0"
End Sub

Private Sub bmesbase_Click()
  cnumerodebase = cadbl(cnumerodebase) + 1
End Sub

Private Sub Command1_Click()
   Dim i As Integer
   Dim vnumc As Double
   Dim vpalet As Integer
   If llista.ListCount = 0 Then Exit Sub
   ratoli "espera"
   For i = 0 To llista.ListCount - 1
     If llista.Selected(i) Then
       vnumc = cadbl(Mid(llista.List(i), 1, InStr(1, llista.List(i), "-") - 1))
       vpalet = cadbl(Mid(llista.List(i), InStr(1, llista.List(i), "-") + 1, 2))
       dbbaixes.Execute "update bobinesent set comandapaletentrega=" + atrim(vnumc) + ",numpaletentrega=" + atrim(vpalet) + ",numdebaseentrega=" + atrim(cadbl(cnumerodebase)) + " where comanda=" + atrim(vnumc) + " and numpalet=" + atrim(cadbl(vpalet))
       dbbaixes.Execute "update bobinesent set numdebaseentrega=" + atrim(cadbl(cnumerodebase)) + " where comandapaletentrega=" + atrim(vnumc) + " and numpaletentrega=" + atrim(cadbl(vpalet))
     End If
   Next i
   carregar_llistapalets cadbl(formvendes.datacapcalera.Recordset!numalbara)
   ratoli "normal"
End Sub

Private Sub Command2_Click()
   Dim vnumc As String
   Dim vpalet As String
   Dim vnumc_origen As String
   Dim vpalet_origen As String
   Dim i As Integer
   If llista.ListIndex = -1 Then Exit Sub
   i = llista.ListIndex
   vnumc_origen = cadbl(Mid(llista.List(i), 1, InStr(1, llista.List(i), "-") - 1))
   vpalet_origen = cadbl(Mid(llista.List(i), InStr(1, llista.List(i), "-") + 1, 2))
   If vnumc_origen = 0 Or vpalet_origen = 0 Then Exit Sub
   vnumc = cadbl(InputBox("Entra la comanda on vols ajuntar aquest palet", "Ajuntar palet"))
   If vnumc = 0 Then Exit Sub
   vpalet = cadbl(InputBox("Entra el numero de palet d'aquesta comanda on ajuntar-lo", "Ajuntar palet"))
   If vpalet = 0 Then Exit Sub
   dbbaixes.Execute "update bobinesent set comandapaletentrega=" + atrim(vnumc) + ", numpaletentrega=" + atrim(cadbl(vpalet)) + " where comanda=" + atrim(vnumc_origen) + " and numpalet=" + atrim(vpalet_origen)
   carregar_llistapalets cadbl(formvendes.datacapcalera.Recordset!numalbara)
End Sub

Private Sub Command8_Click()

End Sub

Private Sub Command3_Click()
  relacionarpaletsambbasesautomaticament
End Sub
Sub relacionarpaletsambbasesautomaticament()
   Dim nbase As Integer
   Dim i As Integer
   Dim vnumc As Double
   Dim vpalet As Integer
   If llista.tag = "Relaciofeta" Then If MsgBox("Ja tens una relació feta vols continuar i eliminar-ho?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   nbase = 1
   If llista.ListCount = 0 Then Exit Sub
   ratoli "espera"
   For i = 0 To llista.ListCount - 1
       vnumc = cadbl(Mid(llista.List(i), 1, InStr(1, llista.List(i), "-") - 1))
       vpalet = cadbl(Mid(llista.List(i), InStr(1, llista.List(i), "-") + 1, 2))
       dbbaixes.Execute "update bobinesent set comandapaletentrega=" + atrim(vnumc) + ",numpaletentrega=" + atrim(vpalet) + ",numdebaseentrega=" + atrim(cadbl(nbase)) + " where comanda=" + atrim(vnumc) + " and numpalet=" + atrim(cadbl(vpalet))
       dbbaixes.Execute "update bobinesent set numdebaseentrega=" + atrim(cadbl(nbase)) + " where comandapaletentrega=" + atrim(vnumc) + " and numpaletentrega=" + atrim(cadbl(vpalet))
       nbase = nbase + 1
   Next i
   carregar_llistapalets cadbl(formvendes.datacapcalera.Recordset!numalbara)
   ratoli "normal"
End Sub

Private Sub Form_Load()
  carregar_llistapalets cadbl(formvendes.datacapcalera.Recordset!numalbara)
  carregar_altresdades
End Sub
Sub carregar_altresdades()
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("select packinglistalbara from clients_envios where id=" + atrim(formvendes.datacapcalera.Recordset!id_direnvio))
  If Not rst.EOF Then
     If atrim(rst!packinglistalbara) = "Cap" Then etavis = "Sense Packing-List"
     If atrim(rst!packinglistalbara) = "Detal Bobina per Bobina" Then etavis = "": cdetallbobines.Value = 1
     If atrim(rst!packinglistalbara) = "Totalitzat" Then etavis = "": cdetallbobines.Value = 0
  End If
  
  
End Sub
Function primerseleccionat() As Integer
   Dim i As Integer
   i = 0
   While i < llista.ListCount - 1
      If llista.Selected(i) Then primerseleccionat = i: GoTo fi
      i = i + 1
   Wend
fi:
End Function
Sub carregar_llistapalets(vnumalbara As Double)
   Dim rst As Recordset
   Dim rstp As Recordset
   Dim vdesti As String
   Dim vindex As Integer
   Set rst = dbbaixes.OpenRecordset("select distinct numpalet,comanda from bobinesent where numalbara=" + atrim(vnumalbara) + " order by comanda")
   vindex = primerseleccionat
   llista.Clear
   llista.tag = ""
   While Not rst.EOF
      Set rstp = dbbaixes.OpenRecordset("select * from bobinesent where comanda=" + atrim(rst!comanda) + " and numpalet=" + atrim(rst!numpalet))
      If Not rstp.EOF Then
         desti = ""
         If cadbl(rstp!comandapaletentrega) = cadbl(rst!comanda) And cadbl(rstp!numpaletentrega) = cadbl(rstp!numpalet) Then
             desti = "Base " + atrim(cadbl(rstp!numdebaseentrega))
             If cadbl(rstp!numdebaseentrega) = 0 Then desti = ""
               Else:
                 If cadbl(rstp!comandapaletentrega) > 0 Then
                    desti = "P: " + atrim(cadbl(rstp!comandapaletentrega)) + "-" + atrim(cadbl(rstp!numpaletentrega))
                 End If
         End If
         If desti <> "" Then
            llista.tag = "Relaciofeta"
             Else: llista.tag = ""
         End If
         llista.AddItem atrim(rst!comanda) + "-" + Format(rst!numpalet, "00") + IIf(desti <> "", " -> " + desti, "")
      End If
      rst.MoveNext
   Wend
   If vindex > 0 Then
     llista.ListIndex = vindex
     llista.Selected(vindex) = True
   End If
   Set rst = Nothing
   Set rstp = Nothing
End Sub

