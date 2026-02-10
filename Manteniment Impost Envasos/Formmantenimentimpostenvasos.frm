VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00EAD9CE&
   Caption         =   "Manteniment Impost Envasos"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   18720
   Icon            =   "Formmantenimentimpostenvasos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   18720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bfacturescompres 
      BackColor       =   &H80000004&
      Caption         =   "Fac. Compres"
      Height          =   300
      Left            =   15990
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   525
      Width           =   1110
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   8160
      Left            =   75
      TabIndex        =   14
      Top             =   1185
      Width           =   18570
      _ExtentX        =   32755
      _ExtentY        =   14393
      _Version        =   393216
      AllowBigSelection=   0   'False
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.CheckBox checknocomptarmermes 
      BackColor       =   &H00EAD9CE&
      Caption         =   "No comptar mermes"
      Height          =   195
      Left            =   5670
      TabIndex        =   12
      Top             =   60
      Width           =   1845
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H006BEBB1&
      Caption         =   "Presentar CSV(592) o A22"
      Height          =   300
      Left            =   15990
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   855
      Width           =   2565
   End
   Begin VB.CommandButton bguardarCSV 
      Caption         =   "Crear CSV"
      Enabled         =   0   'False
      Height          =   675
      Left            =   17145
      Picture         =   "Formmantenimentimpostenvasos.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   150
      Width           =   1410
   End
   Begin VB.ComboBox combomes 
      Height          =   315
      ItemData        =   "Formmantenimentimpostenvasos.frx":0B14
      Left            =   1710
      List            =   "Formmantenimentimpostenvasos.frx":0B3F
      TabIndex        =   8
      Top             =   15
      Width           =   1950
   End
   Begin VB.TextBox climitperiode 
      Height          =   285
      Left            =   -105
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   120
      Picture         =   "Formmantenimentimpostenvasos.frx":0BA1
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   780
      Width           =   540
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crear el resum Espanya e Importació    A22"
      Height          =   390
      Left            =   3450
      TabIndex        =   2
      Top             =   375
      Width           =   4050
   End
   Begin VB.Data dataintracomunitari 
      Caption         =   "dataintracomunitari"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "taula_impost"
      Top             =   765
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear el resum Intracomunitari 592"
      Height          =   390
      Left            =   330
      TabIndex        =   1
      Top             =   345
      Width           =   2895
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Formmantenimentimpostenvasos.frx":11BB
      Height          =   8085
      Left            =   120
      OleObjectBlob   =   "Formmantenimentimpostenvasos.frx":11D9
      TabIndex        =   0
      Top             =   1170
      Visible         =   0   'False
      Width           =   18510
   End
   Begin VB.Label ettrimestre 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3750
      TabIndex        =   13
      Top             =   75
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label etsumakilosmermaivendes 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00ED823A&
      Height          =   1035
      Left            =   7590
      TabIndex        =   11
      Top             =   120
      Width           =   10425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Limit del periode:"
      Height          =   240
      Left            =   420
      TabIndex        =   7
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label etsumakilos 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005C31DD&
      Height          =   300
      Left            =   10995
      TabIndex        =   4
      Top             =   825
      Width           =   5010
   End
   Begin VB.Label etllistat 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   690
      TabIndex        =   3
      Top             =   765
      Width           =   6825
   End
   Begin VB.Menu mhistorics 
      Caption         =   "Historics"
      Begin VB.Menu mhistoricintra 
         Caption         =   "Historic Intracomunitaries 592"
      End
      Begin VB.Menu mhisImpEsp 
         Caption         =   "Historic Importació i Espanya  A22"
      End
   End
   Begin VB.Menu mresumkgimpost 
      Caption         =   "Resum Kg Impost Comprats-Venuts"
   End
   Begin VB.Menu mresummermes 
      Caption         =   "Resum mermes i tan%"
   End
   Begin VB.Menu mgenerarequeriment 
      Caption         =   "Generar Requeriment"
   End
   Begin VB.Menu mabonamentsclients 
      Caption         =   "Abonaments Factures Clients amb Impost"
   End
   Begin VB.Menu mjustificants 
      Caption         =   "Justificants Merma"
   End
   Begin VB.Menu tanxcentmermes 
      Caption         =   "Llistat % de mermes entre dates"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vNomFitxerControlFacturesCompra As String
Dim vCopiarFacturesCompraiVenta As Boolean
Dim vRutaFacturesPdfSap As String
Dim vRutaFacturesCompres As String
Dim vRutaFacturesVendes As String
Dim vmsgERROR As String
Dim vsumakilos As Double
Dim vsumakilos2 As Double
Dim vcomandescomptades As String
Dim dbsap As Database
Dim vtipusllistat As String
Dim vMermesNoAfegides As String

Private Sub bfacturescompres_Click()
  Dim rst As Recordset
  Dim rstp As Recordset
  Dim rstc As Recordset
  Dim rstsap As Recordset
  Dim vfact As String
  Dim vfactcompres As String
  Dim vtipusproveidor As String
  Dim vsql As String
  
  vtipusproveidor = "='Intracomunitari'"
  If InStr(1, etllistat, "A22") > 0 Then vtipusproveidor = "<>'Intracomunitari'"
  
  Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
  Set rst = dbtmp.OpenRecordset("select * from taula_impost where lotinplacsa>200000 and apuntperdeclarar=true order by concepte,data")
  If Not rst.EOF Then rst.MoveLast: rst.MoveFirst
  ratoli "rellotge"
  vsql = "SELECT distinct parcials.idpalet as numpalet FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi "

  While Not rst.EOF
     Set rstp = dbstocks.OpenRecordset(vsql + " where comanda='" + atrim(rst!lotinplacsa) + "' and proveidors.tipusproveidorIMPOST" + vtipusproveidor)
     While Not rstp.EOF
        Set rstc = dbstocks.OpenRecordset("select numalbaraprov,data from albaransbip where kgimpostenvasos>0 and kgimpostenvasos<>null and numpalet=" + atrim(rstp!numpalet))
        If Not rstc.EOF Then
            Set rstsap = dbtmp.OpenRecordset("select * from importada_albarans_compres_inplacsa where year(docduedate)=" + atrim(Year(rstc!Data)) + " and numatcard='" + atrim(rstc!numalbaraprov) + "' order by numentrada desc")
            If Not rstsap.EOF Then
                vfact = atrim(rstsap!facturaprov)
                If atrim(rstsap!DUA) <> "" Then vfact = atrim(rstsap!DUA)
                If InStr(1, vfactcompres, vfact) = 0 Then vfactcompres = vfactcompres + vfact + ";" + atrim(rstsap!docduedate) + ";" + atrim(rstsap!nomproveidor) + ";" + Format(cadbl(rstsap!doctotal), "#,##0.00") + vbNewLine
            End If
        End If
        rstp.MoveNext
     Wend
     Me.Caption = atrim(rst.AbsolutePosition) + " / " + atrim(rst.RecordCount)
     DoEvents
     rst.MoveNext
  Wend
  ratoli "normal"
  If vfactcompres <> "" Then
    Open "c:\temp\factures compres A22.csv" For Output As 1
    Print #1, "Factures Compres Espanya i Importació"
    Print #1, vfactcompres
    Close #1
    If existeix("c:\temp\factures compres a22.csv") Then obrir_document "c:\temp\factures compres a22.csv"
    'Clipboard.SetText "Factures Compres Espanya i Importació" + vbNewLine + vfactcompres: MsgBox "Factures de compres copiades al portapapers."
  End If
  Set rst = Nothing
  Set rstc = Nothing
  Set rstp = Nothing
  Me.Caption = "Manteniment Impost Envasos"
End Sub

Private Sub bguardarCSV_Click()
  Dim vassentament As Long
  Dim rst As Recordset
  Dim vlinia As String
  Dim vnomfitxerCSV As String
  Dim vdatafitxer As Date
  If Not IsDate(etllistat.Tag) Then MsgBox "Per fer el CSV has de carregar un Historic.", vbCritical, "Error": Exit Sub
  vdatafitxer = etllistat.Tag
  vnomfitxerCSV = "c:\temp\CSVpresentacióImpostEnvasos.csv"
  Set rst = dbtmp.OpenRecordset("select * from taula_impost where apuntperdeclarar order by concepte,data")
  Open vnomfitxerCSV For Output As #1
  vassentament = 1
  Print #1, "Número Asiento;Fecha Hecho Contabilizado;Concepto;Clave Producto;Descripción Producto;Régimen Fiscal;Justificante;Prov./Dest.: Tipo Documento;Prov./Dest.: Nº documento;Prov./Dest.: Razón social;Kilogramos;Kilogramos No Reciclados;Observaciones"
  While Not rst.EOF
    If rst!kilosnoreciclats > 0 Then
        carregar_dades_a_linia vlinia, vassentament, rst, vdatafitxer
        Print #1, vlinia
        vassentament = vassentament + 1
    End If
    rst.MoveNext
  Wend
  Close #1
  
  If existeix(vnomfitxerCSV) Then obrir_document vnomfitxerCSV
  Set rst = Nothing
End Sub
Sub carregar_dades_a_linia(vlinia As String, vassentament As Long, rst As Recordset, vdatafitxer As Date)
  Dim vdata As Date
  'vlinia = atrim(vassentament) + ";"
  vlinia = atrim(rst!numassentament) + ";"
  vdata = IIf(Not IsNull(rst!datacomptableSAP), rst!datacomptableSAP, rst!Data)
  If rst!concepte = 3 And Month(vdata) <> Month(vdatafitxer) Then vdata = vdatafitxer
  vlinia = vlinia + Format(vdata, "dd/mm/yyyy") + ";"
  vlinia = vlinia + atrim(rst!concepte) + ";"
  vlinia = vlinia + atrim(rst!clauproducte) + ";"
  vlinia = vlinia + "FILM;"
  vlinia = vlinia + IIf(cadbl(rst!concepte) = 1, atrim(rst!regimfiscal), "") + ";"
  vlinia = vlinia + atrim(rst!justificant) + ";"
  If rst!concepte <> 3 Then   'si es merma no volen que es possin aquests camps
        vlinia = vlinia + "2;"
        vlinia = vlinia + atrim(rst!nifdestinatari) + ";"
        vlinia = vlinia + substituir(atrim(rst!nomdestinatari), "´", "'") + ";"
        Else: vlinia = vlinia + ";;;"
  End If
  vlinia = vlinia + passaradecimal(atrim(rst!kilos)) + ";"
  vlinia = vlinia + passaradecimal(atrim(rst!kilosnoreciclats)) + ";"
  vlinia = vlinia + atrim(rst!observacions)
End Sub

Private Sub combomes_Click()
   climitperiode = Format(DateAdd("d", -1, "01/" + Format(combomes.ItemData(combomes.ListIndex), "00") + "/" + Trim(Year(Now))), "dd/mm/yyyy")
   ettrimestre = Format(DateAdd("d", -1, "01/" + Format(combomes.ItemData(combomes.ListIndex), "00") + "/" + Trim(Year(Now))), "q") + " Trimestre"
   ettrimestre.Tag = Format(DateAdd("d", -1, "01/" + Format(combomes.ItemData(combomes.ListIndex), "00") + "/" + Trim(Year(Now))), "q")
End Sub

Private Sub Command1_Click()
  'posso a zero el comptador de kgassignats a cada factura TEMPORAL PER AQUEST CALCUL
  dbtmp.Execute "update facturesSAPreciclatge set kgtemporals=0"
  ettrimestre.Visible = False
  vsumakilos = 0: vsumakilos2 = 0
  vtipusllistat = "I"
  etllistat = "Llistat de compres, vendes i mermes INTRACOMUNITARIES"
  etllistat.Tag = "I"
  crear_resum_compres_intracomunitaries
  etsumakilos = "Compres IntraCom. Total_Kg: " + atrim(vsumakilos) + "   Base_Imposable: " + atrim(vsumakilos2)
  crear_resum_ventes_intracomunitaries
  crear_resum_abonos_intracomunitaris
  crear_resum_abonos_RECILAR_INTRA
  If checknocomptarmermes <> 1 Then
     crear_resum_parcials_100i300_Intra
     crear_resum_mermes_vendeS_intracomunitaries
     'crear_resum_parcials_300_Devolucions "Intracomunitari"
  End If

  bguardarCSV.Enabled = False
  bfacturescompres.Enabled = False
  sumar_mermes_i_vendes
  dataintracomunitari.Refresh
  configurar_reixa
  emplenar_reixa
End Sub
Sub crear_resum_abonos_RECILAR_INTRA()
   Dim rst As Recordset
   Dim rstabonos As Recordset
   Dim rstimpost As Recordset
   Dim rstfactura As Recordset
   Dim rstc As Recordset
   Dim vwere As String
   Dim vsql As String
   Dim vkgcapa As Double
   Dim vnumfactura As String
   Dim vcolormerma As String
   Dim vtotsSI As Boolean
   Dim vnomproveidor As String
   Dim vnifproveidor As String
   Dim vkgvendaINTRA As Double
   
   Set rstabonos = dbtmp.OpenRecordset("select * from abonosclients where numremesadestruccio_592=0 or numremesadestruccio_592=null")
   Set rstimpost = dbtmp.OpenRecordset("select * from taula_impost")
   
   While Not rstabonos.EOF
           'vwere = " (PaisVenta<>'ES' or regimfiscal<>'') and KgVentaAd_Intracom>0 and (((ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Intra) Is Null Or (ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Intra)=0)) "
           Set rstc = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(rstabonos!lotinplacsa))
           vwere = " (impostenvasos.comanda=" + atrim(rstc!comanda) + " or impostenvasos.comanda=" + atrim(rstc!linkcomanda1) + " or impostenvasos.comanda=" + atrim(rstc!linkcomanda2) + ") "
           vsql = "SELECT ImpostEnvasos.*, capcaleraalbara.*, (select distinct datafactura from Importada_LiniesFacturesSAP_Inplacsa where NumFact=[capcaleraalbara].[numfacturaSAP] ) AS Datafactura, materials.tanpercentimpostenvasos, clients.nom AS nomclient, clients_codisSAP.nif "
           vsql = vsql + " FROM (((ImpostEnvasos LEFT JOIN capcaleraalbara ON ImpostEnvasos.numalbara = capcaleraalbara.numalbara) LEFT JOIN (comandes LEFT JOIN materials ON comandes.materialex = materials.codi) ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN clients_codisSAP ON capcaleraalbara.codiclient = clients_codisSAP.codiSAP "
           Set rst = dbtmp.OpenRecordset(vsql + " where " + vwere)
           vtotsSI = True
           While Not rst.EOF
               vkgvendaINTRA = cadbl(rst!kgventaad_intracom)
               If vkgvendaINTRA > 0 Then
                vkgcapa = (totalKgtoteslescapes(rst!comanda, "592") * 100) / (totalKgtoteslescapes(rst!comanda, "A22") + totalKgtoteslescapes(rst!comanda, "592")) 'saber el % TOTAL
                vtotalimpost = (rstabonos!totaimpost) * (vkgcapa / 100)
                vkgcapa = (vkgvendaINTRA * 100) / totalKgtoteslescapes(rst!comanda, "592") 'saber el % DE CAPA
               ' If rstabonos!lotinplacsa = 212407 Then Stop
                vkgcapa = vtotalimpost * (vkgcapa / 100)
                
                rstimpost.AddNew
                rstimpost!concepte = 3
                rstimpost!clauproducte = "B"
                vcolormerma = rstabonos!colorreciclat
                rstimpost!kilos = Redondejar((vkgcapa * ((rst!tanpercentimpostenvasos / 100) + 1)) - vkgcapa, 3)
                rstimpost!kilosnoreciclats = Redondejar(vkgcapa, 3) 'Redondejar(rstabonos!totaimpost, 3)
                rstimpost!Data = dataFACTURA_REC(IIf(IsNull(rst!datafactura), 0, rst!datafactura), vcolormerma, vnumfactura, cadbl(rstimpost!kilos), vnomproveidor, vnifproveidor)
                rstimpost!justificant = vnumfactura
                rstimpost!regimfiscal = "J"
                rstimpost!nomdestinatari = vnomproveidor 'rst!nomclient
                rstimpost!nifdestinatari = vnifproveidor 'rst!nif
                rstimpost!observacions = "DESTRUCCIÓ ABONO " + atrim(rstabonos!numfacturaabono) + " " + UCase(vcolormerma) + " amb compra Intracomunitaria."
                'rstimpost!observacions = "Abono no Espanya amb compra Intracomunitaria."
                rstimpost!lotinplacsa = rst!comanda
                'If DateDiff("m", rstimpost!Data, climitperiode) >= 0 Then
                  If cadbl(rstimpost!justificant) > 0 Then
                        rstimpost!apuntperdeclarar = True
                  End If
               ' End If
                If rstimpost!apuntperdeclarar = False Then vtotsSI = False
                rstimpost!identificador = rstabonos!id
                rstimpost.Update
               End If
               rst.MoveNext
           Wend
           If Not vtotsSI Then dbtmp.Execute "update taula_impost set apuntperdeclarar=false where identificador=" + atrim(rstabonos!id)
           rstabonos.MoveNext
   Wend
fi:
   Set rstc = Nothing
   Set rst = Nothing
   Set rstimpost = Nothing
   dataintracomunitari.Refresh
End Sub

Sub crear_resum_parcials_300_Devolucions(vtipusproveidor As String)
Dim rst As Recordset
  
  Dim vsql As String
  Dim rstimpost As Recordset
  Dim vnumfactura As String
  Dim vimportmerma As Double
  Dim vnomproveidor As String
  Dim vnifproveidor As String
  Dim vcolormerma As String
  Dim vtipusparcial As String
  Dim vdatafactura As String
  Dim vdatainici As Date
  Dim vdatafi As Date
  vdatainici = buscar_data_inicifi("inici", IIf(cadbl(combomes.Tag) = 0, combomes.ListIndex + 1, cadbl(combomes.Tag)), cadbl(IIf(ettrimestre.Visible, ettrimestre.Tag, 0)))
  vdatafi = buscar_data_inicifi("fi", IIf(cadbl(combomes.Tag) = 0, combomes.ListIndex + 1, cadbl(combomes.Tag)), cadbl(IIf(ettrimestre.Visible, ettrimestre.Tag, 0)))
  vtipusparcial = "300"
inici:
  vsql = "SELECT Parcials.id,Parcials.idpalet, Parcials.idbobina, Parcials.metres, Parcials.comanda, Parcials.data, Parcials.operari, Palets.teimpost, materials.colorreciclatge, materials.tanpercentimpostenvasos, ([parcials].[metres]*[bobines].[pesdelproveidor])/[bobines].[mts] AS Kg_recuperar100, ([Kg_recuperar100]*[materials].[tanpercentimpostenvasos])/100 AS Kg_baseimposable, Parcials.numremesa, proveidors.tipusproveidorIMPOST "
  vsql = vsql + " FROM (((Parcials INNER JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN Bobines ON (Parcials.idbobina = Bobines.Idbobina) AND (Parcials.idpalet = Bobines.Idpalet)) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi"
  vsql = vsql + " WHERE (((Parcials.comanda)='" + vtipusparcial + "') AND ((Palets.teimpost)=True) AND (Parcials.numremesa=0 or Parcials.numremesa=null) AND ((proveidors.tipusproveidorIMPOST)='" + vtipusproveidor + "'));"

  Set rst = dbtmp.OpenRecordset(vsql)
  Set rstimpost = dbtmp.OpenRecordset("select * from taula_impost")
  While Not rst.EOF
        vdatafactura = "": vnumfactura = "": vnomproveidor = "": vnifproveidor = ""
        data_factura_devolucio rst, vdatafactura, vnumfactura, vnomproveidor, vnifproveidor
        'If vnumfactura = "" Or vdatafactura = "" Then GoTo proxim
        vcolormerma = rst!colorreciclatge
        'vnumfactura = ""
        rstimpost.AddNew
        rstimpost!concepte = 4
        rstimpost!clauproducte = IIf(vtipusproveidor = "Importació" Or vtipusproveidor = "Intracomunitari", "B", "G")
        rstimpost!Data = IIf(vdatafactura = "", Null, vdatafactura)
        rstimpost!justificant = vnumfactura
        rstimpost!kilos = Redondejar(cadbl(rst!Kg_recuperar100), 3)
        rstimpost!kilosnoreciclats = Redondejar(cadbl(rst!Kg_baseimposable), 3)
        rstimpost!regimfiscal = "J"
        rstimpost!nomdestinatari = vnomproveidor
        rstimpost!nifdestinatari = vnifproveidor
        rstimpost!observacions = "DEVOLUCIÓ " + vtipusparcial + "s " + UCase(vcolormerma) + " amb compra " + vtipusproveidor + "."
        rstimpost!lotinplacsa = rst!idpalet
        If atrim(vdatafactura) <> "" And atrim(vnumfactura) <> "" Then
            If DateValue(vdatafactura) >= DateValue(vdatainici) And DateValue(vdatafactura) <= DateValue(vdatafi) Then
                 rstimpost!apuntperdeclarar = True
            End If
        End If
        rstimpost!identificador = rst!id
        rstimpost.Update
proxim:
        rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
Sub data_factura_devolucio(rst As Recordset, vdatafactura As String, vnumfactura As String, vnomproveidor As String, vnifproveidor As String)
     Dim rst2 As Recordset
     Dim rstp As Recordset
     Set rst2 = dbtmp.OpenRecordset("select * from devoluciomaterial_linies where idparcial300=" + atrim(rst!id))
     If Not rst2.EOF Then
         Set rst2 = dbtmp.OpenRecordset("select * from devoluciomaterial where id=" + atrim(rst2!idcapcalera))
         If Not rst2.EOF Then
             Set rstp = dbtmp.OpenRecordset("SELECT albaransbip.numalbara, albaransbip.nomproveidorcomercial, proveidors_codisSAP.nomproveidor, proveidors_codisSAP.Nif FROM Parcials LEFT JOIN (albaransbip LEFT JOIN proveidors_codisSAP ON albaransbip.codiproveidorcomercial = proveidors_codisSAP.codiSAP) ON Parcials.idpalet = albaransbip.numpalet where parcials.id=" + atrim(rst!id))
             If Not rstp.EOF Then
                    vnumfactura = atrim(rst2!numerofacturasap)
                    vdatafactura = atrim(rst2!datafacturasap)
                    vnomproveidor = atrim(rstp!nomproveidorcomercial)
                    vnifproveidor = atrim(rstp!nif)
              End If
         End If
     End If
     Set rst2 = Nothing
End Sub
Sub crear_resum_parcials_100i300_Intra()
Dim rst As Recordset
  Dim vtipusproveidor As String
  Dim vsql As String
  Dim rstimpost As Recordset
  Dim vnumfactura As String
  Dim vimportmerma As Double
  Dim vnomproveidor As String
  Dim vnifproveidor As String
  Dim vcolormerma As String
  Dim vtipusparcial As String
  
  vtipusparcial = "100"
  vtipusproveidor = "Intracomunitari"
inici:
  vsql = "SELECT Parcials.id,Parcials.idpalet, Parcials.idbobina, Parcials.metres, Parcials.comanda, Parcials.data, Parcials.operari, Palets.teimpost, materials.colorreciclatge, materials.tanpercentimpostenvasos, ([parcials].[metres]*[bobines].[pesdelproveidor])/[bobines].[mts] AS Kg_recuperar100, ([Kg_recuperar100]*[materials].[tanpercentimpostenvasos])/100 AS Kg_baseimposable, Parcials.numremesa, proveidors.tipusproveidorIMPOST "
  vsql = vsql + " FROM (((Parcials INNER JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN Bobines ON (Parcials.idbobina = Bobines.Idbobina) AND (Parcials.idpalet = Bobines.Idpalet)) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi"
  vsql = vsql + " WHERE (((Parcials.comanda)='" + vtipusparcial + "') AND ((Palets.teimpost)=True) AND (Parcials.numremesa=0 or Parcials.numremesa=null) AND ((proveidors.tipusproveidorIMPOST)='" + vtipusproveidor + "'));"

  Set rst = dbtmp.OpenRecordset(vsql)
  Set rstimpost = dbtmp.OpenRecordset("select * from taula_impost")
  While Not rst.EOF
        vcolormerma = rst!colorreciclatge
        vnumfactura = ""
        rstimpost.AddNew
        rstimpost!concepte = 3
        rstimpost!clauproducte = "B"
        vimportmerma = Redondejar(cadbl(rst!Kg_baseimposable), 3)
        rstimpost!Data = dataFACTURA_REC(IIf(IsNull(rst!Data), 0, rst!Data), rst!colorreciclatge, vnumfactura, vimportmerma, vnomproveidor, vnifproveidor)
        rstimpost!justificant = vnumfactura
        If vnumfactura = "" Then rstimpost!Data = Format(rst!Data, "dd/mm/yy")
        rstimpost!kilos = Redondejar(cadbl(rst!Kg_recuperar100), 3)
        rstimpost!kilosnoreciclats = Redondejar(cadbl(rst!Kg_baseimposable), 3)
        rstimpost!regimfiscal = "J"
        rstimpost!nomdestinatari = vnomproveidor
        rstimpost!nifdestinatari = vnifproveidor
        rstimpost!observacions = "MERMA " + vtipusparcial + "s " + UCase(vcolormerma) + " amb compra Intracomunitaria."
        rstimpost!lotinplacsa = rst!idpalet
        If atrim(rstimpost!justificant) <> "" Then
               rstimpost!apuntperdeclarar = True
        End If
        rstimpost!identificador = rst!id
        rstimpost.Update
        rst.MoveNext
  Wend
  If vtipusparcial <> "400" Then vtipusparcial = "400": GoTo inici
  Set rst = Nothing
End Sub
Sub sumar_mermes_i_vendes_ImpIEsp()
   Dim rstimpost As Recordset
   Dim vtotal As Double
   Dim vbaseimp As Double
   Dim vkgvenda1 As Double
   Dim vkgvenda2 As Double
   Dim vkgvendaE1 As Double
   Dim vkgvendaE2 As Double
   Dim vkgvendaK1 As Double
   Dim vkgvendaK2 As Double
   'except de pagar
   Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where concepte=2 and clauproducte='A' and apuntperdeclarar=true and regimfiscal='E'")
   If cadbl(rstimpost!kg) > 0 Then vkgvendaE1 = Redondejar(cadbl(rstimpost!kg), 2)
   Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where concepte=2 and clauproducte='A' and apuntperdeclarar=true and regimfiscal='K'")
   If cadbl(rstimpost!kg) > 0 Then vkgvendaK1 = Redondejar(cadbl(rstimpost!kg), 2)
   Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where concepte=2 and clauproducte='D' and apuntperdeclarar=true and regimfiscal='E'")
   If cadbl(rstimpost!kg) > 0 Then vkgvendaE2 = Redondejar(cadbl(rstimpost!kg), 2)
   Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where concepte=2 and clauproducte='D' and apuntperdeclarar=true and regimfiscal='K'")
   If cadbl(rstimpost!kg) > 0 Then vkgvendaK2 = Redondejar(cadbl(rstimpost!kg), 2)
   'vendes importació
   Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where concepte=2 and apuntperdeclarar=true and clauproducte='A' and regimfiscal='A'")
   etsumakilosmermaivendes = "Importació->Vendes: " + atrim(Redondejar(cadbl(rstimpost!kg), 2)) + " / " + atrim(Redondejar(cadbl(rstimpost!kgrec), 2)) + IIf(vkgvendaE1 > 0, " VendaE= " + atrim(vkgvendaE1) + "Kg ", "") + IIf(vkgvendaK1 > 0, " VendaE= " + atrim(vkgvendaK1) + "Kg ", "")
   vtotal = Redondejar(cadbl(rstimpost!kg), 2)
   vbaseimp = Redondejar(cadbl(rstimpost!kgrec), 2)
   'merma importacio
   Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where concepte=3 and apuntperdeclarar=true and clauproducte='B'")
   etsumakilosmermaivendes = etsumakilosmermaivendes + "    Suma Mermes: " + atrim(Redondejar(cadbl(rstimpost!kg), 2)) + " / " + atrim(Redondejar(cadbl(rstimpost!kgrec), 2))
   vtotal = Redondejar(vtotal + cadbl(rstimpost!kg), 2)
   vbaseimp = Redondejar(vbaseimp + cadbl(rstimpost!kgrec), 2)
   Set rstimpost = Nothing
   
   'vendes espanya
   Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where (concepte=2) and apuntperdeclarar=true and clauproducte='D' and regimfiscal='A'")
   vkgvenda1 = cadbl(rstimpost!kg)
   vkgvenda2 = cadbl(rstimpost!kgrec)
   'etsumakilosmermaivendes = etsumakilosmermaivendes + vbNewLine + "Espanyol--> Suma Vendes: " + atrim(Redondejar(cadbl(rstimpost!kg), 2)) + " / " + atrim(Redondejar(cadbl(rstimpost!kgrec), 2))
   vtotal = vtotal + Redondejar(cadbl(rstimpost!kg), 2)
   vbaseimp = vbaseimp + Redondejar(cadbl(rstimpost!kgrec), 2)
   
   'abonos clients
   Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where (concepte=4) and apuntperdeclarar=true") ' and clauproducte='D'")
   etsumakilosmermaivendes = etsumakilosmermaivendes + vbNewLine + "Espanyol->Vendes: " + atrim(Redondejar(vkgvenda1 - cadbl(rstimpost!kg), 2)) + " / " + atrim(Redondejar(vkgvenda2 - cadbl(rstimpost!kgrec), 2)) + IIf(vkgvendaE2 > 0, " VendaE= " + atrim(vkgvendaE2) + "Kg ", "") + IIf(vkgvendaK2 > 0, " VendaE= " + atrim(vkgvendaK2) + "Kg ", "")
   vtotal = vtotal - Redondejar(cadbl(rstimpost!kg), 2)
   vbaseimp = vbaseimp - Redondejar(cadbl(rstimpost!kgrec), 2)
   'merma espanya
   Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where concepte=3 and apuntperdeclarar=true and clauproducte='G'")
   etsumakilosmermaivendes = etsumakilosmermaivendes + "    Suma Mermes: " + atrim(Redondejar(cadbl(rstimpost!kg), 2)) + " / " + atrim(Redondejar(cadbl(rstimpost!kgrec), 2))
   vtotal = Redondejar(vtotal + cadbl(rstimpost!kg), 2)
   vbaseimp = Redondejar(vbaseimp + cadbl(rstimpost!kgrec), 2)
   Set rstimpost = Nothing
   
   etsumakilosmermaivendes = etsumakilosmermaivendes + vbNewLine + "TOTAL: " + atrim(vtotal) + " / " + atrim(vbaseimp)
   
End Sub
Sub sumar_mermes_i_vendes()
   Dim rstimpost As Recordset
   Dim vtotal As Double
   Dim vbaseimp As Double
   Dim vmermaEsp As Double
   Dim vmermaImp As Double
   etsumakilos = ""
   'except de pagar
   Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where concepte=2 and apuntperdeclarar=true and regimfiscal<>'A'")
   If cadbl(rstimpost!kg) > 0 Then MsgBox "Hi ha vendes de regim fiscal EXCEPT de pagar, E o K", vbCritical, "Atenció"
   'COMPRES INTRACOMUNITARIES NOMES
   If vtipusllistat = "I" Then
        Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where concepte=1 and apuntperdeclarar=true")
        etsumakilos = "Compres IntraCom. Total_Kg: " + atrim(rstimpost!kgrec) + "   Base_Imposable: " + atrim(rstimpost!kg)
   End If
   'VENDES
   Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where concepte=2 and apuntperdeclarar=true")
   etsumakilosmermaivendes = "Suma Vendes: " + atrim(Redondejar(cadbl(rstimpost!kg), 2)) + " / " + atrim(Redondejar(cadbl(rstimpost!kgrec), 2))
   vtotal = Redondejar(cadbl(rstimpost!kg), 2)
   vbaseimp = Redondejar(cadbl(rstimpost!kgrec), 2)
   'ABONOS
   Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where concepte=4 and apuntperdeclarar=true")
   If cadbl(rstimpost!kg) > 0 Then
         etsumakilosmermaivendes = etsumakilosmermaivendes + vbNewLine + "Suma Abonos: " + atrim(Redondejar(cadbl(rstimpost!kg), 2)) + " / " + atrim(Redondejar(cadbl(rstimpost!kgrec), 2))
         vtotal = Redondejar(vtotal - cadbl(rstimpost!kg), 2)
         vbaseimp = Redondejar(vbaseimp - cadbl(rstimpost!kgrec), 2)
   End If
   'MERMES
      'Importació  B i Intracomunitaries
   Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where (concepte=3) and apuntperdeclarar=true and clauproducte='B'")
   vtotal = Redondejar(vtotal + cadbl(rstimpost!kg), 2)
   vbaseimp = Redondejar(vbaseimp + cadbl(rstimpost!kgrec), 2)
   vmermaImp = Redondejar(cadbl(rstimpost!kg), 2)
   If vtipusllistat = "I" Then
           etsumakilosmermaivendes = etsumakilosmermaivendes + vbNewLine + " Mermes Intracomunitaries [B]: " + atrim(Redondejar(cadbl(rstimpost!kg), 2)) + "Kg / " + atrim(Redondejar(cadbl(rstimpost!kgrec), 2)) + "Kg"
       Else: etsumakilosmermaivendes = etsumakilosmermaivendes + vbNewLine + " Mermes Imp[B]: " + atrim(Redondejar(cadbl(rstimpost!kg), 2)) + "Kg / " + atrim(Redondejar(cadbl(rstimpost!kgrec), 2)) + "Kg"
   End If
      'Espanya  G
   If vtipusllistat <> "I" Then
    Set rstimpost = dbtmp.OpenRecordset("select sum(kilos) as kg, sum(kilosnoreciclats) as kgrec from taula_impost where (concepte=3) and apuntperdeclarar=true and clauproducte='G'")
    vtotal = Redondejar(vtotal + cadbl(rstimpost!kg), 2)
    vbaseimp = Redondejar(vbaseimp + cadbl(rstimpost!kgrec), 2)
    vmermaEsp = Redondejar(cadbl(rstimpost!kg), 2)
    etsumakilosmermaivendes = etsumakilosmermaivendes + " Mermes Esp[G]: " + atrim(Redondejar(cadbl(rstimpost!kg), 2)) + "Kg / " + atrim(Redondejar(cadbl(rstimpost!kgrec), 2)) + "Kg"
   End If
   
   Set rstimpost = Nothing
   'TOTALS
   etsumakilosmermaivendes = etsumakilosmermaivendes + vbNewLine + "TOTAL: " + atrim(vtotal) + " / " + atrim(vbaseimp)
   
End Sub
Sub crear_resum_mermes_vendeS_intracomunitaries()
  Dim vsql As String
  Dim rst As Recordset
  Dim rstcolor As Recordset
  Dim rstimpost As Recordset
  Dim vwere As String
  Dim vnumfactura As String
  Dim vimportmerma As Double
  Dim vcolormerma As String
  Dim vproximaseccio As String
  
  vsql = "SELECT ImpostEnvasos.*, materials.colorreciclatge,comandes.proximaseccio FROM (ImpostEnvasos LEFT JOIN comandes ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN materials ON comandes.materialex = materials.codi where Num_remesa_ImpostEnv_merma_Intra=0 or Num_remesa_ImpostEnv_merma_Intra=null order by impostenvasos.comanda;"
  Set rstcolor = dbtmp.OpenRecordset(vsql)
  If rstcolor.EOF Then MsgBox "No hi ha registre d 'IMPOSTOS.": GoTo fi
  
  Set rstimpost = dbtmp.OpenRecordset("select * from taula_impost")
   
  vwere = " KgMermaAd_Intracom>0 and (Num_remesa_ImpostEnv_merma_Intra=0 or Num_remesa_ImpostEnv_merma_Intra=null) and tipusdeentrega='T'"
  
  'vsql = "SELECT ImpostEnvasos.*, capcaleraalbara.*, (select distinct datafactura from Importada_LiniesFacturesSAP_Inplacsa where NumFact=[capcaleraalbara].[numfacturaSAP] ) AS Datafactura, materials.tanpercentimpostenvasos, clients.nom AS nomclient, clients_codisSAP.nif, liniesalbara.tipusdeentrega "
  'vsql = vsql + " FROM ((((ImpostEnvasos LEFT JOIN capcaleraalbara ON ImpostEnvasos.numalbara = capcaleraalbara.numalbara) LEFT JOIN (comandes LEFT JOIN materials ON comandes.materialex = materials.codi) ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN clients_codisSAP ON capcaleraalbara.codiclient = clients_codisSAP.codiSAP) LEFT JOIN liniesalbara ON ImpostEnvasos.id = liniesalbara.id "
  vsql = "SELECT ImpostEnvasos.*, capcaleraalbara.*, (select distinct datafactura from Importada_LiniesFacturesSAP_Inplacsa where NumFact=[capcaleraalbara].[numfacturaSAP] ) AS Datafactura, materials.tanpercentimpostenvasos, clients.nom AS nomclient, clients_codisSAP.nif, liniesalbara.tipusdeentrega "
  vsql = vsql + " FROM ((((ImpostEnvasos LEFT JOIN capcaleraalbara ON ImpostEnvasos.numalbara = capcaleraalbara.numalbara) LEFT JOIN (comandes LEFT JOIN materials ON comandes.materialex = materials.codi) ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN clients_codisSAP ON capcaleraalbara.codiclient = clients_codisSAP.codiSAP) LEFT JOIN liniesalbara ON ImpostEnvasos.idliniaalbara = liniesalbara.id"

'  Clipboard.Clear
'  Clipboard.SetText vsql + " where " + vwere
  Set rst = dbtmp.OpenRecordset(vsql + " where " + vwere)
  While Not rst.EOF
        vnumfactura = ""
        rstcolor.FindFirst "id=" + atrim(rst!id)
        If rstcolor.NoMatch Then MsgBox "No s'ha trobat el color de la merma de la comanda " + atrim(rst!comanda): GoTo cont
        vcolormerma = rstcolor!colorreciclatge
          'sumo tots els colors i resto el color del material de la merma
        vimportcolormerma = (cadbl(rst!kgmermaimpost_ad_capa_verd) + cadbl(rst!kgmermaimpost_ad_capa_blau) + cadbl(rst!kgmermaimpost_ad_capa_vermell)) - cadbl(rst.Fields("KgMermaIMPOST_AD_capa_" + atrim(vcolormerma)))
          'resto l'import del material d'ajust del total de la merma (el del mateix color ja està tret)
        vimportmerma = cadbl(rst!kgmermaad_intracom) - vimportcolormerma
            'LA MERMA D'AJUST SI ES EL MATEIX COLOR QUE LA DEL MATERIAL SURTIRÀ JUNT, SINO ES FARA LINIA NOVA PER CADA COLOR
        
registrenou:
        vnumfactura = ""
        rstimpost.AddNew
        rstimpost!concepte = 3
        rstimpost!clauproducte = "B"
        vproximaseccio = estatcomanda(rst!comanda)
        If vproximaseccio = "T" Then   'If rstcolor!proximaseccio = "T" Then
           rstimpost!Data = dataFACTURA_REC(IIf(IsNull(rst!datafactura), 0, rst!datafactura), vcolormerma, vnumfactura, vimportmerma)
           rstimpost!justificant = vnumfactura
        End If
       ' If vnumfactura = "" Then MsgBox atrim(rst!comanda)
        rstimpost!kilos = Redondejar((cadbl(vimportmerma) * ((rst!tanpercentimpostenvasos / 100) + 1)) - cadbl(vimportmerma), 3)
        rstimpost!kilosnoreciclats = Redondejar(cadbl(vimportmerma), 3)
        rstimpost!regimfiscal = "J"
        rstimpost!nomdestinatari = rst!nomclient
        rstimpost!nifdestinatari = rst!nif
        rstimpost!observacions = "MERMA " + UCase(vcolormerma) + IIf(vcolormerma <> rstcolor!colorreciclatge, " AJUST", "") + " amb compra Intracomunitaria."
        rstimpost!lotinplacsa = rst!comanda
        If vproximaseccio = "T" And atrim(rstimpost!justificant) <> "" Then
                rstimpost!apuntperdeclarar = True
        End If
        rstimpost!identificador = rst!id
        rstimpost.Update
colors:
        If vcolormerma = "Verd" And rstcolor!colorreciclatge <> "Blau" Then vcolormerma = "Blau": vimportmerma = cadbl(rst.Fields("KgMermaIMPOST_AD_capa_" + atrim(vcolormerma))): If vimportmerma > 0 Then GoTo registrenou Else GoTo colors
        If vcolormerma = "Blau" And rstcolor!colorreciclatge <> "Vermell" Then vcolormerma = "Vermell": vimportmerma = cadbl(rst.Fields("KgMermaIMPOST_AD_capa_" + atrim(vcolormerma))): If vimportmerma > 0 Then GoTo registrenou Else GoTo colors
        If vcolormerma = "Vermell" And rstcolor!colorreciclatge <> "Verd" Then vcolormerma = "Verd": vimportmerma = cadbl(rst.Fields("KgMermaIMPOST_AD_capa_" + atrim(vcolormerma))): If vimportmerma > 0 Then GoTo registrenou Else GoTo colors
cont:
        rst.MoveNext
  Wend
   
fi:
   Set rst = Nothing
   Set rstimpost = Nothing
   Set rstcolor = Nothing

  
End Sub
Function estatcomanda(vnumc As Double) As String
   Dim rst As Recordset
   If vnumc = 0 Then Exit Function
    Set rst = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(vnumc))
    If Not rst.EOF Then
       Set rst = dbtmp.OpenRecordset("select proximaseccio from comandes where proximaseccio='T' and (comanda=" + atrim(rst!comanda) + " or comanda=" + atrim(rst!linkcomanda1) + " or comanda=" + atrim(rst!linkcomanda2) + ") order by comanda")
       If Not rst.EOF Then
           estatcomanda = rst!proximaseccio
       End If
    End If
   Set rst = Nothing
End Function
Function buscar_data_inicifi(vtipus As String, vmes As Double, vtrimestre As Double) As Date
   If vtrimestre > 0 Then  'si es trimestre
      If vtipus = "inici" Then
           buscar_data_inicifi = "01/" + atrim((vtrimestre * 3) - 2) + "/" + atrim(IIf(Month(Now) = 1, Year(Now) - 1, Year(Now)))
          Else
             buscar_data_inicifi = atrim(Day(DateSerial(Year(Now), (vtrimestre * 3) + 1, 0))) + "/" + atrim((vtrimestre * 3)) + "/" + atrim(IIf(Month(Now) = 1, Year(Now) - 1, Year(Now)))
      End If
   End If
   If vtrimestre = 0 Then  'si es un mes
      If vtipus = "inici" Then
           buscar_data_inicifi = "01/" + atrim(vmes) + "/" + atrim(IIf(Month(Now) = 1, Year(Now) - 1, Year(Now)))
          Else
             buscar_data_inicifi = atrim(Day(DateSerial(Year(Now), vmes + 1, 0))) + "/" + atrim(vmes) + "/" + atrim(IIf(Month(Now) = 1, Year(Now) - 1, Year(Now)))
      End If
   End If
End Function
Function dataFACTURA_REC(vdata As Date, vcolor As String, vnumfactura As String, vkgdemerma As Double, Optional vnomproveidor As String, Optional vnifproveidor As String, Optional vclauproducte As String) As Date
   Dim rst As Recordset
   Dim vdatainici As Date
   Dim vdatafi As Date
   Dim vFactor As Double
   Static vultimvalor As Double
   vdatainici = buscar_data_inicifi("inici", IIf(cadbl(combomes.Tag) = 0, combomes.ListIndex + 1, cadbl(combomes.Tag)), cadbl(IIf(ettrimestre.Visible, ettrimestre.Tag, 0)))
   vdatafi = buscar_data_inicifi("fi", IIf(cadbl(combomes.Tag) = 0, combomes.ListIndex + 1, cadbl(combomes.Tag)), cadbl(IIf(ettrimestre.Visible, ettrimestre.Tag, 0)))
   If vclauproducte = "B" Then If vultimvalor = 1 Then vultimvalor = 0: GoTo fi
   vultimvalor = 1
   'Set rst = dbtmp.OpenRecordset("select * from facturesSAPreciclatge where tipus like 'DESP" + UCase(vcolor) + "*' and (datafactura>=#" + atrim(Format(vdatainici, "mm/dd/yyyy")) + "# and datafactura<=#" + atrim(Format(vdatafi, "mm/dd/yyyy")) + "#) order by datafactura asc")
   Set rst = dbtmp.OpenRecordset("select * from facturesSAPreciclatge where tipus like 'DESP*' and (datafactura>=#" + atrim(Format(vdatainici, "mm/dd/yyyy")) + "# and datafactura<=#" + atrim(Format(vdatafi, "mm/dd/yyyy")) + "#) order by datafactura asc")
   If Not rst.EOF Then
     'If vtipusllistat = "I" Then vFactor = 0.9
     'If vtipusllistat <> "I" Then
     '    vFactor = 4.5 / 6
     '    If Month(rst!datafactura) < 3 Then vFactor = 1
     'End If
     vFactor = 1
     While (cadbl(rst!kgassignatsdemerma) + cadbl(rst!kgtemporals)) > (cadbl(rst!kgfactura) * vFactor)
       rst.MoveNext
       If rst.EOF Then GoTo fi
       'If vtipusllistat = "I" Then vFactor = 0.92
       'If vtipusllistat <> "I" Then
       '  vFactor = 4.5 / 6
       '  If Month(rst!datafactura) < 3 Then vFactor = 1
       'End If
     Wend
     vnumfactura = rst!numerofactura
    ' If vnumfactura = 6453698 Then Stop
     dataFACTURA_REC = rst!datafactura
     vnomproveidor = atrim(rst!nomproveidor)
     vnifproveidor = atrim(rst!nifproveidor)
     rst.Edit
     rst!kgtemporals = rst!kgtemporals + vkgdemerma
     rst.Update
       Else: dataFACTURA_REC = vdata: vnumfactura = ""
   End If
fi:
   Set rst = Nothing
End Function
Private Sub Command2_Click()
  'posso a zero el comptador de kgassignats a cada factura TEMPORAL PER AQUEST CALCUL
  dbtmp.Execute "update facturesSAPreciclatge set kgtemporals=0"
  ettrimestre.Visible = True
  bguardarCSV.Enabled = False
  vsumakilos = 0: vsumakilos2 = 0
  etllistat = "Llistat A22 de vendes i mermes IMPORTACIÓ I ESPANYA"
  etllistat.Tag = "E+I"
  vtipusllistat = "E+I"
  crear_resum_ventes_Esp_i_Imp
 ' crear_resum_abonos_Esp_i_Imp
 ' crear_resum_abonos_RECILAR_ImpIEsp
  If checknocomptarmermes <> 1 Then
    crear_resum_parcials_100i300_Imp_Esp "Espanyol"
    crear_resum_parcials_100i300_Imp_Esp "Importació"
    crear_resum_mermes_vendeS_Esp_i_Imp
   ' crear_resum_parcials_300_Devolucions "Importació"
' NO   crear_resum_parcials_300_Devolucions "Espanyol"  ' nocal mirarlo perquè ens ho abona el proveidor
  End If
  etsumakilos = ""
  'sumar_mermes_i_vendes
  sumar_mermes_i_vendes_ImpIEsp
  dataintracomunitari.Refresh
  configurar_reixa
  emplenar_reixa
  bfacturescompres.Enabled = True
End Sub
Sub crear_resum_abonos_RECILAR_ImpIEsp()
   Dim rst As Recordset
   Dim rstabonos As Recordset
   Dim rstimpost As Recordset
   Dim rstfactura As Recordset
   Dim rstc As Recordset
   Dim vwere As String
   Dim vsql As String
   Dim vkgcapa As Double
   Dim vnumfactura As String
   Dim vcolormerma As String
   Dim vtotsSI As Boolean
   Dim vnomproveidor As String
   Dim vnifproveidor As String
   Dim vsumaEspImp As Double
   Dim vtotalimpost As Double
   
   Set rstabonos = dbtmp.OpenRecordset("select * from abonosclients where numremesadestruccio_A22=0 or numremesadestruccio_A22=null AND TOTAIMPOST>0")
   Set rstimpost = dbtmp.OpenRecordset("select * from taula_impost")
   
   While Not rstabonos.EOF
           'vwere = " (PaisVenta<>'ES' or regimfiscal<>'') and KgVentaAd_Intracom>0 and (((ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Intra) Is Null Or (ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Intra)=0)) "
           Set rstc = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(rstabonos!lotinplacsa))
           vwere = " (impostenvasos.comanda=" + atrim(rstc!comanda) + " or impostenvasos.comanda=" + atrim(rstc!linkcomanda1) + " or impostenvasos.comanda=" + atrim(rstc!linkcomanda2) + ") "
           vsql = "SELECT ImpostEnvasos.*, capcaleraalbara.*, (select distinct datafactura from Importada_LiniesFacturesSAP_Inplacsa where NumFact=[capcaleraalbara].[numfacturaSAP] ) AS Datafactura, materials.tanpercentimpostenvasos, clients.nom AS nomclient, clients_codisSAP.nif "
           vsql = vsql + " FROM (((ImpostEnvasos LEFT JOIN capcaleraalbara ON ImpostEnvasos.numalbara = capcaleraalbara.numalbara) LEFT JOIN (comandes LEFT JOIN materials ON comandes.materialex = materials.codi) ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN clients_codisSAP ON capcaleraalbara.codiclient = clients_codisSAP.codiSAP "
           Set rst = dbtmp.OpenRecordset(vsql + " where " + vwere)
           vtotsSI = True
           While Not rst.EOF
                vsumaEspImp = cadbl(rst!kgventaespanya) + cadbl(rst!kgventaimp_mes_esp)
               If cadbl(vsumaEspImp) > 0 Then
                'If rstabonos!lotinplacsa = 212407 Then Stop
                vkgcapa = (totalKgtoteslescapes(rst!comanda, "A22") * 100) / (totalKgtoteslescapes(rst!comanda, "A22") + totalKgtoteslescapes(rst!comanda, "592")) 'saber el % TOTAL
                vtotalimpost = (rstabonos!totaimpost) * (vkgcapa / 100)
                vkgcapa = (vsumaEspImp * 100) / totalKgtoteslescapes(rst!comanda, "A22") 'saber el % DE CAPA
                'If rstabonos!lotinplacsa = 212407 Then Stop
                vkgcapa = vtotalimpost * (vkgcapa / 100)
                rstimpost.AddNew
                rstimpost!concepte = 3
                rstimpost!clauproducte = IIf(vtipusproveidor = "Importació", "B", "G")
                vcolormerma = rstabonos!colorreciclat
                rstimpost!kilos = Redondejar((vkgcapa * ((rst!tanpercentimpostenvasos / 100) + 1)) - vkgcapa, 3)
                rstimpost!kilosnoreciclats = Redondejar(vkgcapa, 3) 'Redondejar(rstabonos!totaimpost, 3)
                rstimpost!Data = dataFACTURA_REC(IIf(IsNull(rst!datafactura), 0, rst!datafactura), vcolormerma, vnumfactura, cadbl(rstimpost!kilos), vnomproveidor, vnifproveidor)
                rstimpost!justificant = vnumfactura
                rstimpost!regimfiscal = "J"
                rstimpost!nomdestinatari = vnomproveidor 'rst!nomclient
                rstimpost!nifdestinatari = vnifproveidor 'rst!nif
                rstimpost!observacions = "DESTRUCCIÓ ABONO " + atrim(rstabonos!numfacturaabono) + " " + UCase(vcolormerma) + " amb compra a Espany o Importació."
                'rstimpost!observacions = "Abono no Espanya amb compra Intracomunitaria."
                rstimpost!lotinplacsa = rst!comanda
               ' If DateDiff("m", rstimpost!Data, climitperiode) >= 0 Then
                  If cadbl(rstimpost!justificant) > 0 Then
                        rstimpost!apuntperdeclarar = True
                  End If
                'End If
                If rstimpost!apuntperdeclarar = False Then vtotsSI = False
                rstimpost!identificador = rstabonos!id
                rstimpost.Update
               End If
               rst.MoveNext
           Wend
           If Not vtotsSI Then dbtmp.Execute "update taula_impost set apuntperdeclarar=false where concepte=3 and identificador=" + atrim(rstabonos!id)
           rstabonos.MoveNext
   Wend
fi:
   Set rstc = Nothing
   Set rst = Nothing
   Set rstimpost = Nothing
   dataintracomunitari.Refresh
End Sub
Sub crear_resum_parcials_100i300_Imp_Esp(vtipusproveidor As String)
  Dim rst As Recordset

  Dim vsql As String
  Dim rstimpost As Recordset
  Dim vnumfactura As String
  Dim vimportmerma As Double
  Dim vnomproveidor As String
  Dim vnifproveidor As String
  Dim vcolormerma As String
  
  vtipusparcial = "100"
 'vtipusproveidor = "Intracomunitari"
inici:

  
  vsql = "SELECT Parcials.id,Parcials.idpalet, Parcials.idbobina, Parcials.metres, Parcials.comanda, Parcials.data, Parcials.operari, Palets.teimpost, materials.colorreciclatge, materials.tanpercentimpostenvasos, ([parcials].[metres]*[bobines].[pesdelproveidor])/[bobines].[mts] AS Kg_recuperar100, ([Kg_recuperar100]*[materials].[tanpercentimpostenvasos])/100 AS Kg_baseimposable, Parcials.numremesa, proveidors.tipusproveidorIMPOST "
  vsql = vsql + " FROM (((Parcials INNER JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN Bobines ON (Parcials.idbobina = Bobines.Idbobina) AND (Parcials.idpalet = Bobines.Idpalet)) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi"
  vsql = vsql + " WHERE (((Parcials.comanda)='" + vtipusparcial + "') AND ((Palets.teimpost)=True) AND ((Parcials.numremesa)=0) AND ((proveidors.tipusproveidorIMPOST)='" + vtipusproveidor + "'));"

  Set rst = dbtmp.OpenRecordset(vsql)
  Set rstimpost = dbtmp.OpenRecordset("select * from taula_impost")
  While Not rst.EOF
        vcolormerma = rst!colorreciclatge
        vnumfactura = ""
        vnomproveidor = "": vnifproveidor = ""
        rstimpost.AddNew
        rstimpost!concepte = 3
        rstimpost!clauproducte = IIf(vtipusproveidor = "Importació", "B", "G")
        vimportmerma = Redondejar(cadbl(rst!Kg_baseimposable), 3)
        rstimpost!Data = dataFACTURA_REC(IIf(IsNull(rst!Data), 0, rst!Data), rst!colorreciclatge, vnumfactura, vimportmerma, vnomproveidor, vnifproveidor)
        rstimpost!justificant = vnumfactura
        rstimpost!kilos = Redondejar(cadbl(rst!Kg_recuperar100), 3)
        rstimpost!kilosnoreciclats = Redondejar(cadbl(rst!Kg_baseimposable), 3)
        rstimpost!regimfiscal = "J"
        rstimpost!nomdestinatari = vnomproveidor
        rstimpost!nifdestinatari = vnifproveidor
        rstimpost!observacions = "MERMA " + vtipusparcial + "s " + UCase(vcolormerma) + " amb compra " + vtipusproveidor + "."
        rstimpost!lotinplacsa = rst!idpalet
        If atrim(rstimpost!justificant) <> "" Then
           ' If Format(rstimpost!Data, "q") <= cadbl(ettrimestre.Tag) Then
                 rstimpost!apuntperdeclarar = True
            'End If
               Else: rstimpost!Data = rst!Data
        End If
        rstimpost!identificador = rst!id
        rstimpost.Update
        rst.MoveNext
  Wend
  If vtipusparcial <> "400" Then vtipusparcial = "400": GoTo inici
  Set rst = Nothing
End Sub
Sub crear_resum_mermes_vendeS_Esp_i_Imp()
  Dim vsql As String
  Dim rst As Recordset
  Dim rstcolor As Recordset
  Dim rstimpost As Recordset
  Dim vwere As String
  Dim vnumfactura As String
  Dim vimportmerma As Double
  Dim vcolormerma As String
  Dim vclauproducte As String
  Dim vproximaseccio As String

 ' vclauproducte = "G"
 ' vsql = "SELECT ImpostEnvasos.*, materials.colorreciclatge,comandes.proximaseccio FROM (ImpostEnvasos LEFT JOIN comandes ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN materials ON comandes.materialex = materials.codi where (Num_remesa_ImpostEnv_merma_Esp=0 or Num_remesa_ImpostEnv_merma_Esp=null) and kgmermaEspanya>0 order by impostenvasos.comanda;"
  vclauproducte = "B"
  vsql = "SELECT ImpostEnvasos.*, materials.colorreciclatge,comandes.proximaseccio FROM (ImpostEnvasos LEFT JOIN comandes ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN materials ON comandes.materialex = materials.codi where (Num_remesa_ImpostEnv_merma_Imp=0 or Num_remesa_ImpostEnv_merma_Imp=null) and KgMermaImp_mes_Esp>0 order by impostenvasos.comanda;"
inici:
  Set rstcolor = dbtmp.OpenRecordset(vsql)
  If rstcolor.EOF Then MsgBox "No hi ha registre d 'IMPOSTOS.": GoTo fi
  
  Set rstimpost = dbtmp.OpenRecordset("select * from taula_impost")
   
  If vclauproducte = "B" Then vwere = " KgMermaImp_mes_Esp>0 and (Num_remesa_ImpostEnv_merma_Imp=0 or Num_remesa_ImpostEnv_merma_Imp=null) and tipusdeentrega='T'"
  If vclauproducte = "G" Then vwere = " KgMermaespanya>0 and (Num_remesa_ImpostEnv_merma_esp=0 or Num_remesa_ImpostEnv_merma_esp=null) and tipusdeentrega='T'"
  'vsql = "SELECT ImpostEnvasos.*, capcaleraalbara.*, (select distinct datafactura from Importada_LiniesFacturesSAP_Inplacsa where NumFact=[capcaleraalbara].[numfacturaSAP] ) AS Datafactura, materials.tanpercentimpostenvasos, clients.nom AS nomclient, clients_codisSAP.nif "
  'vsql = vsql + " FROM (((ImpostEnvasos LEFT JOIN capcaleraalbara ON ImpostEnvasos.numalbara = capcaleraalbara.numalbara) LEFT JOIN (comandes LEFT JOIN materials ON comandes.materialex = materials.codi) ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN clients_codisSAP ON capcaleraalbara.codiclient = clients_codisSAP.codiSAP "
  vsql = "SELECT ImpostEnvasos.*, capcaleraalbara.*, (select distinct datafactura from Importada_LiniesFacturesSAP_Inplacsa where NumFact=[capcaleraalbara].[numfacturaSAP] ) AS Datafactura, materials.tanpercentimpostenvasos, clients.nom AS nomclient, clients_codisSAP.nif, liniesalbara.tipusdeentrega "
  vsql = vsql + " FROM ((((ImpostEnvasos LEFT JOIN capcaleraalbara ON ImpostEnvasos.numalbara = capcaleraalbara.numalbara) LEFT JOIN (comandes LEFT JOIN materials ON comandes.materialex = materials.codi) ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN clients_codisSAP ON capcaleraalbara.codiclient = clients_codisSAP.codiSAP) LEFT JOIN liniesalbara ON ImpostEnvasos.idliniaalbara = liniesalbara.id"
  
 ' Clipboard.Clear
 ' Clipboard.SetText vsql + " where " + vwere
  
  Set rst = dbtmp.OpenRecordset(vsql + " where " + vwere)
  While Not rst.EOF
        vnumfactura = ""
        rstcolor.FindFirst "id=" + atrim(rst!id)
        If rstcolor.NoMatch Then MsgBox "No s'ha trobat el color de la merma de la comanda " + atrim(rst!comanda): GoTo cont
        If vclauproducte = "B" Then
                vcolormerma = atrim(rstcolor!colorreciclatge)
                If vcolormerma = "" Then vcolormerma = "VERD"
                'sumo tots els colors i resto el color del material de la merma
                vimportcolormerma = (cadbl(rst!kgmermaimpost_ie_capa_verd) + cadbl(rst!kgmermaimpost_ie_capa_blau) + cadbl(rst!kgmermaimpost_ie_capa_vermell)) - cadbl(rst.Fields("KgMermaIMPOST_IE_capa_" + atrim(vcolormerma)))
                'resto l'import del material d'ajust del total de la merma (el del mateix color ja està tret)
                vimportmerma = cadbl(rst!kgMERMAimp_mes_esp) - vimportcolormerma
                'LA MERMA D'AJUST SI ES EL MATEIX COLOR QUE LA DEL MATERIAL SURTIRÀ JUNT, SINO ES FARA LINIA NOVA PER CADA COLOR
        End If
        If vclauproducte = "G" Then
                vcolormerma = atrim(rstcolor!colorreciclatge)
                If vcolormerma = "" Then vcolormerma = "VERD"
                'sumo tots els colors i resto el color del material de la merma
                vimportcolormerma = (cadbl(rst!kgmermaimpost_es_capa_verd) + cadbl(rst!kgmermaimpost_es_capa_blau) + cadbl(rst!kgmermaimpost_es_capa_vermell)) - cadbl(rst.Fields("KgMermaIMPOST_ES_capa_" + atrim(vcolormerma)))
                'resto l'import del material d'ajust del total de la merma (el del mateix color ja està tret)
                vimportmerma = cadbl(rst!kgMERMAespanya) - vimportcolormerma
                'LA MERMA D'AJUST SI ES EL MATEIX COLOR QUE LA DEL MATERIAL SURTIRÀ JUNT, SINO ES FARA LINIA NOVA PER CADA COLOR
        End If
        
registrenou:
        rstimpost.AddNew
        rstimpost!concepte = 3
        rstimpost!clauproducte = vclauproducte
        vproximaseccio = estatcomanda(rst!comanda)
        If vproximaseccio = "T" Then
            rstimpost!Data = dataFACTURA_REC(IIf(IsNull(rst!datafactura), 0, rst!datafactura), vcolormerma, vnumfactura, vimportmerma, , , vclauproducte)
            rstimpost!justificant = vnumfactura
           
        End If
        rstimpost!kilos = Redondejar((cadbl(vimportmerma) * ((rst!tanpercentimpostenvasos / 100) + 1)) - cadbl(vimportmerma), 3)
        rstimpost!kilosnoreciclats = Redondejar(cadbl(vimportmerma), 3)
        rstimpost!regimfiscal = "J"
        rstimpost!nomdestinatari = rst!nomclient
        rstimpost!nifdestinatari = rst!nif
        rstimpost!observacions = "MERMA " + UCase(vcolormerma) + IIf(vcolormerma <> rstcolor!colorreciclatge, " AJUST", "") + " amb compra " + atrim(IIf(vclauproducte = "B", "a Importació", "a Espanya"))
        rstimpost!lotinplacsa = rst!comanda
        If vproximaseccio = "T" And atrim(rstimpost!justificant) <> "" Then
             If Format(rstimpost!Data, "q") <= cadbl(ettrimestre.Tag) Then
                rstimpost!apuntperdeclarar = True
             End If
        End If
        rstimpost!identificador = rst!id
        rstimpost.Update
colors:
        If vcolormerma = "Verd" And rstcolor!colorreciclatge <> "Blau" Then vcolormerma = "Blau": vimportmerma = cadbl(rst.Fields("KgMermaIMPOST_AD_capa_" + atrim(vcolormerma))): If vimportmerma > 0 Then GoTo registrenou Else GoTo colors
        If vcolormerma = "Blau" And rstcolor!colorreciclatge <> "Vermell" Then vcolormerma = "Vermell": vimportmerma = cadbl(rst.Fields("KgMermaIMPOST_AD_capa_" + atrim(vcolormerma))): If vimportmerma > 0 Then GoTo registrenou Else GoTo colors
        If vcolormerma = "Vermell" And rstcolor!colorreciclatge <> "Verd" Then vcolormerma = "Verd": vimportmerma = cadbl(rst.Fields("KgMermaIMPOST_AD_capa_" + atrim(vcolormerma))): If vimportmerma > 0 Then GoTo registrenou Else GoTo colors
        
cont:
        rst.MoveNext
  Wend
  If vclauproducte = "B" Then
     ' vclauproducte = "B"
     ' vsql = "SELECT ImpostEnvasos.*, materials.colorreciclatge,comandes.proximaseccio FROM (ImpostEnvasos LEFT JOIN comandes ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN materials ON comandes.materialex = materials.codi where (Num_remesa_ImpostEnv_merma_Imp=0 or Num_remesa_ImpostEnv_merma_Imp=null) and KgMermaImp_mes_Esp>0 order by impostenvasos.comanda;"
      vclauproducte = "G"
      vsql = "SELECT ImpostEnvasos.*, materials.colorreciclatge,comandes.proximaseccio FROM (ImpostEnvasos LEFT JOIN comandes ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN materials ON comandes.materialex = materials.codi where (Num_remesa_ImpostEnv_merma_Esp=0 or Num_remesa_ImpostEnv_merma_Esp=null) and kgmermaEspanya>0 order by impostenvasos.comanda;"
      GoTo inici
  End If
fi:
   Set rst = Nothing
   Set rstimpost = Nothing
   Set rstcolor = Nothing

  

End Sub
Sub crear_resum_ventes_Esp_i_Imp()
   Dim rst As Recordset
   Dim rstimpost As Recordset
   Dim rstfactura As Recordset
   Dim vwere As String
   Dim vsql As String
   Dim vtipusventa As String
   Dim vkgimpost As Double
   Dim vdatainici As Date
   Dim vdatafi As Date
   vdatainici = buscar_data_inicifi("inici", IIf(cadbl(combomes.Tag) = 0, combomes.ListIndex + 1, cadbl(combomes.Tag)), cadbl(IIf(ettrimestre.Visible, ettrimestre.Tag, 0)))
   vdatafi = buscar_data_inicifi("fi", IIf(cadbl(combomes.Tag) = 0, combomes.ListIndex + 1, cadbl(combomes.Tag)), cadbl(IIf(ettrimestre.Visible, ettrimestre.Tag, 0)))
   
   dbtmp.Execute "delete * from taula_impost"
   Set rstimpost = dbtmp.OpenRecordset("select * from taula_impost")
   If Not rstimpost.EOF Then MsgBox "No s'han eliminat els registres temporals.": GoTo fi
   dataintracomunitari.Refresh
   DoEvents
      
   Set rstimpost = dbtmp.OpenRecordset("select * from taula_impost")
   
   
   vwere = " (PaisVenta<>'ES' or regimfiscal<>'') and KgVentaImp_mes_Esp>0 and (((ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Imp) Is Null Or (ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Imp)=0)) "
   vtipusventa = "Importació"
   
inici:
   vsql = "SELECT ImpostEnvasos.*, capcaleraalbara.*, (select distinct datafactura from Importada_LiniesFacturesSAP_Inplacsa where NumFact=[capcaleraalbara].[numfacturaSAP] ) AS Datafactura, materials.tanpercentimpostenvasos, clients.nom AS nomclient, clients_codisSAP.nif "
   vsql = vsql + " FROM (((ImpostEnvasos LEFT JOIN capcaleraalbara ON ImpostEnvasos.numalbara = capcaleraalbara.numalbara) LEFT JOIN (comandes LEFT JOIN materials ON comandes.materialex = materials.codi) ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN clients_codisSAP ON capcaleraalbara.codiclient = clients_codisSAP.codiSAP "
   Set rst = dbtmp.OpenRecordset(vsql + " where " + vwere)
   While Not rst.EOF
       
        rstimpost.AddNew
        rstimpost!concepte = 2
        rstimpost!clauproducte = IIf(vtipusventa = "Importació", "A", "D")
        rstimpost!Data = rst!datafactura
        rstimpost!justificant = rst!numfacturasap
        vkgimpost = IIf(vtipusventa = "Importació", rst!kgventaimp_mes_esp, rst!kgventaespanya)
        rstimpost!kilos = Redondejar((vkgimpost * ((rst!tanpercentimpostenvasos / 100) + 1)) - vkgimpost, 2)
        rstimpost!kilosnoreciclats = Redondejar(vkgimpost, 2)
        rstimpost!regimfiscal = IIf(atrim(rst!regimfiscal) = "", "A", rst!regimfiscal)
        rstimpost!nomdestinatari = rst!nomclient
        rstimpost!nifdestinatari = rst!nif
        rstimpost!observacions = "Venta fora d'Espanya amb compra per " + vtipusventa + "."
        If rstimpost!regimfiscal <> "A" Then rstimpost!observacions = "Venta amb règim fiscal exempt (Lletra " + rstimpost!regimfiscal + ")"
        rstimpost!lotinplacsa = rst!comanda
        If Not IsNull(rst!datafactura) Then
       '     If Format(rst!datafactura, "q") <= cadbl(ettrimestre.Tag) Then
             If DateValue(rst!datafactura) >= DateValue(vdatainici) And DateValue(rst!datafactura) <= DateValue(vdatafi) Then
                If cadbl(rstimpost!justificant) > 0 Then
                    rstimpost!apuntperdeclarar = True
                End If
             End If
        '    End If
        End If
        rstimpost!identificador = rst!id
        rstimpost.Update
       
       rst.MoveNext
   Wend
   If vtipusventa = "Importació" Then
        vwere = " (PaisVenta<>'ES' or regimfiscal<>'') and KgVentaespanya>0 and (((ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Esp) Is Null Or (ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Esp)=0)) "
        vtipusventa = "Adquirent"
        GoTo inici
   End If
   
fi:
   Set rst = Nothing
   Set rstimpost = Nothing
   dataintracomunitari.Refresh
End Sub





Private Sub etfacmermaverd_Click()

End Sub

Private Sub etfacmermavermella_Click()

End Sub

Function treuresimbols(desc As String) As String
   desc = substituir(desc, ":", "_")
   desc = substituir(desc, "'", "´")
   desc = substituir(desc, "|", "_")
   desc = substituir(desc, ";", "_")
   treuresimbols = desc
End Function
Function substituir(cadena As String, buscar As String, canviar As String, Optional vcanvis As Long) As String
   If atrim(buscar) = atrim(canviar) Then GoTo fi
   cadena = "  " + cadena
   While InStr(1, cadena, buscar) > 0
    comença = InStr(1, cadena, buscar) - 1
    If comença < 1 Then substituir = cadena: Exit Function
    acaba = comença + Len(buscar) + 1
    cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
    vcanvis = vcanvis + 1
   Wend
fi:
   substituir = atrim(cadena)
   'MsgBox linia
End Function
Private Sub Command3_Click()
   Dim rst As Recordset
   Dim vnomfitxerCSV As String
   vnomfitxerCSV = "c:\temp\" + treuresimbols(etllistat) + ".csv" 'Llistat_Resum_ImpostEnv_Intracomunitari.csv"
   
   Set rst = dataintracomunitari.Recordset
   rst.MoveFirst
   Open vnomfitxerCSV For Output As #1
   'Print #1, "Assentament;Apunt per declarar;Lot_Inplacsa;Concepte;Clau_Producte;Data;Justificant;Kilos;Kilos_NoReciclat;RegimFiscal;Nom_Destinatari;Nif_destinatari;Observacions"
   For i = 0 To rst.Fields.Count - 1
      linia = linia + IIf(linia = "", "", ";") + """" + atrim(UCase(rst.Fields(i).Name)) + """"
    Next i
    Print #1, linia
   While Not rst.EOF
    linia = ""
    For i = 0 To rst.Fields.Count - 1
      linia = linia + IIf(linia = "", "", ";") + """" + atrim(rst.Fields(i)) + """"
    Next i
    Print #1, linia
    rst.MoveNext
   Wend
   Close #1
   wait 2
   obrir_document vnomfitxerCSV
   Set rst = Nothing
End Sub
Function comprovar_quelaremesasiguiunica(vNumRemesa As Double) As Boolean
  Dim rst As Recordset
  comprovar_quelaremesasiguiunica = True
  Set rst = dbtmp.OpenRecordset("select * from Remeses_Taula_Impost_Intracomunitaria where numremesa=" + atrim(vNumRemesa))
  If Not rst.EOF Then comprovar_quelaremesasiguiunica = False
  Set rst = dbtmp.OpenRecordset("select * from Remeses_Taula_Impost_ImpIEsp where numremesa=" + atrim(vNumRemesa))
  If Not rst.EOF Then comprovar_quelaremesasiguiunica = False
  Set rst = Nothing
End Function
Private Sub Command4_Click()
  Dim v As String
  Dim vNumRemesa As String
  Dim vassentament As Double
  
  If etllistat = "" Then MsgBox "No puc registrar les dades si primer no esculls quines dades son.": Exit Sub
  If etllistat.Tag = "" Then MsgBox "No hi ha dades generades.", vbCritical, "Error": Exit Sub
  
  If UCase(InputBoxEx("Escriu la contrasenya per poder generar els fitxers.", "Contrasenya", , , , , , SPassword)) <> "INPLACSA" Then MsgBox "La contrasenya no es correcte": Exit Sub
  vassentament = cadbl(llegir_ini("Impost Envasos", "UltimAssentament", rutadelfitxer(cami) + "valorsprograma.ini"))
  ' SI etllistat.Tag = "I" ES INTRACOMUNITARIES
  vNumRemesa = IIf(etllistat.Tag = "I", "1", "2") + Format(DateAdd("m", -1, Now), "yyyymm")
  v = InputBox(etllistat + vbNewLine + "Aquest procés passarà tots els registres implicats a entregats amb numero de REMESA " + vNumRemesa + vbNewLine + " Escriu [CORRECTE] per acceptar-ho.")
  If UCase(v) <> "CORRECTE" Then Exit Sub
  If MsgBox("Es gravarà la remesa Nº: " + vNumRemesa + vbNewLine + "ESTÀS SEGUR QUE ES LA CORRECTE?", vbExclamation + vbDefaultButton2 + vbYesNo, "COMPROVACIÓ") = vbNo Then Exit Sub
  If comprovar_quelaremesasiguiunica(cadbl(vNumRemesa)) = False Then MsgBox "Aquesta remesa ja s'ha utilitzat,ERROR", vbCritical, "ERROR": Exit Sub
  dbtmp.Execute "update taula_impost set numremesa=" + atrim(vNumRemesa)
  'nomes posso numero d'assentament si es intracomunitaria per presentar el CSV
  If etllistat.Tag = "I" Then possar_assentaments vassentament
  wait 1
  
  If etllistat.Tag = "I" Then
     guardar_historics_remesa "Remeses_Taula_Impost_Intracomunitaria"
  End If
  If etllistat.Tag = "E+I" Then
     guardar_historics_remesa "Remeses_Taula_Impost_ImpIEsp"
  End If
  pasarnumremesa_a_totselsregistresimplicats cadbl(vNumRemesa)
  consolidar_totals_mermes
  
  escriure_ini "Impost Envasos", "UltimAssentament", Trim(vassentament), rutadelfitxer(cami) + "valorsprograma.ini"
  MsgBox "PROCÉS ACABAT"
End Sub
Sub consolidar_totals_mermes()
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from facturesSAPreciclatge where kgtemporals>0")
  While Not rst.EOF
    rst.Edit
    rst!kgassignatsdemerma = cadbl(rst!kgassignatsdemerma) + cadbl(rst!kgtemporals)
    rst!kgtemporals = 0
    rst.Update
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
Sub possar_assentaments(vassentament As Double)
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from taula_impost WHERE apuntperdeclarar=true order by concepte,data")
  While Not rst.EOF
    vassentament = vassentament + 1
    rst.Edit
    rst!numassentament = vassentament
    rst.Update
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
Sub pasarnumremesa_a_totselsregistresimplicats(vNumRemesa As Long)
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from taula_impost WHERE apuntperdeclarar=true")
  If etllistat.Tag = "I" Then
    While Not rst.EOF
        If rst!concepte = 1 Then registro_compres_intracomunitaries rst, vNumRemesa
        If rst!concepte = 2 Then registro_vendes_intracomunitaries rst, vNumRemesa
        If rst!concepte = 3 Then
            registro_mermes_intracomunitaries rst, vNumRemesa
        End If
        If rst!concepte = 4 Then registro_abonos_intracomunitaries rst, vNumRemesa
        rst.MoveNext
    Wend
  End If
  If etllistat.Tag = "E+I" Then
    While Not rst.EOF
        If rst!concepte = 2 Then registro_vendes_EspoImp rst, vNumRemesa
        If rst!concepte = 3 Then
            registro_mermes_EspoImp rst, vNumRemesa
        End If
        If rst!concepte = 4 Then registro_abonos_EspoImp rst, vNumRemesa
        rst.MoveNext
    Wend
  End If
    
  Set rst = Nothing
End Sub
Sub registro_vendes_EspoImp(rst As Recordset, vNumRemesa As Long)
    If rst!clauproducte = "A" Then dbtmp.Execute "update ImpostEnvasos set Num_remesa_ImpostEnv_Venta_Imp=" + atrim(vNumRemesa) + " where id=" + atrim(rst!identificador)
    If rst!clauproducte = "D" Then dbtmp.Execute "update ImpostEnvasos set Num_remesa_ImpostEnv_Venta_Esp=" + atrim(vNumRemesa) + " where id=" + atrim(rst!identificador)
End Sub
Sub registro_mermes_EspoImp(rst As Recordset, vNumRemesa As Long)
    If InStr(1, rst!observacions, "MERMA 400s") = 0 And InStr(1, rst!observacions, "MERMA 100s") = 0 And InStr(1, rst!observacions, "DEVOLUCIÓ 300s") = 0 Then
       If rst!clauproducte = "B" Then dbtmp.Execute "update ImpostEnvasos set Num_remesa_ImpostEnv_Merma_Imp=" + atrim(vNumRemesa) + " where id=" + atrim(rst!identificador)
       If rst!clauproducte = "G" Then dbtmp.Execute "update ImpostEnvasos set Num_remesa_ImpostEnv_Merma_Esp=" + atrim(vNumRemesa) + " where id=" + atrim(rst!identificador)
    End If
    If InStr(1, rst!observacions, "MERMA 400s") > 0 Or InStr(1, rst!observacions, "MERMA 100s") > 0 Or InStr(1, rst!observacions, "DEVOLUCIÓ 300s") > 0 Then
       dbtmp.Execute "update parcials set numremesa=" + atrim(vNumRemesa) + " where id=" + atrim(rst!identificador)
    End If
    If InStr(1, rst!observacions, "DESTRUCCIÓ ABONO") > 0 Then
       dbtmp.Execute "update abonosclients set numremesadestruccio_a22=" + atrim(vNumRemesa) + " where id=" + atrim(rst!identificador)
    End If
End Sub

Sub registro_mermes_intracomunitaries(rst As Recordset, vNumRemesa As Long)
  If InStr(1, rst!observacions, "MERMA 400s") = 0 And InStr(1, rst!observacions, "MERMA 100s") = 0 And InStr(1, rst!observacions, "MERMA 300s") = 0 And InStr(1, rst!observacions, "DEVOLUCIÓ ") = 0 Then dbtmp.Execute "update ImpostEnvasos set Num_remesa_ImpostEnv_Merma_Intra=" + atrim(vNumRemesa) + " where id=" + atrim(rst!identificador)
  If InStr(1, rst!observacions, "MERMA 400s") > 0 Or InStr(1, rst!observacions, "MERMA 100s") > 0 Or InStr(1, rst!observacions, "DEVOLUCIÓ 300s") > 0 Then
      dbtmp.Execute "update parcials set numremesa=" + atrim(vNumRemesa) + " where id=" + atrim(rst!identificador)
  End If
  If InStr(1, rst!observacions, "DESTRUCCIÓ ABONO") > 0 Then
       dbtmp.Execute "update abonosclients set numremesadestruccio_592=" + atrim(vNumRemesa) + " where id=" + atrim(rst!identificador)
    End If
    
    
  'Num_remesa_ImpostEnv_Merma
End Sub
Sub registro_vendes_intracomunitaries(rst As Recordset, vNumRemesa As Long)
   dbtmp.Execute "update ImpostEnvasos set Num_remesa_ImpostEnv_Venta_Intra=" + atrim(vNumRemesa) + " where id=" + atrim(rst!identificador)
  'Num_remesa_ImpostEnv_Venta
End Sub

Sub registro_abonos_EspoImp(rst As Recordset, vNumRemesa As Long)
   dbtmp.Execute "update abonosclients set Numremesa_A22=" + atrim(vNumRemesa) + " where id=" + atrim(rst!identificador)
  'Num_remesa_ImpostEnv_Venta
End Sub

Sub registro_abonos_intracomunitaries(rst As Recordset, vNumRemesa As Long)
   dbtmp.Execute "update abonosclients set Numremesa_592=" + atrim(vNumRemesa) + " where id=" + atrim(rst!identificador)
  'Num_remesa_ImpostEnv_Venta
End Sub

Sub registro_compres_intracomunitaries(rst As Recordset, vNumRemesa As Long)
  dbtmp.Execute "update albaransbip set Num_remesa_ImpostEnv=" + atrim(vNumRemesa) + " where id=" + atrim(rst!identificador)
  'Num_remesa_ImpostEnv
End Sub
Sub guardar_historics_remesa(vnomtaula As String)
   dbtmp.Execute "insert into " + vnomtaula + " select * from taula_impost where apuntperdeclarar=true"
End Sub

Private Sub Command5_Click()
End Sub

Private Sub DBGrid1_Click()
  Me.Caption = DBGrid1.Columns(DBGrid1.Col).Width
End Sub

Private Sub etsumakilosmermaivendes_DblClick()
    Clipboard.Clear
    Clipboard.SetText etsumakilosmermaivendes.Caption + vbNewLine + etsumakilos.Caption
    MsgBox "Dades copiades al portapapers."
End Sub

Sub emplenar_reixa()
   DBGrid1.Visible = False
   dataintracomunitari.Refresh
   reixa.Rows = 1
   reixa.Visible = False
   While Not dataintracomunitari.Recordset.EOF
        reixa.Rows = reixa.Rows + 1
        For i = 0 To DBGrid1.Columns.Count - 1
          DBGrid1.Col = i
          reixa.TextMatrix(reixa.Rows - 1, i + 1) = DBGrid1.text
        Next i
        If Not dataintracomunitari.Recordset!apuntperdeclarar Then posarcolorfilareixa reixa.Rows - 1, &HC0C0FF
        dataintracomunitari.Recordset.MoveNext
   Wend
   'DBGrid1.Visible = True
   reixa.Visible = True
   If Form1.Visible Then Form1.SetFocus
   AppActivate "Manteniment Impost Envasos"
End Sub
Sub posarcolorfilareixa(vfila, vcolor As Double)
   reixa.Row = vfila
   reixa.RowSel = vfila
   reixa.Col = 1
   reixa.ColSel = reixa.Cols - 1
   reixa.CellBackColor = vcolor
End Sub
Sub configurar_reixa()
   Dim i As Byte
   Dim vamplades As Variant
   Dim vnoms As Variant
   vamplades = Array(500, 1100, 450, 800, 650, 450, 1300, 1550, 1150, 1550, 800, 2400, 1540, 2500)
   vnoms = Array("", "Assentament", "Apuntperdeclarar", "LotInp", "Concepte", "Clauproducte", "data", "justificant", "kilos", "Kg_Base_Imposable", "regimfiscal", "nomdestinatari", "nifdestinatari", "observacions")
   reixa.Cols = 13 + 1
   For i = 0 To reixa.Cols - 1
       reixa.ColWidth(i) = vamplades(i)
       reixa.TextMatrix(0, i) = vnoms(i)
   Next i
End Sub
Function ObtenerLíneaComando(Optional MaxArgs)
    'Declara las variables.
    Dim c, LíneaComando, LonLínComando, ArgIn, i, NúmArgs
    'Ver si MaxArgs está.
    If IsMissing(MaxArgs) Then MaxArgs = 10
    'Crea una matriz del tamaño correcto.
    ReDim ArgArray(MaxArgs)
    NúmArgs = 0: ArgIn = False
    'Obtiene los argumentos de la línea de comandos.
    LíneaComando = Command()
    LonLínComando = Len(LíneaComando)
    'Recorre la línea de comando carácter a carácter
    'a la vez.

For i = 1 To LonLínComando
        c = Mid(LíneaComando, i, 1)
        'Comprueba espacio o tabulación.
        If (c <> " " And c <> vbTab) Then
            'Ningún espacio o tabulación.
            'Comprueba si está en el argumento.
            If Not ArgIn Then
            'Empieza el nuevo argumento.
            'Comprueba para más argumentos.
                If NúmArgs = MaxArgs Then Exit For
                    NúmArgs = NúmArgs + 1
                    ArgIn = True
                End If
            'Agrega el carácter al argumento actual.

ArgArray(NúmArgs) = ArgArray(NúmArgs) + c
        Else
            'Encontró un espacio o tabulador.
            'Establece ArgIn a False.
            ArgIn = False
        End If
    Next i
    'Redimensiona la matriz lo suficiente para contener los argumentos.
    'ReDim Preserve ArgArray(NúmArgs)
    'Devuelve la matriz en nombre de la función.
    ObtenerLíneaComando = ArgArray()
End Function

Private Sub Form_Load()
   Dim arguments As Variant
   
   arguments = ObtenerLíneaComando
   cami = llegir_ini("General", "cami", "comandes.ini")
   If UCase(Environ("computername")) = "SERVERPRODU" Then cami = "C:\Dades\progcomandes\dades\comandes.mdb"
   Me.Caption = cami
   Set dbtmp = OpenDatabase(rutadelfitxer(cami) + "ImpostEnvasos.mdb")
   
   If UCase(arguments(1)) = "ABONOS" Then
         formabonosclients.Show 1
         End
   End If
   dataintracomunitari.DatabaseName = rutadelfitxer(cami) + "ImpostEnvasos.mdb"
   dbtmp.Execute "delete * from taula_impost"
   dataintracomunitari.RecordSource = "select * from taula_impost order by concepte,data"
   dataintracomunitari.Refresh
   configurar_reixa
   emplenar_reixa
  ' actualitzar_factures_merma_SAP
   'climitperiode = Format(DateAdd("d", -1, "01/" + Trim(Month(Now)) + "/" + Trim(Year(Now))), "dd/mm/yyyy")
   combomes.ListIndex = Format(DateAdd("d", -1, "01/" + Trim(Month(Now)) + "/" + Trim(Year(Now))), "m") - 1
  
End Sub

Sub actualitzar_factures_merma_SAP()
  Dim vsql As String
  Dim rst As Recordset
  Dim rstfacts As Recordset
  Dim vsqlwhere As String
  Dim vultimadata As Date
  Exit Sub
  vultimadata = "1/1/2023"
  Set rstfacts = dbtmp.OpenRecordset("select * from facturesSAPreciclatge")
  If Not rstfacts.EOF Then vultimadata = rstfacts!datafactura
  vsql = "SELECT First(Importada_LiniesFacturesSAP_Inplacsa.Dataalbara) AS DataFactura,First(Importada_LiniesFacturesSAP_Inplacsa.ItemCode) AS Tipus, Sum(IIf([Quantity] Is Null,0,[Quantity]/1000000)) AS Kg, Importada_LiniesFacturesSAP_Inplacsa.NumFact"
  vsql = vsql + ", First(Importada_LiniesFacturesSAP_Inplacsa.NomClient) AS nomclient, First(Importada_LiniesFacturesSAP_Inplacsa.Codicomptable) AS PrimeroDeCodicomptable, First(Importada_LiniesFacturesSAP_Inplacsa.Nif) AS nifclient"
  vsql = vsql + " From Importada_LiniesFacturesSAP_Inplacsa "
  'reciclatge VERD
  vsqlwhere = " Where (((Importada_LiniesFacturesSAP_Inplacsa.ItemCode) = 'DESPVERD01') and (Importada_LiniesFacturesSAP_Inplacsa.Dataalbara>#" + atrim(Format(vultimadata, "mm/dd/yy")) + "#) And ((Importada_LiniesFacturesSAP_Inplacsa.CANCELED) = 'N')) GROUP BY Importada_LiniesFacturesSAP_Inplacsa.NumFact;"
'  Clipboard.Clear
'  Clipboard.SetText vsql + vsqlwhere
  Set rst = dbtmp.OpenRecordset(vsql + vsqlwhere)
  afegir_factures_SAP_reciclatge rst, rstfacts
  vsqlwhere = " Where (((Importada_LiniesFacturesSAP_Inplacsa.ItemCode) = 'DESPVERD02') and (Importada_LiniesFacturesSAP_Inplacsa.Dataalbara>#" + atrim(Format(vultimadata, "mm/dd/yy")) + "#) And ((Importada_LiniesFacturesSAP_Inplacsa.CANCELED) = 'N')) GROUP BY Importada_LiniesFacturesSAP_Inplacsa.NumFact;"
  Set rst = dbtmp.OpenRecordset(vsql + vsqlwhere)
  afegir_factures_SAP_reciclatge rst, rstfacts
  'reciclatge BLAU
  vsqlwhere = " Where (((Importada_LiniesFacturesSAP_Inplacsa.ItemCode) = 'DESPBLAU01') and (Importada_LiniesFacturesSAP_Inplacsa.Dataalbara>#" + atrim(Format(vultimadata, "mm/dd/yy")) + "#) And ((Importada_LiniesFacturesSAP_Inplacsa.CANCELED) = 'N')) GROUP BY Importada_LiniesFacturesSAP_Inplacsa.NumFact;"
  Set rst = dbtmp.OpenRecordset(vsql + vsqlwhere)
  afegir_factures_SAP_reciclatge rst, rstfacts
  vsqlwhere = " Where (((Importada_LiniesFacturesSAP_Inplacsa.ItemCode) = 'DESPBLAU02') and (Importada_LiniesFacturesSAP_Inplacsa.Dataalbara>#" + atrim(Format(vultimadata, "mm/dd/yy")) + "#) And ((Importada_LiniesFacturesSAP_Inplacsa.CANCELED) ='N')) GROUP BY Importada_LiniesFacturesSAP_Inplacsa.NumFact;"
  Set rst = dbtmp.OpenRecordset(vsql + vsqlwhere)
  afegir_factures_SAP_reciclatge rst, rstfacts
  'reciclatge VERMELL
  vsqlwhere = " Where (((Importada_LiniesFacturesSAP_Inplacsa.ItemCode) = 'DESPVERMELL01') and (Importada_LiniesFacturesSAP_Inplacsa.Dataalbara>#" + atrim(Format(vultimadata, "mm/dd/yy")) + "#) And ((Importada_LiniesFacturesSAP_Inplacsa.CANCELED) = 'N')) GROUP BY Importada_LiniesFacturesSAP_Inplacsa.NumFact;"
  Set rst = dbtmp.OpenRecordset(vsql + vsqlwhere)
  afegir_factures_SAP_reciclatge rst, rstfacts
  vsqlwhere = " Where (((Importada_LiniesFacturesSAP_Inplacsa.ItemCode) = 'DESPVERMELL02') and (Importada_LiniesFacturesSAP_Inplacsa.Dataalbara>#" + atrim(Format(vultimadata, "mm/dd/yy")) + "#) And ((Importada_LiniesFacturesSAP_Inplacsa.CANCELED) = 'N')) GROUP BY Importada_LiniesFacturesSAP_Inplacsa.NumFact;"
  Set rst = dbtmp.OpenRecordset(vsql + vsqlwhere)
  afegir_factures_SAP_reciclatge rst, rstfacts
  
  
  Set rst = Nothing
End Sub
Sub afegir_factures_SAP_reciclatge(rst As Recordset, rstfacts As Recordset)
 While Not rst.EOF
     rstfacts.FindFirst "numerofactura=" + atrim(rst!numfact) + " and tipus='" + atrim(rst!tipus) + "'"
     If rstfacts.NoMatch Then
        rstfacts.AddNew
        rstfacts!datafactura = rst!datafactura
        rstfacts!numerofactura = rst!numfact
        rstfacts!tipus = rst!tipus
        rstfacts!kgfactura = rst!kg
        rstfacts!nomproveidor = atrim(rst!nomclient)
        rstfacts!nifproveidor = atrim(rst!nifclient)
        rstfacts.Update
     End If
     rst.MoveNext
  Wend
End Sub
Sub crear_resum_abonos_Esp_i_Imp()
   Dim rst As Recordset
   Dim rstabonos As Recordset
   Dim rstimpost As Recordset
   Dim rstfactura As Recordset
   Dim rstc As Recordset
   Dim vwere As String
   Dim vsql As String
   Dim vsumaEspImp As Double
   Dim vkgcapa As Double
   Dim vtotalimpost As Double
   
   Set rstabonos = dbtmp.OpenRecordset("select * from abonosclients where TOTAIMPOST>0 AND numremesa_a22=0 or numremesa_a22=null")
   Set rstimpost = dbtmp.OpenRecordset("select * from taula_impost")
   
   While Not rstabonos.EOF
           'vwere = " (PaisVenta<>'ES' or regimfiscal<>'') and KgVentaAd_Intracom>0 and (((ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Intra) Is Null Or (ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Intra)=0)) "
           Set rstc = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(rstabonos!lotinplacsa))
           vwere = " (impostenvasos.comanda=" + atrim(rstc!comanda) + " or impostenvasos.comanda=" + atrim(rstc!linkcomanda1) + " or impostenvasos.comanda=" + atrim(rstc!linkcomanda2) + ") "
           vsql = "SELECT ImpostEnvasos.*, capcaleraalbara.*, (select distinct datafactura from Importada_LiniesFacturesSAP_Inplacsa where NumFact=[capcaleraalbara].[numfacturaSAP] ) AS Datafactura, materials.tanpercentimpostenvasos, clients.nom AS nomclient, clients_codisSAP.nif "
           vsql = vsql + " FROM (((ImpostEnvasos LEFT JOIN capcaleraalbara ON ImpostEnvasos.numalbara = capcaleraalbara.numalbara) LEFT JOIN (comandes LEFT JOIN materials ON comandes.materialex = materials.codi) ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN clients_codisSAP ON capcaleraalbara.codiclient = clients_codisSAP.codiSAP "
        
        
           Set rst = dbtmp.OpenRecordset(vsql + " where " + vwere)
           While Not rst.EOF
               vsumaEspImp = cadbl(rst!kgventaespanya) + cadbl(rst!kgventaimp_mes_esp)
               If vsumaEspImp > 0 Then
                vkgcapa = (totalKgtoteslescapes(rst!comanda, "A22") * 100) / (totalKgtoteslescapes(rst!comanda, "A22") + totalKgtoteslescapes(rst!comanda, "592")) 'saber el % DEL TOTAL
                vtotalimpost = (rstabonos!totaimpost / 0.45) * (vkgcapa / 100)
                vkgcapa = (vsumaEspImp * 100) / totalKgtoteslescapes(rst!comanda, "A22") 'saber el % DE CAPA
                If rstabonos!lotinplacsa = 212407 Then Stop
                vkgcapa = vtotalimpost * (vkgcapa / 100)
                rstimpost.AddNew
                rstimpost!concepte = 4
                rstimpost!clauproducte = IIf(vtipusproveidor = "Importació", "B", "G")
                rstimpost!Data = rstabonos!datafacturaabono
                rstimpost!justificant = rstabonos!numfacturaabono
                rstimpost!kilos = Redondejar((vkgcapa * ((rst!tanpercentimpostenvasos / 100) + 1)) - vkgcapa, 3)
                rstimpost!kilosnoreciclats = Redondejar(vkgcapa, 3) 'Redondejar(rstabonos!totaimpost, 3)
                vsumakilos2 = vsumakilos2 + rstimpost!kilosnoreciclats
                vsumakilos = vsumakilos + rstimpost!kilos
                rstimpost!regimfiscal = "A"
                rstimpost!nomdestinatari = rst!nomclient
                rstimpost!nifdestinatari = rst!nif
                rstimpost!observacions = "Abono no Espanya amb compra Espanya o Importació."
                rstimpost!lotinplacsa = rst!comanda
               ' If DateDiff("m", rstabonos!datafacturaabono, climitperiode) >= 0 Then
                  If cadbl(rstimpost!justificant) > 0 Then
                        rstimpost!apuntperdeclarar = True
                  End If
                'End If
                rstimpost!identificador = rstabonos!id
                rstimpost.Update
               End If
               rst.MoveNext
           Wend
       rstabonos.MoveNext
   Wend
fi:
       Set rstc = Nothing
   Set rst = Nothing
   Set rstimpost = Nothing
   dataintracomunitari.Refresh
End Sub

Sub crear_resum_abonos_intracomunitaris()
   Dim rst As Recordset
   Dim rstabonos As Recordset
   Dim rstimpost As Recordset
   Dim rstfactura As Recordset
   Dim rstc As Recordset
   Dim vwere As String
   Dim vsql As String
   Dim vkgcapa As Double
   Dim vtotalimpost As Double
   
   Set rstabonos = dbtmp.OpenRecordset("select * from abonosclients where numremesa_592=0 or numremesa_592=null")
   Set rstimpost = dbtmp.OpenRecordset("select * from taula_impost")
   
   While Not rstabonos.EOF
           'vwere = " (PaisVenta<>'ES' or regimfiscal<>'') and KgVentaAd_Intracom>0 and (((ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Intra) Is Null Or (ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Intra)=0)) "
           Set rstc = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(rstabonos!lotinplacsa))
           vwere = " (impostenvasos.comanda=" + atrim(rstc!comanda) + " or impostenvasos.comanda=" + atrim(rstc!linkcomanda1) + " or impostenvasos.comanda=" + atrim(rstc!linkcomanda2) + ") "
           vsql = "SELECT ImpostEnvasos.*, capcaleraalbara.*, (select distinct datafactura from Importada_LiniesFacturesSAP_Inplacsa where NumFact=[capcaleraalbara].[numfacturaSAP] ) AS Datafactura, materials.tanpercentimpostenvasos, clients.nom AS nomclient, clients_codisSAP.nif "
           vsql = vsql + " FROM (((ImpostEnvasos LEFT JOIN capcaleraalbara ON ImpostEnvasos.numalbara = capcaleraalbara.numalbara) LEFT JOIN (comandes LEFT JOIN materials ON comandes.materialex = materials.codi) ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN clients_codisSAP ON capcaleraalbara.codiclient = clients_codisSAP.codiSAP "
           Set rst = dbtmp.OpenRecordset(vsql + " where " + vwere)
           While Not rst.EOF
               If cadbl(rst!kgventaad_intracom) > 0 Then
                vkgcapa = (totalKgtoteslescapes(rst!comanda, "592") * 100) / (totalKgtoteslescapes(rst!comanda, "A22") + totalKgtoteslescapes(rst!comanda, "592")) 'saber el % TOTAL
                vtotalimpost = (rstabonos!totaimpost / 0.45) * (vkgcapa / 100)
                vkgcapa = (cadbl(rst!kgventaad_intracom) * 100) / totalKgtoteslescapes(rst!comanda, "592") 'saber el % DE CAPA
                vkgcapa = vtotalimpost * (vkgcapa / 100)
               
                rstimpost.AddNew
                rstimpost!concepte = 4
                rstimpost!clauproducte = "B"
                rstimpost!Data = rstabonos!datafacturaabono
                rstimpost!justificant = rstabonos!numfacturaabono
                rstimpost!kilos = Redondejar((vkgcapa * ((rst!tanpercentimpostenvasos / 100) + 1)) - vkgcapa, 3)
                rstimpost!kilosnoreciclats = Redondejar(vkgcapa, 3) 'Redondejar(rstabonos!totaimpost, 3)
                'vsumakilos2 = vsumakilos2 - rstimpost!kilosnoreciclats
                'vsumakilos = vsumakilos - rstimpost!kilos
                rstimpost!regimfiscal = "A"
                rstimpost!nomdestinatari = rst!nomclient
                rstimpost!nifdestinatari = rst!nif
                rstimpost!observacions = "Abono no Espanya amb compra Intracomunitaria."
                rstimpost!lotinplacsa = rst!comanda
               ' If DateDiff("m", rstabonos!datafacturaabono, climitperiode) >= 0 Then
                  If cadbl(rstimpost!justificant) > 0 Then
                        rstimpost!apuntperdeclarar = True
                  End If
               ' End If
                rstimpost!identificador = rstabonos!id
                rstimpost.Update
               End If
               rst.MoveNext
           Wend
       rstabonos.MoveNext
   Wend
fi:
   Set rstc = Nothing
   Set rst = Nothing
   Set rstimpost = Nothing
   dataintracomunitari.Refresh
End Sub
Function totalKgtoteslescapes(vnumc As Double, vtipus As String) As Double
   Dim rst As Recordset
   Dim rsti As Recordset
   Set rst = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(vnumc))
   If rst.EOF Then GoTo fi
   Set rsti = dbtmp.OpenRecordset("select sum(kgventaespanya+kgventaimp_mes_esp) as T_A22,sum(kgventaad_intracom) as T_592 from impostenvasos where comanda=" + atrim(rst!comanda) + " or comanda=" + atrim(cadbl(rst!linkcomanda1)) + " or comanda=" + atrim(cadbl(rst!linkcomanda2)))
   If rsti.EOF Then GoTo fi
   If vtipus = "592" Then totalKgtoteslescapes = cadbl(rsti!T_592)
   If vtipus = "A22" Then totalKgtoteslescapes = cadbl(rsti!T_A22)
   'If vtipus = "A22" Then If cadbl(rsti!T_A22) > 0 Then hihamaterialIMPOST = True
   'If vtipus = "592" Then If cadbl(rsti!T_592) > 0 Then hihamaterialIMPOST = True
fi:
   Set rst = Nothing
   Set rsti = Nothing

End Function
Sub crear_resum_ventes_intracomunitaries()
   Dim rst As Recordset
   Dim rstimpost As Recordset
   Dim rstfactura As Recordset
   Dim vwere As String
   Dim vsql As String
   
   Set rstimpost = dbtmp.OpenRecordset("select * from taula_impost")
   
   
   vwere = " (PaisVenta<>'ES' or regimfiscal<>'') and KgVentaAd_Intracom>0 and (((ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Intra) Is Null Or (ImpostEnvasos.Num_remesa_ImpostEnv_Venta_Intra)=0)) "
   vsql = "SELECT ImpostEnvasos.*, capcaleraalbara.*, (select distinct datafactura from Importada_LiniesFacturesSAP_Inplacsa where NumFact=[capcaleraalbara].[numfacturaSAP] ) AS Datafactura, materials.tanpercentimpostenvasos, clients.nom AS nomclient, clients_codisSAP.nif "
   vsql = vsql + " FROM (((ImpostEnvasos LEFT JOIN capcaleraalbara ON ImpostEnvasos.numalbara = capcaleraalbara.numalbara) LEFT JOIN (comandes LEFT JOIN materials ON comandes.materialex = materials.codi) ON ImpostEnvasos.comanda = comandes.comanda) LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN clients_codisSAP ON capcaleraalbara.codiclient = clients_codisSAP.codiSAP "


   Set rst = dbtmp.OpenRecordset(vsql + " where " + vwere)
   While Not rst.EOF
        rstimpost.AddNew
        rstimpost!concepte = 2
        rstimpost!clauproducte = "B"
        rstimpost!Data = rst!datafactura
        rstimpost!justificant = rst!numfacturasap
        rstimpost!kilos = Redondejar((rst!kgventaad_intracom * ((rst!tanpercentimpostenvasos / 100) + 1)) - rst!kgventaad_intracom, 3)
        rstimpost!kilosnoreciclats = Redondejar(rst!kgventaad_intracom, 3)
        vsumakilos2 = vsumakilos2 + rstimpost!kilosnoreciclats
        vsumakilos = vsumakilos + rstimpost!kilos
        rstimpost!regimfiscal = IIf(atrim(rst!regimfiscal) = "", "A", rst!regimfiscal)
        rstimpost!nomdestinatari = rst!nomclient
        rstimpost!nifdestinatari = rst!nif
        rstimpost!observacions = "Venta no Espanya amb compra Intracomunitaria."
        If rstimpost!regimfiscal <> "A" Then rstimpost!observacions = "Venta amb règim fiscal exempt (Lletra " + rstimpost!regimfiscal + ")"
        rstimpost!lotinplacsa = rst!comanda
        If DateDiff("m", rst!datafactura, climitperiode) >= 0 Then
          If cadbl(rstimpost!justificant) > 0 Then
                rstimpost!apuntperdeclarar = True
          End If
        End If
        rstimpost!identificador = rst!id
        rstimpost.Update
        rst.MoveNext
   Wend
   
fi:
   Set rst = Nothing
   Set rstimpost = Nothing
   dataintracomunitari.Refresh
End Sub
Sub crear_resum_compres_intracomunitaries()
   Dim rst As Recordset
   Dim rstimpost As Recordset
   Dim rstfactura As Recordset
   If Not IsDate(climitperiode) Then MsgBox "La data posada com a limit no es correcte.": Exit Sub
   dbtmp.Execute "delete * from taula_impost"
   Set rstimpost = dbtmp.OpenRecordset("select * from taula_impost")
   If Not rstimpost.EOF Then MsgBox "No s'han eliminat els registres temporals.": GoTo fi
   dataintracomunitari.Refresh
   DoEvents
   Set rst = dbtmp.OpenRecordset("select * from agrupacio_dades_albaransbip") ' where Num_remesa_ImpostEnv=0 or Num_remesa_ImpostEnv=null")
   While Not rst.EOF
     Set rstfactura = dbtmp.OpenRecordset("select * from Importada_Albarans_Compres_Inplacsa where  NumAtCard='" + atrim(rst!numalbaraprov) + "' and codicomptable='" + atrim(rst!codiproveidorcomercial) + "' order by facturaprov DESC")
     vdata = IIf(rstfactura.EOF, 0, rstfactura!FechaContableFactura)
     If vdata = 0 Then vdata = climitperiode
     'If DateDiff("m", vdata, climitperiode) = 0 Or DateDiff("m", vdata, climitperiode) = 1 Then
     If DateDiff("m", vdata, climitperiode) = 0 Then
        If Not rstfactura.EOF Then
            rstimpost.AddNew
            rstimpost!concepte = 1
            rstimpost!lotinplacsa = rst!numalbara
            rstimpost!clauproducte = "B"
            rstimpost!Data = vdata 'rstfactura!datafactura
            rstimpost!datacomptableSAP = rstfactura!FechaContableFactura
            rstimpost!justificant = rstfactura!facturaprov
            rstimpost!kilos = rst!kgbaseimposableimpostenvasos
            rstimpost!kilosnoreciclats = rst!KgImpostEnvasos
                vsumakilos2 = vsumakilos2 + rstimpost!kilosnoreciclats
                vsumakilos = vsumakilos + rstimpost!kilos
            rstimpost!regimfiscal = "A"
            rstimpost!euroscompra = cadbl(rst!importnet)
            rstimpost!nomdestinatari = rst!nomproveidor
            rstimpost!nifdestinatari = rst!nif
            rstimpost!observacions = "Compra Intracomunitaria"
            If rstimpost!justificant <> "" Then
                    rstimpost!apuntperdeclarar = True
            End If
            If rst!Num_remesa_ImpostEnv <> 0 And Not IsNull(rst!Num_remesa_ImpostEnv) Then
               rstimpost!apuntperdeclarar = False
            End If
            rstimpost!identificador = rst!id
            rstimpost.Update
         End If
     End If
     rst.MoveNext
   Wend
   
fi:
   Set rst = Nothing
   Set rstimpost = Nothing
   dataintracomunitari.Refresh
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label2_DblClick()

End Sub

Private Sub Form_Resize()
   On Error Resume Next
   bguardarCSV.Left = Form1.Width - bguardarCSV.Width - 500
   bfacturescompres.Left = bguardarCSV.Left - bfacturescompres.Width
   DBGrid1.Width = Form1.Width - 500
   DBGrid1.Height = Form1.Height - DBGrid1.Top - 900
End Sub
Sub generar_liniaventa(vidTAULAIMPOST As Long, vKGadeclarar As Double, rst As Recordset, vliniaventes As String)
    Dim rstalb As Recordset
    Dim rstsap As Recordset
    Dim rstventa As Recordset
    Dim vfactventa As String
    Dim vdataventa As Date
  '  Set rstventa = dbtmp.OpenRecordset("SELECT * from impostenvasos where id=" + atrim(vidTAULAIMPOST))
  '  If rstventa.EOF Then MsgBox "NO S'HA TROBAT LA VENTA A LA TAULA IMPOSTENVASOS"
  '  Set rstventa = dbtmp.OpenRecordset("select * from liniesalbara where numalbara=" + atrim(rstventa!numalbara) + " and lotinplacsa=" + atrim(numcomandaprincipal(rstventa!comanda)))
  '  If rstventa.EOF Then MsgBox "No s'ha trobat el registre IDENTIFICADOR a liniesalbarans els KG DE QUANTITAT VENUDA S'AGAFARÀN DE LES LINIES D'ALBARÀ"
    vfactventa = atrim(rst!justificant)
    vdataventa = atrim(rst!Data)
  '  If vfactventa = 23000445 Then Stop
    comprovarsiDUA vfactventa, vdataventa
    vkgbruts = Redondejar(cadbl(vKGadeclarar), 3)
    If vkgbruts = 0 Then Stop
    vliniaventes = atrim(vfactventa) + ";" + atrim(vdataventa) + ";" + atrim(rst!nifdestinatari) + ";" + atrim(rst!nomdestinatari) + ";" + atrim(vkgbruts) + ";" + atrim(Redondejar(vkgbruts * 0.45, 3))
    
End Sub
Sub generar_liniacompra(vnumpalet As Double, vnumbobina As Double, vliniacompra As String, Optional vnommaterial As String, Optional vtipusproveidor As String, Optional vTKgParcial As Double)
   Dim rstsap As Recordset
   Dim rst As Recordset
   Dim rstprov As Recordset
   Dim vsumakgimpost As Double
   Dim vnumalb As String
   Dim vdomiciliprov  As String
   Dim vDUA As String
   Dim vdatafacturaprov As String
   Dim vnumfacturaproveidor As String
   Dim vnominifproveidor As String
   
   Set rst = dbtmpb.OpenRecordset("select * from palets where idpalet=" + atrim(vnumpalet))
   If cadbl(rst!link_numpalet) > 0 Then
       Set rst = dbtmpb.OpenRecordset("select * from palets where idpalet=" + atrim(rst!link_numpalet))
       vnumpalet = rst!idpalet
   End If
   If cadbl(rst!numalb) = 1 Then
      ' vnumpalet = 0
       Set rst = dbtmpb.OpenRecordset("SELECT rebobinadores.comanda, rebobinadores.tipus, bobinesreb.numerodebobina, bobinesentreb.palet, bobinesentreb.bobina FROM (bobinesreb RIGHT JOIN rebobinadores ON bobinesreb.controlid = rebobinadores.Id) LEFT JOIN bobinesentreb ON bobinesreb.Id = bobinesentreb.id WHERE (((rebobinadores.comanda)=" + atrim(rst!numlot) + ") AND ((rebobinadores.tipus)='F') AND ((bobinesreb.numerodebobina)=" + atrim(vnumbobina) + "));")
       If Not rst.EOF Then vnumpalet = rst!palet
   End If
   Set rst = dbtmp.OpenRecordset("select * from albaransbip where numpalet=" + atrim(vnumpalet))
   If rst.EOF Then MsgBox "No s'ha trobat el palet a albaransbip"
   vnommaterial = atrim(rst!descripcio)
   vnumalb = atrim(rst!numalbaraprov)
   
   'Clipboard.Clear
   'Clipboard.SetText "select * from Factures_albarans_Inplacsa where U_GSP_INFABLOTE='" + atrim(rst!numlotproveidor) + "' and numatcard='" + atrim(vnumalb) + "' order by datafactura desc"
   vlotproveidor = rst!numlotproveidor
   If vnumalb = "9000157682" And vlotproveidor = "2000223130-50" Then vlotproveidor = "2000223442-50"
   If vnumalb = "9000157682" And vlotproveidor = "2000223442-10" Then vlotproveidor = "2000223442-80"
   If vlotproveidor = "171960" Then vlotproveidor = "171690"
   Set rst = dbsap.OpenRecordset("select * from Factures_albarans_Inplacsa where U_GSP_INFABLOTE='" + atrim(vlotproveidor) + "' and numatcard='" + atrim(vnumalb) + "' order by datafactura desc")
   If rst.EOF Then MsgBox "No s'ha trobat el numero d'albarà de proveidor a les Factures de SAP"
   vnumfacturaproveidor = ""
   vdatafacturaprov = ""
   vnominifproveidor = ""
   'Trec l'avis pero es grava al fitxer   If atrim(rst!facturaprov) = "" And vnumalb <> "232373" Then MsgBox "No s'ha trobat la factura del LOT:" + atrim(vlotproveidor) + " de l'albarà " + atrim(vnumalb)
   Set rstsap = dbsap.OpenRecordset("select * from Facturesiliniescompres where numfacturaprov='" + atrim(rst!facturaprov) + "' order by U_IFG_340Dua desc")
   If Not rstsap.EOF Then If atrim(rstsap!U_IFG_340Dua) = "" Then Set rstsap = dbsap.OpenRecordset("select * from Facturesiliniescompres where itemcode='IMP_ENV' AND numfacturaprov='" + atrim(rst!facturaprov) + "'")
   If rstsap.EOF Then
       If vtipusproveidor = "Importació" Then vmsgERROR = vmsgERROR + "Aquesta factura es d'importació i no hi ha DUA associat." + vbNewLine
       vmsgERROR = vmsgERROR + "No s'ha trobat la Factura al SAP del LOT:" + atrim(vlotproveidor) + " de l'albarà " + atrim(vnumalb) + vbNewLine
       vnumfacturaproveidor = ""
       vdatafacturaprov = 0
      Else
        vnumfacturaproveidor = atrim(rstsap!numfacturaprov)
        vdatafacturaprov = rstsap!datafacturaprov
   End If
'   If vtipusproveidor = "Importació" And vnumfacturaproveidor <> "25ES0008553001YDL0" Then Stop
   'If vnumfacturaproveidor = "" Then Stop
   vnominifproveidor = ""
   If Not rstsap.EOF Then
     Set rstprov = dbtmp.OpenRecordset("select * from proveidors_codissap where codiSAP=" + atrim(rstsap!codicomptableproveidor))
     If Not rstprov.EOF Then
          vdomiciliprov = atrim(rstprov!provincia)
          vnominifproveidor = atrim(rstprov!nomproveidor) + " " + atrim(rstprov!nif)
     End If
   End If
 ' AMB AIXÓ SUMO EL TOTAL D'IMPOST PAGAT A ADUANA potser no faix servir l'import
   vDUA = ""
   If atrim(rstsap!U_IFG_340Dua) <> "" Then
        vDUA = atrim(rstsap!U_IFG_340Dua)
        Set rstsap = dbsap.OpenRecordset("select * from Facturesiliniescompres where itemcode='IMP_ENV' AND " + IIf(atrim(rstsap!U_IFG_340Dua) <> "", "U_IFG_340Dua='" + atrim(rstsap!U_IFG_340Dua) + "'", "numfacturaprov='" + atrim(rst!facturaprov) + "'"))
        If rstsap.EOF Then MsgBox "No he trobat l'import de l'impost pagat a aduanes."
   End If
   While Not rstsap.EOF
      vDUA = atrim(rstsap!U_IFG_340Dua)
      If vDUA = "" Then
          vDUA = IIf(cadbl(Mid(atrim(rstsap!numfacturaprov), 1, 2)) > 0 And Mid(UCase(atrim(rstsap!numfacturaprov)), 3, 2) = "ES", atrim(rstsap!numfacturaprov), "")
      End If
      If vDUA <> "" And cadbl(rstsap!Quantity) = 1 Then
           vsumakgimpost = vsumakgimpost + (cadbl(rstsap!Price) / 0.45)
           vDUA = atrim(rstsap!U_IFG_340Dua)
            Else: vsumakgimpost = vsumakgimpost + cadbl(rstsap!Quantity)
      End If
      rstsap.MoveNext
   Wend
   If vnumfacturaproveidor <> "" Then
            rstsap.MoveFirst
             ' Else: MsgBox "Num factura compra no trobada a SAP"
   End If
   
     'ATENCIÓ POSSO UN ESPAI DAVANT EL NUMERO DE FACTURA DE PROVEIDOR PERQUE EL CSV M'HO TRACTI COM A TEXTE I
      'NO PERDI ELS ZEROS DAVANT DE LA FACTURA
   'vliniacompra = " " + atrim(vnumfacturaproveidor) + ";" + atrim(vdatafacturaprov) + ";" + vtipusproveidor + ";" + vDUA + ";" + vnominifproveidor + ";" + vdomiciliprov + ";" + vnommaterial + ";" + atrim(vTKgParcial) + ";" + atrim(vTKgParcial) + ";" + atrim(vTKgParcial * 0.45)
     'MODIFICACIÓ TOTAL IMPOST PAGAT A LA FACTURA
     
    ' If vdatafacturaprov = "0" Then Stop
   vliniacompra = """" + atrim(vnumfacturaproveidor) + """" + ";" + atrim(vdatafacturaprov) + ";" + vtipusproveidor + ";" + vDUA + ";" + vnominifproveidor + ";" + vdomiciliprov + ";" + vnommaterial + ";" + atrim(vsumakgimpost) + ";" + atrim(vsumakgimpost) + ";" + atrim(vsumakgimpost * 0.45)
   
fi:
   If atrim(vliniacompra) = "" Then Stop
   If vnumfacturaproveidor <> "" Then
         buscar_factura_compra vnumfacturaproveidor, CVDate(vdatafacturaprov)
         sumar_kg_totalitzaciofacturacompra vnumfacturaproveidor, CVDate(vdatafacturaprov), vnominifproveidor, vTKgParcial
   End If
   
   
   Set rstsap = Nothing
   Set rst = Nothing
   Set rstprov = Nothing
End Sub
Sub sumar_kg_totalitzaciofacturacompra(vnumfact As String, vdata As Date, vnom As String, vkg As Double)
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from TotalsFacturesProveidorsA22 where numfactura='" + atrim(vnumfact) + "' and datafactura=#" + Format(vdata, "mm/dd/yyyy") + "#")
   If rst.EOF Then Set rst = dbtmp.OpenRecordset("select * from TotalsFacturesProveidorsA22 where numfactura='" + atrim(cadbl(vnumfact)) + "' and datafactura=#" + Format(vdata, "mm/dd/yyyy") + "#")
   If Not rst.EOF Then
       rst.Edit
       rst!kgtmp = cadbl(rst!kgtmp) + vkg
       rst.Update
       If cadbl(rst!kgtmp) + rst!kgacumulats > rst!kg Then
            'aixó vol dir que estem demanant mes kg a hisenda dels que tenia la factura
            If Not existeix(vNomFitxerControlFacturesCompra) Then
                Open vNomFitxerControlFacturesCompra For Output As #9
                  Else: Open vNomFitxerControlFacturesCompra For Append As #9
            End If
            Print #9, vnumfact + ";" + atrim(vdata) + ";" + vnom + ";" + atrim(vkg)
            Close #9
       End If
   End If
   Set rst = Nothing
End Sub
Sub buscar_factura_compra(vnumfacturacompra As String, vdatacompra As Date)
   Dim rstfactures As Recordset
   Dim vnomfitxer As String
   Set rstfactures = dbtmp.OpenRecordset("Select * from LlistaDirectoriFacturesSAP where ucase(nomfitxer) like '*" + UCase(vnumfacturacompra) + "*'")
   If Not rstfactures.EOF Then
      rstfactures.MoveLast: rstfactures.MoveFirst
      If rstfactures.RecordCount > 1 Then
         While Not rstfactures.EOF
            If Year(rstfactures!datafitxer) = Year(vdatacompra) Then GoTo cont
            rstfactures.MoveNext
         Wend
cont:
      End If
      If Not rstfactures.EOF Then
          vnomfitxer = vRutaFacturesPdfSap + rstfactures!rutarelativa + "\" + rstfactures!nomfitxer
          If existeix(vnomfitxer) Then Copiar_Fitxer vnomfitxer, vRutaFacturesCompres
            Else: Print #4, "Compra;" + vnumfacturacompra + ";No Trobada"
      End If
        Else: Print #4, "Compra;" + vnumfacturacompra + ";No Trobada"
   End If
   Set rstfactures = Nothing
End Sub
Sub buscar_factura_venda(vnumfacturavenda As String)
   Dim rstfactures As Recordset
   Dim vnomfitxer As String
   Set rstfactures = dbtmp.OpenRecordset("Select * from LlistaDirectoriFacturesSAP where nomfitxer like '" + vnumfacturavenda + "*'")
   If Not rstfactures.EOF Then
        vnomfitxer = vRutaFacturesPdfSap + rstfactures!rutarelativa + "\" + rstfactures!nomfitxer
        If existeix(vnomfitxer) Then Copiar_Fitxer vnomfitxer, vRutaFacturesVendes
          Else: Print #4, "Venda;" + vnumfacturavenda + ";No Trobada"
   End If
   Set rstfactures = Nothing
End Sub

Sub generar_liniamerma(vidTAULAIMPOST As Long, vKGadeclarar As Double, rst As Recordset, vliniaventes As String, vlletra As String, vnommaterial As String)
    Dim rstalb As Recordset
    Dim rstsap As Recordset
    Dim rstventa As Recordset
    Dim vfactventa As String
    Dim vdataventa As Date
    Dim rsttransport As Recordset
    Dim vwhere As String
    Dim vsql As String
    'Set rstventa = dbtmp.OpenRecordset("SELECT * from impostenvasos where id=" + atrim(vidTAULAIMPOST))
    'If rstventa.EOF Then MsgBox "NO S'HA TROBAT LA VENTA A LA TAULA IMPOSTENVASOS"
    'Set rstventa = dbtmp.OpenRecordset("select * from liniesalbara where numalbara=" + atrim(rstventa!numalbara) + " and lotinplacsa=" + atrim(numcomandaprincipal(rstventa!comanda)))
    'If rstventa.EOF Then MsgBox "No s'ha trobat el registre IDENTIFICADOR a liniesalbarans els KG DE QUANTITAT VENUDA S'AGAFARÀN DE LES LINIES D'ALBARÀ"
    vwhere = " where [impostenvasos].[id]=" + atrim(vidTAULAIMPOST)
    vsql = "SELECT ImpostEnvasos.id, ImpostEnvasos.idliniaalbara, transportistes.descripcio, clients_codisSAP.nomclient, clients_codisSAP.nif "
    vsql = vsql + " FROM ((transportistes RIGHT JOIN (capcaleraalbara RIGHT JOIN liniesalbara ON capcaleraalbara.numalbara = liniesalbara.numalbara) ON transportistes.codi = capcaleraalbara.id_transport) RIGHT JOIN ImpostEnvasos ON liniesalbara.id = ImpostEnvasos.idliniaalbara) LEFT JOIN clients_codisSAP ON capcaleraalbara.codiclient = clients_codisSAP.codiSAP "
    Set rsttransport = dbtmp.OpenRecordset(vsql + vwhere)
    
   'Trec el justificant perque SIDE diu que no correspon
'    vfactventa = atrim(rst!justificant)
  '  vdataventa = atrim(rst!Data)  'data del justificant

    vkgbruts = Redondejar(cadbl(vKGadeclarar), 3)
    'If vkgbruts = 0 Then Stop
    vnomtransport = ""
    vnominifclient = ""
    'If Not rsttransport.EOF Then
    '         vnomtransport = atrim(rsttransport![descripcio])
    '         vnominifclient = atrim(rsttransport!nomclient) + " " + atrim(rsttransport!nif)
    'End If
    vliniaventes = vlletra + ";" + atrim(vfactventa) + ";" + atrim(vdataventa) + ";" + vnominifclient + ";" + vnommaterial + ";" + atrim(vkgbruts) + ";" + vnomtransport + ";" + atrim(vkgbruts) + ";" + atrim(Redondejar(vkgbruts * 0.45, 3))
      
End Sub



Sub generar_liniaventa2(vidTAULAIMPOST As Long, vKGadeclarar As Double, rst As Recordset, vliniaventes As String, vlletra As String, vnommaterial As String)
    Dim rstalb As Recordset
    Dim rstsap As Recordset
    Dim rstventa As Recordset
    Dim vfactventa As String
    Dim vdataventa As Date
    Dim rsttransport As Recordset
    Dim rstCMR As Recordset
    Dim vwhere As String
    Dim vsql As String
    Dim vsql2 As String
    Dim vnumCMR As String
    Dim vincoterm As String
    Dim vnumFraTransport As String
    
    'Set rstventa = dbtmp.OpenRecordset("SELECT * from impostenvasos where id=" + atrim(vidTAULAIMPOST))
    'If rstventa.EOF Then MsgBox "NO S'HA TROBAT LA VENTA A LA TAULA IMPOSTENVASOS"
    'Set rstventa = dbtmp.OpenRecordset("select * from liniesalbara where numalbara=" + atrim(rstventa!numalbara) + " and lotinplacsa=" + atrim(numcomandaprincipal(rstventa!comanda)))
    'If rstventa.EOF Then MsgBox "No s'ha trobat el registre IDENTIFICADOR a liniesalbarans els KG DE QUANTITAT VENUDA S'AGAFARÀN DE LES LINIES D'ALBARÀ"
    vwhere = " where [impostenvasos].[id]=" + atrim(vidTAULAIMPOST)
    'vsql = "SELECT liniesalbara.lotinplacsa,capcaleraalbara.numalbara,ImpostEnvasos.id, ImpostEnvasos.idliniaalbara, transportistes.descripcio, clients_codisSAP.nomclient, clients_codisSAP.nif "
    vsql = "SELECT liniesalbara.lotinplacsa, capcaleraalbara.numalbara, ImpostEnvasos.id, ImpostEnvasos.idliniaalbara, transportistes.descripcio, clients_codisSAP.nomclient, clients_codisSAP.nif "
    'vsql = vsql + " FROM ((transportistes RIGHT JOIN (capcaleraalbara RIGHT JOIN liniesalbara ON capcaleraalbara.numalbara = liniesalbara.numalbara) ON transportistes.codi = capcaleraalbara.id_transport) RIGHT JOIN ImpostEnvasos ON liniesalbara.id = ImpostEnvasos.idliniaalbara) LEFT JOIN clients_codisSAP ON capcaleraalbara.codiclient = clients_codisSAP.codiSAP "
    vsql2 = " FROM ((((capcaleraalbara RIGHT JOIN liniesalbara ON capcaleraalbara.numalbara = liniesalbara.numalbara) RIGHT JOIN ImpostEnvasos ON liniesalbara.id = ImpostEnvasos.idliniaalbara) LEFT JOIN clients_codisSAP ON capcaleraalbara.codiclient = clients_codisSAP.codiSAP) LEFT JOIN Transportistes_avisos ON capcaleraalbara.numalbara = Transportistes_avisos.numalbara) LEFT JOIN transportistes ON Transportistes_avisos.coditransport = transportistes.codi "
    Set rsttransport = dbtmp.OpenRecordset(vsql + vsql2 + vwhere)
    
    vfactventa = atrim(rst!justificant)
    vdataventa = atrim(rst!Data)
'    comprovarsiDUA vfactventa, vdataventa
    vkgbruts = Redondejar(cadbl(vKGadeclarar), 3)
    vnumFraTransport = ""
    If vkgbruts = 0 Then Stop
    If Not rsttransport.EOF Then
             vnomtransport = atrim(rsttransport![descripcio])
             vnominifclient = atrim(rsttransport!nomclient) + " " + atrim(rsttransport!nif)
             Set rstCMR = dbtmp.OpenRecordset("select numeroavis from transportistes_avisos where numalbara=" + atrim(cadbl(rsttransport![numalbara])))
             vnumCMR = "No hi ha CMR fet"
             If Not rstCMR.EOF Then
                  vnumCMR = atrim(rstCMR!numeroavis)
                  vnumFraTransport = guardarCMRifacturaTransportista(vnumCMR)
             End If
             Set rstCMR = dbtmp.OpenRecordset("select incoterm_envio from comandes_extres where comanda=" + atrim(cadbl(rsttransport!lotinplacsa)))
             If Not rstCMR.EOF Then
                  vincoterm = atrim(rstCMR!incoterm_envio)
             End If
    End If
    
    vliniaventes = vlletra + ";" + atrim(vfactventa) + ";" + atrim(vdataventa) + ";" + vnominifclient + ";" + vnommaterial + ";" + atrim(vkgbruts) + ";" + vnomtransport + ";" + atrim(vkgbruts) + ";" + atrim(Redondejar(vkgbruts * 0.45, 3)) + ";" + vincoterm + ";_" + vnumCMR + ";" + vnumFraTransport
'   MsgBox vliniaventes
    If vfactventa <> "" Then buscar_factura_venda vfactventa
End Sub
Function treuresimbolsnovalidsnomfitxer(desc As String) As String
   desc = substituir(desc, "\", "_")
   desc = substituir(desc, "/", "_")
   desc = substituir(desc, "|", "_")
   desc = substituir(desc, ":", ";")
   desc = substituir(desc, "?", "¿")
   desc = substituir(desc, "*", "x")
   desc = substituir(desc, """", "'")
   desc = substituir(desc, ">", "+")
   desc = substituir(desc, "<", "-")
   treuresimbolsnovalidsnomfitxer = desc
End Function

Function guardarCMRifacturaTransportista(vnumCMR As String) As String
    Dim vdirectoriOrigenCMRSiFactures As String
    Dim vdirectoridesti As String
    Dim vnomfitxerFra As String
    Dim rst As Recordset
    vdirectoridesti = "c:\temp\requeriment\CmrsiFacturesdetransport"
    vdirectoriOrigenCMRSiFactures = "\\ord_copies\AlbaransSAPClients"
    vsql = "SELECT transportistes_factures_CMR.numeroCMR, transportistes_factures.ID,transportistes_factures.numerofactura, transportistes_factures.datafactura, transportistes_factures.escanejada FROM transportistes_factures LEFT JOIN transportistes_factures_CMR ON transportistes_factures.id = transportistes_factures_CMR.id "
    vsql = vsql + " where numeroCMR='" + atrim(vnumCMR) + "'"
    Set rst = dbtmp.OpenRecordset(vsql)
     
    If Not rst.EOF Then
        vnomfitxerFra = "FraTrans_" + atrim(Format(rst!id, "000000")) + " [" + Format(rst!datafactura, "dd-mm-yy") + "] " + treuresimbolsnovalidsnomfitxer(rst!numerofactura) + ".pdf"
        If existeix(vdirectoriOrigenCMRSiFactures + "\CMRs\CMR_" + vnumCMR + ".pdf") Then
              If Not existeix(vdirectoridesti + "\" + vnumCMR + ".pdf") Then
                   Copiar_Fitxer vdirectoriOrigenCMRSiFactures + "\CMRs\CMR_" + vnumCMR + ".pdf", vdirectoridesti + "\" + vnumCMR + ".pdf"
              End If
                Else: guardarCMRifacturaTransportista = " [No PDF del CMR] " + vnomfitxerFra
        End If
        If existeix(vdirectoriOrigenCMRSiFactures + "\FacturesTransport\" + vnomfitxerFra) Then
              If Not existeix(vdirectoridesti + "\" + vnomfitxerFra) Then
                   Copiar_Fitxer vdirectoriOrigenCMRSiFactures + "\FacturesTransport\" + vnomfitxerFra, vdirectoridesti + "\" + vnomfitxerFra
              End If
               Else: guardarCMRifacturaTransportista = " [No PDF de la FraTrans] " + vnomfitxerFra
        End If
        If guardarCMRifacturaTransportista = "" Then guardarCMRifacturaTransportista = atrim(rst!datafactura) + " Nº:" + atrim(rst!numerofactura)
          Else: guardarCMRifacturaTransportista = "Fra Relacionada amb el CMR No Trobada"
    End If
    Set rst = Nothing
End Function
Sub generar_liniacompra2(vnumpalet As Double, vnumbobina As Double, vliniacompra As String)
   Dim rstsap As Recordset
   Dim rst As Recordset
   Dim rstprov As Recordset
   Dim vsumakgimpost As Double
   Dim vnumalb As String
   Set rst = dbtmpb.OpenRecordset("select * from palets where idpalet=" + atrim(vnumpalet))
   If cadbl(rst!numalb) = 1 Then
       Set rst = dbtmpb.OpenRecordset("SELECT rebobinadores.comanda, rebobinadores.tipus, bobinesreb.numerodebobina, bobinesentreb.palet, bobinesentreb.bobina FROM (bobinesreb RIGHT JOIN rebobinadores ON bobinesreb.controlid = rebobinadores.Id) LEFT JOIN bobinesentreb ON bobinesreb.Id = bobinesentreb.id WHERE (((rebobinadores.comanda)=" + atrim(rst!numlot) + ") AND ((rebobinadores.tipus)='F') AND ((bobinesreb.numerodebobina)=" + atrim(vnumbobina) + "));")
       vnumpalet = rst!palet
   End If
   Set rst = dbtmp.OpenRecordset("select * from albaransbip where numpalet=" + atrim(vnumpalet))
   If rst.EOF Then MsgBox "No s'ha trobat el palet a albaransbip"
   vnumalb = atrim(rst!numalbaraprov)
   
   Set rst = dbsap.OpenRecordset("select * from Factures_albarans_Inplacsa where numatcard='" + atrim(vnumalb) + "'")
   If rst.EOF Then MsgBox "No s'ha trobat el numero d'albarà de proveidor a les Factures de SAP": GoTo fi
   Set rstsap = dbsap.OpenRecordset("select * from Facturesiliniescompres where itemcode='IMP_ENV' AND numfacturaprov='" + atrim(rst!facturaprov) + "'")
   If rstsap.EOF Then GoTo fi
   'Set rstprov = dbtmp.OpenRecordset("select * from proveidors_codissap where codiSAP=" + atrim(rstsap!codicomptableproveidor))
   'If rstprov.EOF Then GoTo fi
   'If rstsap!numfacturaprov = "23ES00081130081940" Then Stop
   While Not rstsap.EOF
      If (atrim(rstsap!U_IFG_340Dua) <> "" Or Mid(rstsap!numfacturaprov, 3, 2) = "ES") And cadbl(rstsap!Quantity) = 1 Then
           vsumakgimpost = vsumakgimpost + Redondejar(cadbl(rstsap!Price) / 0.45, 2)
            Else: vsumakgimpost = vsumakgimpost + cadbl(rstsap!Quantity)
      End If
      rstsap.MoveNext
   Wend
   rstsap.MoveFirst
   vliniacompra = atrim(rstsap!datafacturaprov) + ";" + atrim(rstsap!numfacturaprov) + ";" + atrim(vsumakgimpost)
fi:
   Set rstsap = Nothing
   Set rst = Nothing
   Set rstprov = Nothing
End Sub

Function metrestotalstipusproveidorXcomanda(vnumc As Double, vtipusproveidor As String, Optional vComptar500 As Boolean) As Double
    Dim rst As Recordset
    Dim vsqlcomptar500 As String
    Dim vsqltipusproveidor As String
    If vtipusproveidor <> "A22" Then vsqltipusproveidor = " and tipusproveidorIMPOST='" + vtipusproveidor + "'"
    If vtipusproveidor = "A22" Then vsqltipusproveidor = " and (tipusproveidorIMPOST='Importació' or tipusproveidorIMPOST='Espanyol')"
    If vtipusproveidor = "" Then vsqltipusproveidor = ""
    If Not vComptar500 Then vsqlcomptar500 = " AND ((Parcials.orcomassignacio)<>'500')"
    Set rst = dbtmp.OpenRecordset("SELECT Parcials.idpalet, Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE teimpost=true and (((Parcials.comanda)='" + atrim(vnumc) + "') " + vsqlcomptar500 + ")" + vsqltipusproveidor)
    While Not rst.EOF
       metrestotalstipusproveidorXcomanda = metrestotalstipusproveidorXcomanda + cadbl(rst!metres)
       rst.MoveNext
    Wend
    Set rst = Nothing
End Function

Private Sub mabonamentsclients_Click()
  formabonosclients.Show 1
End Sub

Private Sub Generar_requeriment_Mermes(Optional vNumRemesa As Double, Optional vAnnexarCSV As Boolean, Optional vvalorA As Double, Optional vvalorD As Double)
    Dim vliniacompra As String
    Dim vliniaventa As String
    Dim vnomfitxer As String
    Dim vnumalbaraprov As String
    Dim vTmtrsTots As Double
    Dim vmtrsGrupImp As Double
    Dim vmtrsGrupEsp As Double
    Dim vKgImp As Double
    Dim vKgEsp As Double
    Dim vTKgParcial As Double
    Dim rst2 As Recordset
    Dim rst As Recordset
    Dim rstTIPUS As Recordset
    Dim vsqltipusproveidor As String
    Dim vnumc As Double
    Dim vKgTots As Double
    Dim vcomptadordeliniesA As Double
    Dim vfactorKgaSumarA As Double
    Dim vcomptadordeliniesD As Double
    Dim vfactorKgaSumarD As Double
    Dim vcontA As Double
    Dim vcontD As Double
    Dim v As String
    Dim vFactor As Double
    Dim vnommaterial As String
    Dim vKgcomptarlinies As Double
    Dim vLiniaFeta As Boolean
    Dim vKgSumatsID As Double
    
    vvalorA = 0.00000000001  'Es perquè no demani valors
    
    vKgcomptarlinies = 25
    vmsgERROR = ""
    If cadbl(vNumRemesa) = 0 Then
         v = escullir_historic_Imp_i_Esp(v)
         vNumRemesa = cadbl(v)
          Else: v = vNumRemesa
    End If
    If cadbl(vvalorA) = 0 Then
      vvalorA = cadbl(InputBox("Entra el valor de la casella MERMES IMPORTACIO B. En Kg", "Valor B"))
      If vvalorA = 0 Then MsgBox "No pot ser zero.", vbCritical, "Error": Exit Sub
      vvalorD = cadbl(InputBox("Entra el valor de la casella MERMES ESPANYA G. En Kg", "Valor G"))
      If vvalorD = 0 Then MsgBox "No pot ser zero.", vbCritical, "Error": Exit Sub
    End If
      'Si es vol fer en Euros enlloc de Kg treure les 2 seguents linies
    vvalorA = vvalorA * 0.45
    vvalorD = vvalorD * 0.45
    
    vfactorKgaSumarA = 0.000001
    vfactorKgaSumarD = 0.000001
    
    
    Set dbtmpb = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
    Set dbsap = OpenDatabase(rutadelfitxer(cami) + "connexiosap.mdb")
tornaracalcular:
    If Not existeix("c:\temp\requeriment") Then MkDir "c:\temp\requeriment"
    vnomfitxer = "c:\temp\requeriment\" + v + " A22.csv"
    If Not vAnnexarCSV Then
         Open vnomfitxer For Output As #1
          Else: Open vnomfitxer For Append As #1
    End If
    If Not vAnnexarCSV Then Print #1, "Nº FACTURA DE ADQUISICIÓN (1);FECHA DE LA ADQUISICIÓN (2);ORIGEN DE LA ADQUISICIÓN (3);REFERENCIA DUA DE IMPORTACIÓN (4);PROVEEDOR (5);DOMICILIO DEL PROVEEDOR (6);PRODUCTO (7);CANTIDAD (8);KG NO RECICLADO (9);IEEPNR  (10);MOTIVO DE LA SOLICITUD DE DEVOLUCIÓN (11);Nº FACTURA VENTA (12);FECHA (13);CLIENTE (14);PRODUCTO (15);CANTIDAD (16);RESPONSABILIDAD  (17);KG NO RECICLADO (18);IEEPNR  (19)"
    Set rst2 = dbtmp.OpenRecordset("select * from Remeses_Taula_Impost_ImpIEsp where kilosnoreciclats>0 and concepte=3 and numremesa=" + atrim(cadbl(v)) + " order by lotinplacsa desc")
    rst2.MoveLast
    rst2.MoveFirst
    While Not rst2.EOF
        DoEvents
        vLiniaFeta = False
        vKgSumatsID = 0
        vnumc = 0
        Me.Caption = atrim(rst2.AbsolutePosition) + "/" + atrim(rst2.RecordCount) + "   vFactorB=" + atrim(vfactorKgaSumarA) + "  " + "vFactorG=" + atrim(vfactorKgaSumarD) + " NºRemesa: " + atrim(vNumRemesa)
         If rst2!lotinplacsa > 150000 Then
              Set rst = dbtmp.OpenRecordset("select * from impostenvasos where id=" + atrim(rst2!identificador))
              If rst.EOF Then
                    GoTo proxima ' MsgBox "No hi ha l'IDENTIFICADOR a IMPOSTENVASOS": GoTo proxima
              End If
              vnumc = rst!comanda
                Else
                  Set rst = dbtmp.OpenRecordset("select * from parcials where id=" + atrim(rst2!identificador))
                  If rst.EOF Then MsgBox "No hi ha el parcial a palets amb aquest ID": GoTo proxima
                  vnumc = rst!idpalet
         End If
         
         If vnumc > 150000 Then
          vmtrsGrupImp = metrestotalstipusproveidorXcomanda(vnumc, "Importació", True)
          vmtrsGrupEsp = metrestotalstipusproveidorXcomanda(vnumc, "Espanyol", True)
          vTmtrsTots = metrestotalstipusproveidorXcomanda(vnumc, "A22", True)
          vKgImp = cadbl(rst!kgMERMAimp_mes_esp) '
          vKgEsp = cadbl(rst!kgMERMAespanya) '
          If rst2!clauproducte = "B" Then vKgImp = rst2!kilosnoreciclats
          If rst2!clauproducte = "G" Then vKgEsp = rst2!kilosnoreciclats
          vKgTots = vKgImp + vKgEsp
          'If vTmtrsTots > 0 Then
          '      vKgEsp = (vmtrsGrupEsp * (vKgTots)) / vTmtrsTots
          '      vKgImp = (vmtrsGrupImp * (vKgTots)) / vTmtrsTots
          'End If
            Else
              vKgImp = cadbl(rst2!kilosnoreciclats) '
              vKgEsp = cadbl(rst2!kilosnoreciclats) '
              vKgTots = cadbl(rst2!kilosnoreciclats)
        End If
              
        
              
        'Tipus Importació
        vsqltipusproveidor = " and tipusproveidorIMPOST='Importació'"
        If vnumc > 150000 Then
              If rst2!clauproducte = "B" Then
                   Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina,Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  teimpost=true and (((Parcials.comanda)='" + atrim(vnumc) + "')) " + vsqltipusproveidor)
                    Else: Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina,Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  Parcials.comanda='-1'") 'Per fer EOF
              End If
                Else: Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina,Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  PARCIALS.ID=" + atrim(rst2!identificador) + vsqltipusproveidor)
        End If
        vSumaTKgParcial = 0
        While Not rstTIPUS.EOF And vKgImp > 0
           vnommaterial = ""
           vliniacompra = ""
           vliniaventa = ""
           If vnumc > 150000 Then
               If vmtrsGrupImp > 0 Then
                   vTKgParcial = Redondejar((cadbl(rstTIPUS!metres) * vKgImp) / vmtrsGrupImp, 3)
                   vFactor = IIf(vTKgParcial > vKgcomptarlinies, vfactorKgaSumarA, 0)
               vFactor = 0
                   vTKgParcial = Redondejar(vTKgParcial + vFactor, 3)
                     Else: vTKgParcial = Redondejar(vTKgParcial + vKgImp, 3)
               End If
                  Else:
                    vTKgParcial = vKgImp
                    vFactor = IIf(vTKgParcial > vKgcomptarlinies, vfactorKgaSumarA, 0)
               vFactor = 0
                    vTKgParcial = Redondejar(vTKgParcial + vFactor, 3)
           End If
           If vfactorKgaSumarA <> 0 Then
                 generar_liniacompra rstTIPUS!idpalet, rstTIPUS!idbobina, vliniacompra, vnommaterial, "Importació", vTKgParcial
                 generar_liniamerma rst!id, vTKgParcial, rst2, vliniaventa, "B", vnommaterial
           End If
           If vTKgParcial > vKgcomptarlinies Then
                vcomptadordeliniesA = vcomptadordeliniesA + 1
           End If
           vcontA = vcontA + vTKgParcial
           vSumaTKgParcial = vSumaTKgParcial + vTKgParcial
           Print #1, vliniacompra + ";" + vliniaventa + ";" + atrim(rst2!identificador)
           vKgSumatsID = vKgSumatsID + vTKgParcial
           vLiniaFeta = True
           rstTIPUS.MoveNext
        Wend
 '       If vSumaTKgParcial <> rst2!kilosnoreciclats And rst2!clauproducte = "B" Then MsgBox atrim(vSumaTKgParcial) + " <> " + atrim(rst2!kilosnoreciclats)
        If vSumaTKgParcial > 0 And rst2!lotinplacsa < 150000 Then GoTo proxima
        
        'Tipus Espanyol
        vsqltipusproveidor = " and tipusproveidorIMPOST='Espanyol'"
        'Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina, Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  teimpost=true and (((Parcials.comanda)='" + atrim(vnumc) + "')) " + vsqltipusproveidor)
        If vnumc > 150000 Then
               If rst2!clauproducte = "G" Then
                    Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina,Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  teimpost=true and (((Parcials.comanda)='" + atrim(vnumc) + "')) " + vsqltipusproveidor)
                      Else: Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina,Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE Parcials.comanda='-1'") 'per fer EOF
               End If
                Else:  Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina,Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  PARCIALS.ID=" + atrim(rst2!identificador) + vsqltipusproveidor)
        End If
        vSumaTKgParcial = 0
        While Not rstTIPUS.EOF And vKgEsp > 0
           vnommaterial = ""
           vliniacompra = ""
           vliniaventa = ""
           If vnumc > 150000 Then
               If vmtrsGrupEsp > 0 Then
                   vTKgParcial = Redondejar((cadbl(rstTIPUS!metres) * vKgEsp) / vmtrsGrupEsp, 3)
                   vFactor = IIf(vTKgParcial > vKgcomptarlinies, vfactorKgaSumarD, 0)
                 vFactor = 0
                   vTKgParcial = Redondejar(vTKgParcial + vFactor, 3)
                     Else: vTKgParcial = Redondejar(vKgEsp + vfactorKgaSumarD, 3)
               End If
                Else:
                   vTKgParcial = vKgEsp
                   vFactor = IIf(vTKgParcial > vKgcomptarlinies, vfactorKgaSumarD, 0)
                 vFactor = 0
                   vTKgParcial = Redondejar(vTKgParcial + vFactor, 3)
           End If
           If vfactorKgaSumarD <> 0 Then
                generar_liniacompra rstTIPUS!idpalet, rstTIPUS!idbobina, vliniacompra, vnommaterial, "Espanya", vTKgParcial
                generar_liniamerma rst!id, vTKgParcial, rst2, vliniaventa, "G", vnommaterial
           End If
           If vTKgParcial > vKgcomptarlinies Then
              vcomptadordeliniesD = vcomptadordeliniesD + 1
           End If
           vcontD = vcontD + vTKgParcial
           vSumaTKgParcial = vSumaTKgParcial + vTKgParcial
           If atrim(vliniacompra + ";" + vliniaventa) <> ";" Then
                Print #1, vliniacompra + ";" + vliniaventa + ";" + atrim(rst2!identificador)
                vKgSumatsID = vKgSumatsID + vTKgParcial
                vLiniaFeta = True
           End If
           rstTIPUS.MoveNext
           
        Wend
'        If vSumaTKgParcial <> rst2!kilosnoreciclats Then MsgBox atrim(vSumaTKgParcial) + " <> " + atrim(rst2!kilosnoreciclats)
        'If vcomptadordeliniesD + vcomptadordeliniesA > 200 Then GoTo Salt1
        If Redondejar(vKgSumatsID, 0) <> Redondejar(cadbl(rst2!kilosnoreciclats), 0) Then Stop
        If Not vLiniaFeta Then
               vMermesNoAfegides = vMermesNoAfegides + atrim(rst2!identificador) + ";" + atrim(rst2!clauproducte) + ";" + atrim(rst2!kilosnoreciclats) + vbNewLine
        End If
           
proxima:
        rst2.MoveNext
    Wend
Salt1:
    If vmsgERROR <> "" Then Print #1, vmsgERROR
    If vMermesNoAfegides <> "" Then Print #1, vbNewLine + vbNewLine + vMermesNoAfegides
    Close #1
    If vfactorKgaSumarA = 0 Or vfactorKgaSumarD = 0 Then
         If vcomptadordeliniesA > 0 Then
           vfactorKgaSumarA = ((vvalorA / 0.45) - vcontA) / vcomptadordeliniesA
             Else: vfactorKgaSumarA = 0.000000000000001
         End If
         vcomptadordeliniesA = 0: vcontA = 0
           vfactorKgaSumarD = ((vvalorD / 0.45) - vcontD) / vcomptadordeliniesD
           vcomptadordeliniesD = 0: vcontD = 0
           GoTo tornaracalcular
    End If
    
    If cadbl(vNumRemesa) = 0 Then obrir_document vnomfitxer
    Set rst = Nothing
    Set rst2 = Nothing
End Sub
Function buscar_impost(vfactura As String, vdata As String) As Double
   Dim rstsap As Recordset
   Set rstsap = dbsap.OpenRecordset("select sum([Quantity]) as Timpost, sum([Price]) as Tpreu from Facturesiliniescompres where itemcode='IMP_ENV' AND numfacturaprov='" + atrim(vfactura) + "' and DataFacturaProv=#" + Format(vdata, "mm/dd/yyyy") + "#")
   If cadbl(rstsap!Timpost) = 0 Then
        Set rstsap = dbsap.OpenRecordset("select sum([Quantity]) as Timpost, sum([Price]) as Tpreu from Facturesiliniescompres where itemcode='IMP_ENV' AND numfacturaprov like '*" + atrim(vfactura) + "' and DataFacturaProv=#" + Format(vdata, "mm/dd/yyyy") + "#")
   End If
   If cadbl(rstsap!Timpost) = 0 Then
          buscar_impost = 0
      Else
        If cadbl(rstsap!Timpost) > 1 Then
             buscar_impost = cadbl(rstsap!Timpost)
              Else: buscar_impost = cadbl(rstsap!Tpreu)
        End If
   '     vnumfacturaproveidor = atrim(rstsap!numfacturaprov)
        'vdatafacturaprov = rstsap!datafacturaprov
   End If
End Function
Sub eliminar_fitxers_i_carpetes(vNomDirectori As String)
   If Not existeix(vNomDirectori) Then
       MkDir vNomDirectori
         Else: If hihaalgudins(vNomDirectori) Then Kill vNomDirectori + "\*.*"
   End If
End Sub
Function hihaalgudins(vNomDirectori As String) As Boolean
    Dim vdir As String
    vdir = Dir(vNomDirectori + "\*.*")
    While vdir <> "" And Not hihaalgudins
       vdir = Dir
       If vdir <> "." Or vdir <> ".." Then hihaalgudins = True
    Wend
End Function
Public Function ReplaceVB5(ByVal sText As String, ByVal sFind As String, ByVal sReplace As String) As String
    Dim iPos As Integer
    ' Si la cadena a cercar és buida, retornem el text original per evitar bucles infinits
    If Len(sFind) = 0 Then
        ReplaceVB5 = sText
        Exit Function
    End If

    iPos = InStr(1, sText, sFind, vbTextCompare) ' vbTextCompare ignora majúscules/minúscules
    Do While iPos > 0
        sText = Left(sText, iPos - 1) & sReplace & Mid(sText, iPos + Len(sFind))
        iPos = InStr(iPos + Len(sReplace), sText, sFind, vbTextCompare)
    Loop
    ReplaceVB5 = sText
End Function
Sub RecorrerCarpetes(ByVal oFolder As Folder)
    Dim oSubFolder As Folder
    Dim oFile As File
    Dim sSQL As String
    Dim sRutaRelativa As String

    ' Processar fitxers
    For Each oFile In oFolder.Files
        ' Comprovar extensió PDF (comparant en minúscules)
        If LCase(Right(oFile.Name, 4)) = ".pdf" Then
            
            ' Calculem la ruta relativa manualment
            sRutaRelativa = ReplaceVB5(oFile.ParentFolder.Path, vRutaFacturesPdfSap, "")
            If sRutaRelativa = "" Then sRutaRelativa = "\"
            
            ' Insert amb tractament de cometes per evitar errors SQL
            sSQL = "INSERT INTO LlistaDirectoriFacturesSAP (rutarelativa, datafitxer, nomfitxer) VALUES (" & _
                   "'" & ReplaceVB5(sRutaRelativa, "'", "''") & "', " & _
                   "#" & Format(oFile.DateLastModified, "yyyy-mm-dd hh:mm:ss") & "#, " & _
                   "'" & ReplaceVB5(oFile.Name, "'", "''") & "')"
            
            dbtmp.Execute sSQL
        End If
    Next

    ' Recórrer subcarpetes
    For Each oSubFolder In oFolder.SubFolders
        RecorrerCarpetes oSubFolder
    Next
End Sub
Sub crear_llista_factures_SAP()
    Dim fso As New FileSystemObject
    Dim oFolder As Folder
    

    dbtmp.Execute "DELETE FROM LlistaDirectoriFacturesSAP"
    
    If fso.FolderExists(vRutaFacturesPdfSap) Then
        Set oFolder = fso.GetFolder(vRutaFacturesPdfSap)
        RecorrerCarpetes oFolder
    End If
    

End Sub
Sub guardar_Totals_factures_A22()
    Dim vsql As String
    dbtmp.Execute "delete * from TotalsFacturesProveidorsA22"
    vsql = "INSERT INTO TotalsFacturesProveidorsA22 ( numfactura, datafactura, nomproveidor, kg, kgacumulats ) "
    vsql = vsql + " SELECT [Suma Factures Proveidors Vs Entregats a Hisenda].[Nº FACTURA DE ADQUISICIÓN (1)], [Suma Factures Proveidors Vs Entregats a Hisenda].[FECHA DE LA ADQUISICIÓN (2)], [Suma Factures Proveidors Vs Entregats a Hisenda].[PrimeroDePROVEEDOR (5)], [Suma Factures Proveidors Vs Entregats a Hisenda].[PrimeroDeKG NO RECICLADO (9)], [Suma Factures Proveidors Vs Entregats a Hisenda].[SumaDeKG NO RECICLADO (18)]"
    vsql = vsql + " FROM [Suma Factures Proveidors Vs Entregats a Hisenda];"
    dbtmp.Execute vsql
End Sub
Private Sub mgenerarequeriment_Click()
    Dim vliniacompra As String
    Dim vliniaventa As String
    Dim vnomfitxer As String
    Dim vnumalbaraprov As String
    Dim vTmtrsTots As Double
    Dim vmtrsGrupImp As Double
    Dim vmtrsGrupEsp As Double
    Dim vKgImp As Double
    Dim vKgEsp As Double
    Dim vTKgParcial As Double
    Dim rst2 As Recordset
    Dim rst As Recordset
    Dim rstp As Recordset
    Dim rstTIPUS As Recordset
    Dim vsqltipusproveidor As String
    Dim vcont As Double
    Dim vnumc As Double
    Dim vKgTots As Double
    Dim vcomptadordelinies As Double
    Dim vfactorKgaSumar As Double
    Dim v As String
    Dim vnomfitxerLOG As String
    vNomFitxerControlFacturesCompra = "c:\temp\requeriment\LogFacturesCompres.csv"
    vCopiarFacturesCompraiVenta = True
    vRutaFacturesPdfSap = "\\SERVIDORSAP\Factures_Compres_Vendes\"
    vRutaFacturesCompres = "c:\temp\requeriment\Factures_Compres\"
    vRutaFacturesVendes = "c:\temp\requeriment\Factures_Vendes\"
    If UCase(InputBox("Vols tornar a crear la llista de factures PDF de SAP?" + vbNewLine + "ESCRIU [SI]", "Atenció")) = "SI" Then crear_llista_factures_SAP
    vnomfitxerLOG = "c:\temp\requeriment\LogCopiaFactures.csv"
    If existeix(vnomfitxerLOG) Then Kill vnomfitxerLOG
    If existeix(vNomFitxerControlFacturesCompra) Then Kill vNomFitxerControlFacturesCompra
    If UCase(InputBox("Si vols generar la taula de totals de factures escriu [Generar]:" + vbNewLine + "AIXO ES FA UN COP NOMÉS ABANS DE COMENÇAR A CALCULAR EL TRIMESTRE SINÓ NO ACUMULARÀ ELS TOTALS ENTRE CALCULS SEGUENTS DINS DEL MATEIX TRIMESTRE.", "TOTALS DE FACTURA ACUMULATS")) = "GENERAR" Then guardar_Totals_factures_A22
    
    eliminar_fitxers_i_carpetes "c:\temp\requeriment\CmrsiFacturesdetransport"
    eliminar_fitxers_i_carpetes vRutaFacturesCompres
    eliminar_fitxers_i_carpetes vRutaFacturesVendes
    If vCopiarFacturesCompraiVenta Then
          Open vnomfitxerLOG For Output As #4
          Print #4, "Llistat de fitxer factures PDF no trobats a " + vRutaFacturesPdfSap
    End If
       'Valor entrat en Kg
    'Generar_requeriment_Vendes , , , , "A"
    Generar_requeriment_Vendes , , , , "E"
    'Generar_requeriment_Mermes
    
    
    If vCopiarFacturesCompraiVenta Then
        Close #4
        wait 1
        If existeix(vnomfitxerLOG) Then obrir_document vnomfitxerLOG
    End If
    If existeix(vNomFitxerControlFacturesCompra) Then obrir_document vNomFitxerControlFacturesCompra
    MsgBox "Fitxer creats a c:\temp\requeriment"
    Exit Sub
    
    
    
    
    
    
    
    v = escullir_historic_Imp_i_Esp(v)
    Set dbtmpb = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
    Set dbsap = OpenDatabase(rutadelfitxer(cami) + "connexiosap.mdb")
tornaracalcular:
    vnomfitxer = "c:\temp\A22_1r_T.csv"
    Open vnomfitxer For Output As #1
    Print #1, "FECHA DE COMPRA/DUA;NºFRA DE COMPRA/DUA;KGS DE PLASTICO NO RECICLADO INTEGRADOS EN LA FRA DE COMPRA;FECHA DE VENTA;NºFRA DE VENTA;KGS DE PLASTICO NO RECICLADO EN LA FRA DE VENTA;TIPO IMPOSITIVO DEL IMPUESTO;IMPORTE DEL IMPUESTO ESPECIAL SOBRE ENVASES DE PLASTICO NO REUTILIZABLE"
  
    Set rst2 = dbtmp.OpenRecordset("select * from Remeses_Taula_Impost_ImpIEsp where concepte=2 and numremesa=" + atrim(cadbl(v)) + " order by clauproducte")
    'Set rst2 = dbtmp.OpenRecordset("select * from Remeses_Taula_Impost_ImpIEsp where concepte=2 and (clauproducte='A' or clauproducte='D') order by data,clauproducte")
    rst2.MoveLast
    rst2.MoveFirst
    While Not rst2.EOF
        DoEvents
        Me.Caption = atrim(rst2.AbsolutePosition) + "/" + atrim(rst2.RecordCount)
         Set rst = dbtmp.OpenRecordset("select * from impostenvasos where id=" + atrim(rst2!identificador))
         If rst.EOF Then MsgBox "No hi ha l'IDENTIFICADOR a IMPOSTENVASOS": GoTo proxima
          vnumc = rst!comanda
          vmtrsGrupImp = metrestotalstipusproveidorXcomanda(vnumc, "Importació")
          vmtrsGrupEsp = metrestotalstipusproveidorXcomanda(vnumc, "Espanyol")
          'vTmtrsTots = metrestotalstipusproveidorXcomanda(vnumc, "")
          vKgImp = cadbl(rst!kgventaimp_mes_esp) '
          vKgEsp = cadbl(rst!kgventaespanya) '
          vKgTots = vKgImp + vKgEsp
          'If vTmtrsTots > 0 Then
          '      vKgEsp = (vmtrsGrupEsp * (vKgTots)) / vTmtrsTots
          '      vKgImp = (vmtrsGrupImp * (vKgTots)) / vTmtrsTots
          'End If
              
        'Tipus Importació
       If rst2!clauproducte = "A" Then
        vsqltipusproveidor = " and tipusproveidorIMPOST='Importació'"
        Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina,Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  teimpost=true and (((Parcials.comanda)='" + atrim(vnumc) + "') AND ((Parcials.orcomassignacio)<>'500'))" + vsqltipusproveidor)
        While Not rstTIPUS.EOF
           vTKgParcial = Redondejar((cadbl(rstTIPUS!metres) * vKgImp) / vmtrsGrupImp, 3)
           generar_liniacompra2 rstTIPUS!idpalet, rstTIPUS!idbobina, vliniacompra
           generar_liniaventa2 rst!id, vTKgParcial, rst2, vliniaventa, "A", ""
        '   If InStr(1, vliniaventa, "23000435;") > 0 Then GoTo proxima
         '  If InStr(1, vliniacompra, ";43000219;") = 0 And InStr(1, vliniaventa, "23000435;") = 0 Then
         '       vTKgparcial = Redondejar(vTKgparcial + vfactorKgaSumar, 3)
         '       vcomptadordelinies = vcomptadordelinies + 1
         '       Else
          '         If InStr(1, vliniacompra, ";43000219;") > 0 Then vTKgparcial = vTKgparcial - 4.5
          ' End If
           vcont = vcont + vTKgParcial
           generar_liniaventa2 rst!id, vTKgParcial, rst2, vliniaventa, "A", ""
           Print #1, vliniacompra + ";" + vliniaventa + ";" + atrim(rst2!clauproducte)
proxima:
           rstTIPUS.MoveNext
        Wend
       End If
'        If vcont > 0 Then MsgBox "Importació " + atrim(vcont) + " --> " + atrim(vKgImp)

        
        'Tipus Espanyol
       If rst2!clauproducte <> "D" And vKgEsp > 0 Then
        vsqltipusproveidor = " and tipusproveidorIMPOST='Espanyol'"
        Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina, Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  teimpost=true and (((Parcials.comanda)='" + atrim(vnumc) + "') AND ((Parcials.orcomassignacio)<>'500'))" + vsqltipusproveidor)
        'Clipboard.Clear
        'Clipboard.SetText "SELECT Parcials.idpalet, parcials.idbobina, Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  teimpost=true and (((Parcials.comanda)='" + atrim(vnumc) + "') AND ((Parcials.orcomassignacio)<>'500'))" + vsqltipusproveidor
        While Not rstTIPUS.EOF
           vTKgParcial = Redondejar((cadbl(rstTIPUS!metres) * vKgEsp) / vmtrsGrupEsp, 3)
           generar_liniacompra2 rstTIPUS!idpalet, rstTIPUS!idbobina, vliniacompra
           generar_liniaventa2 rst!id, vTKgParcial, rst2, vliniaventa, "D", ""
          ' If InStr(1, vliniaventa, "23000435;") > 0 Then GoTo proxima2
          ' If InStr(1, vliniacompra, ";43000219;") = 0 And InStr(1, vliniaventa, "23000435;") = 0 Then
          '      vTKgparcial = Redondejar(vTKgparcial + vfactorKgaSumar, 3)
          '      vcomptadordelinies = vcomptadordelinies + 1
           '       Else
           '         If InStr(1, vliniacompra, ";43000219;") > 0 Then vTKgparcial = vTKgparcial - 4.5
           'End If
           vcont = vcont + vTKgParcial
           generar_liniaventa2 rst!id, vTKgParcial, rst2, vliniaventa, "D", ""
           Print #1, vliniacompra + ";" + vliniaventa + ";" + atrim(rst2!clauproducte)
           rstTIPUS.MoveNext
        Wend
      End If
proxima2:
      rst2.MoveNext
    Wend
    
 GoTo fi   'salto al final
 
merma:
    Set rst2 = dbtmp.OpenRecordset("select * from Remeses_Taula_Impost_ImpIEsp where concepte=3 and lotinplacsa>200000 and numremesa=" + atrim(cadbl(v)) + " order by clauproducte")
    rst2.MoveLast
    rst2.MoveFirst
    While Not rst2.EOF
        DoEvents
        Me.Caption = atrim(rst2.AbsolutePosition) + "/" + atrim(rst2.RecordCount)
         Set rst = dbtmp.OpenRecordset("select * from impostenvasos where id=" + atrim(rst2!identificador))
         If rst.EOF Then MsgBox "No hi ha l'IDENTIFICADOR a IMPOSTENVASOS": GoTo proxima3
          vnumc = rst!comanda
          vmtrsGrupImp = metrestotalstipusproveidorXcomanda(vnumc, "Importació")
          vmtrsGrupEsp = metrestotalstipusproveidorXcomanda(vnumc, "Espanyol")
          'vTmtrsTots = metrestotalstipusproveidorXcomanda(vnumc, "")
          vKgImp = cadbl(rst!kgventaimp_mes_esp) '
          vKgEsp = cadbl(rst!kgventaespanya) '
          vKgTots = vKgImp + vKgEsp
          'If vTmtrsTots > 0 Then
          '      vKgEsp = (vmtrsGrupEsp * (vKgTots)) / vTmtrsTots
          '      vKgImp = (vmtrsGrupImp * (vKgTots)) / vTmtrsTots
          'End If
              
        'Tipus Importació
       If vmtrsGrupImp > 0 Then
        vsqltipusproveidor = " and tipusproveidorIMPOST='Importació'"
        Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina,Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  teimpost=true and (((Parcials.comanda)='" + atrim(vnumc) + "'))" + vsqltipusproveidor)
        If Not rstTIPUS.EOF And rst2!kilos > 0 Then
           vTKgParcial = rst2!kilos
            generar_liniacompra2 rstTIPUS!idpalet, rstTIPUS!idbobina, vliniacompra
           generar_liniaventa2 rst!id, vTKgParcial, rst2, vliniaventa, "A", ""
           If InStr(1, vliniaventa, "23000435;") > 0 Then Stop: GoTo proxima3
           If InStr(1, vliniacompra, ";43000219;") = 0 And InStr(1, vliniaventa, "23000435;") = 0 Then
                vTKgParcial = Redondejar(vTKgParcial + vfactorKgaSumar, 3)
                vcomptadordelinies = vcomptadordelinies + 1
                Else
                   If InStr(1, vliniacompra, ";43000219;") > 0 Then vTKgParcial = vTKgParcial - 4.5
                   ' If InStr(1, vliniaventa, "23000435;") > 0 Then Stop
           End If
           vcont = vcont + vTKgParcial
           generar_liniaventa2 rst!id, vTKgParcial, rst2, vliniaventa, "A", ""
           Print #1, vliniacompra + ";" + vliniaventa + ";" + atrim(rst2!clauproducte)
proxima3:
       End If
      End If
      
      If vmtrsGrupEsp > 0 Then
        vsqltipusproveidor = " and tipusproveidorIMPOST='Espanyol'"
        Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina,Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  teimpost=true and (((Parcials.comanda)='" + atrim(vnumc) + "'))" + vsqltipusproveidor)
        If Not rstTIPUS.EOF And rst2!kilos > 0 Then
           vTKgParcial = rst2!kilos
            generar_liniacompra2 rstTIPUS!idpalet, rstTIPUS!idbobina, vliniacompra
           generar_liniaventa2 rst!id, vTKgParcial, rst2, vliniaventa, "D", ""
           If InStr(1, vliniaventa, "23000435;") > 0 Then Stop: GoTo proxima4
           If InStr(1, vliniacompra, ";43000219;") = 0 And InStr(1, vliniaventa, "23000435;") = 0 Then
                vTKgParcial = Redondejar(vTKgParcial + vfactorKgaSumar, 3)
                vcomptadordelinies = vcomptadordelinies + 1
                Else
                   If InStr(1, vliniacompra, ";43000219;") > 0 Then vTKgParcial = vTKgParcial - 4.5
                   ' If InStr(1, vliniaventa, "23000435;") > 0 Then Stop
           End If
           vcont = vcont + vTKgParcial
           generar_liniaventa2 rst!id, vTKgParcial, rst2, vliniaventa, "D", ""
           Print #1, vliniacompra + ";" + vliniaventa + ";" + atrim(rst2!clauproducte)
proxima4:
        End If
      End If
      rst2.MoveNext
     Wend
  
mermabobines:
  'MERMA DE BOBINES SUELTES
      
    Set rst2 = dbtmp.OpenRecordset("select * from Remeses_Taula_Impost_ImpIEsp where concepte=3 and lotinplacsa<200000 and numremesa=" + atrim(cadbl(v)) + " order by clauproducte")
    rst2.MoveLast
    rst2.MoveFirst
    While Not rst2.EOF
        DoEvents
        Me.Caption = atrim(rst2.AbsolutePosition) + "/" + atrim(rst2.RecordCount)
        Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina,Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  Parcials.id=" + atrim(rst2!identificador))
        If Not rstTIPUS.EOF Then
           vTKgParcial = rst2!kilos
           generar_liniacompra2 rstTIPUS!idpalet, rstTIPUS!idbobina, vliniacompra
           generar_liniaventa2 1, vTKgParcial, rst2, vliniaventa, "A", ""
           If InStr(1, vliniaventa, "23000435;") > 0 Then Stop: GoTo proxima5
           If InStr(1, vliniacompra, ";43000219;") = 0 And InStr(1, vliniaventa, "23000435;") = 0 Then
                vTKgParcial = Redondejar(vTKgParcial + vfactorKgaSumar, 3)
                vcomptadordelinies = vcomptadordelinies + 1
                Else
                   If InStr(1, vliniacompra, ";43000219;") > 0 Then vTKgParcial = vTKgParcial - 4.5
                   ' If InStr(1, vliniaventa, "23000435;") > 0 Then Stop
           End If
           vcont = vcont + vTKgParcial
           generar_liniaventa2 1, vTKgParcial, rst2, vliniaventa, "D", ""
           Print #1, vliniacompra + ";" + vliniaventa + ";" + atrim(rst2!clauproducte)
proxima5:
       End If
      
      rst2.MoveNext
    Wend
      
      
fi:
      
    Close #1
    'If vfactorKgaSumar = 0 Then
    '       vfactorKgaSumar = ((36622.515 / 0.45) - vcont) / vcomptadordelinies
    '       vcomptadordelinies = 0: vcont = 0
    '       GoTo tornaracalcular
    'End If
    
    obrir_document vnomfitxer
    Set rst = Nothing
    Set rst2 = Nothing
End Sub
Private Sub Generar_requeriment_Vendes(Optional vNumRemesa As Double, Optional vAnnexarCSV As Boolean, Optional vvalorA As Double, Optional vvalorD As Double, Optional vLletraFiscal As String)
    Dim vliniacompra As String
    Dim vliniaventa As String
    Dim vnomfitxer As String
    Dim vnumalbaraprov As String
    Dim vTmtrsTots As Double
    Dim vmtrsGrupImp As Double
    Dim vmtrsGrupEsp As Double
    Dim vKgImp As Double
    Dim vKgEsp As Double
    Dim vTKgParcial As Double
    Dim rst2 As Recordset
    Dim rst As Recordset
    Dim rstTIPUS As Recordset
    Dim vsqltipusproveidor As String
    Dim vcontA As Double
    Dim vcontD As Double
    Dim vnumc As Double
    Dim vKgTots As Double
    Dim vcomptadordeliniesA As Double
    Dim vfactorKgaSumarA As Double
    Dim vcomptadordeliniesD As Double
    Dim vfactorKgaSumarD As Double
    Dim vsumaImp As Double
    Dim vsumaEsp As Double
    Dim v As String
    Dim vnommaterial As String
    
    
    vvalorA = 0.00000000001  'Es perquè no demani valors
    
    If cadbl(vNumRemesa) = 0 Then
         v = escullir_historic_Imp_i_Esp(v)
         vNumRemesa = cadbl(v)
          Else: v = vNumRemesa
    End If
    If cadbl(vvalorA) = 0 Then
            vvalorA = cadbl(InputBox("Entra el valor de la casella A. En Kg", "Valor A"))
            If vvalorA = 0 Then MsgBox "No pot ser zero.", vbCritical, "Error": Exit Sub
            vvalorD = cadbl(InputBox("Entra el valor de la casella D. En Kg", "Valor D"))
            If vvalorD = 0 Then MsgBox "No pot ser zero.", vbCritical, "Error": Exit Sub
    End If
      'Si es vol fer en Euros enlloc de Kg treure les 2 seguents linies
    vvalorA = vvalorA * 0.45
    vvalorD = vvalorD * 0.45
    
    'vLletraFiscal = "E"
    
    If atrim(vLletraFiscal) = "" Then vLletraFiscal = "A"   'Si no passo lletra suposo que es Venda fora d'Espanya
    
    Set dbtmpb = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
    Set dbsap = OpenDatabase(rutadelfitxer(cami) + "connexiosap.mdb")
   vfactorKgaSumarD = 0.00000000000001
    vfactorKgaSumarA = 0.00000000000001
tornaracalcular:
    If Not existeix("c:\temp\requeriment") Then MkDir "c:\temp\requeriment"
    vnomfitxer = "c:\temp\requeriment\" + v + " A22.csv"
    If Not vAnnexarCSV Then
         Open vnomfitxer For Output As #1
          Else: Open vnomfitxer For Append As #1
    End If
    If Not vAnnexarCSV Then Print #1, "Nº FACTURA DE ADQUISICIÓN (1);FECHA DE LA ADQUISICIÓN (2);ORIGEN DE LA ADQUISICIÓN (3);REFERENCIA DUA DE IMPORTACIÓN (4);PROVEEDOR (5);DOMICILIO DEL PROVEEDOR (6);PRODUCTO (7);CANTIDAD (8);KG NO RECICLADO (9);IEEPNR  (10);MOTIVO DE LA SOLICITUD DE DEVOLUCIÓN (11);Nº FACTURA VENTA (12);FECHA (13);CLIENTE (14);PRODUCTO (15);CANTIDAD (16);RESPONSABILIDAD  (17);KG NO RECICLADO (18);IEEPNR  (19);INCOTERM;CMR;Factura Transportista"
    Set rst2 = dbtmp.OpenRecordset("select * from Remeses_Taula_Impost_ImpIEsp where concepte=2 and regimfiscal='" + vLletraFiscal + "' and numremesa=" + atrim(cadbl(v)))
    rst2.MoveLast
    rst2.MoveFirst
    While Not rst2.EOF
        DoEvents
        vsumaImp = 0: vsumaEsp = 0
        Me.Caption = atrim(rst2.AbsolutePosition) + "/" + atrim(rst2.RecordCount) + " vFactorA=" + atrim(vfactorKgaSumarA) + "  " + "vFactorD=" + atrim(vfactorKgaSumarD) + " NºRemesa: " + atrim(vNumRemesa)
         Set rst = dbtmp.OpenRecordset("select * from impostenvasos where id=" + atrim(rst2!identificador))
         If rst.EOF Then GoTo proxima ' MsgBox "No hi ha l'IDENTIFICADOR a IMPOSTENVASOS": GoTo proxima
          vnumc = rst!comanda
       '   If InStr(1, vnumerosdelotsfets, atrim(vnumc)) > 0 Then GoTo proxima
          vmtrsGrupImp = metrestotalstipusproveidorXcomanda(vnumc, "Importació")
          vmtrsGrupEsp = metrestotalstipusproveidorXcomanda(vnumc, "Espanyol")
          vTmtrsTots = metrestotalstipusproveidorXcomanda(vnumc, "A22")
          vKgImp = cadbl(rst!kgventaimp_mes_esp) '
          vKgEsp = cadbl(rst!kgventaespanya) '
          vKgTots = vKgImp + vKgEsp
          If vTmtrsTots > 0 Then
                vKgEsp = (vmtrsGrupEsp * (vKgTots)) / vTmtrsTots
                vKgImp = (vmtrsGrupImp * (vKgTots)) / vTmtrsTots
          End If
         ' If Redondejar(vKgEsp, 0) <> Redondejar(cadbl(rst!kgventaespanya), 0) Then Stop
          'If Redondejar(vKgImp, 0) <> Redondejar(cadbl(rst!kgventaimp_mes_esp), 0) Then Stop
              
        'Tipus Importació
        If rst2!clauproducte <> "A" Then GoTo Imp_Espanyol
        vsqltipusproveidor = " and tipusproveidorIMPOST='Importació'"
        Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina,Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  teimpost=true and (((Parcials.comanda)='" + atrim(vnumc) + "') AND ((Parcials.orcomassignacio)<>'500'))" + vsqltipusproveidor)
        While Not rstTIPUS.EOF
           vnommaterial = ""
           vliniacompra = ""
           vliniaventa = ""
           vTKgParcial = Redondejar((cadbl(rstTIPUS!metres) * vKgImp) / vmtrsGrupImp, 3)
           vTKgParcial = Redondejar(vTKgParcial + IIf(vTKgParcial > 30, vfactorKgaSumarA, 0), 3)
           If vfactorKgaSumarA <> 0 Then
                 generar_liniacompra rstTIPUS!idpalet, rstTIPUS!idbobina, vliniacompra, vnommaterial, "Importació", vTKgParcial
                 'generar_liniaventa rst!id, vTKgparcial, rst2, vliniaventa
                 'If InStr(1, vliniaventa, "23000435;") > 0 Then GoTo proxima
                 'If InStr(1, vliniacompra, ";43000219;") = 0 And InStr(1, vliniaventa, "23000435;") = 0 Then
                      'vTKgparcial = Redondejar(vTKgparcial + vfactorKgaSumar, 3)
                      'vcomptadordeliniesA = vcomptadordeliniesA + 1
                '      Else
                '         If InStr(1, vliniacompra, ";43000219;") > 0 Then vTKgparcial = vTKgparcial - 4.5
                ' End If
                generar_liniaventa2 rst!id, vTKgParcial, rst2, vliniaventa, rst2!clauproducte, vnommaterial  'vLletraFiscal
          End If
          If vTKgParcial > 10 Then vcomptadordeliniesA = vcomptadordeliniesA + 1
          vcontA = vcontA + vTKgParcial
          vsumaImp = vsumaImp + vTKgParcial
           Print #1, vliniacompra + ";" + vliniaventa
           rstTIPUS.MoveNext
        Wend
'        If vcont > 0 Then MsgBox "Importació " + atrim(vcont) + " --> " + atrim(vKgImp)
       If Redondejar(vKgImp, 0) <> Redondejar(vsumaImp, 0) Then Stop  'SEGURAMENT ES COMPRA EXTERNA(REVENTA) S'HA D'ENTRAR MANUALMENT
       
Imp_Espanyol:
        'Tipus Espanyol
        If rst2!clauproducte <> "D" Then GoTo proxima
        If rst2!identificador = 0 Then Stop
        vsqltipusproveidor = " and tipusproveidorIMPOST='Espanyol'"
        Set rstTIPUS = dbtmp.OpenRecordset("SELECT Parcials.idpalet, parcials.idbobina, Parcials.comanda, Parcials.orcomassignacio, Parcials.metres, proveidors.tipusproveidorIMPOST FROM ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE  teimpost=true and (((Parcials.comanda)='" + atrim(vnumc) + "') AND ((Parcials.orcomassignacio)<>'500'))" + vsqltipusproveidor)
        While Not rstTIPUS.EOF
           vnommaterial = ""
           vliniacompra = ""
           vliniaventa = ""
           vTKgParcial = Redondejar((cadbl(rstTIPUS!metres) * vKgEsp) / vmtrsGrupEsp, 3)
           vTKgParcial = Redondejar(vTKgParcial + IIf(vTKgParcial > 30, vfactorKgaSumarD, 0), 3)
           If vfactorKgaSumarD <> 0 Then
                generar_liniacompra rstTIPUS!idpalet, rstTIPUS!idbobina, vliniacompra, vnommaterial, "Espanya", vTKgParcial
                'generar_liniaventa rst!id, vTKgparcial, rst2, vliniaventa
                'If InStr(1, vliniaventa, "23000435;") > 0 Then GoTo proxima
                'If InStr(1, vliniacompra, ";43000219;") = 0 And InStr(1, vliniaventa, "23000435;") = 0 Then
                     'vTKgparcial = Redondejar(vTKgparcial + vfactorKgaSumar, 3)
                     'vcomptadordeliniesD = vcomptadordeliniesD + 1
                '       Else
                '         If InStr(1, vliniacompra, ";43000219;") > 0 Then vTKgparcial = vTKgparcial - 4.5
                'End If
                generar_liniaventa2 rst!id, vTKgParcial, rst2, vliniaventa, rst2!clauproducte, vnommaterial  'vLletraFiscal
           End If
           If vTKgParcial > 10 Then vcomptadordeliniesD = vcomptadordeliniesD + 1
           vcontD = vcontD + vTKgParcial
           vsumaEsp = vsumaEsp + vTKgParcial
           If atrim(vliniacompra + ";" + vliniaventa) <> ";" Then Print #1, vliniacompra + ";" + vliniaventa
           rstTIPUS.MoveNext
        Wend
        If Redondejar(cadbl(vKgEsp), 0) <> Redondejar(cadbl(vsumaEsp), 0) Then Stop 'SEGURAMENT ES COMPRA EXTERNA(REVENTA) S'HA D'ENTRAR MANUALMENT
proxima:
       ' If rst2!clauproducte = "D" And Redondejar(rst2!kilosnoreciclats, 0) <> Redondejar(vsumaEsp, 0) Then Stop
       ' If rst2!clauproducte = "A" And Redondejar(rst2!kilosnoreciclats, 0) <> Redondejar(vsumaImp, 0) Then Stop
       vnumerosdelotsfets = vnumerosdelotsfets + " " + atrim(vnumc)
       rst2.MoveNext
    Wend
    Close #1
    If vfactorKgaSumarA = 0 Or vfactorKgaSumarD = 0 Then
           If vcomptadordeliniesA > 0 Then vfactorKgaSumarA = ((vvalorA / 0.45) - vcontA) / vcomptadordeliniesA
           vcomptadordeliniesA = 0: vcontA = 0
           If vcomptadordeliniesD > 0 Then vfactorKgaSumarD = ((vvalorD / 0.45) - vcontD) / vcomptadordeliniesD
           vcomptadordeliniesD = 0: vcontD = 0
           GoTo tornaracalcular
    End If
    
    If cadbl(vNumRemesa) = 0 Then obrir_document vnomfitxer
    Set rst = Nothing
    Set rst2 = Nothing
End Sub
Sub generarrequerimentPRIMERINTENT_NOVALID()
Dim rst As Recordset
   Dim vmtrsimp As Double
   Dim vmtrsesp As Double
   Dim vtanximp As Double
   Dim vcomptador As Double
   Dim vtanxesp As Double
   Dim rstcompra As Recordset
   Dim rstventa As Recordset
   Dim rstimpenv As Recordset
   Dim rstsap As Recordset
   Dim rstalb As Recordset
   Dim vtotal As Double
   Dim vnumc As Double
   Dim vfactcompra As String
   Dim vdatacompra As Date
   Dim vfactventa As String
   Dim vdataventa As Date
   Dim vliniacompresI As String
   Dim vliniacompresE As String
   Dim vliniaventes As String
   Dim vnomfitxer As String
   Dim dbsap As Database
   
   Set dbsap = OpenDatabase(rutadelfitxer(cami) + "connexiosap.mdb")
   vnomfitxer = "c:\temp\A22_1r_T.csv"
   Set rst = dbtmp.OpenRecordset("select * from Remeses_Taula_Impost_ImpIEsp where concepte=2") 'paisventa<>'ES' and month(data)<4 and Imp_mes_Esp_KgIMPOST>0")
   rst.MoveLast
   rst.MoveFirst
   Open vnomfitxer For Output As #1
   Print #1, "IDENTIF.PROVEEDOR;NOMBRE PROVEEDOR;DUA IMPORTACION/NUMERO FACTURA;FECHA DUA O FECHA FACTURA;IMPUESTO PAGADO DUA O SOPORTADO EN FACTURA;MERCANCIA;CANTIDAD MERCANCIA ADQUIRIDA(KG);NºFACTURA SALIDA TAI O DUA EXPORTACION;FECHA FACTURA SALIDA TAI O FECHA DUA EXPORTACION;IDENTIFICACION CLIENTE;NOMBRE CLIENTE;CANTIDAD MERCANCIA VENDIDA;IMPUESTO SOLICITADO A DEVOLVER"
   While Not rst.EOF
       Me.Caption = atrim(rst!lotinplacsa) + "   " + atrim(rst.AbsolutePosition) + "/" + atrim(rst.RecordCount) + "->" + atrim(vcomptador + 1): DoEvents
       vnumc = rst!lotinplacsa
'       Clipboard.Clear
'       Clipboard.SetText "SELECT Parcials.comanda, Palets.Idpalet, Parcials.metres, proveidors.tipusproveidorIMPOST, proveidors.nom FROM (materials RIGHT JOIN (Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) ON materials.codi = Palets.codimatprognou) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE (((Parcials.comanda)='" + atrim(vnumc) + "')  AND ((Parcials.orcomassignacio)<>'500'));"
      Set rstcompra = dbtmp.OpenRecordset("SELECT palets.Numalb,Parcials.comanda, Palets.Idpalet, Parcials.metres, proveidors.tipusproveidorIMPOST, proveidors.nom,albaransbip.KgImpostEnvasos, albaransbip.codiproveidorcomercial, albaransbip.nomproveidorcomercial, albaransbip.numalbaraprov, Palets.teimpost, proveidors_codisSAP.Nif FROM (((materials RIGHT JOIN (Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) ON materials.codi = Palets.codimatprognou) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi) LEFT JOIN albaransbip ON Palets.Idpalet = albaransbip.numpalet) LEFT JOIN proveidors_codisSAP ON albaransbip.codiproveidorcomercial = proveidors_codisSAP.codiSAP WHERE (((Parcials.comanda)='" + atrim(vnumc) + "') AND ((Parcials.orcomassignacio)<>'500') AND ((Palets.teimpost)=True));")
      '"SELECT Parcials.comanda, Palets.Idpalet, Parcials.metres, proveidors.tipusproveidorIMPOST, proveidors.nom FROM (materials RIGHT JOIN (Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) ON materials.codi = Palets.codimatprognou) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE (((Parcials.comanda)='" + atrim(vnumc) + "')  AND ((Parcials.orcomassignacio)<>'500'));")
      vmtrsesp = 0
      vmtrsimp = 0
      vliniacompresI = ""
      vliniacompresE = ""
      vliniaventes = ""
      vfactcompra = ""
      vdatacompra = 0
      If Not rstcompra.EOF Then
          rstcompra.FindFirst "tipusproveidorimpost='Espanyol'"
          vkgimpostcompra = 0
          If Not rstcompra.NoMatch Then
          Set rstalb = dbtmp.OpenRecordset("SELECT albaransbip.*, proveidors_codisSAP.Nif FROM albaransbip LEFT JOIN proveidors_codisSAP ON albaransbip.codiproveidorcomercial = proveidors_codisSAP.codiSAP WHERE (((albaransbip.numpalet)=" + atrim(IIf(atrim(rstcompra!numalbaraprov) <> "", rstcompra!idpalet, cadbl(rstcompra!numalbaraprov))) + "));")
              If rstalb.EOF Then MsgBox "LINIA ALBARANSBIP NO TROBADA"
              Set rstsap = dbtmp.OpenRecordset("select numfacturasap,facturaprov,datafactura,DUA,DataDUA from Importada_Albarans_Compres_Inplacsa where NumAtCard='" + atrim(rstalb!numalbaraprov) + "' order by numfacturasap desc")
              If Not rstsap.EOF Then
                If atrim(rstsap!facturaprov) = "" Then
                       If rstsap.EOF Then MsgBox "Sense NªFactura a l'albara de prov: " + rstcompra!numalbaraprov: End
                       MsgBox "Numero de factura no trobat albara de proveidor " + atrim(rstalb!numalbaraprov)
                End If
                vfactcompra = atrim(rstsap!facturaprov)
                If IsDate(rstsap!datafactura) Then vdatacompra = rstsap!datafactura
                Set rstimpenv = dbsap.OpenRecordset("select * from Facturesiliniescompres where itemcode='IMP_ENV' and docnum=" + atrim(rstsap!numfacturasap))
                vkgimpostcompra = 0
                If atrim(rstsap!DUA) <> "" Then
                      While Not rstimpenv.EOF
                           vkgimpostcompra = vkgimpostcompra + (cadbl(rstimpenv!Price))
                           rstimpenv.MoveNext
                      Wend
                End If
                If atrim(rstsap!DUA) = "" Then
                      While Not rstimpenv.EOF
                           vkgimpostcompra = vkgimpostcompra + (cadbl(rstimpenv!Quantity))
                           rstimpenv.MoveNext
                      Wend
                End If
                 Else: MsgBox "Factura de l'albarà " + atrim(rstalb!numalbaraprov) + " no trobada."
              End If
              If vkgimpostcompra = 0 Then vkgimpostcompra = rstcompra!KgImpostEnvasos
              vliniacompresE = atrim(rstalb!nif) + ";" + atrim(rstalb!nomproveidorcomercial) + ";" + atrim(vfactcompra) + ";" + atrim(vdatacompra) + ";" + atrim(vkgimpostcompra * 0.45) + ";" + atrim(rstalb!descripcio) + ";" + atrim(vkgimpostcompra)
          End If
          While Not rstcompra.NoMatch
              vmtrsesp = vmtrsesp + cadbl(rstcompra!metres)
              rstcompra.FindNext "tipusproveidorimpost='Espanyol'"
              'If Not rstcompra.NoMatch Then MsgBox atrim(vnumc) + " -> Dos proveidors"
              'GoTo proxima
          Wend
          rstcompra.FindFirst "tipusproveidorimpost='Importació'"
          vkgimpostcompra = 0
          If Not rstcompra.NoMatch Then
              Set rstalb = dbtmp.OpenRecordset("SELECT albaransbip.*, proveidors_codisSAP.Nif FROM albaransbip LEFT JOIN proveidors_codisSAP ON albaransbip.codiproveidorcomercial = proveidors_codisSAP.codiSAP WHERE (((albaransbip.numpalet)=" + atrim(IIf(atrim(rstcompra!numalbaraprov) <> "", rstcompra!idpalet, cadbl(rstcompra!numalbaraprov))) + "));")
              If rstalb.EOF Then MsgBox "LINIA ALBARANSBIP NO TROBADA"
              Set rstsap = dbtmp.OpenRecordset("select numfacturasap,facturaprov,datafactura,DUA,DataDUA from Importada_Albarans_Compres_Inplacsa where NumAtCard='" + atrim(rstalb!numalbaraprov) + "'")
              If Not rstsap.EOF Then
                 If atrim(rstsap!facturaprov) = "" Then If rstsap.EOF Then MsgBox "Sense NªFactura a l'albara de prov: " + rstcompra!numalbaraprov: End
                 vfactcompra = IIf(rstsap!DUA <> "", rstsap!DUA, rstsap!facturaprov)
                 vdatacompra = IIf(rstsap!DUA <> "", atrim(rstsap!DataDUA), rstsap!datafactura)
                 Set rstimpenv = dbsap.OpenRecordset("select * from Facturesiliniescompres where itemcode='IMP_ENV' and docnum=" + atrim(rstsap!numfacturasap))
                 If atrim(rstsap!DUA) <> "" Then
                      While Not rstimpenv.EOF
                           vkgimpostcompra = vkgimpostcompra + (cadbl(rstimpenv!Price))
                           rstimpenv.MoveNext
                      Wend
                 End If
                 If atrim(rstsap!DUA) = "" Then
                      While Not rstimpenv.EOF
                           vkgimpostcompra = vkgimpostcompra + (cadbl(rstimpenv!Quantity))
                           rstimpenv.MoveNext
                      Wend
                 End If
              End If
              If vkgimpostcompra = 0 Then vkgimpostcompra = rstcompra!KgImpostEnvasos
              vliniacompresI = atrim(rstalb!nif) + ";" + atrim(rstalb!nomproveidorcomercial) + ";" + atrim(vfactcompra) + ";" + atrim(vdatacompra) + ";" + atrim(vkgimpostcompra * 0.45) + ";" + atrim(rstalb!descripcio) + ";" + atrim(vkgimpostcompra)
          End If
          While Not rstcompra.NoMatch
              vmtrsimp = vmtrsimp + cadbl(rstcompra!metres)
              rstcompra.FindNext "tipusproveidorimpost='Espanyol'"
              'If Not rstcompra.NoMatch Then MsgBox atrim(vnumc) + " -> Dos proveidors"
              'GoTo proxima
          Wend
      End If
proxima:
      If vmtrsimp + vmtrsesp > 0 Then
        vtanximp = (vmtrsimp * 100) / (vmtrsimp + vmtrsesp) / 100
        vtanxesp = (vmtrsesp * 100) / (vmtrsimp + vmtrsesp) / 100
        'MsgBox atrim(vnumc) + "  Imp: " + atrim(rst!kilosnoreciclats * vtanximp) + "  Esp: " + atrim(rst!kilosnoreciclats * vtanxesp)
           Else: MsgBox atrim(vnumc) + " valor de packinglist zero"
      End If
      vtotal = vtotal + (rst!kilosnoreciclats + 1.7771)
      vcomptador = vcomptador + 1
      Set rstventa = dbtmp.OpenRecordset("SELECT * from impostenvasos where id=" + atrim(rst!identificador))
      If rstventa.EOF Then GoTo unaaltra
      Set rstventa = dbtmp.OpenRecordset("select * from liniesalbara where numalbara=" + atrim(rstventa!numalbara) + " and lotinplacsa=" + atrim(numcomandaprincipal(rstventa!comanda)))
      If rstventa.EOF Then MsgBox "No s'ha trobat el registre IDENTIFICADOR a liniesalbarans els KG DE QUANTITAT VENUDA S'AGAFARÀN DE LES LINIES D'ALBARÀ"
      vfactventa = atrim(rst!justificant)
      vdataventa = atrim(rst!Data)
      comprovarsiDUA vfactventa, vdataventa
      vkgbruts = cadbl(rstventa!kgtotalsbruts)
      If vkgbruts = 0 Then Stop
      If vkgbruts = 0 Then
          Set rstcompra = dbtmp.OpenRecordset("select * from liniesalbara where lotinplacsa=" + atrim(rst!lotinplacsa))
          If Not rstcompra.EOF Then vkgbruts = cadbl(rstcompra!kgtotalsbruts)
      End If
      If vtanxesp > 0 Then
              'el 1.7771 es per compesar el total entregat a hisenda
          vliniaventes = atrim(vfactventa) + ";" + atrim(vdataventa) + ";" + atrim(rst!nifdestinatari) + ";" + atrim(rst!nomdestinatari) + ";" + atrim(vkgbruts) + ";" + atrim(Redondejar((rst!kilosnoreciclats + 1.7771) * vtanxesp, 2)) 'el 1.7771 es per compesar el total entregat a hisenda
          Print #1, vliniacompresE + ";"; vliniaventes
      End If
      If vtanximp > 0 Then
             'el 1.7771 es per compesar el total entregat a hisenda
           vliniaventes = atrim(vfactventa) + ";" + atrim(vdataventa) + ";" + atrim(rst!nifdestinatari) + ";" + atrim(rst!nomdestinatari) + ";" + atrim(vkgbruts) + ";" + atrim(Redondejar((rst!kilosnoreciclats + 1.7771) * vtanximp, 2)) 'el 1.7771 es per compesar el total entregat a hisenda
           Print #1, vliniacompresI + ";"; vliniaventes
      End If
unaaltra:
      rst.MoveNext
   Wend
   Close #1
   If existeix(vnomfitxer) Then obrir_document vnomfitxer
   MsgBox "fi  " + atrim(vtotal) + " Kg"
   Set rst = Nothing
End Sub
Sub comprovarsiDUA(vfactventa As String, vdataventa As Date)
   Dim rst As Recordset
   'set rst=dbtmp.OpenRecordset(
   
   Set rst = Nothing
End Sub
Function numcomandaprincipal(vnumc As Double) As Double
  Dim rst As Recordset
  Dim v1 As Double
  Dim v2 As Double
  Dim v3 As Double
   Set rst = dbtmp.OpenRecordset("select linkcomanda1,linkcomanda2,comanda from comandes where comanda=" + atrim(vnumc))
   v1 = rst!comanda
   v2 = rst!linkcomanda1
   v3 = rst!linkcomanda2
   If v1 = 0 Then v1 = 999999
   If v2 = 0 Then v2 = 999999
   If v3 = 0 Then v3 = 999999
   numcomandaprincipal = v1
   If v2 < v1 Then numcomandaprincipal = v2
   If v3 < numcomandaprincipal Then numcomandaprincipal = v3
  Set rst = Nothing
End Function
Private Sub mhisImpEsp_Click()
 Dim v As String
   Dim vtitol As String
   v = escullir_historic_Imp_i_Esp(vtitol)
   If cadbl(v) = 0 Then Exit Sub
   dbtmp.Execute "delete * from taula_impost"
   dbtmp.Execute "insert into taula_impost select * from Remeses_Taula_Impost_ImpIEsp where numremesa=" + v
   dataintracomunitari.RecordSource = "select * from taula_impost order by concepte,data"
   dataintracomunitari.Refresh
   etllistat.Caption = "Historic A22 Importació_i_Espanya   Remesa:" + v + " " + UCase(vtitol)
   'etllistat.Tag = "01/" + Mid(v, 6) + "/" + Mid(v, 2, 4)
   etllistat.Tag = ""
   vtipusllistat = "E+I"
   configurar_reixa
  emplenar_reixa
  'sumar_mermes_i_vendes
  sumar_mermes_i_vendes_ImpIEsp
  bguardarCSV.Enabled = True
End Sub
Function escullir_historic_Imp_i_Esp(vtitol As String) As String
  Load formseleccio
  formseleccio.sortirs.Tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "ImpostEnvasos.mdb"
  formseleccio.Data1.RecordSource = "SELECT Remeses_Taula_Impost_ImpIEsp.numremesa, Format(MAX(data),'mmmm') AS Mes, Year(MAX(data)) AS Any From Remeses_Taula_Impost_ImpIEsp GROUP BY Remeses_Taula_Impost_ImpIEsp.numremesa ORDER BY Remeses_Taula_Impost_ImpIEsp.numremesa DESC"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 1500
  formseleccio.DBGrid2.Columns(1).Width = 1000
  formseleccio.DBGrid2.Columns(2).Width = 1000
  formseleccio.DBGrid2.Font.Size = 16
  formseleccio.Width = 5000
  formseleccio.Left = formseleccio.Left - 3000
  formseleccio.Show 1
  If seleccioret = 1 Then
           escullir_historic_Imp_i_Esp = cadbl(formseleccio.DBGrid2.Columns("numremesa"))
           vtitol = formseleccio.DBGrid2.Columns("Mes") + " " + formseleccio.DBGrid2.Columns("Any")
   End If
   Unload formseleccio
   
End Function

Private Sub mjustificants_Click()
Load formaltarep
  formaltarep.Caption = "Manteniment justificants merma"
  formaltarep.Tag = "justificants merma"
  formaltarep.Data1.DatabaseName = rutadelfitxer(cami) + "ImpostEnvasos.mdb"
  formaltarep.Data1.RecordSource = "select datafactura,numerofactura,tipus,kgfactura,nomproveidor,nifproveidor from facturesSAPreciclatge order by datafactura desc"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Width = 12000
 formaltarep.Width = 12400
 formaltarep.DBGrid1.Columns(0).Width = 1700
 formaltarep.DBGrid1.Columns(1).Width = 1700
 formaltarep.DBGrid1.Columns(2).Width = 1300
 formaltarep.DBGrid1.Columns(3).Width = 1400
 formaltarep.DBGrid1.Columns(4).Width = 2000
 formaltarep.DBGrid1.Columns(5).Width = 2500
 formaltarep.DBGrid1.AllowAddNew = False
  formaltarep.Show
End Sub

Private Sub mresumkgimpost_Click()
   Dim vImpostKgcomprats As Double
   Dim vImpostKgalmagatzem As Double
   Dim vImpostKgproduitNoEnviat As Double
   Dim vImpostKgproduint As Double
   Dim vImpostKgVenda As Double
   Dim vImpostKgMerma As Double
   Dim vnomfitxerCSV As String
   Dim vImpostKgMermaDevolucio As Double
   Dim vlinia As String
   If MsgBox("Aquest llistat triga uns segons un cop accepteu aquest missatge." + vbNewLine + "VOLS CONTINUAR?", vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbNo Then Exit Sub
   ratoli "espera"
   vImpostKgcomprats = calcular_KgImpost_compres
   vImpostKgalmagatzem = calcular_KgImpost_almagatzem + calcular_KgambImpostXcomanda("(CDbl([comanda]))>2000 And (CDbl([comanda]))<3000)")
   vImpostKgproduitNoEnviat = calcular_KgNoEnviats_i_Produint("noenviat")
   vImpostKgproduint = calcular_KgNoEnviats_i_Produint("produint")
   vImpostKgVenda = calcular_KgImpostVenuts
   vImpostKgMerma = calcular_KgImpostMerma
   vImpostKgMermaDevolucio = calcular_KgImpostMerma_devolucio
   
   vnomfitxerCSV = "c:\temp\LlistatEstadisticaImpostos.csv"
   Open vnomfitxerCSV For Output As #1
   vlinia = "ESTADISTICA D'IMPOSTOS COMPRAT/VENUT/MERMA"
   Print #1, vlinia
   vlinia = "============================================="
   Print #1, vlinia
   vlinia = ""
   Print #1, vlinia
   vlinia = "(1)Total Kg COMPRATS:;" + atrim(Redondejar(vImpostKgcomprats))
   Print #1, vlinia
   vlinia = "(2)Total Kg DISPONIBLE al magatzem:(Grups inclosos);" + atrim(Redondejar(vImpostKgalmagatzem))
   Print #1, vlinia
   vlinia = "(3)Total Kg producte ACABAT al magatzem:;" + atrim(Redondejar(vImpostKgproduitNoEnviat))
   Print #1, vlinia
   vlinia = "(4)Total Kg producte PRODUINT al magatzem:;" + atrim(Redondejar(vImpostKgproduint))
   Print #1, vlinia
   vlinia = "(5)Total Kg producte VENUTS(o devolució hisenda):;" + atrim(Redondejar(vImpostKgVenda))
   Print #1, vlinia
   vlinia = "(6)Total Kg producte MERMA(inclos 100 i 300):;" + atrim(Redondejar(vImpostKgMerma))
   Print #1, vlinia
   vlinia = "(7)Total Kg producte MERMA(inclos 100 i 300) (Devolució hisenda):;" + atrim(Redondejar(vImpostKgMermaDevolucio))
   Print #1, vlinia
   vlinia = "============================================="
   Print #1, vlinia
   vlinia = "Total:1-(2+3+4+5+6+7);" + atrim(Redondejar((vImpostKgcomprats - (vImpostKgalmagatzem + vImpostKgproduitNoEnviat + vImpostKgproduint + vImpostKgVenda + vImpostKgMerma + vImpostKgMermaDevolucio)), 0))
    Print #1, vlinia
   Close #1
   
  ratoli "normal"
  If existeix(vnomfitxerCSV) Then obrir_document vnomfitxerCSV
     
   
End Sub
Function calcular_KgImpostVenuts() As Double
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select Sum([KgVentaAd_Intracom]) AS TKgVenda from ImpostEnvasos where comanda not in (" + vcomandescomptades + ")")
   If Not rst.EOF Then calcular_KgImpostVenuts = cadbl(rst!TKgVenda)
   
   Set rst = dbtmp.OpenRecordset("select Sum([KgVentaImp_mes_Esp]) AS TKgVenda from ImpostEnvasos where comanda not in (" + vcomandescomptades + ")")
   If Not rst.EOF Then calcular_KgImpostVenuts = calcular_KgImpostVenuts + cadbl(rst!TKgVenda)
   
   Set rst = dbtmp.OpenRecordset("select Sum([KgventaEspanya]) AS TKgVenda from ImpostEnvasos where comanda not in (" + vcomandescomptades + ")")
   If Not rst.EOF Then calcular_KgImpostVenuts = calcular_KgImpostVenuts + cadbl(rst!TKgVenda)
   
   Set rst = Nothing
End Function
Function calcular_KgImpostMerma_devolucio() As Double
   Dim rst As Recordset
   Dim vsubsql As String
   vsubsql = "SELECT ImpostEnvasos.comanda FROM ImpostEnvasos LEFT JOIN liniesalbara ON ImpostEnvasos.id = liniesalbara.id WHERE liniesalbara.tipusdeentrega='T' AND (ImpostEnvasos.Num_remesa_ImpostEnv_Merma_Intra <>  Null and ImpostEnvasos.Num_remesa_ImpostEnv_Merma_Intra>0)"
   Set rst = dbtmp.OpenRecordset("select sum(KgMermaAd_Intracom) as TKgMerma from ImpostEnvasos where comanda not in (" + vcomandescomptades + ") and comanda in (" + vsubsql + ")")
   If Not rst.EOF Then calcular_KgImpostMerma_devolucio = calcular_KgImpostMerma_devolucio + cadbl(rst!TKgMerma)
   
   vsubsql = "SELECT ImpostEnvasos.comanda FROM ImpostEnvasos LEFT JOIN liniesalbara ON ImpostEnvasos.id = liniesalbara.id WHERE liniesalbara.tipusdeentrega='T' AND (ImpostEnvasos.Num_remesa_ImpostEnv_Merma_Imp <> Null and ImpostEnvasos.Num_remesa_ImpostEnv_Merma_Imp>0)"
   Set rst = dbtmp.OpenRecordset("select sum(KgMermaImp_mes_Esp) as TKgMerma from ImpostEnvasos where comanda not in (" + vcomandescomptades + ") and comanda in (" + vsubsql + ")")
   If Not rst.EOF Then calcular_KgImpostMerma_devolucio = calcular_KgImpostMerma_devolucio + cadbl(rst!TKgMerma)
   
   vsubsql = "SELECT ImpostEnvasos.comanda FROM ImpostEnvasos LEFT JOIN liniesalbara ON ImpostEnvasos.id = liniesalbara.id WHERE liniesalbara.tipusdeentrega='T' AND (ImpostEnvasos.Num_remesa_ImpostEnv_Merma_Esp <> Null Or ImpostEnvasos.Num_remesa_ImpostEnv_Merma_Esp>0)"
   Set rst = dbtmp.OpenRecordset("select sum(KgMermaEspanya) as TKgMerma from ImpostEnvasos where comanda not in (" + vcomandescomptades + ") and comanda in (" + vsubsql + ")")
   If Not rst.EOF Then calcular_KgImpostMerma_devolucio = calcular_KgImpostMerma_devolucio + cadbl(rst!TKgMerma)
   
   calcular_KgImpostMerma_devolucio = calcular_KgImpostMerma_devolucio + calcular_KgambImpostXcomanda("(CDbl([comanda]))=300 or (CDbl([comanda]))=100) and (numremesa>0 and numremesa<>null)")
   Set rst = Nothing
End Function

Function calcular_KgImpostMerma() As Double
   Dim rst As Recordset
   Dim vsubsql As String
   vsubsql = "SELECT ImpostEnvasos.comanda FROM ImpostEnvasos LEFT JOIN liniesalbara ON ImpostEnvasos.id = liniesalbara.id WHERE liniesalbara.tipusdeentrega='T' AND (ImpostEnvasos.Num_remesa_ImpostEnv_Merma_Intra Is Null Or ImpostEnvasos.Num_remesa_ImpostEnv_Merma_Intra=0)"
   Set rst = dbtmp.OpenRecordset("select sum(KgMermaAd_Intracom) as TKgMerma from ImpostEnvasos where comanda not in (" + vcomandescomptades + ") and comanda in (" + vsubsql + ")")
   If Not rst.EOF Then calcular_KgImpostMerma = calcular_KgImpostMerma + cadbl(rst!TKgMerma)
   
   vsubsql = "SELECT ImpostEnvasos.comanda FROM ImpostEnvasos LEFT JOIN liniesalbara ON ImpostEnvasos.id = liniesalbara.id WHERE liniesalbara.tipusdeentrega='T' AND (ImpostEnvasos.Num_remesa_ImpostEnv_Merma_Imp Is Null Or ImpostEnvasos.Num_remesa_ImpostEnv_Merma_Imp=0)"
   Set rst = dbtmp.OpenRecordset("select sum(KgMermaImp_mes_Esp) as TKgMerma from ImpostEnvasos where comanda not in (" + vcomandescomptades + ") and comanda in (" + vsubsql + ")")
   If Not rst.EOF Then calcular_KgImpostMerma = calcular_KgImpostMerma + cadbl(rst!TKgMerma)
   
   vsubsql = "SELECT ImpostEnvasos.comanda FROM ImpostEnvasos LEFT JOIN liniesalbara ON ImpostEnvasos.id = liniesalbara.id WHERE liniesalbara.tipusdeentrega='T' AND (ImpostEnvasos.Num_remesa_ImpostEnv_Merma_Esp Is Null Or ImpostEnvasos.Num_remesa_ImpostEnv_Merma_Esp=0)"
   Set rst = dbtmp.OpenRecordset("select sum(KgMermaEspanya) as TKgMerma from ImpostEnvasos where comanda not in (" + vcomandescomptades + ") and comanda in (" + vsubsql + ")")
   If Not rst.EOF Then calcular_KgImpostMerma = calcular_KgImpostMerma + cadbl(rst!TKgMerma)
   
   calcular_KgImpostMerma = calcular_KgImpostMerma + calcular_KgambImpostXcomanda("(CDbl([comanda]))=300 or (CDbl([comanda]))=100) and (numremesa=0 or numremesa=null)")
   Set rst = Nothing
End Function

Function calcular_KgambImpostXcomanda(vsubsql As String) As Double
   Dim vsql As String
   Dim rst As Recordset
   
   vsql = "SELECT Sum((([parcials].[metres]*[bobines].[pesdelproveidor])/[bobines].[mts])*([materials].[tanpercentimpostenvasos]/100)) AS KgambImpost"
   vsql = vsql + " FROM ((Palets LEFT JOIN materials ON Palets.codimatprognou = materials.codi) RIGHT JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) RIGHT JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet) "
   vsql = vsql + " WHERE (((Palets.teimpost)=True) AND (" + vsubsql + ");"
   'vsql = vsql + " WHERE (((Palets.teimpost)=True) AND (" + vsubsql + ");"
   Set rst = dbtmp.OpenRecordset(vsql)
   If Not rst.EOF Then calcular_KgambImpostXcomanda = cadbl(rst!KgambImpost)
   
   Set rst = Nothing
End Function
Function calcular_KgNoEnviats_i_Produint(vtipus As String) As Double
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim vkg As Double
   Dim vmsg As String
   Dim vnomfitxer As String
   
   If vtipus = "produint" Then vtipus = " proximaseccio<>'V' and proximaseccio<>'T' and comanda>200000 and client<>7 and client<>6842"
   If vtipus = "noenviat" Then vtipus = " (proximaseccio='V' or proximaseccio='P') and comanda>200000 and client<>7 or (client=6842 and comanda>200000)"
   
   vsql = "SELECT Sum((([parcials].[metres]*[bobines].[pesdelproveidor])/[bobines].[mts])*([materials].[tanpercentimpostenvasos]/100)) AS KgambImpost"
   vsql = vsql + " FROM ((Palets LEFT JOIN materials ON Palets.codimatprognou = materials.codi) RIGHT JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) RIGHT JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet) "
   vsql = vsql + " WHERE Palets.teimpost=True "

   
   'Clipboard.Clear
   'Clipboard.SetText vsql
   
   Set rstc = dbtmp.OpenRecordset("select * from comandes where " + vtipus)
   vkg = 0
   vmsg = vtipus + vbNewLine
   While Not rstc.EOF
      Set rst = dbtmp.OpenRecordset(vsql + " and Parcials.comanda='" + atrim(rstc!comanda) + "';")
      If Not rst.EOF Then
        If InStr(1, vcomandescomptades, atrim(rstc!comanda)) = 0 Then
          vcomandescomptades = vcomandescomptades + IIf(vcomandescomptades <> "", ",", "") + atrim(rstc!comanda)
          vkg = vkg + cadbl(rst!KgambImpost)
          If cadbl(rst!KgambImpost) > 0 Then vmsg = vmsg + vbNewLine + atrim(rstc!comanda) + ";" + atrim(rstc!proximaseccio) + ";" + atrim(rstc!client) + ";" + atrim(rstc!producte) + ";" + atrim(cadbl(rst!KgambImpost))
        End If
      End If
      rstc.MoveNext
   Wend
   calcular_KgNoEnviats_i_Produint = vkg
   vnomfitxer = "C:\TEMP\comandes_produint_" + Format(Now, "nn_ss") + ".csv"
   Open vnomfitxer For Output As #1
   Print #1, vmsg
   Close #1
   obrir_document vnomfitxer
  
   Set rstc = Nothing
   Set rst = Nothing
End Function
Function calcular_KgImpost_almagatzem() As Double
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("SELECT Sum((([bobines].[disponible]*[bobines].[pesdelproveidor])/[bobines].[mts])*([materials].[tanpercentimpostenvasos]/100)) AS KgambImpost FROM (Palets RIGHT JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi WHERE (((Palets.teimpost)=True));")
   If Not rst.EOF Then calcular_KgImpost_almagatzem = cadbl(rst!KgambImpost)
   Set rst = Nothing
End Function
Function calcular_KgImpost_compres() As Double
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select sum(kgimpostenvasos) as TotalKgImpost from albaransbip where (((albaransbip.KgImpostEnvasos)>0) AND ((albaransbip.numpalet)<>0 and (albaransbip.numpalet)<> Null));")
   If Not rst.EOF Then calcular_KgImpost_compres = cadbl(rst!TotalKgImpost)
   Set rst = Nothing
End Function

Private Sub mhistoricintra_Click()
   Dim v As String
   Dim vtitol As String
   v = escullir_historic_Intra(vtitol)
   If cadbl(v) = 0 Then Exit Sub
   dbtmp.Execute "delete * from taula_impost"
   dbtmp.Execute "insert into taula_impost select * from Remeses_Taula_Impost_Intracomunitaria where numremesa=" + v
   dataintracomunitari.RecordSource = "select * from taula_impost order by concepte,data"
   dataintracomunitari.Refresh
   etllistat.Caption = "Historic Intracomunitari Remesa:" + v + " " + UCase(vtitol)
   etllistat.Tag = "01/" + Mid(v, 6) + "/" + Mid(v, 2, 4)
   configurar_reixa
  emplenar_reixa
  vtipusllistat = "I"
  sumar_mermes_i_vendes
  bguardarCSV.Enabled = True
End Sub
Function escullir_historic_Intra(vtitol As String) As String
  Load formseleccio
  formseleccio.sortirs.Tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "ImpostEnvasos.mdb"
  formseleccio.Data1.RecordSource = "SELECT Remeses_Taula_Impost_Intracomunitaria.numremesa, Format(max(data),'mmmm') AS Mes, Year(max(data)) AS Any From Remeses_Taula_Impost_Intracomunitaria where concepte=1 GROUP BY Remeses_Taula_Impost_Intracomunitaria.numremesa ORDER BY Remeses_Taula_Impost_Intracomunitaria.numremesa DESC"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 1500
  formseleccio.DBGrid2.Columns(1).Width = 1000
  formseleccio.DBGrid2.Columns(2).Width = 1000
  formseleccio.DBGrid2.Font.Size = 16
  formseleccio.Width = 5000
  formseleccio.Left = formseleccio.Left - 3000
  formseleccio.Show 1
  If seleccioret = 1 Then
           escullir_historic_Intra = cadbl(formseleccio.DBGrid2.Columns("numremesa"))
           vtitol = formseleccio.DBGrid2.Columns("Mes") + " " + formseleccio.DBGrid2.Columns("Any")
   End If
   Unload formseleccio
   
End Function

Private Sub mresummermes_Click()
   Dim vinici As String
   Dim vfi As String
   Dim rstm As Recordset
   Dim vtotalventa As Double
   Dim vtotalmerma As Double
   Dim vfitxerCSV As String
   Dim rstc As Recordset
   Dim rsts As Recordset
   Dim vlinia As String
   
   vfitxerCSV = "c:\temp\llistat_percentatge_mermesImpost.CSV"
   vinici = InputBox("Entra la data d'inici de la consulta de mermes.", "Data inici")
   If StrPtr(vinici) = 0 Then Exit Sub
   If Not IsDate(vinici) Then Exit Sub
   vfi = InputBox("Entra la data de fi de la consulta de mermes.", "Data fi")
   If StrPtr(vinici) = 0 Then Exit Sub
   If Not IsDate(vfi) Then Exit Sub
   Set rstc = dbtmp.OpenRecordset("Select comandes.*,clients.nom FROM comandes LEFT JOIN clients ON comandes.client = clients.codi")
   Set rstm = dbtmp.OpenRecordset("SELECT impostenvasos.*, liniesalbara.tipusdeentrega FROM impostenvasos INNER JOIN liniesalbara ON impostenvasos.id = liniesalbara.id where data>=#" + Format(vinici, "mm/dd/yy") + "# and data<=#" + Format(vfi, "mm/dd/yy") + "# and (kgmermaad_intracom>0 or kgmermaimp_mes_esp>0)")
   Open vfitxerCSV For Output As #1
   Print #1, "Comanda;CodiCli;Nom Client;Total Venta;Total Merma;%MermaVsVenta"
   While Not rstm.EOF
     If atrim(rstm!tipusdeentrega) <> "P" Then  'nomes miro la merma si son entregues totals
      rstc.FindFirst "comanda=" + atrim(rstm!comanda)
      If Not rstc.NoMatch Then
        vtanpercent = 0
        Set rsts = dbtmp.OpenRecordset("select sum(kgventaimp_mes_esp+kgventaad_intracom) as sumaventa from impostenvasos where comanda=" + atrim(rstc!comanda))
        'vtotalventa = cadbl(rstm!kgventaad_intracom) + cadbl(rstm!kgventaimp_mes_esp)
        If Not rsts.EOF Then
            vtotalventa = cadbl(rsts!sumaventa)
            vtotalmerma = cadbl(rstm!kgmermaad_intracom) + cadbl(rstm!kgMERMAimp_mes_esp)
            If vtotalventa > 0 Then vtanpercent = Redondejar((vtotalmerma * 100) / vtotalventa, 0)
            vlinia = atrim(rstm!comanda) + ";" + atrim(rstc!client) + ";" + atrim(rstc!nom) + ";" + atrim(vtotalventa) + ";" + atrim(vtotalmerma) + ";" + atrim(vtanpercent)
            Print #1, vlinia
        End If
      End If
     End If
     rstm.MoveNext
   Wend
   Close 1
   If existeix(vfitxerCSV) Then obrir_document vfitxerCSV
End Sub

Private Sub tanxcentmermes_Click()
 Dim rstimpost As Recordset
   Dim rstimpost2 As Recordset
   Dim rstlinies As Recordset
   Dim rstlinia As Recordset
   Dim vkgmerma As Double
   Dim vcalcultanx100merma As Double
   Dim vkgtotalimpost As Double
   Dim vdata As String
   Dim vnomfitxer As String
   Dim vinici As String
   Dim vfi As String
   
   vinici = InputBox("Entra la data d'inici de la consulta de mermes.")
   vfi = InputBox("Entra la data fi de la consulta de mermes.")
   
   
   vnomfitxer = "c:\temp\CSV_Mermes.csv"
   Set rstlinia = dbtmp.OpenRecordset("SELECT Max(liniesalbara.id) AS Idlinia FROM capcaleraalbara RIGHT JOIN liniesalbara ON capcaleraalbara.numalbara = liniesalbara.numalbara Where [dataalbara]>=#" + Format(vinici, "mm/dd/yy") + "# and [dataalbara]<=#" + Format(vfi, "mm/dd/yy") + "# GROUP BY liniesalbara.lotinplacsa ORDER BY Max(capcaleraalbara.dataalbara);")
   Set rstimpost = dbtmp.OpenRecordset("select distinct comanda from impostenvasos where idliniaalbara<>null and year(data)=2024 ")
   Open vnomfitxer For Output As 1
   Print #1, "Albarà;Data albarà;Lot Inplacsa;Kg Impost;% de merma"
   With rstlinia
   While Not .EOF
     vkgmerma = 0
     vkgtotalimpost = 0
     Set rstlinies = dbtmp.OpenRecordset("SELECT capcaleraalbara.*, liniesalbara.* FROM capcaleraalbara RIGHT JOIN liniesalbara ON capcaleraalbara.numalbara = liniesalbara.numalbara where  id=" + atrim(cadbl(rstlinia!idlinia)))
     If rstlinies.EOF Then GoTo proxim
     vkgmerma = Redondejar(calculartanxcentmermadellot(rstlinies!lotinplacsa), 1)
     vdata = rstlinies!dataalbara
     vkgtotalimpost = rstlinies!KgImpostEnvasos
     If vkgmerma > 0 Then Print #1, atrim(rstlinies![capcaleraalbara.numalbara]) + ";" + atrim(vdata) + ";" + atrim(rstlinies!lotinplacsa) + ";" + atrim(vkgtotalimpost) + ";" + atrim(vkgmerma)
     
proxim:
     .MoveNext
   Wend
   End With
   Close 1
   If existeix(vnomfitxer) Then obrir_document vnomfitxer
End Sub
Function calculartanxcentmermadellot(vnumc As Double) As Double
   Dim rstimpost2 As Recordset
   Dim rstc As Recordset
   Set rstc = dbtmp.OpenRecordset("select linkcomanda1,linkcomanda2,comanda from comandes where comanda=" + atrim(vnumc))
   Set rstimpost2 = dbtmp.OpenRecordset("select * from impostenvasos where (kgmermaad_intracom>0 or kgmermaimp_mes_esp>0 or kgmermaespanya>0) and (comanda=" + atrim(vnumc) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2) + ")")
   Set rstimpost2 = dbtmp.OpenRecordset("select * from impostenvasos where idliniaalbara=" + atrim(cadbl(rstimpost2!idliniaalbara)))
   While Not rstimpost2.EOF
        vkgmerma = vkgmerma + cadbl(rstimpost2!kgMERMAimp_mes_esp) + cadbl(rstimpost2!kgmermaad_intracom) + cadbl(rstimpost2!kgMERMAespanya)
        vkgtotalimpost = vkgtotalimpost + cadbl(rstimpost2!Imp_mes_Esp_KgIMPOST) + cadbl(rstimpost2!Ad_Intracom_KgIMPOST) + cadbl(rstimpost2!Espanya_KgIMPOST) + cadbl(rstimpost2!KgMermaIMPOST_IE_capa) + cadbl(rstimpost2!KgMermaIMPOST_AD_capa) + cadbl(rstimpost2!KgMermaIMPOST_ES_capa)
        rstimpost2.MoveNext
     Wend
   If vkgtotalimpost > 0 Then calculartanxcentmermadellot = ((vkgmerma * 100) / vkgtotalimpost)
   Set rstimpost2 = Nothing
   Set rstc = Nothing
End Function

