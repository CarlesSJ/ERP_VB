VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formtransportistes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transportistes"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   19425
   Icon            =   "formtransportistes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   19425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bfiltre 
      Height          =   360
      Left            =   13860
      Picture         =   "formtransportistes.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "actualitzar"
      Top             =   450
      Width           =   360
   End
   Begin VB.TextBox cfiltrerecullida 
      Height          =   300
      Left            =   12630
      TabIndex        =   7
      Top             =   525
      Width           =   1215
   End
   Begin VB.CommandButton beliminaralbara 
      Height          =   300
      Left            =   4980
      Picture         =   "formtransportistes.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "Afegir un albarà a aquest avís."
      Top             =   1575
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton bafegiralbara 
      Height          =   300
      Left            =   4980
      Picture         =   "formtransportistes.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "Afegir un albarà a aquest avís."
      Top             =   1275
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.ListBox cllistaalbarans 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   3135
      TabIndex        =   5
      Top             =   1275
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Height          =   390
      Left            =   1185
      Picture         =   "formtransportistes.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Avisar al transportista."
      Top             =   390
      Width           =   960
   End
   Begin VB.CommandButton Command3 
      Height          =   390
      Left            =   225
      Picture         =   "formtransportistes.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   390
      Width           =   960
   End
   Begin VB.CommandButton Command7 
      Height          =   390
      Left            =   2145
      Picture         =   "formtransportistes.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir el CMR"
      Top             =   390
      Width           =   960
   End
   Begin VB.Data datatransportistes 
      Caption         =   "datatransportistes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4845
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Tots_transportistes_envios"
      Top             =   405
      Visible         =   0   'False
      Width           =   2745
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "formtransportistes.frx":2886
      Height          =   6210
      Left            =   165
      OleObjectBlob   =   "formtransportistes.frx":28A3
      TabIndex        =   0
      Top             =   840
      Width           =   19215
   End
   Begin VB.Menu mutils 
      Caption         =   "Utilitats"
      Begin VB.Menu memails 
         Caption         =   "Manteniment emails transportistes."
      End
   End
End
Attribute VB_Name = "formtransportistes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function enviaremail(sSendTo As String, sSubject As String, sText As String, sCosmissatge As String, adjunt As String, Optional noensenyarinterficie As Boolean) As Boolean
  Dim usuarim As String
  Dim contrasenyam As String
  Dim destinatari As String
  Dim vnomcarpeta As String
  Dim vadjunt As String
  
 
   usuarim = llegir_ini("Enviomails", "usuari", "comandes.ini")
   contrasenyam = llegir_ini("Enviomails", "contrasenya", "comandes.ini")
   If usuarim = "{[}]" Or contrasenyam = "{[}]" Then
      escriure_ini "Enviomails", "usuari", "expedicions@inplacsa.com", "comandes.ini"
      escriure_ini "Enviomails", "contrasenya", "isseyiznlzqmvtvt", "comandes.ini"
      'MsgBox "L'usuari o la contrasenya no estan entrades", vbCritical, "Error": Exit Function
   End If
   vadjunt = adjunt
   vnomcarpeta = "\\serverprodu\Dades\progcomandes\dades\spoolerenviament\" + nomordinador + "_" + Format(Now, "yymmdd_hhnnss")
   
   If Not existeix(vnomcarpeta) Then MkDir vnomcarpeta
   escriure_ini "Capcalera", "apuntperenviar", "No", vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "data", Now, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "nomordinador", nomordinador, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "usuari", usuarim, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "contrasenya", contrasenyam, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "destinatari", sSendTo, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "remitent", usuarim, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "assumpte", sSubject, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "adjunt", vnomcarpeta + "\" + substituirtot(vadjunt, rutadelfitxer(vadjunt), ""), vnomcarpeta + "\dadesmail.txt"
   Copiar_Fitxer adjunt, vnomcarpeta
   Open "c:\temp\cosmissatge.txt" For Output As #2
   Print #2, sCosmissatge
   Close #2
   Copiar_Fitxer "c:\temp\cosmissatge.txt", vnomcarpeta
   Kill "c:\temp\cosmissatge.txt"
  
   escriure_ini "Capcalera", "apuntperenviar", "Si", vnomcarpeta + "\dadesmail.txt"
   enviaremail = True
   
End Function


Function substituirtot(ByVal cadena As String, buscar As String, canviar As String) As String
   Dim comença As Integer
   Dim acaba As Integer
   If buscar = canviar Then GoTo fi
   While InStr(1, cadena, buscar) > 0
    comença = InStr(1, cadena, buscar) - 1
    
    If comença < 1 And InStr(1, cadena, buscar) <> 1 Then substituirtot = cadena: Exit Function
    acaba = comença + Len(buscar) + 1
    cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   Wend
fi:
   substituirtot = cadena
   'MsgBox linia
End Function


Private Sub bafegiralbara_Click()
  Dim vnumalb As Double
  Dim rstalb As Recordset
  Dim vnumavis As String
  If datatransportistes.Recordset!dataavis <> Null Then MsgBox "No es pot modificar els albarans si ja has demanat transport, primer elimina la data d'avís de transportista.", vbCritical, "Error": Exit Sub
  vnumalb = escullir_albara_enviament(datatransportistes.Recordset!idenvio, datatransportistes.Recordset!id_transport)
  If vnumalb = 0 Then Exit Sub
  Set rstalb = dbtmp.OpenRecordset("select * from Transportistes_avisos")
  rstalb.AddNew
  rstalb!idenvio = datatransportistes.Recordset!idenvio
  rstalb!coditransport = datatransportistes.Recordset!id_transport
  rstalb!numeroavis = datatransportistes.Recordset!numeroavis
  vnumavis = datatransportistes.Recordset!numeroavis
  rstalb!numalbara = vnumalb
  'rstalb!dataavis = Now
  rstalb!datarecullida = datatransportistes.Recordset!datarecullida
  rstalb.Update
   
   Set rstalb = Nothing
   datatransportistes.Refresh
   datatransportistes.Recordset.FindFirst "numeroavis='" + vnumavis + "'"
   If Not datatransportistes.Recordset.EOF Then
     carregar_albaransalallista datatransportistes.Recordset!numeroavis
   End If
   posarbotoafegiralbara
End Sub

Private Sub beliminaralbara_Click()
  Dim vnumalb As String
  Dim vnumavis As String
  If datatransportistes.Recordset!dataavis <> Null Then MsgBox "No es pot modificar els albarans si ja has demanat transport, primer elimina la data d'avís de transportista.", vbCritical, "Error": Exit Sub
  If cllistaalbarans.ListCount < 1 Then MsgBox "No hi ha cap albarà per eliminar", vbCritical, "Error": Exit Sub
  If cllistaalbarans.ListIndex = -1 Then cllistaalbarans.ListIndex = 0
  vnumalb = cadbl(Mid(cllistaalbarans.Text, 5))
  If MsgBox("Segur que vols eliminar l'albarà " + atrim(vnumalb) + " de la recullida?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
      vnumavis = datatransportistes.Recordset!numeroavis
      dbtmp.Execute "delete * from Transportistes_avisos where numalbara=" + atrim(vnumalb)
      datatransportistes.Refresh
      datatransportistes.Recordset.FindFirst "numeroavis='" + vnumavis + "'"
      If datatransportistes.Recordset.EOF Then
         carregar_albaransalallista 0
          Else: carregar_albaransalallista datatransportistes.Recordset!numeroavis
      End If
      posarbotoafegiralbara
  End If
End Sub

Private Sub bfiltre_Click()
   actualitzar_consulta_envios
   
End Sub
Sub actualitzar_consulta_envios()
   cllistaalbarans.visible = False
   bafegiralbara.visible = False
   beliminaralbara.visible = False
    If IsDate(cfiltrerecullida) Then
        datatransportistes.RecordSource = "select * from Tots_transportistes_envios where datarecullida>=#" + Format(cfiltrerecullida, "mm/dd/yy") + "# order by datarecullida"
        datatransportistes.Refresh
        
        If Not datatransportistes.Recordset.EOF Then datatransportistes.Recordset.MoveLast
        
          Else
            datatransportistes.RecordSource = "select * from Tots_transportistes_envios order by datarecullida desc"
            datatransportistes.Refresh
    End If
    
    If Not datatransportistes.Recordset.EOF Then datatransportistes.Recordset.MoveFirst: reixa.Row = 0
End Sub

Private Sub cllistaalbarans_DblClick()
   formvendes.datacapcalera.Recordset.FindFirst "numalbara=" + atrim(cadbl(Mid(cllistaalbarans, 5)))
   If Not formvendes.datacapcalera.Recordset.NoMatch Then Unload formtransportistes
End Sub

Private Sub cllistaalbarans_LostFocus()
'bafegiralbara.visible = False
End Sub

Private Sub Command1_Click()
   Dim rst As Recordset
   Dim vdia As Date
   Dim vidtransport As Double
   Dim vemails As String
   Dim vfitxeradjunt As String
   Dim vVectordiasetmana As Variant
   Dim vsql As String
   Dim vpais As String
   
   vVectordiasetmana = Array("Dilluns", "Dimarts", "Dimecres", "Dijous", "Divendres", "Dissabte", "Diumenge")
   
   escullir_transportidia vdia, vidtransport
   If vidtransport = 0 Then Exit Sub
   
 'aqui hauria de fer el select dels registres agrupats per pais i si aquest transportista
  '  te diferents emails per país hauria de treure-l's un per un
    'si es aquest cas fer primer un msgbox avisant que es passarà pais a pais
   Set rst = dbtmp.OpenRecordset("select first(pais) as unpais from tots_transportistes_envios where id_transport=" + atrim(vidtransport) + " and dataavis=null and datarecullida=#" + atrim(Format(vdia, "mm/dd/yy")) + "# group by pais")
   
   If Not rst.EOF Then
      vemails = emailsdetransportistaXrPais(vidtransport, rst!unpais)
      If vemails = "" Then Exit Sub
      rst.MoveLast: rst.MoveFirst
      If rst.RecordCount > 1 Then MsgBox "Aquest transportista te " + atrim(rst.RecordCount) + " enviaments a diferents països, els ensenyaré un per un.", vbInformation, "Atenció"
        Else: GoTo prepararenviament
   End If
   While Not rst.EOF
     vemails = ""
     If atrim(rst!unpais) <> "" Then
        vpais = atrim(rst!unpais)
prepararenviament:
        preparar_enviament_al_transportista vdia, vidtransport, vpais, vfitxeradjunt
        If MsgBox("Es correcte l'enviament que has vist?" + vbNewLine + "Vols enviar-lo al transportista per demanar l'enviament?", vbInformation + vbDefaultButton2 + vbYesNo, "Confirmació") = vbYes Then
           vemails = "seguimentexpedicions@inplacsa.com;expedicions@inplacsa.com"
           vemails = vemails + "; " + emailsdetransportistaXrPais(vidtransport, vpais)
           enviaremail vemails, vVectordiasetmana(WeekDay(vdia, vbMonday) - 1) + ": EXPEDICIÓ PER " + UCase(nompais(vpais)), "", "Bon dia." + vbNewLine + vbNewLine + "Adjunto recollida." + vbNewLine + "Gràcies." + vbNewLine + vbNewLine + "Salutacions:" + vbNewLine + "DEPARTAMENT D'EXPEDICIONS (INPLACSA)." + vbNewLine + vbNewLine + "HORARI MAGATZEM --> MATÍ : DE 7.30 A 12.30h  // TARDA : DE 15 A 16.30 h", vfitxeradjunt, False
         'despres d'enviar passar-lo a enviat posar data d'avís
           vsql = "select numeroavis from tots_transportistes_envios where dataavis=null and id_transport=" + atrim(vidtransport) + " and datarecullida=#" + Format(vdia, "mm/dd/yy") + "# and pais='" + atrim(vpais) + "'"
           dbtmp.Execute "update transportistes_avisos set dataavis=now where numeroavis in (" + vsql + ")"
           'enviaremailgeneric "miquel.inplacsa@gmail.com", "Sql despres enviament transportistes", "update transportistes_avisos set dataavis=now where numeroavis in (" + vsql + ")"
           MsgBox "E-Mail enviat, comprova a enviats que s'ha enviat correctament a:" + vbNewLine + vemails, vbInformation, "Atenció"
           datatransportistes.Refresh
        End If
     End If
     If vpais = "" Then GoTo fi
     rst.MoveNext
   Wend
fi:
Set rst = Nothing
End Sub
Function nompais(vpais As String) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select nompais from paisos where codipais='" + atrim(vpais) + "'")
  If Not rst.EOF Then nompais = atrim(rst!nompais)
  Set rst = Nothing
End Function

Function emailsdetransportistaXrPais(vidtransport As Double, vpais As String) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select email_expedicions,email_copiaexpedicions from transportistes where codi=" + atrim(vidtransport))
  If rst.EOF Then Exit Function
  If atrim(rst!email_expedicions) = "" Then MsgBox "No hi ha direcció d'email principal per aquest transportista.", vbCritical, "Error": Exit Function
  emailsdetransportistaXrPais = atrim(rst!email_expedicions) + IIf(rst!email_copiaexpedicions <> "", ";" + rst!email_copiaexpedicions, "")
  Set rst = dbtmp.OpenRecordset("select emailscontacte from transportistes_emailsperpais where id_transport=" + atrim(vidtransport) + " and codipais='" + vpais + "'")
  While Not rst.EOF
    emailsdetransportistaXrPais = IIf(emailsdetransportistaXrPais <> "", emailsdetransportistaXrPais + ";" + rst!emailscontacte, rst!emailscontacte)
    rst.MoveNext
  Wend
  Set rst = Nothing
End Function
Sub preparar_enviament_al_transportista(vdia As Date, vidtransport As Double, vpais As String, vfitxerPDF As String)

  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "solicitudtransport.rpt", 1)
 ' oreport.SQLQueryString = ""
  oreport.RecordSelectionFormula = "isnull({Tots_transportistes_envios.dataavis}) and {Tots_transportistes_envios.datarecullida}=#" + atrim(Format(vdia, "mm/dd/yy")) + "# and {Tots_transportistes_envios.id_transport}=" + atrim(vidtransport)
  If vpais <> "" Then oreport.RecordSelectionFormula = oreport.RecordSelectionFormula + " and {Tots_transportistes_envios.pais}='" + vpais + "'"
        
'  Clipboard.Clear
'  Clipboard.SetText oreport.RecordSelectionFormula
  
  'oreport.FormulaFields.GetItemByName("nomdirenvio").Text = "'" + treure_apostruf(etinfodelclient.tag) + "'"
  'oreport.SQLQueryString = ""
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "vendes.mdb"
  'oreport.Database.Tables.Item(2).Location = rutadelfitxer(cami) + "vendes.mdb"
  oreport.DiscardSavedData
  oreport.VerifyOnEveryPrint = False
  
   Load veurereport
   veurereport.width = 15000
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
  vfitxerPDF = "c:\temp\PDFtemporalsexpedicions\TransportsInplacsa_" + Format(Now, "ddmmyy") + ".pdf"
  borrarfitxerstemporalsPDF
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat
  oreport.ExportOptions.DiskFileName = vfitxerPDF
  oreport.ExportOptions.PDFExportAllPages = True
  oreport.Export False
End Sub
Sub borrarfitxerstemporalsPDF()
   On Error Resume Next
   Kill "c:\temp\PDFtemporalsexpedicions\*.*"
   MkDir "c:\temp\PDFtemporalsexpedicions"
End Sub
Sub escullir_transportidia(vdia As Date, vidtransport As Double)
  Dim vsql As String
  ', Count(Transportistes_avisos.numeroavis) AS [NºdeDestins]
  vsql = "SELECT Transportistes_avisos.datarecullida AS Data_Recullida, First(transportistes.descripcio) AS Transportista, transportistes.codi "
  vsql = vsql + " FROM (Transportistes_avisos LEFT JOIN Clients_envios ON Transportistes_avisos.idenvio = Clients_envios.id) LEFT JOIN transportistes ON Transportistes_avisos.coditransport = transportistes.codi "
  vsql = vsql + " Where (((Transportistes_avisos.dataavis) Is Null)) GROUP BY Transportistes_avisos.datarecullida, transportistes.codi"
  'Clipboard.Clear
'  Clipboard.SetText vsql
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
  formseleccio.Data1.RecordSource = vsql
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(2).visible = False
  formseleccio.DBGrid2.Columns(0).width = 1600
  formseleccio.DBGrid2.Columns(1).width = 3000
'  formseleccio.DBGrid2.Columns(3).width = 1300
  formseleccio.width = 7500
  formseleccio.Left = ((Screen.width / 2) - (formseleccio.width / 2))
  If formseleccio.Data1.Recordset.EOF Then MsgBox "NO HI HA RECOLLIDA PENDENT.": Exit Sub
  formseleccio.Data1.Recordset.MoveLast
  formseleccio.Data1.Recordset.MoveFirst
  formseleccio.Show 1
  If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           vdia = formseleccio.DBGrid2.Columns("Data_Recullida")
           vidtransport = cadbl(formseleccio.DBGrid2.Columns("codi"))
        End If
  End If
  Unload formseleccio
End Sub
Private Sub Command3_Click()
   Dim vnumalb As Double
   Dim vnumavis As String
   vnumalb = escullir_albara_enviament
   If cadbl(vnumalb) = 0 Then vnumalb = cadbl(InputBox("No has escullit cap número d'albarà, vols entrar-lo manualment? Escriu-lo si vols.", "Albarà manual"))
   afegir_albara_anumerodavis vnumalb, vnumavis
   datatransportistes.Refresh
   datatransportistes.Recordset.FindFirst "numeroavis='" + atrim(vnumavis) + "'"
End Sub
Sub afegir_albara_anumerodavis(vnumalb As Double, vnumavis As String)
  Dim rst As Recordset
  Dim rstalb As Recordset
  Dim vdatarecullida As String
  If vnumalb = 0 Then Exit Sub
  vdatarecullida = InputBox("Entra la data que vols que vingui el transportista.", "Data recullida", Format(DateAdd("d", 1, Now), "dd/mm/yy"))
  If Not IsDate(vdatarecullida) Then Exit Sub
  'vnumavis = escullir_dia_avisSiIgual
  vnumavis = Format(Now, "yymmddhhmmss")
  Set rstalb = dbtmp.OpenRecordset("select * from transportistes_avisos")
  Set rst = dbtmp.OpenRecordset("select * from capcaleraalbara where numalbara=" + atrim(vnumalb))
  If rst.EOF Then Exit Sub
  rstalb.FindFirst "idenvio=" + atrim(rst!id_direnvio) + " and coditransport=" + atrim(rst!id_transport) + " and datarecullida=#" + Format(vdatarecullida, "mm/dd/yy") + "#"
  If Not rstalb.NoMatch Then
      If MsgBox("Veig que ja has fet un avís per aquest dia amb aquesta direcció d'enviament i aquest transportista, vols utilitzar el mateix avís [SI] o fer-ne un altra[NO]?", vbExclamation + vbDefaultButton1 + vbYesNo, "Atenció") = vbYes Then
            vnumavis = rstalb!numeroavis
      End If
  End If
  rstalb.AddNew
  rstalb!idenvio = rst!id_direnvio
  rstalb!coditransport = rst!id_transport
    rstalb!numeroavis = vnumavis
  rstalb!numalbara = vnumalb
  'rstalb!dataavis = Now
  rstalb!INCOTERMS = crearINCOTERMS(vnumalb)
  rstalb!datarecullida = vdatarecullida
  rstalb.Update
  actualitzar_avis_registre_enviaments vnumalb, vnumavis
End Sub
Function crearINCOTERMS(vnumalb As Double) As String
  Dim rst As Recordset
  Dim vnumc As Double
  Set rst = dbtmp.OpenRecordset("select lotinplacsa from liniesalbara where numalbara=" + atrim(vnumalb))
  If Not rst.EOF Then
     vnumc = rst!lotinplacsa
     Set rst = dbtmp.OpenRecordset("select incoterm_envio from comandes_extres where comanda=" + atrim(vnumc))
     If Not rst.EOF Then
         crearINCOTERMS = atrim(rst!incoterm_envio)
         Set rst = dbtmp.OpenRecordset("select direnvio from comandes where comanda=" + atrim(vnumc))
         If Not rst.EOF Then
             Set rst = dbtmp.OpenRecordset("select poblacioe from clients_envios where id=" + atrim(rst!direnvio))
             If Not rst.EOF Then crearINCOTERMS = Mid(atrim(crearINCOTERMS + " - " + atrim(rst!poblacioe)), 1, 49)
         End If
     End If
  End If
  Set rst = Nothing
  
End Function
Sub actualitzar_avis_registre_enviaments(vnumalb As Double, vnumavis As String)
  Dim rst As Recordset
  Dim rste As Recordset
  Set rst = dbtmp.OpenRecordset("select lotinplacsa from liniesalbara where numalbara=" + atrim(vnumalb))
  If Not rst.EOF Then
      Set rste = dbtmp.OpenRecordset("select * from registre_enviaments where comandesrelacionades like '*" + Trim(rst!lotinplacsa) + " *' and (numeroavis='' or numeroavis=null)")
      If rste.EOF Then GoTo fi
      rste.Edit
      rste!numeroavis = vnumavis
      rste.Update
  End If
fi:
  Set rst = Nothing
End Sub
Function escullir_albara_enviament(Optional videnvio As Double, Optional vidtransport As Double) As Double
  Dim vsql As String
  Dim vsubsql As String
  If videnvio <> 0 Then vsubsql = " and (id_direnvio=" + atrim(videnvio) + " and id_transport=" + atrim(vidtransport) + ") "
  vsql = "SELECT capcaleraalbara.numalbara, transportistes.descripcio, Clients_envios.nome,clients_envios.poblacioe, Clients_envios.provinciae "
  vsql = vsql + " FROM (capcaleraalbara LEFT JOIN transportistes ON capcaleraalbara.id_transport = transportistes.codi) LEFT JOIN Clients_envios ON capcaleraalbara.id_direnvio = Clients_envios.id "
  vsql = vsql + " Where (capcaleraalbara.numalbara not in (select numalbara from Transportistes_avisos)) and (((transportistes.descripcio) <> '') And ((capcaleraalbara.dataenvioasap) Is Null))"
  vsql = vsql + vsubsql + " ORDER BY transportistes.descripcio;"
  'Clipboard.Clear
 ' Clipboard.SetText vsql
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
  formseleccio.Data1.RecordSource = vsql
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(3).width = 2500
  formseleccio.DBGrid2.Columns(4).width = 2500
  formseleccio.DBGrid2.Columns(2).width = 2500
  formseleccio.width = 15000
  formseleccio.Left = ((Screen.width / 2) - (formseleccio.width / 2))
  If formseleccio.Data1.Recordset.EOF Then
        MsgBox "NO HI HA CAP ALBARÀ PENDENT D'ENVIAR AMB TRANSPORTISTA ESCULLIT." + vbNewLine + "PODRIA SER QUE JA L'HAGIS PUJAT AL SAP."
        escullir_albara_enviament = cadbl(InputBox("Escriu el numero d'albarà manualment si vols.", "Albarà manual"))
        GoTo fi
  End If
  formseleccio.Data1.Recordset.MoveLast
  formseleccio.Data1.Recordset.MoveFirst
  formseleccio.Show 1
  If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           escullir_albara_enviament = cadbl(formseleccio.DBGrid2.Columns("NUMALBARA"))
        End If
  End If
fi:
  Unload formseleccio
End Function

Private Sub Command5_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command6_Click()

End Sub

Function generarINCOTERMS(vnumavis As String) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from transportistes_avisos where numeroavis='" + atrim(vnumavis) + "'")
  If Not rst.EOF Then
      generarINCOTERMS = crearINCOTERMS(cadbl(rst!numalbara))
  End If
  Set rst = Nothing
End Function
Private Sub Command7_Click()
  Dim rst As Recordset
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vmatricula As String
  Dim vmatricularemolc As String
  Dim vnumavis As String
  Dim vfitxerPDF As String
  Dim vdirenvio As String
  Dim vincoterm As String
  
  If datatransportistes.Recordset.EOF Then Exit Sub
  vnumavis = datatransportistes.Recordset!numeroavis
  
  vmatricula = InputBox("Escriu la matricula del camió/tractora:", "Matricula", atrim(datatransportistes.Recordset!matriculacamio))
  vmatricularemolc = InputBox("Escriu la matricula del remolc:", "Matricula", atrim(datatransportistes.Recordset!matricularemolc))
  dbtmp.Execute "update Transportistes_avisos set matricularemolc='" + vmatricularemolc + "',matriculacamio='" + vmatricula + "' where numeroavis='" + atrim(vnumavis) + "'"
  vincoterm = atrim(datatransportistes.Recordset!INCOTERMS)
  If vincoterm = "" Then
      vincoterm = generarINCOTERMS(vnumavis)
      dbtmp.Execute "update Transportistes_avisos set incoterms='" + vincoterm + "' where numeroavis='" + atrim(vnumavis) + "'"
  End If
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "CMR_expedicions.rpt", 1)
  
  'oreport.FormulaFields.GetItemByName("nomdirenvio").Text = "'" + treure_apostruf(etinfodelclient.tag) + "'"
  oreport.FormulaFields.GetItemByName("matricules").Text = """" + vmatricula + " <br>" + vmatricularemolc + """"
  oreport.FormulaFields.GetItemByName("dataenviament").Text = "'" + atrim(datatransportistes.Recordset!datarecullida) + "'"
  oreport.FormulaFields.GetItemByName("numbases").Text = "'" + atrim(datatransportistes.Recordset!bases) + " BASES'"
  oreport.FormulaFields.GetItemByName("numpalets").Text = "'" + atrim(datatransportistes.Recordset!palets) + " PALETS'"
  oreport.FormulaFields.GetItemByName("metrescubics").Text = IIf(datatransportistes.Recordset!metres3 > 0, "'" + atrim(datatransportistes.Recordset!metres3) + " M3'", "''")
  oreport.FormulaFields.GetItemByName("numeroenviament").Text = "'" + atrim(datatransportistes.Recordset!numeroavis) + "'"
  oreport.FormulaFields.GetItemByName("relacioalbarans").Text = "'DELIVERY NOTES: " + atrim(buscaralbaransSAPdaquestenviament(vnumavis)) + "'" ' + atrim(buscarnumcomandaclientdaquestenviament(vnumavis), datatransportistes.Recordset!idenvio) + "'"
  oreport.FormulaFields.GetItemByName("totalkg").Text = atrim(datatransportistes.Recordset!kgs)
  oreport.FormulaFields.GetItemByName("nomtransport").Text = "'" + nomdeltransportista(datatransportistes.Recordset!id_transport) + "'"
  oreport.FormulaFields.GetItemByName("direccioclient").Text = "'" + nomdelclient(datatransportistes.Recordset!idenvio, False) + "'"
  oreport.FormulaFields.GetItemByName("llocentrega").Text = "'" + nomdelclient(datatransportistes.Recordset!idenvio, True) + "'"
  oreport.FormulaFields.GetItemByName("observacionsexpedicions").Text = "'" + observacionsdelremitent(datatransportistes.Recordset!idenvio, vnumavis) + "'"
  oreport.FormulaFields.GetItemByName("INCOTERMcondicionsentrega").Text = "'" + atrim(vincoterm) + "'"
  
  If InStr(1, oreport.FormulaFields.GetItemByName("relacioalbarans").Text, " 0") > 0 Then MsgBox "Hi ha algun albarà que encara no està pujat a SAP.", vbCritical, "Error"
 ' MsgBox oreport.FormulaFields.GetItemByName("relacioalbarans").Text
  oreport.SQLQueryString = ""
'  oreport.Database.Tables.Item(0).Location = ""
  oreport.DiscardSavedData
  oreport.VerifyOnEveryPrint = False
  
   Load veurereport
   veurereport.width = 15000
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
   If MsgBox("Vols enviar el CMR per email a comercial?", vbDefaultButton2 + vbYesNo + vbDefaultButton2, "Enviar CMR a comercial") = vbYes Then
       vfitxerPDF = "CMR_" + atrim(datatransportistes.Recordset!descripcio) + "(" + atrim(datatransportistes.Recordset!idenvio) + ")_" + atrim(datatransportistes.Recordset!datarecullida) + "_Nº:" + atrim(datatransportistes.Recordset!numeroavis) + ".pdf"
       vfitxerPDF = treuresimbolsnovalidsnomfitxer(vfitxerPDF)
       vfitxerPDF = "c:\temp\PDFtemporalsexpedicions\" + vfitxerPDF
       oreport.ExportOptions.DestinationType = crEDTDiskFile
       oreport.ExportOptions.FormatType = crEFTPortableDocFormat
       oreport.ExportOptions.DiskFileName = vfitxerPDF
       oreport.ExportOptions.PDFExportAllPages = True
       oreport.Export False
       'recepcion_inplacsa@inplacsa.com
       vdirenvio = nomdelclient(datatransportistes.Recordset!idenvio, False)
       vdirenvio = substituirtot(vdirenvio, "<br>", vbNewLine)
       If existeix(vfitxerPDF) Then enviaremail "recepcion_inplacsa@inplacsa.com", "CMR " + atrim(datatransportistes.Recordset!descripcio) + " NºCMR:" + atrim(datatransportistes.Recordset!numeroavis), "", "Adjunto CMR recollida." + vbNewLine + atrim(datatransportistes.Recordset!descripcio) + vbNewLine + vbNewLine + vdirenvio, vfitxerPDF, False
   End If
   datatransportistes.Refresh
   datatransportistes.Recordset.FindFirst "numeroavis='" + atrim(vnumavis) + "'"
End Sub
Function observacionsdelremitent(videnvio As Double, vnumavis As String) As String
   Dim rst As Recordset
   Dim vcomandaclient As String
   Set rst = dbtmp.OpenRecordset("select * from clients_envios where id=" + atrim(videnvio))
   If Not rst.EOF Then
      observacionsdelremitent = treure_apostruf(atrim(rst!cmr_observacions))
      If rst!cmr_comandaclient Then
          Set rst = dbtmp.OpenRecordset("SELECT Transportistes_avisos.numeroavis, comandes.comandaclient, liniesalbara.lotinplacsa FROM (Transportistes_avisos LEFT JOIN liniesalbara ON Transportistes_avisos.numalbara = liniesalbara.numalbara) LEFT JOIN comandes ON liniesalbara.lotinplacsa = comandes.comanda WHERE (((Transportistes_avisos.numeroavis)='" + vnumavis + "'));")
          While Not rst.EOF
             If atrim(rst!comandaclient) <> "" Then vcomandaclient = vcomandaclient + IIf(vcomandaclient <> "", ";", "") + atrim(rst!comandaclient)
             rst.MoveNext
          Wend
          If vcomandaclient <> "" Then vcomandaclient = "<b><i>Ref: " + vcomandaclient
      End If
   End If
   observacionsdelremitent = possar_caracters_html("<center>" + substituirtot(observacionsdelremitent + "<br>" + vbNewLine + vcomandaclient, vbNewLine, "<br>"))
   Set rst = Nothing
   
End Function
Function nomdelclient(videnvio As Double, vllocentrega As Boolean) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("SELECT clients.*, Clients_envios.* FROM clients RIGHT JOIN Clients_envios ON clients.codi = Clients_envios.codi Where clients_envios.id = " + atrim(videnvio))
  If rst.EOF Then Exit Function
  If Not vllocentrega Then
      nomdelclient = atrim(rst!nom) + "<br>" + atrim(rst!domicili) + "<br>" + atrim(rst!codipostal) + " " + atrim(rst!poblacio) + "(" + atrim(rst!provincia) + ")"
        Else
         nomdelclient = atrim(rst![clients_envios.nome]) + "<br>" + atrim(rst![clients_envios.domicilie]) + "<br>" + atrim(rst![clients_envios.codipostale]) + " " + atrim(rst![clients_envios.poblacioe]) + "(" + atrim(rst![clients_envios.provinciae]) + ") " + "[" + atrim(rst![pais]) + "]"
  End If
  nomdelclient = treure_apostruf(nomdelclient)
  nomdelclient = possar_caracters_html(nomdelclient)
End Function
Function nomdeltransportista(vidtransport As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from transportistes where codi=" + atrim(vidtransport))
   If Not rst.EOF Then
       nomdeltransportista = atrim(rst!descripcio) + " <br>" + atrim(rst!direccio) + "<br>" + atrim(rst!CPpoblacioPais)
   End If
   nomdeltransportista = possar_caracters_html(nomdeltransportista)
   Set rst = Nothing
End Function
Function possar_caracters_html(ByVal v As String) As String
'    v = substituirtot(v, "&", "&amp;")
    v = substituirtot(v, "À", "&Agrave;")
    v = substituirtot(v, "Á", "&Aacute;")
    v = substituirtot(v, "Â", "&Acirc;")
    v = substituirtot(v, "Ã", "&Atilde;")
    v = substituirtot(v, "Ä", "&Auml;")
    v = substituirtot(v, "È", "&Egrave;")
    v = substituirtot(v, "É", "&Eacute;")
    v = substituirtot(v, "Ê", "&Ecirc;")
    v = substituirtot(v, "Ë", "&Euml;")
    v = substituirtot(v, "Ì", "&Igrave;")
    v = substituirtot(v, "Í", "&Iacute;")
    v = substituirtot(v, "Î", "&Icirc;")
    v = substituirtot(v, "Ï", "&Iuml;")
    v = substituirtot(v, "Ò", "&Ograve;")
    v = substituirtot(v, "Ó", "&Oacute;")
    v = substituirtot(v, "Ô", "&Ocirc;")
    v = substituirtot(v, "Õ", "&Otilde;")
    v = substituirtot(v, "Ö", "&Ouml;")
    v = substituirtot(v, "Ù", "&Ugrave;")
    v = substituirtot(v, "Ú", "&Uacute;")
    v = substituirtot(v, "Û", "&Ucirc;")
    v = substituirtot(v, "Ü", "&Uuml;")
    v = substituirtot(v, "Š", "&Scaron;")
    v = substituirtot(v, "Ý", "&Yacute;")
    v = substituirtot(v, "Ÿ", "&Yuml;")
    v = substituirtot(v, "à", "&agrave;")
    v = substituirtot(v, "á", "&aacute;")
    v = substituirtot(v, "â", "&acirc;")
    v = substituirtot(v, "ã", "&atilde;")
    v = substituirtot(v, "ä", "&auml;")
    v = substituirtot(v, "è", "&egrave;")
    v = substituirtot(v, "é", "&eacute;")
    v = substituirtot(v, "ê", "&ecirc;")
    v = substituirtot(v, "ë", "&euml;")
    v = substituirtot(v, "ì", "&igrave;")
    v = substituirtot(v, "í", "&iacute;")
    v = substituirtot(v, "î", "&icirc;")
    v = substituirtot(v, "ï", "&iuml;")
    v = substituirtot(v, "ò", "&ograve;")
    v = substituirtot(v, "ó", "&oacute;")
    v = substituirtot(v, "ô", "&ocirc;")
    v = substituirtot(v, "õ", "&otilde;")
    v = substituirtot(v, "ö", "&ouml;")
    v = substituirtot(v, "ù", "&ugrave;")
    v = substituirtot(v, "ú", "&uacute;")
    v = substituirtot(v, "û", "&ucirc;")
    v = substituirtot(v, "ü", "&uuml;")
    v = substituirtot(v, "š", "&scaron;")
    v = substituirtot(v, "ý", "&yacute;")
    v = substituirtot(v, "ÿ", "&yuml;")
    v = substituirtot(v, "ç", "&#231;")
    v = substituirtot(v, "Ç", "&#199;")
    v = substituirtot(v, "ñ", "&#241;")
    v = substituirtot(v, "Ñ", "&#209,")

    
    
    possar_caracters_html = v
End Function
Function buscaralbaransSAPdaquestenviament(vnumavis As String) As String
   Dim rst As Recordset
   Dim vsql As String
   
   vsql = "SELECT Transportistes_avisos.numeroavis, capcaleraalbara.numalbaraSAP FROM Transportistes_avisos LEFT JOIN capcaleraalbara ON Transportistes_avisos.numalbara = capcaleraalbara.numalbara "
   vsql = vsql + " where numeroavis='" + vnumavis + "'"
   Set rst = dbtmp.OpenRecordset(vsql)
   While Not rst.EOF
      buscaralbaransSAPdaquestenviament = buscaralbaransSAPdaquestenviament + IIf(buscaralbaransSAPdaquestenviament <> "", " - ", "") + atrim(rst!numalbaraSAP)
      rst.MoveNext
   Wend
  Set rst = Nothing
End Function
Function buscarnumcomandaclientdaquestenviament(vnumavis As String, videnvio As Long) As String
   Dim rst As Recordset
   Dim vsql As String
   vsql = "SELECT Transportistes_avisos.numeroavis, liniesalbara.numcomandacli, liniesalbara.lotinplacsa, Transportistes_avisos.numalbara, Clients_envios.avisrebobinadora FROM ((Transportistes_avisos LEFT JOIN liniesalbara ON Transportistes_avisos.numalbara = liniesalbara.numalbara) LEFT JOIN capcaleraalbara ON Transportistes_avisos.numalbara = capcaleraalbara.numalbara) LEFT JOIN Clients_envios ON capcaleraalbara.id_direnvio = Clients_envios.id "
   vsql = vsql + " where numeroavis='" + vnumavis + "'"
   Set rst = dbtmp.OpenRecordset(vsql)
'   Clipboard.Clear
'   Clipboard.SetText vsql
   While Not rst.EOF
      buscarnumcomandaclientdaquestenviament = buscarnumcomandaclientdaquestenviament + IIf(buscarnumcomandaclientdaquestenviament <> "", " - ", "") + "[" + atrim(rst!numcomandacli) + "]"
      rst.MoveNext
   Wend
   If buscarnumcomandaclientdaquestenviament <> "" Then buscarnumcomandaclientdaquestenviament = "<br>Customer request: " + buscarnumcomandaclientdaquestenviament
   buscarnumcomandaclientdaquestenviament = possar_caracters_html(buscarnumcomandaclientdaquestenviament)
  Set rst = Nothing
End Function



Private Sub Form_Load()
  Set dbtmp = OpenDatabase(rutadelfitxer(cami) + "vendes.mdb")
  datatransportistes.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
  cfiltrerecullida = Date
  actualitzar_consulta_envios
  
  'actualitzar_avis_registre_enviaments
End Sub

Private Sub memails_Click()
   Formmantenimenttransportistes.Show 1
End Sub

Private Sub reixa_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
   posarbotoafegiralbara
End Sub

Private Sub reixa_DblClick()
  Dim vdata As String
  Dim vnumavis As String
  If reixa.Columns(reixa.col).DataField = "dataavis" Then
     If reixa.Columns(reixa.col) <> "" Then
       If MsgBox("Vols borrar aquesta data d'avís?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
            dbtmp.Execute "update Transportistes_avisos set dataavis=null where numeroavis='" + atrim(datatransportistes.Recordset!numeroavis) + "'"
            vnumavis = datatransportistes.Recordset!numeroavis
            datatransportistes.Refresh
            datatransportistes.Recordset.FindFirst "numeroavis='" + atrim(vnumavis) + "'"
            MsgBox "RECORDA QUE SI ELIMINES LA DATA D'AVIS TENS QUE TORNAR A DEMANAR EL SERVEI AL TRANSPORTISTA.", vbCritical, "A T E N C I Ó"
       End If
          Else
            vdata = InputBox("Entra la data d'avís que vols possar?", "Atenció", Format(Now, "dd/mm/yy"))
            If IsDate(vdata) Then
                dbtmp.Execute "update Transportistes_avisos set dataavis=#" + Format(vdata, "mm/dd/yy") + "# where numeroavis='" + atrim(datatransportistes.Recordset!numeroavis) + "'"
                vnumavis = datatransportistes.Recordset!numeroavis
                datatransportistes.Refresh
                datatransportistes.Recordset.FindFirst "numeroavis='" + atrim(vnumavis) + "'"
                MsgBox "RECORDA QUE POSSANT LA DATA MANUALMENT EL TRANSPORTISTA, NI ELS CORREUS INTERNS D'INPLACSA, REBEN L'AVÍS.", vbCritical, "A T E N C I Ó"
            End If
            vdata = ""
     End If
  End If
    If reixa.Columns(reixa.col).DataField = "datarecullida" Then
     If reixa.Columns(reixa.col) <> "" Then
       If Not datatransportistes.Recordset.EOF Then
         If Not IsNull(datatransportistes.Recordset!dataavis) Then MsgBox "No pots canviar la data de recullida amb una data d'avis feta, ELIMINA PRIMER LA DATA D'AVÍS", vbCritical, "ERROR": Exit Sub
       End If
       vdata = InputBox("Entra la nova data de recullida.", "Nova data de recullida")
       If IsDate(vdata) Then
            dbtmp.Execute "update Transportistes_avisos set datarecullida=#" + Format(vdata, "mm/dd/yy") + "# where numeroavis='" + atrim(datatransportistes.Recordset!numeroavis) + "'"
            vnumavis = datatransportistes.Recordset!numeroavis
            datatransportistes.Refresh
            datatransportistes.Recordset.FindFirst "numeroavis='" + atrim(vnumavis) + "'"
       End If
     End If
  End If
End Sub

Private Sub reixa_GotFocus()
     cllistaalbarans.visible = True
End Sub

Private Sub reixa_LostFocus()
   'cllistaalbarans.visible = False
End Sub

Private Sub reixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  posarbotoafegiralbara
End Sub
Sub posarbotoafegiralbara()
'   bafegiralbara.Top = (reixa.Row) * reixa.RowHeight + reixa.Top + 220
'   bafegiralbara.Height = reixa.RowHeight
'   bafegiralbara.Left = reixa.Left + reixa.Columns(2).Left - bafegiralbara.width
   cllistaalbarans.Top = (reixa.Row) * reixa.RowHeight + reixa.Top + 220 + reixa.RowHeight
 'cllistaalbarans.Height = reixa.RowHeight
   cllistaalbarans.Left = reixa.Left + reixa.Columns(1).Left
   cllistaalbarans.width = reixa.Columns(1).width
   If Not datatransportistes.Recordset.EOF Then carregar_albaransalallista datatransportistes.Recordset!numeroavis
   If cllistaalbarans.ListCount > 0 Then
     cllistaalbarans.visible = True
     bafegiralbara.visible = True
     beliminaralbara.visible = True
      Else:
        cllistaalbarans.visible = False
        bafegiralbara.visible = False
        beliminaralbara.visible = False
   End If
'  bafegiralbara.visible = True
'  beliminaralbara.visible = True
  bafegiralbara.Left = cllistaalbarans.Left + (cllistaalbarans.width)
  bafegiralbara.Top = cllistaalbarans.Top
  beliminaralbara.Left = cllistaalbarans.Left + (cllistaalbarans.width)
  beliminaralbara.Top = bafegiralbara.Top + bafegiralbara.Height
  bafegiralbara.ZOrder 0
  beliminaralbara.ZOrder 0
End Sub
Sub carregar_albaransalallista(vnumeroavis As String)
  Dim rst As Recordset
  cllistaalbarans.Clear
  
  Set rst = dbtmp.OpenRecordset("select * from transportistes_avisos where numeroavis='" + atrim(vnumeroavis) + "'")
  While Not rst.EOF
     cllistaalbarans.AddItem "Alb: " + atrim(rst!numalbara) + vbNewLine
   '  actualitzar_avis_registre_enviaments rst!numalbara, datatransportistes.Recordset!numeroavis
     rst.MoveNext
  Wend
  cllistaalbarans.Height = cllistaalbarans.ListCount * 400
  Set rst = Nothing
End Sub
Private Sub reixa_RowResize(Cancel As Integer)
   posarbotoafegiralbara
End Sub

