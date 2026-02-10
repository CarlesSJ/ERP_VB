VERSION 5.00
Begin VB.Form clients_envios 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "clients_envios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub camp1_Change()

End Sub

Private Sub alta_Click()
alta_registre
End Sub

Private Sub Check10_Click()
 Check9.Value = 0
 Text44.Text = "0"
End Sub

Private Sub Check13_Click()
  If Check13.Value > 0 Then
     framepesnet.Visible = True
      carregar_pesosnets
       Else: framepesnet.Visible = False
  End If
End Sub
Sub carregar_pesosnets()
  datapesnet.RecordSource = "select * from taulapesnet where idenvio=" + atrim(cadbl(envios.Tag))
  datapesnet.Refresh
  If datapesnet.Recordset.EOF Then
   For i = 0 To 120 Step 5
     datapesnet.Recordset.AddNew
     datapesnet.Recordset!idenvio = cadbl(envios.Tag)
     datapesnet.Recordset!mida = i
     datapesnet.Recordset.Update
   Next i
    Else: datapesnet.Recordset.MoveFirst
  End If
End Sub
Private Sub Check9_Click()
Check10.Value = 0
End Sub

Private Sub clients_Reposition()
  carregar_lookups
  clients.Caption = "Clients:  " + atrim(cadbl(clients.Recordset.AbsolutePosition) + 1) + " de " + atrim(clients.Recordset.RecordCount)
  If clients.EditMode = 0 Then areadatos.Enabled = False
End Sub

Sub triarrepresentant()
  Load formseleccio
  formseleccio.Data1.DatabaseName = clients.DatabaseName
  formseleccio.Data1.RecordSource = "select * from representants"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text26.Text = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   Text27.Text = atrim(formseleccio.Data1.Recordset!nom)
  End If
  Unload formseleccio
  
End Sub

Sub triarformapag()
  Load formseleccio
  formseleccio.Data1.DatabaseName = clients.DatabaseName
  formseleccio.Data1.RecordSource = "select * from [formes de pagament]"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
    If cadbl((formseleccio.Data1.Recordset!codi)) > 0 Then
      Text28.Text = atrim((formseleccio.Data1.Recordset!codi))
      Text29.Text = atrim(formseleccio.Data1.Recordset!descripcio)
       Else
        Text28.Text = "0"
        Text29.Text = ""
    End If
  End If
  Unload formseleccio
  
End Sub


Private Sub combo_alcadapalet_Click()
 Dim combo As Object
 Dim vcamp As String
 Set combo = combo_alcadapalet
 vcamp = "alcadapalet"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_cert_qualitat_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_cert_qualitat
 vcamp = "cert_qualitat"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_conosprotectors_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_conosprotectors
 vcamp = "conosprotectors"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_emb_anonim_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_emb_anonim
 vcamp = "emb_anonim"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_guardarmostres_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_guardarmostres
 vcamp = "guardarmostres"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_protecciob_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_protecciob
 vcamp = "tipusprotecciob"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)

End Sub

Private Sub combo_protecciop_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_protecciop
 vcamp = "tipusprotecciop"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_protecciospr_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_protecciospr
 vcamp = "tipusprotecciospr"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_tipuspalet_Click()
 If combo_tipuspalet.ListIndex <> -1 Then envios.Recordset!tipuspalet = combo_tipuspalet.ItemData(combo_tipuspalet.ListIndex)
End Sub

Private Sub Combo1_Click()

End Sub

Private Sub Command3_Click()
  If Not comandesafectades.Visible Then
    comandesafectades.Left = 1725
    comandesafectades.Top = 1275
    comandesafectades.Visible = True
    comandesafectades.Caption = "Buscant comandes un moment sisplau..."
    DoEvents
    ratoli "espera"
    comandes.RecordSource = "select comanda,texteimpressio,puntrisc from comandes where proximaseccio<>'T' and client=" + atrim(cadbl(clients.Recordset!codi)) + " order by comanda DESC"
'    comandes.RecordSource = "select comanda,texteimpressio,'            ' as RISC, puntrisc from comandes "
    comandes.Refresh
    While comandes.Recordset.RecordCount = 1
      DoEvents
    Wend
    reixa.Rows = 500
    i = 0
    reixa.row = 0
    reixa.ColWidth(0) = 100 * 9: reixa.ColWidth(1) = 1000 * 4: reixa.ColWidth(2) = 300 * 4
    reixa.Col = 2
    DoEvents
    reixa.TextMatrix(0, 0) = "COMANDA": reixa.TextMatrix(0, 1) = "TEXTE": reixa.TextMatrix(0, 2) = "RISC"
    While Not comandes.Recordset.EOF
       'If Not comandes.Recordset.EOF Then reixa.Rows = comandes.Recordset.RecordCount+1
       i = comandes.Recordset.AbsolutePosition + 1
       reixa.TextMatrix(i, 1) = atrim(comandes.Recordset!texteimpressio)
       reixa.TextMatrix(i, 0) = atrim(comandes.Recordset!comanda)
       reixa.TextMatrix(i, 2) = IIf(comandes.Recordset!puntrisc = 1, "Vermell", IIf(comandes.Recordset!puntrisc = 2, "Verd", ""))
       reixa.row = i
       reixa.CellBackColor = IIf(reixa.Text = "Vermell", QBColor(12), IIf(reixa.Text = "Verd", QBColor(10), QBColor(15)))
       comandes.Recordset.MoveNext
    Wend
    ratoli "normal"
    comandesafectades.Caption = "Comandes Afectades"
   Else: comandesafectades.Visible = False
  End If

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()
 r = obre_fitxer(ruta_relativa_docs, 2)
 Text41 = Mid(r, Len(ruta_relativa_docs) + 2)
End Sub

Private Sub Command2_Click()
 r = obre_fitxer(ruta_relativa_docs, 2)
 Text42 = Mid(r, Len(ruta_relativa_docs) + 2)
End Sub

Private Sub Command4_Click()
   
   If Not envios.Recordset.EOF Then
     If MsgBox("Segur que vols borrar aquest enviament?", vbCritical + 4, "Atenció") = vbYes Then
      envios.Recordset.Delete
      envios.Refresh
     End If
   End If
End Sub

Private Sub Command5_Click()
  envios.Recordset.AddNew
  envios.Recordset!codi = clients.Recordset!codi
  nomenvio.SetFocus
End Sub

Private Sub Command6_Click()
  If envios.Recordset.EditMode <> 0 Then
     envios.Recordset.Update
  End If
  missatgeenviament.Caption = "": envios.RecordSource = " select * from Clients_envios where codi=" + atrim(cadbl(clients.Recordset!codi))
  envios.Refresh
  If Not envios.Recordset.EOF Then envios.Recordset.MoveLast
End Sub

Private Sub Command7_Click()
If buscant Then Exit Sub
Frame2.Visible = Not Frame2.Visible
palets.Visible = Frame2.Visible
If Not Frame2.Visible Then Exit Sub

envios.RecordSource = " select * from Clients_envios where codi=" + atrim(cadbl(clients.Recordset!codi))
envios.Refresh

If envios.Recordset.EOF Then
     missatgeenviament.Caption = "Enviament únic. DADES GENÈRIQUES"
     Set envios.Recordset = clients.Recordset
     envios.RecordSource = clients.RecordSource
     envios.Tag = cadbl(clients.Recordset!codi) - (cadbl(clients.Recordset!codi) * 2)
   Else:
      missatgeenviament.Caption = ""
      envios.RecordSource = " select * from Clients_envios where codi=" + atrim(cadbl(clients.Recordset!codi))
      envios.Refresh
      
End If
If Frame2.Visible Then
   carregar_combos
   kg.Value = 0: mtrs.Value = 0: unitats.Value = 0: peces.Value = 0
End If
If clients.EditMode > 0 And Not envios.Recordset.EOF Then envios.Recordset.Edit
End Sub
Sub carregar_combos()
   Dim dbcomandes As Database
   Dim rstc As Recordset
   Dim combo As Object
   Dim Combo2 As Object
   Dim Combo3 As Object
   Set dbcomandes = OpenDatabase(llegir_ini("General", "cami", "comandes.ini"))
   'alçades palet
   Set rstc = dbcomandes.OpenRecordset("select distinct familia from productes")
   Set combo = qproducte
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     If atrim(rstc!familia) <> "" Then
        r = atrim(rstc!familia)
        Select Case (atrim(rstc!familia))
          Case "B"
             r = "Bosses"
          Case "F"
             r = "Formats"
        End Select
        combo.AddItem r
     End If
     'combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend

'tipus palets
   Set rstc = dbcomandes.OpenRecordset("tipuspalets")
   Set combo = combo_tipuspalet
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
'alçades palet
   Set rstc = dbcomandes.OpenRecordset("alcadespalets")
   Set combo = combo_alcadapalet
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
 'tipus proteccio
   Set rstc = dbcomandes.OpenRecordset("tipusproteccions")
   Set combo = combo_protecciob: Set Combo2 = combo_protecciop: Set Combo3 = combo_protecciospr
   r = combo.Text: combo.Clear: combo.Text = r
   r = Combo2.Text: Combo2.Clear: Combo2.Text = r
   r = Combo3.Text: Combo3.Clear: Combo3.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio): combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     Combo2.AddItem atrim(rstc!descripcio): Combo2.ItemData(Combo2.NewIndex) = cadbl(rstc!codi)
     Combo3.AddItem atrim(rstc!descripcio): Combo3.ItemData(Combo3.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
'embalatges
'alçades palet
   Set rstc = dbcomandes.OpenRecordset("embalatgesanonims")
   Set combo = combo_emb_anonim
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
'certificat qualitat
'alçades palet
   Set rstc = dbcomandes.OpenRecordset("cert_qualitat")
   Set combo = combo_cert_qualitat
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
'guardarmostres
   'alçades palet
   Set rstc = dbcomandes.OpenRecordset("guardarmostres")
   Set combo = combo_guardarmostres
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
   
'conos protectors
'alçades palet
   Set rstc = dbcomandes.OpenRecordset("conosprotectors")
   Set combo = combo_conosprotectors
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
End Sub

Private Sub Command8_Click()
  r = InputBox("Entra el pes que vols iguals a totes les mides.", "Entrada pes igual")
  If cadbl(r) >= 0 Then
    datapesnet.Recordset.MoveFirst
    While Not datapesnet.Recordset.EOF
       datapesnet.Recordset.Edit
       datapesnet.Recordset!pes = cadbl(r)
       datapesnet.Recordset.Update
       datapesnet.Recordset.MoveNext
    Wend
    datapesnet.Recordset.MoveFirst
  End If
End Sub

Private Sub Command9_Click()
 Dim taulatemp As String
 If clients.Tag = "" Then clients.Tag = " where codi>0"
  taulatemp = "c:\temporal.mdb"
  If existeix(taulatemp) Then Kill taulatemp
  DBEngine.CreateDatabase taulatemp, dbLangGeneral
  dbtmp.Execute ("select * into temporal in '" + taulatemp + "' from clients " + clients.Tag)
 report.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistatclients1.rpt"
 report.DataFiles(0) = taulatemp
 report.Destination = crptToWindow
 report.Action = 1

End Sub

Private Sub consultar_Click()
  buscant = True
  alta_registre
  deixartotblanc
  
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub eliminar_Click()
 On Error GoTo err
  If MsgBox("Segur que vols Eliminar?", vbYesNo + vbCritical, "Atenció") = 6 Then
    clients.Recordset.Delete
    clients.Recordset.MoveNext
    If clients.Recordset.EOF Then clients.Recordset.MovePrevious
  End If
 Exit Sub
err:
  MsgBox "No s'ha pogut eliminar possiblement perque tingui registres relacionats. O bé no hi ha res per eliminar."
End Sub

Private Sub envios_Reposition()
  If buscant Then Exit Sub
  carregar_altres_envio
  If clients.EditMode > 0 And Not envios.Recordset.EOF Then envios.Recordset.Edit
  If missatgeenviament <> "" Then
     envios.Caption = "-----": envios.Enabled = False
    Else:
       If Not envios.Recordset.EOF Then
         envios.Tag = envios.Recordset("id")
        Else: envios.Tag = cadbl(clients.Recordset!codi) - (cadbl(clients.Recordset!codi) * 2)
       End If
       envios.Enabled = True: envios.Caption = Trim(envios.Recordset.AbsolutePosition + 1) + "/" + Trim(envios.Recordset.RecordCount)
  End If
     
End Sub

Sub obreenvios()
Set dbenvios = OpenDatabase(envios.DatabaseName)
End Sub

Sub carregar_altres_envio()
  obreenvios
  If envios.Recordset.EOF Then Exit Sub
  combo_tipuspalet = possar_descripcio("tipuspalets", "descripcio", "codi", cadbl(envios.Recordset!tipuspalet))
  combo_alcadapalet = possar_descripcio("alcadespalets", "descripcio", "codi", cadbl(envios.Recordset!alcadapalet))
  Set dbenvios = Nothing
End Sub
Function possar_descripcio(vtaula As String, vdescripcio As String, vbuscara As String, vvalorbuscat As String)
  Dim rstenvio As Recordset
  Set rstenvio = dbenvios.OpenRecordset("Select " + vdescripcio + " from " + vtaula + " where " + vbuscara + "=" + atrim(cadbl(vvalorbuscat)))
  If Not rstenvio.EOF Then
        possar_descripcio = atrim(rstenvio.Fields(vdescripcio))
     Else: possar_descripcio = ""
  End If
  
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 65 Then alta_registre: KeyCode = 0
If KeyCode = 69 Then buscar_registre
If KeyCode = 27 Then cancelar_registre
If KeyCode = 112 Then gravar_registre
If KeyCode = 13 Then SendKeys "{TAB}": KeyCode = 0

End Sub
Sub buscar_registre()

End Sub
Sub alta_registre()
 If areadatos.Enabled = False Then
      areadatos.Enabled = True
      clients.Recordset.AddNew
      DoEvents
      Text1.Enabled = True
      'busco el mes gran i el poso a codi +1
      If Not buscant Then
        Set rsttmp = dbtmp.OpenRecordset("select max(codi) as [grancodi] from clients")
        If Not rsttmp.EOF Then
          Text1 = atrim(cadbl(rsttmp!grancodi) + 1)
         Else: Text1 = "1"
        End If
      End If
      Text1.SetFocus
 End If
End Sub
Sub gravar_registre()
 If areadatos.Enabled And Not buscant Then
    Command6_Click 'gravar els enviaments
    Frame2.Visible = False
    Text1.Enabled = False
    sortir.SetFocus
    DoEvents
    If Screen.ActiveControl.Name = "sortir" Then
      If envios.Recordset.EditMode > 0 Then envios.Recordset.Update
      If clients.EditMode > 0 Then clients.Recordset.Update
      areadatos.Enabled = False
      clients.Recordset.Bookmark = clients.Recordset.LastModified
    End If
 End If
 If buscant Then finalitzarbusqueda
End Sub
Sub cancelar_registre()
  If clients.Recordset.EditMode > 0 Then
   If envios.Recordset.EditMode > 0 Then envios.Recordset.CancelUpdate
   clients.Recordset.CancelUpdate
   areadatos.Enabled = False
   Text1.Enabled = False
   buscant = False
   carregar_lookups
     Else: Unload Me
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
  If KeyAscii > 50 Then KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub gravar_Click()
gravar_registre
End Sub

Private Sub kg_Click()
If Screen.ActiveControl.Name = "kg" Then gravar_quantitat
End Sub
Sub gravar_quantitat()
If qproducte = "" Then Exit Sub
  
  Set rsttmp = dbtmp.OpenRecordset("select * from unitatsxproducte where idproducte='" + atrim(qproducte) + "' and idenvio=" + atrim(cadbl(envios.Tag)))
  If rsttmp.EOF Then
     rsttmp.AddNew
     rsttmp!idenvio = envios.Tag
     rsttmp!idproducte = qproducte
   Else
     rsttmp.Edit
  End If
    rsttmp!kg = kg.Value
    rsttmp!mtrs = mtrs.Value
    rsttmp!pcs = peces.Value
    rsttmp!unts = unitats.Value
  rsttmp.Update
End Sub

Private Sub modificar_Click()
   areadatos.Enabled = True
   clients.Recordset.Edit
   Text2.SetFocus
End Sub

Private Sub mtrs_Click()
If Screen.ActiveControl.Name = "mtrs" Then gravar_quantitat
End Sub

Private Sub paperfrontal_Click()
 Dim combo As Object
 Dim vcamp As String
 
 If atrim(paperfrontal.Text) = "" Then
     framepaperfrontal.Enabled = False
       Else: framepaperfrontal.Enabled = True
 End If
 
 Set combo = paperfrontal
 vcamp = "pfpaperfrontal"
 'If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.(combo.ListIndex)
End Sub

Private Sub peces_Click()
If Screen.ActiveControl.Name = "peces" Then gravar_quantitat
End Sub

Private Sub qproducte_Click()
  If qproducte.ListIndex > -1 Then
    Set rsttmp = dbtmp.OpenRecordset("select * from unitatsxproducte where idproducte='" + atrim(qproducte) + "' and idenvio=" + atrim(cadbl(envios.Tag)))
    If Not rsttmp.EOF Then
      kg.Value = rsttmp!kg
      mtrs.Value = rsttmp!mtrs
      peces.Value = rsttmp!pcs
      unitats.Value = rsttmp!unts
        Else
         kg.Value = 0
         mtrs.Value = 0
         peces.Value = 0
         unitats.Value = 0
    End If
  End If
End Sub

Private Sub risc_Change()
If risc = "" And Not buscant Then risc = ",00"
End Sub

Private Sub riscpla_Change()
If riscpla = "" And Not buscant Then riscpla = ",00"
End Sub

Private Sub sortir_Click()
 Unload Me
End Sub

Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_LostFocus()
 If Not buscant And clients.Recordset.EditMode > 0 Then
   Set rsttmp = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(Text1.Text)))
   If rsttmp.RecordCount > 0 Then MsgBox "Aquest codi ja existeix haurieu de canviar-lo": If areadatos.Enabled Then Text1.SetFocus
 End If
End Sub

Private Sub Text26_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then triarrepresentant
End Sub

Private Sub Text27_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then triarrepresentant
End Sub

Private Sub Text28_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then triarformapag
End Sub

Private Sub Text29_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then triarformapag
End Sub

Private Sub Text41_Click()
 
 r = "cmd /c "
 If existeix("c:\windows\command\start.exe") Then r = "start "
 r = Shell(r + Chr$(34) + ruta_relativa_docs + "\" + ActiveControl.Text + Chr$(34), vbMinimizedFocus)
End Sub

Private Sub Text42_Click()
 r = "cmd /c "
 If existeix("c:\windows\command\start.exe") Then r = "start "
 r = Shell(r + Chr$(34) + ruta_relativa_docs + "\" + ActiveControl.Text + Chr$(34), vbMinimizedFocus)
End Sub

Private Sub Text44_LostFocus()
 If Text44 = "" Then Text44.Text = "0"
End Sub

Private Sub Timer1_Timer()
  estattaula.Caption = textestattaula(clients.EditMode)
  If estattaula.ForeColor <> QBColor(0) Then
     estattaula.ForeColor = QBColor(0)
    Else: estattaula.ForeColor = QBColor(14)
  End If
End Sub


Sub recorregutregistres()
 Dim objecte As Object
 Dim protegir As Boolean
 protegir = IIf(llegir_ini("general", "protegircamps", "comandes.ini") = "si", True, False)
 If Not protegir Then escriure_ini "general", "protegircamps", "no", "comandes.ini"
 queryorder = ""
 querywhere = ""
 'On Error Resume Next
 For Each objecte In Me
    If TypeOf objecte Is TextBox Then
      If objecte.DataField <> "" Then ' Si Texto es igual "Hola".
        If objecte.Text <> "" Then evaluarcontingut objecte.DataField, objecte.Text, clients.Recordset.Fields(objecte.DataField).Type
     End If
     
    End If
Next

End Sub
Sub colocarbloqueig()
 Dim objecte As Object
 Dim protegir As Boolean
 protegir = IIf(llegir_ini("general", "protegircamps", "comandes.ini") = "no", False, True)
 If protegir Then escriure_ini "general", "protegircamps", "si", "comandes.ini"
 
 queryorder = ""
 querywhere = ""
 'On Error Resume Next
 For Each objecte In Me
    If TypeOf objecte Is TextBox Then
     If objecte.Tag = "protegits" Then
        If protegir Then
          objecte.Locked = True
         Else
          objecte.BackColor = QBColor(15)
        End If
     End If
     End If
     
   
Next

End Sub

Function evaluarcontingut(camp As String, valor As String, tipusdato As Byte) As String
  Dim rest As String
  rest = ""
  evaluarcontingut = ""
  If triarordre(camp, valor) Then Exit Function
  If tipusdato = 10 Then
   If InStr(1, valor, "*") Or InStr(1, valor, "?") Then
      rest = " like '" + valor + "'"
     Else
       If InStr(1, valor, ">") Or InStr(1, valor, "<") Or InStr(1, valor, "=") Then
           rest = "='" + valor + "'"
        Else: rest = "=" + "'" + IIf(valor = " ", "", valor) + "'"
       End If
   End If
  End If
  If tipusdato <> 10 Then
    If InStr(1, valor, ">") Or InStr(1, valor, "<") Or InStr(1, valor, "=") Then
           rest = atrim(cadbl(valor))
        Else: rest = "=" + atrim(cadbl(valor))
    End If
  End If
  rest = camp + rest
  evaluarcontingut = rest
  
  If querywhere = "" Then
     querywhere = rest
    Else
     querywhere = querywhere + " and " + rest + " "
  End If
End Function

Function triarordre(camp As String, valorord As String) As Boolean
  Dim ord As String
  triarordre = False
  If InStr(1, valorord, "<<") Then ord = camp + " " + " ASC"
  If InStr(1, valorord, ">>") Then ord = camp + " " + " DESC"
  If ord <> "" Then
      triarordre = True
    Else: Exit Function
  End If
  If queryorder = "" Then
     queryorder = ord
   Else: queryorder = queryorder + ", " + ord
  End If
  
End Function
Sub finalitzarbusqueda()
 ratoli "espera"
 recorregutregistres
 If clients.Recordset.EditMode > 0 Then clients.Recordset.CancelUpdate
 buscant = False
 Text1.Enabled = True
 areadatos.Enabled = False
 If queryorder <> "" Then queryorder = " Order By " + queryorder
 If querywhere <> "" Then querywhere = " Where " + querywhere
 clients.RecordSource = "select * from clients " + querywhere + queryorder
 clients.Tag = querywhere + queryorder
 clients.Refresh
 If Not clients.Recordset.EOF Then clients.Recordset.MoveLast
 ratoli "normal"
End Sub

Sub deixartotblanc()
Frame2.Visible = False
palets.Visible = False
 For Each objecte In Me
    If TypeOf objecte Is TextBox Then
      If objecte.DataField <> "" Then ' Si Texto es igual "Hola".
        objecte.Text = ""
     End If
    End If
Next



End Sub

Sub carregar_lookups()
 Frame2.Visible = False: palets.Visible = False
 
 If clients.Recordset.EOF And clients.Recordset.BOF Then Exit Sub
 ' carrego els envios
 'If envios.RecordSource <> clients.RecordSource Then
  envios.RecordSource = " select * from Clients_envios where codi=" + atrim(cadbl(clients.Recordset!codi))
 'End If
 envios.Refresh
If Not envios.Recordset.EOF Then envios.Recordset.MoveLast
 'LOOKUP DE REPRESENTANT
  Set rsttmp = dbtmp.OpenRecordset("select nom from representants where codi=" + atrim(cadbl(clients.Recordset!representant)))
  If Not rsttmp.EOF Then
     Text27 = rsttmp!nom
    Else: Text27 = ""
  End If
  
  
  'LOOKUP DE formade pag
  Set rsttmp = dbtmp.OpenRecordset("select descripcio from [formes de pagament] where codi='" + atrim((clients.Recordset!formapag)) + "'")
  If Not rsttmp.EOF Then
     Text29 = rsttmp!descripcio
    Else: Text29 = ""
  End If
  
  Set rsttmp = Nothing
End Sub
Sub possarvalordcamps()
On Error Resume Next
 For Each objecte In Me
    If TypeOf objecte Is TextBox Then
      If objecte.DataField <> "" Then
         objecte.MaxLength = clients.Recordset.Fields(objecte.DataField).Size
      End If
    End If
Next

End Sub

Private Sub unitats_Click()
If Screen.ActiveControl.Name = "unitats" Then gravar_quantitat
End Sub

