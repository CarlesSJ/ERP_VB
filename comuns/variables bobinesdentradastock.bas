Attribute VB_Name = "varialbesbobentrada"
Global espesorbobina As Double
Global lletraseccio As String
Global ncomanda As Double
Global ncomanda2 As Double
Global explicacio As String
Global PoB As String
Global vtipusimpresio As String
Sub netejarreport(rpt As CrystalReport)
  Dim i As Byte
  rpt.ReportFileName = ""
  
  For i = 1 To 20
     rpt.Formulas(i) = ""
     rpt.DataFiles(i) = ""
  Next i
  
End Sub


Sub estatdelabobina(palet As Double, bobina As Double, grup As Double, ByVal comanda As Double, Optional ByVal comanda2 As Double)
   Dim rstb As Recordset
   Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "Palets.mdb")
   Load mantenimentbobina
    If grup = 0 Then
      Set rstb = dbstocks.OpenRecordset("select orcomassignacio from parcials where idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)) + " and (cdbl(orcomassignacio)<3000 and cdbl(orcomassignacio)>1999)")
      If Not rstb.EOF Then grup = rstb!orcomassignacio
    End If
    If comanda2 > 0 And grup = 0 Then
     Set rstb = dbstocks.OpenRecordset("select * from parcials where idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)) + " and (cdbl(comanda)=" + atrim(comanda) + " or cdbl(comanda)=" + atrim(comanda2) + ")")
     If Not rstb.EOF Then
        comanda = cadbl(rstb!comanda)
       Else: MsgBox "No s'ha trobat aquest palet assignat.": Exit Sub
     End If
    End If
   mantenimentbobina.palet = atrim(palet)
   mantenimentbobina.bobina = atrim(bobina)
   mantenimentbobina.grup = atrim(grup)
   mantenimentbobina.comanda = atrim(comanda)
   If seccioanterior(comanda) <> "E" Then mantenimentbobina.comanda = comanda2
   mantenimentbobina.Show 1
    Set rstb = Nothing
End Sub
Sub demanar_final_palet_bobina_stock(palet As Double, bobina As Double)
 Dim metres As Double
 Dim metresant As Double

'es una bobina d'estock
         metres = cadbl(ncomanda)
         'carregar_bobinesdentrada "metresbobinadisponible", , palet, bobina, metres
         metres = bobinesdentrada.calcular_mtrsdispreals(palet, bobina)
         metresant = metres
         metres = cadbl(InputBox("La bobina " + atrim(palet) + "/" + atrim(bobina) + " tenia " + atrim(metres) + " Mtrs." + Chr(10) + Chr(13) + " Quants metres has gastat?", "Bobina no acabada"))
         If (metresant - metres) < 500 Then
              If (metresant - metres) < 500 Then MsgBox "Bobines de menys de 500 metres es donen per gastades.", vbInformation, "Atenció"
             'If MsgBox("Has entrat  mes Metres dels que quedaven disponibles." + Chr(10) + Chr(13) + " La donu com a gastada total?", vbExclamation + vbYesNo, "Atenció") = vbYes Then
                carregar_bobinesdentrada "metresbobinaassignar", metresant, palet, bobina, ncomanda, , ncomanda2
                carregar_bobinesdentrada "marcarutilitzada", , palet, bobina, ncomanda, True, ncomanda2
             'End If
            Else:
              carregar_bobinesdentrada "metresbobinaassignar", metres, palet, bobina, ncomanda
              carregar_bobinesdentrada "marcarutilitzada", , palet, bobina, ncomanda, True, ncomanda2
              If bobinesdentrada.calcular_mtrsdispreals(palet, bobina) Then carregar_bobinesdentrada "imprimirbobina", , palet, bobina
         End If
End Sub

Sub demanar_paletibobina(palet As Double, bobina As Double, Optional desb As Byte)
  Unload entradabobina
  If lletraseccio <> "L" Then
   entradabobina.etdesb.visible = False
   entradabobina.desb.visible = False
    Else: entradabobina.desb = "2"
  End If
  Load entradabobina
  'entradabobina.palet.SetFocus
  entradabobina.Show 1
  If cadbl(entradabobina.palet) > 0 Then
     palet = cadbl(entradabobina.palet)
     bobina = cadbl(entradabobina.bobina)
     desb = cadbl(entradabobina.desb)
     
  End If
  Unload entradabobina
End Sub
Sub enviaremailsishaentratmanualment(venviatdesdemanual As Boolean, palet As Double, bobina As Double)
   Dim venviar As Boolean
   Dim dbcomandes As Database
   Dim rst As Recordset
   Dim cos As String
   
   cos = Chr(13) + Chr(10) + "Operari: " + atrim(numop) + Chr(13) + Chr(10) + "Seccio: " + lletraseccio + Chr(13) + Chr(10) + "Palet: " + atrim(palet) + "/" + atrim(bobina)
   Set dbcomandes = DBEngine.OpenDatabase(rutadelfitxer(cami) + "comandes.mdb", , True)
   Set rst = dbcomandes.OpenRecordset("SELECT  comandes.comanda, clients.codi, clients.nom, comandes.refclient, comandes.marcailinia FROM clients INNER JOIN comandes ON clients.codi = comandes.client where comanda=" + atrim(ncomanda))
   If Not rst.EOF Then cos = Chr(13) + Chr(10) + "Codi Client: " + atrim(rst!codi) + " - " + atrim(rst!nom) + Chr(13) + Chr(10) + "Ref.Client: " + atrim(rst!refclient) + Chr(10) + Chr(13) + "Texte Imp.: " + atrim(rst!marcailinia)
   
   If lletraseccio = "R" Then GoTo fi
   If Form1.botoensenyarpacking.tag = "afegidamanualmentcaixa" Then venviar = True
   If Form1.botoensenyarpacking.tag = "afegidamanualment" Then venviar = True
   cos = Chr(13) + Chr(10) + "Seccio: " + atrim(lletraseccio) + Chr(13) + Chr(10) + "Operari: " + atrim(numop) + Chr(13) + Chr(10) + "Màquina: " + atrim(nummaq) + Chr(13) + Chr(10) + cos
   If venviar Then enviaremailgeneric "Avisbobinesentradessenseescanejar", "Bobina entrada manualment Lot:" + atrim(ncomanda), cos
fi:
   Form1.botoensenyarpacking.tag = ""
   Set rst = Nothing
   Set dbcomandes = Nothing
End Sub
Sub marcaranteriorscomagastades(Optional primer_proces As Boolean)
   Dim bobinesent As Control
   Dim rstb As Recordset
   Dim palet As Double
   Dim bobina As Double
   Dim utilitzada As Boolean
   Set bobinesent = Form1.bobinesent
   bobinesent.Refresh
   utilitzada = True
   ratoli "espera"
   While Not bobinesent.Recordset.EOF
         If lletraseccio = "L" Then If desb <> bobinesent.Recordset!desb Then GoTo cont
         If seccioanterior(ncomanda) = "E" Then PoB = "P"
         If lletraseccio = "L" Then
            PoB = UCase(atrim(bobinesent.Recordset!paletobobina))
          ' Else: PoB = "P"
         End If
         palet = cadbl(bobinesent.Recordset!palet)
         bobina = cadbl(bobinesent.Recordset!bobina)
         'aquesta linia es un apanyu per saber si es palet o bobina s'ha darreglar
'         If bobinesent.Recordset!palet > 120000 Then PoB = "B"
         'carregar_bobinesdentrada "marcarutilitzada", , cadbl(bobinesent.Recordset!palet), cadbl(bobinesent.Recordset!bobina), ncomanda, True, ncomanda2
         If PoB = "P" Then
            ' If Not esdestoc(palet, bobina) Then
                carregar_bobinesdentrada "mirarsiutilitzada", , cadbl(bobinesent.Recordset!palet), (bobinesent.Recordset!bobina), ncomanda, utilitzada, ncomanda2, primer_proces
                  'Else
         '           metresreals = bobinesdentrada.calcular_mtrsdispreals(palet, bobina)
         '           If metresreals > 0 Then utilitzada = False
                    
         '    End If
             
             If Not utilitzada Then estatdelabobina cadbl(bobinesent.Recordset!palet), cadbl(bobinesent.Recordset!bobina), 0, ncomanda, ncomanda2
           Else: If PoB = "B" Then carregar_bobinesdentrada "marcarutilitzada", , cadbl(bobinesent.Recordset!palet), (bobinesent.Recordset!bobina), ncomanda, utilitzada, ncomanda2, primer_proces
         End If
         wait 1
         'bobinesdentrada.imprimir_bobinaparcial cadbl(bobinesent.Recordset!palet), cadbl(bobinesent.Recordset!bobina)
cont:
         bobinesent.Recordset.MoveNext
   Wend
   ratoli "normal"
End Sub
Function esdestoc(palet As Double, bobina As Double) As Boolean
   Dim rstesd As Recordset
   Set rstesd = dbstocks.OpenRecordset("select * from parcials where (cdbl(orcomassignacio)<3000 and cdbl(orcomassignacio)>2000) and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
   If Not rstesd.EOF Then
      esdestoc = True
     Else: esdestoc = False
   End If
End Function
Sub afegir_labobinadentrada(palet As Double, bobina As Double, Optional desb As Byte, Optional proces_invertit As Boolean)
  Dim jaexisteix As Boolean
  Dim utilitzada As Boolean
  Dim bobinesent As Control
  Dim bobines As Control
  Set bobinesent = Form1.bobinesent
  Set bobines = Form1.bobines
  
       bobinesent.Refresh
       
       While Not bobinesent.Recordset.EOF
         If cadbl(bobinesent.Recordset!palet) = palet And cadbl(bobinesent.Recordset!bobina) = bobina Then jaexisteix = True
         bobinesent.Recordset.MoveNext
       Wend
       If Not jaexisteix Then
        If lletraseccio = "R" Or lletraseccio = "I" Or lletraseccio = "L" Then marcaranteriorscomagastades proces_invertit
        ratoli "espera"
        bobinesent.Recordset.AddNew
        bobinesent.Recordset!id = bobines.Recordset!id
        If lletraseccio = "L" Then
           PoB = IIf(palet < 120000, "p", "b")
           bobinesent.Recordset!desb = IIf(desb = 1 Or desb = 2, desb, 1)
           'carregar_bobinesdentrada "marcarutilitzada", , palet, bobina, ncomanda, utilitzada, ncomanda2
           'If utilitzada Then PoB = UCase(PoB)
           bobinesent.Recordset!paletobobina = PoB
        End If
        bobinesent.Recordset!palet = palet
        bobinesent.Recordset!bobina = bobina
        bobinesent.Recordset.Update
        bobinesent.Refresh
        bobinesent.Recordset.Bookmark = bobinesent.Recordset.LastModified
        bobinesent.UpdateControls
        enviaremailsishaentratmanualment False, palet, bobina
          Else
            MsgBox "Aquesta bobina ja està afegida", vbCritical, "Repeticio de bobina"
            palet = 0  'posso a 0 per controlar que ja esta afegida quan surti del procediment
       End If
    ratoli "normal"
End Sub
Sub assignar_dbbaixes(dbdebaixes As Database)
  If estaobertalabd(dbdebaixes) Then If InStr(1, LCase(dbdebaixes.Name), "baixes.mdb") Then Exit Sub
   If InStr(1, LCase(dbtmpb.Name), "baixes.mdb") > 0 Then Set dbbaixes = dbtmpb
   If InStr(1, LCase(dbtmp.Name), "baixes.mdb") > 0 Then Set dbbaixes = dbtmp
   If InStr(1, LCase(dbbaixes.Name), "baixes.mdb") > 0 Then Set dbdebaixes = dbbaixes
End Sub
Function seccioanterior(numc As Double) As String
   Dim rsttemp As Recordset
   Dim seccio As String
   Set rsttemp = dbtmp.OpenRecordset("SELECT comandes.comanda, productes.ruta as rutap ,linkcomanda2 FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(numc) + "));")
   seccio = "E"
   If Not rsttemp.EOF Then
      If Len(rsttemp!rutap) > 1 Then
        seccio = Mid(rsttemp!rutap, InStr(1, rsttemp!rutap, lletraseccio) - 1, 1)
      End If
   End If
   seccioanterior = seccio
End Function
Public Sub carregar_bobinesdentrada(Optional funcio As String, Optional multiselect As Double, Optional palet As Double, Optional bobina As Double, Optional numcomanda As Double, Optional utilitzada As Boolean, Optional numcomanda2 As Double, Optional laminadoraprocesinvertit As Boolean, Optional vnodeixarfercheckaacceptarbobines As Boolean, Optional notancarform As Boolean)
    Dim cont As Integer
    Dim rsttemp As Recordset
    Dim rstb As Recordset
    Dim noutilitzades As String
    Dim nomtaula As String
    Dim nomtaula2 As String
    Dim seccio As String
    Dim seccio2 As String
    Dim dbbaixes As Database
    Dim numc As Double
    Dim orcomassignacio As String
    Dim mtrsb As Double
    Dim comandaassignacio As Double
    Dim rr As String
    Dim selecciocomandes As String
    Dim linkcomanda2 As Double
    Dim jahepassatperbobinestransformades As Boolean
    Unload bobinesdentrada
    'ensenyar
    'marcarutilitzada si es s'afageig la paraula demanar farà una pregunta abans de fer-ho
    'mirarsiutilitzada  retorna a la variable utilitzada si està utilitzada o no
    'metresbobina    es retornen els metres a numcomanda   si hi ha assignar el valor de multiselect  sagafarà com a metres a possar a aquesta bobina   si s'afageig disponible son els metres disponibles de la bobina jumbo
    'carregarbobines  carrega les bobines i surt sense ensenyar res noutilitzades o siutilitzades
    'imprimirbobina
    If bobinesdentrada.visible = True Then notancarform = True
    assignar_dbbaixes dbbaixes
    
    
    comandaassignacio = numcomanda
    
    ' seccio de la comanda primera
    Set rsttemp = dbtmp.OpenRecordset("SELECT comandes.refilatd,comandes.comanda, productes.ruta as rutap ,linkcomanda2 FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(numcomanda) + "));")
    If Not rsttemp.EOF Then
      If Len(rsttemp!rutap) > 1 Then
        seccio = Mid(rsttemp!rutap, InStr(1, rsttemp!rutap, lletraseccio) - 1, 1)
        If seccio = "R" Then nomtaula = "bobinesreb"
        If seccio = "I" Then nomtaula = "bobinesimp"
        If seccio = "L" Then nomtaula = "bobineslam"
        If seccio = "I" And ncomanda2 = ncomanda + 2 And Not laminadoraprocesinvertit Then nomtaula = "bobineslam": seccio = "L"
        linkcomanda2 = cadbl(rsttemp!linkcomanda2)
        'If cadbl(rsttemp!linkcomanda2) > 0 And lletraseccio = "R" Then comandaassignacio = rsttemp!linkcomanda2
      End If
    End If
    
    ' seccio de la comanda segona
    Set rsttemp = dbtmp.OpenRecordset("SELECT comandes.comanda, productes.ruta as rutap ,linkcomanda2 FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(numcomanda2) + "));")
    If Not rsttemp.EOF Then
      If InStr(1, rsttemp!rutap, lletraseccio) > 1 Then
        seccio2 = Mid(rsttemp!rutap, InStr(1, rsttemp!rutap, lletraseccio) - 1, 1)
        If seccio2 = "I" Then nomtaula2 = "bobinesimp"
        If seccio2 = "L" Then nomtaula2 = "bobineslam"
        If ncomanda2 = ncomanda + 2 Then nomtaula2 = "bobineslam": seccio2 = "L"
        'If cadbl(rsttemp!linkcomanda2) > 0 And lletraseccio = "R" Then comandaassignacio = rsttemp!linkcomanda2
      End If
    End If
    
    
    If funcio = "" Then funcio = "ensenyar"
    If Not notancarform Then
      Unload bobinesdentrada
      Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "Palets.mdb")
    End If
    
    If funcio = "imprimirbobina" Then
       Set rstconsulta = dbtemp.OpenRecordset("select * from selecciobobentrada ")
       bobinesdentrada.imprimir_bobinaparcial palet, bobina, , 1
       Unload bobinesdentrada
       Exit Sub
    End If
    
    'miro si la bobina està utilitzada i ho retorno a utilitzada
    If InStr(1, funcio, "mirarsiutilitzada") > 0 And palet > 0 And bobina > 0 And numcomanda > 0 Then
     If palet < (numcomanda - 5) Then
      'faig una busqueda de parcials per si enlloc de comanda es numero de grup els <10000
      'canvio el numcomanda pel numcomanda2 perque en els complexes la segona
         ' proces porta l'estock
      numc = numcomanda
      If numcomanda2 > 0 Then numc = numcomanda2
       Set rsttemp = dbstocks.OpenRecordset("select comanda from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and comanda='" + atrim(numc) + "'")
       If rsttemp.EOF Then
         Set rsttemp = dbstocks.OpenRecordset("select orcomassignacio from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
         If Not rsttemp.EOF Then numc = rsttemp!orcomassignacio
             'aqui he canviat comanda per orcomassignació i al select també
                '12/12/21  he tornat a posar orcomassignació per problemes amb la detecció de la bobina
       End If
       Set rsttemp = dbstocks.OpenRecordset("select utilitzada from parcials  where orcomassignacio<>'500' and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and comanda='" + atrim(numc) + "'")
       If Not rsttemp.EOF Then utilitzada = rsttemp!utilitzada
         Else
           'ara es una bobina feta a inplacsa
             'miro si es de impresores o laminadores i la marcu
           If nomtaula = "bobinesimp" Then instsql = "SELECT impressores.comanda, bobinesimp.numerodebobina, bobinesimp.utilitzadaabaixa as utilitzada FROM impressores INNER JOIN bobinesimp ON impressores.Id = bobinesimp.controlid WHERE (((impressores.comanda)=" + atrim(numcomanda) + ") AND ((bobinesimp.numerodebobina)=" + atrim(bobina) + "));"
           If nomtaula = "bobineslam" Then
               numc = numcomanda
               If comandaassignacio = numc + 2 Then numc = comandaassignacio
               If lletraseccio = "R" Then
                  numc = numcomanda
                  If numcomanda2 - numcomanda > 1 Then
                   If laminadoraprocesinvertit Then
                        numc = numcomanda
                       Else: numc = numcomanda2
                   End If
                  End If
               End If
               instsql = "SELECT laminadores.comanda, bobineslam.numerodebobina, bobineslam.utilitzadaabaixa as utilitzada FROM laminadores INNER JOIN bobineslam ON laminadores.Id = bobineslam.controlid WHERE (((laminadores.comanda)=" + atrim(numc) + ") AND ((bobineslam.numerodebobina)=" + atrim(bobina) + "));"
           End If
           If instsql <> "" Then
            Set rstttmp = dbbaixes.OpenRecordset(instsql)
            If Not rstttmp.EOF Then utilitzada = rstttmp!utilitzada
           End If
      End If
      Exit Sub
    End If
    
    
    If InStr(1, funcio, "marcarutilitzada") > 0 And palet > 0 And bobina > 0 And numcomanda > 0 Then
      If InStr(1, funcio, "demanar") > 0 Then
         If MsgBox("Es final de bobina?" + Chr(10) + Chr(13) + " Bobina: " + atrim(palet) + "/" + atrim(bobina), vbQuestion + vbYesNo, "Bobines") = vbYes Then
            utilitzada = True
           Else: utilitzada = False
         End If
      End If
      'si el palet es mes petit que el numerodecomanda (em donu 5 de marge amb 2 n'hi hauria prou) es un palet
      If palet < (numcomanda - 5) Then
       'faig una busqueda de parcials per si enlloc de comanda es numero de grup els <10000
       numc = numcomanda
       If lletraseccio = "R" Then comassignacio = numcomanda
       If numcomanda2 > 0 Then numc2 = numcomanda2
       orcomassignacio = numc
       Set rsttemp = dbstocks.OpenRecordset("select comanda,orcomassignacio from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and (comanda='" + atrim(numcomanda2) + "' or comanda='" + atrim(numcomanda) + "')")
       If rsttemp.EOF Then
         Set rsttemp = dbstocks.OpenRecordset("select comanda from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
         If Not rsttemp.EOF Then
            numc = cadbl(rsttemp!comanda)
            If numc = 0 Then numc = numc2
         End If
          Else
            orcomassignacio = cadbl(rsttemp!orcomassignacio)
            If orcomassignacio = 0 Then orcomassignacio = numcomanda2
            numc = rsttemp!comanda
       End If
       If utilitzada Then
          Set rsttemp = dbstocks.OpenRecordset("select * from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and (comanda='" + atrim(numcomanda2) + "' or comanda='" + atrim(numcomanda) + "')")
          If Not rsttemp.EOF Then
           comandaassignacio = rsttemp!comanda
           numc = comandaassignacio
           If comandaassignacio < 10000 Then
              numc = comandaassignacio
              comandaassignacio = numcomanda
             If numcomanda2 > 0 Then comandaassignacio = numcomanda2
           End If
           If Not rsttemp!utilitzada Then
            dbstocks.Execute "update parcials set seccio='" + lletraseccio + "',operari=" + atrim(cadbl(numop)) + ",comanda='" + atrim(comandaassignacio) + "',data=now,utilitzada=true where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and comanda='" + atrim(numc) + "'"
           End If
           If orcomassignacio < 10000 Then
             mtrsb = bobinesdentrada.calcular_mtrsdispreals(palet, bobina)
             If mtrsb > 0 Then dbstocks.Execute "insert into parcials (idpalet,idbobina,metres,comanda,orcomassignacio) values (" + atrim(palet) + "," + atrim(bobina) + "," + atrim(mtrsb) + "," + atrim(numc) + "," + atrim(orcomassignacio) + ")"
           End If
          End If
           Else: dbstocks.Execute "update parcials set seccio='',operari=0,data=null,utilitzada=false,comanda='" + orcomassignacio + "' where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and comanda='" + atrim(numc) + "'"
       End If
         Else
           'ara es una bobina feta a inplacsa
             'miro si es de impresores o laminadores i la marcu
           If lletraseccio <> "R" Then comandaassignacio = palet
           Set rsttmp = dbtmp.OpenRecordset("select producte from comandes where comanda=" + atrim(palet))
           If Not rsttmp.EOF Then
              nomtaula = "bobinesimp"
              If rsttmp!producte = "PC2" Then
                 nomtaula = "bobineslam"
                   Else
                     If Not laminadoraprocesinvertit And (numcomanda2 - numcomanda) = 2 And lletraseccio = "L" Then comandaassignacio = numcomanda: nomtaula = "bobineslam"
                     If Not laminadoraprocesinvertit And (numcomanda2 - numcomanda) = 2 And lletraseccio = "R" Then comandaassignacio = numcomanda2: nomtaula = "bobineslam"
                     If laminadoraprocesinvertit And (numcomanda2 - numcomanda) = 2 Then comandaassignacio = numcomanda: nomtaula = "bobineslam"
                     If (numcomanda2 - numcomanda) = 1 And laminadoraprocesinvertit Then comandaassignacio = numcomanda: nomtaula = "bobineslam"
              End If
continuareb:
              If numcomanda2 - numcomanda = 1 And lletraseccio = "R" Then nomtaula = "bobineslam"
           End If
           If nomtaula = "bobinesimp" Then instsql = "UPDATE impressores INNER JOIN bobinesimp ON impressores.Id = bobinesimp.controlid SET bobinesimp.utilitzadaabaixa = " + IIf(utilitzada, "True", "False") + " WHERE (((impressores.comanda)=" + atrim(comandaassignacio) + ") AND ((bobinesimp.numerodebobina)=" + atrim(bobina) + "));"
           If nomtaula = "bobineslam" Then instsql = "UPDATE laminadores INNER JOIN bobineslam ON laminadores.Id = bobineslam.controlid SET bobineslam.utilitzadaabaixa = " + IIf(utilitzada, "True", "False") + " WHERE (((laminadores.comanda)=" + atrim(comandaassignacio) + ") AND ((bobineslam.numerodebobina)=" + atrim(bobina) + "));"
           
           If nomtaula <> "" Then dbbaixes.Execute instsql
      End If
      Exit Sub
    End If
    
    If InStr(1, funcio, "metresbobina") > 0 Then
     If InStr(1, funcio, "assignar") = 0 Then
      If InStr(1, funcio, "disponible") = 0 Then
       Set rsttemp = dbstocks.OpenRecordset("select metres from parcials where comanda='" + atrim(numcomanda) + "' and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
       If Not rsttemp.EOF Then numcomanda = rsttemp!metres
        Else
           Set rsttemp = dbstocks.OpenRecordset("select sum(metres) as disponibles from parcials where  utilitzada and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
           numcomanda = 0
           If Not rsttemp.EOF Then
             numcomanda = cadbl(rsttemp!disponibles)
             Set rsttemp = dbstocks.OpenRecordset("select mts  from bobines where  idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
             If Not rsttemp.EOF Then numcomanda = rsttemp!mts - numcomanda
           End If
      End If
       Set rsttemp = Nothing
         Else:
             dbstocks.Execute "update parcials set seccio='',operari=0,data=null,utilitzada=false,metres=" + atrim(multiselect) + " where (comanda='" + atrim(numcomanda) + "' or comanda='" + atrim(numcomanda2) + "') and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina)
             bobinesdentrada.actualitzar_metres_disponibles palet, bobina
     End If
     Exit Sub
    End If
    
    If estaobertalabd(dbtemp) Then
        If Not notancarform Then
          dbtemp.Execute "delete * from selecciobobentrada"
        End If
     Else: Exit Sub
    End If
    
    'carrego les bobines d'estock
    numc = numcomanda
carregarstock:
    If InStr(1, funcio, "noutilitzades") > 0 Then noutilitzades = " and utilitzada=false  "
    If InStr(1, funcio, "siutilitzades") > 0 Then noutilitzades = " and utilitzadaabaixa=true "
    If (lletraseccio = "S" Or lletraseccio = "R") And seccio <> "E" Then GoTo bobinestransformades
    If laminadoraprocesinvertit And (numcomanda2 - numcomanda) > 1 And seccio <> "E" Then GoTo bobinestransformades
    If (numcomanda2 - numcomanda) = 1 And laminadoraprocesinvertit Then selecciocomandes = " and (comanda='" + atrim(numcomanda) + "' or comanda='" + atrim(numcomanda2) + "')"
    If (numcomanda2 - numcomanda) = 1 And Not laminadoraprocesinvertit Then
           If seccio = "E" Then
              selecciocomandes = " and (comanda='" + atrim(numcomanda2) + "' or comanda='" + atrim(numcomanda) + "')"
             Else:
               If seccio2 <> "I" Then selecciocomandes = " and comanda='" + atrim(numcomanda2) + "'"
           End If
    End If
    If (numcomanda2 - numcomanda) = 2 And Not laminadoraprocesinvertit Then selecciocomandes = " and comanda='" + atrim(numcomanda2) + "'"
    If (numcomanda2 - numcomanda) = 2 And laminadoraprocesinvertit Then selecciocomandes = " and comanda='" + atrim(numcomanda) + "'"
    If selecciocomandes = "" Then selecciocomandes = " and comanda='" + atrim(numcomanda) + "' "
    ' IIf(seccio = "I" Or seccio = "L", "", " or comanda='" + atrim(numcomanda) + "'") + ") "
    Set rstbobinesdentrada = dbstocks.OpenRecordset("select * from parcials where orcomassignacio<>'500' " + selecciocomandes + noutilitzades + " order by idpalet,idbobina")
    If Not rstbobinesdentrada.EOF Then Set rstb = dbstocks.OpenRecordset("select * from palets where idpalet=" + atrim(cadbl(rstbobinesdentrada!idpalet)))
    Set rstconsulta = dbtemp.OpenRecordset("select * from selecciobobentrada ")
    While Not rstbobinesdentrada.EOF
        rstconsulta.AddNew
        rstconsulta!idpalet = rstbobinesdentrada!idpalet
        rstconsulta!idbobina = rstbobinesdentrada!idbobina
        rstconsulta!utilitzada = rstbobinesdentrada!utilitzada
        rstconsulta!metres = rstbobinesdentrada!metres
        rstconsulta!tipus = IIf(bobinesdentrada.esrestu(rstbobinesdentrada!idpalet, rstbobinesdentrada!idbobina), "R", "")
        rstconsulta!tipus = IIf(bobinesdentrada.esparcial(rstbobinesdentrada!idpalet, rstbobinesdentrada!idbobina), "P", rstconsulta!tipus)
        rstconsulta!tipus = IIf(bobinesdentrada.calcular_mtrsdispreals(rstbobinesdentrada!idpalet, rstbobinesdentrada!idbobina) = 0, "Z", rstconsulta!tipus)
        If rstconsulta!tipus = "" Or IsNull(rstconsulta!tipus) Then rstconsulta!tipus = "O"
        rstconsulta!taula = "parcials"
        If Not rstb.EOF Then espesorbobina = cadbl(rstb!micres)
        rstconsulta.Update
        rstbobinesdentrada.MoveNext
        cont = cont + 1
    Wend
    'If numcomanda2 > 0 And numc <> numcomanda2 Then numc = numcomanda2: GoTo carregarstock
bobinestransformades:
   'carrego les bobinestransformades o sigui les que tenen el numero de comanda
   If (seccio = "E" Or seccio = "") And seccio2 <> "I" Then GoTo fibobinestransformades
    If InStr(1, funcio, "noutilitzades") > 0 Then noutilitzades = " and utilitzadaabaixa=false"
    If InStr(1, funcio, "siutilitzades") > 0 Then noutilitzades = " and utilitzadaabaixa=true"
    rr = atrim(numc)
    If seccio = "I" Then Set rstbobinesdentrada = dbtmpb.OpenRecordset("SELECT impressores.comanda, bobinesimp.* FROM impressores INNER JOIN bobinesimp ON impressores.Id = bobinesimp.controlid WHERE (impressores.comanda)=" + atrim(comandaassignacio) + " " + noutilitzades + " order by numerodebobina")
    If seccio2 = "I" Then Set rstbobinesdentrada = dbtmpb.OpenRecordset("SELECT impressores.comanda, bobinesimp.* FROM impressores INNER JOIN bobinesimp ON impressores.Id = bobinesimp.controlid WHERE (impressores.comanda)=" + atrim(numcomanda2) + " " + noutilitzades + " order by numerodebobina"): rr = atrim(numcomanda2)
    If seccio = "R" Then Set rstbobinesdentrada = dbtmpb.OpenRecordset("SELECT rebobinadores.comanda, bobinesreb.* FROM rebobinadores INNER JOIN bobinesreb ON rebobinadores.Id = bobinesreb.controlid WHERE (rebobinadores.comanda)=" + atrim(comandaassignacio) + " " + noutilitzades + " order by numerodebobina")
    If seccio = "L" Then
      'rr = ""
      'If numcomanda + 2 = numcomanda2 Or numcomanda2 + 2 = numcomanda Then
      '      rr = " or laminadores.comanda=" + atrim(comandaassignacio) 'atrim(IIf(numcomanda2 > numcomanda, numcomanda2, numcomanda))
      'End If
      'MsgBox "SELECT laminadores.comanda, bobineslam.* FROM laminadores INNER JOIN bobineslam ON laminadores.Id = bobineslam.controlid WHERE (laminadores.comanda=" + atrim(comandaassignacio) + rr + ") " + noutilitzades + " order by numerodebobina"
      If lletraseccio = "R" And Not laminadoraprocesinvertit And linkcomanda2 > 0 Then numc = atrim(numcomanda + 2)
        Set rstbobinesdentrada = dbtmpb.OpenRecordset("SELECT laminadores.comanda, bobineslam.* FROM laminadores INNER JOIN bobineslam ON laminadores.Id = bobineslam.controlid WHERE (laminadores.comanda=" + atrim(numc) + ") " + noutilitzades + " order by numerodebobina")
      
      If lletraseccio = "R" And Not laminadoraprocesinvertit Then rr = atrim(numcomanda)
    End If
    Set rstconsulta = dbtemp.OpenRecordset("select * from selecciobobentrada ")
    While Not rstbobinesdentrada.EOF And (seccio <> "" Or seccio2 = "I")
       If cadbl(rstbobinesdentrada!comanda) Then
        rstconsulta.AddNew
        rstconsulta!idpalet = cadbl(rr) 'numcomanda
        rstconsulta!idbobina = rstbobinesdentrada!numerodebobina
        rstconsulta!utilitzada = rstbobinesdentrada!utilitzadaabaixa
        rstconsulta!metres = cadbl(rstbobinesdentrada!metres)
        espesorbobina = cadbl(rstbobinesdentrada!espessor)
        'rstconsulta!tipus = IIf(bobinesdentrada.esrestu(rstbobinesdentrada!idpalet, rstbobinesdentrada!idbobina), "R", "")
        'rstconsulta!tipus = IIf(bobinesdentrada.esparcial(rstbobinesdentrada!idpalet, rstbobinesdentrada!idbobina), "P", rstconsulta!tipus)
         rstconsulta!tipus = seccio
        rstconsulta!taula = nomtaula
        rstconsulta!idb = rstbobinesdentrada!id
        rstconsulta.Update
       End If
        rstbobinesdentrada.MoveNext
        cont = cont + 1
    Wend
fibobinestransformades:
    If laminadoraprocesinvertit And (numcomanda2 - numcomanda) = 2 And Not jahepassatperbobinestransformades And lletraseccio <> "R" Then jahepassatperbobinestransformades = True: seccio = "L": nomtaula = "bobineslam": numc = numcomanda2: GoTo bobinestransformades
    Set rstconsulta = dbtemp.OpenRecordset("select * from selecciobobentrada order by tipus,idpalet,idbobina")
    If InStr(1, funcio, "carregarbobines") > 0 Then Load bobinesdentrada: Exit Sub
    If cont > 0 Then
       If rstconsulta.EOF Then Exit Sub
       rstconsulta.MoveFirst
       bobinesdentrada.multiseleccio = multiselect
       If InStr(1, funcio, "ensenyar") > 0 Then
          bobinesdentrada.Show 1
          'If vnodeixarfercheckaacceptarbobines Then bobinesdentrada.Command1.Enabled = False
          
         ' enviaremailsishaentratmanualment
         
       End If
       If bobinesdentrada.tag = "acceptar" Then
          For i = 1 To bobinesdentrada.reixa.Rows - 1
             If bobinesdentrada.reixa.TextMatrix(i, bobinesdentrada.columnadelcamp("seleccionat")) = "1" Then
               palet = cadbl(bobinesdentrada.reixa.TextMatrix(i, bobinesdentrada.columnadelcamp("idpalet")))
               bobina = cadbl(bobinesdentrada.reixa.TextMatrix(i, bobinesdentrada.columnadelcamp("idbobina")))
               PoB = "B"
               If bobinesdentrada.reixa.TextMatrix(i, bobinesdentrada.columnadelcamp("taula")) = "parcials" Then PoB = "P"
             End If
          Next i
          
       End If
    End If
    
    
    Set rstconsulta = Nothing
    If Not notancarform Then Unload bobinesdentrada
    
    
        Set rsttemp = Nothing
    Set rstb = Nothing
  '  Set dbbaixes = Nothing
    
End Sub
Sub obrestocks2(Optional noobrirbd As Boolean)
 Dim camistocks As String
camistocks = llegir_ini("General", "ruta_stocks", "comandes.ini")
If camistocks = "{[}]" Then escriure_ini "General", "ruta_stocks", rutadelfitxer(cami) + "palets.mdb", "comandes.ini"
camistocks = llegir_ini("General", "ruta_stocks", "comandes.ini")
If Not noobrirbd Then
   Set dbstocks = OpenDatabase(camistocks)
End If
  
End Sub
Function isvisible(vnomform As String) As Boolean
  Dim f
  For Each f In Forms
   If f.Name = vnomform Then
         If f.visible Then isvisible = True
   End If
  Next
End Function
Function isloaded(vnomform As String) As Boolean
  Dim f
  For Each f In Forms
   If f.Name = vnomform Then
         isloaded = True
   End If
  Next
End Function
Function estaobertalabd(db As Database) As Boolean
  Dim vf As String
  On Error GoTo err
  vf = db.Name
  estaobertalabd = True
  Exit Function
err:
  estaobertalabd = False
End Function

Public Function comprovarsireciclarmaterial(vnumc As Double, Optional vtotsencolats As Boolean) As Double
    Dim v1 As Double
    Dim v2 As Double
    Dim v3 As Double
    Dim c1 As Double
    Dim c2 As Double
    Dim vvalormesgran As Byte
    Dim v As String
    Dim vcolorverd As Double
    Dim vcolorblau As Double
    Dim vcolorvermell As Double
    Dim vcolor As Double
    mirarsilacomandaestaacabada vnumc, c1, c2
    vcolorverd = &HFF00&
    vcolorblau = &HF3B378
    vcolorvermell = &HFF&
    v1 = NUMEROCOLORmaterialdelacomanda(atrim(c1))
    v2 = NUMEROCOLORmaterialdelacomanda(atrim(c2))
    vvalormesgran = IIf((IIf(v1 > v2, v1, v2)) > v3, (IIf(v1 > v2, v1, v2)), v3)
    comprovarsireciclarmaterial = IIf(vvalormesgran = 1, vcolorverd, IIf(vvalormesgran = 2, vcolorblau, vcolorvermell))
  End Function
  Public Function NUMEROCOLORmaterialdelacomanda(vnumc As Double) As Double
    Dim rst As Recordset
    NUMEROCOLORmaterialdelacomanda = 0
    Set rst = dbtmpb.OpenRecordset("SELECT materials.colorreciclatge FROM comandes INNER JOIN materials ON comandes.materialex = materials.codi where comanda=" + atrim(cadbl(vnumc)))
    If Not rst.EOF Then NUMEROCOLORmaterialdelacomanda = IIf(atrim(rst!colorreciclatge) = "Verd", 1, IIf(atrim(rst!colorreciclatge) = "Blau", 2, 3))
    Set rst = Nothing
End Function
  
 Public Function comprovarsireciclarmaterial_VELL(vnumc As Double, Optional vtotsencolats As Boolean) As Double
    Dim v1 As String
    Dim v2 As String
    Dim v3 As String
    Dim c1 As Double
    Dim c2 As Double
    Dim v As String
    Dim vcolorverd As Double
    Dim vcolorblau As Double
    Dim vcolorvermell As Double
    Dim vcolor As Double
    mirarsilacomandaestaacabada vnumc, c1, c2
    vcolorverd = &HFF00&
    vcolorblau = &HF3B378
    vcolorvermell = &HFF&
    v1 = materialdelacomanda(atrim(c1)) + "  "
    v2 = materialdelacomanda(atrim(c2)) + "  "
    v = atrim(Mid(v1, 1, InStr(1, v1, " ")))
    If v = "PEAD" Then v = "PEBD" ' en ramon i en miralles han dit que els PEAD son PEBD
    'v1 = v + IIf(InStr(1, v1, "EVOH") > 0, " EVOH", "")
    If InStr(1, v1, "EVOH") > 0 Then
       v1 = "EVOH"
        Else: v1 = v
    End If
    v = atrim(Mid(v2, 1, InStr(1, v2, " ")))
    If v = "PEAD" Then v = "PEBD"  ' en ramon i en miralles han dit que els PEAD son PEBD
    If InStr(1, v2, "EVOH") > 0 Then
        v2 = "EVOH"
         Else: v2 = v
    End If
    'reciclarmaterial1.BackColor = vcolorvermell
    'reciclarmaterial2.BackColor = vcolorvermell
    vcolor = vcolorvermell
    v1 = atrim(v1)
    v2 = atrim(v2)
    If v1 = "PEBD" Then vcolor = vcolorverd
    If v1 = "OPP" Or v1 = "EVOH" Or v1 = "CPP" Then vcolor = vcolorblau
    
    If atrim(v2) <> "" And vcolor <> vcolorvermell Then
      If v2 = "PEBD" And v1 = "PEBD" Then
         vcolor = vcolorverd
          Else
            If v2 = "OPP" Or v2 = "EVOH" Or v2 = "CPP" Then
               vcolor = vcolorblau
                Else: If Not (v2 = "PEBD" And vcolor = vcolorblau) Then vcolor = vcolorvermell
            End If
      End If
    End If
    comprovarsireciclarmaterial_VELL = vcolor
        
       
  End Function


  Public Function materialdelacomanda(vnumc As String) As String
    Dim rst As Recordset
    Set rst = dbtmpb.OpenRecordset("SELECT comandes.comanda, materials.descripcio FROM comandes INNER JOIN materials ON comandes.materialex = materials.codi where comanda=" + atrim(cadbl(vnumc)))
    If Not rst.EOF Then materialdelacomanda = atrim(rst!descripcio)
    Set rst = Nothing
  End Function
Public Sub mirarsilacomandaestaacabada(vnumc As Double, c1 As Double, c2 As Double)
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("select * from laminadorestot where comanda=" + atrim(vnumc))
   If rst.EOF Then c1 = vnumc: GoTo fi
   If Not IsNull(rst!acavada) Then
    If rst!acavada Then
      If rst!comdesb1 = vnumc Then
           c1 = cadbl(rst!comdesb2)
           c2 = cadbl(rst!comdesb1)
            Else
              c1 = cadbl(rst!comdesb1)
              c2 = cadbl(rst!comdesb2)
      End If
     Else: c1 = vnumc: c2 = 0
    End If
   End If
fi:
   If c1 = 0 And c2 = 0 Then c1 = vnumc
   Set rst = Nothing
End Sub
'Sub convertirScanambPaletiBobina(vcodi As String, vpalet As Double, vbob As Double)
'   Dim vcont As Double
'   vcodi = atrim(vcodi)
'   While vcont < Len(vcodi)
'      If Not IsNumeric(Mid(vcodi, vcont + 1, 1)) Then
'        vpalet = cadbl(Mid(vcodi, 1, vcont))
'        If Len(vcodi) >= vcont + 2 Then vbob = cadbl(Mid(vcodi, vcont + 2))
'        GoTo sortir
'      End If
'      vcont = vcont + 1
'   Wend
'sortir:
'End Sub

