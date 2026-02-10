Attribute VB_Name = "funcionscomunes"
Sub enviaremailgeneric(destinatari As String, assumpte As String, cos As String)
   Dim dbenvio As Database
   If atrim(cos) = "" Then Exit Sub
   Set dbenvio = OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
   dbenvio.Execute "insert into envios_mails (data,destinatari,assumpte,cos) values (now,'" + destinatari + "','" + treuresimbols(assumpte) + "','" + treuresimbols(cos) + "')"
   Set dbenvio = Nothing
End Sub
Function treuresimbols(desc As String) As String
   desc = substituir(desc, ":", "_")
   desc = substituir(desc, "'", "´")
   desc = substituir(desc, "|", "_")
   desc = substituir(desc, ";", "_")
   treuresimbols = desc
End Function
Function substituir(cadena As String, buscar As String, canviar As String) As String
   comença = InStr(1, cadena, buscar) - 1
   If comença < 1 Then substituir = cadena: Exit Function
   acaba = comença + Len(buscar) + 1
   cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   substituir = cadena
   'MsgBox linia
End Function

Sub imprimiretiquetabossaclixesdemanantimpresora(numtreball As Double, ordremodificacio As Byte, llistat As CrystalReport, controlarrepeticio As Boolean)
  ' escullirimpresoradetiquetes
  'If llegir_ini("Clixes", "impresoraetiquetesnom", fitxerini) <> "{[}]" Then
    llistat.PrinterName = llegir_ini("Clixes", "impresoraetiquetesnom", fitxerini)
    llistat.PrinterDriver = llegir_ini("Clixes", "impresoraetiquetesdriver", fitxerini)
    llistat.PrinterPort = llegir_ini("Clixes", "impresoraetiquetesport", fitxerini)
    llistat.PrinterSelect
'    llistat.de
    
    escriure_ini "clixes", "impresoraetiquetesdriver", llistat.PrinterDriver, fitxerini
    escriure_ini "clixes", "impresoraetiquetesport", llistat.PrinterPort, fitxerini
    escriure_ini "clixes", "impresoraetiquetesnom", llistat.PrinterName, fitxerini
  ' End If
   imprimiretiquetabossaclixes numtreball, ordremodificacio, llistat, controlarrepeticio
   llistat.Reset
End Sub
Sub escullirimpresoradetiquetes()
   Load escullirimpresora
   escullirimpresora.emplenarllistaimpresores llegir_ini("clixes", "impresoraetiquetesnom", fitxerini)
   escullirimpresora.Show 1
   If escullirimpresora.tag = "acceptar" Then
      escriure_ini "clixes", "impresoraetiquetesdriver", escullirimpresora.driverimpresora, fitxerini
      escriure_ini "clixes", "impresoraetiquetesport", escullirimpresora.portimpresora, fitxerini
      escriure_ini "clixes", "impresoraetiquetesnom", escullirimpresora.nomimpresora, fitxerini
   End If
   Unload escullirimpresora
End Sub

Sub imprimiretiquetabossaclixes(numtreball As Double, ordremodificacio As Byte, llistat As CrystalReport, controlarrepeticio As Boolean)
   Dim rstref As Recordset
   Dim vreferencia As String
   Dim rstc As Recordset
   Dim rstt As Recordset
   Dim vample As String
   Dim refalt As String
   Dim vrefalt As String
   Dim vrefalt2 As String
   Dim cont As Byte
   Set rstref = dbclixes.OpenRecordset("SELECT DISTINCT Clientsvinculats.refclient, clientsvinculats.refclientalternatives, Clixes.id_treball FROM Clixes INNER JOIN Clientsvinculats ON Clixes.id_treball = Clientsvinculats.id_treball WHERE (((Clientsvinculats.refclient)<>'') AND ((Clixes.id_treball)=" + atrim(numtreball) + "));")
   vreferencia = ""
   While Not rstref.EOF
      
      refalt = atrim(rstref!refclient)
filtrar:
      If InStr(1, refalt, "/") = 0 And InStr(1, refalt, "|") = 0 Then
        If InStr(1, vreferencia, atrim(refalt) + " ¦ ") = 0 Then vreferencia = vreferencia + atrim(refalt) + " ¦ "
      End If
      If InStr(1, refalt, "/") > 0 And InStr(1, refalt, "|") = 0 Then refalt = refalt + "|"
      While InStr(1, refalt, "|")
          vrefalt = atrim(Mid(refalt, 1, InStr(1, refalt, "|") - 1))
          'While InStr(1, vrefalt, "/") > 0
            'vrefalt2 = refalt 'Mid(refalt, 1, InStr(1, refalt, "/") - 1)
            'If InStr(1, vreferencia, atrim(vrefalt2) + " ¦") = 0 Then vreferencia = vreferencia + atrim(vrefalt2) + " ¦ "
            'vrefalt = Mid(vrefalt, InStr(1, vrefalt, "/") + 1)
        '  Wend
          If InStr(1, vrefalt, "|") > 0 Then vrefalt = atrim(Mid(vrefalt, 1, InStr(1, vrefalt, "|") - 1))
          If InStr(1, vreferencia, atrim(vrefalt) + " ¦") = 0 Then vreferencia = vreferencia + atrim(vrefalt) + " ¦ "
          refalt = Mid(refalt, InStr(1, refalt, "|") + 1)
          If InStr(1, refalt, "|") = 0 And Len(atrim(refalt)) > 3 Then refalt = refalt + "|"
      Wend
      refalt = atrim(rstref!refclientalternatives)
      cont = cont + 1
      If cont = 1 Then GoTo filtrar
      cont = 0
      rstref.MoveNext
   Wend
   Set rstt = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(numtreball))
   If rstt.EOF Then Exit Sub
   If controlarrepeticio And atrim(rstt!controlimpresioetiquetabossa) = atrim(atrim(atrim(rstt!arxiu) + "  " + vreferencia)) Then Exit Sub
   Set rstc = dbcomandes.OpenRecordset("SELECT comandes.numtreball, comandes.comanda, comandes.*, productes.ruta, clients.nom FROM (comandes INNER JOIN productes ON comandes.producte = productes.codi) INNER JOIN clients ON comandes.client = clients.codi Where (((comandes.numtreball) = " + atrim(numtreball) + ")) ORDER BY comandes.comanda DESC;")
   'If rstc.EOF Then MsgBox "No s'ha trobat cap comanda relacionada.": Exit Sub
   'falta passar els datos al report i fer el report
  'llenço el llistat
   For i = 0 To 50
      llistat.Formulas(i) = ""
   Next i
   generar_codisdebarres numtreball, atrim(rstt!arxiu)
   If MsgBox("Possa paper A4 adhesiu a la impresora i fes acceptar per continuar." + Chr(10) + " O bé Cancelar per no imprimir l'etiqueta.", vbInformation + vbOKCancel + vbDefaultButton2, "Impresió etiqueta de la bossa clixé") = vbCancel Then Exit Sub
   
   llistat.Formulas(0) = "nomclient='" + atrim(rstt!nomclienttemporal) + "'"
   llistat.Formulas(1) = "liniaimarca='" + atrim(rstt!marca) + " - " + atrim(rstt!linia) + "'"
   llistat.Formulas(2) = "ubicacio='" + atrim(rstt!arxiu) + "'"
   llistat.Formulas(3) = "numtreball=" + atrim(numtreball)
   llistat.Formulas(4) = "refclient='" + atrim(vreferencia) + "'"
   llistat.Formulas(5) = "compartitamb='" + mirarsicompartit(numtreball, ordremodificacio) + "'"
   If Not rstc.EOF Then
        If InStr(1, rstc!ruta, "S") > 0 Then
              vample = "AmpleS: " + atrim(cadbl(rstc!amplesol)) + "cm Llarg: " + atrim(cadbl(rstc!longitudsol)) + "cm"
            Else
             If InStr(1, rstc!ruta, "R") > 0 Then
                  vample = "AmpleR: " + atrim(cadbl(rstc!amplereb)) + "cm"
             End If
        End If
   End If
   llistat.Formulas(6) = "refilat='" + atrim(vample) + "'"
   
   vreferencia = atrim(atrim(rstt!arxiu) + "  " + vreferencia)
   llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetabossaclixes.rpt"
   llistat.DiscardSavedData = True
   llistat.CopiesToPrinter = 1
   llistat.DataFiles(0) = rutadelfitxer(cami) + "clixesnous.mdb"
   llistat.Destination = crptToPrinter
   
   
   If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
   llistat.Action = 1
   
   dbclixes.Execute "update  clixes set controlimpresioetiquetabossa=""" + vreferencia + """ where id_treball=" + atrim(numtreball)
   For i = 0 To 50
      llistat.Formulas(i) = ""
   Next i
   wait 2
End Sub

Sub imprimiretiquetabossasoldadores(vnumc As Double, llistat As CrystalReport, controlarrepeticio As Boolean)
   Dim rstref As Recordset
   Dim vreferencia As String
   Dim rstc As Recordset
   Dim rstt As Recordset
   Dim vample As String
   Dim refalt As String
   Dim vrefalt As String
   Dim vrefalt2 As String
   Dim vnumtreball As String
   Dim cont As Byte
   Set rstc = dbcomandes.OpenRecordset("SELECT  comandesmesextres.*, productes.ruta FROM (comandesmesextres INNER JOIN productes ON comandesmesextres.producte = productes.codi) INNER JOIN clients ON comandesmesextres.client = clients.codi Where comanda = " + atrim(vnumc) + " ORDER BY comanda DESC;")
   If rstc.EOF Then Exit Sub
   vnumtreball = atrim(rstc!numerobossasoldadores)

   Set rstref = dbclixes.OpenRecordset("SELECT DISTINCT Clientsvinculats.refclient, clientsvinculats.refclientalternatives, Clixes.id_treball FROM Clixes INNER JOIN Clientsvinculats ON Clixes.id_treball = Clientsvinculats.id_treball WHERE (((Clientsvinculats.refclient)<>'') AND ((Clixes.id_treball)=" + atrim(cadbl(vnumtreball)) + "));")
   vreferencia = ""
   While Not rstref.EOF
      
      refalt = atrim(rstref!refclient)
filtrar:
      If InStr(1, refalt, "/") = 0 And InStr(1, refalt, "|") = 0 Then
        If InStr(1, vreferencia, atrim(refalt) + " ¦ ") = 0 Then vreferencia = vreferencia + atrim(refalt) + " ¦ "
      End If
      If InStr(1, refalt, "/") > 0 And InStr(1, refalt, "|") = 0 Then refalt = refalt + "|"
      While InStr(1, refalt, "|")
          vrefalt = atrim(Mid(refalt, 1, InStr(1, refalt, "|") - 1))
          While InStr(1, vrefalt, "/") > 0
            vrefalt2 = Mid(refalt, 1, InStr(1, refalt, "/") - 1)
            If InStr(1, vreferencia, atrim(vrefalt2) + " ¦") = 0 Then vreferencia = vreferencia + atrim(vrefalt2) + " ¦ "
            vrefalt = Mid(vrefalt, InStr(1, vrefalt, "/") + 1)
          Wend
          If InStr(1, vrefalt, "|") > 0 Then vrefalt = atrim(Mid(vrefalt, 1, InStr(1, vrefalt, "|") - 1))
          If InStr(1, vreferencia, atrim(vrefalt) + " ¦") = 0 Then vreferencia = vreferencia + atrim(vrefalt) + " ¦ "
          refalt = Mid(refalt, InStr(1, refalt, "|") + 1)
          If InStr(1, refalt, "|") = 0 And Len(atrim(refalt)) > 3 Then refalt = refalt + "|"
      Wend
      refalt = atrim(rstref!refclientalternatives)
      cont = cont + 1
      If cont = 1 Then GoTo filtrar
      cont = 0
      rstref.MoveNext
   Wend
   Set rstt = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(cadbl(vnumtreball)))
   
   Set rstc = dbcomandes.OpenRecordset("SELECT comandesmesextres.*, productes.ruta, clients.nom FROM (comandesmesextres INNER JOIN productes ON comandesmesextres.producte = productes.codi) INNER JOIN clients ON comandesmesextres.client = clients.codi Where comanda = " + atrim(vnumc) + " ORDER BY comandes.comanda DESC;")
   If Not rstt.EOF Then
       vnomclient = atrim(rstt!nomclienttemporal)
       vmarcailinia = atrim(rstt!marca) + " - " + atrim(rstt!linia)
       Else
          vnomclient = atrim(rstc!nom)
          vmarcailinia = ""
   End If
   'If rstc.EOF Then MsgBox "No s'ha trobat cap comanda relacionada.": Exit Sub
   'falta passar els datos al report i fer el report
  'llenço el llistat
   For i = 0 To 50
      llistat.Formulas(i) = ""
   Next i
   generar_codisdebarres cadbl(vnumtreball), vnumtreball
   'If MsgBox("Possa paper A4 adhesiu a la impresora i fes acceptar per continuar." + Chr(10) + " O bé Cancelar per no imprimir l'etiqueta.", vbInformation + vbOKCancel + vbDefaultButton2, "Impresió etiqueta de la bossa clixé") = vbCancel Then Exit Sub
   
   llistat.Formulas(0) = "nomclient='" + treure_apostruf(vnomclient) + "'"
   llistat.Formulas(1) = "liniaimarca='" + vmarcailinia + "'"
   llistat.Formulas(2) = "ubicacio=''"
   llistat.Formulas(3) = "numtreball='" + atrim(vnumtreball) + "'"
   llistat.Formulas(4) = "refclient='" + atrim(vreferencia) + "'"
   llistat.Formulas(5) = "compartitamb=''"
   If Not rstc.EOF Then
        If InStr(1, rstc![productes.ruta], "S") > 0 Then
              vample = "AmpleS: " + atrim(cadbl(rstc!amplesol)) + "cm Llarg: " + atrim(cadbl(rstc!longitudsol)) + "cm"
            Else
             If InStr(1, rstc![productes.ruta], "R") > 0 Then
                  vample = "AmpleR: " + atrim(cadbl(rstc!amplereb)) + "cm"
             End If
        End If
   End If
   llistat.Formulas(6) = "refilat='" + atrim(vample) + "'"
   
   'vreferencia = atrim(atrim(rstt!arxiu) + "  " + vreferencia)
   llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetabossasodadores.rpt"
   llistat.DiscardSavedData = True
   llistat.CopiesToPrinter = 1
   llistat.DataFiles(0) = rutadelfitxer(cami) + "clixesnous.mdb"
   llistat.Destination = crptToPrinter
   
   
   If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
   llistat.Action = 1
   
   'dbclixes.Execute "update  clixes set controlimpresioetiquetabossa=""" + vreferencia + """ where id_treball=" + atrim(numtreball)
   For i = 0 To 50
      llistat.Formulas(i) = ""
   Next i
   wait 2
End Sub

Sub generar_codisdebarres(vnumtreball As Double, varxiu As String)
   Dim rst As Recordset
   Dim dbtintestmp As Database
   '  escriure_ini "Tbarcode", "nomfitxer", "c:\temp\prova1.bmp", "generartbarcode.ini"
 '  escriure_ini "Tbarcode", "pixelsample", "1000", "generartbarcode.ini"
 '  escriure_ini "Tbarcode", "pixelsalt", "800", "generartbarcode.ini"
 '  escriure_ini "Tbarcode", "text", "xl-935", "generartbarcode.ini"
 '  escriure_ini "Tbarcode", "printdatatext", "0", "generartbarcode.ini"
 '  escriure_ini "Tbarcode", "tipusbarcode", "62", "generartbarcode.ini"
   '62 es full asci
   '13 as ean 13
   
   '  GENERA EL CODI DE BARRES DEL NUMERO DE TREBALL
   escriure_ini "Tbarcode", "nomfitxer", "c:\temp\~vnumtreball.bmp", "generartbarcode.ini"
   escriure_ini "Tbarcode", "pixelsample", "1000", "generartbarcode.ini"
   escriure_ini "Tbarcode", "pixelsalt", "800", "generartbarcode.ini"
   escriure_ini "Tbarcode", "text", atrim(vnumtreball), "generartbarcode.ini"
   escriure_ini "Tbarcode", "printdatatext", "0", "generartbarcode.ini"
   escriure_ini "Tbarcode", "tipusbarcode", "62", "generartbarcode.ini"
   Shell llegir_ini("General", "rutallistats", "comandes.ini") + "generarimatgedecodidebarres.exe"
   '62 es full asci
   '13 as ean 13
   
   '  GENERA EL CODI DE BARRES DE L'ARXIU
   escriure_ini "Tbarcode", "nomfitxer", "c:\temp\~varxiu.bmp", "generartbarcode.ini"
   escriure_ini "Tbarcode", "pixelsample", "1000", "generartbarcode.ini"
   escriure_ini "Tbarcode", "pixelsalt", "800", "generartbarcode.ini"
   escriure_ini "Tbarcode", "text", atrim(varxiu), "generartbarcode.ini"
   escriure_ini "Tbarcode", "printdatatext", "0", "generartbarcode.ini"
   escriure_ini "Tbarcode", "tipusbarcode", "62", "generartbarcode.ini"
   Shell llegir_ini("General", "rutallistats", "comandes.ini") + "generarimatgedecodidebarres.exe"
   '62 es full asci
   '13 as ean 13
   Set dbtintestmp = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
   Set rst = dbtintestmp.OpenRecordset("select * from valorsgenerals")
   If rst.EOF Then
        rst.AddNew
      Else: rst.Edit
   End If
   r = copiafoto("c:\temp\~vnumtreball.bmp", rst!imatge)
   r = copiafoto("c:\temp\~varxiu.bmp", rst!imatge2)
   'rst!codiop = cadbl(cpersona.tag)
   rst.Update
   Set rst = Nothing
   Set dbtintestmp = Nothing
End Sub
Function copiafoto(foto As String, fldTO As Field)

'This function takes the source field image and copies it
'into the destination field.
'The function first saves the image in the source field to a
'temp file on disc. Then reads this temp file into
'the destination field.
'The temp file is then deleted
'On Error Resume Next

Dim iFieldSize  As Long
Dim varChunk    As Variant
Dim baData()    As Byte
Dim iOffset     As Long
Dim sFName      As String
Dim iFileNum    As Long
Dim cnt         As Long
Dim z()         As Byte

Const CONCHUNKSIZE As Long = 16384

Dim iChunks As Long
Dim iFragmentSize As Long
    
    'Get a unique random filename
    If Not existeix(foto) Then Exit Function
    sFName = foto
    
    Open sFName For Binary Access Read As #1
    ReDim z(FileLen(sFName))
    Get #1, , z()
     fldTO.AppendChunk z
    Close #1
    
    'Delete the file
    'Kill (sFName)
    
End Function


Function mirarsicompartit(numc As Double, vordremodificacio As Byte) As String
  Dim rst As Recordset
  mirarsicompartit = ""
  'Set rst = dbclixes.OpenRecordset("select * from tintes where tinterlinkambid_treball,id_treball=" + atrim(numc))
  'he fet aquest canvi HE TRET ORDREMODIFICACIO DELA SELECCIO
    ' "select distinct id_treball from tintes where ordremodificacio=" + atrim(vordremodificacio) + " and id_tinter in(select tinterlinkambid_treball from tintes where tinterlinkambid_treball and id_treball=" + atrim(numc) + ")"
    
'selecciono tots els treballs que aquest treball va a estirar clixe (ancora groga d'AQUEST)
  Set rst = dbclixes.OpenRecordset("select distinct id_treball from tintes where  id_tinter in (select tinterlinkambid_treball from tintes where tinterlinkambid_treball and id_treball=" + atrim(numc) + " and ordremodificacio=" + atrim(vordremodificacio) + ")")
  While Not rst.EOF
      mirarsicompartit = mirarsicompartit + atrim(rst!id_treball) + " | "
      rst.MoveNext
  Wend
 GoTo fi
  'salto el seguent pas perque diuen que no volen veure els altres que estiren
  
'selecciono tots els treballs que estiren clixe d'aquest  (ancora groga DELS ALTRES)
  Set rst = dbclixes.OpenRecordset("select distinct id_treball from tintes where  tinterlinkambid_treball in (select id_tinter from tintes where comparteix and (ordremodificacio=" + atrim(vordremodificacio) + " and id_treball=" + atrim(numc) + "))")
 ' Clipboard.Clear
 ' Clipboard.SetText "select distinct id_treball from tintes where  tinterlinkambid_treball in (select id_tinter from tintes where comparteix and (ordremodificacio=" + atrim(vordremodificacio) + " and id_treball=" + atrim(numc) + "))"
  While Not rst.EOF
      mirarsicompartit = mirarsicompartit + atrim(rst!id_treball) + " | "
      rst.MoveNext
  Wend
fi:
  If mirarsicompartit <> "" Then mirarsicompartit = "Compartit amb: " + mirarsicompartit
End Function

Sub posardiferenciesacomandadeltreball(numc As Double)
   Dim i As Byte
   Dim cilindre As Double
   Dim desarroll As Double
   Dim continu As String
   Dim rstc As Recordset
   Dim rstclixe As Recordset
   Dim rstmodificacio As Recordset
   Dim rsttintes As Recordset
   Dim rsttintesllaunes As Recordset
   Dim rstlink As Recordset
   Dim treball As Integer
   Dim modificacio As Integer
   Dim arxiumontadora As String
    Dim vtinters As Byte
   Dim vnomtinta As String
   
   Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
   If rstc.EOF Then GoTo noinfo
   treball = cadbl(rstc!numtreball): modificacio = cadbl(rstc!numordremodificacio)
   If modificacio = 0 Then modificacio = 1
   Set rstclixe = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(cadbl(rstc!numtreball)))
   If rstclixe.EOF Then GoTo noinfo
   Set rstmodificacio = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordre=" + atrim(cadbl(modificacio)))
   If rstmodificacio.EOF Then GoTo noinfo
   rstc.Edit
   Set rsttintes = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordremodificacio=" + atrim(cadbl(modificacio)) + " order by ordretinter")
   Set rsttintesllaunes = dbclixes.OpenRecordset("select codi,descripcio from tintes_llaunes", , dbonly)
   continu = ""
   If rsttintes.EOF Then GoTo noinfo
      'comprovo tintes
      For i = 1 To 8
        rsttintes.FindFirst "ordretinter = " + atrim(i)
        If Not rsttintes.NoMatch Then
          If atrim(rsttintes!color) <> "" Or cadbl(rsttintes!tinterlinkambid_treball) > 0 Then vtinters = vtinters + 1
          Set rstlink = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(IIf(rsttintes!tinterlinkambid_treball > 0, rsttintes!tinterlinkambid_treball, rsttintes!id_tinter)))
          rsttintesllaunes.FindFirst "codi='" + atrim(rstlink!coditinta) + "'"
          If rsttintesllaunes.NoMatch And atrim(rsttintes!color) = "" Then posartinterdecomandaazero rstc, i: GoTo proxima
          vnomtinta = ""
          If atrim(rsttintes!color) <> "" And cadbl(rsttintes!tinterlinkambid_treball) = 0 Then
              vnomtinta = atrim(rsttintes!color)
                Else: vnomtinta = atrim(rsttintesllaunes!descripcio)
          End If
          rstc.Fields("tinta" + atrim(i) + "a") = vnomtinta
          rstc.Fields("lin" + atrim(i)) = atrim(rstlink!anilox)
        '  rstc.Fields("tinta" + atrim(i) + "b") = atrim(rstlink!observacions)
          rstc.Fields("detalltinter" + atrim(i)) = atrim(rstlink!detalltinter)
          If continu <> "S" And atrim(rstlink!color) <> "" Then
            rstc!continu = IIf(Not rstlink!continuu, "N", "S")
            rstc!cilindres = cadbl(rstlink!cilindre)
            If cadbl(rstlink!desarroll) > cadbl(rstc!dessarroll) Then rstc!dessarroll = cadbl(rstlink!desarroll)
            continu = rstc!continu
              Else: If atrim(rstlink!color) <> "" Then rstc!dessarroll = 0
          End If
            Else: posartinterdecomandaazero rstc, i
        End If
proxima:
      Next i
      copiarobservacionstreballacomanda treball, modificacio, numc
      Set rstlink = Nothing
     If larutahiha(rstc!producte, "R") And atrim(rstc!amplereb) <> atrim(rstmodificacio!amplelamina) Then
     rstc!amplereb = rstmodificacio!amplelamina
   End If
   
   arxiumontadora = buscararxiumontadora(cadbl(rstmodificacio!id_treball), cadbl(rstmodificacio!ordre), cadbl(rstc!client), cadbl(rstc!direnvio))
   rstc!arxiumontadora = arxiumontadora
   If vtinters <> cadbl(rstc!numerotintes) Then rstc!numerotintes = vtinters
   rstc!arxiu = rstclixe!arxiu
   rstc!formaimp = atrim(rstmodificacio!formaimpresio)
   rstc!gruixpol = cadbl(rstmodificacio!gruixpolimer)
   'rstc!codibarras = IIf(Len(rstclixe!codidebarres) > 15, Mid(atrim(rstclixe!codidebarres), Len(atrim(rstclixe!codidebarres)) - 14), atrim(rstclixe!codidebarres))
   If Len(atrim(rstclixe!codidebarres)) > 15 Then
           rstc!codibarras = Mid(atrim(rstclixe!codidebarres), Len(atrim(rstclixe!codidebarres)) - 14)
            Else: rstc!codibarras = atrim(rstclixe!codidebarres)
   End If
   rstc!cmaquina = IIf(atrim(rstclixe!redcilindrefw) <> "", atrim(rstclixe!reduccioxmetre), atrim(rstc!cmaquina))
 '  rstc!amplereb = cadbl(rstmodificacio!amplelamina)
   
   rstc.Update
   Set rsttintesllaunes = Nothing
   Set rstc = Nothing
   Set rsttintes = Nothing
   Set rstmodificacio = Nothing
   Exit Sub
noinfo:

End Sub
Sub posartinterdecomandaazero(rstc As Recordset, i As Byte)
   rstc.Fields("tinta" + atrim(i) + "a") = ""
   rstc.Fields("lin" + atrim(i)) = 0
   rstc.Fields("detalltinter" + atrim(i)) = ""
End Sub

Sub imprimirdiferenciescomandaitreball(numc As Double)
 ' Dim rst As Recordset
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
 
  Set rst = dbclixes.OpenRecordset("select * from diferenciescomandaitreball where comanda=" + atrim(numc))
  If rst.EOF Then Exit Sub
 ' llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "incidenciescomandaitreball.rpt"
 ' If Not existeix("c:\ordprog.ini") Then llistat.Destination = crptToPrinter
 ' llistat.DataFiles(0) = rutadelfitxer(cami) + "clixesnous.mdb"
 ' llistat.SelectionFormula = "{diferenciescomandaitreball.comanda}=" + atrim(numc)
 ' llistat.Action = 1
  
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "incidenciescomandaitreball.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "clixesnous.mdb"
  oreport.RecordSelectionFormula = "{diferenciescomandaitreball.comanda}=" + atrim(numc)
  
  oreport.DiscardSavedData
  'If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
   ' Else
   '   oreport.PrintOut False, 1
 ' End If
  
End Sub

Function mirardiferenciescomandaitreball(numc As Double) As Boolean
   Dim i As Byte
   Dim cilindre As Double
   Dim desarroll As Double
   Dim continu As String
   Dim rstc As Recordset
   Dim rstclixe As Recordset
   Dim rstmodificacio As Recordset
   Dim rsttintes As Recordset
   Dim rsttintesllaunes As Recordset
   Dim rstlink As Recordset
   Dim treball As Integer
   Dim arxiumontadora As String
   Dim modificacio As Integer
   Dim vtinters As Byte
   Dim vnomtinta As String
   
   dbclixes.Execute "delete * from diferenciescomandaitreball where comanda=" + atrim(numc)
   Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
   If rstc.EOF Then GoTo noinfo
   treball = cadbl(rstc!numtreball): modificacio = cadbl(rstc!numordremodificacio)
   If modificacio = 0 Then modificacio = 1
   Set rstclixe = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(cadbl(rstc!numtreball)))
   If rstclixe.EOF Then GoTo noinfo
   Set rstmodificacio = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordre=" + atrim(cadbl(modificacio)))
   If rstmodificacio.EOF Then GoTo noinfo
   continu = ""
   Set rsttintes = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordremodificacio=" + atrim(cadbl(modificacio)) + " order by ordretinter")
   Set rsttintesllaunes = dbclixes.OpenRecordset("select codi,descripcio from tintes_llaunes", , dbonly)
   If rsttintes.EOF Then GoTo noinfo
      'comprovo tintes
      
      For i = 1 To 8
        rsttintes.FindFirst "ordretinter = " + atrim(i)
        If rsttintes.NoMatch Then GoTo proxima
        Set rstlink = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(IIf(rsttintes!tinterlinkambid_treball > 0, rsttintes!tinterlinkambid_treball, rsttintes!id_tinter)))
        If rstlink.EOF Then
            If atrim(rstc.Fields("tinta" + atrim(i) + "a")) <> "" Then
                posardiferencia "Tinter Nº " + atrim(i), atrim(rstc.Fields("tinta" + atrim(i) + "a")), "<Sense Tinta>", treball, modificacio, numc
            End If
            GoTo proxima
        End If
        If atrim(rsttintes!color) <> "" Or cadbl(rsttintes!tinterlinkambid_treball) > 0 Then vtinters = vtinters + 1
        rsttintesllaunes.FindFirst "codi='" + atrim(rstlink!coditinta) + "'"
        If rsttintesllaunes.NoMatch And atrim(rsttintes!color) = "" Then
            If atrim(rstc.Fields("tinta" + atrim(i) + "a")) <> "" Then
              posardiferencia "Tinter Nº " + atrim(i), atrim(rstc.Fields("tinta" + atrim(i) + "a")), "<Sense Tinta>", treball, modificacio, numc
            End If
            GoTo proxima
        End If
        vnomtinta = IIf(Not rsttintesllaunes.NoMatch, atrim(rsttintesllaunes!descripcio), atrim(rsttintes!color))
        If atrim(rstc.Fields("tinta" + atrim(i) + "a")) <> vnomtinta Then posardiferencia "Tinter Nº " + atrim(i), atrim(rstc.Fields("tinta" + atrim(i) + "a")), vnomtinta, treball, modificacio, numc
        If atrim(rstc.Fields("detalltinter" + atrim(i))) <> atrim(rstlink!detalltinter) Then posardiferencia "Detalltinter Nº " + atrim(i), atrim(rstc.Fields("detalltinter" + atrim(i))), atrim(rstlink!detalltinter), treball, modificacio, numc
        If cadbl(atrim(rstc.Fields("lin" + atrim(i)))) <> cadbl(atrim(rstlink!anilox)) Then posardiferencia "Anilox Nº " + atrim(i), atrim(rstc.Fields("lin" + atrim(i))), atrim(rstlink!anilox), treball, modificacio, numc
          'If atrim(rstc.Fields("tinta" + atrim(i) + "b")) <> atrim(rstlink!observacions) Then posardiferencia "Observacions Nº " + atrim(i), atrim(rstc.Fields("tinta" + atrim(i) + "b")), atrim(rstlink!observacions), treball, modificacio, numc
          
          
          'agafo els primers valors
         ' If continu = "" Then continu = IIf(Not rstlink!continuu, "N", "S")
         ' If cilindre = 0 Then cilindre = cadbl(rstlink!cilindre)
         ' If desarroll = 0 Then desarroll = cadbl(rstlink!desarroll)
         If continu <> "S" And atrim(rstlink!color) <> "" Then
            continu = IIf(Not rstlink!continuu, "N", "S")
            cilindre = cadbl(rstlink!cilindre)
            If cadbl(rstlink!desarroll) > cadbl(desarroll) Then desarroll = cadbl(rstlink!desarroll)
            
            '  Else: If atrim(rstlink!color) <> "" Then desarroll = 0 'HEM TRET AQUESTA OPCIO AMB LA EVA PERQUÈ NO SABEM PERQUÈ LA VAM POSAR
         End If
proxima:
      Next i
      If diferenciesobservacionstreballicomanda(treball, modificacio, numc) Then posardiferencia "Observacions diferents", "Diferencies", "Diferencies", treball, modificacio, numc
      Set rstlink = Nothing
      If vtinters <> cadbl(rstc!numerotintes) Then posardiferencia "NºTinters", atrim(rstc!numerotintes), atrim(vtinters), treball, modificacio, numc
      If IIf(atrim(rstc!continu) = "", "N", atrim(rstc!continu)) <> continu Then posardiferencia "Continuu", atrim(rstc!continu), continu, treball, modificacio, numc
      If cadbl(rstc!cilindres) <> cilindre Then posardiferencia "Cilindre", cadbl(rstc!cilindres), atrim(cilindre), treball, modificacio, numc
      If cadbl(rstc!dessarroll) <> desarroll Then posardiferencia "Desarroll", cadbl(rstc!dessarroll), atrim(desarroll), treball, modificacio, numc
   If larutahiha(rstc!producte, "R") And atrim(rstc!amplereb) <> atrim(rstmodificacio!amplelamina) Then
     posardiferencia "Ample lamina", atrim(rstc!amplereb), atrim(rstmodificacio!amplelamina), treball, modificacio, numc
   End If
   
   If atrim(rstc!formaimp) <> atrim(rstmodificacio!formaimpresio) Then posardiferencia "Forma Impresio", atrim(rstc!formaimp), atrim(rstmodificacio!formaimpresio), treball, modificacio, numc
   If cadbl(rstc!gruixpol) <> cadbl(rstmodificacio!gruixpolimer) Then posardiferencia "Gruix Pol.", atrim(rstc!gruixpol), atrim(rstmodificacio!gruixpolimer), treball, modificacio, numc
   If atrim(rstc!codibarras) <> atrim(rstclixe!codidebarres) Then
     If Len(rstclixe!codidebarres) > 15 Then
       If Mid(rstclixe!codidebarres, Len(rstclixe!codidebarres) - 14) <> atrim(rstc!codibarras) Then
           posardiferencia "Codi de Barres", atrim(rstc!codibarras), atrim(rstclixe!codidebarres), treball, modificacio, numc
       End If
          Else
               posardiferencia "Codi de Barres", atrim(rstc!codibarras), atrim(rstclixe!codidebarres), treball, modificacio, numc
     End If
   End If
   If atrim(rstc!cmaquina) <> atrim(rstclixe!reduccioxmetre) And atrim(rstclixe!redcilindrefw) <> "" Then posardiferencia "Reduccio per metre", atrim(rstc!cmaquina), atrim(rstclixe!reduccioxmetre), treball, modificacio, numc
   
   If atrim(rstc!arxiu) <> atrim(rstclixe!arxiu) Then posardiferencia "Arxiu Clixe", atrim(rstc!arxiu), atrim(rstclixe!arxiu), treball, modificacio, numc
   arxiumontadora = buscararxiumontadora(cadbl(rstmodificacio!id_treball), cadbl(rstmodificacio!ordre), cadbl(rstc!client), cadbl(rstc!direnvio))
   If atrim(rstc!arxiumontadora) <> arxiumontadora Then posardiferencia "Arxiu Muntadora", atrim(rstc!arxiumontadora), atrim(arxiumontadora), treball, modificacio, numc
noinfo:
   Set rstc = dbclixes.OpenRecordset("select * from diferenciescomandaitreball where comanda=" + atrim(numc))
   If rstc.EOF Then
       mirardiferenciescomandaitreball = False
        Else: mirardiferenciescomandaitreball = True
   End If

End Function
Function larutahiha(producte As String, seccio As String) As Boolean
   Dim rstp As Recordset
   Set rstp = dbcomandes.OpenRecordset("select ruta from productes where codi='" + atrim(producte) + "'")
   If rstp.EOF Then Exit Function
   If InStr(1, rstp!ruta, seccio) > 0 Then
       larutahiha = True
      Else: larutahiha = False
   End If
End Function
Sub posardiferencia(camp As String, valorcomanda As String, valortreball As String, treball As Integer, modificacio As Integer, numc As Double)
   Dim valors As String
   valors = atrim(treball) + "," + atrim(modificacio) + "," + atrim(numc) + ",'" + treure_apostruf(camp) + "','" + treure_apostruf(valorcomanda) + "','" + treure_apostruf(valortreball) + "'"
   dbclixes.Execute "insert into diferenciescomandaitreball (id_treball,ordremodificacio,comanda,camp,valorcomanda,valortreball) values (" + valors + ")"
End Sub

Function diferenciesobservacionstreballicomanda(idtreball As Integer, ordremodificacio As Integer, numc As Double) As Boolean
   Dim vc1 As String
   Dim vc2 As String
   Dim vt1 As String
   Dim vt2 As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from tintes_observacions where id_treball=" + atrim(idtreball) + " and ordre=" + atrim(ordremodificacio) + " order by id")
   If Not rst.EOF Then vt1 = atrim(rst!observacio): rst.MoveNext
   If Not rst.EOF Then vt2 = atrim(rst!observacio): rst.MoveNext
   
   Set rst = dbcomandes.OpenRecordset("select * from comandes_observacionstintes where comanda=" + atrim(numc) + " order by id")
   If Not rst.EOF Then vc1 = atrim(rst!observacio): rst.MoveNext
   If Not rst.EOF Then vc2 = atrim(rst!observacio): rst.MoveNext
   
   If treure_apostruf(vt1) <> vc1 Then diferenciesobservacionstreballicomanda = True
   If treure_apostruf(vt2) <> vc2 Then diferenciesobservacionstreballicomanda = True
   
   Set rst = Nothing
End Function

Function buscararxiumontadora(treball As Double, modificacio As Double, client As Double, direnvio As Double) As String
  Dim rst As Recordset
  Set rst = dbclixes.OpenRecordset("select * from clientsvinculats where id_treball=" + atrim(treball) + " and ordremodificacio=" + atrim(modificacio) + " and codiclient=" + atrim(client) + " and direnvio=" + atrim(direnvio))
  If Not rst.EOF Then buscararxiumontadora = atrim(rst!codimuntadora)
  If rst.EOF Then
        Set rst = dbclixes.OpenRecordset("select * from clientsvinculats where codimuntadora<>'' and id_treball=" + atrim(treball))
        If Not rst.EOF Then buscararxiumontadora = atrim(rst!codimuntadora)
  End If
  Set rst = Nothing
End Function


Function copiarobservacionstreballacomanda(idtreball As Integer, ordremodificacio As Integer, numc As Double) As Boolean
   Dim vc1 As String
   Dim vc2 As String
   Dim vt1 As String
   Dim vt2 As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from tintes_observacions where id_treball=" + atrim(idtreball) + " and ordre=" + atrim(ordremodificacio) + " order by id")
   If Not rst.EOF Then vt1 = atrim(rst!observacio): rst.MoveNext
   If Not rst.EOF Then vt2 = atrim(rst!observacio): rst.MoveNext
   dbcomandes.Execute "delete * from comandes_observacionstintes where comanda=" + atrim(numc)
   If vt1 <> "" Then
    dbcomandes.Execute "insert into comandes_observacionstintes (comanda,observacio) values (" + atrim(numc) + ",'" + treure_apostruf(vt1) + "')"
    If vt2 <> "" Then dbcomandes.Execute "insert into comandes_observacionstintes (comanda,observacio) values (" + atrim(numc) + ",'" + treure_apostruf(vt2) + "')"
   End If
   Set rst = Nothing
End Function

Public Function nomordinador() As String
   nomordinador = Environ("computername")
End Function
Function enviaremail(sSendTo As String, sSubject As String, sText As String, Optional adjunt As String, Optional vidavis As Long, Optional adjunt2 As String, Optional adjunt3 As String) As Boolean
  Dim usuarim As String
  Dim contrasenyam As String
  Dim destinatari As String
  Dim vnomcarpeta As String
  Dim vadjunt As String
  Dim vadjunt2 As String
  Dim vadjunt3 As String
  Dim vv As String
  vv = llegir_ini("destinataris", sSendTo, "enviarservidor.ini")
  If vv = "{[}]" Then
    sSendTo = sSendTo
     Else: sSendTo = vv
  End If
  vadjunt = adjunt
  vadjunt2 = adjunt2
  vadjunt3 = adjunt3
  vnomcarpeta = "\\serverprodu\Dades\progcomandes\dades\spoolerenviament\" + nomordinador + "_" + Format(Now, "yymmdd_hhnnss")
'  usuarim = llegir_ini("dadesservidor", "usrsmtp", "enviarservidor.ini")
'  contrasenyam = llegir_ini("dadesservidor", "passsmtp", "enviarservidor.ini")
  If usuarim = "{[}]" Or contrasenyam = "{[}]" Then
      escriure_ini "dadesservidor", "usrsmtp", " ", "enviarservidor.ini"
      escriure_ini "dadesservidor", "passsmtp", " ", "enviarservidor.ini"
      MsgBox "L'usuari o la contrasenya no estan entrades", vbCritical, "Error": Exit Function
  End If
  If Not existeix(vnomcarpeta) Then MkDir vnomcarpeta
  escriure_ini "Capcalera", "apuntperenviar", "No", vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "data", Now, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "nomordinador", nomordinador, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "usuari", usuarim, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "contrasenya", contrasenyam, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "destinatari", sSendTo, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "remitent", "incidencies@inplacsa.com", vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "assumpte", treure_apostruf(sSubject), vnomcarpeta + "\dadesmail.txt"
  If existeix(vadjunt) Then
     Copiar_Fitxer vadjunt, vnomcarpeta
     vadjunt = substituirtot(vadjunt, rutadelfitxer(vadjunt), vnomcarpeta + "\")
     escriure_ini "Capcalera", "adjunt", vadjunt, vnomcarpeta + "\dadesmail.txt"
  End If
  If existeix(vadjunt2) Then
     Copiar_Fitxer vadjunt2, vnomcarpeta
     vadjunt2 = substituirtot(vadjunt2, rutadelfitxer(vadjunt2), vnomcarpeta + "\")
     escriure_ini "Capcalera", "adjunt2", vadjunt2, vnomcarpeta + "\dadesmail.txt"
  End If
  If existeix(vadjunt3) Then
     Copiar_Fitxer vadjunt3, vnomcarpeta
     vadjunt3 = substituirtot(vadjunt3, rutadelfitxer(vadjunt3), vnomcarpeta + "\")
     escriure_ini "Capcalera", "adjunt3", vadjunt3, vnomcarpeta + "\dadesmail.txt"
  End If
  
  If LCase(sText) <> "c:\temp\cosmissatge.txt" Then
        Open "c:\temp\cosmissatge.txt" For Output As #2
        Print #2, sText
        passarliniesdavisosalfitxertxt vidavis
        Close #2
   End If
   Copiar_Fitxer "c:\temp\cosmissatge.txt", vnomcarpeta
   If existeix("c:\temp\cosmissatge.txt") Then Kill "c:\temp\cosmissatge.txt"
   escriure_ini "Capcalera", "apuntperenviar", "Si", vnomcarpeta + "\dadesmail.txt"
   wait 1
   
End Function
Sub passarliniesdavisosalfitxertxt(vidavis As Long)
    Dim rst As Recordset
    Dim v As String
    If vidavis = 0 Then Exit Sub
    Set rst = db.OpenRecordset("select * from envios_mails_linies where id_envio=" + atrim(vidavis))
    If Not rst.EOF Then
       Print #2, ""
       Print #2, ""
    End If
    While Not rst.EOF
      v = atrim(rst!descripcio)
      If Len(v) > 0 Then
        'If InStr(1, v, Chr(10)) = 0 Then v = v + Chr(10)
        Print #2, v
      End If
      rst.MoveNext
    Wend
    Set rst = Nothing
End Sub

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

