Attribute VB_Name = "FuncionsGlobals"
'Global arguments As Variant

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess _
    As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Function proximdianatural() As Date
   proximdianatural = DateAdd("d", 1, Now)
   While Format(proximdianatural, "w", vbMonday) > 5
      proximdianatural = DateAdd("d", 1, proximdianatural)
   Wend
End Function
Sub actualitzar_bobinesent(numc As Double, Optional vruta As String)
  Dim ultimaseccio As String
  Dim seccions As Variant
  Dim seccionsbob As Variant
  Dim ordre As String
  Dim nomtaula As String
  Dim nomsubtaula As String
  Dim idscontrol As String
  Dim rstbob As Recordset
  Dim rstc As Recordset
  
  '''''    !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!11
  '''''
  '''''  ATENCIÓ QUALSEVOL CANVI EN AQUESTA FUNCIO TAMBÉ S'HA DE FER AL MODUL DE VENTES
  '''''
  '''''    !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!11
  If vruta <> "" Then ruta = vruta
  ordre = "EILRS"
  seccions = Array("extrussores", "impressores", "laminadores", "rebobinadores", "soldadores")
  seccionsbob = Array("Bobinesext", "bobinesimp", "bobineslam", "bobinesreb", "bobinessol")
  If ruta <> "" Then ultimaseccio = Mid(ruta, Len(ruta), 1)
  nomtaula = seccions(InStr(1, ordre, ultimaseccio) - 1)
  nomsubtaula = seccionsbob(InStr(1, ordre, ultimaseccio) - 1)
  nomordre = "numerodebobina"
  If nomsubtaula = "bobinessol" Then nomordre = "numerodesac"
  'busco les dades de la comanda per poder calcular el pes de soladores
  Set rstc = dbtmp.OpenRecordset("SELECT comandes.*, comandes_extres.codicomptable,comandes_extres.solpesgrmcm2 FROM comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda where comandes.comanda=" + atrim(numc))
  'Miro tots els registres de la taula principal per fer la busqueda a les bobines(subtaula)
  Set rsttmp = dbtmpb.OpenRecordset("select * from " + nomtaula + " where comanda=" + atrim(cadbl(numc)))
  While Not rsttmp.EOF
     If idscontrol <> "" Then
         idscontrol = idscontrol + " or controlid=" + atrim(cadbl(rsttmp!ID))
        Else: idscontrol = " controlid=" + atrim(cadbl(rsttmp!ID))
     End If
     rsttmp.MoveNext
  Wend
  
  'Faig la busqueda de la subtaula i les entro a bobinesent
  If idscontrol <> "" Then
   Set rsttmp = dbtmpb.OpenRecordset("select * from " + nomsubtaula + " where " + idscontrol + " order by " + nomordre + " ASC")
   While Not rsttmp.EOF
     'If rsttmp.RecordCount = "F" Then
      On Error Resume Next
      numbob = cadbl(rsttmp!numerodebobina)
      numbob = cadbl(rsttmp!numerodesac)
      On Error GoTo 0
      'comprovo que si la bobina ja està afegida a les bobines d'entrega i si no la afegeixo o la modifico
      Set rstbob = dbtmpb.OpenRecordset("select *  from bobinesent where comanda=" + atrim(cadbl(numc)) + " and numbob=" + atrim(numbob))
      If rstbob.EOF Then
          'afegeixo la bobina
         dbtmp.Execute "update comandes set seccioactual='P' where comanda=" + atrim(numc)
         rstbob.AddNew
         rstbob!comanda = atrim(cadbl(numc))
         rstbob!controlid = rsttmp!controlid
         
         rstbob!seccio = ultimaseccio
        
         If ultimaseccio = "S" Then
              rstbob!metresisacs = cadbl(rsttmp!unitatsxsac)
              rstbob!kilosiunitats = Redondejar(calcularpesxrpeça(rstc) * cadbl(rsttmp!unitatsxsac), 2)
              rstbob!numbob = rsttmp!numerodesac
              rstbob!numpalet = cadbl(rsttmp!palet)
              If rstbob!numpalet = 0 Then rstbob!numpalet = 1
           Else
              rstbob!numbob = rsttmp!numerodebobina
              rstbob!metresisacs = rsttmp!metres
              On Error Resume Next
              rstbob!kilosnets = Redondejar(rsttmp!pesnet, 2)
              rstbob!numpalet = rsttmp!palet
              If rstbob!numpalet = 0 Then rstbob!numpalet = 1
              rstbob!kilosiunitats = Redondejar(rsttmp!kilos, 2)
              On Error GoTo 0
         End If
         rstbob.Update
          Else
            'modifico la bobina si correspont fer-ho
            If ultimaseccio = "S" Then
              If rstbob!metresisacs <> rsttmp!unitatsxsac Or cadbl(rstbob!numpalet) <> cadbl(rsttmp!palet) Then
                 rstbob.Edit
                 rstbob!metresisacs = rsttmp!unitatsxsac
                 rstbob!kilosiunitats = Redondejar(calcularpesxrpeça(rstc) * cadbl(rsttmp!unitatsxsac), 2)
                 rstbob!numpalet = cadbl(rsttmp!palet)
                 rstbob.Update
              End If
           Else
              If rstbob!metresisacs <> rsttmp!metres Then rstbob.Edit: rstbob!metresisacs = rsttmp!metres: rstbob.Update
              On Error Resume Next
              vpesnet = cadbl(rsttmp!pesnet)
              If cadbl(rstbob!kilosiunitats) <> cadbl(rsttmp!kilos) Or cadbl(rstbob!kilosnets) <> cadbl(vpesnet) Then
                 rstbob.Edit
                 rstbob!kilosiunitats = Redondejar(cadbl(rsttmp!kilos), 2)
                 rstbob!kilosnets = Redondejar(cadbl(rsttmp!pesnet), 2)
                 rstbob.Update
              End If
              If cadbl(rstbob!numpalet) <> rsttmp!palet Then rstbob.Edit: rstbob!numpalet = rsttmp!palet: rstbob.Update
              On Error GoTo 0
         End If
      End If
    ' End If
      rsttmp.MoveNext
   Wend
   dbtmpb.Execute ("delete *  from bobinesent where comanda=" + atrim(cadbl(numc)) + " and numbob>" + atrim(cadbl(numbob)))
   
   Set rstbob = dbtmpb.OpenRecordset("select numbob from bobinesent where comanda=" + atrim(cadbl(numc)) + " order by numbob ASC")
   c = IIf(ultimaseccio = "S", "numerodesac", "numerodebobina")

   While Not rstbob.EOF
        Set rsttmp = dbtmpb.OpenRecordset("select " + c + " from " + nomsubtaula + " where (" + idscontrol + ")and " + c + "=" + atrim(cadbl(rstbob!numbob)) + " order by " + nomordre + " ASC")
        'Set rsttmp = dbtmpb.OpenRecordset("select " + c + " from " + nomtaula + " where comanda=" + atrim(cadbl(numc)) + " and numbob=" + atrim(cadbl(rstbob.Fields![c])))
        'MsgBox Trim(rsttmp!numerodebobina) + "     " + Trim(rstbob!numbob)
        If rsttmp.EOF Then
           dbtmpb.Execute ("delete *  from bobinesent where comanda=" + atrim(cadbl(numc)) + " and numbob=" + atrim(cadbl(rstbob!numbob)))
        End If
        rstbob.MoveNext
   Wend
  End If
  Set rstc = Nothing
  Set rstbob = Nothing
End Sub
Function calcularpesxrpeça(rst As Recordset) As Double
    Dim pesgrmcm2 As Double
    If cadbl(rst!cantitatsol) = 0 Then Exit Function
    pesgrmcm2 = cadbl(rst!solpesgrmcm2)
    calcularpesxrpeça = pesgrmcm2 * (cadbl(rst!amplesol) * (cadbl(rst!longitudsol) + (cadbl(rst!solapasol) / 2)))
    calcularpesxrpeça = calcularpesxrpeça * IIf(rst!migelaboratsol = "L", 1, 2)
    'calcularpesxrpeça = cadbl(rst!cantitatsol) * calcularpesxrpeça
End Function


Sub comprovarsitepreuassignatosinoenviarunmail(numc As Double)
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("SELECT comandes.pvp, comandes.numpressupost,comandes.proximaseccio,comandes.producte,clients.grupdeclient FROM comandes LEFT JOIN clients ON comandes.client = clients.codi Where comanda = " + atrim(numc))
  If rst.EOF Then Exit Sub
  If InStr(1, atrim(rst!producte), "PC") > 0 Then GoTo fi
  If atrim(rst!proximaseccio) = "T" Then GoTo fi
  If atrim(rst!numpressupost) = "PROVA" Then GoTo fi
  If atrim(rst!grupdeclient) = "INPLACSA" Then GoTo fi
  If atrim(rst!grupdeclient) <> "ARDO" Then
        If cadbl(rst!pvp) = 0 Then
           avisarnohihapreu numc
        End If
      Else
         If atrim(rst!numpressupost) = "" Then
          'sha desactivat despres de possar les firmes  avisarnohihapreu numc, atrim(rst!grupdeclient)
        End If
  End If
fi:
  Set rst = Nothing
      
End Sub
Sub avisarnohihapreu(numc As Double, Optional grup As String)
   Dim rutamdb As String
   Dim dbavisos As Database
   Dim dbcomandes As Database
   Dim rsta As Recordset
   Dim destinatari As String
   Dim cos As String
   Dim assumpte As String
   Exit Sub  'ARA ES CONTROLA EL PREU AMB UN ENVIAMENT DOS VEGADES AL DIA DE TOTES LES COMANANDES
   assumpte = "La comanda " + atrim(numc) + " no te preu " + IIf(grup <> "", "ni NºPressupost ", "") + "i ja està en producció."
   destinatari = "incidenciesdePVP"
   rutamdb = rutadelfitxer(cami) + "avisosincidencies.mdb"
   Set dbavisos = DBEngine.OpenDatabase(rutamdb)
   Set dbcomandes = DBEngine.OpenDatabase(rutadelfitxer(cami) + "comandes.mdb", , True)
   Set rst = dbcomandes.OpenRecordset("SELECT  comandes.comanda, clients.codi, clients.nom, comandes.refclient, comandes.marcailinia FROM clients INNER JOIN comandes ON clients.codi = comandes.client where comanda=" + atrim(numc))
   Set rsta = dbavisos.OpenRecordset("select * from envios_mails where assumpte='" + atrim(assumpte) + "'")
   If rsta.EOF And Not rst.EOF Then
      cos = Chr(13) + Chr(10) + "Codi Client: " + atrim(rst!codi) + " - " + atrim(rst!nom) + Chr(13) + Chr(10) + "Ref.Client: " + atrim(rst!refclient) + Chr(10) + Chr(13) + "Texte Imp.: " + atrim(rst!marcailinia)
      dbavisos.Execute "insert into envios_mails (data,destinatari,assumpte,cos) values (now,'" + destinatari + "','" + atrim(assumpte) + "','" + atrim(cos) + "')"
   End If
   Set rsta = Nothing
   Set rst = Nothing
   dbavisos.Close
   Set dbavisos = Nothing
   Set dbcomandes = Nothing
End Sub


'========== Codigo realizado por CULD ==========
'============= culd_@hotmail.com ===============
'La funcion "EAN13_Valido" devuelve si el codigo
'control del EAN13 es VALIDO...
'El algoritmo utilizado es el descrito en la
'siguiente pagina web
'http://latecladeescape.com/w0/recetas-algoritmicas/validar-codigos-ean.html
'La function "EAN13_Control" devuelve el numero de
'control correspondiente para un codigo EAN13 de
'12 digitos (asi devuelve el control que seria el 13)
'===============================================
 
Public Function EAN13_Valido(Codigo As String) As Boolean
'Variables a utilizar
Dim X As Integer
Dim SumaPar As Integer
Dim SumaImpar As Integer
Dim Resto As Integer
Dim Control As Integer
 
'Comprobar que el código tiene 13 dígitos. De no ser así, no es correcto.
If Len(Codigo) <> 13 Then
    EAN13_Valido = False
    Exit Function
End If
 
'Sumar los dígitos de lugares pares por un lado y los de los impares por otro, pero sin incuir el último dígito.
For X = 1 To 12
    If X Mod 2 = 0 Then
        SumaPar = SumaPar + CInt(Mid(Codigo, X, 1))
    Else
        SumaImpar = SumaImpar + CInt(Mid(Codigo, X, 1))
    End If
Next X
 
'multiplicar la suma de los pares por 3.
SumaPar = SumaPar * 3
 
'Sumar el resultado de los pares y el de los impares y hallar el resto de la división por 10.
Resto = (SumaPar + SumaImpar) Mod 10
 
'Realizar la operación 10 menos ese resto y ese es el dígito de control
Control = 10 - Resto
 
'Si como resultado sale 10, entenderemos que el dígito de control es 0.
If Control = 10 Then
    If CInt(Right(Codigo, 1)) = 0 Then
        EAN13_Valido = True
        Exit Function
    Else
        EAN13_Valido = False
        Exit Function
    End If
End If
 
'Comprobar que el dígito de control que hemos calculado y el último dígito del código EAN coinciden
If CInt(Right(Codigo, 1)) = Control Then
    EAN13_Valido = True
    Exit Function
Else
    EAN13_Valido = False
    Exit Function
End If
End Function
 
Public Function EAN13_Control(Codigo As String) As Integer
'Variables a utilizar
Dim X As Integer
Dim SumaPar As Integer
Dim SumaImpar As Integer
Dim Resto As Integer
Dim Control As Integer
 
'Comprobar que el código tiene 12 dígitos. De no ser así, no es correcto.
'devuelvo un numero mayor a 9
If Len(Codigo) <> 12 Then
    EAN13_Control = 10
    Exit Function
End If
 
'Sumar los dígitos de lugares pares por un lado y los de los impares por otro, pero sin incuir el último dígito.
For X = 1 To 12
    If X Mod 2 = 0 Then
        SumaPar = SumaPar + CInt(Mid(Codigo, X, 1))
    Else
        SumaImpar = SumaImpar + CInt(Mid(Codigo, X, 1))
    End If
Next X
 
'multiplicar la suma de los pares por 3.
SumaPar = SumaPar * 3
 
'Sumar el resultado de los pares y el de los impares y hallar el resto de la división por 10.
Resto = (SumaPar + SumaImpar) Mod 10
 
'Realizar la operación 10 menos ese resto y ese es el dígito de control
Control = 10 - Resto
 
'Si como resultado sale 10, entenderemos que el dígito de control es 0.
'de lo contrario, el control es el numero que salio
If Control = 10 Then
    EAN13_Control = 0
Else
    EAN13_Control = Control
End If
End Function
Function cabool(valor As Variant) As Boolean
  If IsNull(valor) Then valor = False
  If atrim(valor) = "" Then valor = False
  If valor = "Sí" Or valor = "S" Then valor = True
  If valor = "No" Or valor = "N" Then valor = False
  If valor = "1" Or valor = "-1" Then valor = True
  If valor = "0" Then valor = False
  If valor Then
    cabool = True
   Else: cabool = False
  End If
End Function

Sub exportarllistatapdf(vllistat As CrystalReport, vnomfitxerRPT As String, vnumc As Double, vcarpetadesti As String)
  Dim a As ReportObjects
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim camp As TextObject
  Dim f  As OLEObject
  Dim vformula As String
  Dim i As Byte
  Dim vcopies As Byte
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(vnomfitxerRPT, 1)
  For i = 1 To oreport.Database.Tables.Count
    oreport.Database.Tables.Item(i).Location = vllistat.DataFiles(0)
  Next i
  'oreport.RecordSelectionFormula = "{Llaunes.numllauna}='" + UCase(atrim(numllauna)) + "'"
  'oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
  'oreport.DiscardSavedData
  convertirformules oreport, vllistat
'  oreport.DisplayProgressDialog = FalsE
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat
  oreport.ExportOptions.DiskFileName = vcarpetadesti + "\" + atrim(vnumc) + "_BaixaImpresores.pdf"
  oreport.ExportOptions.PDFExportAllPages = True
  oreport.Export False
  For i = 1 To vllistat.PrinterCopies
     oreport.PrintOut False
     wait 1
  Next i
  
End Sub
Sub convertirformules(oreport As CRAXDDRT.Report, vllistat As CrystalReport)
  Dim i As Byte
  Dim vn As String
  Dim vv As String
  Dim v As String
  i = 0
  While vllistat.Formulas(i) <> ""
     v = vllistat.Formulas(i)
     vn = Mid(v, 1, InStr(1, v, "=") - 1)
     vv = Mid(v, InStr(1, v, "=") + 1)
     oreport.FormulaFields.GetItemByName(vn).Text = vv
     i = i + 1
  Wend
End Sub
Sub obrirtancar_rele(vDispositiu As String, vobrirotancar As String, vporta As String)
  Shell llegir_ini("General", "rutallistats", "comandes.ini") + "usb_rele\usb_rele.exe " + vDispositiu + " " + vobrirotancar + " " + vporta, vbHide
  'Clipboard.Clear
  'Clipboard.SetText llegir_ini("General", "rutallistats", "comandes.ini") + "usb_rele.exe " + vdispositiu + " " + vobrirotancar + " " + vporta
End Sub
Sub sonar_sirena(vTipus As String)
 Dim vDispositiu As String
 If LCase(arguments(1)) = "desbobinadors" Then
        escriure_ini "Impresores_Compartida", "SonarSirena_maq_" + atrim(nummaq), vTipus, rutadelfitxer(cami) + "valorsprograma.ini"
        Exit Sub
 End If
 vDispositiu = llegir_ini("General", "ReleSirena1", "comandes.ini")
 If vDispositiu = "{[}]" Or vDispositiu = "" Then
    vDispositiu = InputBox("Falta el ID del RelèUSB.", "Relé de la Sirena")
    escriure_ini "General", "ReleSirena1", vDispositiu, "comandes.ini": Exit Sub
 End If
 If vDispositiu = "" Then Exit Sub
 If vTipus = "tancar" Then obrirtancar_rele vDispositiu, "close", "01"
 If vTipus = "intermitent" Then
    obrirtancar_rele vDispositiu, "open", "01"
    Sleep 200
    obrirtancar_rele vDispositiu, "close", "01"
    Sleep 200
    obrirtancar_rele vDispositiu, "open", "01"
    Sleep 200
    obrirtancar_rele vDispositiu, "close", "01"
    Sleep 200
    obrirtancar_rele vDispositiu, "open", "01"
    Sleep 200
    obrirtancar_rele vDispositiu, "close", "01"
 End If
 If vTipus = "continuu" Then
     obrirtancar_rele vDispositiu, "open", "01"
     wait 4
     obrirtancar_rele vDispositiu, "close", "01"
 End If
  If vTipus = "unpitu" Then
     obrirtancar_rele vDispositiu, "open", "01"
     wait 1
     obrirtancar_rele vDispositiu, "close", "01"
 End If

End Sub
Sub obrestocks(Optional noobrirbd As Boolean)
 Dim camistocks As String
' Set ws = DBEngine.CreateWorkspace("", "admin", "")
 ' If estaobertstocks Then dbtemp.Execute "delete * from selecciobobentrada": Exit Sub
camistocks = llegir_ini("General", "ruta_stocks", "comandes.ini")
'If camistocks = "{[}]" Then camistocks = "\\Ser2\documentos\Stock Reclamaciones\Estoc inplacsa.mdb"
'If Not existeix(camistocks) Then
'    MsgBox "Error obrint la la base de dades de Estocs (Palets) intentarem obrir la BD per defecte", vbCritical, "Error"
'    camistocks = "\\serverprodu\dades\progcomandes\dades\palets.mdb"
'End If

If camistocks = "{[}]" Then escriure_ini "General", "ruta_stocks", rutadelfitxer(cami) + "palets.mdb", "comandes.ini"
camistocks = llegir_ini("General", "ruta_stocks", "comandes.ini")
If Not noobrirbd Then
   Set dbstocks = OpenDatabase(camistocks)
 '  dbtemp.Execute "delete * from selecciobobentrada"
End If
  
End Sub

Function comprovarsilabobinaesvalida(vbobina As String, verror As String, vnumc As Double) As Boolean
    Dim vresp As String
    Dim vpalet As Double
    Dim vbob As Double
    Dim rst As Recordset
    Dim vgrup As Double
    convertirScanambPaletiBobina vbobina, vpalet, vbob
    vbobina = Trim(vpalet) + "/" + Trim(vbob)
    Set rst = dbtmp.OpenRecordset("SELECT comandes_extres.materialexacte, comandes.materialex, comandes_extres.comanda FROM comandes_extres INNER JOIN comandes ON comandes_extres.comanda = comandes.comanda where comandes_extres.comanda=" + atrim(vnumc))
   ' vpalet = cadbl(Mid(" " + vbobina, 1, InStr(1, vbobina + "  ", "/")))
   ' vbob = cadbl(Mid(vbobina, InStr(1, vbobina + "  ", "/") + 1))
    obrestocks
    valorsdajust vnumc, vgrup, "", 0
    If rst!materialexacte Then
      If Not comprovar_materialexacte(vpalet, cadbl(rst!materialex)) Then verror = "Aquest material no es exactament el que demana el client.": GoTo fi
    End If
    vresp = comprovarsieselmateixmaterial(vpalet, vbob, vnumc, vgrup, 0, True)
    
    If InStr(1, vresp, "#materialerror") > 0 Then
       If Not estadinspackinglistogrup(atrim(vpalet) + "/" + atrim(vbob), atrim(vnumc), atrim(vgrup)) Then
          verror = "Aquest material no coincideix amb el de la comanda."
           Else: verror = ""
       End If
    End If
fi:
    Set rst = Nothing
    If verror = "" Then comprovarsilabobinaesvalida = True
End Function

Function comprovar_materialexacte(vpalet As Double, vcodimaterialexacte As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbstocks.OpenRecordset("select codimatprognou from palets where idpalet=" + atrim(vpalet))
   If Not rst.EOF Then
      If cadbl(rst!codimatprognou) = cadbl(vcodimaterialexacte) Then comprovar_materialexacte = True
   End If
   Set rst = Nothing
End Function
Function valorsdajust(numc As Double, grupdestoc As Double, vtexte As String, vgrupmaterialcompatible As Double) As String
  Dim rstopcions As Recordset
  Dim rstgrup As Recordset
  Dim t As String
  Dim sisaj As Byte
  
   Set rstopcions = dbstocks.OpenRecordset("select * from opcionsdajust where comanda=" + atrim(numc))
   If Not rstopcions.EOF Then
     sisaj = atrim(cadbl(rstopcions!sistemadajust))
     grupdestoc = cadbl(rstopcions!grupdestoc)
     If sisaj > 0 Then
      t = atrim(cadbl(rstopcions!mtrsajust)) + " Mtrs D'AJUST.  "
      If sisaj = 1 Then t = t + " S'HA D'UTILITZAR MATERIAL PER LLENÇAR."
      If sisaj = 2 Then
        If cadbl(rstopcions!grupdestoc) > 0 Then
           Set rstgrup = dbstocks.OpenRecordset("select numerogrup,nomdelgrup from grupsdepalets where numerogrup=" + atrim(cadbl(rstopcions!grupdestoc)))
           If Not rstgrup.EOF Then
            t = t + " S'HA D'UTILITZAR MATERIAL D'ESTOC DEL " + UCase(rstgrup!nomdelgrup)
            grupdestoc = cadbl(rstgrup!numerogrup)
           End If
        End If
      End If
      If sisaj = 3 And cadbl(rstopcions!paletajust) > 0 Then t = t + " S'HA D'UTILITZAR EL PALET " + atrim(rstopcions!paletajust) + " BOB: " + atrim(rstopcions!bobinaajust)
     End If
   End If
   vtexte = t
   If grupdestoc > 0 Then
    Set rsttmp = dbstocks.OpenRecordset("select codigrupmaterialscompatibles from grupsdepalets where numerogrup=" + atrim(grupdestoc))
    vgrupmaterialcompatible = cadbl(rsttmp!codigrupmaterialscompatibles)
   End If
   Set rstopcions = Nothing
   Set rstgrup = Nothing
End Function
Function comprovarsieselmateixmaterial(palet As Double, bobina As Double, numc As Double, vgrup As Double, vgrupmaterialcompatible As Double, Optional vnoavisar As Boolean) As String
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim vmsg As String
   Dim vsubfams As String
   Dim vcodimatpalet As Double
 
   Set rst = dbtmp.OpenRecordset("select materialex from comandes where comanda=" + atrim(numc))
   If rst.EOF Then Exit Function
   If vgrup = 0 Then Set rst = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rst!materialex)))
   If vgrup > 0 Then Set rst = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(vgrup))
   If rst.EOF Then Exit Function
   Set rst2 = dbstocks.OpenRecordset("select codimatprognou from palets where idpalet=" + atrim(palet))
   If rst2.EOF Then Exit Function
   vcodimatpalet = IIf(cadbl(vgrupmaterialcompatible) > 0, buscarfamiliescompatibles(cadbl(vgrupmaterialcompatible), vsubfams), rst2!codimatprognou)
   Set rst2 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(vcodimatpalet))
   If rst2.EOF Then Exit Function
   'If vsubfams <> "" Then
   vsubfams = atrim(rst2!subfamilia) + "," + vsubfams
   'If cadbl(rst!familia) <> cadbl(rst2!familia) Or cadbl(rst!subfamilia) <> cadbl(rst2!subfamilia) Or cadbl(rst!familiacol) <> cadbl(rst2!familiacol) Or cadbl(rst!subfamiliacol) <> cadbl(rst2!subfamiliacol) Or InStr(1, atrim(rst2!subfamilia) + ",", vsubfams) = 0 Then
   If cadbl(rst!familia) <> cadbl(rst2!familia) Or cadbl(rst!familiacol) <> cadbl(rst2!familiacol) Or cadbl(rst!subfamiliacol) <> cadbl(rst2!subfamiliacol) Or InStr(1, vsubfams, atrim(rst!subfamilia) + ",") = 0 Then
       If Not vnoavisar Then
         If MsgBox("Les families del material de comanda i les d'aquesta bobina no coinideixen." + Chr(10) + "VOLS CONTINUAR UTILITZANT AQUEST MATERIAL?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
            comprovarsieselmateixmaterial = InputBox("Entra una explicació perquè utilitzes aquest material i l'assignat.", "Comentari"): GoTo fi
             Else: comprovarsieselmateixmaterial = "#materialerror"
          End If
           Else: comprovarsieselmateixmaterial = "#materialerror"
       End If
   End If
   Set rst = dbstocks.OpenRecordset("select * from parcials where comanda='" + atrim(numc) + "' and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
   If rst.EOF Then comprovarsieselmateixmaterial = comprovarsieselmateixmaterial + " #noesdelpackinglist"
fi:
   Set rst = Nothing
   
End Function

Function buscarfamiliescompatibles(vgrupcompatible As Double, vsubfams As String) As Double
   Dim rst As Recordset
   Dim vsql As String
   If vgrupcompatible = 0 Then Exit Function
   vsql = "SELECT grupsmaterialscompatibles.*, grupsmaterialscompatibles_linies.* "
   vsql = vsql + " FROM grupsmaterialscompatibles INNER JOIN grupsmaterialscompatibles_linies ON grupsmaterialscompatibles.numerodegrup = grupsmaterialscompatibles_linies.idgrupsdematerialscompatibles"
   vsql = vsql + " where grupsmaterialscompatibles.numerodegrup=" + atrim(vgrupcompatible)
   Set rst = dbstocks.OpenRecordset(vsql)
   If Not rst.EOF Then buscarfamiliescompatibles = cadbl(rst!codimaterialprincipal)
   vsubfams = ""
   While Not rst.EOF
     vsubfams = vsubfams + IIf(vsubfams <> "", ",", "") + atrim(rst!codisubfamilia)
     rst.MoveNext
   Wend
   Set rst = Nothing
End Function
Function estadinspackinglistogrup(vbobina As String, vnumc As String, vgrup As String) As Boolean
    Dim vresp As String
    Dim vpalet As Double
    Dim vbob As Double
    Dim rst As Recordset
   ' vgrup = cadbl(form1.veuregrupsdestoc.tag)
   ' vnumc = cadbl(form1.comanda.text)
    convertirScanambPaletiBobina vbobina, vpalet, vbob
    vbobina = Trim(vpalet) + "/" + Trim(vbob)
    obrestocks
    Set rst = dbstocks.OpenRecordset("select * from parcials where idbobina=" + atrim(vbob) + " and idpalet=" + atrim(vpalet) + " and comanda='" + atrim(IIf(vgrup > 0, vgrup, vnumc)) + "'")
    If Not rst.EOF Then estadinspackinglistogrup = True
 End Function

Sub convertirScanambPaletiBobina(vcodi As String, vpalet As Double, vbob As Double)
   Dim vcont As Double
   vcodi = atrim(vcodi)
   While vcont < Len(vcodi)
      If Not IsNumeric(Mid(vcodi, vcont + 1, 1)) Then
        vpalet = cadbl(Mid(vcodi, 1, vcont))
        If Len(vcodi) >= vcont + 2 Then vbob = cadbl(Mid(vcodi, vcont + 2))
        GoTo sortir
      End If
      vcont = vcont + 1
   Wend
sortir:
End Sub
