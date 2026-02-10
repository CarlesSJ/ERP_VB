Attribute VB_Name = "compartittintes"
Global conODBC As DAO.Connection
Global wsODBC As DAO.Workspace



Function substituir(cadena As String, buscar As String, canviar As String) As String
   If buscar = canviar Then GoTo fi
   cadena = "  " + cadena
   While InStr(1, cadena, buscar) > 0
    comença = InStr(1, cadena, buscar) - 1
    If comença < 1 Then substituir = cadena: Exit Function
    acaba = comença + Len(buscar) + 1
    cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   Wend
fi:
   substituir = atrim(cadena)
   'MsgBox linia
End Function


Sub actualitzar_estocdecomponents()
  Dim rstink As Recordset
  Dim rst As Recordset
  'SELECT dbo_tblComponenti.IdComponente, dbo_tblComponenti.StockA FROM dbo_tblComponenti;
  Set rstink = conODBC.OpenRecordset("SELECT dbo.tblComponenti.IdComponente, dbo.tblComponenti.StockA,dbo.tblComponenti.BatchCodeA FROM dbo.tblComponenti;")
  Set rst = dbtintes.OpenRecordset("select * from componentsbase")
  While Not rstink.EOF
     rst.FindFirst "idcomponent=" + atrim(rstink!idcomponente)
     If Not rst.NoMatch Then
       rst.Edit
       rst!estocactual = cadbl(rstink!StockA) / 1000
       rst.Update
       passarllaunaasalaiactualitzarKG atrim(rstink!BatchCodeA), Redondejar(cadbl(rst!estocactual), 1)
     End If
     rstink.MoveNext
  Wend
End Sub
Sub passarllaunaasalaiactualitzarKG(vnumlotinkmaker As String, vestocactual As Double)
  Dim rstll As Recordset
  vnumlotinkmaker = UCase(vnumlotinkmaker)
  Set rstll = dbtintes.OpenRecordset("select * from llaunes where numllauna='" + atrim(vnumlotinkmaker) + "'")
  If Not rstll.EOF Then
    If rstll!situacio <> "S" Then
     dbtintes.Execute "insert into historialsituacions (data,situacio,numllauna) values (now,'SALA','" + atrim(vnumlotinkmaker) + "')"
     rstll.Edit
     rstll!situacio = "SALA"
     rstll.Update
    End If
    If rstll!capacitatactual <> vestocactual Then
'        rstll.Edit
'        rstll!capacitatactual = vestocactual
'        rstll.Update
       modificarcapacitatllaunaaldosificadordeinkmaker rstll!id, atrim(rstll!numllauna), cadbl(rstll!capacitatactual), vestocactual
    End If
  End If
  Set rstll = Nothing

End Sub

Sub modificarcapacitatllaunaaldosificadordeinkmaker(idnumllauna As Long, vllauna As String, vkgactuals As Double, kg As Double)
  Dim rsthistoria As Recordset
  Set rsthistoria = dbtintes.OpenRecordset("select * from historiallauna where tipusmoviment='I' and comanda=0 and idnumllauna=" + atrim(idnumllauna) + " order by id desc")
  If rsthistoria.EOF Then
     rsthistoria.AddNew
      Else: rsthistoria.Edit
  End If
  rsthistoria!idnumllauna = idnumllauna
  rsthistoria!Data = Now
  rsthistoria!numrecarrega = numproximarecarrega(vllauna, True)
  rsthistoria!tipusmoviment = "I"
  rsthistoria!comanda = 0
  rsthistoria!formula = ""
  rsthistoria!kg = vkgactuals - kg
  rsthistoria.Update
  calcularkgdisponiblesllauna vllauna
End Sub

Sub barrejardosllaunes(Optional numllaunaperbuidar As String, Optional numllaunaaonbuidar As String, Optional kgabarrejar As Double)
  Dim idhistoriabarreja As Long
  Dim desctintes As String
  Dim rstll As Recordset
  Dim rsthistoria As Recordset
  Dim rstlldesti As Recordset
  Dim vdemanarkg As Boolean
  If numllaunaperbuidar = "" And numllaunaaonbuidar = "" And kgabarrejar = 0 Then vdemanarkg = True
  If atrim(numllaunaperbuidar) = "" Then
    numllaunaperbuidar = InputBox("Entra el Nº de llauna que vols buidar.", "Barrejar Llaunes")
    If atrim(numllaunaperbuidar) = "" Then Exit Sub
  End If
  Set rstll = dbtintes.OpenRecordset("SELECT Llaunes.*, tintes.descripcio FROM tintes LEFT JOIN Llaunes ON tintes.idtinta = Llaunes.idtinta where numllauna='" + atrim(numllaunaperbuidar) + "'")
  If rstll.EOF Then MsgBox "Aquesta llauna no existeix.", vbCritical, "Atenció": Exit Sub
  If Not rstll!activa Then MsgBox "Aquesta llauna no està activa no pots barrejar-la", vbCritical, "Atenció": Exit Sub
  If atrim(numllaunaaonbuidar) = "" Then
    numllaunaaonbuidar = InputBox("Ara entra el Nº de llauna on vols buidar la " + UCase(numllaunaperbuidar), "Barrejar Llaunes")
    If atrim(numllaunaaonbuidar) = "" Then Exit Sub
  End If
  Set rstlldesti = dbtintes.OpenRecordset("SELECT Llaunes.*, tintes.descripcio FROM tintes LEFT JOIN Llaunes ON tintes.idtinta = Llaunes.idtinta where numllauna='" + atrim(numllaunaaonbuidar) + "'")
  If rstlldesti.EOF Then MsgBox "Aquesta llauna no existeix.", vbCritical, "Atenció": Exit Sub
  If Not rstlldesti!activa Then MsgBox "Aquesta llauna no està activa no pots barrejar-la", vbCritical, "Atenció": Exit Sub
  desctintes = Chr(10) + "Origen: " + UCase(rstll!numllauna) + " --> " + UCase(rstll!descripcio) + Chr(10)
  desctintes = desctintes + "Destí: " + UCase(rstlldesti!numllauna) + " --> " + UCase(rstlldesti!descripcio)
  If rstll!idtinta <> rstlldesti!idtinta Then
     If MsgBox("Les tintes de les dues llaunes son diferents" + Chr(10) + " Vols barrejar-les igualment?" + Chr(10) + desctintes, vbCritical, "Atenció") = vbNo Then Exit Sub
  End If
    calcularkgdisponiblesllauna rstll!numllauna
    calcularkgdisponiblesllauna rstlldesti!numllauna
    If vdemanarkg Then kgabarrejar = cadbl(InputBox("Entra els Kg que vols barrejar", "Atenció", atrim(rstll!capacitatactual)))
    Set rsthistoria = dbtintes.OpenRecordset("select * from historiallauna order by id")
    If (kgabarrejar > 0 And kgabarrejar > rstll!capacitatactual) Or kgabarrejar = 0 Then kgabarrejar = rstll!capacitatactual
    rsthistoria.AddNew
    rsthistoria!idnumllauna = rstlldesti!id
    rsthistoria!Data = Now
    rsthistoria!numrecarrega = numproximarecarrega(numllaunaaonbuidar, True)
    rsthistoria!tipusmoviment = "B"
    rsthistoria!formula = ""
    rsthistoria!kg = kgabarrejar
    rsthistoria.Update
    rsthistoria.MoveLast
    idhistoriabarreja = rsthistoria!id
    rsthistoria.AddNew
    rsthistoria!idnumllauna = rstll!id
    rsthistoria!Data = Now
    rsthistoria!numrecarrega = numproximarecarrega(numllaunaperbuidar, True)
    rsthistoria!tipusmoviment = "V"
    rsthistoria!formula = ""
    rsthistoria!idhistoriabarreja = idhistoriabarreja
    rsthistoria!kg = kgabarrejar
    rsthistoria.Update
    dbtintes.Execute "update llaunes set situacio='" + atrim(rstll!situacio) + "' where numllauna='" + atrim(rstlldesti!numllauna) + "'"
    calcularkgdisponiblesllauna rstll!numllauna
    calcularkgdisponiblesllauna rstlldesti!numllauna

  Set rsthistoria = Nothing
  Set rstll = Nothing
  Set rstlldesti = Nothing
End Sub

Sub calcularkgdisponiblesllauna(numllauna As String, Optional retornpes As Double, Optional vultimretorn As Boolean)
   Dim rstllauna As Recordset
   Dim rsthistoria As Recordset
   Dim kg As Double
   Dim vretorn As Boolean
   kg = 0
   Set rstllauna = dbtintes.OpenRecordset("select * from llaunes where numllauna='" + atrim(numllauna) + "'")
   If rstllauna.EOF Then Exit Sub
   Set rsthistoria = dbtintes.OpenRecordset("select * from historiallauna where idnumllauna=" + atrim(rstllauna!id) + " order by data")
   If Not rsthistoria.EOF Then
      'vuidat
     '' rsthistoria.FindLast "tipusmoviment='V'"
     '' If Not rsthistoria.NoMatch Then kg = rsthistoria!kg: GoTo cont
      'retorn
      rsthistoria.FindLast "tipusmoviment='R'"
      While Not rsthistoria.NoMatch And vultimretorn And kg = 0
         kg = cadbl(rsthistoria!kg)
         rsthistoria.FindPrevious "tipusmoviment='R'"
         If Not rsthistoria.NoMatch Then kg = cadbl(rsthistoria!kg)
      Wend
      If Not rsthistoria.NoMatch Then
         kg = rsthistoria!kg
         vretorn = True
      End If
   End If
   If kg = 0 And Not vretorn Then
      If Not rsthistoria.EOF Then
          rsthistoria.MoveFirst
            Else: MsgBox "Hi ha un error amb la historia d'aquesta llauna. " + atrim(numllauna), vbCritical, "Error": Exit Sub
      End If
   End If
   While Not rsthistoria.EOF
   'carrega
     If rsthistoria!tipusmoviment = "C" Then kg = kg + rsthistoria!kg
   'carrega desde reconvertir
     If rsthistoria!tipusmoviment = "K" Then kg = kg + rsthistoria!kg
   'carrega desde barreja
     If rsthistoria!tipusmoviment = "B" Then kg = kg + rsthistoria!kg
   'vuidar
     If rsthistoria!tipusmoviment = "V" Then kg = kg - rsthistoria!kg
   'gastat a impresores
     If rsthistoria!tipusmoviment = "I" Or rsthistoria!tipusmoviment = "L" Then kg = kg - rsthistoria!kg
     rsthistoria.MoveNext
   Wend
cont:
   rstllauna.Edit
   rstllauna!capacitatactual = Redondejar(kg, 1)
   retornpes = rstllauna!capacitatactual
   If kg = 0 Then rstllauna!activa = False: rstllauna!aimpresores = False
   rstllauna.Update
End Sub
Sub ferelretorndetinta(vnllauna As String, vpesnet As Double, Optional novulletiqueta As Boolean, Optional vdatainventari As String)
    Dim rsthistoria As Recordset
    Dim rstll As Recordset
    Dim vretornpes As Double

    Set rstll = dbtintes.OpenRecordset("select * from llaunes where numllauna='" + atrim(vnllauna) + "'")
    If rstll.EOF Then MsgBox "Error... no he trobat aquesta llauna.": Exit Sub
    
    Set rsthistoria = dbtintes.OpenRecordset("select * from historiallauna")
    rsthistoria.AddNew
    rsthistoria!idnumllauna = rstll!id
    rsthistoria!numrecarrega = recarregamesgran(rstll!id)
    rsthistoria!Data = Now
    rsthistoria!tipusmoviment = "R"
    rsthistoria!formula = ""
    rsthistoria!kg = Redondejar(cadbl(vpesnet), 1)
    rsthistoria!datainventari = vdatainventari
    rsthistoria.Update
    calcularkgdisponiblesllauna rstll!numllauna, vretornpes
    'If vretornpes < 1 Then
   '    If MsgBox("Aquesta llauna està buida... Et cal etiqueta?", vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then GoTo noetiqueta
   ' End If
    If Not novulletiqueta Then
      imprimir_etiqueta rstll!numllauna
    End If
noetiqueta:
    Set rstll = Nothing
    Set rsthistoria = Nothing
End Sub
Sub fercompraEstocminim(vnumllauna As String)
   Dim rst As Recordset
'   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.idtinta, Sum(Llaunes.capacitatactual) AS totalkg From Llaunes Where llaunes.capacitatactual>0 and Llaunes.activa = True and Llaunes.idtinta=" + atrim(vidtinta) + " GROUP BY Llaunes.idtinta;")
'   Set rst2 = dbtintes.OpenRecordset("SELECT count(Llaunes.idtinta) as Tllaunes from Llaunes Where llaunes.capacitatactual>0.9 and Llaunes.activa = True and Llaunes.idtinta=" + atrim(vidtinta) + " GROUP BY Llaunes.idtinta;")
 '  If Not rst.EOF Then etkgtotals = atrim(rst2!Tllaunes) + " Llaunes" + Chr(10) + atrim(Redondejar(rst!totalkg, 1)) + " Kg"
      
   
   Set rst = Nothing
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
    
    Open sFName For Binary Access Read As #5
    ReDim z(FileLen(sFName))
    Get #5, , z()
     fldTO.AppendChunk z
    Close #5
    
    'Delete the file
    'Kill (sFName)
    
End Function
Function buscarcodiformula(numllauna As String) As String
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, tintes.codi, tintes.descripcio, historiallauna.formula, Llaunes.situacio, historiallauna.data FROM tintes LEFT JOIN (Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) ON tintes.idtinta = Llaunes.idtinta Where (((historiallauna.tipusmoviment) = 'C' or (historiallauna.tipusmoviment) = 'K') And ( numllauna='" + atrim(numllauna) + "')) ORDER BY historiallauna.tipusmoviment;")
   'Clipboard.Clear
   'Clipboard.SetText "SELECT Llaunes.numllauna, tintes.codi, tintes.descripcio, historiallauna.formula, Llaunes.situacio, historiallauna.data FROM tintes LEFT JOIN (Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) ON tintes.idtinta = Llaunes.idtinta Where (((historiallauna.tipusmoviment) = 'C' or (historiallauna.tipusmoviment) = 'K') And ( numllauna='" + atrim(numllauna) + "')) ORDER BY historiallauna.tipusmoviment;"
   If Not rst.EOF Then buscarcodiformula = atrim(rst!formula)
   Set rst = Nothing
End Function

Sub imprimir_etiqueta(numllauna As String)
 ' Dim rst As Recordset
  Dim a As ReportObjects
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim camp As TextObject
  Dim f  As OLEObject
  Dim vformula As String
  Dim vcopies As Byte
  vcopies = 1
  If nomdelmaterialdelcontenidor(numllauna) <> "" Then vcopies = 2
  vformula = buscarcodiformula(numllauna)
  crearcodibarresaltemp UCase(numllauna)
  crearcodibarresdelaformulaaltemp vformula
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "etiqueta_llaunes_A5.rpt", 1) '"etiqueta_llaunes.rpt"
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "tintes.mdb"
  oreport.RecordSelectionFormula = "{Llaunes.numllauna}='" + UCase(atrim(numllauna)) + "'"
  oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
  oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
 ' oreport.Sections("D").ReportObjects.Item("serie2").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
'  oreport.Sections("D").ReportObjects.Item("matriculacontenidor1").BackColor = QBColor(5)
  oreport.DiscardSavedData
  oreport.Sections("D").ReportObjects.Item("recuperador").Suppress = True
  oreport.Sections("D").ReportObjects.Item("numllauna2").Suppress = True
  possarformulesdetall oreport, numllauna
  
  oreport.FormulaFields.GetItemByName("numllaunarecuperador").text = "'" + UCase(atrim(numllauna)) + "'"
  If vformula <> "" Then oreport.FormulaFields.GetItemByName("vformula").text = "'(" + vformula + ")'"
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
    Else
      oreport.DisplayProgressDialog = False
      
      oreport.PrintOut False, 1
      If CInt(vcopies) > 1 Then
      
          oreport.Sections("D").ReportObjects.Item("recuperador").Suppress = True
          oreport.Sections("D").ReportObjects.Item("numllaunarecuperador1").Suppress = False
          oreport.Sections("D").ReportObjects.Item("vmatriculacontenidor2").Suppress = False
          oreport.Sections("D").ReportObjects.Item("numllauna2").Suppress = False
          wait 1
          oreport.PrintOut False, 1
      End If
              
  End If
  
End Sub
Sub possarformulesdetall(oreport As CRAXDDRT.Report, numllauna As String)
  Dim i As Byte
  Dim rst As Recordset
  Dim rstn As Recordset
  Dim detall As String
  Dim vllauna As String
  Dim vnumlotinplacsa As String
  For i = 1 To 5
    oreport.FormulaFields.GetItemByName("Detall" + atrim(i)).text = "''"
  Next i
  Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, historiallauna.* FROM Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna WHERE (((Llaunes.numllauna)='" + atrim(numllauna) + "')) order by data Desc;")
  If rst.EOF Then Exit Sub
  If cadbl(rst!id) = 0 Then Exit Sub
  'rst.MoveLast
  For i = 1 To 5
    If Not rst.EOF And Not rst.BOF Then
       If atrim(cadbl(rst!idhistoriabarreja)) > 0 Then
            Set rstn = dbtintes.OpenRecordset("select numllauna from llaunes where id=(select idnumllauna from historiallauna where id=" + atrim(cadbl(rst!idhistoriabarreja)) + ")")
            If Not rstn.EOF Then vllauna = "  ->" + rstn!numllauna
       End If

       Set rstn = dbtintes.OpenRecordset("select numllauna from llaunes where id=(select idnumllauna from historiallauna where idhistoriabarreja=" + atrim(cadbl(rst!id)) + ")")
       If Not rstn.EOF Then vllauna = "  <-" + rstn!numllauna
       
       detall = Format(rst!Data, "dd/mm/yy hh:nn") + " - " + atrim(rst!tipusmoviment) + " - " + atrim(IIf(rst!comanda > 0, rst!comanda, "")) + "   " + atrim(Redondejar(rst!kg, 1)) + " Kg" + vllauna
       oreport.FormulaFields.GetItemByName("Detall" + atrim(i)).text = "'" + detall + "'"
       vllauna = ""
       rst.MoveNext
    End If
  vnumlotinplacsa = buscarlotinplacsadelallauna(numllauna)
  Next i
  oreport.FormulaFields.GetItemByName("numerolotinplacsa").text = "'" + vnumlotinplacsa + "'"
  Set rst = Nothing
  
End Sub
Function nomdelmaterialdelcontenidor(vnumllauna As String) As String
  Dim rst As Recordset
  vsql = "SELECT Llaunes.numllauna, Contenidors_material.descripcio, Llaunes.idmaterialcontenidor FROM Llaunes LEFT JOIN Contenidors_material ON Llaunes.idmaterialcontenidor = Contenidors_material.codi where numllauna='" + vnumllauna + "'"
  Set rst = dbtintes.OpenRecordset(vsql)
  If Not rst.EOF Then
      nomdelmaterialdelcontenidor = atrim(rst!descripcio)
  End If
  Set rst = Nothing
End Function
Function buscarlotinplacsadelallauna(vnumllauna As String) As String
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, historiallaunalots.numlotbase, historiallaunalots.idcomponent FROM (historiallauna RIGHT JOIN Llaunes ON historiallauna.idnumllauna = Llaunes.id) LEFT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria WHERE (((historiallaunalots.idcomponent)=0) and numllauna='" + atrim(vnumllauna) + "'" + ");")
   If Not rst.EOF Then
      buscarlotinplacsadelallauna = atrim(rst!numlotbase)
   End If
End Function
Sub creardirecotoritemp()
 On Error Resume Next
 MkDir "c:\temp"
End Sub

Sub crearcodibarresaltemp(numcodi As String)
  Dim rst As Recordset
   Dim vinici As Date
   If existeix("c:\temp\CBLlauna.bmp") Then Kill "c:\temp\CBLlauna.bmp"
   creardirecotoritemp
   If atrim(numcodi) = "" Then Exit Sub
   
   '  GENERA EL CODI DE BARRES DEL NUMERO DE TREBALL
   escriure_ini "Tbarcode", "nomfitxer", "c:\temp\CBLlauna.bmp", "generartbarcode.ini"
   escriure_ini "Tbarcode", "pixelsample", "1000", "generartbarcode.ini"
   escriure_ini "Tbarcode", "pixelsalt", "800", "generartbarcode.ini"
   escriure_ini "Tbarcode", "text", atrim(numcodi), "generartbarcode.ini"
   escriure_ini "Tbarcode", "printdatatext", "1", "generartbarcode.ini"
   escriure_ini "Tbarcode", "tipusbarcode", "62", "generartbarcode.ini"
   Shell llegir_ini("General", "rutallistats", "comandes.ini") + "generarimatgedecodidebarres.exe"
   '62 es full asci
   '13 as ean 13
   
   
   
   'controlcodidebarres.PrintDataText = True
   'controlcodidebarres.Enabled = True
   'controlcodidebarres.Text = numcodi
   'controlcodidebarres.SaveImage "c:\temp\CBLlauna", eIMBmp, 1000, 800, 600, 600
   vinici = Now
   While Not existeix("c:\temp\CBLlauna.bmp") And DateDiff("s", vinici, Now) < 5
     DoEvents
   Wend
   dbtintes.Execute "delete * from generarcodidebarres"
   Set rst = dbtintes.OpenRecordset("select * from generarcodidebarres")
    rst.AddNew
    copiafoto "c:\temp\CBLlauna.bmp", rst!codidebarres
    rst!numllauna = numcodi
   rst.Update
End Sub
Sub crearcodibarresdelaformulaaltemp(numcodi As String)
  Dim rst As Recordset
  Dim vinicia As Date
   If atrim(numcodi) = "" Then Exit Sub
   If existeix("c:\temp\CBLlaunaformula.bmp") Then Kill "c:\temp\CBLlaunaformula.bmp"
   
    '  GENERA EL CODI DE BARRES DEL NUMERO DE TREBALL
   escriure_ini "Tbarcode", "nomfitxer", "c:\temp\CBLlaunaformula.bmp", "generartbarcode.ini"
   escriure_ini "Tbarcode", "pixelsample", "2000", "generartbarcode.ini"
   escriure_ini "Tbarcode", "pixelsalt", "500", "generartbarcode.ini"
   escriure_ini "Tbarcode", "text", atrim(numcodi), "generartbarcode.ini"
   escriure_ini "Tbarcode", "printdatatext", "0", "generartbarcode.ini"
   escriure_ini "Tbarcode", "tipusbarcode", "62", "generartbarcode.ini"
   Shell llegir_ini("General", "rutallistats", "comandes.ini") + "generarimatgedecodidebarres.exe"
   '62 es full asci
   '13 as ean 13
   
   'controlcodidebarres.PrintDataText = False
   'controlcodidebarres.Enabled = True
   'controlcodidebarres.Text = numcodi
   'controlcodidebarres.SaveImage "c:\temp\CBLlaunaformula", eIMBmp, 2000, 500, 600, 600
   vinici = Now
   While Not existeix("c:\temp\CBLlaunaformula.bmp") And DateDiff("s", vinici, Now) < 5
     DoEvents
   Wend
   Set rst = dbtintes.OpenRecordset("select * from generarcodidebarres")
   If rst.EOF Then Exit Sub
    rst.Edit
    copiafoto "c:\temp\CBLlaunaformula.bmp", rst!codidebarresformula
   rst.Update
End Sub
Function posarcolorserie(numllauna As String) As Double
  Dim rst As Recordset
  Dim nomcolor As String
  Dim numcolor As Double
  posarcolorserie = QBColor(15)
  Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, colorsetiquetes.codicolor FROM ((subfamiliestintes RIGHT JOIN tintes ON subfamiliestintes.codi = tintes.idsubfamilia) RIGHT JOIN Llaunes ON tintes.idtinta = Llaunes.idtinta) LEFT JOIN colorsetiquetes ON subfamiliestintes.color = colorsetiquetes.nomcolor where numllauna='" + atrim(numllauna) + "';")
  If rst.EOF Then Exit Function
  numcolor = cadbl(rst!codicolor)
  If numcolor = 0 Then numcolor = 15
  posarcolorserie = QBColor(numcolor)
  
End Function

Function triarreferencia(idtinta As Long) As Long
  Load formseleccio
  formseleccio.caption = "Selecciona referencia proveidor"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "select id,referencia,nomproveidor from tintesreferencies where idtinta=" + atrim(idtinta) + " order by predeterminada"
  formseleccio.refrescar
  If formseleccio.Data1.Recordset.EOF Then GoTo fi
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 1000
  formseleccio.DBGrid2.Columns(2).width = 1500
  formseleccio.Show 1
  If seleccioret = 1 Then
   triarreferencia = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
fi:
  Unload formseleccio
End Function
Function crearnovallauna(Optional iddelatinta As Long, Optional idrefproveidor As Long) As String
  Dim rstllauna As Recordset
  Dim numnovallauna As Long
  Dim rsttinta As Recordset
  If cadbl(iddelatinta) = 0 Then iddelatinta = triartinta
  If cadbl(iddelatinta) = 0 Then Exit Function
  If cadbl(idrefproveidor) = 0 Then idrefproveidor = triarreferencia(iddelatinta)
  Set rstllauna = dbtintes.OpenRecordset("select numllauna from contadors")
  numnovallauna = rstllauna!numllauna + 1
  dbtintes.Execute "update contadors set numllauna=[numllauna]+1"
  Set rstllauna = dbtintes.OpenRecordset("select * from llaunes")
  rstllauna.AddNew
  rstllauna!numllauna = "A" + atrim(numnovallauna)
  crearnovallauna = rstllauna!numllauna
  rstllauna!idtinta = iddelatinta
  rstllauna!id_refproveidor = idrefproveidor
  rstllauna!situacio = ""
  rstllauna!activa = True
  rstllauna.Update
  Set rstllauna = Nothing
  Set rsttinta = Nothing
End Function
Function numproximarecarrega(nllauna As String, Optional ultimarecarrega As Boolean) As Long
   Dim rst As Recordset
   numproximarecarrega = 0
   Set rst = dbtintes.OpenRecordset("select max(numrecarrega) as mrec from historiallauna where idnumllauna in (select id from llaunes where numllauna='" + atrim(nllauna) + "') group by idnumllauna")
   If Not rst.EOF Then numproximarecarrega = cadbl(rst!mrec)
   If Not ultimarecarrega Or numproximarecarrega = 0 Then numproximarecarrega = numproximarecarrega + 1
   
End Function
Function recarregamesgran(idllauna As Double) As Double
    Dim rst As Recordset
    Set rst = dbtintes.OpenRecordset("Select max(numrecarrega) as lagran from historiallauna where idnumllauna=" + atrim(idllauna) + " group by idnumllauna")
    recarregamesgran = cadbl(rst!lagran)
    Set rst = Nothing
End Function

Sub actualitzarcarguescomponents()
   Dim contador As Integer
   Dim longitud As Byte
   Dim rstfink As DAO.Recordset
   Dim rstf As Recordset
   Dim inst_sql As String
   Dim vllaunavella As String
   Dim vllaunanova As String
   'Dim conODBC As DAO.Connection
   
   inst_sql = "SELECT StartEventDateAndTime, CodComponente, dbo.tblLogBookDetail.idcomponente,DescComponente, OperationQuantity, BatchCode FROM dbo.tblLogBook LEFT JOIN dbo.tblLogBookDetail ON dbo.tblLogBook.IDLogBook = dbo.tblLogBookDetail.IDLogBook "
   inst_sql = inst_sql + " Where (((dbo.tblLogBook.IDOperation) = 50)) ORDER BY dbo.tblLogBook.StartEventDateAndTime DESC;"
      
   ratoli "espera"
'   formtintes.Enabled = False
   'Set wsODBC = CreateWorkspace("", "tintes", "", dbUseODBC)
   'Set conODBC = wsODBC.OpenConnection("connexiosql", , True, "ODBC;DATABASE=InkmakerDB;UID=sa;PWD=Mak2008;DSN=tintes")
   Set rstfink = conODBC.OpenRecordset(inst_sql, dbOpenSnapshot)
   If rstfink.EOF Then MsgBox "No s'ha trobat cap component al INKMAKER.", vbCritical, "Atenció": GoTo fi
   rstfink.MoveLast
   rstfink.MoveFirst
   contador = 0
   While Not rstfink.EOF And contador < 50
      If cadbl(rstfink!OperationQuantity) >= 0 Then
        crearnovacarga rstfink, conODBC
      End If
      contador = contador + 1
      rstfink.MoveNext
   Wend
   'conODBC.Close
fi:
   Set rstf = dbtintes.OpenRecordset("select * from componentsbase")
   While Not rstf.EOF
      Set rstfink = dbtintes.OpenRecordset("select * from detallnumeroslotsbase where idcomponent=" + atrim(rstf!idcomponent) + " order by data Desc")
      If Not rstfink.EOF Then dbtintes.Execute "update  llaunes set situacio='DOS' where numllauna='" + atrim(rstfink!numerodelot) + "'"
      rstf.MoveNext
   Wend
   
   Set rstfink = Nothing
   Set rstf = Nothing
   ratoli "normal"
'   formtintes.Enabled = True
End Sub
Sub crearnovacarga(rstfink As Recordset, conODBC)
  Dim rstc As Recordset
  Dim vllaunavella As String
  Set rstc = dbtintes.OpenRecordset("select * from detallnumeroslotsbase where idcomponent=" + atrim(rstfink!idcomponente) + " order by id desc")
  If Not rstc.EOF Then vllaunavella = atrim(rstc!numerodelot)
  If Not IsDate(rstfink!StartEventDateAndTime) Then GoTo fi
  rstc.FindFirst "data=#" + Format(rstfink!StartEventDateAndTime, "mm/dd/yy hh:nn:ss") + "#"
  If rstc.NoMatch Then
    If atrim(rstfink!batchcode) <> "" Then
     rstc.AddNew
     rstc!Data = Format(rstfink!StartEventDateAndTime, "dd/mm/yy hh:nn:ss")
     rstc!idcomponent = rstfink!idcomponente
     rstc!numerodelot = rstfink!batchcode
     rstc.Update
     comprovar_llaunacoincideixambdosificador vllaunavella, atrim(rstfink!batchcode), rstfink
    End If
    
  End If
fi:
  
  Set rstc = Nothing
End Sub
Sub comprovar_llaunacoincideixambdosificador(vllaunavella As String, vllaunanova As String, rstfink As Recordset)
   Dim rstlln As Recordset
   Dim rstllv As Recordset
   Set rstllv = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, tintes.idfamilia, tintes.idsubfamilia, tintes.idfamcolor, tintes.idsubfamcolor FROM Llaunes INNER JOIN tintes ON Llaunes.idtinta = tintes.idtinta where numllauna='" + vllaunavella + "'")
   Set rstlln = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, tintes.idfamilia, tintes.idsubfamilia, tintes.idfamcolor, tintes.idsubfamcolor FROM Llaunes INNER JOIN tintes ON Llaunes.idtinta = tintes.idtinta where numllauna='" + vllaunanova + "'")
   If rstlln.EOF Or rstllv.EOF Then enviaremailgeneric "controlestoctintes", "Error de llauna al dosificador de " + atrim(rstfink!DescComponente), "La llauna nova " + atrim(vllaunanova) + " o l'anterior " + atrim(vllaunavella) + " no existeixen a la base de dades.": GoTo fi
   If rstllv!idfamilia <> rstlln!idfamilia Or rstllv!idsubfamilia <> rstlln!idsubfamilia Or rstllv!idfamcolor <> rstlln!idfamcolor Or rstllv!idsubfamcolor <> rstlln!idsubfamcolor Then
        enviaremailgeneric "controlestoctintes", "Error de llauna al dosificador de " + atrim(rstfink!DescComponente), "La llauna " + atrim(vllaunanova) + " del dosificador " + atrim(rstfink!DescComponente) + " no correspont amb l'anterior " + atrim(vllaunavella)
       'enviar missatge families equivocades per aquest dosificador
   End If
fi:
End Sub
Function treuresimbols(desc As String) As String
'   desc = substituir(desc, ":", "_")
   desc = substituir(desc, "'", "´")
   desc = substituir(desc, "|", "_")
   desc = substituir(desc, ";", "_")
   treuresimbols = desc
End Function

Sub enviaremailgeneric(destinatari As String, assumpte As String, cos As String)
   Dim dbenvio As Database
   If atrim(cos) = "" Then Exit Sub
    
   Set dbenvio = OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
   dbenvio.Execute "insert into envios_mails (data,destinatari,assumpte,cos) values (now,'" + destinatari + "','" + treuresimbols(assumpte) + "','" + treuresimbols(cos) + "')"
   Set dbenvio = Nothing
End Sub

Function preukg_actual_component(vidcomponent As Double) As String
   Dim vsql As String
   Dim rst As Recordset
   vsql = "SELECT First(detallnumeroslotsbase.numerodelot) AS numerodelot, Max(detallnumeroslotsbase.data) AS MáxDedata, First(DetallFormules.[%decomponent]) AS tanxcent FROM (Formules LEFT JOIN DetallFormules ON Formules.idformula = DetallFormules.IDFormula) LEFT JOIN detallnumeroslotsbase ON DetallFormules.IdComponente = detallnumeroslotsbase.idcomponent WHERE "
   vsql = vsql + " (((detallnumeroslotsbase.idcomponent)=" + atrim(vidcomponent) + ")) group by detallnumeroslotsbase.data order by Max(detallnumeroslotsbase.data) Desc;"
   Set rst = dbtintes.OpenRecordset(vsql)
   If Not rst.EOF Then preukg_actual_component = atrim(rst!numerodelot)
   Set rst = Nothing
End Function
Function saber_preu_kg_tinta_llauna(vnumerodelot As String, Optional vcodiformula) As Double
   Dim rst As Recordset
   Dim vsql As String
   If vnumerodelot <> "" Then saber_preu_kg_tinta_llauna = calcular_preu_kg_tinta(vnumerodelot)
   If vcodiformula <> "" Then
       'actualitzarcarguescomponents
       'buscar preu de cost calculant el %de tinta de cada component utilitzat
       vsql = "SELECT detallnumeroslotsbase.idcomponent, First(detallnumeroslotsbase.numerodelot) AS numerodelot, Max(detallnumeroslotsbase.data) AS MáxDedata, First(DetallFormules.[%decomponent]) AS [tanxcent] FROM (Formules LEFT JOIN DetallFormules ON Formules.idformula = DetallFormules.IDFormula) LEFT JOIN detallnumeroslotsbase ON DetallFormules.IdComponente = detallnumeroslotsbase.idcomponent Where (((Formules.codiformula) = '"
       vsql = vsql + vcodiformula + "')) GROUP BY detallnumeroslotsbase.idcomponent;"
       Set rst = dbtintes.OpenRecordset(vsql)
       While Not rst.EOF
          vpreukg = calcular_preu_kg_tinta(atrim(rst!numerodelot))
          If vpreukg = 0 Then vpreukg = calcular_preu_kg_tinta(preukg_actual_component(cadbl(rst!idcomponent)))
          saber_preu_kg_tinta_llauna = saber_preu_kg_tinta_llauna + (vpreukg * (cadbl(rst!tanxcent) / 100))
          If vpreukg = 0 Then
             'saber_preu_kg_tinta_llauna = calcular_preu_kg_tinta(preukg_actual_component(rst!idcomponent))
             saber_preu_kg_tinta_llauna = 0: GoTo fi
          End If
          rst.MoveNext
       Wend
       'If saber_preu_kg_tinta_llauna > 0 Then Stop
   End If
fi:
   Set rst = Nothing
   
End Function
Function preukg_delallauna(vnumllauna As String) As String
  Dim rst As Recordset
  Set rst = dbtintes.OpenRecordset("select preuxrkilo from llaunes where numllauna='" + vnumllauna + "'")
  If Not rst.EOF Then preukg_delallauna = cadbl(rst!preuxrkilo)
  Set rst = Nothing
End Function
Function calcular_preu_kg_tinta(Optional vnumerodelot As String) As Double
  ' Dim dbcompres As Database
   Dim rstcompres As Recordset
   
   If vnumerodelot <> "" Then
       If Mid(vnumerodelot, 1, 1) = "A" And (Len(vnumerodelot) < 7 And Len(vnumerodelot) > 4) Then
        calcular_preu_kg_tinta = cadbl(preukg_delallauna(vnumerodelot))
        Exit Function
       End If
      'buscar preu de cost a les compres
      'Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
      Set rstcompres = dbcompres.OpenRecordset("select * from albaransbip where numlotproveidor='" + atrim(vnumerodelot) + "'")
      If Not rstcompres.EOF Then
        If cadbl(rstcompres!quantitat) > 0 Then
         calcular_preu_kg_tinta = Redondejar(cadbl(rstcompres!preu), 2)
        End If
      End If
      Set rstcompres = Nothing
'      Set dbcompres = Nothing
   End If
  
End Function

