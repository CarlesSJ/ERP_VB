VERSION 5.00
Begin VB.Form formidproducte 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "formidproducte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Dim arguments As Variant
  fitxerini = "comandes.ini"
  cami = llegir_ini("General", "cami", fitxerini)
  ruta_relativa_docs = "\\ser2\documentos\Pautacli"
  ruta_documentacio_clixes = "\\ser2\d\documentacioclixes"
  '"c:\misdoc~1\commandes\comandes.mdb"
  If existeix("c:\ordprog.ini") Then cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
  Set dbcomandes = DBEngine.OpenDatabase(cami)
  arguments = ObtenerLíneaComando
  If cadbl(arguments(1)) > 0 Then numcomanda = atrim(cadbl(arguments(1)))
  'numcomanda = 158923
  If cadbl(numcomanda) > 0 Then crear_id_producte
  Set dbcomandes = Nothing
  End
End Sub
Sub crear_id_producte()
    Dim rstid As Recordset
    Dim rstc As Recordset
    Dim rstcextra As Recordset
    Dim rstp As Recordset
    Dim rsti As Recordset
    Dim rstm As Recordset
    Dim capes As String
    Dim espesor As Double
    Dim ample As Double
    Dim longitud As Double
    Dim id_families As Long
    Dim were As String
    id_families = 0
    Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(numcomanda)))
    If rstc.EOF Then Exit Sub
    Set rstcextra = dbcomandes.OpenRecordset("select * from comandes_extres where comanda=" + atrim(cadbl(numcomanda)))
    Set rstp = dbcomandes.OpenRecordset("select ruta from productes where codi='" + atrim(rstc!producte) + "'")
    If rstp.EOF Then Exit Sub
    Set rstm = dbcomandes.OpenRecordset("select id_familia from materials where codi=" + atrim(cadbl(rstc!materialex)))
    If Not rstm.EOF Then id_families = cadbl(rstm!id_familia)
    If InStr(1, "VPT", rstc!proximaseccio) = 0 Or rstc!producte = "PC" Or rstc!producte = "PC2" Then Exit Sub
    capes = IIf(cadbl(rstc!linkcomanda2) > 0, 3, IIf(cadbl(rstc!linkcomanda1) > 0, 2, 1))
    espesor = sumaespesor(cadbl(numcomanda), cadbl(rstc!linkcomanda1), cadbl(rstc!linkcomanda2))
    If ultimaseccio(rstp!ruta) = "R" Then
      ample = cadbl(rstc!amplereb)
    End If
    If ultimaseccio(rstp!ruta) = "S" Then
         ample = cadbl(rstc!amplesol)
         longitud = cadbl(rstc!longitudsol)
    End If
    If ultimaseccio(rstp!ruta) = "I" Then
         ample = cadbl(rstc!ampleesq)
    End If
    If ultimaseccio(rstp!ruta) = "E" Then
         ample = cadbl(rstc!ampleesq)
    End If
    If ultimaseccio(rstp!ruta) = "L" Then
         ample = cadbl(rstc!amplelaminar)
    End If
    
    were = "client=" + atrim(cadbl(rstc!client)) + " and producte='" + atrim(rstc!producte) + "' and capes=" + atrim(cadbl(capes))
    were = were + " and id_treball=" + atrim(cadbl(rstc!numtreball)) + " and espesor=" + passaradecimalpunt(cadbl(espesor)) + " and id_families=" + atrim(id_families)
    were = were + " and ample=" + passaradecimalpunt(atrim(ample)) + " and longitud=" + passaradecimalpunt(atrim(longitud))
possarid:
    Set rsti = dbcomandes.OpenRecordset("select * from idproductes where " + were)
    If Not rsti.EOF Then
        rstcextra.Edit
        rstcextra!id_inplacsa = rsti!id_inplacsa
        rstcextra.Update
        GoTo fi
    End If
    rsti.AddNew
    rsti!client = cadbl(rstc!client)
    rsti!producte = atrim(rstc!producte)
    rsti!capes = capes
    rsti!id_treball = cadbl(rstc!numtreball)
    rsti!espesor = espesor
    rsti!id_families = id_families
    rsti!ample = ample
    rsti!longitud = longitud
    rsti.Update
    GoTo possarid
    
fi:
    Set rstcextra = Nothing
    Set rstc = Nothing
    Set rstm = Nothing
    Set rstp = Nothing
    Set rsti = Nothing
    
End Sub

Function sumaespesor(numc As Double, numc2 As Double, numc3 As Double) As Double
  Dim rstc As Recordset
  Dim rstc2 As Recordset
  Dim rstc3 As Recordset
  Set rstc = dbcomandes.OpenRecordset("select espessor from comandes where comanda=" + atrim(numc))
  Set rstc2 = dbcomandes.OpenRecordset("select espessor from comandes where comanda=" + atrim(numc2))
  Set rstc3 = dbcomandes.OpenRecordset("select espessor from comandes where comanda=" + atrim(numc3))
  If Not rstc.EOF Then sumaespesor = sumaespesor + cadbl(rstc!espessor)
  If Not rstc2.EOF Then sumaespesor = sumaespesor + cadbl(rstc2!espessor)
  If Not rstc3.EOF Then sumaespesor = sumaespesor + cadbl(rstc3!espessor)
  Set rstc = Nothing
  Set rstc2 = Nothing
  Set rstc3 = Nothing
  
End Function
Function ultimaseccio(ruta As String) As String

   ultimaseccio = Mid(ruta, Len(ruta), 1)
End Function
