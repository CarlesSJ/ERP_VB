VERSION 5.00
Begin VB.Form comprovaretrebo 
   Caption         =   "Comprovar Etiqueta de Rebobinadora"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Height          =   435
      Left            =   5970
      Picture         =   "comprovaretrebo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "No hi ha Vistiplau "
      Top             =   5340
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Height          =   435
      Left            =   4920
      Picture         =   "comprovaretrebo.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Donar Vistiplau"
      Top             =   5340
      Width           =   975
   End
   Begin VB.ListBox linia 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      Left            =   45
      TabIndex        =   0
      Top             =   135
      Width           =   6375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   6450
      X2              =   6450
      Y1              =   0
      Y2              =   5355
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Si prems el verd donaràs el VISTIPLAU  el vermell el treu."
      Height          =   375
      Left            =   780
      TabIndex        =   3
      Top             =   5475
      Width           =   4530
   End
End
Attribute VB_Name = "comprovaretrebo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub Command1_Click()
  comprovaretrebo.Tag = "OK"
  comprovaretrebo.Hide
End Sub

Private Sub Command2_Click()
comprovaretrebo.Tag = "NO"
comprovaretrebo.Hide
End Sub

Private Sub Form_Load()
Set dbtmpb = OpenDatabase(llegir_ini("General", "camibaixes", fitxerini))
imprimir_bobina numcomanda
comprovaretrebo.Tag = ""
End Sub
Function micresmaterial(descripcio As String, espesor As Double, tubolam As String) As Double
  r = espesor
  If descripcio = "GALGUES" Then
            If tubolam = "T" Then
                 r = Format(espesor / 4, "#,##0")
                  Else: r = Format(espesor / 2, "#,##0")
            End If
  End If
  If InStr(1, descripcio, "GR/") > 0 Then
    micresmaterial = espesor * -1
  End If
  micresmaterial = r
End Function

Private Sub List1_Click()

End Sub
Function desc_mat(numlot As String, ordre As Byte)
  Dim rsttmp3 As Recordset
  Dim rsttmp2 As Recordset
  Dim esp As Double
  If numlot = 0 Then Exit Function
  Set rsttmp3 = dbtmp.OpenRecordset("select materialex,colorex,espessor,mesuraesp,tubolam from comandes where comanda=" + atrim(numlot))
  
  If Not rsttmp3.EOF Then
     Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rsttmp3!mesuraesp)))
     If Not rsttmp2.EOF Then esp = micresmaterial(rsttmp2!descripcio, rsttmp3!espessor, rsttmp3!tubolam)
     Set rsttmp2 = dbtmp.OpenRecordset("select familia from materials where codi=" + atrim(cadbl(rsttmp3!materialex)))
    If Not rsttmp2.EOF Then
       Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(cadbl(rsttmp2!familia)))
       If Not rsttmp2.EOF Then desc_mat = atrim(rsttmp2!descripcio)
    End If
  End If
  If desc_mat <> "" Then desc_mat = desc_mat + "(" + atrim(esp) + ")"
  If ordre > 1 And desc_mat <> "" Then desc_mat = " + " + desc_mat
End Function
Sub possar_valors_taula_reb(numcom As String, idbobina As Double, situacioet As String)
   Dim rstbob As Recordset
   Dim rstcom As Recordset
   Dim rstenvio As Recordset
   Dim idio As String
   
   Dim rst2 As Recordset
   Dim ruta As String
   
   taula_tmp = "tmp_reb_empalmes_p"
   Set rstcom = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(numcom)))
   Set rstenvio = dbtmp.OpenRecordset("select * from clients_envios where id=" + atrim(cadbl(rstcom!direnvio)))
   Set rst2 = dbtmp.OpenRecordset("select * from productes where codi='" + atrim(rstcom!producte) + "'")
   If Not rstcom.EOF And Not rstenvio.EOF And Not rst2.EOF Then
      Set rstopcionset = dbtmp.OpenRecordset("select * from clients_etbobina where id_envio=" + atrim(rstenvio!ID))
      If rstopcionset.EOF Then rstopcionset.AddNew: rstopcionset!id_envio = rstenvio!ID: rstopcionset.Update: rstopcionset.MoveFirst
      rsttmp.AddNew
      rsttmp!idiomaclient = atrim(rstenvio!Idioma)
      If rsttmp!idiomaclient = "" Then rsttmp!idiomaclient = "ES"
      'rsttmp!idiomaclient = "EN"
      rsttmp!etmostra = rstopcionset!etmostra
      rsttmp!comandacli = atrim(rstcom!comandaclient)
      rsttmp!pesbobina = cadbl(9999)
      rsttmp!refclient = atrim(rstcom!refclient)
      rsttmp!numcomanda = atrim(rstcom!comanda)
      rsttmp!texteimpresio = IIf(InStr(1, rst2!ruta, "I") > 0, atrim(rstcom!texteimpressio), "")
      rsttmp!codiproducte = ""
      rsttmp!dataproduccio = "01/01/01"
      rsttmp!material = desc_mat(rstcom!comanda, 1) + desc_mat(cadbl(rstcom!linkcomanda1), 2) + desc_mat(cadbl(rstcom!linkcomanda2), 3)
      rsttmp!midarebobinat = "999"
      rsttmp!desarroll = rstcom!dessarroll
      rsttmp!peces = 9999
      If (atrim(rstcom!continu) <> "S" And rstcom!dessarroll > 0) Then rsttmp!peces = Fix(99999 / cadbl(rstcom!dessarroll))
      rsttmp!numbob = 9
      rsttmp!metresbob = 999
      rsttmp!codibarres = rstcom!codibarras
      rsttmp!obsetiqueta = IIf(atrim(rstopcionset!obsetiq) <> "", atrim(rstopcionset!obsetiq), atrim(rstcom!obsetiq))
      rsttmp!situacioet = situacioet
      If atrim(rstopcionset!campcodibarres) <> "" Then
        rsttmp!campcodibarres = rstcom.Fields(rstopcionset!campcodibarres) ' s'ha de agafar el que possi a client
        rsttmp!tipuscodibarres = rstopcionset!tipuscodibarres ' s'ha de agafar el qu epossi a client
      End If
      rsttmp!inplacsasino = IIf(cadbl(rstenvio!emb_anonim) = 0, "INPLACSA", "")
      rsttmp!nomclient = atrim(rstenvio!nome)
      rsttmp!operari = 99
      idio = IIf(rsttmp!idiomaclient <> "ES", "EN", rsttmp!idiomaclient)
      rsttmp!descproducte = atrim(rst2.Fields("descpelclient_" + idio))
      rsttmp.Update
        Else: MsgBox "Hi ha hagut un error de client d'envio o de comanda. NO ES POT IMPRIMIR LA ETIQUETA": Exit Sub
   End If
   rsttmp.MoveFirst
End Sub
Sub crear_taula_rev_empalmes()
  Dim camps(100, 2) As String
  taula_tmp = "tmp_reb_empalmes_p"
  If Not existeixlataula(dbtmpb.Name, taula_tmp) Then
 ' On Error Resume Next
 '  dbtmpb.Execute "drop table " + taula_tmp
 ' On Error GoTo 0
        i = 1
        camps(i, 1) = "comandacli": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "pesbobina": camps(i, 2) = "double": i = i + 1
        camps(i, 1) = "refclient": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "numcomanda": camps(i, 2) = "double": i = i + 1
        camps(i, 1) = "texteimpresio": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "codiproducte": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "dataproduccio": camps(i, 2) = "date": i = i + 1
        camps(i, 1) = "material": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "midarebobinat": camps(i, 2) = "double": i = i + 1
        camps(i, 1) = "desarroll": camps(i, 2) = "double": i = i + 1
        camps(i, 1) = "peces": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "numbob": camps(i, 2) = "integer": i = i + 1
        camps(i, 1) = "metresbob": camps(i, 2) = "double": i = i + 1
        camps(i, 1) = "codibarres": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "obsetiqueta": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "situacioet": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "inplacsasino": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "nomclient": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "descproducte": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "operari": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "campcodibarres": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "tipuscodibarres": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "etmostra": camps(i, 2) = "bit": i = i + 1
        camps(i, 1) = "idiomaclient": camps(i, 2) = "string": i = i + 1
        dbtmpb.Execute ("create table " + taula_tmp + " (n string)")
        For i = 1 To 100
          If camps(i, 1) <> "" Then
             dbtmpb.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
              Else: i = 1000
          End If
        Next i
          Else: dbtmpb.Execute "delete * from " + taula_tmp
   End If
  'ample double,plegat double,solapa double,espessor double,metres double,kilos double)"
  'dbtmpb.Execute ("create table tmp_lam_empalmes (" + camps + camps2 + camps3 + camps4) + ")"
End Sub
Sub imprimir_bobina(numcomanda As String)
 taula_tmp = "tmp_reb_empalmes_p"
 Set rsttmp = Nothing: r = ""
crear_taula_rev_empalmes
Set rsttmp = dbtmpb.OpenRecordset(taula_tmp)
possar_valors_taula_reb numcomanda, 0, ""
If rsttmp.EOF Then Exit Sub
preparar_etiqueta_zebra

'imprimir_etiqueta_zebra
Set rsttmp = Nothing
End Sub

Sub preparar_etiqueta_zebra()
   Dim v As String
   linia = ""
   
   With rsttmp
   idiomaclient = !idiomaclient
   possar_codidebarres
   linia.AddItem sitoca(!inplacsasino, "inplacsasino") + " " + sitoca(retallar(!nomclient, 22), "nomclient")
   linia.AddItem sitoca(retallar(Idioma("Producto: ") + atrim(!descproducte), 40), "descproducte")
   linia.AddItem sitoca(Idioma("RefC:") + !refclient, "refclient") + " " + sitoca(Idioma("PedC:") + !comandacli, "comandacli")
   linia.AddItem sitoca(retallar(!material, 40), "material")
   linia.AddItem sitoca(retallar(!texteimpresio, 40), "texteimpresio")
   linia.AddItem sitoca(Idioma("Ancho:") + atrim(!midarebobinat) + " m/m ", "midarebobinat") + sitoca(Idioma("Desar:") + atrim(!desarroll) + " m/m", "desarroll")
   linia.AddItem sitoca(retallar(IIf(!obsetiqueta <> "", Idioma("Obs.Et:") + !obsetiqueta, ""), 40), "obsetiqueta")
   linia.AddItem sitoca(Idioma("NºBob:") + atrim(!numbob), "numbob")
   linia.AddItem sitoca(Idioma("Peso:") + atrim(Format(!pesbobina, "#,##0.0")) + " Kg", "pesbobina")
   linia.AddItem sitoca(Idioma("Long:") + atrim(Format(!metresbob, "#,##0")) + " Mts", "metresbob")
   linia.AddItem sitoca(IIf(!peces > 0, Idioma("Unidades:") + atrim(Format(!peces, "#,##0")), ""), "peces")
   linia.AddItem sitoca(Format(!dataproduccio, "dd/mm/yy"), "dataproduccio") + "    " + sitoca(Idioma("Op:") + !operari, "operari") + " " + sitoca(Idioma("Lote: ") + Format(!numcomanda, "#,##0"), "numcomanda") + "  " + !situacioet
   linia.AddItem " "
   r = atrim(rstopcionset!etinteriorbob)
   If r <> "" Then linia.AddItem "ET.INT.BOBINA:(" + r + ")"
   If rstopcionset!etmostra Then linia.AddItem "TREURE ETIQUETA MOSTRA PEL CLIENT"
   End With
End Sub
Sub possar_codidebarres()
 
  If atrim(rsttmp!tipuscodibarres) <> "" Then linia.AddItem "CODI DE BARRES: (" + atrim(rsttmp!tipuscodibarres) + ")  " + atrim(rsttmp!campcodibarres)
End Sub
Function Idioma(txt As String) As String
 Dim v As String
 Dim fitxeridioma As String
 
 If idiomaclient = "" Then idiomaclient = "ES"
 fitxeridioma = llegir_ini("General", "rutallistats", "comandes.ini") + idiomaclient + "_etiquetareb.txt"
 f = llegir_ini("Idioma", txt, fitxeridioma)
 'If f = "{[}]" Then escriure_ini "Idioma", txt, txt, fitxeridioma: f = txt
 Idioma = f
End Function
Function sitoca(txt As String, camp As String) As String
  sitoca = ""
  If Not rstopcionset.Fields(camp) Then sitoca = txt
  If atrim(rsttmp.Fields(camp)) = "" Then sitoca = ""
End Function
Function retallar(txt As String, tamany As Integer) As String
   retallar = Mid(txt, 1, tamany)
End Function
'Sub substituir(buscar As String, canviar As String)
'   comença = InStr(1, linia, buscar) - 1
'   If comença < 1 Then Exit Sub
'   acaba = comença + Len(buscar) + 1
'   linia = Mid(linia, 1, comença) + canviar + Mid(linia, acaba)
'End Sub

Sub imprimir_etiqueta_zebra()
  Dim nomord As String * 255
  GetComputerName nomord, 255
  Open "c:\temp\etiquetareb.prn" For Output As #2
  Print #2, linia.Text
  Close #2
  'linia = ""
  nomord = Mid(nomord, 1, InStr(1, nomord, Chr$(0)) - 1)
  
  Shell "c:\windows\system32\cmd.exe /c type c:\temp\etiquetareb.prn>\\" + atrim(nomord) + "\zebra"
  Shell "c:\windows\system32\cmd.exe /c type " + llegir_ini("General", "rutallistats", "comandes.ini") + "graficetareb1.prn>\\" + atrim(nomord) + "\zebra"
End Sub
