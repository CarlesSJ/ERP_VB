VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formbobinesaimpresores 
   Caption         =   "Bobines portades a Impresora"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   6060
   ClipControls    =   0   'False
   Icon            =   "formbobinesaimpresores.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5550
      Top             =   30
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0017D062&
      Caption         =   "Ordre Imp."
      Height          =   765
      Left            =   45
      Picture         =   "formbobinesaimpresores.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   210
      Width           =   1155
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C4DA45&
      Caption         =   "per retornar"
      Height          =   765
      Left            =   4830
      Picture         =   "formbobinesaimpresores.frx":065C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   210
      Width           =   1155
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00ED823A&
      Caption         =   "Diam.Picus"
      Height          =   765
      Left            =   3630
      Picture         =   "formbobinesaimpresores.frx":0BE6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   210
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00EAD9CE&
      Caption         =   "a màquina"
      Height          =   765
      Left            =   2445
      Picture         =   "formbobinesaimpresores.frx":0DCB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   210
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "Penjades"
      Height          =   765
      Left            =   1245
      Picture         =   "formbobinesaimpresores.frx":1406
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   210
      Width           =   1155
   End
   Begin VB.Frame framemaquina 
      BackColor       =   &H00FDDECE&
      Height          =   6285
      Left            =   45
      TabIndex        =   0
      Top             =   1020
      Width           =   5910
      Begin VB.TextBox ccodidebarres 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1035
         MaxLength       =   10
         TabIndex        =   2
         Top             =   165
         Width           =   2790
      End
      Begin MSFlexGridLib.MSFlexGrid reixa 
         Height          =   5505
         Left            =   315
         TabIndex        =   3
         Top             =   720
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   9710
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   435
         Left            =   3915
         Picture         =   "formbobinesaimpresores.frx":1956
         Stretch         =   -1  'True
         Top             =   180
         Width           =   450
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Bobina:"
         Height          =   240
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   1050
      End
   End
   Begin VB.Frame Framediametrepicus 
      BackColor       =   &H00ED823A&
      Height          =   6255
      Left            =   30
      TabIndex        =   14
      Top             =   1050
      Width           =   5925
      Begin VB.TextBox ccodidebarrespicus 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1035
         MaxLength       =   10
         TabIndex        =   15
         Top             =   165
         Width           =   2790
      End
      Begin MSFlexGridLib.MSFlexGrid reixapicus 
         Height          =   5370
         Left            =   525
         TabIndex        =   16
         Top             =   825
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   9472
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label etmanualment 
         BackStyle       =   0  'Transparent
         Caption         =   "Manualment amb CTRL apretat."
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   2460
         TabIndex        =   18
         Top             =   630
         Width           =   2940
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Bobina:"
         Height          =   240
         Left            =   135
         TabIndex        =   17
         Top             =   225
         Width           =   1050
      End
      Begin VB.Image Image5 
         Height          =   435
         Left            =   3915
         Picture         =   "formbobinesaimpresores.frx":2359
         Stretch         =   -1  'True
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.Frame framepenjades 
      BackColor       =   &H00808000&
      Caption         =   "Bobines penjades"
      Height          =   6210
      Left            =   90
      TabIndex        =   6
      Top             =   1080
      Width           =   5850
      Begin VB.CommandButton treurebobina2 
         Height          =   480
         Left            =   4170
         Picture         =   "formbobinesaimpresores.frx":2D5C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Treure aquesta bobina."
         Top             =   2685
         Width           =   645
      End
      Begin VB.CommandButton treurebobina1 
         Height          =   480
         Left            =   4155
         Picture         =   "formbobinesaimpresores.frx":32E6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Treure aquesta bobina."
         Top             =   1725
         Width           =   645
      End
      Begin VB.TextBox ccodidebarres2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1095
         MaxLength       =   10
         TabIndex        =   9
         Top             =   540
         Width           =   2790
      End
      Begin VB.TextBox cbob1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1680
         Width           =   2640
      End
      Begin VB.TextBox cbob2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2640
         Width           =   2625
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Desb/Bob:"
         Height          =   240
         Left            =   195
         TabIndex        =   10
         Top             =   600
         Width           =   1050
      End
      Begin VB.Image Image2 
         Height          =   435
         Left            =   3975
         Picture         =   "formbobinesaimpresores.frx":3870
         Stretch         =   -1  'True
         Top             =   555
         Width           =   450
      End
      Begin VB.Image Image3 
         Height          =   855
         Left            =   390
         Picture         =   "formbobinesaimpresores.frx":4273
         Stretch         =   -1  'True
         Top             =   1350
         Width           =   915
      End
      Begin VB.Image Image4 
         Height          =   810
         Left            =   285
         Picture         =   "formbobinesaimpresores.frx":47C7
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   1065
      End
   End
   Begin VB.Frame Frameretornar 
      BackColor       =   &H00C4DA45&
      Height          =   6255
      Left            =   15
      TabIndex        =   20
      Top             =   1065
      Width           =   5940
      Begin VB.TextBox ccodibarresretorn 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1035
         MaxLength       =   10
         TabIndex        =   21
         Top             =   165
         Width           =   2790
      End
      Begin MSFlexGridLib.MSFlexGrid reixaretorn 
         Height          =   5385
         Left            =   360
         TabIndex        =   22
         Top             =   720
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   9499
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Bobina:"
         Height          =   240
         Left            =   135
         TabIndex        =   23
         Top             =   225
         Width           =   1050
      End
      Begin VB.Image Image6 
         Height          =   435
         Left            =   3915
         Picture         =   "formbobinesaimpresores.frx":4D17
         Stretch         =   -1  'True
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.Label etoperari 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   210
      TabIndex        =   24
      Top             =   15
      Width           =   5010
   End
   Begin VB.Menu mutils 
      Caption         =   "Utilitats"
      Begin VB.Menu mbuscarbob 
         Caption         =   "Buscar bobina a magatzem"
      End
      Begin VB.Menu mhistoricbobinesamaquines 
         Caption         =   "Historic de bobines portades a màquina"
      End
   End
End
Attribute VB_Name = "formbobinesaimpresores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VScr As Integer, HScr As Integer
  Dim VFactor As Single, HFactor As Single
Private Sub ccodibarresretorn_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vcodidb As String
   If KeyCode = 13 Then
       vcodidb = substituir(ccodibarresretorn, "-", "/")
       vcodidb = substituir(vcodidb, "'", "/")
       treurelabobinadeimpresores vcodidb
       KeyCode = 0
   End If
End Sub
Sub treurelabobinadeimpresores(vbobina As String)
   Dim rst As Recordset
   If MsgBox("Vols fer el retorn d'aquesta bobina?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
        Set rst = dbtmpb.OpenRecordset("select * from impresores_bobinesamaquina where numbobina='" + atrim(vbobina) + "'")
        If Not rst.EOF Then
               rst.Delete
             Else: MsgBox "No he trobat aquesta bobina a la llista de bobines a impresores.", vbCritical, "Error"
        End If
   End If
   Set rst = Nothing
   carregar_bobines_retorn
   ccodibarresretorn = ""
   ccodibarresretorn.SetFocus
End Sub

Private Sub ccodidebarres2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
       ccodidebarres2 = substituir(ccodidebarres2, "·", "#")
       If ccodidebarres2 = "1#" Or ccodidebarres2 = "2#" Then
           ccodidebarres2.SetFocus
           ccodidebarres2.SelStart = Len(ccodidebarres2)
           KeyCode = 0
           sonar_sirena "unpitu"
           GoTo fi
       End If
       ccodidebarres2 = substituir(ccodidebarres2, "-", "/")
       ccodidebarres2 = substituir(ccodidebarres2, "'", "/")
       If Mid(ccodidebarres2 + "   ", 2, 1) = "#" Then
          afegir_bobinaamaquina ccodidebarres2
       End If
       ccodidebarres2 = ""
       If ccodidebarres2.visible Then ccodidebarres2.SetFocus
       KeyCode = 0
   End If
fi:
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
Function llegir_comanda_compartida_impresores() As String
  Dim vdata As String
  Dim vcomanda As String
  vdata = llegir_ini("Impresores_Compartida", "dataihora_maq_" + atrim(nummaq), rutadelfitxer(cami) + "valorsprograma.ini")
  vcomanda = llegir_ini("Impresores_Compartida", "comanda_maq_" + atrim(nummaq), rutadelfitxer(cami) + "valorsprograma.ini")
  If Not existeix("c:\ordprog.ini") Then If DateDiff("n", vdata, Now) > 2 Then MsgBox "Error amb connexió compartida de la tablet amb el programa principal.", vbCritical, "Error": GoTo fi
  llegir_comanda_compartida_impresores = vcomanda
  etoperari = llegir_ini("Impresores_Compartida", "nomoperari2_maq_" + atrim(nummaq), rutadelfitxer(cami) + "valorsprograma.ini")
  If UCase(arguments(1)) = "DESBOBINADORS" Then
        numop = cadbl(llegir_ini("Impresores_Compartida", "numop_maq_" + atrim(nummaq), rutadelfitxer(cami) + "valorsprograma.ini"))
        numop2 = cadbl(llegir_ini("Impresores_Compartida", "numop2_maq_" + atrim(nummaq), rutadelfitxer(cami) + "valorsprograma.ini"))
  End If
fi:
End Function
Sub afegir_bobinaamaquina(vbobina As String)
  Dim vdesb As Double
  Dim verror As String
  Dim vesvalida As Boolean
  Dim vpalet As Double
  Dim vbob As Double
  Dim vnumc As Double
  Dim vgrup As Double
  Dim vtexte As String
  vdesb = cadbl(Mid(vbobina + "   ", 1, 1))
  vbobina = atrim(Mid(vbobina + "   ", 3))
  vnumc = cadbl(llegir_comanda_compartida_impresores)
  If vdesb = 0 Then Exit Sub
  vesvalida = comprovarsilabobinaesvalida(vbobina, verror, cadbl(vnumc))
  valorsdajust vnumc, vgrup, vtexte, 0
  If Not vesvalida Then
     sonar_sirena "continuu"
     missatgeerrorbobina verror
     ccodidebarres2 = ""
       'encara que el material no sigui vàlid si està dins del packinglist o grup s´acceptarà
     
     If estadinspackinglistogrup(vbobina, atrim(vnumc), atrim(vgrup)) Then

       MsgBox "El material no es exactament el mateix però com que està dins del packinglist es permet utilitzar-lo." + Chr(13) + " SISPLAU ASSEGURE-VOS QUE ES CORRECTE.", vbCritical, "A T E N C I Ó"
       vesvalida = True
     End If
  End If
  
  If vesvalida Then
    cbob1 = llegir_ini("Bobines_Desbobinadors_" + atrim(nummaq), "Bobina1", rutadelfitxer(cami) + "valorsprograma.ini")
    escriure_ini "Bobines_Desbobinadors_" + atrim(nummaq), "Bobina" + atrim(vdesb), vbobina, rutadelfitxer(cami) + "valorsprograma.ini"
    escriure_ini "Bobines_Desbobinadors_" + atrim(nummaq), "Horabob" + atrim(vdesb), Now, rutadelfitxer(cami) + "valorsprograma.ini"
    convertirScanambPaletiBobina vbobina, vpalet, vbob
    dbstocks.Execute "update bobines set sit='IMP' WHERE idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbob)
    carregar_bobines_desbobinadors
    sonar_sirena "intermitent"
  End If
End Sub
 Sub missatgeerrorbobina(verror As String)
      Load fcalculant
      fcalculant.Command1.BackColor = QBColor(12)
      fcalculant.Command1.caption = verror + Chr(13) + "Fes Click"
      fcalculant.Show 1
      Unload fcalculant
      wait 1
 End Sub
 

Private Sub ccodidebarres_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim vcodidb As String
   If KeyCode = 13 Then
       vcodidb = substituir(ccodidebarres, "-", "/")
       vcodidb = substituir(vcodidb, "'", "/")
       afegir_bobinaalareixa vcodidb
       KeyCode = 0
   End If
End Sub
Sub afegir_bobinaalareixa(vbobina As String)
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim vp As String
   Dim vb As String
   separarpaletibobina atrim(vbobina), vp, vb
   vbobina = substituir(vbobina, "'", "/")
   Set rst2 = dbtmpb.OpenRecordset("select * from impresores_bobinesamaquina where numbobina='" + vbobina + "'")
   If Not rst2.EOF Then GoTo fi
   Set rst = dbstocks.OpenRecordset("select * from bobines where trim(idpalet)&'/'&trim(idbobina)='" + vbobina + "'")
   If Not rst.EOF Then
    reixa.AddItem vbobina, 1
    dbtmpb.Execute "insert into impresores_bobinesamaquina (numbobina,maquina,operari) values ('" + vbobina + "'," + atrim(nummaq) + "," + atrim(numop2) + ")"
    dbstocks.Execute "update bobines set sit='IMP' where idpalet=" + Trim(vp) + " and idbobina=" + atrim(vb)
    sonar_sirena "intermitent"
     Else:
       sonar_sirena "unpitu"
       MsgBox "No he localitzat aquesta bobina a al base de dades.", vbCritical, "Error"
       
   End If
fi:
   ccodidebarres.text = ""
   ccodidebarres.SetFocus
   Set rst = Nothing
   Set rst2 = Nothing
End Sub
Sub carregar_bobines_picus()
  Dim rst As Recordset
  Dim vpalet As Double
  Dim vbob As Double
  Dim vbobina As String
  Dim vmtrs As Double
  Dim rstb As Recordset
  reixapicus.Clear
  reixapicus.Rows = 1
  reixapicus.TextMatrix(0, 0) = "NºBobina"
  'primer trec totes les bobines que ja no cal que estiguin dins la taula
  'Set rst = dbtmpb.OpenRecordset("Select * from impresores_bobinesamaquina where diametrerevisat=false and  data>#" + Format(DateAdd("m", -3, Now), "mm/dd/yy") + "# order by data desc")
  Set rst = dbtmpb.OpenRecordset("select * from bobines_pendent_revisar_diametre where seccio='I' and maquina=" + atrim(nummaq))
  While Not rst.EOF
    vbobina = atrim(rst!palet) + "/" + atrim(rst!bobina)
    vpalet = cadbl(rst!palet) 'cadbl(Mid(" " + vbobina, 1, InStr(1, vbobina + "  ", "/")))
    vbob = cadbl(rst!bobina)
    vmtrs = bobinesdentrada.calcular_mtrsdispreals(vpalet, vbob)
    If vmtrs > 0 Then
          reixapicus.AddItem vbobina
            Else: rst.Delete
    End If
    'Set rstb = dbtmpb.OpenRecordset("select sit,mts from bobines where idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbob))
    'If Not rstb.EOF Then
       'vmtrs = bobinesdentrada.calcular_mtrsdispreals(vpalet, vbob)
       'If vmtrs <> rstb!mts And rstb!sit = "IMP" Then reixapicus.AddItem rst!numbobina
    'End If
    rst.MoveNext
  Wend
  
  Set rst = Nothing
  Set rstb = Nothing
  

End Sub
Sub carregar_bobines_retorn()
  Dim rst As Recordset
  Dim vpalet As String
  Dim vbob As String
  Dim vbobina As String
  Dim vmtrs As Double
  Dim rstp As Recordset
  Dim rstpo As Recordset
  Dim vsqlparcialsordreimpresio As String
  
  vsqlparcialsordreimpresio = "SELECT Parcials.idpalet, Parcials.idbobina, impresores_ordreimpresio.comanda FROM Parcials LEFT JOIN impresores_ordreimpresio ON cdbl(Parcials.comanda) = impresores_ordreimpresio.comanda Where impresores_ordreimpresio.comanda <> Null And parcials.utilitzada = False "
  Set rstpo = dbtmpb.OpenRecordset(vsqlparcialsordreimpresio)
  'passarbobines_acabadesambIMP_fora
  'ensenyo totes les valides
  Set rst = dbtmpb.OpenRecordset("Select * from impresores_bobinesamaquina where data>#" + Format(DateAdd("m", -3, Now), "mm/dd/yy") + "# and maquina=" + atrim(nummaq) + " order by data desc")
  reixaretorn.TextMatrix(0, 0) = "Bobina"
  reixaretorn.ColWidth(0) = 3500
  reixaretorn.ColAlignment(0) = 3
  reixaretorn.Rows = 1
  While Not rst.EOF
    vbobina = atrim(rst!numbobina)
    separarpaletibobina atrim(vbobina), vpalet, vbob
    rstpo.FindFirst "idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbob)
    If rstpo.NoMatch Then
        ' If cadbl(vpalet) = 54183 Then Stop
         Set rstp = dbtmpb.OpenRecordset("select * from parcials where idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbob) + " and ((cdbl(comanda)>10000 and not utilitzada) or (cdbl(comanda)>2000 and cdbl(comanda)<10000))")
        ' Clipboard.Clear
        ' Clipboard.SetText "select * from parcials where idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbob) + " and ((cdbl(comanda)>10000 and not utilitzada) or (cdbl(comanda)>2000 and cdbl(comanda)<10000))"
         If rstp.EOF Then reixaretorn.AddItem vpalet + "/" + Format(vbob, "00")
    End If
    rst.MoveNext
  Wend
  Set rst = Nothing
  Set rstpo = Nothing
  Set rstp = Nothing
  
End Sub

Sub carregar_bobines_baixades()
  Dim rst As Recordset
  Dim vpalet As Double
  Dim vbob As Double
  Dim vbobina As String
  Dim vmtrs As Double
  Dim rstb As Recordset
  
  'primer trec totes les bobines que ja no cal que estiguin dins la taula
  'Set rst = dbtmpb.OpenRecordset("Select * from impresores_bobinesamaquina where data>#" + Format(DateAdd("m", -3, Now), "mm/dd/yy") + "# order by data desc")
  'While Not rst.EOF
  '  vbobina = rst!numbobina
  '  vpalet = cadbl(Mid(" " + vbobina, 1, InStr(1, vbobina + "  ", "/")))
  '  vbob = cadbl(Mid(vbobina, InStr(1, vbobina + "  ", "/") + 1))
    'Set rstb = dbtmpb.OpenRecordset("select sit from bobines where idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbob))
    'If Not rstb.EOF Then
   '    vmtrs = rstb!bobinesdentrada.calcular_mtrsdispreals(vpalet, vbob)
   '    If vmtrs <= 500 Then
   '         rst.Delete
   '        Else: If rstb.sit <> "IMP" Then rst.Delete
   '    End If
    'End If
   ' rst.MoveNext
  'Wend
  
  'ensenyo totes les valides
  Set rst = dbtmpb.OpenRecordset("Select * from impresores_bobinesamaquina where data>#" + Format(DateAdd("m", -3, Now), "mm/dd/yy") + "# order by data desc")
  reixa.Rows = 1
  While Not rst.EOF
    reixa.AddItem rst!numbobina
    rst.MoveNext
  Wend
  Set rst = Nothing
  Set rstb = Nothing
  
End Sub

Private Sub ccodidebarrespicus_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim vcodidb As String
 Dim vpalet As String
 Dim vbob As String
 Dim vcanvifet As Boolean
 
 If KeyCode = 13 Then
       vcodidb = substituir(ccodidebarrespicus, "-", "/")
       vcodidb = substituir(vcodidb, "'", "/")
       separarpaletibobina atrim(vcodidb), vpalet, vbob
       If comprovar_si_shautilitzat(cadbl(vpalet), cadbl(vbob)) = False Then
          If MsgBox("Aquesta bobina s'ha afegit a una comanda però encara no s'ha informat dels metres gastats, potser encara no s'ha tancat la comanda." + vbNewLine + "Vols ajustar per DIAMETRE igualment?", vbCritical + vbDefaultButton2 + vbYesNo, "Error") = vbNo Then Exit Sub
       End If
       ajustar_diametre_real vcodidb, vcanvifet
       If vcanvifet Then dbtmpb.Execute "delete * from bobines_pendent_revisar_diametre where seccio='I' and maquina=" + atrim(nummaq) + " and palet=" + atrim(vpalet) + " and bobina=" + atrim(vbob)
       carregar_bobines_picus
       KeyCode = 0
 End If
End Sub
Function comprovar_si_shautilitzat(vpalet As Double, vbob As Double) As Boolean
   Dim rst As Recordset
   Dim rstp As Recordset
   Dim vsql As String
   
   vsql = "SELECT impressores.comanda FROM impressores LEFT JOIN (bobinesentimp RIGHT JOIN bobinesimp ON bobinesentimp.id = bobinesimp.Id) ON impressores.Id = bobinesimp.controlid WHERE (((bobinesentimp.palet)=" + atrim(vpalet) + ") AND ((bobinesentimp.bobina)=" + atrim(vbob) + "))"
   Set rst = dbtmpb.OpenRecordset(vsql)
   Set rstp = dbtmpb.OpenRecordset("select * from parcials where idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbob))
   If Not rst.EOF Then comprovar_si_shautilitzat = True
   While Not rst.EOF
     rstp.FindFirst "comanda='" + atrim(rst!comanda) + "'"
     If rstp.NoMatch Then
        comprovar_si_shautilitzat = False
         Else: If rstp!utilitzada = False Then comprovar_si_shautilitzat = False
     End If
     rst.MoveNext
   Wend
   
   Set rst = Nothing
End Function
Sub actualitzar_metresxrdiametre(vpalet As String, vbob As String, vmetresanteriors As Double, vmetresnous As Double)
    Dim rst As Recordset
    Dim vmetresbob As Double
    Dim vvalues As String
    Dim vmetresactualitzar As Double
    'Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
    'vmetresanteriors = bobinesdentrada.calcular_mtrsdispreals(cadbl(vpalet), cadbl(vbob))
    vmetresactualitzar = Redondejar(vmetresanteriors - vmetresnous, 0)
    Set rst = dbstocks.OpenRecordset("select * from parcials where comanda='444' and idpalet=" + atrim(cadbl(vpalet)) + " and idbobina=" + atrim(cadbl(vbob)))
    vvalues = "(" + atrim(vpalet) + "," + atrim(vbob) + ",True,'444',now," + atrim(cadbl(numop2)) + ",'I','Actualització metres per diametre.')"
    If rst.EOF Then dbstocks.Execute "insert into parcials (idpalet,idbobina,utilitzada,comanda,data,operari,seccio,observacions) values " + vvalues
    Set rst = dbstocks.OpenRecordset("select * from parcials where comanda='444' and idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbob))
    If Not rst.EOF Then
       rst.Edit: rst!metres = rst!metres + vmetresactualitzar: rst!Data = Now: rst.Update
       bobinesdentrada.actualitzar_metres_disponibles cadbl(vpalet), cadbl(vbob)
    End If
    Set rst = Nothing
End Sub
Sub separarpaletibobina(vnumbob As String, vpalet As String, vbob As String)
    If vnumbob = "" Then Exit Sub
    If InStr(1, vnumbob, "/") = 0 Then Exit Sub
    vpalet = cadbl(Mid(vnumbob, 1, InStr(1, vnumbob, "/") - 1))
    vbob = cadbl(substituir(vnumbob, vpalet + "/", ""))
End Sub

Sub ajustar_diametre_real(vbobina As String, vcanvifet As Boolean)
   Dim vpalet As String
   Dim vbob As String
   Dim vdiametrenou As String
   Dim vmetresanteriors As Double
   Dim vmetres As Double
   Dim rstbob As Recordset
   Dim v As String
   'Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
   separarpaletibobina atrim(vbobina), vpalet, vbob
   If cadbl(vpalet) = 0 Or cadbl(vbob) = 0 Then Exit Sub
   Set rstbob = dbtmpb.OpenRecordset("select * from bobines_pendent_revisar_diametre where bobina=" + atrim(vbob) + " and palet=" + atrim(vpalet) + " and seccio='I' and maquina=" + atrim(nummaq))
   If rstbob.EOF Then If MsgBox("No he trobat la bobina " + vbobina + vbNewLine + "VOLS UTILITZAR-LA IGUALMENT?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   Set rstbob = dbstocks.OpenRecordset("select * from bobines where idbobina=" + atrim(vbob) + " and idpalet=" + atrim(vpalet))
   If rstbob.EOF Then MsgBox "No he trobat la bobina " + vbobina, vbCritical, "Atenció": Exit Sub
   vmetresanteriors = bobinesdentrada.calcular_mtrsdispreals(cadbl(vpalet), cadbl(vbob))
demanardiametre:
   vdiametrenou = cadbl(InputBox("Entra el diametre actual de la bobina " + vpalet + "/" + vbob, "Nou diametre"))
   If vdiametrenou > 85 Then MsgBox "Aquesta mida no es possible, una bobina no pot fer mes de 80cm de diametre.", vbCritical, "Atenció": GoTo demanardiametre
   If cadbl(vdiametrenou) = 0 Then Exit Sub
   vmetres = Redondejar(calcular_metresambdiametre(cadbl(vpalet), cadbl(vbob), cadbl(vdiametrenou)), cadbl(rstbob!tamanycanutu))
   
   If cadbl(vmetres - vmetresanteriors) >= 1000 Or cadbl(vmetres - vmetresanteriors) <= -1000 Then
      If MsgBox("Aquest diametre que dius serien uns " + atrim(vmetres) + " metres desfase de " + atrim(cadbl(vmetres - vmetresanteriors)) + " metres." + vbNewLine + "Creus que es correcte?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then GoTo demanardiametre
'      enviaremailgeneric "missatgesgenericsimpresores", "444-Ajust de " + atrim(cadbl(vmetres - vmetresanteriors)) + " metres per diametre. " + atrim(Now), "El diametres suggerit es " + atrim(vdiametrenou) + " cm." + vbNewLine + atrim(vmetresanteriors) + "m. abans i " + atrim(vmetres) + "m despres del canvi." + vbNewLine + vbNewLine + atrim(Now) + " -- (" + nommaq + ") - " + nomoperari
      mantenimentbobina.passaravis atrim(vpalet), atrim(vbob), "444-Ajust de " + vbobina + " " + atrim(cadbl(vmetres - vmetresanteriors)) + " metres per diametre. " + atrim(Now), form1.comanda, "El diametres suggerit es " + atrim(vdiametrenou) + " cm." + vbNewLine + atrim(vmetresanteriors) + "m. abans i " + atrim(vmetres) + "m despres del canvi." + vbNewLine + vbNewLine + atrim(Now) + " -- (" + nommaq + ") - " + nomoperari
      v = "El diametres suggerit es " + atrim(vdiametrenou) + " cm." + vbNewLine + atrim(vmetresanteriors) + "m. abans i " + atrim(vmetres) + "m despres del canvi." + vbNewLine + vbNewLine + atrim(Now) + " -- (" + nommaq + ") - " + nomoperari
      enviaremailgeneric "tintes@inplacsa.com", "444-Ajust de " + vbobina + " " + atrim(cadbl(vmetres - vmetresanteriors)) + " metres per diametre. " + atrim(Now), v
   End If
   dbtmpb.Execute "delete * from parcials where comanda='444' and idbobina=" + atrim(vbob) + " and idpalet=" + atrim(vpalet)
   vmetres = Redondejar(calcular_metresambdiametre(cadbl(vpalet), cadbl(vbob), cadbl(vdiametrenou)), cadbl(rstbob!tamanycanutu))
   If vmetres <> 0 Then
     vmetresanteriors = bobinesdentrada.calcular_mtrsdispreals(cadbl(vpalet), cadbl(vbob))
     actualitzar_metresxrdiametre vpalet, vbob, vmetresanteriors, vmetres: wait 1
   End If
   If UCase(arguments(1)) = "DESBOBINADORS" Then
          escriure_ini "Impresores_Compartida", "imprimirBobina_maq_" + atrim(nummaq), atrim(vbob), rutadelfitxer(cami) + "valorsprograma.ini"
          escriure_ini "Impresores_Compartida", "imprimirPalet_maq_" + atrim(nummaq), atrim(vpalet), rutadelfitxer(cami) + "valorsprograma.ini"
        Else
         bobinesdentrada.imprimir_bobinaparcial cadbl(vpalet), cadbl(vbob), , 1
   End If
   ccodidebarrespicus = ""
   vcanvifet = True
   If ccodidebarrespicus.visible Then ccodidebarrespicus.SetFocus
End Sub
Function calcular_metresambdiametre(palet As Double, bobina As Double, vdiametre As Double, Optional canutu As Double) As Double
     Dim rstp As Recordset
  Dim rstb As Recordset
  Dim metres As Double
  Dim micres As Double
  Dim diametre As Double
  Dim pi As Double
  If cadbl(canutu) = 0 Then canutu = 15.2
  If canutu < 10 Then canutu = canutu + 2 'afegeixo l'amplada del cartrò del canutu
  If canutu >= 10 Then canutu = canutu + 2.8 'afegeixo l'amplada del cartrò del canutu
  '3,1416*(Diametro maximo^2-Diametro corazon^2)/(4*Espesor)
  Set rstp = dbstocks.OpenRecordset("select micres,grmsm2,codimatprognou from palets where idpalet=" + atrim(palet))
  If rstp.EOF Then GoTo fi
  Set rstb = dbstocks.OpenRecordset("select * from materials where codi=" + atrim(rstp!codimatprognou))
  If Not rstp.EOF Then
    pi = 4 * Atn(1)
    vdiametre = vdiametre / 100
    canutu = canutu / 100
    micres = cadbl(rstp!micres)
    If micres = 0 Then micres = cadbl(rstb!micresdelsgrm2)
    If micres = 0 Then GoTo fi
    micres = (micres * 0.0001) / 100
    diametre = (((vdiametre * vdiametre) - (canutu * canutu)) * pi) / (4 * micres)
    'diametre = Sqr(((metres * micres) / pi) + (canutu * canutu)) * 200
    calcular_metresambdiametre = Redondejar(diametre, 0)
    'If cadbl(calcular_metresambdiametre) < 9 Then calcular_metresambdiametre = "0"
  End If
fi:
  Set rstp = Nothing
  Set rstb = Nothing
End Function

Private Sub Command1_Click()
   framepenjades.Top = Command1.Top + Command1.Height
   framepenjades.Left = 50
   framemaquina.visible = False
   framepenjades.visible = True
   Framediametrepicus.visible = False
   Frameretornar.visible = False
   
   If ccodidebarres2.visible Then ccodidebarres2.SetFocus
  ' If LCase(arguments(1)) = "desbobinadors" Then 'redimensionarformulari Me
  '     formbobinesaimpresores.Top = 0
  '     formbobinesaimpresores.Left = -105
  ' End If
   
End Sub

Private Sub Command2_Click()
   framemaquina.Top = Command1.Top + Command1.Height
   framemaquina.Left = 50
   framemaquina.visible = True
   framepenjades.visible = False
   Framediametrepicus.visible = False
   Frameretornar.visible = False
   ratoli "espera"
   act_des_activarbotons False
   carregar_bobines_baixades
   act_des_activarbotons True
   ratoli "normal"
   If ccodidebarres.visible Then ccodidebarres.SetFocus
   'If LCase(arguments(1)) = "desbobinadors" Then 'redimensionarformulari Me
   '    formbobinesaimpresores.Top = 0
   '    formbobinesaimpresores.Left = -105
    'End If
End Sub
Sub act_des_activarbotons(vActivar As Boolean)
  Command1.Enabled = vActivar
  Command2.Enabled = vActivar
  Command3.Enabled = vActivar
  Command4.Enabled = vActivar
  Command5.Enabled = vActivar
End Sub

Private Sub Command3_Click()
   
   framemaquina.Top = Command1.Top + Command1.Height
   framemaquina.Left = 50
   Framediametrepicus.Top = framemaquina.Top
   Framediametrepicus.Left = framemaquina.Left
   framemaquina.visible = False
   framepenjades.visible = False
   Framediametrepicus.visible = True
   Frameretornar.visible = False
   ratoli "espera"
   act_des_activarbotons False
   carregar_bobines_picus
   act_des_activarbotons True
   ratoli "normal"
   If ccodidebarrespicus.visible Then ccodidebarrespicus.SetFocus
End Sub

Private Sub Command4_Click()
   Frameretornar.Top = Command1.Top + Command1.Height
   Frameretornar.Left = 50
   framemaquina.visible = False
   framepenjades.visible = False
   Framediametrepicus.visible = False
   Frameretornar.visible = True
   ratoli "espera"
   act_des_activarbotons False
   carregar_bobines_retorn
   act_des_activarbotons True
   ratoli "normal"
   If ccodibarresretorn.visible Then ccodibarresretorn.SetFocus
End Sub

Private Sub Command5_Click()
 act_des_activarbotons False
 carregar_llista_ordreimpressio
 act_des_activarbotons True
End Sub
Sub carregar_llista_ordreimpressio(Optional vnumc As String, Optional vnoobrirla As Boolean)
  'Dim vnumc As String
  Dim vcomandaactual As Double
  Dim vcomandafingerprint As Double
  Dim vnumcimp As Double
  If UCase(arguments(1)) <> "DESBOBINADORS" Then Exit Sub
  ratoli "espera"
  vcomandaactual = cadbl(comanda)
  vcomandafingerprint = vcomandaactual
  Load formordreimpresio
  'redimensionarformulari formordreimpresio
  formordreimpresio.reixa.row = 1
  formordreimpresio.reixa_refrescarfila
  formordreimpresio.reixa.width = formbobinesaimpresores.width - 1000
  formordreimpresio.reixa.Height = formbobinesaimpresores.Height - 800
  formordreimpresio.width = formbobinesaimpresores.width
  formordreimpresio.Height = formbobinesaimpresores.Height
  formordreimpresio.Command3.Left = formbobinesaimpresores.width - formordreimpresio.Command3.width - 100
  formordreimpresio.bimprimir.Left = formbobinesaimpresores.width - formordreimpresio.bimprimir.width - 100
  formordreimpresio.bbobinesamaquina.Left = formbobinesaimpresores.width - formordreimpresio.bbobinesamaquina.width - 100
  formordreimpresio.Frame6.Left = formbobinesaimpresores.width - formordreimpresio.Frame6.width - 100
  formordreimpresio.Frame6.Top = formordreimpresio.bimprimir.Top + formordreimpresio.bimprimir.Height + 200
  'formordreimpresio.Command1.Left = formbobinesaimpresores.width - formordreimpresio.Command1.width
  formordreimpresio.Show
  formordreimpresio.Top = formbobinesaimpresores.Top
  formordreimpresio.Left = formbobinesaimpresores.Left
  ratoli "normal"
  While Screen.ActiveForm.Name = "formordreimpresio" Or Screen.ActiveForm.Name = "obsidtreball" Or Screen.ActiveForm.Name = "veurereport"
     If seleccioret = 5 Then
         'form1.imprimir_packinglistTICKET cadbl(formordreimpresio.reixa.TextMatrix(formordreimpresio.reixa.row, 0)), False
         vnumcimp = cadbl(formordreimpresio.reixa.TextMatrix(formordreimpresio.reixa.row, 0))
         escriure_ini "Impresores_Compartida", "imprimirPKGLST_maq_" + atrim(nummaq), atrim(vnumcimp), rutadelfitxer(cami) + "valorsprograma.ini"
         seleccioret = 0
     End If
     DoEvents
  Wend
  
  Unload formordreimpresio
End Sub

Private Sub Form_Activate()
 llegir_comanda_compartida_impresores
 Command1_Click
 If LCase(arguments(1)) = "desbobinadors" Then 'redimensionarformulari Me
       formbobinesaimpresores.Top = 0
       formbobinesaimpresores.Left = -105
       Timer1.Enabled = True
 End If
End Sub
Sub redimensionarformulari(vform As Form)
  Static vfets As String
  If InStr(1, vfets, vform.Name) > 0 Then Exit Sub
  vfets = vfets + " " + vform.Name + " "
  HScr = Screen.width / Screen.TwipsPerPixelX
    VScr = Screen.Height / Screen.TwipsPerPixelY
    VFactor = VScr / 600
    HFactor = (800 * cadbl(VFactor)) / 600
    vform.width = vform.width * HFactor
    vform.Height = vform.Height * VFactor
  On Error Resume Next
  For Each ctl In vform
    ctl.Top = ctl.Top * VFactor
    ctl.Height = ctl.Height * VFactor
    ctl.Left = ctl.Left * HFactor
    ctl.width = ctl.width * HFactor
    ctl.FontSize = ctl.FontSize * HFactor
  Next
End Sub
Private Sub Form_Load()
  reixa.TextMatrix(0, 0) = "NºBobina"
  reixa.col = 0
  reixa.ColWidth(0) = 5000
  reixa.ColAlignment(0) = 3
  reixa.Rows = 1
  
  reixapicus.TextMatrix(0, 0) = "NºBobina"
  reixapicus.col = 0
  reixapicus.ColWidth(0) = 5000
  reixapicus.ColAlignment(0) = 3
  reixapicus.Rows = 1
  Framediametrepicus.Left = reixa.Left
  Framediametrepicus.Top = reixa.Top
  
  Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
  
'  carregar_bobines_baixades
  carregar_bobines_desbobinadors
 ' carregar_bobines_picus
  
End Sub
Sub carregar_bobines_desbobinadors()
   form1.actualitzarestatbobinesdesbobinadors
   cbob1 = llegir_ini("Bobines_Desbobinadors_" + atrim(nummaq), "Bobina1", rutadelfitxer(cami) + "valorsprograma.ini")
   cbob2 = llegir_ini("Bobines_Desbobinadors_" + atrim(nummaq), "Bobina2", rutadelfitxer(cami) + "valorsprograma.ini")
   If cbob1 = "{[}]" Then cbob1 = ""
   If cbob2 = "{[}]" Then cbob2 = ""
   
End Sub

Private Sub mbuscarbob_Click()
   If arguments(1) = "DESBOBINADORS" Then MsgBox "Aquesta opció no es pot fer desde la TABLET fer-ho desde l'ordinador", vbCritical, "Atenció": Exit Sub
   escriure_ini "Baixes", "numcomanda", form1.comanda, "comandes.ini"
   Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "palets.exe comandes.ini FiltrarBobinesImpresores", vbNormalFocus
End Sub

Private Sub mhistoricbobinesamaquines_Click()
  Load formseleccionou
  formseleccionou.caption = "Bobines a màquina"
  formseleccionou.Data1.DatabaseName = rutadelfitxer(cami) + "baixes"
  formseleccionou.Data1.RecordSource = "SELECT numbobina,data,operari,maquina from impresores_bobinesamaquina order by maquina,operari asc ,data desc"
  formseleccionou.refrescar
  formseleccionou.DBGrid2.Columns(0).width = 1000
  formseleccionou.DBGrid2.Columns(1).width = 2000
  formseleccionou.DBGrid2.Columns(2).width = 800
  formseleccionou.DBGrid2.Columns(3).width = 800
  formseleccionou.width = 6000
  formseleccionou.Show 1
     
End Sub

Private Sub reixa_DblClick()
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("select * from impresores_bobinesamaquina where numbobina='" + reixa.text + "'")
   If rst.EOF Then Exit Sub
   If cadbl(rst!operari) <> numop2 Then MsgBox "Aquesta bobina l'ha afegit l'Operari " + atrim(rst!operari) + " només ell pot eliminar-la.", vbCritical, "Error": GoTo fi
   If MsgBox("Vols eliminar aquesta bobina de la llista?", vbInformation + vbDefaultButton2 + vbYesNo, "Eliminar de la llista") = vbYes Then
       dbtmpb.Execute "delete * from impresores_bobinesamaquina where numbobina='" + reixa.text + "'"
       carregar_bobines_baixades
   End If
fi:
   Set rst = Nothing
End Sub

Private Sub reixapicus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 2 Then ccodidebarrespicus = reixapicus.text: ccodidebarrespicus.SetFocus: Sendkeys "{ENTER}"
End Sub

Private Sub Timer1_Timer()
   If LCase(arguments(1)) = "desbobinadors" Then llegir_comanda_compartida_impresores
End Sub

Private Sub treurebobina1_Click()
   If MsgBox("Estas segur que vols treure aquesta bobina del DESBOBINADOR 1?", vbExclamation + vbDefaultButton2 + vbYesNo, "TREURE BOBINA") = vbYes Then
      escriure_ini "Bobines_Desbobinadors_" + atrim(nummaq), "Bobina1", "", rutadelfitxer(cami) + "valorsprograma.ini"
      escriure_ini "Bobines_Desbobinadors_" + atrim(nummaq), "Horabob1", "", rutadelfitxer(cami) + "valorsprograma.ini"
      carregar_bobines_desbobinadors
   End If
End Sub

Private Sub treurebobina2_Click()
If MsgBox("Estas segur que vols treure aquesta bobina del DESBOBINADOR 2?", vbExclamation + vbDefaultButton2 + vbYesNo, "TREURE BOBINA") = vbYes Then
      escriure_ini "Bobines_Desbobinadors_" + atrim(nummaq), "Bobina2", "", rutadelfitxer(cami) + "valorsprograma.ini"
      escriure_ini "Bobines_Desbobinadors_" + atrim(nummaq), "Horabob2", "", rutadelfitxer(cami) + "valorsprograma.ini"
      carregar_bobines_desbobinadors
   End If
End Sub
