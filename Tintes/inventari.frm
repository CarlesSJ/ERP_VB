VERSION 5.00
Begin VB.Form forminventari 
   Caption         =   "Regularització de llaunes"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10770
   Icon            =   "inventari.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Llegir llaunes Off-line"
      Height          =   435
      Left            =   8130
      TabIndex        =   18
      Top             =   45
      Width           =   1710
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estantaries"
      Height          =   7860
      Left            =   465
      TabIndex        =   0
      Top             =   420
      Width           =   9885
      Begin VB.CommandButton Command3 
         Caption         =   "Llistat de llaunes pendents d'inventari"
         Height          =   570
         Left            =   6165
         TabIndex        =   17
         Top             =   3975
         Width           =   1665
      End
      Begin VB.CommandButton bimprimirllauna 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Llauna"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   3735
         Picture         =   "inventari.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Imprimir llauna"
         Top             =   3765
         Width           =   810
      End
      Begin VB.CommandButton Command33 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Situació"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2925
         Picture         =   "inventari.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Canviar situacio de la llauna"
         Top             =   3765
         Width           =   810
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Llistat de llaunes sense moviment"
         Height          =   570
         Left            =   7860
         TabIndex        =   14
         Top             =   3975
         Width           =   1665
      End
      Begin VB.Frame Frame3 
         Caption         =   "Informació de la llauna"
         Height          =   2415
         Left            =   2565
         TabIndex        =   5
         Top             =   1185
         Width           =   6975
         Begin VB.TextBox datainventari 
            Height          =   300
            Left            =   5760
            TabIndex        =   12
            ToolTipText     =   "Només s'ha de canviar al principi de l'inventari."
            Top             =   1920
            Width           =   1020
         End
         Begin VB.CommandButton Command1 
            Height          =   450
            Left            =   2355
            Picture         =   "inventari.frx":1BCE
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "S'actualitzarà els kilos de la llauna."
            Top             =   1845
            Width           =   645
         End
         Begin VB.TextBox pesactual 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1335
            TabIndex        =   8
            Top             =   1890
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Inventari"
            Height          =   210
            Left            =   5730
            TabIndex        =   13
            Top             =   1650
            Width           =   1155
         End
         Begin VB.Label etpesnet 
            Height          =   330
            Left            =   3180
            TabIndex        =   11
            Top             =   1950
            Width           =   1800
         End
         Begin VB.Label ettara 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tara Llauna:"
            Height          =   195
            Left            =   180
            TabIndex        =   10
            Top             =   1500
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Pes actual:"
            Height          =   210
            Left            =   180
            TabIndex        =   7
            Top             =   1950
            Width           =   990
         End
         Begin VB.Label etinfo 
            BackColor       =   &H00E0E0E0&
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
            Left            =   120
            TabIndex        =   6
            Top             =   285
            Width           =   6660
         End
      End
      Begin VB.TextBox numllauna 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   285
         TabIndex        =   4
         Top             =   1245
         Width           =   1980
      End
      Begin VB.ListBox llistallaunes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5820
         Left            =   270
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1830
         Width           =   2010
      End
      Begin VB.Frame Frame2 
         Caption         =   "Situació"
         Height          =   870
         Left            =   300
         TabIndex        =   1
         Top             =   315
         Width           =   2040
         Begin VB.ComboBox combosituacio 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   150
            TabIndex        =   2
            Top             =   285
            Width           =   1650
         End
      End
   End
   Begin VB.Label etllistat 
      BackStyle       =   0  'Transparent
      Caption         =   "Generant el llistat de llaunes escanejades offline"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   975
      TabIndex        =   19
      Top             =   45
      Visible         =   0   'False
      Width           =   9285
   End
End
Attribute VB_Name = "forminventari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub carregarsituacionsalcombo()
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from situacionsllaunes order by situacio")
   While Not rst.EOF
     combosituacio.AddItem UCase(rst!situacio)
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Private Sub bimprimirllauna_Click()
   Dim vnumllauna As String
   vnumllauna = InputBox("Escriu la llauna que vols imprimir.", "Imprimir Llauna", numllauna)
   If vnumllauna <> "" Then imprimir_etiqueta vnumllauna
End Sub

Private Sub combosituacio_Click()
   carregar_llistallaunes
End Sub

Private Sub combosituacio_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub Command1_Click()
   Dim vpesnet As Double
   If Not IsDate(datainventari) Then MsgBox "No hi ha una data d'inventari vàlida.", vbCritical, "Error": Exit Sub
   vpesnet = cadbl(pesactual) - cadbl(ettara.tag)
   ferelretorndetinta numllauna, vpesnet, True, datainventari
End Sub

Private Sub Command2_Click()
   Dim vsql As String
   Dim vdata As String
    vdata = InputBox("Entra la data limit del moviment de llaunes.", "Ultim moviment")
    If Not IsDate(vdata) Then MsgBox "Aquesta data no es vàlida", vbCritical, "Error": Exit Sub
   'vsql = "SELECT Llaunes.numllauna FROM historiallauna LEFT JOIN Llaunes ON historiallauna.idnumllauna = Llaunes.id where historiallauna.datainventari=null GROUP BY Llaunes.numllauna "
   'vsql = vsql + " HAVING (((Max(historiallauna.data))<#" + Format(vdata, "mm/dd/yy") + "#));"
   vsql = "SELECT Llaunes.numllauna FROM historiallauna LEFT JOIN Llaunes ON historiallauna.idnumllauna = Llaunes.id Where (((Llaunes.activa) = True)) GROUP BY Llaunes.numllauna "
   vsql = vsql + " HAVING (((Max(historiallauna.data))<#" + Format(vdata, "mm/dd/yy") + "#));"

   If borrartaulatempllistatllaunessensemoviment Then
    dbtintes.Execute "select * into tmp_llistatllaunessensemoviment from dadesllaunes where numllauna in (" + vsql + ")"
    ratoli "espera"
    wait 3
    ratoli "normal"
    ferelllistat vdata
   End If
End Sub
Sub ferelllistat(vdata As String)
Dim rst As Recordset
  
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatdellaunesnoutilitzades.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "tintes.mdb"
  'oreport.RecordSelectionFormula = "{Llaunes.numllauna}='" + atrim(numllauna) + "'"
'  oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
  oreport.FormulaFields.GetItemByName("vdata").Text = "'(" + vdata + ")'"
  
  'oreport.PaperOrientation = crLandscape
  oreport.DiscardSavedData
 ' If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
  '  Else
  '    oreport.DisplayProgressDialog = False
 '     oreport.PrintOut False, 1
 ' End If
End Sub
Function borrartaulatempllistatllaunessensemoviment() As Boolean
  On Error GoTo errorllistatobert
  borrartaulatempllistatllaunessensemoviment = True
  dbtintes.Execute "drop table  tmp_llistatllaunessensemoviment "
   
   Exit Function
errorllistatobert:
   If err.Number <> 3376 Then
     MsgBox "Hi ha un error al generar el llistat, tanca el programa i torna-ho a provar", vbCritical, "Error"
     borrartaulatempllistatllaunessensemoviment = False
   End If
End Function

Private Sub Command3_Click()
  Dim nomfitxertemporal As String
  Dim vsionoinventariat As String
  Dim vdatainventari As String
  vdatainventari = InputBox("Entra la data d'inventari que vols consultar.", "Inventari", datainventari)
  vsionoinventariat = UCase(InputBox("Vols la llaunes inventariades(S) o les no inventariades(N).", "Inventari", "N"))
  If vdatainventari = "" Then Exit Sub
  crear_taula_temporal_llinventari nomfitxertemporal
  emplanar_llistat_inventari_tintes nomfitxertemporal, vdatainventari, vsionoinventariat
  ferelllistat_inventari nomfitxertemporal, vdatainventari, vsionoinventariat
  
End Sub
Sub ferelllistat_inventari(nomfitxertemporal As String, vdatainventari, vsionoinventariat As String)
  
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatinventari_tintes.rpt", 1)
  oreport.Database.Tables.Item(1).Location = nomfitxertemporal
  'oreport.RecordSelectionFormula = "{Llaunes.numllauna}='" + atrim(numllauna) + "'"
'  oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
  oreport.FormulaFields.GetItemByName("titolinforme").Text = "'" + IIf(vsionoinventariat = "N", "Llistat de llaunes NO inventeriades i sense moviment a dia ", "Llistat de llaunes INVENTERIADES a dia ") + vdatainventari + "'"
  
  'oreport.PaperOrientation = crLandscape
  oreport.DiscardSavedData
 ' If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
  '  Else
  '    oreport.DisplayProgressDialog = False
 '     oreport.PrintOut False, 1
 ' End If
End Sub
Sub emplanar_llistat_inventari_tintes(nomfitxertemporal As String, vdatainventari As String, vsionoinventariat As String)
  Dim rst As Recordset
  Dim vsql As String
  
  Dim rsttemp As Recordset
  Dim dbtemp As Database
  If Not existeix(nomfitxertemporal) Then MsgBox "No s'ha creat el fitxer temporal. Torna-ho a provar mes tard.", vbCritical, "Error": Exit Sub
  Set dbtemp = OpenDatabase(nomfitxertemporal)
  Set rsttemp = dbtemp.OpenRecordset("select * from llistat_inventari")
  vsql = "SELECT Llaunes.numllauna, tintes.descripcio, Last(Llaunes.situacio) AS ÚltimoDesituacio, last(historiallauna.data) as ultimmoviment FROM (Llaunes RIGHT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) LEFT JOIN tintes ON Llaunes.idtinta = tintes.idtinta "
  vsql = vsql + " Where (((Llaunes.activa) = True)) GROUP BY Llaunes.numllauna, tintes.descripcio HAVING (((Llaunes.numllauna) " + IIf(vsionoinventariat = "N", "Not", "") + " In (SELECT Llaunes.numllauna FROM Llaunes RIGHT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna"
  vsql = vsql + " WHERE (((historiallauna.datainventari)='" + vdatainventari + "'))))) order by last(historiallauna.data) DESC;"

  Set rst = dbtintes.OpenRecordset(vsql)
  If rst.EOF Then MsgBox "No hi ha registres", vbCritical, "Error": Exit Sub
  While Not rst.EOF
    If DateDiff("d", vdatainventari, rst!ultimmoviment) < 0 Or vsionoinventariat = "S" Then
     rsttemp.AddNew
     rsttemp!llauna = rst!numllauna
     rsttemp!descripcio = rst!descripcio
     rsttemp!situacio = rst!ÚltimoDeSituacio
     rsttemp.Update
    End If
    rst.MoveNext
  Wend
  Set rst = Nothing
  Set rsttemp = Nothing
  Set dbtemp = Nothing
End Sub
Sub crear_taula_temporal_llinventari(nomfitxertemporal As String)
  
  nomfitxertemporal = "c:\temp\~tintes_inv" + Format(Now, "ddmmhhnnss") + ".mdb"
  On Error Resume Next
   MkDir "c:\temp"
   Kill "c:\temp\~tintes_inv*.*"
   DBEngine.CreateDatabase nomfitxertemporal, dbLangGeneral, dbVersion10
   Set dbtemp = OpenDatabase(nomfitxertemporal)
   'dbtemp.Execute "drop table tmp_imp_empalmes"
  On Error GoTo 0
  camps = "llauna string,descripcio string,situacio string"
  dbtemp.Execute ("create table llistat_inventari (" + camps) + ")"
End Sub

Private Sub Command33_Click()
  Dim i As Byte
  Load formsituacio
  formsituacio.numllauna = numllauna
  formsituacio.afegeixllauna
  formsituacio.carregarsituacionsalcombo
  formsituacio.Show 1
End Sub

Private Sub Command4_Click()
  Open "c:\temp\llaunes_offline.txt" For Output As #1
  Close 1
  Shell "notepad 'c:\temp\llaunes_offline.txt'", vbNormalFocus
  End
End Sub

Private Sub datainventari_LostFocus()
   If Not IsDate(datainventari) Then MsgBox "Data erronea", vbCritical, "Error": datainventari = "": Exit Sub
   escriure_ini "Tintes", "ultiminventari", datainventari, "comandes.ini"
End Sub

Private Sub Form_Activate()
  If Not etllistat.visible Then comprovar_llaunes_offline
End Sub

Private Sub Form_Load()
  carregarsituacionsalcombo
  datainventari = llegir_ini("Tintes", "ultiminventari", "comandes.ini")
  If datainventari = "{[}]" Then
    datainventari = Format(Now, "dd/mm/yy")
    escriure_ini "Tintes", "ultiminventari", datainventari, "comandes.ini"
  End If
  
  
End Sub
Sub fer_el_llistat()
  'llistat_inventari_llaunesoffline.rpt
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vllaunes As String
  Dim vnumc As Double
  Dim vprinter As Printer
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistat_inventari_llaunesoffline.rpt")
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "tintes.mdb"
  'oreport.RecordSelectionFormula = "{Llaunes.numllauna}='" + UCase(atrim(numllauna)) + "'"
   'report.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
 ' oreport.Sections("D").ReportObjects.Item("serie2").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
  oreport.DiscardSavedData
 
  'oreport.FormulaFields.GetItemByName("observacions").Text = "'" + atrim(observacions) + "'"
  'oreport.PaperOrientation = crPortrait
  'If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
 '   Else
 '     oreport.DisplayProgressDialog = False
 '     oreport.PrintOut False, 1
 ' End If
  'Unload Me
  
End Sub
Sub comprovar_llaunes_offline()
  Dim v As String
  If existeix("c:\temp\llaunes_offline.txt") Then
    Open "c:\temp\llaunes_offline.txt" For Input As #1
    If EOF(1) Then GoTo fi
    Line Input #1, v
    If Len(atrim(v)) > 4 Then
      DoEvents
      etllistat.visible = True
      Frame1.Enabled = False
      Command4.Enabled = False
      generar_llistat_llaunesoffline v
      wait 3
      fer_el_llistat
      Close 1
      Kill "c:\temp\llaunes_offline.txt"
      Unload forminventari
    End If
  End If
  
fi:
  Close 1
  
End Sub
Sub generar_llistat_llaunesoffline(v As String)
   dbtintes.Execute "delete * from llistat_llaunesoffline"
   While v <> "" And Not EOF(1)
      dbtintes.Execute "insert into llistat_llaunesoffline (numllauna) values ('" + v + "')"
      Line Input #1, v
      If EOF(1) And v <> "" Then dbtintes.Execute "insert into llistat_llaunesoffline (numllauna) values ('" + v + "')"
   Wend
   dbtintes.Execute "UPDATE llistat_llaunesoffline LEFT JOIN Llaunes ON llistat_llaunesoffline.numllauna = Llaunes.numllauna SET llistat_llaunesoffline.observacio = IIf([llaunes].[activa]=True,'Activada','Anulada');"
   dbtintes.Execute "INSERT INTO llistat_llaunesoffline ( numllauna, observacio ) SELECT Llaunes.numllauna, 'NO ESCANEJADA' AS Expr1 From Llaunes WHERE (((Llaunes.activa)=True) AND ((Llaunes.numllauna) Not In (SELECT numllauna FROM LLISTAT_LLAUNESOFFLINE)));"
   dbtintes.Execute "UPDATE (llistat_llaunesoffline INNER JOIN Llaunes ON llistat_llaunesoffline.numllauna = Llaunes.numllauna) INNER JOIN tintes ON Llaunes.idtinta = tintes.idtinta SET llistat_llaunesoffline.descripcio = [tintes].[descripcio], llistat_llaunesoffline.situacio = [llaunes].[situacio];"

End Sub
Sub carregar_llistallaunes()
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from llaunes where situacio='" + atrim(combosituacio.Text) + "' and activa=true order by mid(numllauna,2)")
   llistallaunes.Clear
   numllauna = ""
   While Not rst.EOF
      llistallaunes.AddItem atrim(rst!numllauna)
      rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Private Sub llistallaunes_Click()
  If Screen.ActiveControl.Name = "llistallaunes" Then numllauna = llistallaunes.Text
End Sub

Sub buscar_llauna(vnumllauna As String)
   Dim i As Integer
   For i = 0 To llistallaunes.ListCount - 1
      If UCase(vnumllauna) = Mid(llistallaunes.List(i), 1, Len(vnumllauna)) Then
        llistallaunes.ListIndex = i
        GoTo fi
      End If
   Next i
   llistallaunes.ListIndex = -1
fi:
End Sub
Sub possar_info(vnumllauna As String)
   Dim rst As Recordset
   Dim rsthistoria As Recordset
   Set rst = dbtintes.OpenRecordset("select * from dadesllaunes where numllauna='" + vnumllauna + "'")
   ettara = "": etinfo = ""
   ettara.tag = ""
   ettara.caption = ""
   pesactual = ""
   etinfo.BackColor = &HE0E0E0
   If Not rst.EOF Then
      Set rsthistoria = dbtintes.OpenRecordset("select * from historiallauna where idnumllauna=" + atrim(rst!id) + " and datainventari='" + datainventari + "'")
      If rsthistoria.EOF Then
         etinfo.BackColor = &HE0E0E0
           Else: etinfo.BackColor = &H8080FF
      End If
      Set rsthistoria = dbtintes.OpenRecordset("select * from historiallauna where idnumllauna=" + atrim(rst!id) + " and datainventari<>'' order by data desc")
      etinfo = atrim(rst!numllauna) + " - " + atrim(rst!descripcio) + Chr(10) + atrim(rst!nombido) + Chr(10) + "Kg: " + atrim(rst!capacitatactual)
      If Not rsthistoria.EOF Then etinfo = etinfo + Chr(10) + "Ultim inventari: " + atrim(rsthistoria!datainventari)
      ettara.caption = "Tara llauna: " + atrim(rst!tara) + " Kg"
      ettara.tag = atrim(rst!tara)
      pesactual = atrim(cadbl(rst!capacitatactual) + cadbl(rst!tara))
   End If
End Sub
Private Sub numllauna_Change()
   carregar_infollauna
End Sub
Sub carregar_infollauna()
   buscar_llauna numllauna
   If UCase(numllauna) = llistallaunes.Text Then
      possar_info numllauna
      pesactual.SetFocus
      pesactual.SelStart = 0
      pesactual.SelLength = Len(pesactual)
       Else: possar_info ""
   End If
End Sub

Private Sub numllauna_KeyUp(KeyCode As Integer, Shift As Integer)
  carregar_infollauna
End Sub

Private Sub pesactual_Change()
   If cadbl(pesactual) > 25 Then If MsgBox("El pes es superior a 25Kg es correcte?", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   etpesnet = "Pes net: " + atrim(cadbl(pesactual) - cadbl(ettara.tag)) + " KG"
   If (cadbl(pesactual) - cadbl(ettara.tag)) < 0 Then etpesnet = "Pes net: 0 KG"
End Sub
