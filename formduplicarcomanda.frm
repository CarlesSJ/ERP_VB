VERSION 5.00
Begin VB.Form formduplicarcomanda 
   Caption         =   "Duplicar Comanda"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6630
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "formduplicarcomanda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6630
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Escullir Client "
      Height          =   3345
      Left            =   195
      TabIndex        =   3
      Top             =   15
      Width           =   6270
      Begin VB.CommandButton Command9 
         Height          =   405
         Index           =   0
         Left            =   3360
         Picture         =   "formduplicarcomanda.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Pdf amb el risc resumit."
         Top             =   1860
         Width           =   675
      End
      Begin VB.CommandButton sortirs 
         Height          =   375
         Left            =   5790
         Picture         =   "formduplicarcomanda.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Sortir del Programa"
         Top             =   120
         Width           =   435
      End
      Begin VB.Timer Timer1 
         Interval        =   600
         Left            =   5790
         Top             =   2340
      End
      Begin VB.CommandButton Command9 
         Height          =   405
         Index           =   1
         Left            =   2655
         Picture         =   "formduplicarcomanda.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir full de control de Risc."
         Top             =   1860
         Width           =   675
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H008080FF&
         Caption         =   "RISC"
         Height          =   330
         Left            =   1410
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Per saber el risc d'aquest client (PROVES)"
         Top             =   135
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   435
         Left            =   5220
         Picture         =   "formduplicarcomanda.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Donar Vistiplau"
         Top             =   1845
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox codicomptable 
         Height          =   315
         Left            =   2670
         TabIndex        =   10
         Top             =   1515
         Width           =   3345
      End
      Begin VB.ComboBox nomclient 
         Height          =   315
         Left            =   2670
         TabIndex        =   6
         Top             =   495
         Width           =   3360
      End
      Begin VB.ComboBox direccioenvio 
         Height          =   315
         Left            =   2670
         TabIndex        =   5
         Top             =   975
         Width           =   3345
      End
      Begin VB.TextBox codiclient 
         Height          =   285
         Left            =   1365
         TabIndex        =   4
         Top             =   495
         Width           =   1005
      End
      Begin VB.Label etformadepagament 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005C31DD&
         Height          =   510
         Left            =   60
         TabIndex        =   21
         Top             =   255
         Width           =   1710
      End
      Begin VB.Label etnopredeterminat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CODI COMPTABLE    NO PREDETERMINAT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   525
         Left            =   3405
         TabIndex        =   18
         Top             =   2265
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label etcomandesazero 
         BackStyle       =   0  'Transparent
         Caption         =   "----"
         ForeColor       =   &H000000FF&
         Height          =   600
         Left            =   60
         TabIndex        =   16
         Top             =   2700
         Width           =   6120
      End
      Begin VB.Label ettotalrisc 
         BackStyle       =   0  'Transparent
         Caption         =   "----"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   45
         TabIndex        =   15
         Top             =   2475
         Width           =   5670
      End
      Begin VB.Label etrisc 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00ED823A&
         Height          =   1875
         Left            =   90
         TabIndex        =   13
         Top             =   780
         Width           =   2550
      End
      Begin VB.Label Label2 
         BackColor       =   &H00209BCA&
         BackStyle       =   0  'Transparent
         Caption         =   "Facturació - Comptable"
         Height          =   285
         Left            =   2670
         TabIndex        =   11
         Top             =   1335
         Width           =   2445
      End
      Begin VB.Label Label1 
         BackColor       =   &H00209BCA&
         BackStyle       =   0  'Transparent
         Caption         =   "Codi Client"
         Height          =   285
         Left            =   1455
         TabIndex        =   9
         Top             =   255
         Width           =   900
      End
      Begin VB.Label Label5 
         BackColor       =   &H00209BCA&
         BackStyle       =   0  'Transparent
         Caption         =   "Nom del Client"
         Height          =   285
         Left            =   2670
         TabIndex        =   8
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label6 
         BackColor       =   &H00209BCA&
         BackStyle       =   0  'Transparent
         Caption         =   "Direcció d'Envio"
         Height          =   285
         Left            =   2670
         TabIndex        =   7
         Top             =   795
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   195
      TabIndex        =   0
      Top             =   3435
      Width           =   6270
      Begin VB.CommandButton repetir 
         Caption         =   "Repeticio / Modificació"
         Height          =   1500
         Left            =   4140
         Picture         =   "formduplicarcomanda.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton canviproducte 
         Caption         =   "Canvi de Producte/Treball"
         Height          =   1500
         Left            =   540
         Picture         =   "formduplicarcomanda.frx":42AC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   1755
      End
   End
End
Attribute VB_Name = "formduplicarcomanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub acceptar_Click()

End Sub

Private Sub canviproducte_Click()
   If Not revisarcredit Then Exit Sub
   If escorrecteelclient Then
      canviproducte.Tag = "1"
        Else: Exit Sub
   End If
   formduplicarcomanda.Tag = ""
   Me.Hide
End Sub
Function escorrecteelclient() As Boolean
   If cadbl(codicomptable.Tag) = 0 Or cadbl(direccioenvio.Tag) = 0 Or cadbl(codiclient) = 0 Then MsgBox "Falta escullir dades dels client.", vbCritical, "Error": Exit Function
   escorrecteelclient = True
End Function

Private Sub codicomptable_DropDown()
   triar_client_codicomptable
   
End Sub


Private Sub Command1_Click()
    If Not revisarcredit Then Exit Sub
    formduplicarcomanda.Tag = ""
    formduplicarcomanda.Hide
End Sub

Private Sub Command2_Click()
   Dim risc As Double
   risc = comprovarrisc(cadbl(codicomptable.Tag))
   If risc > -1 Then
    MsgBox "Aquest client te un risc assignat de :" + atrim(risc)
   End If
End Sub

Private Sub Command9_Click(Index As Integer)
  If Index = 1 Then informe_credit_unclient codicomptable.Tag
  If Index = 0 Then llistat_resum_credit codicomptable.Tag
End Sub
Sub llistat_resum_credit(vcodicomptable As String)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.report
  Dim vrisc As TipusVrisc
  Dim vriscreal As Double
  Dim vresum As String
  Dim vriscsuposat As Double
  Dim vnomfitxerPDF As String
  calcular_credit_delclient cadbl(vcodicomptable), vrisc
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistat_resum_credit.rpt", 1)
  'oreport.Database.Tables.Item(1).Location = ""
  vriscreal = Redondejar(vrisc.creditsap - vrisc.creditgastatsap - vrisc.valoralbaranspendentsSAP - vrisc.valorestoc, 0)
  vriscsuposat = Redondejar(vriscreal - vrisc.valorproduccio - vrisc.valorpendent, 0)
  oreport.FormulaFields.GetItemByName("creditconcedit").Text = cadbl(Redondejar(vrisc.creditsap))
  oreport.FormulaFields.GetItemByName("creditconsumit").Text = cadbl(Redondejar(vrisc.creditgastatsap, 0))
  oreport.FormulaFields.GetItemByName("albarapendent").Text = cadbl(Redondejar(vrisc.valoralbaranspendentsSAP))
  oreport.FormulaFields.GetItemByName("comandesestoc").Text = cadbl(Redondejar(vrisc.valorestoc))
  oreport.FormulaFields.GetItemByName("riscreal").Text = cadbl(vriscreal)
  oreport.FormulaFields.GetItemByName("enproduccio").Text = cadbl(Redondejar(vrisc.valorproduccio))
  oreport.FormulaFields.GetItemByName("pendent").Text = cadbl(Redondejar(vrisc.valorpendent))
  oreport.FormulaFields.GetItemByName("riscsuposat").Text = cadbl(Redondejar(vriscsuposat))
  oreport.FormulaFields.GetItemByName("comandesazero").Text = """" + Trim(vrisc.comandesazerodetall) + "   ==============     " + "    " + "Total Kg: " + Format(Redondejar(vrisc.comandesazeroTotalKg), "#,##0") + "Kg" + """"
  oreport.FormulaFields.GetItemByName("nomclient").Text = """" + Trim(vrisc.nomdelclient) + """"

  
'  Load veurereport
'  veurereport.CRViewer.ReportSource = oreport
'  veurereport.CRViewer.DisplayGroupTree = False
'  veurereport.CRViewer.ViewReport
'  veurereport.WindowState = 2
'  veurereport.Show 1
   If Not existeix("c:\temp\riscclients") Then MkDir "c:\temp\riscclients"
   borrar_temporal_credit
   vnomfitxerPDF = "c:\temp\riscclients\Informe_risc_client_" + vrisc.nomdelclient + ".pdf"
   oreport.ExportOptions.DiskFileName = vnomfitxerPDF
   oreport.ExportOptions.PDFExportAllPages = True
   oreport.ExportOptions.FormatType = crEFTPortableDocFormat
   oreport.ExportOptions.DestinationType = crEDTDiskFile
   oreport.Export False
   If existeix(vnomfitxerPDF) Then
        Shell "c:\windows\system32\cmd.exe /c start mailto:"
        wait 2
        idp = ShellExecute(Me.hwnd, "Open", "c:\windows\explorer.exe", " " + "c:\temp\riscclients", "", 1)
          Else: MsgBox "Error no s'ha creat el fitxer PDF.", vbCritical, "Error"
   End If
End Sub
Sub borrar_temporal_credit()
 On Error Resume Next
  Kill "c:\temp\riscclients\*.*"
End Sub

Private Sub direccioenvio_DropDown()
  triar_client_direnvio
End Sub

Function comprovarrisc(codic As Double) As Double
  Dim parametres As String
  Dim resp As String
  If existeix("c:\temp\risc.ini") Then Kill "c:\temp\risc.ini"
  parametres = " Empresa=INPLACSA_FB Funcion=Riesgo Cliente=" + atrim(codic) + " Archivo=c:\temp\risc.ini"
'  MsgBox "\\serconta\home\OPENSOFT\BIN\PCalculaDato.exe " + parametres
  Shell "\\serconta\home\OPENSOFT\BIN\PCalculaDato.exe " + parametres
  ratoli "espera"
  resp = llegir_ini("Archivo", "Estado", "c:\temp\risc.ini")
  While resp <> "OK" And resp <> "ERROR"
    DoEvents
    resp = llegir_ini("Archivo", "Estado", "c:\temp\risc.ini")
  Wend
  If resp = "OK" Then comprovarrisc = cadbl(llegir_ini("Detalle riesgo", "Riesgo concedido", "c:\temp\risc.ini"))
  If resp = "ERROR" Then comprovarrisc = -99999999
  ratoli "normal"
End Function

Private Sub etcomandesazero_DblClick()
InputBox "Comandes amb valor zero", "Comandes zero", etcomandesazero
End Sub
Sub comprovarsicodicomptablepredeterminat()
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select codicomptable,nomclient,predeterminat from clients_codiscomptables where codifabricacio=" + atrim(cadbl(codiclient)) + " and predeterminat")
  If Not rst.EOF Then
      If atrim(rst!codicomptable) <> codicomptable.Tag Then
         etnopredeterminat.Tag = "1"
          Else: etnopredeterminat.Tag = ""
      End If
  End If
  Set rst = Nothing
End Sub

Private Sub Form_Activate()
   comprovarsicodicomptablepredeterminat
   If Frame2.Tag = "0" Then Exit Sub
   If direccioenvio = "" And Frame2.Tag <> "1" Then
     Frame2.Tag = "0"
     Unload formseleccio
     triar_client_direnvio
     'wait 1
   End If
   comprovarsiexisteixcodicomptablealsap
   If codicomptable = "" And Frame2.Tag <> "1" Then
      Frame2.Tag = "0"
      Unload formseleccio
      triar_client_codicomptable
      'wait 1
   End If
   Frame2.Tag = "1"
End Sub
Sub comprovarsiexisteixcodicomptablealsap()
   Dim rst As Recordset
   If cadbl(codicomptable.Tag) = 0 Then Exit Sub
   Set rst = dbtmp.OpenRecordset("select * from clients_codisSAP where codiSAP=" + atrim(cadbl(codicomptable.Tag)))
   If rst.EOF Then
      MsgBox "El codi comptable " + codicomptable.Tag + " no existeix a SAP o cal actualitzar les dades de fabricació amb el SAP." + Chr(10) + " S'ha d'assignar un altra codi de client comptable", vbCritical + vbOKOnly, "Atenció"
      codicomptable = ""
      codicomptable.Tag = ""
   End If
   Set rst = Nothing
End Sub
Private Sub Form_Click()
  'nomclient_DropDown
End Sub

Private Sub Form_Load()
  
   canviproducte.Tag = ""
   repetir.Tag = ""
   formduplicarcomanda.Tag = "sortir"
   codiclient = cadbl(formcomandes.Data1.Recordset!client)
   nomclient = atrim(formcomandes.nomclient)
   direccioenvio.Tag = cadbl(formcomandes.Data1.Recordset!direnvio)
   direccioenvio = formcomandes.nomclient.Tag
   codicomptable.Tag = formcomandes.Text32(3).Tag
   codicomptable = formcomandes.Text32(3)
  ensenyarelcredit
  ratoli "normal"
End Sub
Sub posarformadepagament(vcodiclient As Double)
   Dim rst As Recordset
   etformadepagament = ""
   Set rst = dbtmp.OpenRecordset("select * from clients_codisSAP where codiSAP=" + atrim(vcodiclient))
   If Not rst.EOF Then
       etformadepagament = "F. pagament:" + vbNewLine + atrim(rst!formadepagament)
   End If
   Set rst = Nothing
End Sub

Private Sub guardar_credit_client(vcodi As Double, vrisc As TipusVrisc)
  Dim rstc As Recordset
  Set rstc = dbtmp.OpenRecordset("Select * from clients_codisSAP where codisap=" + atrim(vcodi))
  If Not rstc.EOF Then
        rstc.Edit
        rstc!creditsap = cadbl(vrisc.creditsap)
        rstc!valordiferencial = Redondejar(cadbl(vrisc.valordiferencial), 0)
        rstc!creditgastatsap = cadbl(vrisc.creditgastatsap)
        rstc!valorestoc = cadbl(vrisc.valorestoc)
        rstc!valorpendent = cadbl(vrisc.valorpendent)
        rstc!valorproduccio = cadbl(vrisc.valorproduccio)
        rstc!valordelsclixes = cadbl(vrisc.valordelsclixes)
        rstc!valoralbaranspendentsSAP = cadbl(vrisc.valoralbaranspendentsSAP)
        rstc.Update
  End If
  Set rstc = Nothing
End Sub
Sub ensenyarelcredit()
   Dim vrisc As TipusVrisc
   Dim vtotalrisc As Double
   'Command9(1).Visible = False
   'calcular_credit_delclient cadbl(codicomptable.Tag), vcomandesaPVPzero, vconcedit, vconsumitsap, vstock, vproduccio, vpendent
   calcular_credit_delclient cadbl(codicomptable.Tag), vrisc
   guardar_credit_client cadbl(codicomptable.Tag), vrisc
   posarformadepagament cadbl(codicomptable.Tag)
  ' MsgBox vrisc.comandesestoc
   etrisc.Tag = atrim(vrisc.creditsap)
   etrisc = "C.Concedit:" + justificar(Format(vrisc.creditsap, "#,##0") + "€", 12, "D") + Chr(10)
   etrisc = etrisc + "C.Consumit:" + justificar(Format(vrisc.creditgastatsap, "#,##0") + "€", 12, "D") + Chr(10)
   etrisc = etrisc + "V.Estoc:   " + justificar(Format(vrisc.valorestoc, "#,##0") + "€", 12, "D") + Chr(10)
   etrisc = etrisc + "V.Produc.: " + justificar(Format(vrisc.valorproduccio, "#,##0") + "€", 12, "D") + Chr(10)
   etrisc = etrisc + "V.Pendent: " + justificar(Format(vrisc.valorpendent, "#,##0") + "€", 12, "D") + Chr(10)
   etrisc = etrisc + "V.Alb.SAP: " + justificar(Format(vrisc.valoralbaranspendentsSAP, "#,##0") + "€", 12, "D") + Chr(10)
   'etrisc = etrisc + "Valor Pendent:    " + atrim(Format(vrisc.valorpendent, "#,##0")) + "€" + Chr(10)
   etrisc = etrisc + "V.Clixes:  " + justificar(Format(vrisc.valordelsclixes, "#,##0") + "€", 12, "D") + Chr(10)
   etrisc = etrisc + "======================"
   vtotalrisc = vrisc.creditsap - vrisc.creditgastatsap - vrisc.valorestoc - vrisc.valorproduccio - vrisc.valorpendent - vrisc.valoralbaranspendentsSAP
   ettotalrisc = "Total Diferencial:  " + atrim(Format(vtotalrisc, "#,##0")) + "€"
   ettotalrisc.Tag = vtotalrisc
   If vtotalrisc < 0 Then
      ettotalrisc.ForeColor = QBColor(12)
      'Command9(1).Visible = True
       Else: ettotalrisc.ForeColor = QBColor(0)
   End If
   'etinfocredit = "CS:" + atrim(Format(vconcedit, "#,##0")) + " CcS:" + atrim(Format(vconsumitsap, "#,##0"))
   'etinfocredit = etinfocredit + " Vst:" + atrim(Format(vstock, "#,##0")) + " Vpr:" + atrim(Format(vproduccio, "#,##0")) + " Vpnt:" + atrim(Format(vpendent, "#,##0"))
   If vrisc.comandesazero <> "" Then
       etcomandesazero = "Comandes a Zero: " + vrisc.comandesazero
        Else: etcomandesazero = ""
   End If
End Sub


Sub triar_client_imp()
 Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codi,nom  from clients"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Width = 4200
  formseleccio.Show 1
  
   If seleccioret = 1 Then
        
        If Not formseleccio.Data1.Recordset.EOF Then
           nomclient = formseleccio.DBGrid2.Columns("nom")
           codiclient = cadbl(formseleccio.DBGrid2.Columns("codi"))
           nomclient.Tag = "nou"
        End If
   End If
    If seleccioret = 9 Then
        nomclient = ""
        codiclient = ""
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   SendKeys "{TAB}"
   'codimuntadora.SetFocus
End Sub
Sub triar_client_direnvio()
  If cadbl(codiclient) < 1 Then MsgBox "Primer has d'escullir un client"
    'Unload formseleccio
   Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select id ,nome,domicilie,poblacioe,provinciae from clients_envios where codi=" + atrim(cadbl(codiclient))
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(2).Width = 900
  formseleccio.Width = 9000
  formseleccio.Left = formseleccio.Left - 3000
  If formseleccio.Data1.Recordset.EOF Then MsgBox "Aquest client no te cap DIRECCIO D'ENVIO ASSIGNADA.": Exit Sub
  formseleccio.Data1.Recordset.MoveLast
  formseleccio.Data1.Recordset.MoveFirst
  If formseleccio.Data1.Recordset.RecordCount > 1 Then
                                                                                                                
     formseleccio.Show 1
     While formseleccio.Visible And seleccioret = 0
       DoEvents
     Wend
    Else: seleccioret = 1
  End If
  
  
  
   If seleccioret = 1 Then
        
        If Not formseleccio.Data1.Recordset.EOF Then
           direccioenvio = formseleccio.DBGrid2.Columns("poblacioe")
           direccioenvio.Tag = cadbl(formseleccio.DBGrid2.Columns("id"))
           'campid_treball.Tag = cadbl(formseleccio.DBGrid2.Columns("ordremodificacio"))
        End If
   End If
    If seleccioret = 9 Then
        direccioenvio = ""
        direccioenvio.Tag = ""
        'campid_treball.Tag = ""
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   SendKeys "{TAB}"
   'codimuntadora.SetFocus
End Sub

Private Sub nomclientnouimp_Change()

End Sub

Private Sub nomclient_DropDown()
  triar_client_imp
  If nomclient.Tag = "nou" Then
    nomclient.Tag = ""
    
    direccioenvio.Tag = ""
    direccioenvio = ""
    direccioenvio = ""
    codicomptable = ""
    codicomptable.Tag = ""
    triar_client_direnvio
    triar_client_codicomptable
  End If
End Sub
Sub triar_client_codicomptable()
  If cadbl(codiclient) < 1 Then MsgBox "Primer has d'escullir un client"
  'Unload formseleccio
   Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codicomptable,nomclient,predeterminat from clients_codiscomptables where codifabricacio=" + atrim(cadbl(codiclient)) + " order by predeterminat asc"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Width = 2000
  'formseleccio.Width = 9000
  'formseleccio.Left = formseleccio.Left - 3000
  If formseleccio.Data1.Recordset.EOF Then MsgBox "Aquest client no te cap CODI COMPTABLE ASSIGNAT.": Exit Sub
  formseleccio.Data1.Recordset.MoveLast
  formseleccio.Data1.Recordset.MoveFirst
  If formseleccio.Data1.Recordset.RecordCount > 1 Then
         formseleccio.Show 1
     While formseleccio.Visible
        DoEvents
     Wend
    Else: seleccioret = 1
  End If
  
   If seleccioret = 1 Then
        
        If Not formseleccio.Data1.Recordset.EOF Then
           codicomptable = formseleccio.DBGrid2.Columns("codicomptable") + " - " + formseleccio.DBGrid2.Columns("nomclient")
           codicomptable.Tag = cadbl(formseleccio.DBGrid2.Columns("codicomptable"))
           'campid_treball.Tag = cadbl(formseleccio.DBGrid2.Columns("ordremodificacio"))
        End If
   End If
    If seleccioret = 9 Then
        codicomptable = ""
        codicomptable.Tag = ""
        'campid_treball.Tag = ""
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   SendKeys "{TAB}"
   ensenyarelcredit
   'codimuntadora.SetFocus
   comprovarsicodicomptablepredeterminat
End Sub
Function esclixevalid(numtreball As Double) As Boolean
   Dim rstc As Recordset
   If numtreball = 0 Then esclixevalid = True: Exit Function
   esclixevalid = False
   Set rstc = dbclixes.OpenRecordset("select databaixaclixe from clixes where id_treball=" + atrim(numtreball))
   If rstc.EOF Then Exit Function
   If Not IsDate(rstc!databaixaclixe) Then esclixevalid = True
End Function
Function revisarcredit() As Boolean
   Dim vmsg As String
   revisarcredit = True
   If cadbl(ettotalrisc.Tag) < 2 Then
       v = InputBox("Aquest client no te prou crèdit per fer aquesta comanda." + Chr(10) + "Si vols continuar igualment escriu [endavant] sino prem cancelar." + Chr(10) + "S'enviarà un email de crèdit insuficient.", "Error de crèdit")
       If UCase(v) <> "ENDAVANT" Then revisarcredit = False: GoTo fi
       vmsg = "S'ha creat una comanda amb diferencial de risc zero o inferior."
       vmsg = vmsg + "Ordinador: " + UCase(nomordinador) + Chr(13) + Chr(10)
       vmsg = vmsg + "Client: " + atrim(nomclient) + Chr(13) + Chr(10)
       vmsg = vmsg + "Direccio: " + direccioenvio + Chr(13) + Chr(10)
       vmsg = vmsg + "Codi Comptable: " + codicomptable + Chr(13) + Chr(10)
       vmsg = vmsg + "Risc : " + etrisc.Tag + "€" + Chr(13) + Chr(10)
       vmsg = vmsg + "Risc Diferencial : " + ettotalrisc.Tag + "€" + Chr(13) + Chr(10)
       codiclient.Tag = vmsg
   End If
fi:
End Function
Private Sub repetir_Click()
   If Not revisarcredit Then Exit Sub
   If escorrecteelclient Then
      repetir.Tag = "1"
       Else: Exit Sub
   End If
   formduplicarcomanda.Tag = ""
   If InStr(1, ruta, "I") > 0 Then
     If Not esclixevalid(cadbl(formcomandes.Data1.Recordset!numtreball)) Then MsgBox "El clixé d'aquesta comanda està donat de baixa, no pots utilitzar-lo", vbCritical, "Atenció": Exit Sub
   End If
   Me.Hide
   
End Sub

Private Sub sortirs_Click()
  formduplicarcomanda.Tag = "sortir"
  Me.Hide
End Sub

Private Sub Timer1_Timer()
  If etnopredeterminat.Tag = "1" Then
      etnopredeterminat.Visible = Not etnopredeterminat.Visible
  End If
End Sub
