VERSION 5.00
Begin VB.Form comprespalets 
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "comprespalets.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "Crear Palet i Acabar Albarà"
      Height          =   465
      Left            =   2925
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5190
      Width           =   1530
   End
   Begin VB.CommandButton tancar 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6060
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton crearpalet 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Crear Palet i continuar"
      Height          =   465
      Left            =   1335
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5190
      Width           =   1530
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dades Albarà de Recepció"
      Height          =   4680
      Left            =   255
      TabIndex        =   0
      Top             =   480
      Width           =   6030
      Begin VB.TextBox impostbaseimp 
         BackColor       =   &H00989FF8&
         Height          =   315
         Left            =   3645
         TabIndex        =   3
         Top             =   2280
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox kgimpostenv 
         BackColor       =   &H00989FF8&
         Height          =   315
         Left            =   5250
         TabIndex        =   4
         Top             =   2310
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton imprimiralbaraproveidor 
         Height          =   330
         Left            =   3870
         Picture         =   "comprespalets.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Imprimir el nostra albarà de  proveïdor."
         Top             =   3780
         Width           =   345
      End
      Begin VB.TextBox preucompra 
         Height          =   315
         Left            =   5100
         TabIndex        =   5
         Top             =   1905
         Width           =   690
      End
      Begin VB.TextBox dataalbprov 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   3540
         Width           =   1140
      End
      Begin VB.TextBox qdepalets 
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   11
         ToolTipText     =   "Quants palets ha portat el proveïdor"
         Top             =   4140
         Width           =   645
      End
      Begin VB.TextBox numalbprov 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   10
         Top             =   3825
         Width           =   1830
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ok"
         Height          =   405
         Left            =   2805
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   180
         Width           =   600
      End
      Begin VB.ComboBox combocomandescompra 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         TabIndex        =   20
         Top             =   1020
         Width           =   4830
      End
      Begin VB.Frame Frame2 
         Caption         =   "Info de la Comanda"
         Height          =   840
         Left            =   135
         TabIndex        =   18
         Top             =   1410
         Width           =   4800
         Begin VB.Label infocomanda 
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   120
            TabIndex        =   19
            Top             =   270
            Width           =   4635
         End
      End
      Begin VB.TextBox numpaletproveidor 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   3240
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.TextBox lotproveidor 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   7
         Top             =   2925
         Width           =   2085
      End
      Begin VB.TextBox datarecepcio 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   2625
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox kgentregats 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   2295
         Width           =   840
      End
      Begin VB.TextBox comandacompra 
         Height          =   285
         Left            =   1275
         TabIndex        =   1
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label etcq 
         BackStyle       =   0  'Transparent
         Caption         =   "Calitat concertada."
         Height          =   165
         Left            =   3840
         TabIndex        =   37
         Top             =   2955
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Image bpdfcq 
         Height          =   750
         Left            =   3885
         OLEDropMode     =   1  'Manual
         Picture         =   "comprespalets.frx":0B14
         Top             =   2655
         Width           =   480
      End
      Begin VB.Label eimpbaseimp 
         BackStyle       =   0  'Transparent
         Caption         =   "Imp. Base Imp."
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2595
         TabIndex        =   35
         Top             =   2340
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label ettipusproveidorIMPOST 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3480
         TabIndex        =   34
         Top             =   225
         Width           =   2430
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00D29F7D&
         BackStyle       =   0  'Transparent
         Caption         =   "Kg Imp. Env:"
         DataField       =   "preucompra"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   4410
         TabIndex        =   33
         Top             =   2355
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00D29F7D&
         BackStyle       =   0  'Transparent
         Caption         =   "Preu o  €/Kg:"
         DataField       =   "preucompra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   4995
         TabIndex        =   28
         Top             =   1665
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quants palets ha portat el proveïdor."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   2685
         TabIndex        =   27
         Top             =   4200
         Width           =   2310
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Alb.Prov:"
         Height          =   240
         Left            =   210
         TabIndex        =   26
         Top             =   3555
         Width           =   1230
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Q. de Palets:"
         Height          =   300
         Left            =   210
         TabIndex        =   25
         Top             =   4170
         Width           =   1020
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Albarà Prov:"
         Height          =   240
         Left            =   210
         TabIndex        =   24
         Top             =   3870
         Width           =   1230
      End
      Begin VB.Label desccomanda 
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   135
         TabIndex        =   23
         Top             =   570
         Width           =   5760
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Materials comprats:"
         Height          =   390
         Left            =   195
         TabIndex        =   21
         Top             =   825
         Width           =   2700
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Palet Prov:"
         Height          =   240
         Left            =   210
         TabIndex        =   17
         Top             =   3285
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Lot Prov:"
         Height          =   240
         Left            =   210
         TabIndex        =   16
         Top             =   2955
         Width           =   1230
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Recepcio:"
         Height          =   240
         Left            =   210
         TabIndex        =   15
         Top             =   2640
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Un. o Kg Entregats:"
         Height          =   300
         Left            =   210
         TabIndex        =   14
         Top             =   2325
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nº Comanda: "
         Height          =   300
         Left            =   165
         TabIndex        =   13
         Top             =   270
         Width           =   1065
      End
      Begin VB.Image imgSIcq 
         Height          =   720
         Left            =   4335
         OLEDropMode     =   1  'Manual
         Picture         =   "comprespalets.frx":2526
         Top             =   2775
         Width           =   720
      End
      Begin VB.Image imgNOcq 
         Height          =   720
         Left            =   4335
         OLEDropMode     =   1  'Manual
         Picture         =   "comprespalets.frx":2D5F
         Top             =   2760
         Width           =   720
      End
   End
   Begin VB.Label etstatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   270
      Left            =   30
      TabIndex        =   36
      Top             =   5715
      Width           =   6345
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Compres - Recepció de Materials."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   45
      TabIndex        =   29
      Top             =   75
      Width           =   4305
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   375
      Left            =   -30
      Top             =   -15
      Width           =   6465
   End
End
Attribute VB_Name = "comprespalets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstcompra As Recordset
Dim dbqualitat As Database

Private Sub bpdfcq_DblClick()
  obrir_document bpdfcq.Tag
End Sub

Private Sub bpdfcq_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   guardar_fitxer_tmpCQ data.Files(1)
End Sub

Private Sub comandacompra_LostFocus()
  'Dim numerocompra As String
  'Dim rstreserva As Recordset
  'Dim desc As String
  'If cadbl(comandacompra) = 0 Then Exit Sub
  'If cadbl(comandacompra) > 10000 Then
  '   numerocompra = atrim(comandacompra)
  '  Else: numerocompra = Format(Now, "yy") + "000000" + Format(comandacompra, "0000"): comandacompra = Format(comandacompra, "0000")
  'End If
  'comandacompra.Tag = numerocompra
  'Set rstcompra = dbtmp.OpenRecordset("select * from compresmaterial where not entregada and numcompra='" + numerocompra + "'")
  'combocomandescompra.Clear
  'While Not rstcompra.EOF
  '  Set rstreserva = dbtmp.OpenRecordset("select * from reserves where idreserva=" + atrim(cadbl(rstcompra!idreserva)))
  '  If Not rstreserva.EOF Then desc = "  Ample: " + atrim(rstreserva!ample) + "  " + "Esp: " + atrim(rstreserva!espesor)
  '  combocomandescompra.AddItem atrim(rstcompra!codimat) + "-" + atrim(rstcompra!descmat) + desc
  '  combocomandescompra.ItemData(combocomandescompra.NewIndex) = atrim(rstcompra!idcomandacompra)
  '
  '  rstcompra.MoveNext
  'Wend
  
End Sub

Private Sub combocomandacompres_Click()
  
End Sub

Private Sub combocomandescompres_Change()

End Sub

Private Sub combocomandescompra_Click()
  Dim unitat As String
 If combocomandescompra.ListIndex = -1 Then Exit Sub
  rstcompres.FindFirst "idliniacompra=" + atrim(combocomandescompra.ItemData(combocomandescompra.ListIndex))
  If Not rstcompres.NoMatch Then
     netejarcamps
     'If rstcompra!entregada Then MsgBox "Aquesta comanda ja està entregada": Set rstcompra = Nothing: comandacompra.Tag = "": Exit Sub
     If rstcompres!totentregat Then MsgBox "Aquesta compra ja està entregada.", vbCritical, "OjO Entregat"
     If atrim(rstcompres!tipusmaterialcomprat) = "" Then rstcompres.Edit: rstcompres!tipusmaterialcomprat = "M": rstcompres.Update
     infocomanda = atrim(rstcompres!codimaterial) + "-" + atrim(rstcompres!nommaterial) + Chr(10) + Chr(13)
     unitat = IIf(atrim(rstcompres!tipusmaterialcomprat) <> "V", "Kg", "Un")
     infocomanda = infocomanda + " " + unitat + " Demanats:" + atrim(rstcompres!quantitatkg) + " " + unitat + " Pendents: " + atrim(rstcompres!quantitatkg - cadbl(rstcompres!kgentregats))
     'If cadbl(rstcompra!kgpendents) > 0 Then infocomanda = infocomanda + "/" + atrim(rstcompra!kgpendents) + " Kg Pndts"
     'kgentregats = atrim(rstcompres!quantitatkg - cadbl(rstcompres!kgentregats))
     kgentregats = 0
     datarecepcio = "" 'Format(Now, "dd/mm/yy")
     preucompra = atrim(cadbl(rstcompres!preu))
     mirartipusmaterialcompratiensenyarcampscorrectes atrim(rstcompres!tipusmaterialcomprat), cadbl(rstcompres!codimaterial)
     kgentregats.SetFocus
  End If
End Sub
Sub mirartipusmaterialcompratiensenyarcampscorrectes(tipusmat As String, vcodimat As Double)
  Dim visibles As Boolean
  Dim rst As Recordset
  Dim rstlinies As Recordset
  Dim vImpEnv As Double
  Set rst = dbtmp.OpenRecordset("select tanpercentimpostenvasos from materials where codi=" + atrim(vcodimat), , ReadOnly)
  If rst.EOF Then GoTo salt
  vImpEnv = cadbl(rst!tanpercentimpostenvasos)
  preucompra.Visible = False
  lblLabels(17).Visible = False
  If tipusmat = "M" Then
    visibles = True
    crearpalet.Caption = "Crear Palet i continuar"
    Command2.Caption = "Crear Palet i Acabar Albarà"
  End If
  If tipusmat <> "M" Then
    visibles = False
    crearpalet.Caption = "Registrar i continuar"
    Command2.Caption = "Registrar i Acabar Albarà"
    preucompra.Visible = True
    lblLabels(17).Visible = True
    If vImpEnv > 0 Then
         visibles = True
         'Set rstlinies = dbcompres.OpenRecordset("select numcomanda from comandesxlinia where idliniacompra=" + atrim(combocomandescompra.ItemData(combocomandescompra.ListIndex)))
         'If Not rstlinies.EOF Then lotproveidor = atrim(rstlinies!numcomanda)
            'LA OANA DIU QUE TREGUEM AIXÓ QUE NO LI ES COMODE, 19/06/2023
         Set rstlinies = Nothing
    End If
  End If
salt:
  Label7.Visible = visibles
  qdepalets.Visible = visibles
  Label10.Visible = visibles
  lblLabels(0).Visible = False
  kgimpostenv.Visible = False
  eimpbaseimp.Visible = False
  impostbaseimp.Visible = False
  If vImpEnv > 0 Then 'And tipusmat = "M"
    lblLabels(0).Visible = True
    kgimpostenv.Visible = True
    eimpbaseimp.Visible = True
    impostbaseimp.Visible = True
    kgimpostenv.Tag = atrim(vImpEnv)
  End If
  
End Sub
Sub netejarcamps()
  numpaletproveidor = ""
  kgentregats = ""
  preucompra = ""
  datarecepcio = ""
  lotproveidor = ""
  numpaletproveidor = ""
  dataalbprov = ""
  numalbprov = ""
  qdepalets = ""
  kgimpostenv = ""
  kgimpostenv.Tag = ""
  impostbaseimp = ""
End Sub
Sub crearpaletnou()
Dim rstreserva As Recordset
 Dim rstmaterial As Recordset
 Dim entregaacavada As Boolean
 If datarecepcio = "" Or lotproveidor = "" Or numpaletproveidor = "" Or quantitatdepalets = "" Then MsgBox "Falten camps per emplenar.": Exit Sub
 If rstcompra.EOF Then MsgBox "No hi ha cap relacio de compra amb aquest Nº de comanda.": Exit Sub
 Set rstreserva = dbtmp.OpenRecordset("select * from reserves where idreserva=" + atrim(cadbl(rstcompra!idreserva)))
 Set rstmaterial = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstcompra!codimat)))
 If rstreserva.EOF Then MsgBox "No he localitzat la reserva corresponents.": Exit Sub
 If rstmaterial.EOF Then MsgBox "No he localitzat el material que fa referencia la reserva.": Exit Sub
 Form1.alta_Click
 With Form1.palets.Recordset
  If .EditMode = 0 Then MsgBox "Error al donar d'alta el palet.": Exit Sub
  !ample = rstreserva!ample
  !plegat = rstreserva!plegat
  !solapa = rstreserva!solapa
  !carestractat = rstreserva!carestractat
  !obert = rstreserva!obert
  !microperforat = rstreserva!microperforat
  !semielaborat = rstreserva!semielaborat
  !micres = rstreserva!espesor
  !codimatprognou = rstreserva!codimat
  !numpaletpro = numpaletproveidor
  !numlot = lotproveidor
  !datarec = datarecepcio
  !numpalet = quantitatdepalets
  .Update
   .Bookmark = .LastModified
 End With
End Sub
Sub crear_unabobina(mtrs As Double)
  Form1.DBGrid1.AllowAddNew = True
  If Form1.DBGrid1.AllowAddNew Then
  Form1.DBGrid1.row = bobines.Recordset.RecordCount
  Form1.DBGrid1.Columns("mts") = mtrs
  If Form1.DBGrid1.EditActive Then Form1.bobines.Recordset.Edit: bobines.Recordset.Update
 End If
End Sub

Private Sub Command1_Click()
  desccomanda = ""
  Set rstcompres = dbcompres.OpenRecordset("SELECT capcalera.numcomanda,capcalera.materialrebut, capcalera.data, capcalera.dataentrega, capcalera.codiproveidorcomercial,capcalera.nomprovcomercial,capcalera.empresa,capcalera.nomprov,capcalera.codiproveidor, liniescompra.* FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra where numcomanda=" + atrim(cadbl(comandacompra)) + ";")
  If Not rstcompres.EOF Then
    desccomanda = atrim(rstcompres!nomprov) + " # Data Ent: " + atrim(rstcompres!dataentrega)
    If cadbl(codiproveidor(rstcompres!codiproveidorcomercial)) = 0 Then MsgBox "Aquest proveidor no te proveidor comercial assignat.", vbCritical, "Error": Exit Sub
    If rstcompres!materialrebut Then MsgBox "Atenció aquesta comanda està tota entregada", vbCritical, "Atenció"
  End If
  combocomandescompra.Clear
  While Not rstcompres.EOF
     combocomandescompra.AddItem atrim(rstcompres!codimaterial) + " - " + atrim(rstcompres!nommaterial) + IIf(atrim(rstcompres!tipusmaterialcomprat) = "M", " Ample:" + atrim(rstcompres!ample) + " Esp: " + atrim(rstcompres!micres) + IIf(rstcompres!totentregat, " ENTREGAT", ""), "")
     combocomandescompra.ItemData(combocomandescompra.NewIndex) = cadbl(rstcompres!idliniacompra)
     rstcompres.MoveNext
  Wend
  etcq.Visible = False
  bpdfcq.Visible = False: bpdfcq.Tag = "": bpdfcq.Enabled = False
  imgNOcq.Visible = False: imgSIcq.Visible = False
  combocomandescompra.SetFocus
  SendKeys "%{DOWN}"
End Sub

Function comprovarvalorsentrats() As Boolean
  Dim desc As String
   If Me.numalbprov = "" Then desc = "Falta Nº d'albarà" + Chr(10) + Chr(13)
   If Me.lotproveidor = "" Then desc = desc + "Falta el Lot de proveïdor" + Chr(10) + Chr(13)
   If Not IsDate(Me.dataalbprov) Then desc = desc + "La data d'albarà no es correcte." + Chr(10) + Chr(13)
  ' If Not IsDate(Me.datarecepcio) Then desc = desc + "La data de recepció no es correcte." + Chr(10) + Chr(13)
   If cadbl(Me.qdepalets) = 0 And rstcompres!tipusmaterialcomprat = "M" Then desc = desc + "La quantitat de palets no es correcte" + Chr(10) + Chr(13)
   If desc <> "" Then MsgBox desc
   comprovarvalorsentrats = IIf(desc <> "", False, True)
End Function

Sub comprovasitotentregat()
  Dim rstc As Recordset
  Set rstc = dbcompres.OpenRecordset("SELECT capcalera.numcomanda,capcalera.materialrebut, capcalera.data, capcalera.dataentrega, capcalera.nomprov, liniescompra.* FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra where not totentregat and numcomanda=" + atrim(cadbl(comandacompra)) + ";")
  If rstc.EOF Then
     dbcompres.Execute "update capcalera set materialrebut=true where numcomanda=" + atrim(cadbl(comandacompra))
  End If
  
End Sub

Private Sub Command2_Click()
   Dim vpaletcreat As Boolean
   Dim vnomproveidor As String
   Dim vcodiproveidor As String
   If imgNOcq.Visible Then If MsgBox("No has assignat el PDF del CQ, vols continuar igualment?", vbDefaultButton2 + vbCritical + vbYesNo, "ATENCIÓ") = vbNo Then Exit Sub
   If cadbl(kgentregats) < 1 Then MsgBox "No hi ha kilos entregats.", vbCritical, "Atenció": Exit Sub
   etstatus = "Creant el palet...": DoEvents
   crearelnoupalet vpaletcreat
   DoEvents
   etstatus = "Comprovant si tot entregat.": DoEvents
   comprovasitotentregat
   vcodiproveidor = codiproveidor(rstcompres!codiproveidorcomercial, vnomproveidor)
   guardar_CQ_sicorrespon vcodiproveidor, vnomproveidor, atrim(lotproveidor)
   If Not vpaletcreat Then GoTo fi
   etstatus = "Imprimint l'albarà."
   Me.Caption = "Imprimint l'albarà": DoEvents
   wait 1
   imprimiralbaraproveidor_Click
   DoEvents
   comprespalets.Hide
   Form1.SetFocus
fi:
   Me.Caption = "Compres - Recepció de Materials.": DoEvents
   etstatus = ""
End Sub

Private Sub crearpalet_Click()
 Dim vpaletcreat As Boolean
 Dim vcodiproveidor As String
 Dim vnomproveidor As String
 If imgNOcq.Visible Then If MsgBox("No has assignat el PDF del CQ, vols continuar igualment?", vbDefaultButton2 + vbCritical + vbYesNo, "ATENCIÓ") = vbNo Then Exit Sub
 If cadbl(kgentregats) < 1 Then MsgBox "No hi ha kilos entregats.", vbCritical, "Atenció": Exit Sub
 crearelnoupalet vpaletcreat
 comprovasitotentregat
 vcodiproveidor = codiproveidor(rstcompres!codiproveidorcomercial, vnomproveidor)
 guardar_CQ_sicorrespon atrim(vcodiproveidor), vnomproveidor, atrim(lotproveidor)
 If Not vpaletcreat Then GoTo fi
 comprespalets.Hide
 Form1.SetFocus
fi:
End Sub
Sub guardar_CQ_sicorrespon(vcodiproveidor As String, vnomproveidor As String, vlotproveidor As String)
  Dim vnomfitxerfinal As String
  Dim vfitxer As String
  If vlotproveidor = "" Or vcodiproveidor = "" Or vnomproveidor = "" Then Exit Sub
  vfitxer = bpdfcq.Tag
  If vfitxer = "" Then Exit Sub
  If Mid(UCase(vfitxer), 1, 15) <> "C:\TEMP\TMP_CQ_" Then Exit Sub
  vnomfitxerfinal = "CQ_" + atrim(vlotproveidor) + " [" + atrim(vcodiproveidor) + "]-" + atrim(vnomproveidor) + ".pdf"
  vnomfitxerfinal = treuresimbolsnovalidsnomfitxer(vnomfitxerfinal)
'  Clipboard.Clear
'  Clipboard.SetText vfitxer
  If existeix(vfitxer) Then
     FileCopy vfitxer, rutadelfitxer(cami) + "Cache_escanejarexpedicions\CQ\" + vnomfitxerfinal
  End If
  If existeix(rutadelfitxer(cami) + "Cache_escanejarexpedicions\CQ\" + vnomfitxerfinal) Then
      dbtmp.Execute "update albaransbip set lotescanejat=true where numlotproveidor='" + atrim(vlotproveidor) + "' and codiproveidorcomercial=" + atrim(vcodiproveidor)
  End If
              
End Sub
Function treuresimbolsnovalidsnomfitxer(desc As String) As String
   desc = substituir(desc, "\", "_")
   desc = substituir(desc, "/", "_")
   desc = substituir(desc, "|", "_")
   desc = substituir(desc, ":", ";")
   desc = substituir(desc, "?", "¿")
   desc = substituir(desc, "*", "x")
   desc = substituir(desc, """", "'")
   desc = substituir(desc, ">", "+")
   desc = substituir(desc, "<", "-")
   treuresimbolsnovalidsnomfitxer = desc
End Function
Sub demanarelmaterialdelcontenidor(vidmaterialcontenidor As Long, vmaterialcontenidor As String, vidproveidorrecuperador As Long, vmatriculacontenidor As String)
  Dim vidliniacompra As Long
  Dim rstref As Recordset
  vidliniacompra = cadbl(combocomandescompra.ItemData(combocomandescompra.ListIndex))
  Set rstref = dbcompres.OpenRecordset("select * from liniesdescripcio where idliniacompra=" + atrim(vidliniacompra) + " and descripcio like 'Ref:*'")
  If Not rstref.EOF Then
      vreferencia = substituir(atrim(rstref!descripcio), "Ref: ", " ")
      Set rstref = dbtintes.OpenRecordset("SELECT tintesreferencies.referencia, tipusbidons.nombido, tipusbidons.capacitat, tipusbidons.litrescompres FROM tintesreferencies LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id where tintesreferencies.referencia='" + atrim(vreferencia) + "'")
      If Not rstref.EOF Then
        If rstref.capacitat > 300 Then
            escullir_material_contenidor vidmaterialcontenidor, vmaterialcontenidor
            escullir_proveidorrecuperador vidproveidorrecuperador
            vmatriculacontenidor = InputBox("Escriu la MATRICULA d'aquest contenidor", "Matricula")
            If Len(vmatriculacontenidor) > 50 Then vmatriculacontenidor = Mid(vmatriculacontenidor, 1, 50)
            If vidmaterialcontenidor = 0 Then vidmaterialcontenidor = -1
            If vidproveidorrecuperador = 0 Then vidproveidorrecuperador = -1
        End If
      End If
  End If
  Set rstref = Nothing
End Sub
Sub escullir_proveidorrecuperador(vid As Long)
  Load formseleccio
  formseleccio.sortirs.Tag = "filtre"
  Set formseleccio.Data1.Recordset = dbcomandes.OpenRecordset("select id,nomcomercial from recuperadorsdecontenidors order by nomcomercial")
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Width = 2000
  formseleccio.Width = 5000
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   vid = formseleccio.DBGrid2.Columns("id")
  End If
  Unload formseleccio
End Sub

Sub escullir_material_contenidor(vidmaterialcontenidor As Long, vmaterialcontenidor As String)
    Dim vtipusmaterial As String
      Load formseleccio
      formseleccio.Caption = "Escull l'albarà "
      formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "tintes.mdb"
      formseleccio.Data1.RecordSource = "select codi,descripcio from contenidors_material order by descripcio"
      formseleccio.refrescar
      formseleccio.DBGrid2.Columns(1).Width = 5800
      formseleccio.DBGrid2.Columns(0).Visible = False
      formseleccio.Width = 6800
      'formseleccio.DBGrid2.Columns(0).Width = 1000
      'formseleccio.DBGrid2.Columns(1).Width = 3000
      formseleccio.Command2.Tag = "0"
      formseleccio.Caption = "Escullir tipus de contenidor"
      formseleccio.Show 1
      If seleccioret = 1 Then
           vmaterialcontenidor = atrim(formseleccio.Data1.Recordset!descripcio)
           vidmaterialcontenidor = cadbl(formseleccio.Data1.Recordset!codi)
      End If
      Unload formseleccio
     
End Sub
Sub crearelnoupalet(Optional vpaletcreat As Boolean)
 Dim rstreserva As Recordset
 Dim rstmaterial As Recordset
 Dim entregaacavada As Boolean
 Dim numiddepalet As Double
 Dim vidproveidorrecuperador As Long
 Dim vidmaterialcontenidor As Long
 Dim vmaterialcontenidor As String
 Dim vmatriculacontenidor As String
 vpaletcreat = True
 If cadbl(kgentregats) = 0 Then MsgBox "No hi ha els kilos entregats": kgentregats.SetFocus: Exit Sub
 If Not comprovarvalorsentrats Then Exit Sub
 Set dbstocks = Form1.palets.Database
 If rstcompres!tipusmaterialcomprat = "T" Then
      demanarelmaterialdelcontenidor vidmaterialcontenidor, vmaterialcontenidor, vidproveidorrecuperador, vmatriculacontenidor
      If vidmaterialcontenidor = -1 Or vidproveidorrecuperador = -1 Then MsgBox "Si no escullis contenidor o proveidor/recuperador no pots continuar", vbCritical, "Atenció": vpaletcreat = False: Exit Sub
 End If
 If cadbl(kgentregats) > 0 Then
    rstcompres.Edit
    rstcompres!kgentregats = cadbl(rstcompres!kgentregats) + cadbl(kgentregats)
    rstcompres.Update
    comprovasitotentregat
      
 End If
 entregaacavada = IIf(MsgBox("Ja has rebut tot el material comprat d'aquest producte de la comanda?", vbYesNo, "Atenció") = vbYes, True, False)
 Me.Caption = "Posant la quantitat entregada a la comanda de compra.": DoEvents
 If rstcompres!tipusmaterialcomprat <> "M" Then
      rstcompres.Edit: rstcompres!totentregat = entregaacavada: rstcompres.Update
      Comprovar_Comandes_Adhesius_muntadora cadbl(comandacompra)
 End If
 Me.Caption = "Posant la comanda de compra a entregada total": DoEvents
 If entregaacavada Then passarliniadecompraaentregada rstcompres!idliniacompra, cadbl(comandacompra)
'   reservarlescomandesassociades rstcompres   ' fins ara es reservaven aqui pero ara ho farem quan arrivi el material realment al entrar la data de recepcio
 numiddepalet = 0
 If rstcompres!tipusmaterialcomprat = "M" Then
    Me.Caption = "Creant el palet": DoEvents
    crearelpalet
    numiddepalet = cadbl(Form1.palets.Recordset!idpalet)
    Me.Caption = "Afegint la bobina": DoEvents
    Form1.editarpalet
    Form1.bobines.Recordset.AddNew
    Form1.bobines.Recordset!idbobina = Form1.labobinamesgran + 1
    Form1.DBGrid1.Columns("Nº Palet Prov.") = numpaletproveidor
 End If
 Me.Caption = "Gravant l'albarà de compra (Per SAP)": DoEvents
 etstatus = "Gravant l'albarà Bip": DoEvents
 gravar_albaracompra_bip numiddepalet, vidmaterialcontenidor, vidproveidorrecuperador, vmatriculacontenidor
 Me.Caption = "Gravant l'albarà de compra (Per SAP) FET": DoEvents
 DoEvents
 
End Sub
Sub Comprovar_Comandes_Adhesius_muntadora(vnumcomanda As Double)
    Dim rst As Recordset
    Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
    Set rst = dbbaixes.OpenRecordset("select * from estoccintaadhesiva")
    If cadbl(rst!comanda1) = vnumcomanda Then
         If MsgBox("Vols desvincular aquesta comanda d'adhesius " + atrim(rst!nomadhesiu1) + " de muntadora?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
            rst.Edit: rst!comanda1 = 0: rst.Update
         End If
    End If
    If cadbl(rst!comanda2) = vnumcomanda Then
         If MsgBox("Vols desvincular aquesta comanda d'adhesius " + atrim(rst!nomadhesiu2) + " de muntadora?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
            rst.Edit: rst!comanda2 = 0: rst.Update
         End If
    End If
    If cadbl(rst!comanda3) = vnumcomanda Then
         If MsgBox("Vols desvincular aquesta comanda d'adhesius " + atrim(rst!nomadhesiu3) + " de muntadora?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
            rst.Edit: rst!comanda3 = 0: rst.Update
         End If
    End If
    If cadbl(rst!comanda4) = vnumcomanda Then
         If MsgBox("Vols desvincular aquesta comanda d'adhesius " + atrim(rst!nomadhesiu4) + " de muntadora?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
            rst.Edit: rst!comanda4 = 0: rst.Update
         End If
    End If
    If cadbl(rst!comanda5) = vnumcomanda Then
         If MsgBox("Vols desvincular aquesta comanda d'adhesius " + atrim(rst!nomadhesiu5) + " de muntadora?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
            rst.Edit: rst!comanda5 = 0: rst.Update
         End If
    End If
    Set rst = Nothing
End Sub
Sub passarliniadecompraaentregada(idlinia As Double, numc As Double)
   Dim rstm As Recordset
   Dim rstt As Recordset
   dbcompres.Execute "update liniescompra set totenviat=true where idliniacompra=" + atrim(idlinia)
   
End Sub
Function buscarproximnumalb(numalbp As String, vcodiprovcomercial As Double) As Double
   Dim rsta As Recordset
   Me.Caption = "Gravant l'albarà de compra (Per SAP) Buscant numero albarà nou (numalbaraprov)": DoEvents
'   Set rsta = dbcompres.OpenRecordset("Select * from albaransbip where numalbaraprov='" + atrim(numalbp) + "'", , ReadOnly)
   Set rsta = dbcompres.OpenRecordset("Select * from albaransbip", , ReadOnly)
   rsta.FindFirst "numalbaraprov='" + atrim(numalbp) + "' and year(data)=year(now) and codiproveidorcomercial=" + atrim(vcodiprovcomercial)
   If Not rsta.NoMatch Then
      buscarproximnumalb = rsta!numalbara
     Else
        buscarproximnumalb = 0
        
        Me.Caption = "Gravant l'albarà de compra (Per SAP) Buscant numero albarà (nUmalbara) ": DoEvents
        Set rsta = dbcompres.OpenRecordset("Select numalbara as gran from albaransbip order by numalbara desc", , ReadOnly)
        If Not rsta.EOF Then buscarproximnumalb = rsta!gran
        buscarproximnumalb = buscarproximnumalb + 1
   End If
   Set rsta = Nothing
End Function
Sub gravar_albaracompra_bip(numpalet As Double, vidmaterialcontenidor As Long, vidproveidorrecuperador As Long, vmatriculacontenidor As String)
  Dim rstalb As Recordset
  Dim ruta As String
  Dim numalb As Double
  Dim vcodiprov As Double
  Me.Caption = "Gravant l'albarà de compra(Per SAP)Buscant numero albarà nou": DoEvents
  numalb = buscarproximnumalb(numalbprov, codiproveidor(rstcompres!codiproveidorcomercial))
  Me.Caption = "Gravant l'albarà(Per SAP)Numero trobat": DoEvents
  If Not rstcompres.EOF Then
     Me.Caption = "Gravant l'albarà(Per SAP)Afengint albarà SAP": DoEvents
     Set rstalb = dbcompres.OpenRecordset("albaransbip")
     rstalb.AddNew
     rstalb!numalbara = numalb
     rstalb!numcomanda = rstcompres!numcomanda
     rstalb!numalbaraprov = numalbprov
     rstalb!empresa = rstcompres!empresa
     rstalb!data = CVDate(dataalbprov)
     Me.Caption = "Gravant l'albarà de compra (Per SAP) Buscant codi de proveidor": DoEvents
     rstalb!codiproveidorcomercial = codiproveidor(rstcompres!codiproveidorcomercial)
     rstalb!nomproveidorcomercial = atrim(rstcompres!nomprovcomercial)
     vcodiprov = rstalb!codiproveidorcomercial
     rstalb!numlotproveidor = atrim(lotproveidor)
     rstalb!vidmaterialcontenidor = cadbl(vidmaterialcontenidor)
     rstalb!idproveidorrecuperador = cadbl(vidproveidorrecuperador)
     rstalb!vmatriculacontenidor = atrim(vmatriculacontenidor)
     rstalb!article = rstcompres!codimaterial
     Me.Caption = "Gravant l'albarà de compra (Per SAP)  Possant la descripcio de la compra": DoEvents
     If Len(combocomandescompra) > 3 Then
        rstalb!descripcio = substituir(Mid(combocomandescompra, InStr(1, combocomandescompra, "-") + 1), "ENTREGAT", "")
     End If
     rstalb!quantitat = cadbl(kgentregats)
     rstalb!preu = cadbl(preucompra) ' cadbl(rstcompres!preu)
     rstalb!importnet = rstalb!quantitat * rstalb!preu
     rstalb!kgimpostenvasos = cadbl(kgimpostenv)
     rstalb!kgbaseimposableimpostenvasos = cadbl(impostbaseimp)
     rstalb!preuImpostEnvasos = cadbl(llegir_ini("General", "PreuImpostEnvasos", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"))
     Me.Caption = "Gravant l'albarà de compra (Per SAP) Generar el fitxer de SAP": DoEvents
     rstalb!nomfitxer = generarfitxeralbaratxt(rstalb)
     rstalb!numpalet = numpalet
     rstalb!idliniacompra = rstcompres!idliniacompra
     rstalb!enviat = False
     revisarsitenimlalbaradeproveidorielCQdellotdeproveidorescanejat numalbprov, lotproveidor, vcodiprov, rstalb
     If calCQ(numalbprov, vcodiprov, dbcompres) Then rstalb!cal_CQ_lot = True
     Me.Caption = "Gravant l'albarà de compra (Per SAP)  Gravant l'albarà": DoEvents
     rstalb.Update
     Me.Caption = "Gravant l'albarà de compra (Per SAP) Actualitzat el bookmark a ultim modificat.": DoEvents
     rstalb.Bookmark = rstalb.LastModified
       Else: Me.Caption = "Gravant l'albarà(Per SAP)Compra perduda": DoEvents
  End If
End Sub
Sub revisarsitenimlalbaradeproveidorielCQdellotdeproveidorescanejat(valbprov As String, vlot As String, vcodiprov As Double, rstalb As Recordset)
    Dim rst As Recordset
    Set rst = dbcompres.OpenRecordset("select * from registre_escanejades_expedicions where nomfitxer like '" + atrim(valbprov) + " [[]" + atrim(vcodiprov) + "*'")
    If Not rst.EOF Then rstalb!albaraescanejat = True
    Set rst = dbcompres.OpenRecordset("select * from registre_escanejades_expedicions where nomfitxer like 'CQ_" + atrim(vlot) + " [[]" + atrim(vcodiprov) + "*'")
    If Not rst.EOF Then rstalb!lotescanejat = True
    Set rst = Nothing
End Sub
Function codiproveidor(codi As Long, Optional vnomproveidor As String) As Double
   Dim rstcodi As Recordset
   Dim rstprov As Recordset
   codiproveidor = 0
   ettipusproveidorIMPOST = ""
   Set rstcodi = dbtmpb.OpenRecordset("select codicomptable,codiproduccio from proveidors_comercial where codi=" + atrim(codi), , ReadOnly)
   If Not rstcodi.EOF Then
      codiproveidor = cadbl(rstcodi!codicomptable)
      Set rstprov = dbtmpb.OpenRecordset("select nom, tipusproveidorIMPOST from proveidors where codi=" + atrim(cadbl(rstcodi!codiproduccio)))
      If Not rstprov.EOF Then
            ettipusproveidorIMPOST = "Tipus IMPOST " + atrim(rstprov!tipusproveidorIMPOST)
            vnomproveidor = atrim(rstprov!nom)
      End If
   End If
   Set rstcodi = Nothing
   Set rstprov = Nothing
End Function
Sub generar_capcalera_fitxer_sap(rstalb As Recordset)
    Dim vdata As Date
    With rstalb
    vdata = !data
    'la linia seguent es feia per passar la data a mes seguent si era just a finals de mes per esperar que el material arrives
    'el comptable diu que així no
    'If Month(vdata) <> Month(Now) Or Year(vdata) <> Year(Now) Then vdata = CVDate("01/" + atrim(Month(Now)) + "/" + atrim(Year(Now)))
    linia = atrim(!codiproveidorcomercial) + ";" + format(vdata, "dd/mm/yy") + ";" + atrim(!numalbaraprov)
    Print #1, linia
    End With
End Sub
Sub generar_linia_fitxer_sap(rstalb As Recordset, nomfitxer As String)  ', vrutasapseidor As String)
   Dim r As String
   Dim linia As String
   Dim ruta As String
   Dim rstidcompra As Recordset
   Dim espesor As Double
   Dim mesuraespesor As String
   Dim numlotfabricacio As String
   Dim comandesrelacionades As String
   Dim numalbaraprov As String
   On Error GoTo errorgravar
  ' ruta = llegir_ini("Compres", "rutasap", "comandes.ini")
  '  r = ruta + "\" + atrim(rstalb!nomfitxer)
    r = nomfitxer
    If Not existeix(r) Then
              Open r For Output As 1
              generar_capcalera_fitxer_sap rstalb
          Else: Open r For Append As 1
    End If
    With rstalb
    Set rstidcompra = dbcompres.OpenRecordset("select * from liniescompra where idliniacompra=" + atrim(!idliniacompra))
    If rstidcompra.EOF Then MsgBox "Error no trobo la compra relacionada amb aquest albarà." + vbNewLine + "EL FITXER DE TRASPAS NO ES CREARÀ CORRECTAMENT.", vbCritical, "Error": Exit Sub
    espesor = IIf(cadbl(rstidcompra!grmm2) = 0, cadbl(rstidcompra!micres), cadbl(rstidcompra!grmm2))
    mesuraespesor = "Micres"
    If cadbl(rstidcompra!grmm2) <> 0 Then mesuraespesor = "Grms/m2"
    numlotfabricacio = substituir(atrim(!numlotproveidor), ";", "_")
    comandesrelacionades = llistadecomandesrelacionades(!idliniacompra)
    numalbaraprov = substituir(atrim(!numalbaraprov), ";", "_")
    linia = substituir(atrim(!article), ";", "_") + ";" + treuresimbols(atrim(!descripcio)) + ";" + atrim(rstidcompra!semielaborat) + ";" + substituir(atrim(rstidcompra!ample), ",", ".") + ";" + substituir(atrim(espesor), ",", ".") + ";" + atrim(mesuraespesor) + ";" + atrim(numlotfabricacio) + ";" + atrim(numalbaraprov) + ";" + atrim(comandesrelacionades) + ";" + substituir(atrim(!quantitat), ",", ".") + ";" + substituir(atrim(!preu), ",", ".")
    '    MsgBox linia
    Print #1, linia
    ' SI HI HA IMPOST D'ENVASOS A LA COMPRA I ENS COBRARÀ IMPOST LA POSO EN UNA ALTRA LINIA PERQUE SURTI A FACTURA
    If cadbl(!kgimpostenvasos) > 0 And elproveidoresESPANYOL(rstalb!codiproveidorcomercial) Then
         linia = "IMP_ENV;Impuesto segun Ley 7/2022 de envases no reutilizables.;;;;;;" + atrim(numalbaraprov) + ";;" + substituir(atrim(!kgimpostenvasos), ",", ".") + ";" + substituir(atrim(!preuImpostEnvasos), ",", ".")
         Print #1, linia
    End If
    Close 1
    traspasdetotlarticleaSAPaunfitxerapartCSV cadbl(!article), atrim(!descripcio), atrim(rstidcompra!tipusmaterialcomprat), nomfitxer ', vrutasapseidor
    End With
    Exit Sub
errorgravar:
      MsgBox err.Description & Chr(10) & "No s'ha gravat el fitxer: " + Chr(10) + r
End Sub
Function elproveidoresESPANYOL(vcodicomptableproveidor As Double) As Boolean
   Dim rst As Recordset
   'On Error GoTo 0
   Set rst = dbcomandes.OpenRecordset("select codiproduccio from proveidors_comercial where codicomptable='" + atrim(vcodicomptableproveidor) + "'")
   If rst.EOF Then Exit Function
   Set rst = dbcomandes.OpenRecordset("select * from proveidors where codi=" + atrim(rst!codiproduccio))
   If Not rst.EOF Then
       If atrim(rst!tipusproveidorIMPOST) = "Espanyol" Then
            elproveidoresESPANYOL = True
       End If
   End If
   Set rst = Nothing
End Function
Sub traspasdetotlarticleaSAPaunfitxerapartCSV(codiarticle As Double, nomarticle As String, tipusmaterialcomprat As String, nomfitxer As String) ', vrutasapseidor As String)
    Dim linia As String
    Dim fitxersapseidor As String
    Dim fitxertraspascodis As String
    fitxertraspascodis = rutadelfitxer(nomfitxer) + "A-Articles.csv"
    '// fitxersapseidor = rutadelfitxer(vrutasapseidor) + "A-Articles.csv"
    If Not existeix(fitxertraspascodis) Then
              Open fitxertraspascodis For Output As 5
          Else:
             comprovarsijaexisteix codiarticle, fitxertraspascodis
             Open fitxertraspascodis For Append As 5
    End If
    linia = atrim(codiarticle) + ";" + treuresimbols(nomarticle)
    possarfamiliesarticle codiarticle, tipusmaterialcomprat, linia
    possarfamiliestinta codiarticle, tipusmaterialcomprat, linia
    Print #5, linia
    Close 5
'//    'copio el fitxer d'articles
'//    If existeix(fitxersapseidor) Then
'//       'Kill fitxersapseidor
'//        concatenar_fitxers fitxertraspascodis, fitxersapseidor
'//      Else: Copiar_Fitxer fitxertraspascodis, fitxersapseidor
'//    End If
End Sub
Sub comprovarsijaexisteix(codiarticle As Double, fitxertraspascodis As String)
    Dim vfitxertmp As String
    vfitxertmp = "c:\temp\A-Articles_tmp.csv"
    Open fitxertraspascodis For Input As 5
    Open vfitxertmp For Output As 6
    While Not EOF(5)
      Line Input #5, v
      If InStr(1, v, ";") > 0 Then
         If cadbl(Mid(v, 1, InStr(1, v, ";") - 1)) <> codiarticle Then
            Print #6, v
         End If
      End If
    Wend
    Close 6
    Close 5
    Kill fitxertraspascodis
    Copiar_Fitxer vfitxertmp, fitxertraspascodis
End Sub
Sub possarfamiliestinta(codiarticle As Double, tipusmaterial As String, linia As String)
    Dim rst As Recordset
    Dim vcamps As String
    Dim vfrom As String
    Dim vwhere As String
    
    vcamps = "SELECT tintes.referenciacolor, seriescolors.codi, seriescolors.descripcio, familiestintes.codi, familiestintes.descripcio, subfamiliestintes.codi, subfamiliestintes.descripcio, familiescolors.codi, familiescolors.descripcio, subfamiliescolors.codi, subfamiliescolors.descripcio "
    vfrom = " FROM ((((tintes LEFT JOIN seriescolors ON tintes.idserie = seriescolors.codi) LEFT JOIN familiestintes ON tintes.idfamilia = familiestintes.codi) LEFT JOIN familiescolors ON tintes.idfamcolor = familiescolors.codi) LEFT JOIN subfamiliestintes ON tintes.idsubfamilia = subfamiliestintes.codi) LEFT JOIN subfamiliescolors ON tintes.idsubfamcolor = subfamiliescolors.codi "
    vwhere = " where tintes.codi='" + atrim(codiarticle) + "'"
    Set rst = dbtintes.OpenRecordset(vcamps + vfrom + vwhere)
    If rst.EOF Or tipusmaterial <> "T" Then linia = linia + ";;;;;;;;;;;": GoTo fi
    With rst
    afegirv atrim(![referenciacolor]), linia
    afegirv atrim(cadbl(![seriescolors.codi])), linia
    afegirv atrim(![seriescolors.descripcio]), linia
    afegirv atrim(cadbl(![familiestintes.codi])), linia
    afegirv atrim(![familiestintes.descripcio]), linia
    afegirv atrim(cadbl(![subfamiliestintes.codi])), linia
    afegirv atrim(![subfamiliestintes.descripcio]), linia
    afegirv atrim(cadbl(![familiescolors.codi])), linia
    afegirv atrim(![familiescolors.descripcio]), linia
    afegirv atrim(cadbl(![subfamiliescolors.codi])), linia
    afegirv atrim(![subfamiliescolors.descripcio]), linia
    End With
fi:
    Set rst = Nothing
End Sub
Sub possarfamiliesarticle(codiarticle As Double, tipusmaterial As String, linia As String)
    Dim rst As Recordset
    Dim vcamps As String
    Dim vfrom As String
    Dim vwhere As String
    vcamps = "SELECT materials.codi, materials.descripcio, familiesmaterials.codi, familiesmaterials.descripcio, subfamiliesmaterials.codi, subfamiliesmaterials.descripcio, familiescolorants.codi, familiescolorants.descripcio, subfamiliescolorants.codi, subfamiliescolorants.descripcio, familiesaditius.codi, familiesaditius.descripcio, subfamiliesaditius.codi, subfamiliesaditius.descripcio "
    vfrom = " FROM (((((materials LEFT JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) LEFT JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi) LEFT JOIN familiescolorants ON materials.familiacol = familiescolorants.codi) LEFT JOIN subfamiliescolorants ON materials.subfamiliacol = subfamiliescolorants.codi) LEFT JOIN familiesaditius ON materials.familiaad = familiesaditius.codi) LEFT JOIN subfamiliesaditius ON materials.subfamiliaad = subfamiliesaditius.codi "
    vwhere = " where materials.codi=" + atrim(codiarticle)
    Set rst = dbcomandes.OpenRecordset(vcamps + vfrom + vwhere)
    If rst.EOF Or tipusmaterial = "T" Then linia = linia + ";;;;;;;;;;;;": GoTo fi
    With rst
    afegirv atrim(cadbl(![familiesmaterials.codi])), linia
    afegirv atrim(![familiesmaterials.descripcio]), linia
    afegirv atrim(cadbl(![subfamiliesmaterials.codi])), linia
    afegirv atrim(![subfamiliesmaterials.descripcio]), linia
    afegirv atrim(cadbl(![familiescolorants.codi])), linia
    afegirv atrim(![familiescolorants.descripcio]), linia
    afegirv atrim(cadbl(![subfamiliescolorants.codi])), linia
    afegirv atrim(![subfamiliescolorants.descripcio]), linia
    afegirv atrim(cadbl(![familiesaditius.codi])), linia
    afegirv atrim(![familiesaditius.descripcio]), linia
    afegirv atrim(cadbl(![subfamiliesaditius.codi])), linia
    afegirv atrim(![subfamiliesaditius.codi]), linia
    End With
fi:
    Set rst = Nothing
End Sub
Sub afegirv(v As String, linia As String)
    linia = linia + ";" + treuresimbols(v)
End Sub
Function llistadecomandesrelacionades(idliniacompra As Double) As String
  Dim v As String
  Dim rst As Recordset
  Set rst = dbcompres.OpenRecordset("select * from comandesxlinia where idliniacompra=" + atrim(idliniacompra))
  While Not rst.EOF
    v = v + " " + atrim(rst!comandavisual)
    rst.MoveNext
  Wend
  Set rst = Nothing
  If Len(v) > 50 Then v = Mid(v, 1, 50)
  llistadecomandesrelacionades = v
End Function
Sub generar_linia_fitxer_bip(rstalb As Recordset)
   Dim r As String
   Dim linia As String
   Dim ruta As String
   On Error GoTo errorgravar
   ruta = llegir_ini("Compres", "rutabip", "comandes.ini")
    r = ruta + "\empre" + format(cadbl(llegir_ini("Compres", "numempresabip_" + atrim(rstalb!empresa), "comandes.ini")), "000") + "\" + atrim(rstalb!nomfitxer)
    
    If Not existeix(r) Then
              Open r For Output As 1
          Else: Open r For Append As 1
    End If
    With rstalb
    linia = atrim(!article) + ":" + treuresimbols(atrim(!descripcio)) + "|" + substituir(atrim(cadbl(!quantitat)), ".", ",") + "|" + substituir(atrim(cadbl(!preu)), ".", ",") + "|0"
'    MsgBox linia
    Print #1, linia
    End With
    Close 1
    Exit Sub
errorgravar:
      MsgBox err.Description & Chr(10) & "No s'ha gravat el fitxer: " + Chr(10) + r
End Sub
Function treuresimbols(desc As String) As String
   desc = substituir(desc, ":", "_")
   desc = substituir(desc, "'", "´")
   desc = substituir(desc, "|", "_")
   desc = substituir(desc, ";", "_")
   treuresimbols = desc
End Function
Function substituir(cadena As String, buscar As String, canviar As String) As String
   cadena = " " + cadena
   comença = InStr(1, cadena, buscar) - 1
   If comença < 1 Then substituir = atrim(cadena): Exit Function
   acaba = comença + Len(buscar) + 1
   cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   substituir = atrim(cadena)
   'MsgBox linia
End Function
Sub reservarlescomandesassociades(rstcompres As Recordset)
  Dim rstl As Recordset
  
  Set rstl = dbcompres.OpenRecordset("select * from comandesxlinia where idliniacompra=" + atrim(cadbl(rstcompres!idliniacompra)))
  
  While Not rstl.EOF
    reservar rstcompres, rstl
    rstl.MoveNext
  Wend
End Sub
Function existeixreservaxcomocli(rstc As Recordset) As Boolean
   Dim rstr As Recordset
   Set dbstocks = Form1.palets.Database
   Set rstr = dbstocks.OpenRecordset("select * from percomandaoclient where numcomanda=" + atrim(cadbl(rstc!numcomanda)))
   If Not rstr.EOF Then
       existeixreservaxcomocli = True
         Else: existeixreservaxcomocli = False
   End If
   If cadbl(rstc!numcomanda) < 120000 Then existeixreservaxcomocli = True
   Set rstr = Nothing
End Function
Sub novareserva(rstc As Recordset, rstr As Recordset)
  rstr.AddNew
  rstr!ample = Redondejar(cadbl(rstc!ample), 1)
  rstr!plegat = cadbl(rstc!plegat)
  rstr!carestractat = atrim(rstc!carestractat)
  rstr!obert = atrim(rstc!obert)
  rstr!microperforat = rstc!microperforat
  rstr!semielaborat = rstc!semielaborat
  rstr!espesor = IIf(cadbl(rstc!grmm2) > 0, cadbl(rstc!grmm2) * -1, cadbl(rstc!micres))
  rstr!familia = rstc!familia
  rstr!subfamilia = rstc!subfamilia
  rstr!familiacol = rstc!familiacol
  rstr!subfamiliacol = rstc!subfamiliacol
  rstr!familiaad = rstc!familiaad
  rstr!subfamiliaad = rstc!subfamiliaad
  rstr.Update
  rstr.Bookmark = rstr.LastModified
End Sub
Function existeixreserva(rstc As Recordset, rstr As Recordset) As Boolean
  Dim r As String
  Dim r2 As String
      r = "ample=" + passaradecimalpunt(rstc!ample) + " and plegat=" + passaradecimalpunt(rstc!plegat)
      r = r + " and solapa=" + passaradecimalpunt(rstc!solapa) + " and carestractat='" + atrim(rstc!carestractat + "'")
      r = r + " and obert='" + atrim(rstc!obert) + "' and microperforat=" + IIf(cabool(rstc!microperforat), "True", "False")
      r = r + " and semielaborat='" + atrim(rstc!semielaborat) + "' and espesor=" + passaradecimalpunt(IIf(rstc!grmm2 > 0, rstc!grmm2 * -1, rstc!micres))
      r2 = "and familia=" + atrim(cadbl(rstc!familia)) + " and subfamilia=" + atrim(cadbl(rstc!subfamilia))
      r2 = r2 + " and familiacol=" + atrim(cadbl(rstc!familiacol)) + " and subfamiliacol=" + atrim(cadbl(rstc!subfamiliacol))
      r2 = r2 + " and familiaad=" + atrim(cadbl(rstc!familiaad)) + " and subfamiliaad=" + atrim(cadbl(rstc!subfamiliaad))
      
      Set rstr = dbtmp.OpenRecordset("select * from reserves where " + r + r2)
      If Not rstr.EOF Then
           existeixreserva = True
         Else: existeixreserva = False
      End If
      
      
      
End Function
Sub reservar(rstc As Recordset, rstl As Recordset)
   Dim rstr As Recordset
   Dim rstxc As Recordset
   Dim rstmc As Recordset
   Dim rstcomextra As Recordset
   Dim metrescomanda As Double
   metrescomanda = 0
   If existeixreservaxcomocli(rstl) Or existeixassignacio(rstl!numcomanda) Then GoTo fi
   Set rstmc = dbtmp.OpenRecordset("select cantitatex from comandes where comanda=" + atrim(cadbl(rstl!numcomanda)))
   Set rstcomextra = dbtmp.OpenRecordset("select materialexacte from comandes_extres where comanda=" + atrim(cadbl(rstl!numcomanda)))
   If Not rstcomextra.EOF Then
    If cabool(rstcomextra!materialexacte) = True Then
        GoTo fi
    End If
   End If
   If Not rstmc.EOF Then metrescomanda = rstmc!cantitatex
'   metrescomanda = Format(compramat.conversiokilos(rstc!codimaterial, cadbl(rstc!ample), cadbl(rstl!kgcompra) * -1, IIf(cadbl(rstc!micres), cadbl(rstc!micres), cadbl(rstc!grmm2) * -1), rstc!semielaborat, cadbl(rstc!solapa)), "#,##0")
   If Not existeixreserva(rstc, rstr) Then
       novareserva rstc, rstr
       
   End If
   rstr.Edit
   rstr!metresreservats = cadbl(rstr!metresreservats) + metrescomanda
   rstr.Update
   dbstocks.Execute "insert into percomandaoclient (idreserva,numcomanda,metres) values (" + atrim(rstr!idreserva) + "," + atrim(rstl!numcomanda) + "," + atrim(metrescomanda) + ")"
fi:
   Set rstmc = Nothing
   Set rstcomextra = Nothing
   Set rstxc = Nothing
   Set rstmc = Nothing
End Sub
Function existeixassignacio(numc As Double) As Boolean
   Dim rsta As Recordset
   existeixassignacio = False
   Set rsta = Form1.palets.Database.OpenRecordset("select * from parcials where cdbl(comanda)=" + atrim(numc))
   If Not rsta.EOF Then
       existeixassignacio = True
   End If
   Set rsta = dbtmp.OpenRecordset("select * from comandes_extres where comanda=" + atrim(numc))
   If Not rsta.EOF Then
      If rsta!assignarstock Then existeixassignacio = True
   End If
   Set rsta = Nothing
End Function
Function buscarmicresmaterial(vcodimat As Double) As Double
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select micresdelsgrm2 from materials where codi=" + atrim(vcodimat))
  If Not rst.EOF Then buscarmicresmaterial = cadbl(rst!micresdelsgrm2)
  Set rst = Nothing
End Function
Sub crearelpalet()
With Form1
  .noupalet
  .txtFields(1) = rstcompres!codimaterial
  .txtFields(2) = rstcompres!ample
  .txtFields(3) = rstcompres!plegat
  .txtFields(4) = rstcompres!solapa
  .txtFields(11) = IIf(cadbl(rstcompres!micres) <> 0, rstcompres!micres, buscarmicresmaterial(rstcompres!codimaterial))
  .Combo1 = atrim(rstcompres!semielaborat)
  .Combo2 = atrim(rstcompres!carestractat)
  If atrim(rstcompres!carestractat) <> "N" Then
      .tractat = "LAMINAR"
     Else
       .tractat = "NO"
  End If
  .Combo3 = rstcompres!obert
  .microp = IIf(rstcompres!microperforat, 1, 0)
  .preucompra = rstcompres!preu
  .txtFields(14) = .preucompra
  .txtFields(6) = Me.numalbprov
  .txtFields(7) = Me.lotproveidor
  .txtFields(8) = Me.dataalbprov
  .txtFields(9) = Me.datarecepcio
  .txtFields(5) = Me.qdepalets
  .txtFields(14) = cadbl(Me.preucompra)
  .preucompra = cadbl(Me.preucompra)
  .actualitzar_vinculats
  .gravar_canvis
 End With
End Sub

Private Sub datarecepcio_LostFocus()
   If LCase(datarecepcio) = "avui" Then datarecepcio = format(Now, "dd/mm/yy")
   If LCase(datarecepcio) = "ahir" Then datarecepcio = format(DateAdd("d", -1, Now), "dd/mm/yy")
   If LCase(datarecepcio) = "demà" Or LCase(datarecepcio) = "dema" Then datarecepcio = format(DateAdd("d", 1, Now), "dd/mm/yy")
End Sub

Private Sub Form_Click()

   
   'Set rsttmp = dbcompres.OpenRecordset("SELECT capcalera.numcomanda, liniescompra.idliniacompra, capcalera.materialrebut, liniescompra.totentregat FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE (((capcalera.materialrebut)=False));")
   ' While Not rsttmp.EOF
   '  If rsttmp!totentregat Then
   '   passarliniadecompraaentregada rsttmp!idliniacompra, cadbl(rsttmp!numcomanda)
   '  End If
   '  rsttmp.MoveNext
   ' Wend
   'demanarelmaterialdelcontenidor cadbl(vid), r, cadbl(vid2)
End Sub
Function calCQ(vlot As String, vcodiproveidor As Double, dbcompres As Database) As Boolean
   Dim rst As Recordset
   Dim rstmat As Recordset
   Dim rstprov As Recordset
   calCQ = True
   Set rstprov = dbcomandes.OpenRecordset("SELECT proveidors.tipusCQ, proveidors.dataCQ, proveidors.codi, proveidors_comercial.codicomptable FROM proveidors LEFT JOIN proveidors_comercial ON proveidors.codi = proveidors_comercial.codiproduccio where proveidors_comercial.codicomptable='" + atrim(vcodiproveidor) + "'")
   If rstprov.EOF Then Exit Function
   Set rst = dbcompres.OpenRecordset("SELECT albaransbip.*, liniescompra.tipusmaterialcomprat FROM albaransbip INNER JOIN liniescompra ON albaransbip.idliniacompra = liniescompra.idliniacompra where codiproveidorcomercial=" + atrim(vcodiproveidor) + " and numlotproveidor='" + atrim(vlot) + "'")
   If Not rst.EOF Then
      If rst!tipusmaterialcomprat <> "T" Then
       Set rstmat = dbcomandes.OpenRecordset("select * from materials where codi=" + atrim(rst!article))
       If rstmat.EOF Then Exit Function
       If rstmat!tipusCQ <> "L" Then calCQ = False
       If Not calCQ Then If atrim(rstprov!tipusCQ) = "L" And atrim(rstmat!tipusCQ) <> "N" Then calCQ = True
         Else
           If atrim(rstprov!tipusCQ) <> "L" Then calCQ = False
      End If
   End If
   Set rst = Nothing
   Set rstmat = Nothing
   Set rstprov = Nothing
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 And Screen.ActiveControl.Name <> "comandacompra" Then KeyCode = 0: SendKeys "{tab}"
End Sub

Private Sub Form_Load()
  'Label1 = "Nº Comanda: " + Format(Now, "yy") + "0000000"
  Set dbcompres = DBEngine.OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
  Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
  Set dbqualitat = OpenDatabase(rutadelfitxer(cami) + "qualitat.mdb")
  If existeix("C:\TEMP\TMP_CQ_*.*") Then Kill "C:\TEMP\TMP_CQ_*.*"
  imgNOcq.Visible = False: imgSIcq.Visible = False
  bpdfcq.Visible = False
End Sub

Private Sub Text1_Change()

End Sub

Sub borrartotselstemporals()
  On Error Resume Next
  Kill "c:\temp\~llac*.*"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set dbtintes = Nothing
End Sub

Private Sub imgNOcq_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   guardar_fitxer_tmpCQ data.Files(1)
End Sub
Sub guardar_fitxer_tmpCQ(vnomfitxer As String)
   Dim v As String
   imgNOcq.Visible = False: imgSIcq.Visible = True
   If existeix(vnomfitxer) Then
        If InStr(1, UCase(vnomfitxer), ".PDF") > 0 Then
            v = substituir(atrim(vnomfitxer), rutadelfitxer(vnomfitxer), "c:\temp\TMP_CQ_")
            FileCopy vnomfitxer, v
            bpdfcq.Tag = v
            bpdfcq.Enabled = True: imgNOcq.Visible = False: imgSIcq.Visible = True
        End If
   End If
End Sub
Private Sub imgSIcq_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    guardar_fitxer_tmpCQ data.Files(1)
End Sub

Private Sub imprimiralbaraproveidor_Click()
   impresiodalbara numalbprov
   'generar_fitxer_bip numalbprov
   'generar_fitxer_sap numalbprov
End Sub
Sub carregar_rutes_sap()
  If llegir_ini("Compres", "rutaSapSeidor_INPLACSA", "comandes.ini") = "{[}]" Then
     escriure_ini "Compres", "rutaSapSeidor_INPLACSA", "\\servidorsap\seidor_COMUNICADOR\ENTALBCOMPRAS\INPLACSA", "comandes.ini"
     escriure_ini "Compres", "rutaSapSeidor_PLASEL", "\\servidorsap\seidor_COMUNICADOR\ENTALBCOMPRAS\PLASEL", "comandes.ini"
  End If
End Sub

Sub generar_fitxer_sap(numalbp As String, vnumprov As String)
   Dim rstalb As Recordset
   Dim nomfitxer As String
   Dim numid As Double
   Dim vany As Double
   Dim vrutasapseidor As String
   vany = Year(Now)
   If Month(Now) = 1 And Day(Now) < 15 Then vany = vany - 1
   Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
   Set rstalb = dbcompres.OpenRecordset("select * from albaransbip where year(data)>=" + atrim(vany) + " and numalbaraprov='" + atrim(numalbp) + "' and codiproveidorcomercial=" + atrim(cadbl(vnumprov)))
   
   If Not rstalb.EOF Then
        carregar_rutes_sap
        If llegir_ini("Compres", "rutasapcompres", "comandes.ini") = "{[}]" Then
           MsgBox "No hi ha la ruta d'importació del SAP" + Chr(10) + "Es possarà la de per defecte" + Chr(10) + "\\servidorsap\SGI_COMUNICADOR\ENTALBCOMPRAS"
           escriure_ini "Compres", "rutasapcompres", "\\servidorsap\seidor_COMUNICADOR\ENTALBCOMPRAS", "comandes.ini"
        End If
          
         'nomfitxer = "C-" + Format(Now, "yymmddhhnnss") + ".csv"
         nomfitxer = "C-" + Mid(treuresimbols(rstalb!nomproveidorcomercial), 1, 12) + "-" + format(rstalb!data, "dd") + "-" + format(rstalb!data, "mm") + "-" + format(rstalb!data, "yyyy") + "-" + albprovsensebarres(rstalb!numalbaraprov) + ".csv"
         r = llegir_ini("Compres", "rutasapcompres", "comandes.ini") + "\" + atrim(rstalb!empresa) + "\" + atrim(nomfitxer)
       '//  vrutasapseidor = llegir_ini("Compres", "rutaSapSeidor_" + UCase(atrim(rstalb!empresa)), "comandes.ini") + "\" + atrim(nomfitxer)
         If existeix(r) Then Kill r
       '//  If existeix(vrutasapseidor) Then Kill vrutasapseidor
          Else: Exit Sub
   End If
   
   While Not rstalb.EOF
      numid = rstalb!id
      'If atrim(rstalb!nomfitxer) = "" Then
      rstalb.Edit: rstalb!nomfitxer = nomfitxer: rstalb!enviat = True: rstalb.Update
      'End If
      rstalb.FindFirst "id=" + atrim(numid)
      If rstalb.NoMatch Then
              MsgBox "No s'ha trobat l'albarà", vbCritical, "Error"
          Else
            generar_linia_fitxer_sap rstalb, r ', vrutasapseidor
      End If
      rstalb.MoveNext
   Wend
   
'//   'copio el fitxer a la ruta de seidor
'//   FileCopy r, vrutasapseidor
 '//  Set dbtintes = Nothing
End Sub
Sub generar_fitxer_bip(numalbp As String)
   Dim rstalb As Recordset
   Dim nomfitxer As String
   If llegir_ini("Compres", "rutabip", "comandes.ini") = "{[}]" Then
     ruta = "\\serconta\home\OpenSoft\COMPLEMENTS\Propuestas de compra"
     ruta = InputBox("No hi ha la ruta pels fitxers de BIP, aquesta es la de perdefecte", "Ruta exportació albarans pel programa de BIP", ruta)
     If Not existeix(ruta) Then MsgBox "Ruta no valida, no s'exportaràn els albarans": Exit Sub
     escriure_ini "Compres", "rutabip", ruta, "comandes.ini"
   End If
   Set rstalb = dbcompres.OpenRecordset("select * from albaransbip where numalbaraprov='" + atrim(numalbp) + "'")
   If Not rstalb.EOF Then
         
         nomfitxer = "CAL-" + atrim(rstalb!codiproveidorcomercial) + "-" + format(rstalb!data, "dd") + "-" + format(rstalb!data, "mm") + "-" + format(rstalb!data, "yyyy") + "-" + albprovsensebarres(rstalb!numalbaraprov) + ".txt"
         r = llegir_ini("Compres", "rutabip", "comandes.ini") + "\empre" + format(cadbl(llegir_ini("Compres", "numempresabip_" + atrim(rstalb!empresa), "comandes.ini")), "000") + "\" + atrim(nomfitxer)
         If existeix(r) Then Kill r
          Else: Exit Sub
   End If
   While Not rstalb.EOF
      nomfitxer = generarfitxeralbaratxt(rstalb)
      rstalb.Edit: rstalb!nomfitxer = nomfitxer: rstalb!enviat = True: rstalb.Update
      generar_linia_fitxer_bip rstalb
      rstalb.MoveNext
   Wend
End Sub
Function albprovsensebarres(valor As String) As String
  Dim symbols As Variant
  Dim i As Long
   symbols = Array("?", "/", "%", "$", "\", "@", "#") 'els sustingut marca l'ultim
  i = 0
  While symbols(i) <> "#"
     While InStr(1, valor, symbols(i)) > 0
       valor = substituir(valor, atrim(symbols(i)), "_")
     Wend
     i = i + 1
  Wend
  albprovsensebarres = valor
End Function
Function escullir_albara(vnumalb As String) As Double
      Load formseleccio
      formseleccio.Caption = "Escull l'albarà "
      formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "palets.mdb"
      formseleccio.Data1.RecordSource = "select data,nomproveidorcomercial,id from albaransbip where numalbaraprov='" + atrim(vnumalb) + "'"
      formseleccio.refrescar
      formseleccio.DBGrid2.Columns(0).Width = 1000
      formseleccio.DBGrid2.Columns(1).Width = 5000
      formseleccio.DBGrid2.Columns(2).Visible = False
      formseleccio.Width = 8200
      formseleccio.Command2.Tag = "0"
      formseleccio.Caption = "Escullir albarà"
      formseleccio.Show 1
      If seleccioret = 1 Then
           escullir_albara = cadbl(formseleccio.Data1.Recordset!id)
      End If
      Unload formseleccio
End Function
Sub impresiodalbara(numalbara As String)
   Dim fitxertemp As String
   Dim rstalbprov As Recordset
   Dim albprov As String
   Dim vidalbara As Double
   
   Static imprimint As Boolean
   If imprimint Then MsgBox "Ja s'esta imprimint un albara espera a que acavi per fer-ne un altra.", vbInformation + vbOKOnly, "Atenci": Exit Sub
   imprimint = True
   Set rstalbprov = dbcompres.OpenRecordset("select distinct data from albaransbip where numalbaraprov='" + atrim(numalbara) + "'")
   If rstalbprov.EOF Then MsgBox "No he trobat aquest albarà": Exit Sub
   rstalbprov.MoveLast
   rstalbprov.MoveFirst
   If rstalbprov.RecordCount > 1 Then
    etstatus = "Escullir albarà.": DoEvents
    vidalbara = escullir_albara(numalbara)
    If vidalbara = 0 Then GoTo fi
    Set rstalbprov = dbcompres.OpenRecordset("select * from albaransbip where id=" + atrim(vidalbara))
    If Not rstalbprov.EOF Then Set rstalbprov = dbcompres.OpenRecordset("select * from albaransbip where numalbaraprov='" + atrim(numalbara) + "' and data=#" + format(rstalbprov!data, "mm/dd/yy") + "#")
      Else: Set rstalbprov = dbcompres.OpenRecordset("select * from albaransbip where numalbaraprov='" + atrim(numalbara) + "'")
   End If
   Me.Caption = "Imprimint borrant temporals": DoEvents
   borrartotselstemporals
   fitxertemp = "c:\temp\~llac" + format(Now, "ddhhnnss") + ".mdb"
   If Not existeix(fitxertemp) Then DBEngine.CreateDatabase fitxertemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
   Set dbconsulta = OpenDatabase(fitxertemp)
   Me.Caption = "Imprimint passar registres a temporals": DoEvents
   passarregistrealataulatemporal cadbl(rstalbprov!numcomanda), fitxertemp, cadbl(rstalbprov!codiproveidorcomercial)
   Me.Caption = "Imprimint passar dades a l'albarà": DoEvents
   possarlesdadesdalbara rstalbprov
   If rstalbprov.EOF Then GoTo fi
   albprov = rstalbprov!numalbaraprov
   
   Me.Caption = "Imprimint imprimir albarà.": DoEvents
   imprimiralbara fitxertemp, albprov
   
fi:
   Set rstalbprov = Nothing
   imprimint = False
End Sub
Sub possarlesdadesdalbara(rsta As Recordset)
   Dim rstcap As Recordset
   Dim rstlinia As Recordset
   Dim rstliniescompra As Recordset
   Dim idcompra As Long
   Dim rstp As Recordset
   Dim subtotal As Double
   Dim rstalbaransbip As Recordset
   Dim rstprov As Recordset
   
   Set rstcap = dbconsulta.OpenRecordset("ll_capcalera")
   idcompra = rstcap!id
   Me.Caption = "Imprimint passar dades a l'albarà  [1)": DoEvents
   dbconsulta.Execute "delete * from ll_liniescompra"
   dbconsulta.Execute "delete * from ll_liniesdescripcio"
   Me.Caption = "Imprimint passar dades a l'albarà  [2)": DoEvents
   Set rstlinia = dbconsulta.OpenRecordset("ll_liniescompra")
   Me.Caption = "Imprimint passar dades a l'albarà  [3]": DoEvents
   Set rstp = dbcompres.OpenRecordset("select id, data from capcalera where numcomanda=" + atrim(rsta!numcomanda))
   Me.Caption = "Imprimint passar dades a l'albarà  [4]": DoEvents
   If Not rstp.EOF Then Set rstliniescompra = dbcompres.OpenRecordset("select * from liniescompra where idcompra=" + atrim(rstp!id))
   Me.Caption = "Imprimint passar dades a l'albarà  [5]": DoEvents
  ' Clipboard.Clear
  ' Clipboard.SetText "SELECT albaransbip.numalbara, proveidors.tipusproveidorIMPOST FROM (albaransbip LEFT JOIN (liniescompra LEFT JOIN capcalera ON liniescompra.idcompra = capcalera.id) ON albaransbip.idliniacompra = liniescompra.idliniacompra) LEFT JOIN proveidors ON capcalera.codiproveidor = proveidors.codi where numalbara=" + atrim(rsta!numalbara)
   'Set rstprov = dbcompres.OpenRecordset("SELECT albaransbip.numalbara, proveidors.tipusproveidorIMPOST FROM (albaransbip LEFT JOIN (liniescompra LEFT JOIN capcalera ON liniescompra.idcompra = capcalera.id) ON albaransbip.idliniacompra = liniescompra.idliniacompra) LEFT JOIN proveidors ON capcalera.codiproveidor = proveidors.codi where numalbara=" + atrim(rsta!numalbara))
   buscar_tipus_proveidor_IMPOST rstprov, atrim(rsta!numalbara)
   If rstprov.EOF Then Set rsta = rstprov: Exit Sub 'posso rsta=rstprov per posar a EOF rsta i així saber que no s'ha trobat l'albarà
    Me.Caption = "Imprimint passar dades a l'albarà  [Capçalera]": DoEvents
   While Not rsta.EOF
     Set rstp = dbcompres.OpenRecordset("select data from capcalera where numcomanda=" + atrim(rsta!numcomanda))
     If Not rstliniescompra.EOF Then rstliniescompra.FindFirst "codimaterial=" + atrim(rsta!article)
        Me.Caption = "Imprimint passar dades a l'albarà [LINIA " + atrim(rsta!article) + "]": DoEvents
     rstlinia.AddNew
     rstlinia!idcompra = idcompra
     rstlinia!codimaterial = rsta!article
     rstlinia!nommaterial = "Ped: " + atrim(rsta!numcomanda) + " Fecha " + format(rstp!data, "dd/mm/yy")
     rstlinia!quantitatkg = rsta!quantitat
     rstlinia!preu = rsta!preu
     If Not rstliniescompra.EOF Then If Not rstliniescompra.NoMatch Then rstlinia!tipusmaterialcomprat = rstliniescompra!tipusmaterialcomprat
    
'     rstlinia!preunet = rsta!quantitat * rsta!preu
     subtotal = subtotal + (rstlinia!preu * rstlinia!quantitatkg)
     rstlinia.Update
     rstlinia.Bookmark = rstlinia.LastModified
     dbconsulta.Execute "Insert into ll_liniesdescripcio (idliniacompra,descripcio,ordre) values (" + atrim(cadbl(rstlinia!idliniacompra)) + ",'" + treure_apostruf(rsta!descripcio) + "',1)"
     If rstprov!tipusproveidorIMPOST = "Espanyol" Then
        If cadbl(rsta!kgimpostenvasos) > 0 Then  'SI HI HA IMPOST HAIG DE CREAR UNA LINIA PER ENSENYAR-LO A L'IMPRESIÓ
          rstlinia.AddNew
          rstlinia!idcompra = idcompra
          rstlinia!codimaterial = "99999"
          rstlinia!nommaterial = "         ^ <Impost Envasos> ^"
          rstlinia!quantitatkg = rsta!kgimpostenvasos
          rstlinia!preu = cadbl(rsta!preuImpostEnvasos)
          If Not rstliniescompra.EOF Then If Not rstliniescompra.NoMatch Then rstlinia!tipusmaterialcomprat = rstliniescompra!tipusmaterialcomprat
          subtotal = subtotal + (rstlinia!preu * rstlinia!quantitatkg)
          rstlinia.Update
        End If
     End If
     rsta.MoveNext
   Wend
   'poso les initats correcte per cada linia
   dbconsulta.Execute "UPDATE ll_liniescompra INNER JOIN materials ON ll_liniescompra.codimaterial = materials.codi SET ll_liniescompra.desc_unitat = [materials].[mesuarespcompra];"
   rsta.MoveFirst
   rstcap.Edit
   rstcap!data = rsta!data
   rstdataentrega = "01/01/2000"
   rstcap!numcomanda = rsta!numalbara
   rstcap!empresa = rsta!empresa
   rstcap!baseimp = subtotal
   rstcap!codicomptable = atrim(rsta!codiproveidorcomercial)
   actualitzalesdadesdelacapcalera rstcap
   rstcap.Update
   Set rstcap = Nothing
   Set rstlinia = Nothing
   Set rstliniescompra = Nothing
End Sub
Sub buscar_tipus_proveidor_IMPOST(rstprov As Recordset, valb As String)
   Set rstprov = dbtmp.OpenRecordset("select * from albaransbip")
   If rstprov.EOF Then Exit Sub

   rstprov.FindFirst "numalbara=" + atrim(valb)
   If rstprov.NoMatch Then MsgBox "Albarà " + atrim(valb) + " no trobat.": Set rstprov = dbtmp.OpenRecordset("select * from albaransbip where 1=0"): Exit Sub
   Me.Caption = "Imprimint passar dades a l'albarà  [6]": DoEvents
   Set rstprov = dbtmp.OpenRecordset("select * from liniescompra where idliniacompra=" + atrim(rstprov!idliniacompra))
   If rstprov.EOF Then Exit Sub
   Me.Caption = "Imprimint passar dades a l'albarà  [7]": DoEvents
   Set rstprov = dbtmp.OpenRecordset("select * from capcalera where id=" + atrim(rstprov!idcompra))
   If rstprov.EOF Then Exit Sub
   Me.Caption = "Imprimint passar dades a l'albarà  [8]": DoEvents
   Set rstprov = dbtmp.OpenRecordset("select tipusproveidorIMPOST from proveidors where codi=" + atrim(rstprov!codiproveidor))
End Sub
Sub actualitzalesdadesdelacapcalera(rstcap As Recordset)
    Dim rst As Recordset
    Set rst = dbcomandes.OpenRecordset("Select * from proveidors_comercial where codicomptable='" + atrim(rstcap!codicomptable) + "'")
    If Not rst.EOF Then
        rstcap!nomprov = atrim(rst!nom)
        rstcap!direccio = atrim(rst!direccio)
        rstcap!codipipoblacio = atrim(rst!codipostal) + "-" + atrim(rst!poblacio)
        rstcap!provincia = atrim(rst!provinciapais)
    End If
    Set rst = Nothing
End Sub
Sub imprimiralbara(fitxertemp, numalbprov As String)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  'passarregistrealataulatemporal cnumcomanda
  wait 2
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "comandescompres.rpt", 1)
  oreport.Database.Tables.Item(1).Location = fitxertemp
  oreport.FormulaFields.GetItemByName("albaraproveidor").Text = "'" + numalbprov + "'"
  oreport.DiscardSavedData
  If existeix("c:\ordprog.ini") Then
     Load veurereport
     veurereport.CRViewer.ReportSource = oreport
     veurereport.CRViewer.DisplayGroupTree = False
     veurereport.CRViewer.ViewReport
     veurereport.Show 1, Me
    Else
      oreport.PrintOut False, 1
  End If
  Unload veurereport
  'comandescompra.SetFocus

End Sub


Sub passarregistrealataulatemporal(numc As Double, fitxertemp As String, proveidorcomercial As Double)
  Dim rstp As Recordset
  Dim i As Byte
  Dim rstm As Recordset
  Dim gran As Integer
  Dim rstf As Recordset
  dbcompres.Execute "select * into ll_capcalera IN '" + fitxertemp + "' from capcalera where numcomanda=" + atrim(numc)
  dbcompres.Execute "SELECT liniescompra.* into ll_liniescompra IN '" + fitxertemp + "' FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE (((capcalera.numcomanda)=" + atrim(numc) + "));"
  dbcompres.Execute "SELECT liniesdescripcio.* into ll_liniesdescripcio IN '" + fitxertemp + "' FROM (capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra) RIGHT JOIN liniesdescripcio ON liniescompra.idliniacompra = liniesdescripcio.idliniacompra WHERE (((capcalera.numcomanda)=" + atrim(numc) + "));"
  dbconsulta.Execute "create index principal ON ll_liniescompra([idcompra]);"
  dbconsulta.Execute "create index segona ON ll_liniescompra([idliniacompra]);"
  dbconsulta.Execute "create index principal ON ll_liniesdescripcio([idliniacompra]);"
  dbconsulta.Execute "create index principal ON ll_capcalera([id]);"
  dbconsulta.Execute "alter table ll_liniescompra add column desc_unitat text"
  dbcompres.Execute "select materials.* into materials in '" + fitxertemp + "' from materials"
  
  Set rstp = dbconsulta.OpenRecordset("select max(idcompra) as gran from ll_liniescompra")
  If rstp.EOF Then Exit Sub
  dbconsulta.Execute "alter table ll_liniesdescripcio add column tipus byte"
  gran = 100
  If Not rstp.EOF Then gran = cadbl(rstp!gran)
  Set rstp = dbtmpb.OpenRecordset("select * from proveidors_comercial where codicomptable='" + atrim(proveidorcomercial) + "'")
  For i = 0 To 9
    Set rstm = dbcompres.OpenRecordset("select * from descripcionsmsgpeu where idmsg=" + atrim(cadbl(rstp.Fields("msg" + atrim(i + 1)))) + " order by ordre")
    If Not rstm.EOF Then dbconsulta.Execute "insert into ll_liniescompra (idcompra) values (" + atrim(gran) + ")"
    Set rstf = dbconsulta.OpenRecordset("select max(idliniacompra) as gran from ll_liniescompra ")
    If Not rstf.EOF Then gran = cadbl(rstf!gran)
    While Not rstm.EOF
      dbconsulta.Execute "insert into ll_liniesdescripcio (idliniacompra,ordre,descripcio,tipus) values (" + atrim(gran) + "," + atrim(rstm!ordre) + ",'" + treure_apostruf(rstm!descripcio) + "',1)"
      rstm.MoveNext
    Wend
    gran = gran + 1
  Next i
  'poso les initats correcte per cada linia
  dbconsulta.Execute "UPDATE ll_liniescompra INNER JOIN materials ON ll_liniescompra.codimaterial = materials.codi SET ll_liniescompra.desc_unitat = [materials].[mesuarespcompra];"
  dbconsulta.Execute "update ll_capcalera set dataentrega=#01/01/2000#"
End Sub



Private Sub kgentregats_GotFocus()
  kgentregats.SelStart = 0
  kgentregats.SelLength = Len(kgentregats)
End Sub

Private Sub kgentregats_KeyPress(KeyAscii As Integer)
  If Chr$(KeyAscii) = "." Then KeyAscii = 0: SendKeys ","
End Sub

Private Sub kgimpostenv_LostFocus()
  Dim vcalculimpost As Double
  If cadbl(kgimpostenv) > 0 Then
      vcalculimpost = (cadbl(impostbaseimp) / 100) * cadbl(kgimpostenv.Tag)
      If kgimpostenv <> vcalculimpost Then MsgBox "El valor de Kg corresponent al " + kgimpostenv.Tag + "% d'Impost d'envasos d'aquest material hauria de ser " + atrim(vcalculimpost) + "Kg." + vbNewLine + "REVISA EL VALOR DEL TAN % DEL MATERIAL.", vbCritical, "ATENCIÓ"
  End If
End Sub

Private Sub lotproveidor_Change()
  If lotproveidor = "" Then bpdfcq.Visible = False: imgNOcq.Visible = False: imgSIcq.Visible = False
  bpdfcq.Tag = ""
End Sub
Sub comprovar_cq(vlot As String, vcodiproveidor As Double, vcodiprovproduccio As Double)
   Dim rst As Recordset
   Dim vnomfitxer As String
   etcq.Visible = False
   bpdfcq.Visible = False: bpdfcq.Tag = "": bpdfcq.Enabled = False
   imgNOcq.Visible = False: imgSIcq.Visible = False
   Set rst = dbqualitat.OpenRecordset("select * from proveidors_qualitat where codiproveidor=" + atrim(vcodiprovproduccio))
   If rst.EOF Then MsgBox "No hi ha informació de QUALITAT en aquest proveidor.", vbCritical, "Atenció": Exit Sub
   If rst!tipuscontrolCQ = "" Then MsgBox "AQUEST PROVEIDOR NO TÉ ASSIGNAT CAP TIPUS DE CONTROL DE QUALITAT.", vbCritical, "ERROR": GoTo fi
   If rst!tipuscontrolCQ = "Calidad concertada" Then
          If DateDiff("d", Now, rst!CQ_datacaducitat) < 0 Then
                 bpdfcq.Visible = False: bpdfcq.Enabled = False: imgSIcq.Visible = False: imgNOcq.Visible = False
                 MsgBox "AQUEST PROVEIDOR TE Calidad concertada PERÒ LA DATA ESTÀ CADUCADA.", vbCritical, "ATENCIÓ"
                 GoTo fi2
                   Else:
                      etcq.Visible = True
                      GoTo fi2
          End If
   End If
   Set rst = dbtmpb.OpenRecordset("select * from registre_escanejades_expedicions where nomfitxer like 'CQ_" + vlot + " ?" + atrim(vcodiproveidor) + "?-*'")
   If Not rst.EOF Then
        vnomfitxer = rst!rutadestilocal + rst!nomfitxer
        If existeix(vnomfitxer) Then
            bpdfcq.Visible = True
            bpdfcq.Enabled = True
            bpdfcq.Tag = vnomfitxer
        End If
       Else: bpdfcq.Visible = True: bpdfcq.Tag = "": bpdfcq.Enabled = False
   End If
fi:
  
   imgNOcq.Visible = Not bpdfcq.Enabled: imgSIcq.Visible = bpdfcq.Enabled
   bpdfcq.Enabled = True
fi2:
    Set rst = Nothing
   
End Sub
Private Sub lotproveidor_LostFocus()
   If combocomandescompra = "" Or bpdfcq.Tag <> "" Then Exit Sub
   comprovar_cq lotproveidor, codiproveidor(rstcompres!codiproveidorcomercial), rstcompres!codiproveidor
End Sub

Private Sub tancar_Click()
  comprespalets.Hide
End Sub

