VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form formliniesmodificacio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniment de Linies de Modificació"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   14190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame framebotons2 
      Height          =   585
      Left            =   30
      TabIndex        =   2
      Top             =   -60
      Width           =   14085
      Begin Crystal.CrystalReport llistat 
         Left            =   2025
         Top             =   105
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton imprimir 
         Height          =   360
         Left            =   960
         Picture         =   "formliniesmodificacio.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   13605
         Picture         =   "formliniesmodificacio.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Sortir"
         Top             =   135
         Width           =   390
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   120
         Picture         =   "formliniesmodificacio.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton guardar 
         Height          =   360
         Left            =   540
         Picture         =   "formliniesmodificacio.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   375
      End
      Begin MSComCtl2.DTPicker picker 
         Height          =   315
         Left            =   3345
         TabIndex        =   7
         Top             =   210
         Visible         =   0   'False
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/mm/yy"
         Format          =   662831107
         CurrentDate     =   41303
      End
   End
   Begin VB.Data modifis 
      Caption         =   "modifis"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1635
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clixes_modifi"
      Top             =   -30
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame framemodi 
      Caption         =   " Linies de Modificació"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5085
      Left            =   45
      TabIndex        =   0
      Top             =   510
      Width           =   14070
      Begin MSDBGrid.DBGrid reixamodifis 
         Bindings        =   "formliniesmodificacio.frx":1628
         Height          =   4785
         Left            =   90
         OleObjectBlob   =   "formliniesmodificacio.frx":163A
         TabIndex        =   1
         Tag             =   "modifis"
         Top             =   225
         Width           =   13785
      End
   End
End
Attribute VB_Name = "formliniesmodificacio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub eliminar_Click()
 If MsgBox("Segur que vols borrar aquesta linia?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
      
                If Not modifis.Recordset.EOF Then
                       modifis.Recordset.Delete
                       modifis.Refresh
                       gravar_canvis
                End If
    
  End If
End Sub
Sub gravar_canvis()
  Dim bk As Long
  bk = modifis.Recordset!ordre
  If Not gravar_reixa(reixamodifis) Then Exit Sub
   
'   If clixes.Recordset.EditMode > 0 Then clixes.Recordset.Update
   'modificar_Click
   On Error GoTo erro
   If modifis.Recordset.EditMode > 0 Then modifis.Recordset.Update
   reixamodifis.SetFocus
   modifis.Recordset.FindFirst "ordre=" + atrim(cadbl(bk))
   Exit Sub
erro:
    MsgBox err.Description
    
End Sub
Function gravar_reixa(reixa As DBGrid) As Boolean
    Dim fila As Double
    gravar_reixa = True
    If reixa.visible And reixa.row > 0 Then
     fila = reixa.row
     reixa.SetFocus
     SendKeys "{down}"
     DoEvents
     If reixa.row <> fila + 1 Then gravar_reixa = False
     reixa.row = fila
    End If
    
End Function

Private Sub Form_Load()
  Dim idtreball As String
  Dim ordremodifi As String
  If cadbl(formclixes.modificacions.Recordset!id_treball) = 0 Then Exit Sub
  idtreball = atrim(formclixes.modificacions.Recordset!id_treball)
  ordremodifi = atrim(formclixes.modificacions.Recordset!ordre)
  modifis.DatabaseName = formclixes.clixes.DatabaseName
  modifis.RecordSource = "select * from clixes_modifi where id_treball=" + idtreball + " and ordremodificacio=" + ordremodifi
  modifis.Refresh
End Sub

Private Sub guardar_Click()
  gravar_canvis
End Sub

Private Sub imprimir_Click()
  formclixes.crear_taules_tmp
   imprimir_modificacions
End Sub
Sub imprimir_modificacions()
Dim rstimp As Recordset
  Set rstimp = dbclixes.OpenRecordset("tmp_clixes_capcalera")
  If Not rstimp.EOF Then
     While Not rstimp.EOF
       rstimp.Delete
       rstimp.MoveNext
     Wend
  End If
  With formclixes
  rstimp.AddNew
   rstimp!id_treball = .clixes.Recordset!id_treball
  rstimp!arxiuclixe = .clixes.Recordset!arxiu
  rstimp!datainici = .modificacions.Recordset!dataobertura
  rstimp!formaimp = .formaimpresio
  rstimp!estatclixe = .clixes.Recordset!estatclixe
  rstimp!client = .nomclient
  If atrim(rstimp!client) = "" Then rstimp!client = .nomclienttemporal
  rstimp!marca = .marcaproducte
  rstimp!linia = .liniaproducte
  'rstimp!representant = .nomrepresentant
  rstimp!proveidor = .nomproveidor
  'rstimp!montadora = clixes.Recordset!montadora
  rstimp!codibarres = .clixes.Recordset!codidebarres
  rstimp!dataentrega = IIf(IsDate(.modificacions.Recordset!datatancament), .modificacions.Recordset!datatancament, .modificacions.Recordset!datatancamentprevista)
  rstimp!observacions = .modificacions.Recordset!observacions
  rstimp!sistemaimpresio = .modificacions.Recordset!sistemadimpresio
  rstimp!bandesclixes = .modificacions.Recordset!bandes
  rstimp!ample = .modificacions.Recordset!amplelamina
  rstimp!desarroll = .modificacions.Recordset!desarroll
  rstimp.Update
  
  Set rstimp = dbclixes.OpenRecordset("tmp_clixes_modifis_linies")
  If Not rstimp.EOF Then
     While Not rstimp.EOF
       rstimp.Delete
       rstimp.MoveNext
     Wend
  End If
  modifis.Recordset.MoveFirst
  While Not modifis.Recordset.EOF
   If atrim(modifis.Recordset!imprimiracomanda) <> "Sí" Then
    rstimp.AddNew
    rstimp!id_treball = modifis.Recordset!id_treball
    rstimp!descripcio = modifis.Recordset!descripcio
    rstimp!inici = modifis.Recordset!data_inici
    rstimp!fi = modifis.Recordset!data_fi
    rstimp.Update
   End If
   modifis.Recordset.MoveNext
  Wend
  End With
  llistat.Formulas(0) = ""
  llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "clixes_modificacions.rpt"
 llistat.DataFiles(0) = camiclixes
  llistat.DiscardSavedData = True
 llistat.Destination = crptToPrinter
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 wait 1
 llistat.Action = 1
  
End Sub


Sub ensenyar_picker(reixa As DBGrid, taula As data)
  If reixa.Columns(reixa.col).Locked = True Then Exit Sub
  r = ""
  If taula.Recordset.Fields(reixa.Columns(reixa.col).DataField).Type = 8 Then
       picker.visible = True
       If IsDate(reixa) Then
          picker.Value = reixa
         Else: picker.Value = Now
       End If
       picker.Move reixa.Container.Left + reixa.Columns(reixa.col).Left, reixa.Container.Top + reixa.Columns(reixa.col).Top
       picker.SetFocus
       SendKeys ("%{down}")
       picker.tag = reixa.Name
         Else: picker.visible = False
  End If
   
End Sub
Sub demanaridestat()
   Load formseleccio
   formseleccio.Data1.DatabaseName = formclixes.clixes.DatabaseName
   formseleccio.Data1.RecordSource = "select * from clixes_estats"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns("id_estat").width = 0
   formseleccio.Show 1
   'If seleccioret = 1 Then   'aixó es va treure quan vam passar a fer liniadimpresió a tintes
   '     If cadbl(formclixes.modificacions.Recordset!numerodelinia) = 0 And formseleccio.DBGrid2.Columns("descripcio") = "CLIXES ENTRATS" Then
   '        MsgBox "No pots passar els CLIXES A ENTRATS sense haver assignat una linia d'impresió a la modificació", vbCritical, "Error"
   '        Exit Sub
   '     End If
   'End If
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           reixamodifis.Columns("descripcioestat") = formseleccio.DBGrid2.Columns("descripcio")
           reixamodifis.Columns("id_estatclixe") = formseleccio.DBGrid2.Columns("id_estat")
        End If
   End If
   If seleccioret = 9 Then
            reixamodifis.Columns("descripcioestat") = ""
           reixamodifis.Columns("id_estatclixe") = Null
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub


Private Sub picker_CloseUp()
formliniesmodificacio.Controls(picker.tag) = picker.Value
 formliniesmodificacio.Controls(picker.tag).SetFocus
End Sub

Private Sub reixamodifis_ButtonClick(ByVal ColIndex As Integer)
 Dim ordre As Long
   Dim nomdata As String
   
   nomdata = reixamodifis.Columns(ColIndex).DataField
   If IsDate(reixamodifis.Columns(ColIndex)) Then
     If UCase(InputBox("Segur que vols eliminar aquesta data?" + Chr(10) + "Escriu ELIMINAR per eliminar-la.", "Eliminar data")) = "ELIMINAR" Then
         ordre = modifis.Recordset!ordre
         
         modifis.Database.Execute "update clixes_modifi set " + nomdata + "=null where id_treball=" + atrim(modifis.Recordset!id_treball) + " AND ordremodificacio=" + atrim(modifis.Recordset!ordremodificacio) + " and ordre=" + atrim(modifis.Recordset!ordre)
         modifis.Refresh
         modifis.Recordset.FindFirst "ordre=" + atrim(ordre)
     End If
   End If
   If nomdata = "descripcioestat" Then
       demanaridestat
       If reixamodifis.Columns("descripcioestat") = "CLIXES ENTRATS" And Not formclixes.modificacions.Recordset!pdfvalid Then
          MsgBox "No tens entrat el PDF en aquesta versió del treball, no t'oblidis de possar-la o copiar-la de la versió anterior", vbCritical, "PDF"
       End If
       If reixamodifis.Columns("descripcioestat") = "CLIXES ENTRATS" And InStr(1, " " + formclixes.etestatclixemod, "REPOSICIÓ DEL CLIXE") > 0 Then
           Set rst = dbclixes.OpenRecordset("select * from reposicionsfotogravador where id_treball=" + atrim(id_treball) + " AND ordremodificacio=" + atrim(ordremodificacio) + " order by dataenviament desc")
           If Not rst.EOF Then
               If rst!modificatperinplacsa Then
                  MsgBox "La reposició de clixes estava MODIFICADA PER INPLACSA comprova que estigui tot correcte.", vbCritical, "Atenció"
               End If
           End If
           'MsgBox "PASSO AVIS A IMPRESORES AVISANT QUE JA ÈS AQUI LA REPOSICIÓ.", vbInformation, "REPOSICIÓ A INPLACSA"
           'passar_avis_reposicio_a_inplacsa    'ara es fa al sortir de modificacions quan ha canviat l'estat a clixes entrats
       End If
   End If
End Sub
Sub passar_avis_reposicio_a_inplacsa()
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from clixesentrats_control where format(dataentrada,'dd/mm/yy')=format('" + atrim(Date) + "','dd/mm/yy') and numtreball=" + atrim(id_treball) + " and versio=" + atrim(ordremodificacio))
   If Not rst.EOF Then GoTo fi
   rst.AddNew
   rst!numtreball = id_treball
   rst!versio = ordremodificacio
   rst!desarroll = formclixes.modificacions.Recordset!desarroll
   rst!dataentrada = Now
   rst!reposicio = True
   rst.Update
fi:
   Set rst = Nothing

End Sub

Private Sub reixamodifis_ColEdit(ByVal ColIndex As Integer)
If reixamodifis.Columns("Estat del Clixé") = "" Then
      If Not possardatafialanteriormodificacio Then Exit Sub
      reixamodifis.col = 1
      demanaridestat
      reixamodifis.col = 2
   End If
End Sub
Private Sub reixamodifis_DblClick()
 'If modifis.Recordset.EditMode = 0 Then modifis.Recordset.Edit
 'modifis.Recordset.Update
 If Not gravar_reixa(reixamodifis) Then Exit Sub
 ensenyar_picker reixamodifis, modifis
 comprovarsiestaseleccionat
 
End Sub

Private Sub reixamodifis_GotFocus()
 reixamodifis.Columns("imprimiracomanda").Locked = True
' If clixes.Recordset.EditMode > 0 Then clixes.UpdateRecord: clixes.Recordset.Edit ': activarframes True
End Sub

Sub comprovarsiestaseleccionat()
  If reixamodifis.Columns("descripcio").Locked Then Exit Sub
 If reixamodifis.Columns(reixamodifis.col).DataField = "imprimiracomanda" Then
     If reixamodifis.Columns("imprimiracomanda") = "Sí" Then
         reixamodifis.Columns("imprimiracomanda") = "No"
        Else:
          If MsgBox("Vols passar tots els pendents de marcar a fets?", vbInformation + vbYesNo, "Atenció") = vbYes Then
            dbtmp.Execute "update clixes_modifi set imprimiracomanda='Sí' where id_treball=" + atrim(cadbl(clixes.Recordset!id_treball))
            modifis.Refresh
            modifis.Recordset.MoveLast
            reixamodifis.Refresh
              Else:
                 reixamodifis.Columns("imprimiracomanda") = "Sí"
                 'If modifis.Recordset.EditMode = 0 Then modifis.Recordset.Edit
                 'modifis.Recordset.Update
                 gravar_reixa reixamodifis

          End If
            
     End If
    
 End If

End Sub
Private Sub reixamodifis_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 32 Then comprovarsiestaseleccionat

End Sub

Private Sub reixamodifis_OnAddNew()
  Dim gran As Integer
  If Not possardatafialanteriormodificacio Then Exit Sub
  gran = albmaxordre(modifis)
    'modifis.Recordset.AddNew
    modifis.Recordset!id_treball = formclixes.modificacions.Recordset!id_treball
    modifis.Recordset!ordremodificacio = formclixes.modificacions.Recordset!ordre
    modifis.Recordset!ordre = gran + 1
    'reixamodifis.Columns("data_inici") = Format(Now, "dd/mm/yy")
End Sub
Function albmaxordre(dbctrl As Control) As Integer
   Dim rs As Recordset
    'If albarans.Recordset.EOF Then albmaxordre = 0
   Set rs = dbctrl.Recordset.Clone
   If Not rs.EOF Then
     rs.MoveLast
     albmaxordre = cadbl(rs!ordre)
    Else: albmaxordre = 0
   End If
End Function
Function possardatafialanteriormodificacio() As Boolean
  Dim rst As Recordset
  possardatafialanteriormodificacio = True
  Set rst = modifis.Recordset.Clone
  If Not rst.EOF Then
    rst.MoveLast
    If Not IsDate(rst!data_fi) Then
       possardatafialanteriormodificacio = False
       If modifis.Recordset.EditMode > 0 Then modifis.Recordset.CancelUpdate
       modifis.Refresh
       modifis.Recordset.FindFirst "ordre=" + atrim(rst!ordre)
       data = InputBox("Entra la data de finalització de la modificació" + Chr(10) + atrim(rst!descripcio), "Data Fi", Format(Now, "dd/mm/yy"))
       If Not IsDate(data) Then
          Exit Function
         Else
           modifis.Recordset.Edit
           modifis.Recordset!data_fi = Format(data, "dd/mm/yy")
           modifis.Recordset.Update
           'modifis.Recordset.Move modifis.Recordset.RecordCount
           reixamodifis.SetFocus
       End If
    End If
  End If
  Set rst = Nothing
End Function
Private Sub sortir_Click()
  Unload formliniesmodificacio
End Sub
Sub imprimir_albarans()
  Dim rstimp As Recordset
  Dim facturats As Boolean
  facturats = True
  If MsgBox("Vols imprimir els albarans NO FACTURATS?" + Chr(10) + Chr(13) + "Si prems No es faran els FACTURATS", vbYesNo + vbDefaultButton1, "Escull") = vbYes Then
      facturats = False
  End If
  Set rstimp = dbclixes.OpenRecordset("tmp_clixes_capcalera")
  If Not rstimp.EOF Then
     While Not rstimp.EOF
       rstimp.Delete
       rstimp.MoveNext
     Wend
  End If
  With formclixes
  rstimp.AddNew
  rstimp!id_treball = .clixes.Recordset!id_treball
  rstimp!arxiuclixe = .clixes.Recordset!arxiu
  rstimp!datainici = .modificacions.Recordset!dataobertura
  rstimp!formaimp = .formaimpresio
  rstimp!estatclixe = .clixes.Recordset!estatclixe
  rstimp!client = .nomclient
  If atrim(rstimp!client) = "" Then rstimp!client = .nomclienttemporal
  rstimp!marca = .marcaproducte
  rstimp!linia = .liniaproducte
  'rstimp!representant = .nomrepresentant
  rstimp!proveidor = .nomproveidor
  'rstimp!montadora = clixes.Recordset!montadora
  rstimp!codibarres = .clixes.Recordset!codidebarres
  rstimp!dataentrega = IIf(IsDate(.modificacions.Recordset!datatancament), .modificacions.Recordset!datatancament, .modificacions.Recordset!datatancamentprevista)
  rstimp!observacions = .modificacions.Recordset!observacions
  rstimp!sistemaimpresio = .modificacions.Recordset!sistemadimpresio
  rstimp!bandesclixes = .modificacions.Recordset!bandes
  rstimp!ample = .modificacions.Recordset!amplelamina
  rstimp!desarroll = .modificacions.Recordset!desarroll
  rstimp.Update
  
  Set rstimp = dbclixes.OpenRecordset("tmp_clixes_albarans_linies")
  If Not rstimp.EOF Then
     While Not rstimp.EOF
       rstimp.Delete
       rstimp.MoveNext
     Wend
  End If
  albarans.Recordset.MoveFirst
  While Not albarans.Recordset.EOF
   If albarans.Recordset!facturat = facturats Then
    rstimp.AddNew
    rstimp!id_treball = albarans.Recordset!id_treball
    rstimp!data = albarans.Recordset!data
    rstimp!numalb = albarans.Recordset!num_alb
    rstimp!quantitat = albarans.Recordset!quantitat
    rstimp!descripcio = albarans.Recordset!descripcio
    rstimp!import = albarans.Recordset!import
    rstimp!facturat = IIf(albarans.Recordset!facturat, "Sí", "No")
    rstimp.Update
   End If
    albarans.Recordset.MoveNext
  Wend
  
  
   llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "clixes_albarans.rpt"
 llistat.DataFiles(0) = camiclixes
 If facturats Then
    llistat.Formulas(0) = "facturatsono='Albarans Facturats'"
   Else
     llistat.Formulas(0) = "facturatsono='Albarans NO Facturats'"
 End If
 llistat.DiscardSavedData = True
 llistat.Destination = crptToPrinter
 wait 1
 llistat.Action = 1
  End With
End Sub
