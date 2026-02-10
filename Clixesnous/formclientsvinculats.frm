VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formclientsvinculats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clients Vinculats"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10530
   ControlBox      =   0   'False
   Icon            =   "formclientsvinculats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4485
   ScaleWidth      =   10530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bsortida 
      Height          =   690
      Left            =   9600
      Picture         =   "formclientsvinculats.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Escullir el tipus de sortida i orientació de la imatge que vol el client."
      Top             =   750
      Width           =   780
   End
   Begin VB.Frame fcrearimp 
      BackColor       =   &H002E80A7&
      Caption         =   "Crear IMP"
      Height          =   1710
      Left            =   2250
      TabIndex        =   18
      Top             =   2730
      Visible         =   0   'False
      Width           =   6030
      Begin VB.Frame Frame2 
         BackColor       =   &H00209BCA&
         Caption         =   "Crear copiant d'un altre Treball"
         Height          =   1320
         Left            =   240
         TabIndex        =   19
         Top             =   270
         Width           =   5490
         Begin VB.ComboBox direccioenvionouimp 
            Height          =   315
            Left            =   1695
            TabIndex        =   27
            Top             =   975
            Width           =   2685
         End
         Begin VB.CommandButton Command2 
            Height          =   390
            Left            =   5010
            Picture         =   "formclientsvinculats.frx":1654
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Sortir"
            Top             =   420
            Width           =   390
         End
         Begin VB.CommandButton Command1 
            Height          =   360
            Left            =   4515
            Picture         =   "formclientsvinculats.frx":1BDE
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   435
            Width           =   420
         End
         Begin VB.ComboBox nomclientnouimp 
            Height          =   315
            Left            =   1710
            TabIndex        =   22
            Top             =   465
            Width           =   2685
         End
         Begin VB.TextBox campid_treball 
            Height          =   285
            Left            =   240
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   465
            Width           =   885
         End
         Begin VB.Label Label6 
            BackColor       =   &H00209BCA&
            Caption         =   "Direcció d'Envio"
            Height          =   285
            Left            =   2070
            TabIndex        =   28
            Top             =   795
            Width           =   1380
         End
         Begin VB.Label Label5 
            BackColor       =   &H00209BCA&
            Caption         =   "Nom del Client"
            Height          =   285
            Left            =   2070
            TabIndex        =   23
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "NºTreball"
            Height          =   225
            Left            =   330
            TabIndex        =   21
            Top             =   240
            Width           =   885
         End
      End
   End
   Begin VB.CommandButton botoimp 
      DisabledPicture =   "formclientsvinculats.frx":2168
      DownPicture     =   "formclientsvinculats.frx":27E2
      Height          =   690
      Left            =   8775
      Picture         =   "formclientsvinculats.frx":4324
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   780
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dades Client vinculat"
      Enabled         =   0   'False
      Height          =   1485
      Left            =   15
      TabIndex        =   5
      Top             =   570
      Width           =   10440
      Begin VB.ComboBox direnvio 
         DataField       =   "nomdirenvio"
         DataSource      =   "datavinculats"
         Height          =   315
         Left            =   3120
         TabIndex        =   29
         Top             =   525
         Width           =   2685
      End
      Begin VB.CheckBox cprincipal 
         Caption         =   "Client Principal"
         DataField       =   "principal"
         DataSource      =   "datavinculats"
         Height          =   210
         Left            =   1800
         TabIndex        =   26
         Top             =   135
         Width           =   1605
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFC0C0&
         DataField       =   "refclientalternatives"
         DataSource      =   "datavinculats"
         Height          =   285
         Left            =   1740
         TabIndex        =   12
         Top             =   1095
         Width           =   7725
      End
      Begin VB.TextBox codidebarres 
         BackColor       =   &H00FFC0C0&
         DataField       =   "refclient"
         DataSource      =   "datavinculats"
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   1095
         Width           =   1530
      End
      Begin VB.TextBox codimuntadora 
         DataField       =   "codimuntadora"
         DataSource      =   "datavinculats"
         Height          =   285
         Left            =   6030
         TabIndex        =   10
         Top             =   525
         Width           =   2325
      End
      Begin VB.ComboBox nomclient 
         DataField       =   "nomclient"
         DataSource      =   "datavinculats"
         Height          =   315
         Left            =   150
         TabIndex        =   8
         Top             =   540
         Width           =   2685
      End
      Begin VB.Label Label7 
         Caption         =   "Direcció d'enviament"
         Height          =   285
         Left            =   3480
         TabIndex        =   30
         Top             =   300
         Width           =   2325
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencies alternatives"
         Height          =   270
         Left            =   2325
         TabIndex        =   16
         Top             =   870
         Width           =   3360
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia client"
         Height          =   270
         Left            =   270
         TabIndex        =   15
         Top             =   870
         Width           =   1275
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Codi Muntadora"
         Height          =   270
         Left            =   6615
         TabIndex        =   14
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Nom del Client"
         Height          =   285
         Left            =   510
         TabIndex        =   9
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label missatgenouimp 
         Caption         =   "Fes Clic al Botó per crear el IMP -->"
         ForeColor       =   &H00008080&
         Height          =   255
         Left            =   6210
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   2730
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "formclientsvinculats.frx":5E66
      Height          =   2310
      Left            =   15
      OleObjectBlob   =   "formclientsvinculats.frx":5E7E
      TabIndex        =   4
      Top             =   2115
      Width           =   10440
   End
   Begin VB.Frame framebotons2 
      Height          =   585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.Data datavinculats 
         Caption         =   "datavinculats"
         Connect         =   "Access"
         DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\Clixesnous.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   4095
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Clientsvinculats"
         Top             =   225
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "formclientsvinculats.frx":744D
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   495
         Picture         =   "formclientsvinculats.frx":79D7
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Modificar Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton guardar 
         Height          =   360
         Left            =   1335
         Picture         =   "formclientsvinculats.frx":7F61
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   915
         Picture         =   "formclientsvinculats.frx":84EB
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   9945
         Picture         =   "formclientsvinculats.frx":8A75
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Sortir"
         Top             =   135
         Width           =   390
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   3375
      Left            =   0
      Picture         =   "formclientsvinculats.frx":8FFF
      Top             =   0
      Width           =   5835
   End
End
Attribute VB_Name = "formclientsvinculats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub triarcomandapervincularclient()
  Dim were As String
  Dim were2 As String
  Dim sql As String
  Dim rst As Recordset
  Dim rstcli As Recordset
  were = "comandes.numtreball=" + atrim(id_treball) + " AND comandes.numordremodificacio=" + atrim(ordremodificacio)
  were2 = "Clientsvinculats.id_treball=" + atrim(id_treball) + " AND Clientsvinculats.ordremodificacio=" + atrim(ordremodificacio)
  
  sql = "SELECT Max(comandes.comanda) AS Comanda, comandes.direnvio as Dir_Envio from comandes Where " + were + " GROUP BY comandes.direnvio HAVING comandes.direnvio Not In (SELECT Clientsvinculats.direnvio from Clientsvinculats WHERE " + were2 + ");"
'  MsgBox sql
  Set rst = dbclixes.OpenRecordset(sql)
  If rst.EOF Then MsgBox ("No hi ha cap comanda per crear client vinculat." + Chr(10) + "POTSER ENCARA NO S'HA DONAT D'ALTA LA COMANDA." + Chr(10) + "O JA ESTÀ DONAT D'ALTA."): Exit Sub
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = camiclixes
  formseleccio.Data1.RecordSource = sql
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).visible = False
  'formseleccio.DBGrid2.Columns(2).Width = 900
  'formseleccio.Width = 9000
  'formseleccio.Left = formseleccio.Left - 3000
  
     formseleccio.Show 1
  
   If seleccioret = 1 Then
        Set rst = datavinculats.Recordset
        Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(formseleccio.DBGrid2.Columns("Comanda")))
        Set rstcli = dbcomandes.OpenRecordset("SELECT Clients_envios.id, clients.nom, Clients_envios.poblacioe FROM Clients_envios INNER JOIN clients ON Clients_envios.codi = clients.codi WHERE Clients_envios.id=" + atrim(cadbl(rstc!direnvio)) + ";")
        If rstcli.EOF Then Exit Sub 'sinotrobo el client també surtu
        rst.AddNew
        rst!id_treball = cadbl(id_treball)
        rst!ordremodificacio = cadbl(ordremodificacio)
        rst!codiclient = rstc!client
        rst!direnvio = rstc!direnvio
        rst!nomclient = atrim(rstcli!nom)
        rst!nomdirenvio = atrim(rstcli!poblacioe)
        rst!codimuntadora = atrim(rstc!arxiumontadora)
        rst!refclient = atrim(rstc!refclient)
        rst!refclientalternatives = atrim(rstc!refclialt)
        rst!arxiuimp = False
        rst.Update
    End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   Set rst = Nothing
   Set rstcli = Nothing
   'codimuntadora.SetFocus
End Sub

Private Sub alta_Click()
    'Frame1.Enabled = True
    triarcomandapervincularclient
    datavinculats.Recordset.Bookmark = datavinculats.Recordset.LastModified
    
End Sub



Private Sub botoimp_Click()
  Dim nomfitxer As String
  Dim vimpnou As Boolean
  Dim vmsg As String
  If datavinculats.Recordset.EOF Or datavinculats.Recordset.BOF Then MsgBox "No hi ha client escullit": Exit Sub
  If datavinculats.Recordset.EditMode > 0 Then
   guardar_Click
  End If
  ratoli "espera"
  nomfitxer = generarfitxer_imp(True)

  If Not datavinculats.Recordset!arxiuimp And (existeix(nomfitxer) Or existeix(nomfitxer + "x")) Then
     MsgBox "Ja existeix l'IMP per aquest client... el vincularé a aquest registre.", vbInformation, "Atenció"
     datavinculats.Recordset.Edit
     datavinculats.Recordset!arxiuimp = True
     datavinculats.Recordset.Update
     datavinculats.Recordset.Move 0
  End If
  ratoli "normal"
  If Not datavinculats.Recordset!arxiuimp Then
      vimpnou = True
      ensenyarframenouimp
  End If
  
  ratoli "normal"
  'Me.Caption = nomfitxer
  obrir_document nomfitxer
  If Not vimpnou And InStr(1, UCase(formclixes.etestatclixe), "CLIXES ENTRATS") > 0 Then
       vmsg = InputBox("Si has fet algun canvi que creus que s'ha de notificar a IMPRESORES per tal de modificar els clixes escriu aquí els canvis per tal de passar un Email a Impresores.", "Canvis Imp")
       If atrim(vmsg) <> "" Then
         enviaremail "impresores@inplacsa.com", "Modificació del IMP del treball " + atrim(id_treball) + "/" + atrim(ordremodificacio) + "       [" + atrim(Now) + "]", "S'ha fet un canvi del IMP d'aquest treball revisar-ho si cal. " + vbNewLine + vbNewLine + "Observació de Disseny:" + vbNewLine + vmsg, nomfitxer
       End If
  End If
  If vimpnou Then enviaremail "qualitat@inplacsa.com", "Creació IMP " + Format(id_treball, "00000") + "\IMP" + Format(id_treball, "00000") + "-" + Format(ordremodificacio, "000") + "-" + Format(cadbl(datavinculats.Recordset!codiclient), "000000") + "_" + atrim(cadbl(datavinculats.Recordset!direnvio)), "Revisar el fitxer IMP"
  demanar_firmar_comandes
End Sub
Sub demanar_firmar_comandes()
  Dim i As Byte
  Dim vnumc As String
  Dim vfirmar As Boolean
  Dim rst As Recordset
  If formclixes.llistadecomandespendents.ListCount > 0 Then
    For i = 0 To formclixes.llistadecomandespendents.ListCount - 1
        vnumc = atrim(Mid(formclixes.llistadecomandespendents.List(i), 1, 7))
        Set rst = dbcomandes.OpenRecordset("select tipus from comandes_firmes where tipus='IM1' and comanda=" + atrim(vnumc))
        If rst.EOF Then vfirmar = True
    Next i
  End If
  
  If formclixes.llistadecomandespendents.ListCount > 0 And vfirmar Then
      If MsgBox("Hi ha comanda/s relacionades amb aquest IMP vols firmar-les com a IMP revisat?", vbExclamation + vbYesNo, "Firma comandes") = vbNo Then Exit Sub
        Else: Exit Sub
  End If
  For i = 0 To formclixes.llistadecomandespendents.ListCount - 1
    vnumc = atrim(Mid(formclixes.llistadecomandespendents.List(i), 1, 7))
    Set rst = dbcomandes.OpenRecordset("select tipus from comandes_firmes where tipus='IM1' and comanda=" + atrim(vnumc))
    If rst.EOF Then dbcomandes.Execute "insert into comandes_firmes (comanda,usuari,tipus,data) values (" + atrim(vnumc) + ",'" + nomordinador + "','IM1',now)"
  Next i
  Set rst = Nothing
End Sub
Function generarfitxer_imp_acopiar() As String
   If cadbl(nomclientnouimp.tag) > 0 And cadbl(direccioenvionouimp.tag) > 0 Then
      formclixes.crearruta ruta_documentacio_clixes + "\" + Format(campid_treball, "00000")
      generarfitxer_imp_acopiar = ruta_documentacio_clixes + "\" + Format(campid_treball, "00000") + "\IMP" + Format(campid_treball, "00000") + "-" + Format(campid_treball.tag, "000") + "-" + Format(nomclientnouimp.tag, "000000") + "_" + atrim(cadbl(direccioenvionouimp.tag)) + ".doc"
      If Not existeix(generarfitxer_imp_acopiar) Then generarfitxer_imp_acopiar = generarfitxer_imp_acopiar + "x"
     Else: generarfitxer_imp_acopiar = ""
   End If
End Function
Function generarfitxer_imp(Optional NodocX As Boolean) As String
      generarfitxer_imp = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\IMP" + Format(id_treball, "00000") + "-" + Format(ordremodificacio, "000") + "-" + Format(cadbl(datavinculats.Recordset!codiclient), "000000") + "_" + atrim(cadbl(datavinculats.Recordset!direnvio)) + ".doc"
      If Not NodocX Then If Not existeix(generarfitxer_imp) Then generarfitxer_imp = generarfitxer_imp + "x"
End Function
Sub ensenyarframenouimp()
  Dim rstn As Recordset
  Set rstn = datavinculats.Recordset.Clone
  fcrearimp.visible = True
  fcrearimp.Top = 660
  fcrearimp.Left = 1755
  campid_treball = atrim(id_treball)
  campid_treball.tag = atrim(ordremodificacio)
  nomclientnouimp = ""
  nomclientnouimp.tag = ""
  If rstn.RecordCount > 1 Then
     rstn.MoveFirst
     While Not rstn.EOF
       If rstn!arxiuimp Then
         nomclientnouimp = rstn!nomclient
         nomclientnouimp.tag = rstn!codiclient
       End If
       rstn.MoveNext
     Wend
  End If
End Sub
Private Sub Combo1_Change()

End Sub



Sub triar_direnvio_vinculats()
  If cadbl(datavinculats.Recordset!codiclient) < 1 Then MsgBox "Primer has d'escullir un client"
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select id ,domicilie,poblacioe,provinciae from clients_envios where codi=" + atrim(cadbl(datavinculats.Recordset!codiclient))
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(2).width = 900
  formseleccio.width = 9000
  formseleccio.Left = formseleccio.Left - 3000
  formseleccio.Show 1
  
   If seleccioret = 1 Then
        
        If Not formseleccio.Data1.Recordset.EOF Then
           direnvio = formseleccio.DBGrid2.Columns("poblacioe")
           datavinculats.Recordset!nomdirenvio = formseleccio.DBGrid2.Columns("poblacioe")
           datavinculats.Recordset!direnvio = cadbl(formseleccio.DBGrid2.Columns("id"))
           'campid_treball.Tag = cadbl(formseleccio.DBGrid2.Columns("ordremodificacio"))
        End If
   End If
    If seleccioret = 9 Then
         direnvio = ""
         datavinculats.Recordset!nomdirenvio = ""
           datavinculats.Recordset!direnvio = 0
        'campid_treball.Tag = ""
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   SendKeys "{TAB}"
   'codimuntadora.SetFocus
End Sub





Private Sub bsortida_Click()
   'generarminiaturapdf
   'If Framesortida.visible = False Then
   '    Framesortida.Top = 585
   '    Framesortida.Left = 30
   Unload formImpClient
   formImpClient.Show 1
       
       'Framesortida.visible = True
       '  Else: MsgBox "Fes guardar els canvis a la sortida de bobina que vol el client.", vbExclamation, "Atenció"
   'End If
End Sub
Private Sub Command1_Click()
   ratoli "espera"
  generarelnouimp
  fcrearimp.visible = False
  ratoli "normal"
  botoimp_Click
End Sub
Sub generarelnouimp()
  Dim origen As String
  Dim desti As String
  origen = generarfitxer_imp_acopiar
  desti = generarfitxer_imp(True)
  If InStr(1, origen, ".docx") > 0 Then desti = desti + "x"
  If origen = "" Then MsgBox "No hi ha cap client i direccio d'envio seleccionat.", vbCritical, "Atenció": Exit Sub
  If existeix(desti) Then MsgBox "Hi ha algun error en els fitxers IMP d'aquest treball i client. JA EXISTEIX EL FITXER IMP PER AQUEST CLIENT I TREBALL", vbCritical, "ATENCIÓ": Exit Sub
  If Not existeix(origen) Then MsgBox "No trobo el fitxer IMP per aquest treball i client que vols copiar.", vbCritical, "Atenció": Exit Sub
  'If InStr(1, origen, ".docx") = 0 Then MsgBox "EL FORMAT DEL FITXER ORIGINAL NO ESTÀ EN EL FORMAT WORD DE LA NOVA VERSIÓ, QUAN EL GUARDIS PENSA A CANVIAR-LO.", vbCritical, "ATENCIÓ"
  On Error GoTo errorcopiant
  Copiar_Fitxer origen, desti
  datavinculats.Recordset.Edit
  datavinculats.Recordset!arxiuimp = True
  datavinculats.Recordset.Update
  datavinculats.Recordset.Move 0
  enviaremail "qualitat@inplacsa.com", "Creació IMP " + Format(id_treball, "00000") + "\IMP" + Format(id_treball, "00000") + "-" + Format(ordremodificacio, "000") + "-" + Format(cadbl(datavinculats.Recordset!codiclient), "000000") + "_" + atrim(cadbl(datavinculats.Recordset!direnvio)), "Revisar el fitxer IMP"
  Exit Sub
errorcopiant:
   MsgBox err.Description
End Sub
Private Sub Command2_Click()
  fcrearimp.visible = False
End Sub
Sub possarunprincipal()
    Dim rst As Recordset
    Set rst = dbclixes.OpenRecordset("select principal from clientsvinculats where principal and id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
    If rst.EOF Then
       Set rst = dbclixes.OpenRecordset("select principal from clientsvinculats where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
       If Not rst.EOF Then
        rst.Edit
        rst!principal = True
        rst.Update
       End If
    End If
    Set rst = Nothing
End Sub


Private Sub Command7_Click()
  
End Sub

Private Sub Command8_Click()
 
End Sub

Private Sub cprincipal_Click()
   If Screen.ActiveControl.Name <> "cprincipal" Then Exit Sub
   If cprincipal.Value = 1 Then MsgBox "Per treure un client de principal has de marcar-ne un altre i aquest es treura sol", vbInformation, "Atenció": cprincipal.Value = 1: Exit Sub
    dbclixes.Execute "update clientsvinculats set principal=false where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio)
    dbclixes.Execute "update clientsvinculats set principal=true where id=" + atrim(datavinculats.Recordset!ID)
End Sub

Private Sub datavinculats_Reposition()
   If Not datavinculats.Recordset.EOF Then
      possarcolorbotoimp
      
   End If
End Sub
Sub actualitzarnovinculats()
     Dim rstc As Recordset
     Dim sies0 As String
     
     If datavinculats.Recordset.EOF Then Exit Sub
     sies0 = " or comandes.numordremodificacio=0 "
     If datavinculats.Recordset!ordremodificacio > 1 Then sies0 = ""
     ratoli "espera"
     While Not datavinculats.Recordset.EOF
'     MsgBox "SELECT Max(comandes.comanda) AS mcomanda, Count(comandes.numtreball) AS numcomandes from comandes WHERE (((comandes.numtreball)=" + atrim(id_treball) + ") and (comandes.numordremodificacio=" + atrim(datavinculats.Recordset!ordremodificacio) + sies0 + ") and comandes.client=" + atrim(datavinculats.Recordset!codiclient) + ");"
      Set rstc = dbcomandes.OpenRecordset("SELECT Max(comandes.comanda) AS mcomanda, Count(comandes.numtreball) AS numcomandes from comandes WHERE (((comandes.numtreball)=" + atrim(id_treball) + ") and (comandes.numordremodificacio=" + atrim(datavinculats.Recordset!ordremodificacio) + sies0 + ") and comandes.direnvio=" + atrim(datavinculats.Recordset!direnvio) + ");")
      If Not rstc.EOF Then
         datavinculats.Recordset.Edit
         datavinculats.Recordset!ultimacomanda = rstc!mcomanda
         datavinculats.Recordset!vegadesimpres = rstc!numcomandes
         datavinculats.Recordset.Update
      End If
      datavinculats.Recordset.MoveNext
      DoEvents
     Wend
     datavinculats.Recordset.MoveFirst
     ratoli "normal"
End Sub

Sub triar_client_direnvio()
  If cadbl(nomclientnouimp.tag) < 1 Then MsgBox "Primer has d'escullir un client"
 Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select id ,domicilie,poblacioe,provinciae from clients_envios where codi=" + atrim(cadbl(nomclientnouimp.tag))
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(2).width = 900
  formseleccio.width = 9000
  formseleccio.Left = formseleccio.Left - 3000
  formseleccio.Show 1
  
   If seleccioret = 1 Then
        
        If Not formseleccio.Data1.Recordset.EOF Then
           direccioenvionouimp = formseleccio.DBGrid2.Columns("poblacioe")
           direccioenvionouimp.tag = cadbl(formseleccio.DBGrid2.Columns("id"))
           'campid_treball.Tag = cadbl(formseleccio.DBGrid2.Columns("ordremodificacio"))
        End If
   End If
    If seleccioret = 9 Then
        direccioenvionouimp = ""
        direccioenvionouimp.tag = ""
        'campid_treball.Tag = ""
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   SendKeys "{TAB}"
   'codimuntadora.SetFocus
End Sub


Private Sub direccioenvionouimp_DropDown()
   triar_client_direnvio
End Sub

Private Sub direnvio_DropDown()
'triar_direnvio_vinculats
MsgBox "Per canviar el Client i direcció d'enviament has de crear un registre nou.", vbInformation, "Atenció"
End Sub

Private Sub eliminar_Click()
  If MsgBox("Segur que vols borrar aquest client vinculat?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
     If InputBox("Escriu la paraula [ELIMINAR] per fer efectiu l'eliminació", "Control de seguretat") = "ELIMINAR" Then
          datavinculats.Recordset.Delete
          datavinculats.Refresh
          If datavinculats.Recordset.EOF Then Exit Sub
          dbclixes.Execute "update clientsvinculats set principal=true where id=" + atrim(datavinculats.Recordset!ID)
          datavinculats.Refresh
          MsgBox "He passat el client: " + atrim(datavinculats.Recordset!nomclient) + " com a client principal.", vbInformation, "Atenció"
     End If
  End If
End Sub

Sub possarcolorbotoimp()
    If datavinculats.Recordset!arxiuimp Then
        botoimp.Picture = botoimp.DownPicture
        missatgenouimp.visible = False
          Else:
             botoimp.Picture = botoimp.DisabledPicture
             missatgenouimp.visible = True
    End If
    If existeix(substituir(formclixes.rutapdftreball, ".pdf", "_mini.gif")) Then
       bsortida.BackColor = &HF1B75F    'blau
         Else: bsortida.BackColor = guardar.BackColor 'color gris botó
    End If
End Sub

Private Sub Form_Load()
   datavinculats.DatabaseName = camiclixes
   datavinculats.RecordSource = "select * from clientsvinculats where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio) + " order by id"
   datavinculats.Refresh
   possarunprincipal
   actualitzarnovinculats
End Sub

Private Sub Framesortida_Click()
  'FileCopy imgpdf.tag, "c:\temp\prova.gif"
 ' TallarImatge "c:\temp\prova.gif", "c:\temp\prova.gif"
'  Set imgpdf = LoadPicture("c:\temp\prova.gif")
End Sub

Private Sub guardar_Click()
   If datavinculats.Recordset.EditMode = 0 Then Exit Sub
   On Error GoTo errors
   datavinculats.Recordset.Update
   datavinculats.Recordset.Bookmark = datavinculats.Recordset.LastModified
   Frame1.Enabled = False
   possarunprincipal
   Exit Sub
errors:
   MsgBox err.Description
   If Frame1.Enabled Then nomclient.SetFocus
End Sub

Private Sub Image2_Click()

End Sub

Private Sub imgpdf_DblClick()
   obrir_document imgpdf.tag
End Sub

Private Sub modificar_Click()
  Frame1.Enabled = True
  datavinculats.Recordset.Edit
  nomclient.SetFocus
End Sub

Private Sub nomclient_DropDown()
'  triar_client
  MsgBox "Per canviar el Client i direcció d'enviament has de crear un registre nou.", vbInformation, "Atenció"
End Sub
Function esclientrepetit(codi As Long) As Boolean
   Dim rst As Recordset
   esclientrepetit = False
   If codi = 0 Then Exit Function
   Set rst = datavinculats.Recordset.Clone
   If Not rst.EOF Then rst.MoveFirst
   rst.FindFirst "codiclient=" + atrim(cadbl(codi))
   If Not rst.NoMatch Then esclientrepetit = True
   
   Set rst = Nothing
End Function
Sub triar_client()
 Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select * from clients"
  formseleccio.refrescar
  formseleccio.Show 1
  
   If seleccioret = 1 Then
        If esclientrepetit(cadbl(formseleccio.DBGrid2.Columns("codi"))) Then MsgBox "Aquest client ja està vinculat", vbCritical, "Error": Exit Sub
        If Not formseleccio.Data1.Recordset.EOF Then
           nomclient = formseleccio.DBGrid2.Columns("nom")
           datavinculats.Recordset!codiclient = cadbl(formseleccio.DBGrid2.Columns("codi"))
        End If
   End If
    If seleccioret = 9 Then
        client = ""
        datavinculats.Recordset!codiclient = Null
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   SendKeys "{TAB}"
   'codimuntadora.SetFocus
End Sub
Sub triar_client_imp()
 Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = camiclixes
  formseleccio.Data1.RecordSource = "select ordremodificacio,codiclient,nomclient from clientsvinculats where id_treball=" + atrim(cadbl(campid_treball)) + " and ordremodificacio=(select max(ordremodificacio) from clientsvinculats where arxiuimp and id_treball=" + atrim(cadbl(campid_treball)) + ")"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(2).width = 1200
  formseleccio.Show 1
  
   If seleccioret = 1 Then
        
        If Not formseleccio.Data1.Recordset.EOF Then
           nomclientnouimp = formseleccio.DBGrid2.Columns("nomclient")
           nomclientnouimp.tag = cadbl(formseleccio.DBGrid2.Columns("codiclient"))
           campid_treball.tag = cadbl(formseleccio.DBGrid2.Columns("ordremodificacio"))
        End If
   End If
    If seleccioret = 9 Then
        nomclientnouimp = ""
        nomclientnouimp.tag = ""
        campid_treball.tag = ""
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   SendKeys "{TAB}"
   'codimuntadora.SetFocus
End Sub

Private Sub nomclientnouimp_DropDown()
  triar_client_imp
End Sub

Private Sub sortir_Click()
  'If Framesortida.visible Then MsgBox "S'han de guardar els canvis de la sortida del PDF.", vbCritical, "Error": Exit Sub
  Unload Me
End Sub

