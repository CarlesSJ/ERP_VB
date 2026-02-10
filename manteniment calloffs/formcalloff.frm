VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form formcalloff 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniment de Call-off"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16110
   Icon            =   "formcalloff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   16110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   60
      Top             =   255
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   210
      TabIndex        =   4
      Top             =   30
      Width           =   7605
      Begin VB.ComboBox comboclient 
         Height          =   315
         Left            =   1515
         TabIndex        =   5
         Top             =   255
         Width           =   5985
      End
      Begin VB.Label Label2 
         Caption         =   "Escull el Client:"
         Height          =   300
         Left            =   180
         TabIndex        =   13
         Top             =   285
         Width           =   1530
      End
   End
   Begin VB.Frame framedades 
      Caption         =   "Dades Call-off"
      Enabled         =   0   'False
      Height          =   6105
      Left            =   180
      TabIndex        =   0
      Top             =   795
      Width           =   15795
      Begin VB.TextBox citem 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   615
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   210
         Width           =   1440
      End
      Begin VB.CheckBox checkamagaentregades 
         Caption         =   "Amaga comandes entregades"
         Height          =   225
         Left            =   4260
         TabIndex        =   18
         Top             =   285
         Value           =   1  'Checked
         Width           =   3645
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H008080FF&
         Caption         =   "Treure Call-Off  -->"
         Height          =   420
         Left            =   5865
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4140
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "<-- Assignar Call-Off"
         Height          =   420
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3615
         Width           =   1575
      End
      Begin VB.CommandButton alta 
         Height          =   420
         Left            =   45
         Picture         =   "formcalloff.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   510
         Width           =   450
      End
      Begin VB.CommandButton consultar 
         Height          =   420
         Left            =   45
         Picture         =   "formcalloff.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Buscar Registres"
         Top             =   945
         Width           =   450
      End
      Begin MSFlexGridLib.MSFlexGrid reixac 
         Height          =   2475
         Left            =   495
         TabIndex        =   6
         Top             =   540
         Width           =   15180
         _ExtentX        =   26776
         _ExtentY        =   4366
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.Data dataitems 
         Caption         =   "dataitems"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   390
         Left            =   330
         Options         =   0
         ReadOnly        =   -1  'True
         RecordsetType   =   2  'Snapshot
         RecordSource    =   ""
         Top             =   1575
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Frame Frame2 
         Caption         =   "Detall de l'Item"
         Height          =   2730
         Left            =   7560
         TabIndex        =   2
         Top             =   3165
         Width           =   8145
         Begin VB.CommandButton Command3 
            Height          =   375
            Left            =   30
            Picture         =   "formcalloff.frx":109E
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Refrescar calloffs assignats"
            Top             =   1230
            Width           =   405
         End
         Begin VB.CommandButton modificar 
            Height          =   375
            Left            =   30
            Picture         =   "formcalloff.frx":1628
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Modificar Registres"
            Top             =   705
            Width           =   405
         End
         Begin VB.Data datadetall 
            Caption         =   "datadetall"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   390
            Left            =   2985
            Options         =   0
            ReadOnly        =   -1  'True
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   915
            Visible         =   0   'False
            Width           =   2280
         End
         Begin VB.CommandButton btreurecomanda 
            Height          =   375
            Left            =   45
            Picture         =   "formcalloff.frx":1BB2
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Eliminacio Registres"
            Top             =   285
            Width           =   405
         End
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "formcalloff.frx":213C
            Height          =   2220
            Left            =   495
            OleObjectBlob   =   "formcalloff.frx":2151
            TabIndex        =   7
            Top             =   270
            Width           =   7545
         End
      End
      Begin MSDBGrid.DBGrid reixaitems 
         Bindings        =   "formcalloff.frx":3040
         Height          =   2550
         Left            =   510
         OleObjectBlob   =   "formcalloff.frx":3054
         TabIndex        =   1
         Top             =   510
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.ListBox llistadepalets 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2340
         Left            =   45
         TabIndex        =   8
         Top             =   3375
         Width           =   5790
      End
      Begin VB.Label Label3 
         Caption         =   "Item:"
         Height          =   240
         Left            =   210
         TabIndex        =   20
         Top             =   270
         Width           =   540
      End
      Begin VB.Label etnumcomanda 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4530
         TabIndex        =   16
         Top             =   3030
         Width           =   1950
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Palet    ------->  Quan. Produïda/Assignada     Nº Call-Off"
         Height          =   285
         Left            =   75
         TabIndex        =   9
         Top             =   3135
         Width           =   4320
      End
   End
End
Attribute VB_Name = "formcalloff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
  afegir_calloff
End Sub
Sub afegir_calloff()
   Dim vnumcalloff As String
   Dim vdata As String
   Dim vitem As String
   Dim vdemanats As String
   Dim vnumcontracte As String
   Dim rstc As Recordset
   If cadbl(comboclient.Tag) = 0 Then MsgBox "Primer has d'escullir el nom del client", vbCritical, "Error": Exit Sub
   'If Not dataitems.Recordset.EOF Then vitem = atrim(dataitems.Recordset!Item)
   vitem = citem
   vitem = InputBox("Entra l'Item demanat.", "Item", vitem)
  ' dataitems.Recordset.FindFirst "item='" + atrim(vitem) + "'"
  
   carregar_item vitem
   vnumcalloff = UCase(InputBox("Entra el Nº de Call-Off:", "Nou Call-Off", vnumcalloff))
   If vnumcalloff = "" Then Exit Sub
   If Mid(vnumcalloff + "   ", 1, 2) <> "45" Then
         If MsgBox("Aquest numero de Call-off no comença per 45" + vbNewLine + "Es correcte?", vbCritical + vbDefaultButton2 + vbYesNo, "Contracte") = vbNo Then Exit Sub
   End If
   Set rstc = datadetall.Database.OpenRecordset("select * from calloffs where item='" + atrim(vitem) + "' and numcalloff='" + atrim(vnumcalloff) + "'")
   If Not rstc.EOF Then MsgBox "Aquest calloff ja existeix en aquesta referencia", vbCritical, "Error": Exit Sub
   If Not datadetall.Recordset.EOF Then
      datadetall.Recordset.FindFirst "numcalloff='" + atrim(vnumcalloff) + "'"
        Else
          vdata = InputBox("Entra la data del Call-Off", "Data")
          If Not IsDate(vdata) Then MsgBox "Aquesta data no es vàlida.", vbInformation, "Atenció": Exit Sub
   End If
   If datadetall.Recordset.NoMatch Then
      vdata = InputBox("Entra la data del Call-Off", "Data")
      If Not IsDate(vdata) Then MsgBox "Aquesta data no es vàlida.", vbInformation, "Atenció": Exit Sub
        Else:
          If Not datadetall.Recordset.EOF Then
           vdata = datadetall.Recordset!Data
          End If
   End If
     'ara ja no volen que es demani el numero de contracte, l'afegirè automàticament quan assignin el calloff a la comanda
   'vnumcontracte = UCase(InputBox("Entra el Nº de Contracte:", "Contracte"))
   'If vnumcontracte <> "" Then
   '  If Mid(vnumcontracte + "   ", 1, 2) <> "46" Then
   '      If MsgBox("Aquest numero de contracte no comença per 46" + vbNewLine + "Es correcte?", vbCritical + vbDefaultButton2 + vbYesNo, "Contracte") = vbNo Then Exit Sub
   '  End If
   'End If
   
   vdemanats = cadbl(InputBox("Entra la quantitat demanada pel client.", "Quantitat"))
   If vdemanats = 0 Then MsgBox "Error de quantitat demanada.", vbCritical, "Error": Exit Sub
   With datadetall.Recordset
   datadetall.Database.Execute "insert into calloffs (client,numcalloff,data,item,demanats,contracte) values (" + atrim(cadbl(comboclient.Tag)) + ",'" + vnumcalloff + "',#" + Format(vdata, "mm/dd/yy") + "#,'" + atrim(vitem) + "'," + atrim(vdemanats) + ",'" + vnumcontracte + "')"
   End With
   carregar_item vitem
   'dataitems.Refresh
   'dataitems.Recordset.FindFirst "item='" + vitem + "'"
   datadetall.Recordset.FindFirst "numcalloff='" + vnumcalloff + "'"
End Sub



Sub triarclient()
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = dataitems.DatabaseName
  formseleccio.Data1.RecordSource = "select * from clients"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   comboclient.Tag = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   comboclient.Text = atrim(formseleccio.Data1.Recordset!nom)
  End If
  Unload formseleccio
  
End Sub


Private Sub btreurecomanda_Click()
   Dim vitem As String
   Dim vnumcalloff As String
   Dim i As Integer
   Dim i2 As Integer
   If datadetall.Recordset.EOF Then Exit Sub
   vitem = atrim(citem)
   vnumcalloff = atrim(datadetall.Recordset!numcalloff)
   If MsgBox("Segur que vols eliminar el call-off " + vnumcalloff + "?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   datadetall.Recordset.Delete
   reixac.col = 0
   i2 = reixac.row
   For i = 1 To reixac.Rows - 1
     reixac.row = i
     dbtmp.Execute "update bobinesent set numcalloff='' where (entregat<>'S' or entregat=null) and comanda=" + atrim(cadbl(reixac.Text)) + " and numcalloff='" + atrim(vnumcalloff) + "'"
   Next i
   dbtmp.Execute "delete * from calloffs_detall where item='" + atrim(vitem) + "' and numcalloff='" + atrim(vnumcalloff) + "'"
   'dataitems.Refresh
   'dataitems.Recordset.FindFirst "item='" + atrim(vitem) + "'"
   carregar_item citem
   reixac.row = i2
End Sub

Private Sub checkamagaentregades_Click()
   carregar_item citem
End Sub

Private Sub comboclient_DropDown()
   triarclient
   'ensenyaritemsdelclient
End Sub
Sub ensenyaritemsdelclient()
   Me.Caption = "Escullint Items diferents"
   'dataitems.RecordSource = "select distinct item from calloffs where client=" + atrim(cadbl(comboclient.Tag)) + " order by item"
  ' dataitems.RecordSource = "select  item from calloffs where client=" + atrim(cadbl(comboclient.Tag)) + " order by item"

   dataitems.Refresh
   framedades.Enabled = True
  
   Me.Caption = "Manteniment de Call-off"
End Sub


Function tecontractelacomanda(vnumc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select comandaclient from comandes where comanda=" + atrim(vnumc), , ReadOnly)
   If rst.EOF Then Exit Function
   tecontractelacomanda = IIf(atrim(rst!comandaclient) = "", False, True)
   Set rst = Nothing
End Function
Private Sub Command1_Click()
   Dim vnompalet As String
   Dim vnumc As Double
   Dim vdemanat As Double
   Dim vitem As String
   Dim vpalet As Integer
   Dim vnumcontracte As String
   Dim vnumdecalloff As String
   
   If datadetall.Recordset.EOF Then MsgBox "Escull un Call-Off per assignar al palet o comanda", vbCritical, "Atenció": Exit Sub
   If cadbl(datadetall.Recordset!demanats) <= cadbl(datadetall.Recordset!assignats) Then MsgBox "Aquest call-off ja te assignades les peces que es demanaven", vbCritical, "Atenció"
   If llistadepalets.ListIndex = -1 Then MsgBox "Primer escull un valor de la llista de palets.", vbCritical, "Atenció": Exit Sub
   If llistadepalets.ItemData(llistadepalets.ListIndex) > 0 And llistadepalets.ItemData(llistadepalets.ListIndex) <> -2 Then If MsgBox("Aquesta linia ja te una assignació, vols substituir-la?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   vnumc = cadbl(reixac.TextMatrix(reixac.row, 0))
   vdemanat = cadbl(reixac.TextMatrix(reixac.row, 5))
   vitem = atrim(citem)
   If Mid(llistadepalets.Text, 1, 7) = "Comanda" Then
      vnompalet = " tota la comanda"
        Else:
           vnompalet = " aquest palet"
           vpalet = cadbl(Mid(llistadepalets.Text, 1, 3))
           If vpalet < 1 Then Exit Sub
   End If
   If Not tecontractelacomanda(vnumc) Then MsgBox "Aquesta comanda no té el contracte entrat, primer hauries d'anar a entrar-lo.", vbCritical, "Error": GoTo fi
   If MsgBox("Segur que vols assignar el Call-Off " + atrim(datadetall.Recordset!numcalloff) + " a " + vnompalet + "?", vbInformation + vbYesNo + vbDefaultButton2, "Assignar Call-Off") = vbNo Then Exit Sub
   enviarinformacioassignaciodecalloffaimpresoresitintes vnumc, vdemanat, vitem, vnumcontracte
   'vnumcontracte = concatenartotselscontractesambaquesITEM
   datadetall.Database.Execute "update calloffs set contracte='" + vnumcontracte + "' where item='" + atrim(vitem) + "' and numcalloff='" + atrim(datadetall.Recordset!numcalloff) + "'"
   'dbtmp.Execute "update comandes set comandaclient='" + atrim(datadetall.Recordset!contracte) + "' where comanda=" + atrim(vnumc)
   If vnompalet = " tota la comanda" Then
      datadetall.Database.Execute "delete * from calloffs_detall where comanda=" + atrim(vnumc) + " and item='" + atrim(vitem) + "' "
      'MsgBox "delete * from calloffs_detall where comanda=" + atrim(vnumc) + " and item='" + atrim(vitem) + "' "
'      MsgBox "Insert into calloffs_detall (numcalloff,item,comanda,assignats) values('" + atrim(datadetall.Recordset!numcalloff) + "','" + vitem + "'," + atrim(vnumc) + "," + atrim(vdemanat) + ")"
      datadetall.Database.Execute "Insert into calloffs_detall (numcalloff,item,comanda,assignats) values('" + atrim(datadetall.Recordset!numcalloff) + "','" + vitem + "'," + atrim(vnumc) + "," + atrim(vdemanat) + ")"
     ' MsgBox "Insert into calloffs_detall (numcalloff,item,comanda,assignats) values('" + atrim(datadetall.Recordset!numcalloff) + "','" + vitem + "'," + atrim(vnumc) + "," + atrim(vdemanat) + ")"
      datadetall.Database.Execute "update bobinesent set numcalloff='' where (entregat<>'S' or entregat=null) and comanda=" + atrim(vnumc)
        Else
          datadetall.Database.Execute "update bobinesent set numcalloff='" + atrim(datadetall.Recordset!numcalloff) + "' where comanda=" + atrim(vnumc) + " and numpalet=" + atrim(vpalet) + " and (entregat<>'S' or entregat=null)"
          datadetall.Database.Execute "delete * from calloffs_detall where comanda=" + atrim(vnumc) + " and item='" + atrim(vitem) + "' and numcalloff='" + atrim(datadetall.Recordset!numcalloff) + "'"
   End If
'   actualitzar_valors_decalloff
   carregar_paletsicalloffs
   emplenarllistadecomandesdelitem atrim(citem), vnumc
   vnumdecalloff = atrim(datadetall.Recordset!numcalloff)
   datadetall.Refresh
   datadetall.Recordset.FindFirst "numcalloff='" + vnumdecalloff + "'"
   MsgBox "Call-Off assignat", vbInformation, "Call-Off"
fi:
End Sub
Sub enviarinformacioassignaciodecalloffaimpresoresitintes(vnumc As Double, vdemanat As Double, vitem As String, vnumcontracte As String)
  Dim rst As Recordset
  Dim vlinia As String
  Dim rstc As Recordset
  Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
  If rst.EOF Then Exit Sub
  vnumcontracte = atrim(rst!comandaclient)
  If atrim(rst!proximaseccio) <> "E" And atrim(rst!proximaseccio) <> "I" Then Exit Sub
  Set rstc = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(cadbl(rst!numtreball)))
  If rstc.EOF Then
     vlinia = ""
       Else: vlinia = atrim(rstc!marca) + " - " + atrim(rstc!linia)
  End If
   'avisaralentraruncalloff
  enviaremailgeneric "avisaralentraruncalloff", "CALL-OFF nou entrat.", "Comanda: " + atrim(vnumc) + "  " + vlinia + "  RefClient: " + atrim(vitem) + "  " + atrim(vdemanat)
   
  Set rstc = Nothing
  Set rst = Nothing
End Sub
Sub actualitzar_valors_decalloff()
  Dim vassignats As Double
  Dim rst As Recordset
  Dim rstc As Recordset
  Dim rstcom As Recordset
  Dim vnumc As Double
  Dim vdesarroll As Double
  Dim vitem As String
  Dim vbookmark As Double
  vnumc = cadbl(reixac.TextMatrix(reixac.row, 0))
  vdesarroll = buscardesarroll(vnumc)
  If vdesarroll = 0 Then Exit Sub
  ratoli "espera"
  vitem = atrim(citem)
  If Not datadetall.Recordset.EOF Then vbookmark = cadbl(datadetall.Recordset![ID])
  datadetall.Refresh
  While Not datadetall.Recordset.EOF
   'posso a zero els assignats per tornar a comptarlos
   datadetall.Recordset.Edit
   datadetall.Recordset!assignats = 0
   datadetall.Recordset.Update
   
   Set rstc = dbtmp.OpenRecordset("select distinct comanda from bobinesent where numcalloff='" + atrim(datadetall.Recordset!numcalloff) + "'", , ReadOnly)
   While Not rstc.EOF
    Set rstcom = dbtmp.OpenRecordset("select refclient from comandes where comanda=" + atrim(rstc!comanda))
    If InStr(1, atrim(rstcom!refclient), " " + vitem) Or InStr(1, atrim(rstcom!refclient), vitem + " ") Then
      Set rst = dbtmp.OpenRecordset("select sum(metresisacs) as tmetres from bobinesent where  comanda=" + atrim(cadbl(rstc!comanda)) + " and numcalloff='" + atrim(datadetall.Recordset!numcalloff) + "' group by comanda", , ReadOnly)
      If Not rst.EOF Then
        datadetall.Recordset.Edit
        datadetall.Recordset!assignats = datadetall.Recordset!assignats + Redondejar(cadbl(rst!tmetres) / (vdesarroll / 1000), 0)
        datadetall.Recordset.Update
        datadetall.Database.Execute "delete * from calloffs_detall where comanda=" + atrim(rstc!comanda) + " and item='" + atrim(vitem) + "'" + " and numcalloff='" + atrim(datadetall.Recordset!numcalloff) + "'"
      End If
      Set rst = dbtmp.OpenRecordset("select * from calloffs_detall where comanda=" + atrim(rstc!comanda) + " and item='" + atrim(vitem) + "' and numcalloff='" + atrim(datadetall.Recordset!numcalloff) + "'")
      If Not rst.EOF Then
        datadetall.Recordset.Edit
        datadetall.Recordset!assignats = datadetall.Recordset!assignats + cadbl(rst!assignats)
        datadetall.Recordset.Update
      End If
    End If
    rstc.MoveNext
   Wend
   datadetall.Recordset.MoveNext
  Wend
  If vbookmark > 0 Then datadetall.Recordset.FindFirst "id=" + atrim(vbookmark)
  Set rst = Nothing
  Set rstcom = Nothing
  Set rstc = Nothing
  ratoli "normal"
End Sub
Function pecesassignadesacalloff(vnumc As Double, vitem As String, vdesarroll As Double) As Double
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select sum(metresisacs) as tmetres from bobinesent where  (numcalloff<>'' and numcalloff<>null) and comanda=" + atrim(vnumc) + "  group by comanda", dbOpenSnapshot, dbReadOnly)
  If Not rst.EOF Then
       pecesassignadesacalloff = Redondejar(cadbl(rst!tmetres) / (vdesarroll / 1000), 0)
  End If
  Set rst = dbtmp.OpenRecordset("select * from calloffs_detall where comanda=" + atrim(vnumc) + " and item='" + atrim(vitem) + "'", dbOpenSnapshot, dbReadOnly)
  If Not rst.EOF Then
       pecesassignadesacalloff = cadbl(rst!assignats)
  End If
  pecesassignadesacalloff = Redondejar(pecesassignadesacalloff, 0)
  Set rst = Nothing
End Function

Private Sub Command2_Click()
Dim vnompalet As String
   Dim vnumc As Double
   Dim vdemanat As Double
   Dim vitem As String
   Dim vpalet As Integer
   If llistadepalets.ListIndex = -1 Then MsgBox "Primer escull un valor de la llista de palets.", vbCritical, "Atenció": Exit Sub
   'If datadetall.Recordset.EOF Then MsgBox "Escull un Call-Off per assignar al palet o comanda", vbCritical, "Atenció": Exit Sub
   If InStr(1, llistadepalets.Text, "Call-Off:") = 0 Then MsgBox "Aquesta linia no te cap assignació de Call-off", vbExclamation, "Atenció": Exit Sub
   vnumc = cadbl(reixac.TextMatrix(reixac.row, 0))
   vdemanat = cadbl(reixac.TextMatrix(reixac.row, 5))
   vitem = atrim(citem)
   If Mid(llistadepalets.Text, 1, 7) = "Comanda" Then
      vnompalet = " tota la comanda"
        Else:
           vnompalet = " aquest palet"
           vpalet = cadbl(Mid(llistadepalets.Text, 1, 3))
           If vpalet < 1 Then Exit Sub
   End If
   'If MsgBox("Segur que vols alliberar el Call-Off " + atrim(datadetall.Recordset!numcalloff) + " a " + vnompalet + "?", vbInformation + vbYesNo + vbDefaultButton2, "Assignar Call-Off") = vbNo Then Exit Sub
   vnumdecalloff = treure_Calloff(llistadepalets.Text)
   datadetall.Recordset.FindFirst "numcalloff='" + vnumdecalloff + "'"
   If MsgBox("Segur que vols alliberar el Call-Off " + vnumdecalloff + " a " + vnompalet + "?", vbInformation + vbYesNo + vbDefaultButton2, "Assignar Call-Off") = vbNo Then Exit Sub
   'If MsgBox("Vols borrar també el número de contracte assignat a la comanda?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
   '   dbtmp.Execute "update comandes set comandaclient='' where comanda=" + atrim(vnumc)
   'End If
   If vnompalet = " tota la comanda" Then
      datadetall.Database.Execute "delete * from calloffs_detall where comanda=" + atrim(vnumc) + " and item='" + atrim(vitem) + "' "
      'datadetall.Database.Execute "Insert into calloffs_detall (numcalloff,item,comanda,assignats) values('" + atrim(datadetall.Recordset!numcalloff) + "','" + vitem + "'," + atrim(vnumc) + "," + atrim(vdemanat) + ")"
        Else
          datadetall.Database.Execute "update bobinesent set numcalloff='' where comanda=" + atrim(vnumc) + " and numpalet=" + atrim(vpalet) + " and (entregat<>'S' or entregat=null)"
          datadetall.Database.Execute "delete * from calloffs_detall where comanda=" + atrim(vnumc) + " and item='" + atrim(vitem) + "' and numcalloff='" + atrim(vnumdecalloff) + "'"
   End If
   datadetall.Database.Execute "update calloffs set contracte='' where item='" + atrim(vitem) + "' and numcalloff='" + atrim(datadetall.Recordset!numcalloff) + "'"
   datadetall.Refresh
   datadetall.Recordset.FindFirst "numcalloff='" + vnumdecalloff + "'"
   'actualitzar_valors_decalloff
   carregar_paletsicalloffs
   emplenarllistadecomandesdelitem atrim(vitem), vnumc
   MsgBox "Call-Off alliberat", vbInformation, "Call-Off"
   
End Sub
Function treure_Calloff(ByVal v As String) As String
    If InStr(1, v, "Call-Off:") > 0 Then
        treure_Calloff = substituir(v, Mid(v, 1, InStr(1, v, "Call-Off:") + 9), "")
        treure_Calloff = substituir(treure_Calloff, "Entregat", "")
    End If
End Function

Private Sub Command3_Click()
actualitzar_valors_decalloff
End Sub

Private Sub consultar_Click()
  Dim vitem As String
  vitem = InputBox("Entra la referencia que vols buscar.", "Buscar")
  'dataitems.Recordset.FindFirst "item='" + atrim(vitem) + "'"
  If atrim(vitem) = "" Then MsgBox "Referencia no valida.", vbCritical, "Error": Exit Sub
  citem = vitem
  carregar_item citem
  
End Sub

Private Sub dataitems_Reposition()
   If Not dataitems.Recordset.EOF Then
       carregar_item dataitems.Recordset!Item
          Else
            datadetall.RecordSource = "select * from calloffs where item='' order by numcalloff"
            datadetall.Refresh
   End If
End Sub
Sub carregar_item(vitem As String)
   If atrim(vitem) = "" Then Exit Sub
   citem = vitem
   datadetall.RecordSource = "select id,numcalloff,data,contracte,demanats,assignats from calloffs where item='" + atrim(vitem) + "' order by numcalloff "
   datadetall.Refresh
   Me.Caption = "Emplenant llista de comandes de l'Item " + atrim(vitem)
   emplenarllistadecomandesdelitem atrim(vitem)
   Me.Caption = "Emplenant palets i calloffs"
   carregar_paletsicalloffs
   Me.Caption = "Manteniment de Call-off"
   If Not datadetall.Recordset.EOF Then datadetall.Recordset.MoveLast
End Sub
Private Sub eliminar_Click()

End Sub

Private Sub Form_Load()
   dataitems.DatabaseName = cami
   datadetall.DatabaseName = cami
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb", , True)
   configurar_reixac
   comboclient.Tag = "6841"
   comboclient.Text = "CROP'S NV"
   Timer1.Interval = 2000
   Timer1.Enabled = True
End Sub
Sub configurar_reixac()
  Dim amplades As Variant
  Dim i As Byte
  reixac.Clear
  amplades = Array(1000, 1000, 1000, 4000, 600, 850, 850, 850, 850, 850)
   reixac.Rows = 0
  reixac.AddItem "Comanda" + Chr(9) + "Data" + Chr(9) + "Treball/V" + Chr(9) + "Linia+Texte" + Chr(9) + "Preu" + Chr(9) + "Demanat T" + Chr(9) + "Produït" + Chr(9) + "Entregat" + Chr(9) + "Pendent" + Chr(9) + "Disponible"
  For i = 0 To reixac.Cols - 1
     reixac.ColWidth(i) = amplades(i)
  Next i
  reixac.Rows = 2
  reixac.FixedRows = 1
  reixac.Rows = 1
End Sub


Function justificar(v As String, longitut As Integer, DoE As String) As String
    v = Mid(v, 1, longitut)
    If DoE = "E" Then
       v = v + Space(longitut - Len(v))
      Else: v = Space(longitut - Len(v)) + v
    End If
    justificar = v
End Function

Sub emplenarllistadecomandesdelitem(vitem As String, Optional vnumc As Double)
   Dim rst As Recordset
   Dim rstt As Recordset
   Dim vacabades As Boolean
   Dim i As Byte
   configurar_reixac
   'comandes NO acabades d'aquest item
   Set rstt = dbtmp.OpenRecordset("select * from comandes where refclient = '" + vitem + "' and client=" + atrim(cadbl(comboclient.Tag)) + " order by datacomanda desc", dbOpenSnapshot, dbReadOnly) ' + " and proximaseccio<>'T' order by datacomanda desc"
   If rstt.EOF Then GoTo fi
   rstt.Filter = "proximaseccio<>'T'"
   Set rst = rstt.OpenRecordset
   vacabades = False
carregarcomandes:
   i = 0
   While Not rst.EOF And i < 5
     If rst!producte <> "PC" And rst!producte <> "PC2" And rst!producte <> "PCP" Then
        'If (InStr(1, " " + atrim(rst!refclient), " " + vitem) Or InStr(1, atrim(rst!refclient) + " ", vitem + " ")) And atrim(rst!refclient) <> "" Then
          carregarcomandaareixac rst, vacabades
          If vacabades Then i = i + 1
          DoEvents
        'End If
     End If
     rst.MoveNext
   Wend
   If checkamagaentregades.Value = 0 Then
        'comandes SI acabades d'aquest item
        If Not vacabades Then
          vacabades = True
          'Set rst = dbtmp.OpenRecordset("select * from comandes where refclient like '*" + vitem + "*' and client=" + atrim(cadbl(comboclient.Tag)) + " and proximaseccio='T' order by comanda DESC", dbOpenSnapshot, dbReadOnly)
          rstt.Filter = "proximaseccio='T'"
          Set rst = rstt.OpenRecordset
          GoTo carregarcomandes
        End If
   End If
   If reixac.Rows > 1 Then
       reixac.row = 1
       reixac.RowSel = 1
       reixac.col = 0
       reixac.ColSel = reixac.Cols - 1
   End If
   If vnumc > 0 Then
     i = 1
     While i < reixac.Rows
        If reixac.TextMatrix(i, 0) = atrim(vnumc) Then
           reixac.row = i
           reixac.RowSel = i
           reixac.col = 0
           reixac.ColSel = reixac.Cols - 1
         End If
        i = i + 1
     Wend
   End If
fi:
   Set rst = Nothing
   
End Sub
Sub carregarcomandaareixac(rstc As Recordset, vacabades As Boolean)
  Dim vnumc As Double
  Dim vrow As Integer
  Dim rstent As Recordset
  Dim vdesarroll As Double
  Dim vpecestotal As Double
  vnumc = cadbl(rstc!comanda)
  reixac.AddItem atrim(vnumc)
  'Set rstc = dbtmp.OpenRecordset("Select * from comandes where comanda=" + atrim(vnumc), dbOpenSnapshot, dbReadOnly)
  'If rstc.EOF Then GoTo fi
  If vacabades Then
        reixac.row = reixac.Rows - 1
        For i = 0 To reixac.Cols - 1
          reixac.col = i
          reixac.CellBackColor = QBColor(12)
        Next i
  End If
  vrow = reixac.Rows - 1
  reixac.TextMatrix(vrow, 1) = Format(rstc!datacomanda, "dd/mm/yy")
  reixac.TextMatrix(vrow, 2) = atrim(rstc!numtreball) + "/" + atrim(rstc!numordremodificacio)
  reixac.TextMatrix(vrow, 3) = atrim(rstc!marcailinia)
  reixac.TextMatrix(vrow, 4) = atrim(rstc!pvp)
  reixac.TextMatrix(vrow, 5) = atrim(cadbl(rstc!tubbaseext))
  
  vdesarroll = buscardesarroll(vnumc)
  If vdesarroll = 0 Then GoTo fi
  Set rstent = dbtmp.OpenRecordset("select sum(metresisacs) as tmetres from bobinesent where  comanda=" + atrim(cadbl(vnumc)) + " group by comanda", dbOpenSnapshot, dbReadOnly)
  If Not rstent.EOF Then reixac.TextMatrix(vrow, 6) = atrim(Redondejar(cadbl(rstent!tmetres) / (vdesarroll / 1000), 0)) 'produit
  
  Set rstent = dbtmp.OpenRecordset("select sum(metresisacs) as tmetres  from bobinesent where (entregat='S') and comanda=" + atrim(vnumc) + " group by numpalet", dbOpenSnapshot, dbReadOnly)
  If Not rstent.EOF Then reixac.TextMatrix(vrow, 7) = atrim(Redondejar(cadbl(rstent!tmetres) / (vdesarroll / 1000), 0)) 'entregat
  Set rstent = dbtmp.OpenRecordset("select sum(metresisacs) as tmetres from bobinesent where (entregat='N' or entregat='' or entregat=null)  and comanda=" + atrim(cadbl(vnumc)) + " group by comanda", dbOpenSnapshot, dbReadOnly)
  If Not rstent.EOF Then reixac.TextMatrix(vrow, 8) = atrim(Redondejar(cadbl(rstent!tmetres) / (vdesarroll / 1000), 0)) 'pendent
  vpecestotal = cadbl(reixac.TextMatrix(vrow, 6))
  If vpecestotal = 0 Then vpecestotal = cadbl(reixac.TextMatrix(vrow, 5))
  reixac.TextMatrix(vrow, 9) = vpecestotal - pecesassignadesacalloff(vnumc, citem, vdesarroll) 'disponible
fi:
  'Set rstc = Nothing
  Set rstent = Nothing
End Sub
Private Sub reixacalloffs_Click()

End Sub

Private Sub llistacomandesdisponibles_Click()

End Sub

Private Sub llistadepalets_Click()
   Dim vestat As Boolean
   If InStr(1, llistadepalets, "Entregat") > 0 Then
       vestat = False
        Else
          vestat = True
   End If
   Command1.Enabled = vestat
   Command2.Enabled = vestat
End Sub

Private Sub modificar_Click()
   Dim vdata As String
   Dim vdemanats As Double
   vdata = InputBox("Entra la data del Call-Off", "Data")
   If Not IsDate(vdata) Then MsgBox "Aquesta data no es vàlida.", vbInformation, "Atenció": Exit Sub
   vdemanats = cadbl(InputBox("Entra la quantitat demanada pel client.", "Quantitat"))
   If vdemanats = 0 Then MsgBox "Error de quantitat demanada.", vbCritical, "Error": Exit Sub
   datadetall.Recordset.Edit
   datadetall.Recordset!Data = CVDate(vdata)
   datadetall.Recordset!demanats = vdemanats
   datadetall.Recordset.Update
End Sub

Private Sub reixac_Click()
   carregar_paletsicalloffs
End Sub
Sub carregar_paletsicalloffs()
 Dim rstent As Recordset
  Dim vdesarroll As Double
  Dim vnumc As Double
  Dim ventregatono As String
  llistadepalets.Clear
  llistadepalets.BackColor = QBColor(15)
  vnumc = cadbl(reixac.TextMatrix(reixac.row, 0))
  etnumcomanda = "Lot: " + atrim(vnumc)
  'ventregatono = "(entregat='N' or entregat='' or entregat=null) and "
  If reixac.CellBackColor = QBColor(12) Then
     llistadepalets.BackColor = QBColor(12): ventregatono = ""
     Command1.Enabled = False: Command2.Enabled = False
       Else: Command1.Enabled = True: Command2.Enabled = True
  End If
  Set rstent = dbtmp.OpenRecordset("select first(numpalet) as tnumpalet, first(entregat) as tentregat,first(numcalloff) as tnumcalloff,sum(metresisacs) as tmetres ,first(seccio) as tseccio from bobinesent where " + ventregatono + " comanda=" + atrim(vnumc) + " group by numpalet", dbOpenSnapshot, dbReadOnly)
  If Not rstent.EOF Then llistadepalets.AddItem generardescripciopalets(rstent, buscardesarroll(vnumc))
  If llistadepalets.ListCount = 0 Then possar_sensepalets
  Set rstent = Nothing
End Sub
Sub possar_sensepalets()
 Dim rst As Recordset
   Dim vitem As String
   Dim vnumc As Double
   vitem = citem
   vnumc = cadbl(reixac.TextMatrix(reixac.row, 0))
   Set rst = dbtmp.OpenRecordset("select * from calloffs_detall where comanda=" + atrim(vnumc) + " and item='" + atrim(vitem) + "'", dbOpenSnapshot, dbReadOnly)
   If Not rst.EOF Then
      llistadepalets.AddItem "Comanda sense producció Call-Off: " + atrim(rst!numcalloff)
        Else: llistadepalets.AddItem "Comanda sense producció"
   End If
   llistadepalets.ItemData(llistadepalets.NewIndex) = -2
   Set rst = Nothing
End Sub
Function buscardesarroll(vnumc As Double) As Double
   Dim rstclixes As Recordset
   Set rstclixes = dbclixes.OpenRecordset("SELECT comandes.comanda, Modificacions.desarroll FROM comandes LEFT JOIN Modificacions ON (comandes.numordremodificacio = Modificacions.ordre) AND (comandes.numtreball = Modificacions.id_treball) WHERE (((comandes.comanda)=" + atrim(vnumc) + "));", dbOpenSnapshot, dbReadOnly)
   If Not rstclixes.EOF Then buscardesarroll = cadbl(rstclixes!desarroll)
   Set rstclixes = Nothing
End Function
Function generardescripciopalets(rstent As Recordset, vdesarroll As Double)
   Dim vpeces As Double
   Dim vnumcalloff As String
  
   While Not rstent.EOF
      If vdesarroll > 0 Then vpeces = Redondejar(cadbl(rstent!tmetres) / IIf(atrim(rstent!tseccio) = "R", vdesarroll / 1000, 1), 0)
     ' If atrim(rstent!tnumcalloff) <> "" Then
      vnumcalloff = IIf(atrim(rstent!tnumcalloff) <> "", "Call-Off: ", "") + atrim(rstent!tnumcalloff) + IIf(rstent!tentregat = "S", " Entregat", "")
      '     Else: vnumcalloff = ""
     ' End If
      If vpeces > 0 Then
        llistadepalets.AddItem atrim(cadbl(rstent!tnumpalet)) + "  --> " + "1 x " + atrim(vpeces) + " Pcs   " + vnumcalloff
        If atrim(rstent!tnumcalloff) <> "" Then
           llistadepalets.ItemData(llistadepalets.NewIndex) = cadbl(vpeces)
            Else: llistadepalets.ItemData(llistadepalets.NewIndex) = -1
        End If
      End If
      rstent.MoveNext
   Wend
      
End Function

Private Sub Timer1_Timer()
  ensenyaritemsdelclient
  Timer1.Enabled = False
End Sub
