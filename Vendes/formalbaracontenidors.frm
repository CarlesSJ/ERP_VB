VERSION 5.00
Begin VB.Form formalbaracontenidors 
   Caption         =   "Albarà contenidors"
   ClientHeight    =   6825
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10770
   Icon            =   "formalbaracontenidors.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Height          =   360
      Left            =   3105
      Picture         =   "formalbaracontenidors.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprimir Albarà d'expedicions"
      Top             =   15
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Height          =   360
      Left            =   225
      Picture         =   "formalbaracontenidors.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Crear nou albarà de venda"
      Top             =   15
      Width           =   765
   End
   Begin VB.CommandButton Command4 
      Height          =   360
      Left            =   990
      Picture         =   "formalbaracontenidors.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Eliminar l'albarà"
      Top             =   15
      Width           =   705
   End
   Begin VB.CommandButton Command6 
      Height          =   360
      Left            =   1695
      Picture         =   "formalbaracontenidors.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Actualitzar/Grabar Registres"
      Top             =   15
      Width           =   705
   End
   Begin VB.CommandButton Command10 
      Height          =   360
      Left            =   2400
      Picture         =   "formalbaracontenidors.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Actualitzar/Grabar Registres"
      Top             =   15
      Width           =   705
   End
   Begin VB.Frame Frame3 
      Caption         =   "Albarans fets"
      Height          =   6240
      Left            =   150
      TabIndex        =   3
      Top             =   390
      Width           =   4770
      Begin VB.ListBox llistaalbarans 
         Height          =   5715
         Left            =   90
         TabIndex        =   4
         Top             =   315
         Width           =   4560
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contenidors en estoc"
      Height          =   5520
      Left            =   5130
      TabIndex        =   0
      Top             =   1095
      Width           =   5310
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   2355
         Picture         =   "formalbaracontenidors.frx":213C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Actualitzar/Grabar Registres"
         Top             =   195
         Width           =   705
      End
      Begin VB.TextBox cbuscar 
         Height          =   330
         Left            =   150
         TabIndex        =   11
         Top             =   225
         Width           =   2190
      End
      Begin VB.ListBox llistacontenidors 
         Height          =   4785
         ItemData        =   "formalbaracontenidors.frx":26C6
         Left            =   120
         List            =   "formalbaracontenidors.frx":26CD
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   585
         Width           =   4995
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Proveïdor - Recuperador"
      Height          =   780
      Left            =   5130
      TabIndex        =   2
      Top             =   315
      Width           =   5295
      Begin VB.TextBox nomproveidor 
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   315
         Width           =   4905
      End
   End
   Begin VB.Label etalbaranou 
      BackStyle       =   0  'Transparent
      Caption         =   "Creant albarà nou -  Escull els contenidors que vols adjuntar-hi."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   4020
      TabIndex        =   13
      Top             =   45
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.Menu mmante 
      Caption         =   "Manteniments"
      Begin VB.Menu mproveidors 
         Caption         =   "Manteniment de Proveïdor-Recuperador"
      End
   End
End
Attribute VB_Name = "formalbaracontenidors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbtintes As Database

Private Sub Command1_Click()
  carregar_contenidorsbuits
End Sub

Private Sub Command3_Click()
   Dim vnumalb As Integer
   Dim vdataentrega As String
   Dim vproveidor As String
   Dim vidproveidor As Long
   Dim rst As Recordset
   nomproveidor = ""
   nomproveidor.tag = ""
   Set rst = dbtmp.OpenRecordset("select * from capcalera_contenidors order by numalbara desc")
   If rst.EOF Then
      vnumalb = 1
       Else: vnumalb = rst!numalbara + 1
   End If
   vdataentrega = InputBox("Entra la data d'entrega dels contenidors.", "Data entrega", Date)
   If Not IsDate(vdataentrega) Then MsgBox "Aquesta data no es correcte", vbCritical, "Error": GoTo fi
   vproveidor = escullir_proveidor(vidproveidor)
   If vproveidor = "" Then MsgBox "No has escullit cap proveidor.", vbCritical, "Error": GoTo fi
   dbtmp.Execute "insert into capcalera_contenidors (numalbara,dataentrega,nomproveidor,idproveidor) values (" + atrim(vnumalb) + ",#" + atrim(Format(vdataentrega, "mm/dd/yyyy")) + "#,'" + atrim(vproveidor) + "'," + atrim(vidproveidor) + ")"
   nomproveidor = vproveidor
   nomproveidor.tag = Str(vidproveidor)
   afegir_bobines True
fi:
   carregar_albarans
   carregar_contenidorsbuits
   llistaalbarans.Selected(0) = True
   Set rst = Nothing
End Sub
Function escullir_proveidor(vid As Long) As String
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  Set formseleccio.Data1.Recordset = dbcomandes.OpenRecordset("select id,nomcomercial from recuperadorsdecontenidors order by nomcomercial")
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 2000
  formseleccio.width = 5000
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   escullir_proveidor = formseleccio.DBGrid2.Columns("nomcomercial")
   vid = formseleccio.DBGrid2.Columns("id")
  End If
  Unload formseleccio
End Function

Private Sub Command4_Click()
   Dim valbara As String
   If llistaalbarans.ListIndex = -1 Then MsgBox "Has d'escullir un albarà.", vbCritical, "Error": Exit Sub
   If UCase(InputBox("Estas segur que vols borrar aquest albarà?" + Chr(10) + "Escriu [eliminar] per eliminar-lo.", "Eliminar")) <> "ELIMINAR" Then Exit Sub
   valbara = atrim(llistaalbarans.ItemData(llistaalbarans.ListIndex))
   If MsgBox("Aquest contenidors de l'albarà estan a fàbrica?", vbExclamation + vbYesNo + vbDefaultButton1, "Atenció") = vbYes Then
       activarllaunesdelalbara cadbl(valbara)
   End If
   dbtmp.Execute "delete * from capcalera_contenidors where numalbara=" + atrim(valbara)
   dbtmp.Execute "delete * from linies_contenidors where numalbara=" + atrim(valbara)
   carregar_albarans
   afegir_bobines False
End Sub
Sub activarllaunesdelalbara(vnumalb As Double)
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("Select * from linies_contenidors where numalbara=" + atrim(vnumalb))
  While Not rst.EOF
    dbtintes.Execute "update llaunes set situacio='FORA',activa=false where numllauna='" + atrim(rst!numllauna) + "'"
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub

Private Sub Command6_Click()
  gravar_albara
 
 
End Sub
Sub gravar_albara()
   If etalbaranou.visible Then guardar_contenidors
End Sub
Sub guardar_contenidors()
   Dim i As Integer
   Dim vllauna As String
   Dim vdesc As String
   Dim valbara As String
   Dim vun As Boolean
   Dim rst As Recordset
   valbara = atrim(llistaalbarans.ItemData(llistaalbarans.ListIndex))
   For i = 0 To llistacontenidors.ListCount - 1
     If llistacontenidors.Selected(i) Then
       vun = True
       vllauna = atrim(Mid(llistacontenidors.List(i), 1, 9))
       vdesc = treure_apostruf(atrim(Mid(llistacontenidors.List(i), 10)))
       Set rst = dbtintes.OpenRecordset("select vmatriculacontenidor from llaunes where numllauna='" + vllauna + "'")
       vidcont = ""
       If Not rst.EOF Then vidcont = atrim(rst!vmatriculacontenidor)
       dbtmp.Execute "insert into linies_contenidors (numalbara,numllauna,descripcio,vmatriculacontenidor) values (" + valbara + ",'" + vllauna + "','" + vdesc + "','" + vidcont + "')"
     End If
   Next i
   If Not vun Then
      MsgBox "No hi ha contenidors escullits", vbCritical, "Error"
        Else:
           afegir_bobines False
   End If
End Sub
Sub afegir_bobines(va As Boolean)
  If Not va Then
       etalbaranou.visible = False
       nomproveidor = ""
       nomproveidor.tag = ""
       llistacontenidors.Clear
       Frame3.Enabled = True
       formalbaracontenidors.width = 5280
         Else
           etalbaranou.visible = True
           formalbaracontenidors.width = 11000
           If llistaalbarans.ListCount > 0 Then llistaalbarans.ListIndex = 0
           Frame3.Enabled = False
  End If
End Sub

Private Sub Command8_Click()
   Dim vnumalb As Double
   gravar_albara
   wait 2
   vnumalb = cadbl(llistaalbarans.ItemData(llistaalbarans.ListIndex))
   If llistaalbarans.ListIndex = -1 Then MsgBox "Has d'escullir un albarà.", vbCritical, "Error": Exit Sub
   imprimir_albara vnumalb
   donardebaixallaunes vnumalb
End Sub
Sub donardebaixallaunes(vnumalb As Double)
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("Select * from linies_contenidors where numalbara=" + atrim(vnumalb))
  While Not rst.EOF
    dbtintes.Execute "update llaunes set situacio='REC',activa=false where numllauna='" + atrim(rst!numllauna) + "'"
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
Sub imprimir_albara(vnumalb As Double)
 
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "albara_contenidors.rpt", 1)
 ' oreport.SQLQueryString = ""
  oreport.RecordSelectionFormula = "{capcalera_contenidors.numalbara}=" + atrim(vnumalb)
  'oreport.FormulaFields.GetItemByName("nomdirenvio").Text = "'" + treure_apostruf(etinfodelclient.tag) + "'"
  'oreport.SQLQueryString = ""
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "vendes.mdb"
  'oreport.Database.Tables.Item(2).Location = rutadelfitxer(cami) + "vendes.mdb"
  oreport.DiscardSavedData
  oreport.VerifyOnEveryPrint = False
  
  
  
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
    Else
      oreport.PrintOut False, 1
  End If
  

End Sub

Private Sub Form_Load()
    formalbaracontenidors.width = 5280
   Set dbtmp = formvendes.datacapcalera.Database
   Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
   carregar_albarans
   carregar_contenidorsbuits
End Sub
Sub carregar_contenidorsbuits()
  Dim rst As Recordset
  Dim vsql As String
  Dim vsubconsulta As String
  Dim vlinia As String
  llistacontenidors.Clear
  vsubconsulta = IIf(cadbl(nomproveidor.tag) > 0, " and llaunes.idproveidorrecuperador=" + atrim(cadbl(nomproveidor.tag)), "")
  vsql = "SELECT Llaunes.numllauna, Contenidors_material.descripcio, Llaunes.idmaterialcontenidor,llaunes.vmatriculacontenidor FROM Llaunes LEFT JOIN Contenidors_material ON Llaunes.idmaterialcontenidor = Contenidors_material.codi "
  vsql = vsql + " where llaunes.situacio<>'REC' and ((Llaunes.numllauna) Not In (select numllauna from linies_contenidors))and Llaunes.idmaterialcontenidor>0 and Llaunes.idmaterialcontenidor<>null  and llaunes.activa=false " + vsubconsulta '+ " ORDER BY numalbara "
  Set rst = dbtintes.OpenRecordset(vsql, , ReadOnly)
  
  While Not rst.EOF
     vlinia = justificar(rst!numllauna, 9, "D") + " " + justificar(atrim(rst!descripcio), 20, "E") '+ atrim(rst!vmatriculacontenidor)
     If cbuscar <> "" Then
       If InStr(1, UCase(vlinia), UCase(cbuscar)) > 0 Then llistacontenidors.AddItem vlinia
         Else: llistacontenidors.AddItem vlinia
     End If
     rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
Sub carregar_albarans()
  Dim rst As Recordset
  Dim v As String
  
  Set rst = dbtmp.OpenRecordset("select * from capcalera_contenidors order by numalbara desc")
  llistaalbarans.Clear
  While Not rst.EOF
    v = justificar(rst!numalbara, 5, "D") + " "
    v = v + justificar(Format(rst!dataentrega, "dd/mm/yy"), 9, "D") + " "
    v = v + atrim(rst!nomproveidor)
    llistaalbarans.AddItem v
    llistaalbarans.ItemData(llistaalbarans.NewIndex) = cadbl(rst!numalbara)
    rst.MoveNext
  Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set dbtmp = Nothing
   Set dbtintes = Nothing
End Sub

Private Sub mproveidors_Click()
  Load formaltarep
  formaltarep.caption = "Manteniment de Recuperadors de contenidors"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from recuperadorsdecontenidors "
  formaltarep.width = 9000
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  
  formaltarep.DBGrid1.Columns(0).visible = False
  formaltarep.DBGrid1.Columns(1).width = 6000
  formaltarep.DBGrid1.Columns(2).width = 2000
  formaltarep.Show 1
End Sub
Function justificar(v As String, longitut As Integer, DoE As String) As String
    v = Mid(v, 1, longitut)
    If DoE = "E" Then
       v = v + Space(longitut - Len(v))
      Else: v = Space(longitut - Len(v)) + v
    End If
    justificar = v
End Function
