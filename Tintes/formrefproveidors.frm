VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formrefproveidors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Referencies de proveïdors"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8925
   Icon            =   "formrefproveidors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datarefproveidor 
      Caption         =   "datarefproveidor"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2910
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"formrefproveidors.frx":058A
      Top             =   180
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Frame Frame8 
      Height          =   2445
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      Begin VB.CommandButton Command31 
         Height          =   330
         Left            =   975
         Picture         =   "formrefproveidors.frx":0611
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Marcar com a Activa/Principal."
         Top             =   195
         Width           =   345
      End
      Begin VB.CommandButton Command30 
         Height          =   330
         Left            =   435
         Picture         =   "formrefproveidors.frx":0B9B
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Eliminar referència"
         Top             =   195
         Width           =   345
      End
      Begin VB.CommandButton Command29 
         Height          =   330
         Left            =   75
         Picture         =   "formrefproveidors.frx":1125
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Nova referència"
         Top             =   195
         Width           =   345
      End
      Begin MSDBGrid.DBGrid reixarefproveidor 
         Bindings        =   "formrefproveidors.frx":16AF
         Height          =   1770
         Left            =   60
         OleObjectBlob   =   "formrefproveidors.frx":16CA
         TabIndex        =   4
         Top             =   525
         Width           =   8730
      End
   End
End
Attribute VB_Name = "formrefproveidors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command29_Click()
    crear_referencia
End Sub
Sub crear_referencia()
   Dim vrefproveidor As String
  Dim vcodibido As Double
  Dim nomproveidor As String
  Dim codiprov As Long
  Dim vestocminim As Double
  'If ccoditinta = "" Then MsgBox "Primer has d'haver guardat la tinta abans d'afegir-hi una referència", vbCritical, "Atenció": Exit Sub
  escullir_proveidor codiprov, nomproveidor
  If codiprov = 0 Then Exit Sub
  'If InStr(1, nomproveidor, "INPLACSA") > 0 Then MsgBox "Amb el proveïdor INPLACSA no es poden afegir referències de proveïdor.", vbCritical, "Atenció": Exit Sub
  vrefproveidor = InputBox("Entra el codi de referencia del proveïdor.", "Nova referència")
  If vrefproveidor = "" Then Exit Sub
  If comprovarsireferenciaestarepetida(vrefproveidor) Then MsgBox "Aquesta referència ja existeix per un altra producte.", vbCritical, "Atenció": Exit Sub
  vcodibido = escullir_bido
  'vestocminim = cadbl(InputBox("Entra els bidons d'estoc mínim que vols tenir.", "Estoc mínim"))
  If vcodibido > 0 Then
    dbtintes.Execute "Insert into tintesreferencies (idtinta,referencia,id_bido,codiproveidor,nomproveidor) values ('" + atrim(formtintes.tintes.Recordset!idtinta) + "','" + treure_apostruf(vrefproveidor) + "'," + atrim(vcodibido) + "," + atrim(codiprov) + ",'" + treure_apostruf(nomproveidor) + "')"
  End If
  datarefproveidor.Refresh
  If datarefproveidor.Recordset.RecordCount = 1 Then
     datarefproveidor.Recordset.Edit
     datarefproveidor.Recordset!predeterminada = True
     datarefproveidor.Recordset.Update
  End If
End Sub

Function comprovarsireferenciaestarepetida(vref As String) As Boolean
  Dim rst As Recordset
  Set rst = dbtintes.OpenRecordset("select * from tintesreferencies where referencia='" + treure_apostruf(vref) + "'")
  If Not rst.EOF Then comprovarsireferenciaestarepetida = True: Exit Function
  Set rst = Nothing
End Function
  
Sub escullir_proveidor(codiprov As Long, nomproveidor As String)

  Load formseleccio
  formseleccio.caption = "Selecciona el proveïdor"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codi,nom,aliastintes from proveidors where aliastintes<>'' order by nom"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 4500
  formseleccio.DBGrid2.Columns(2).width = 800
  formseleccio.Show 1
  If seleccioret = 1 Then
   nomproveidor = atrim(formseleccio.Data1.Recordset!nom)
   codiprov = formseleccio.Data1.Recordset!codi
  End If
  Unload formseleccio
End Sub
Function escullir_bido() As Integer
  
  Load formseleccio
  formseleccio.width = formseleccio.width + (formseleccio.width / 2)
  formseleccio.caption = "Escull un bidó"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "select * from tipusbidons  order by capacitat"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 4500
  'formseleccio.DBGrid2.Columns(2).Width = 800
  formseleccio.Data1.Recordset.FindFirst "capacitat=20"
  formseleccio.Show 1
  escullir_bido = 0
  If seleccioret = 1 Then
   escullir_bido = formseleccio.Data1.Recordset!id
  End If
  Unload formseleccio
  
End Function

Function hiharelacionsreferencies() As Boolean
   Dim rst As Recordset
  ' Set rst = Data1.Database.OpenRecordset("select )
   If Not rst.EOF Then hiharelacions = True
   Set rst = Nothing
End Function

Private Sub Command30_Click()
   'If ccoditinta = "" Then MsgBox "Primer has d'haver guardat la tinta abans d'eliminar una referència", vbCritical, "Atenció": Exit Sub
   If datarefproveidor.Recordset.EOF Then Exit Sub
   If hiharelacioreferencia(datarefproveidor.Recordset!id) Then MsgBox "Hi ha una Llauna amb aquesta referència no pots eliminar-la", vbCritical, "Error": Exit Sub
   If MsgBox("Segur que vols eliminar aquesta referència?", vbCritical + vbYesNo + vbDefaultButton2, "Eliminar Referència") = vbNo Then Exit Sub
   datarefproveidor.Recordset.Delete
   datarefproveidor.Refresh
   possarunareferenciapredeterminada
End Sub
Sub possarunareferenciapredeterminada()
   Dim refpre As Boolean
   If datarefproveidor.Recordset.EOF Then Exit Sub
   datarefproveidor.Recordset.MoveFirst
   
   While Not datarefproveidor.Recordset.EOF
     If datarefproveidor.Recordset!predeterminada Then refpre = True
     datarefproveidor.Recordset.MoveNext
   Wend
   If Not refpre Then
     datarefproveidor.Refresh
     datarefproveidor.Recordset.Edit
     datarefproveidor.Recordset!predeterminada = True
     datarefproveidor.Recordset.Update
   End If
End Sub
Function hiharelacioreferencia(nid As Integer) As Boolean
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select numllauna from llaunes where id_refproveidor=" + atrim(nid))
   If Not rst.EOF Then hiharelacioreferencia = True
   Set rst = Nothing
End Function
Private Sub Command31_Click()
    Dim vsql As String
    If datarefproveidor.Recordset.EOF Then Exit Sub
    If Not datarefproveidor.Recordset!predeterminada Then
        If MsgBox("Segur que vols marcar la referència " + atrim(datarefproveidor.Recordset!referencia) + " com a Principal?", vbInformation + vbYesNo, "Referència Principal") = vbNo Then Exit Sub
        vsql = "UPDATE tintesreferencies LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id SET tintesreferencies.predeterminada = False "
        vsql = vsql + " WHERE (((tipusbidons.capacitat)=" + atrim(datarefproveidor.Recordset!capacitat) + ") AND ((tintesreferencies.idtinta)=" + atrim(cadbl(formtintes.tintes.Recordset!idtinta)) + "));"
        dbtintes.Execute vsql
        dbtintes.Execute "update tintesreferencies set predeterminada=true where id=" + atrim(datarefproveidor.Recordset!id)
          Else
            If MsgBox("Segur que vols DES-marcar la referència " + atrim(datarefproveidor.Recordset!referencia) + " com a Principal?", vbInformation + vbYesNo, "Referència Principal") = vbNo Then Exit Sub
            vsql = "UPDATE tintesreferencies LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id SET tintesreferencies.predeterminada = False "
            vsql = vsql + " WHERE (((tipusbidons.capacitat)=" + atrim(datarefproveidor.Recordset!capacitat) + ") AND ((tintesreferencies.idtinta)=" + atrim(cadbl(formtintes.tintes.Recordset!idtinta)) + "));"
            dbtintes.Execute vsql
            'dbtintes.Execute "update tintesreferencies set predeterminada=true where id=" + atrim(datarefproveidor.Recordset!id)
    End If
    datarefproveidor.Refresh
    
End Sub

Private Sub Form_Load()
  datarefproveidor.DatabaseName = rutadelfitxer(cami) + "tintes.mdb"
  datarefproveidor.RecordSource = "SELECT tintesreferencies.id, tintesreferencies.referencia, tintesreferencies.predeterminada,tintesreferencies.nomproveidor,tipusbidons.capacitat,tipusbidons.nombido,tintesreferencies.id_bido FROM tintesreferencies LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id where idtinta=" + atrim(cadbl(formtintes.tintes.Recordset!idtinta)) + " order by predeterminada ;"
  datarefproveidor.Refresh

End Sub

Private Sub reixarefproveidor_DblClick()
  Dim vestocminim As Double
  Dim vref As String
  Dim codiprov As Long
  Dim vcodibido As Long
  Dim nomproveidor As String
  If datarefproveidor.Recordset.EOF Then Exit Sub
  If reixarefproveidor.col = 1 Then Command31_Click: Exit Sub
  If reixarefproveidor.col = 2 Then
      vref = InputBox("Escriu la referencia nova.", "Canvi de referencia", reixarefproveidor)
      If Len(vref) > 1 Then
         dbtintes.Execute "update tintesreferencies set referencia='" + atrim(vref) + "' where id=" + atrim(datarefproveidor.Recordset!id)
      End If
      datarefproveidor.Refresh
  End If
  If reixarefproveidor.col = 3 Then
      escullir_proveidor codiprov, nomproveidor
      If cadbl(codiprov) > 0 Then
         dbtintes.Execute "update tintesreferencies set codiproveidor=" + atrim(cadbl(codiprov)) + " ,nomproveidor='" + atrim(nomproveidor) + "' where id=" + atrim(datarefproveidor.Recordset!id)
      End If
      datarefproveidor.Refresh
  End If
  If reixarefproveidor.col = 4 Then
      vcodibido = escullir_bido
       If cadbl(vcodibido) > 0 Then
         dbtintes.Execute "update tintesreferencies set id_bido=" + atrim(cadbl(vcodibido)) + " where id=" + atrim(datarefproveidor.Recordset!id)
      End If
      datarefproveidor.Refresh
  End If
'  If reixarefproveidor.col = 4 Then
'     vestocminim = cadbl(InputBox("Entra els bidons d'estoc mínim que vols tenir.", "Estoc mínim"))
'     datarefproveidor.Recordset.Edit
'     datarefproveidor.Recordset!estocminim = vestocminim
'     datarefproveidor.Recordset.Update
'     datarefproveidor.Refresh
'  End If
  
End Sub
