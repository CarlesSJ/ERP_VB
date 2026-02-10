VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formrecepciorma 
   Caption         =   "Recepcio de RMA"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15315
   Icon            =   "formrecepciorma.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   15315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Height          =   390
      Left            =   2385
      Picture         =   "formrecepciorma.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Actualitzar/Grabar Registres"
      Top             =   870
      Width           =   705
   End
   Begin VB.CommandButton Command6 
      Height          =   390
      Left            =   1530
      Picture         =   "formrecepciorma.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Actualitzar/Grabar Registres"
      Top             =   870
      Width           =   705
   End
   Begin VB.CommandButton Command5 
      Height          =   390
      Left            =   1125
      Picture         =   "formrecepciorma.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   870
      Width           =   405
   End
   Begin VB.CommandButton Command4 
      Height          =   390
      Left            =   720
      Picture         =   "formrecepciorma.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   870
      Width           =   405
   End
   Begin VB.CommandButton Command3 
      Height          =   390
      Left            =   300
      Picture         =   "formrecepciorma.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   870
      Width           =   405
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ensenya Acabades"
      Height          =   465
      Left            =   13515
      TabIndex        =   3
      Top             =   795
      Width           =   1650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ensenya Pendents"
      Height          =   465
      Left            =   11790
      TabIndex        =   2
      Top             =   795
      Width           =   1650
   End
   Begin VB.Data datarma 
      Caption         =   "datarma"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   12495
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "bobinesrma"
      Top             =   75
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Frame Frame1 
      Height          =   5940
      Left            =   150
      TabIndex        =   0
      Top             =   1215
      Width           =   15045
      Begin MSDBGrid.DBGrid reixa 
         Bindings        =   "formrecepciorma.frx":213C
         Height          =   5685
         Left            =   75
         OleObjectBlob   =   "formrecepciorma.frx":214E
         TabIndex        =   1
         Top             =   180
         Width           =   14865
      End
   End
End
Attribute VB_Name = "formrecepciorma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  datarma.RecordSource = "select * from bobinesrma where estatrma='' order by datarecepcio desc"
  datarma.Refresh
End Sub

Private Sub Command2_Click()
  datarma.RecordSource = "select * from bobinesrma where estatrma<>'' order by datarecepcio desc"
  datarma.Refresh
End Sub

Private Sub Command3_Click()
   If datarma.Recordset.EditMode > 0 Then MsgBox "Ja està editant", vbCritical + vbYesNo + vbDefaultButton2, "Atenció": Exit Sub
   datarma.RecordSource = "select * from bobinesrma where estatrma='' order by datarecepcio desc"
   datarma.Refresh
   datarma.Recordset.AddNew
   datarma.Recordset!Datarecepcio = Date
   datarma.Recordset.Update
   datarma.Recordset.Bookmark = datarma.Recordset.LastModified
   reixa.AllowUpdate = True
   reixa.col = 2
   reixa.SetFocus
   
End Sub

Private Sub Command4_Click()
   If datarma.Recordset.EditMode > 0 Then MsgBox "Ja està editant", vbCritical + vbYesNo + vbDefaultButton2, "Atenció": Exit Sub
 '  datarma.Recordset.Edit
   reixa.AllowUpdate = True
   reixa.col = 2
   reixa.SetFocus
End Sub

Private Sub Command7_Click()
  ' imprimir albarà agrupant pel client i la data
  
  ' quan hagi imprès passar les bobines de la rebobinadora a un palet nou, borrar kilos i metres i palet xo guardar-ho al
  '   mateix registre de bobinesrma s'han de crear els camps marcar el camp de utilitzadaabaixa per controlar que es rma
  'un cop impres no s'ha de poder borrar ni modificar els camps nomes el pvp
  
  
End Sub

Private Sub Form_Load()
   datarma.DatabaseName = rutadelfitxer(cami) + "baixes.mdb"
   datarma.RecordSource = "select * from bobinesrma where estatrma='' order by datarecepcio desc"
   datarma.RecordSource = "bobinesrma"
   datarma.Refresh
End Sub

Private Sub reixa_AfterColEdit(ByVal ColIndex As Integer)
  If reixa.AllowUpdate Then
    If reixa.Columns(ColIndex).DataField = "Comanda" Then possardadesdelacomanda reixa.Columns("Comanda")
   
  End If
End Sub

Sub possardadesdelacomanda(vnumc As Double)
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("SELECT comandes.client, clients.nom FROM clients RIGHT JOIN comandes ON clients.codi = comandes.client where comandes.comanda=" + atrim(vnumc))
   If Not rst.EOF Then
       reixa.Columns("client") = atrim(rst!client)
       reixa.Columns("nomclient") = atrim(rst!nom)
   End If
   Set rst = Nothing
End Sub

Private Sub reixa_ButtonClick(ByVal ColIndex As Integer)
   If reixa.Columns(ColIndex).DataField = "numbob" Then
     reixa = triarbobina(cadbl(reixa.Columns("Comanda")))
   End If
End Sub
Function triarbobina(vnumc As Double) As String
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "baixes.mdb"
  formseleccio.Data1.RecordSource = "SELECT bobinesreb.numerodebobina as Numbob, bobinesreb.kilos as Kilos, bobinesreb.metres as Metres, bobinesreb.palet FROM bobinesreb RIGHT JOIN rebobinadores ON bobinesreb.controlid = rebobinadores.Id Where (((bobinesreb.numerodebobina) > 0) And ((rebobinadores.comanda) = " + atrim(vnumc) + ")) ORDER BY bobinesreb.numerodebobina;"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).visible = False
  'formseleccio.DBGrid2.Columns(2).width = 900
  formseleccio.width = 9000
  If formseleccio.Data1.Recordset.RecordCount > 1 Then
     formseleccio.Show 1
    Else: MsgBox "Nomes hi ha cap bobina de baixa de rebobinadora per aquesta comanda", vbCritical, "Error"
  End If
  If seleccioret = 1 Then
        
        If Not formseleccio.Data1.Recordset.EOF Then
           triarbobina = cadbl(formseleccio.DBGrid2.Columns("Numbob"))
        End If
   End If
   Unload formseleccio
End Function
