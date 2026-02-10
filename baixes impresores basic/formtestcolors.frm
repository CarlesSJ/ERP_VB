VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formtestcolors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test de colors (Fotogravador)"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11370
   Icon            =   "formtestcolors.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datatestcolors 
      Caption         =   "datatestcolors"
      Connect         =   "Access"
      DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\clixesnous.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"formtestcolors.frx":058A
      Top             =   450
      Visible         =   0   'False
      Width           =   3180
   End
   Begin MSDBGrid.DBGrid reixatestcolors 
      Bindings        =   "formtestcolors.frx":06C6
      Height          =   6405
      Left            =   0
      OleObjectBlob   =   "formtestcolors.frx":06DF
      TabIndex        =   0
      Top             =   0
      Width           =   11310
   End
End
Attribute VB_Name = "formtestcolors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   If formtestcolors.Tag = "sortir" Then Unload Me
End Sub

Private Sub Form_Load()
   Dim id_treball As Double
   Dim ordremodificacio As Double
   Dim numc As Double
   Dim rst As Recordset
   Dim rstt As Recordset
   Dim rstc As Recordset
   Dim instsql As String
   Dim dbclixes As Database
   Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + Form1.comanda)
   If rstc.EOF Then Exit Sub
   datatestcolors.DatabaseName = rutadelfitxer(cami) + "clixesnous.mdb"
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb", , True)
   Set rstt = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(rstc!numtreball) + " and ordre=" + atrim(rstc!numordremodificacio))
   If rstt.EOF Then Exit Sub
   instsql = "SELECT First(Modificacions.id_treball) AS pid_treball, First(Modificacions.ordre) AS pordre, Max(comandes.comanda) AS mcomanda FROM comandes LEFT JOIN (Clixes RIGHT JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball) ON (comandes.numordremodificacio = Modificacions.ordre) AND (comandes.numtreball = Modificacions.id_treball) "
   instsql = instsql + " Where (((Modificacions.fotograbador) =" + atrim(rstt!fotograbador) + ") And ((UCase([marca])) = 'TEST') And ((Modificacions.sistemadimpresio) = '" + rstt!sistemadimpresio + "'))"
   instsql = instsql + "ORDER BY Max(comandes.comanda) DESC;"
   Set rst = dbclixes.OpenRecordset(instsql)
   If cadbl(rst!pid_Treball) = 0 Then formtestcolors.Hide: GoTo sortir
   datatestcolors.RecordSource = "SELECT testscolors.id_treball, testscolors.nummodificacio, testscolors.comanda, testscolors.id_anilox, testscolors.color,aniloxos.nommaquina, aniloxos.lineatura, testscolors.densitat, testscolors.densitatrf FROM testscolors LEFT JOIN aniloxos ON testscolors.id_anilox = aniloxos.id "
   datatestcolors.RecordSource = datatestcolors.RecordSource + " where id_treball=" + atrim(rst!pid_Treball) + " and nummodificacio=" + atrim(rst!pordre) + " and comanda=" + atrim(rst!mcomanda)
   datatestcolors.RecordSource = datatestcolors.RecordSource + " order by nommaquina,lineatura;"
   datatestcolors.Refresh
   Set rst = Nothing
   Set rstt = Nothing
   Set rstc = Nothing
    Set dbclixes = Nothing
    Exit Sub
sortir:
    formtestcolors.Tag = "sortir"
End Sub
