VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formtestcolors 
   Caption         =   "Resultats Test colors dels fotogravadors"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13425
   Icon            =   "formtestcolors.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton alta 
      Height          =   360
      Left            =   1740
      Picture         =   "formtestcolors.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Alta treball"
      Top             =   195
      Width           =   420
   End
   Begin VB.CommandButton eliminar 
      Height          =   360
      Left            =   2160
      Picture         =   "formtestcolors.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminacio del treball"
      Top             =   195
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Height          =   360
      Left            =   2580
      Picture         =   "formtestcolors.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Actualitzar Registres"
      Top             =   195
      Width           =   420
   End
   Begin VB.ListBox llistacomandes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   75
      TabIndex        =   1
      Top             =   330
      Width           =   1650
   End
   Begin VB.Data datatestcolors 
      Caption         =   "datatestcolors"
      Connect         =   "Access"
      DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\clixesnous.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   4155
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"formtestcolors.frx":1628
      Top             =   1020
      Visible         =   0   'False
      Width           =   3180
   End
   Begin MSDBGrid.DBGrid reixatestcolors 
      Bindings        =   "formtestcolors.frx":1764
      Height          =   6405
      Left            =   1755
      OleObjectBlob   =   "formtestcolors.frx":177D
      TabIndex        =   0
      Top             =   570
      Width           =   11310
   End
   Begin VB.Label Label1 
      Caption         =   "Comandes Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   195
      TabIndex        =   2
      Top             =   90
      Width           =   1725
   End
End
Attribute VB_Name = "formtestcolors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_Click()

End Sub

Private Sub alta_Click()
   crear_nou_anilox
End Sub

Private Sub Command1_Click()
  If datatestcolors.Recordset.EditMode > 0 Then
    datatestcolors.Recordset.Update
  End If
  datatestcolors.Recordset.Move 0
End Sub

Private Sub eliminar_Click()
   If datatestcolors.Recordset.EditMode > 0 Then MsgBox "Estas editant algun registre, primer guarda les dades.", vbCritical, "Atenció": Exit Sub
   If MsgBox("Segur que vols borrar aquesta linia?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   datatestcolors.Recordset.Delete
   datatestcolors.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 112 Then Command1_Click
End Sub

Private Sub Form_Load()
   datatestcolors.DatabaseName = camiclixes
   carregar_llistacomandes
   If llistacomandes.ListCount = 1 Then
      llistacomandes.ListIndex = 0
   End If
  carregar_aniloxos cadbl(llistacomandes.List(llistacomandes.ListIndex))
End Sub
Sub carregar_llistacomandes()
   Dim rst As Recordset
   llistacomandes.Clear
   
   'Set rst = dbclixes.OpenRecordset("select distinct comanda from testscolors where id_treball=" + atrim(id_treball) + " and nummodificacio=" + atrim(ordremodificacio))
   Set rst = dbclixes.OpenRecordset("select  comanda from comandes where numtreball=" + atrim(id_treball) + " and numordremodificacio=" + atrim(ordremodificacio))
   While Not rst.EOF
     llistacomandes.AddItem atrim(rst!comanda)
     rst.MoveNext
   Wend
  Set rst = Nothing
End Sub
Sub carregar_aniloxosnous(numc As Double)
  Dim rsta As Recordset
  datatestcolors.RecordSource = "SELECT testscolors.id_treball, testscolors.nummodificacio, testscolors.comanda, testscolors.id_anilox, testscolors.color,aniloxos.nommaquina, aniloxos.lineatura, testscolors.densitat, testscolors.densitatrf FROM testscolors LEFT JOIN aniloxos ON testscolors.id_anilox = aniloxos.id "
  datatestcolors.RecordSource = datatestcolors.RecordSource + " where id_treball=" + atrim(id_treball) + " and nummodificacio=" + atrim(ordremodificacio) + " and comanda=" + atrim(numc)
  datatestcolors.RecordSource = datatestcolors.RecordSource + " order by nommaquina,lineatura;"
  datatestcolors.Refresh
  Set rsta = dbclixes.OpenRecordset("select * from aniloxos order by id", dbOpenSnapshot)
  While Not rsta.EOF
        datatestcolors.Recordset.FindFirst "id_anilox=" + atrim(rsta!ID)
        If datatestcolors.Recordset.NoMatch Then
         'crear_nou_anilox rsta
        End If
     rsta.MoveNext
  Wend
  datatestcolors.Refresh
End Sub
Sub crear_nou_anilox()
   datatestcolors.Recordset.AddNew
   datatestcolors.Recordset!id_treball = id_treball
   datatestcolors.Recordset!nummodificacio = ordremodificacio
   datatestcolors.Recordset!comanda = cadbl(llistacomandes.List(llistacomandes.ListIndex))
'   datatestcolors.Recordset!id_anilox = rsta!ID
   datatestcolors.Recordset.Update
End Sub

Private Sub llistacomandes_Click()
   'carregar_aniloxosnous cadbl(llistacomandes.List(0))
     carregar_aniloxos cadbl(llistacomandes.List(llistacomandes.ListIndex))
End Sub
Sub carregar_aniloxos(numc As Double)
  Dim rsta As Recordset
  datatestcolors.RecordSource = "SELECT testscolors.id_treball, testscolors.nummodificacio, testscolors.comanda, testscolors.id_anilox,testscolors.color, aniloxos.nommaquina, aniloxos.lineatura, testscolors.densitat, testscolors.densitatrf FROM testscolors LEFT JOIN aniloxos ON testscolors.id_anilox = aniloxos.id "
  datatestcolors.RecordSource = datatestcolors.RecordSource + " where id_treball=" + atrim(id_treball) + " and nummodificacio=" + atrim(ordremodificacio) + " and comanda=" + atrim(numc)
  datatestcolors.RecordSource = datatestcolors.RecordSource + " order by nommaquina,lineatura;"
  datatestcolors.Refresh
End Sub


Private Sub reixatestcolors_ButtonClick(ByVal ColIndex As Integer)
   If reixatestcolors.Columns(ColIndex).Caption = "Nom Impresora" Or reixatestcolors.Columns(ColIndex).Caption = "Anilox" Then
      'If reixatestcolors.row = datatestcolors.Recordset.RecordCount Then reixatestcolors.Columns("Anilox") = 0
      escullir_impresora
   End If
   If reixatestcolors.Columns(ColIndex).Caption = "Color Bàsic" Then
      escullir_color
   End If
End Sub
Sub creartaulacolors()
    On Error GoTo fi
    dbclixes.Execute "create table tmp_colorsbasics (color string)"
    dbclixes.Execute "insert into tmp_colorsbasics (color) values ('Blanc')"
    dbclixes.Execute "insert into tmp_colorsbasics (color) values ('Negre')"
    dbclixes.Execute "insert into tmp_colorsbasics (color) values ('Groc')"
    dbclixes.Execute "insert into tmp_colorsbasics (color) values ('Vermell')"
    dbclixes.Execute "insert into tmp_colorsbasics (color) values ('Blau')"
    
fi:
End Sub
Sub escullir_color()
   Dim rst As Recordset
   Dim nomimp As String
   creartaulacolors
   If datatestcolors.Recordset.EditMode = 0 Then datatestcolors.Recordset.Edit
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
    formseleccio.Data1.RecordSource = "select *  from tmp_colorsbasics"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.DBGrid2.Columns(0).Width = 1500
   formseleccio.refrescar
   formseleccio.Show 1
   
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
              reixatestcolors.Text = atrim(formseleccio.DBGrid2.Columns("color"))
        End If
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub


Sub escullir_impresora()
   Dim rst As Recordset
   Dim nomimp As String
   Set rst = dbclixes.OpenRecordset("select distinct nommaquina  from aniloxos where nommaquina<>'' order by nommaquina asc")
   If rst.EOF Then Exit Sub
   If datatestcolors.Recordset.EditMode = 0 Then datatestcolors.Recordset.Edit
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
    formseleccio.Data1.RecordSource = "select distinct nommaquina  from aniloxos where nommaquina<>'' order by nommaquina asc"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.DBGrid2.Columns(0).Width = 1500
   formseleccio.refrescar
   formseleccio.Show 1
   
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
              nomimp = atrim(formseleccio.DBGrid2.Columns("nommaquina"))
        End If
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   
   If nomimp <> "" Then escullir_anilox nomimp
   
End Sub

Sub escullir_anilox(nomimp As String)
   Dim rst As Recordset
   
   Set rst = dbclixes.OpenRecordset("select nommaquina ,lineatura as anilox from aniloxos where nommaquina='" + atrim(nomimp) + "' order by lineatura asc")
   If rst.EOF Then Exit Sub
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
    formseleccio.Data1.RecordSource = "select id,nommaquina ,lineatura as anilox from aniloxos where nommaquina='" + atrim(nomimp) + "' order by lineatura asc"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns(0).visible = False
   formseleccio.DBGrid2.Columns(1).Width = 1600
   formseleccio.DBGrid2.Columns(2).Width = 500
   formseleccio.Show 1
   
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
             datatestcolors.Recordset!id_anilox = cadbl(formseleccio.DBGrid2.Columns("id"))
        End If
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   If datatestcolors.Recordset.EditMode > 0 Then datatestcolors.Recordset.Update
   datatestcolors.Recordset.Move 0
   reixatestcolors.col = 3
   reixatestcolors.SetFocus
End Sub




















