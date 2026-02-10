VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form llistatpantones 
   Caption         =   "Llistat Pantones"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   Icon            =   "llistatpantones.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   540
      Top             =   105
   End
   Begin VB.TextBox textabuscar 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3030
      TabIndex        =   1
      Top             =   45
      Width           =   3450
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\tintes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7260
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "dadesllaunes"
      Top             =   45
      Visible         =   0   'False
      Width           =   2145
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "llistatpantones.frx":058A
      Height          =   7080
      Left            =   105
      OleObjectBlob   =   "llistatpantones.frx":059A
      TabIndex        =   0
      Top             =   405
      Width           =   11340
   End
   Begin VB.Label condicio 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   75
      Width           =   2865
   End
End
Attribute VB_Name = "llistatpantones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If Screen.ActiveControl.Name = "DBGrid1" Then
    condicio = DBGrid1.Columns(DBGrid1.Col).Caption + " like '*"
    condicio.Tag = DBGrid1.Columns(DBGrid1.Col).DataField + " like '*"
  End If
End Sub

Private Sub Form_Activate()
 condicio = DBGrid1.Columns(DBGrid1.Col).Caption + " like '*"
 condicio.Tag = DBGrid1.Columns(DBGrid1.Col).DataField + " like '*"
 textabuscar.SetFocus
End Sub

Private Sub Form_Load()
  On Error Resume Next
  'Set dbpantones = OpenDatabase(llegir_ini("General", "rutallistats", "comandes.ini") + "connexio stock tintes pantones.mdb")
  'dbpantones.Execute "drop table llistat"
  'dbpantones.Execute ("SELECT DISTINCTROW Estocpantones.Situacio, Estocpantones.numlata, Estocpantones.Kgpan, Pantones.despan, Familiacol.descol, Familiades.desfam, Pantones.refpan, Estocpantones.lofab, Estocpantones.datamanpan INTO llistat FROM (Familiades INNER JOIN (Familiacol INNER JOIN Pantones ON Familiacol.cdcol = Pantones.cdpancol) ON Familiades.cdfam = Pantones.cdpanfam) INNER JOIN Estocpantones ON Pantones.cdpant = Estocpantones.cdpan Where (((Estocpantones.databaixa) Is Null))ORDER BY pantones.despan;")
  'Set rstpantones = dbpantones.OpenRecordset("llistat")
  
  'Set Data1.Recordset = rstpantones
  Data1.DatabaseName = rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "tintes.mdb"
  Data1.RecordSource = "select * from dadesllaunes"
  Data1.Refresh
  
  DBGrid1.Refresh
  DBGrid1.Col = 2
  
End Sub

Private Sub textabuscar_Change()
 Dim c As Byte
r = condicio.Tag + textabuscar + "*'"
'Set rstpantones = dbpantones.OpenRecordset("select * from dadesllaunes where " + r)
'Set Data1.Recordset = rstpantones
Data1.RecordSource = "select * from dadesllaunes where " + r
Data1.Refresh
DBGrid1.Tag = DBGrid1.Col
'DBGrid1.Refresh
End Sub

