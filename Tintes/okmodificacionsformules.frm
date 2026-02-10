VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formokformules 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ok's modificacions de les fòrmules"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11145
   Icon            =   "okmodificacionsformules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Tot Si"
      Height          =   240
      Left            =   9855
      TabIndex        =   1
      Top             =   15
      Width           =   900
   End
   Begin VB.Data dataokformules 
      Caption         =   "dataokformules"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5145
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   510
      Visible         =   0   'False
      Width           =   2385
   End
   Begin MSDBGrid.DBGrid reixaoks 
      Bindings        =   "okmodificacionsformules.frx":058A
      Height          =   7650
      Left            =   90
      OleObjectBlob   =   "okmodificacionsformules.frx":05A3
      TabIndex        =   0
      Top             =   270
      Width           =   10950
   End
End
Attribute VB_Name = "formokformules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
dataokformules.Refresh
  While Not dataokformules.Recordset.EOF
     dataokformules.Recordset.Edit
     dataokformules.Recordset!ok = True
     dataokformules.Recordset.Update
     dataokformules.Recordset.MoveNext
   Wend
   dataokformules.Refresh
End Sub

Private Sub Form_Load()
   dataokformules.DatabaseName = camitintes
   dataokformules.RecordSource = "select * from test_actualitzacioformules order by formula"
   
End Sub

Private Sub reixaoks_ButtonClick(ByVal ColIndex As Integer)
   reixaoks = Not dataokformules.Recordset.Fields(ColIndex)
End Sub

