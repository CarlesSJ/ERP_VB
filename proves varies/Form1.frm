VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   360
      Left            =   5220
      TabIndex        =   5
      Top             =   555
      Width           =   1410
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   420
      Left            =   5610
      TabIndex        =   4
      Top             =   2025
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   330
      Left            =   5100
      TabIndex        =   3
      Top             =   1155
      Width           =   450
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   540
      Left            =   2745
      TabIndex        =   2
      Top             =   2355
      Width           =   1740
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1995
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "comandes"
      Top             =   510
      Width           =   3075
   End
   Begin VB.TextBox Text2 
      Height          =   570
      Left            =   2565
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1770
      Width           =   1995
   End
   Begin VB.TextBox Text1 
      DataField       =   "comanda"
      DataSource      =   "data1"
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1125
      Width           =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbcomandes As Database
Dim sqlconn As Connection

Private Sub Command1_Click()
  Dim rst As Recordset
  Dim sql As String
  sql = "SELECT refinplacsa, first(producte) as Pr,last(refclient) as Ref_, count(*) as Q,Max(datacomanda) AS maxdata, Max(comanda) AS maxcomanda From comandesmesextres where  ( client=6841)and producte<>'PC' and producte<>'PC2' and producte<>'PCP' and producte<>'PCI3'  GROUP BY refinplacsa ORDER BY Max(datacomanda) DESC;"
  'Clipboard.Clear
  'Clipboard.SetText sql
  vhora = Now
  Set rst = Data1.Database.OpenRecordset(sql)
  MsgBox Trim(DateDiff("s", vhora, Now)) + " Segons"
  If Not rst.EOF Then
     rst.MoveLast
     MsgBox rst.RecordCount
  End If
End Sub

Private Sub Command2_Click()
   Data1.Recordset.Edit
   Data1.Recordset!comanda = Data1.Recordset!comanda
   Data1.Recordset.Update
End Sub

Private Sub Command3_Click()
  Dim rst As Recordset
  'Set rst = dbcomandes.OpenRecordset("Select * from comandes")
  Set rst = Data1.Recordset
  While Not rst.EOF
     Me.Caption = Trim(rst!comanda)
     rst.MoveNext
  Wend
End Sub

Private Sub Command4_Click()
  While Not Data1.Recordset.EOF
    Data1.Recordset.MoveNext
    DoEvents
  Wend
End Sub

Private Sub Form_Load()
 ' Set sqlconn = DBEngine.OpenConnection("Serverprodu", 0, False, "Driver={SQL Server}; Server=serverprodu; uid=sa; pwd=Ipc123")
   Set dbcomandes = OpenDatabase("", 0, False, "Driver={SQL Server}; Server=serverprodu; Database=comandes--SQL2; uid=sa; pwd=Ipc123")
   Set Data1.Recordset = dbcomandes.OpenRecordset("comandes")
End Sub
