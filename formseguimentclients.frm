VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formseguimentclients 
   Caption         =   "Seguiment de clients"
   ClientHeight    =   11160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19005
   LinkTopic       =   "Form1"
   ScaleHeight     =   11160
   ScaleWidth      =   19005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Height          =   360
      Left            =   450
      Picture         =   "formseguimentclients.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Alta  Registres"
      Top             =   5685
      Width           =   420
   End
   Begin VB.ComboBox combosino 
      Height          =   315
      ItemData        =   "formseguimentclients.frx":058A
      Left            =   9570
      List            =   "formseguimentclients.frx":0594
      TabIndex        =   5
      Top             =   7230
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.ComboBox Comboaccio 
      Height          =   315
      ItemData        =   "formseguimentclients.frx":05A5
      Left            =   2490
      List            =   "formseguimentclients.frx":05C1
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.Data datalinies 
      Caption         =   "datalinies"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4935
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from clients_seguiment_linies"
      Top             =   5655
      Visible         =   0   'False
      Width           =   5130
   End
   Begin MSDBGrid.DBGrid reixalinies 
      Bindings        =   "formseguimentclients.frx":061B
      Height          =   4890
      Left            =   360
      OleObjectBlob   =   "formseguimentclients.frx":0630
      TabIndex        =   3
      Top             =   6060
      Width           =   18435
   End
   Begin VB.CommandButton alta 
      Height          =   360
      Left            =   360
      Picture         =   "formseguimentclients.frx":154B
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Alta  Registres"
      Top             =   120
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   4095
      Picture         =   "formseguimentclients.frx":1AD5
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   180
      Width           =   2745
   End
   Begin VB.Data dataseguiment 
      Caption         =   "dataseguiment"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   870
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from clients_seguiment order by nomempresa"
      Top             =   75
      Visible         =   0   'False
      Width           =   2850
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "formseguimentclients.frx":205F
      Height          =   5055
      Left            =   270
      OleObjectBlob   =   "formseguimentclients.frx":2077
      TabIndex        =   0
      Top             =   510
      Width           =   18375
   End
End
Attribute VB_Name = "formseguimentclients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
  reixa.AllowAddNew = True
  dataseguiment.Recordset.AddNew
  reixa.SetFocus
End Sub

Private Sub Comboaccio_Click()
  If datalinies.Recordset.EditMode > 0 Then reixalinies.Text = Comboaccio
End Sub

Private Sub combosino_Click()
   If datalinies.Recordset.EditMode > 0 Then reixalinies.Text = combosino
End Sub

Private Sub Command1_Click()
   Dim v As String
   
   v = InputBox("Escriu el nom de l'empresa que busques.", "Nom empresa")
   If atrim(v) = "" Then Exit Sub
   dataseguiment.RecordSource = "select * from clients_seguiment where nomempresa like '*" + atrim(v) + "*' order by nomempresa"
   dataseguiment.Refresh
End Sub

Private Sub Command2_Click()
   datalinies.Recordset.AddNew
   reixalinies.AllowAddNew = True
   reixalinies.col = 0
   datalinies.Recordset!ID = dataseguiment.Recordset!ID
   reixalinies.SetFocus
End Sub

Private Sub reixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   datalinies.RecordSource = "select * from clients_seguiment_linies where id=" + atrim(dataseguiment.Recordset!ID)
   datalinies.Refresh
End Sub

Private Sub reixalinies_BeforeUpdate(Cancel As Integer)
    datalinies.Recordset!ID = dataseguiment.Recordset!ID
End Sub

Private Sub reixalinies_ButtonClick(ByVal ColIndex As Integer)
If datalinies.Recordset.EOF Then Exit Sub
If reixalinies.Columns(ColIndex).DataField = "accio" Then
    If datalinies.Recordset.EditMode = 0 Then datalinies.Recordset.Edit
    Comboaccio.Left = reixalinies.Columns(ColIndex).Left + reixalinies.Left
    Comboaccio.Top = ((reixalinies.RowHeight) * (reixalinies.row + 1)) + reixalinies.Top - 70
    Comboaccio.Width = reixalinies.Columns(ColIndex).Width
   ' Comboaccio.Height = reixalinies.RowHeight
    Comboaccio.Visible = True
End If
If reixalinies.Columns(ColIndex).DataField = "tancat" Then
    If datalinies.Recordset.EditMode = 0 Then datalinies.Recordset.Edit
    combosino.Left = reixalinies.Columns(ColIndex).Left + reixalinies.Left
    combosino.Top = ((reixalinies.RowHeight) * (reixalinies.row + 1)) + reixalinies.Top - 70
    combosino.Width = reixalinies.Columns(ColIndex).Width
   ' combosino.Height = reixalinies.RowHeight
    combosino.Visible = True
End If
End Sub

Private Sub reixalinies_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   Comboaccio.Visible = False
   combosino.Visible = False
   reixalinies_ButtonClick reixalinies.col
End Sub

Private Sub Form_Load()
   datalinies.DatabaseName = rutadelfitxer(cami) + "comandes.mdb"
   dataseguiment.DatabaseName = rutadelfitxer(cami) + "comandes.mdb"
   dataseguiment.RecordSource = " select * from clients_seguiment order by nomempresa"
End Sub

