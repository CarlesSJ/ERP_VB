VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formcamises 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniment de camises"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14370
   Icon            =   "formcamises.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   14370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   465
      Left            =   1005
      Picture         =   "formcamises.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Borrar camisa"
      Top             =   1920
      Width           =   645
   End
   Begin VB.CommandButton afegir 
      Height          =   465
      Left            =   360
      Picture         =   "formcamises.frx":1254
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Afegir camises"
      Top             =   1920
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecció"
      Height          =   1035
      Left            =   180
      TabIndex        =   1
      Top             =   300
      Width           =   3345
      Begin VB.ComboBox combogrup 
         Height          =   315
         ItemData        =   "formcamises.frx":17DE
         Left            =   2265
         List            =   "formcamises.frx":17E0
         TabIndex        =   6
         Top             =   540
         Width           =   645
      End
      Begin VB.ComboBox combocolor 
         Height          =   315
         Left            =   1215
         TabIndex        =   4
         Top             =   555
         Width           =   900
      End
      Begin VB.ComboBox combodesarroll 
         Height          =   315
         Left            =   225
         TabIndex        =   2
         Top             =   570
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "Grup"
         Height          =   210
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Color"
         Height          =   210
         Left            =   1335
         TabIndex        =   5
         Top             =   255
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Desarroll"
         Height          =   210
         Left            =   240
         TabIndex        =   3
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.Data datacamises 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   480
      Left            =   7380
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "formcamises.frx":17E2
      Height          =   5310
      Left            =   315
      OleObjectBlob   =   "formcamises.frx":17F8
      TabIndex        =   0
      Top             =   2400
      Width           =   13935
   End
End
Attribute VB_Name = "formcamises"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub afegir_Click()
   Dim vQcamises As Double
   Dim vcolor As String
   Dim vGrup As String
   Dim vDesarroll As Double
   Dim rst As Recordset
   Dim vRef As String
   Dim vAny As Double
   Dim i As Byte
   
   vQcamises = cadbl(InputBox("Quantes camises d'aquest grup vols entrar?", "Quantitat"))
   If vQcamises = 0 Or vQcamises > 30 Then Exit Sub
   vDesarroll = cadbl(InputBox("Escriu el desarroll que vols crear les camises.", "Desarroll"))
   If vDesarroll = 0 Then Exit Sub
   vcolor = UCase(InputBox("Escriu el color que vols per la camisa. [B]-Blau  [N]-Negre.", "Color"))
   If vcolor <> "B" And vcolor <> "N" Then Exit Sub
   vGrup = UCase(InputBox("Escriu el grup de camises que vols crear. ex: [A] [B]", "Grup"))
   If Mid(vGrup + "", 1, 1) = "" Then MsgBox "Aquest grup no ès vàlid.", vbCritical, "Error": Exit Sub
   vRef = UCase(atrim(InputBox("Escriu la referència de les camises que vols crear. ex: [R64030-04]", "Referència")))
   If vRef = "" Then Exit Sub
   vAny = cadbl(InputBox("Escriu l'any de compra de les camises.", "Any"))
   If vAny = 0 Then Exit Sub
   Set rst = dbcomandes.OpenRecordset("select * from muntadora_camises where desarroll=" + atrim(vDesarroll) + " and color='" + vcolor + "' and grup='" + atrim(vGrup) + "'")
   If Not rst.EOF Then MsgBox "Ja hi ha les dades d'aquestes camises entrades.", vbCritical, "Error": GoTo fi
   For i = 1 To vQcamises
     rst.AddNew
     rst!desarroll = vDesarroll
     rst!color = vcolor
     rst!grup = vGrup
     rst!referencia = vRef
     rst!numcamisa = i
     rst!Any = vAny
     rst.Update
     
   Next i
fi:
   carregar_combodesarroll
   combodesarroll = vDesarroll
   combocolor = IIf(vcolor = "B", "Blau", "Negre")
   combogrup = vGrup
   filtrar_dades_camises
   
   Set rst = Nothing
End Sub

Private Sub combocolor_Click()
   carregar_combogrup
   combogrup = ""
End Sub

Private Sub combodesarroll_Click()
  carregar_combocolor
  combocolor = ""
  combogrup = ""
End Sub

Private Sub combogrup_Click()
   filtrar_dades_camises
End Sub
Sub filtrar_dades_camises()
   datacamises.DatabaseName = rutadelfitxer(cami) + "comandes.mdb"
   datacamises.RecordSource = "select * from muntadora_camises where desarroll=" + combodesarroll + " and color='" + Mid(combocolor + " ", 1, 1) + "' and grup='" + atrim(combogrup) + "'"
   datacamises.Refresh
End Sub

Private Sub Command1_Click()
   If MsgBox("Segur que vols eliminar aquesta camisa?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   datacamises.Recordset.Delete
   carregar_combodesarroll
End Sub

Private Sub Form_Load()
   carregar_combodesarroll
End Sub
Sub carregar_combodesarroll()
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select distinct desarroll from muntadora_camises order by desarroll asc")
   combodesarroll.Clear
   While Not rst.EOF
     combodesarroll.AddItem atrim(rst!desarroll)
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Sub carregar_combocolor()
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select distinct color from muntadora_camises order by color asc")
   combocolor.Clear
   While Not rst.EOF
     combocolor.AddItem IIf(atrim(rst!color) = "B", "BLAU", "NEGRE")
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Sub carregar_combogrup()
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select distinct grup from muntadora_camises order by grup asc")
   combogrup.Clear
   While Not rst.EOF
     combogrup.AddItem atrim(rst!grup)
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Private Sub reixa_DblClick()
  Dim v As String
  v = reixa.Text
  v = InputBox("Entra el valor d'aquest camp.", reixa.Columns(reixa.col).caption, v)
  If v <> reixa.Text And v <> "" Then
     reixa.EditActive = True
     reixa.Text = v
     reixa.EditActive = False
     reixa.Refresh
  End If
End Sub
