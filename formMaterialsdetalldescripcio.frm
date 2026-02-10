VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formMaterialsdetalldescripcio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Linies de detall del material (Compres)"
   ClientHeight    =   4590
   ClientLeft      =   1575
   ClientTop       =   1470
   ClientWidth     =   6150
   Icon            =   "formMaterialsdetalldescripcio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox comborefinplacsa 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   255
      Width           =   2430
   End
   Begin VB.Data dataliniesmaterial 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3705
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from materials_liniesdescripcio"
      Top             =   150
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CommandButton alta 
      Height          =   360
      Left            =   195
      Picture         =   "formMaterialsdetalldescripcio.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Alta  Registres"
      Top             =   225
      Width           =   420
   End
   Begin VB.CommandButton eliminar 
      Height          =   360
      Left            =   630
      Picture         =   "formMaterialsdetalldescripcio.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Eliminacio Registres"
      Top             =   225
      Width           =   420
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "formMaterialsdetalldescripcio.frx":109E
      Height          =   3615
      Left            =   165
      OleObjectBlob   =   "formMaterialsdetalldescripcio.frx":10BB
      TabIndex        =   0
      Top             =   645
      Width           =   5850
   End
   Begin VB.Label Label1 
      Caption         =   "Referencia Inplacsa"
      Height          =   255
      Left            =   1710
      TabIndex        =   4
      Top             =   45
      Width           =   2145
   End
End
Attribute VB_Name = "formMaterialsdetalldescripcio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
   If comborefinplacsa = "" Then MsgBox "No hi ha refinplacsa entrada": Exit Sub
   If Not existeixrefinplacsa(comborefinplacsa) Then MsgBox "La Referència d'inplacsa " + comborefinplacsa + " no existeix.", vbCritical, "Error": Exit Sub
   If Not dataliniesmaterial.Recordset.EOF Then dataliniesmaterial.Recordset.MoveLast
   dataliniesmaterial.Recordset.AddNew
   dataliniesmaterial.Recordset!codimaterial = cadbl(formMaterialsdetalldescripcio.Tag)
   dataliniesmaterial.Recordset!refinplacsa = UCase(atrim(comborefinplacsa))
   dataliniesmaterial.Recordset.Update
   dataliniesmaterial.Recordset.MoveLast
   reixa.SetFocus
End Sub
Function existeixrefinplacsa(vref As String) As Boolean
   Dim rst As Recordset
   vref = UCase(vref)
   Set rst = dbtmp.OpenRecordset("Select * from comandes_extres where refinplacsa='" + atrim(vref) + "'")
   If Not rst.EOF Then existeixrefinplacsa = True
   Set rst = Nothing
End Function


Private Sub comborefinplacsa_Click()
  actualitzardades
End Sub

Private Sub comborefinplacsa_GotFocus()
  carregar_combo_refinplacsa Me.Tag
End Sub

Private Sub comborefinplacsa_LostFocus()
actualitzardades
End Sub

Private Sub eliminar_Click()
   If dataliniesmaterial.Recordset.EOF Then MsgBox "No hi ha cap linies escullida.", vbCritical, "Error": Exit Sub
   If MsgBox("Segur que vols eliminar aquesta linia?", vbDefaultButton2 + vbYesNo + vbExclamation, "Atenció") = vbNo Then Exit Sub
   dataliniesmaterial.Recordset.Delete
   dataliniesmaterial.Refresh
End Sub
Sub actualitzardades()
   dataliniesmaterial.DatabaseName = cami
   dataliniesmaterial.RecordSource = "Select * from materials_liniesdescripcio where codimaterial=" + atrim(formMaterialsdetalldescripcio.Tag) + " and refinplacsa='" + atrim(comborefinplacsa) + "'"
   dataliniesmaterial.Refresh
End Sub
Private Sub Form_Activate()
carregar_combo_refinplacsa Me.Tag
End Sub

Sub carregar_combo_refinplacsa(vcodimat As String)
  Dim rst As Recordset
  Dim vultim As String
  Set rst = dbtmp.OpenRecordset("Select distinct refinplacsa from materials_liniesdescripcio where codimaterial=" + atrim(vcodimat))
  vultim = comborefinplacsa
  comborefinplacsa.Clear
  While Not rst.EOF
     If atrim(rst!refinplacsa) <> "" Then comborefinplacsa.AddItem atrim(rst!refinplacsa)
     rst.MoveNext
  Wend
  If comborefinplacsa.ListCount > 0 Then comborefinplacsa.ListIndex = 0
  If vultim <> "" Then comborefinplacsa = vultim
  Set rst = Nothing
End Sub
