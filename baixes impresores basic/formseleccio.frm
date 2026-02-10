VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formseleccio 
   BackColor       =   &H80000005&
   Caption         =   "c"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6060
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bimprimir 
      Caption         =   "Pack'list"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   8970
      Picture         =   "formseleccio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2310
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   39.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   8970
      Picture         =   "formseleccio.frx":033B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1425
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Height          =   810
      Left            =   8985
      Picture         =   "formseleccio.frx":0BE0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   570
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "formseleccio.frx":1419
      Height          =   5925
      Left            =   90
      OleObjectBlob   =   "formseleccio.frx":1429
      TabIndex        =   0
      Top             =   135
      Width           =   8805
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1275
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1965
   End
End
Attribute VB_Name = "formseleccio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub refrescar()
 Dim tipusdato As Byte
 Dim grandoto As Integer
 Dim espais As Byte

 Data1.Refresh
 DBGrid2.Refresh
 DBGrid2.ReBind
 DBGrid2.AllowUpdate = False
 On Error GoTo fi
 For i = 0 To 50
   tipusdato = Data1.Recordset.Fields(DBGrid2.Columns(i).DataField).Type
   grandato = Data1.Recordset.Fields(DBGrid2.Columns(i).DataField).Size
   If grandato < 5 Then grandato = 5
   DBGrid2.Columns(i).width = grandato * 200
   DBGrid2.Columns(i).caption = UCase(DBGrid2.Columns(i).caption)
    
 Next i
fi:
If formseleccio.tag = "1" Then DBGrid2.Columns(0).width = 0
End Sub



Private Sub Command2_Click()
 
End Sub


Function possarfiltre()
   Dim va As String
   Dim res As String
   Dim andowhere As String
   va = formseleccio.tag
   andowhere = " where "
   If InStr(1, va, "where") Then andowhere = " and "
   res = va
   If InStr(1, va, "order") Then res = Mid(va, 1, InStr(1, va, "order") - 1)
   res = res + andowhere + (DBGrid2.Columns(DBGrid2.col).DataField + " like '*" + Text1.text + "*'")
   If InStr(1, va, "order") Then res = Mid(va, InStr(1, va, "order"))
   possarfiltre = res
End Function

Private Sub bimprimir_Click()
  seleccioret = 5
  dbtmpb.Execute "update muntadoratot set packinglistimpresaimpresores=true where comanda=" + atrim(cadbl(Data1.Recordset!num_comanda))
  Me.Hide
End Sub

Private Sub Command1_Click()
  seleccioret = 0
  Me.Hide
End Sub

Private Sub Command3_Click()
acceptar
End Sub

Private Sub Data1_Reposition()
 If DBGrid2.tag = "" Then DBGrid2.tag = Data1.RecordSource

End Sub

Sub comprovarsijashaimpresuncop(vnumc As Double)
   Dim vcolorimpres As Double
   Dim vcolornormal As Double
   Dim rst As Recordset
   
   vcolorimpres = &H5C31DD
   vcolornormal = &H8000000F
   bimprimir.BackColor = vcolornormal
   Set rst = dbtmpb.OpenRecordset("select packinglistimpresaimpresores from muntadoratot where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
       If rst!packinglistimpresaimpresores Then bimprimir.BackColor = vcolorimpres
   End If
   Set rst = Nothing
End Sub
Private Sub DBGrid2_DblClick()
  acceptar
End Sub

Private Sub dbgrid2_HeadClick(ByVal ColIndex As Integer)
  Data1.RecordSource = DBGrid2.tag + " order by " + DBGrid2.Columns(ColIndex).DataField
  refrescar
End Sub

Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
 If KeyCode = 38 And DBGrid2.row = 0 Then Text1.SetFocus
End Sub

Private Sub DBGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then acceptar
End Sub

Private Sub DBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 If bimprimir.tag = "comprovarimpresio" And Not Data1.Recordset.EOF Then
       comprovarsijashaimpresuncop cadbl(Data1.Recordset!num_comanda)
   End If
End Sub

Private Sub Form_Activate()
If DBGrid2.Columns.Count > 1 Then
 DBGrid2.col = 1
 DBGrid2.SetFocus
 DoEvents
End If


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 110 Then KeyCode = 188
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
End Sub

Private Sub Form_Load()
  centerscreen Me
 
End Sub

Private Sub Form_Resize()
'On Error Resume Next
'DBGrid2.Width = formseleccio.Width - 200
'Frame1.Width = formseleccio.Width - 400
''DBGrid2.Height = formseleccio.Height - DBGrid2.Top - 700
End Sub

Private Sub Form_Unload(Cancel As Integer)
seleccioret = 0
End Sub


Sub acceptar()
  seleccioret = 1
  Me.Hide
End Sub



Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 
End Sub
