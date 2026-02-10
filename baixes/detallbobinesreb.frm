VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form detallbobreb 
   Caption         =   "Detall de Bobines Rebobinadores"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   Icon            =   "detallbobinesreb.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Saltar a Kilos"
      Height          =   240
      Left            =   4155
      TabIndex        =   6
      Top             =   5385
      Width           =   1545
   End
   Begin VB.CommandButton sortir 
      Height          =   450
      Left            =   5145
      Picture         =   "detallbobinesreb.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Sortir a Menú"
      Top             =   45
      Width           =   450
   End
   Begin VB.CommandButton alta 
      Height          =   450
      Left            =   255
      Picture         =   "detallbobinesreb.frx":0944
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Alta  Registres"
      Top             =   30
      Width           =   450
   End
   Begin VB.CommandButton eliminar 
      Height          =   450
      Left            =   840
      Picture         =   "detallbobinesreb.frx":0D76
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   45
      Width           =   465
   End
   Begin VB.CommandButton gravar 
      Height          =   450
      Left            =   4635
      Picture         =   "detallbobinesreb.frx":1088
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Guardar Registres"
      Top             =   45
      Width           =   450
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4365
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "bobinesreb"
      Top             =   75
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "detallbobinesreb.frx":13CA
      Height          =   4860
      Left            =   210
      OleObjectBlob   =   "detallbobinesreb.frx":13DA
      TabIndex        =   0
      Top             =   510
      Width           =   5445
   End
   Begin VB.Label Label11 
      Caption         =   "Prem F2 per sel.leccionar Taules..."
      Height          =   225
      Left            =   255
      TabIndex        =   5
      Top             =   5370
      Width           =   3570
   End
End
Attribute VB_Name = "detallbobreb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub nova_bobina()
  Dim rstmp As Recordset
  Dim rsttmp2 As Recordset
  Dim col As Byte
  Dim elgran As Double
  DBGrid1.Tag = "afegint"
  If Not data1.Recordset.EOF Then
   If data1.Recordset.EditMode = 0 Then data1.Recordset.Edit
   data1.Recordset.Update
  End If
  Set rsttmp2 = dbtmpb.OpenRecordset("select id  from rebobinadores where comanda=" + atrim(Rebobinadores.bobines.Recordset!comanda))
  elgran = 0
  While Not rsttmp2.EOF
   Set rstmp = dbtmpb.OpenRecordset("select max(numerodebobina) as elgran from bobinesreb where controlid=" + atrim(rsttmp2!id))
   If Not rstmp.EOF Then
      If cadbl(rstmp!elgran) > elgran Then elgran = cadbl(rstmp!elgran)
   End If
   rsttmp2.MoveNext
  Wend
  Set rstmp = dbtmpb.OpenRecordset("select * from bobinesreb where controlid=" + atrim(Rebobinadores.bobines.Recordset!id) + " and numerodebobina=" + atrim(elgran))
  data1.Recordset.AddNew
  data1.Recordset!numerodebobina = elgran + 1
  data1.Recordset!controlid = atrim(Rebobinadores.bobines.Recordset!id)
  data1.Recordset!numempalmes = 0
  col = 0
  If Not rstmp.EOF Then
     data1.Recordset!operari1 = rstmp!operari1
     data1.Recordset!metres = rstmp!metres
     data1.Recordset!kilos = rstmp!kilos
     If rstmp!operari1 > 0 Then col = IIf(Check1.Value = 1, 2, 3)
     'col = 2
   Else: col = 0
  End If
  data1.Recordset.Update
  'Data1.Refresh
  DBGrid1.Refresh
  data1.Recordset.MoveLast
  'DBGrid1.Refresh
  DoEvents
  DBGrid1.col = col
  focusreixa
  Set rstmp = Nothing
If DBGrid1.Text = "0" Then DBGrid1.SelLength = Len(DBGrid1.Text)
DBGrid1.Tag = ""
End Sub

Sub focusreixa()
 If ActiveControl.Name <> "DBGrid1" Then
   DBGrid1.SetFocus
 End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub alta_Click()
 nova_bobina
End Sub

Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
DBGrid1_RowColChange DBGrid1.Row, DBGrid1.col
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim saltacol As Byte
  saltacol = 3
'  If Check1.Value = 1 Then saltacol = 2
  If KeyCode = 13 Then
    KeyCode = 0
    If DBGrid1.col = saltacol Then
       'DBGrid1_RowColChange DBGrid1.Row, DBGrid1.col
       nova_bobina
     Else: SendKeys "{TAB}"
    End If
  End If
  If KeyCode = 113 And DBGrid1.col = 0 Then
    triaroperaris
  End If
End Sub

Private Sub DBGrid1_LostFocus()
'If cadbl(DBGrid1.Columns(2).Text) = 0 And Check1 = 0 Then
'If Not data1.Recordset.EOF Then data1.Recordset.Delete
'     data1.Refresh
'     DBGrid1.Refresh
'End If
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

   'comprova si hem escrit el numero amb separat per .
  If LastCol >= 0 Then
   If IsNumeric(DBGrid1.Columns(LastCol)) Then
      If InStr(1, DBGrid1.Columns(LastCol), ".") Then
         DBGrid1.Columns(LastCol) = Mid(DBGrid1.Columns(LastCol), 1, InStr(1, DBGrid1.Columns(LastCol), ".") - 1) + "," + Mid(DBGrid1.Columns(LastCol), InStr(1, DBGrid1.Columns(LastCol), ".") + 1)
      End If
   End If
  End If

    If DBGrid1.Text = "0" Then DBGrid1.SelLength = Len(DBGrid1.Text)
    If LastCol = 0 Then
      If cadbl(DBGrid1.Columns(LastCol)) <> 0 Then
       Set rsttmp = dbtmp.OpenRecordset("select codi from operaris where maquina='R' and codi=" + atrim(cadbl(DBGrid1.Columns(LastCol))))
       If rsttmp.EOF Then MsgBox "Aquest Operari no Existeix": DBGrid1.Columns(LastCol) = "": DBGrid1.col = LastCol
      End If
    End If
    If LastCol >= 0 Then
      If atrim(DBGrid1.Columns(LastCol).Text) = "" Then DBGrid1.Columns(LastCol).Text = "0"
    End If
End Sub

Private Sub eliminar_Click()
  If MsgBox("Segur que vols borrar aquesta bobina?", vbCritical + 4, "Atenció") = vbYes Then
     If Not data1.Recordset.EOF Then data1.Recordset.Delete
     data1.Refresh
     DBGrid1.Refresh
  End If
End Sub

Private Sub Form_Activate()
 If DBGrid1.Tag = "primera" Then
    DBGrid1.Tag = ""
    'nova_bobina
    
 End If
End Sub
Sub triaroperaris()
  Load formseleccio
  formseleccio.Caption = "Triar Operaris"
  formseleccio.data1.DatabaseName = camicomandes
  formseleccio.data1.RecordSource = "select * from operaris where maquina='R'"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   DBGrid1.Text = atrim(formseleccio.data1.Recordset!codi)
  ' nomextrussora(0).Caption = atrim(formseleccio.data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then sortir_Click
  If KeyCode = 112 Then gravar_Click
End Sub

Private Sub Form_Load()
centerscreen Me
DBGrid1.Tag = "primera"
data1.DatabaseName = cami
data1.RecordSource = "select * from bobinesreb where controlid=" + atrim(Rebobinadores.bobines.Recordset!id) + " order by numerodebobina"

data1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Set dbtmpb = Nothing
End Sub

Private Sub gravar_Click()
 If data1.Recordset.EditMode = 0 Then If Not data1.Recordset.EOF Then data1.Recordset.Edit
If data1.Recordset.EditMode > 0 Then data1.Recordset.Update
End Sub

Private Sub sortir_Click()
  DBGrid1_LostFocus
  Unload detallbobreb
End Sub

