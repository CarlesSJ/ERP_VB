VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formtotals 
   Caption         =   "Totals Producció"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   Icon            =   "formtotals.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4995
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   135
      Visible         =   0   'False
      Width           =   1740
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "formtotals.frx":0442
      Height          =   8610
      Left            =   15
      OleObjectBlob   =   "formtotals.frx":0452
      TabIndex        =   0
      Top             =   60
      Width           =   10425
   End
End
Attribute VB_Name = "formtotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DBGrid1_DblClick()
 formcomandes.Tag = "100"
 formtotals.Caption = "Situant el registre actual..."
 While formcomandes.Data1.Recordset!comanda <> Data1.Recordset!comanda
   If formcomandes.Data1.Recordset!comanda > Data1.Recordset!comanda Then
      formcomandes.Data1.Recordset.MoveNext
     Else: formcomandes.Data1.Recordset.MovePrevious
   End If
   DoEvents
   If formcomandes.Data1.Recordset.EOF Or formcomandes.Data1.Recordset.BOF Then Exit Sub
 Wend
 formcomandes.Tag = ""
 formcomandes.carregar_lookups
 Unload formtotals
End Sub

Private Sub Form_Activate()
 Set rsttmp = dbtmp.OpenRecordset("select comanda from comandes " + querywhere + " order by comanda desc", , dbReadOnly )
With rsttmp
  formtotals.Caption = "Carregant"
While Not .EOF And cont < 3000
  selcomandes = selcomandes + IIf(selcomandes <> "", " or ", "") + "comanda = " + atrim(cadbl(!comanda))
  .MoveNext
  cont = cont + 1
  If (cont Mod 8) = 0 Then
   DoEvents
   formtotals.Caption = formtotals.Caption + "."
   If Len(formtotals.Caption) = 100 Then formtotals.Caption = "Carregant"
  End If
Wend
Set rsttmp = Nothing
End With
Data1.DatabaseName = llegir_ini("General", "camibaixes", fitxerini)
Data1.RecordSource = recordsourcetotals + " where " + selcomandes + " order by comanda desc"
Data1.Refresh
DBGrid1.Refresh
If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
  Data1.Recordset.MoveLast
  Data1.Recordset.MoveFirst
End If
formtotals.Caption = "Totals Producció "
If cont = 3000 Then formtotals.Caption = "Totals Producció " + " (ATENCIÓ Superat el límit de 3000) "
formtotals.Caption = formtotals.Caption + "  Registres: " + atrim(Data1.Recordset.RecordCount)
For i = 0 To DBGrid1.Columns.Count - 1
  DBGrid1.Columns(i).Caption = UCase(DBGrid1.Columns(i).Caption)
Next i
arreglar_reixa
End Sub
Sub arreglar_reixa()
  Dim colu As Byte
  colu = cadbl(formtotals.Tag)
  If colu = 0 Then colu = 4
  On Error Resume Next
  For i = 1 To colu
     DBGrid1.Columns(i).Width = DBGrid1.Columns(i).Width / 1.5
  Next i
  
End Sub
Private Sub Form_Load()
 Dim selcomandes As String
 Dim cont As Integer
centerscreen Me

End Sub
