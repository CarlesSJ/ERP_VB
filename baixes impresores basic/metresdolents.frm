VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form metresdolents 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Metres impresos dolents"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   ControlBox      =   0   'False
   Icon            =   "metresdolents.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton sortir 
      Height          =   525
      Left            =   5835
      Picture         =   "metresdolents.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Sortida."
      Top             =   0
      Width           =   570
   End
   Begin VB.CommandButton alta 
      Height          =   450
      Left            =   120
      Picture         =   "metresdolents.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Alta  Registres"
      Top             =   60
      Width           =   405
   End
   Begin VB.CommandButton eliminar 
      Height          =   450
      Left            =   630
      Picture         =   "metresdolents.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   60
      Width           =   465
   End
   Begin VB.Data datadolents 
      Caption         =   "datadolents"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   975
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2790
      Visible         =   0   'False
      Width           =   2265
   End
   Begin MSDBGrid.DBGrid reixadolents 
      Bindings        =   "metresdolents.frx":1628
      Height          =   2400
      Left            =   60
      OleObjectBlob   =   "metresdolents.frx":163E
      TabIndex        =   0
      Top             =   570
      Width           =   6315
   End
End
Attribute VB_Name = "metresdolents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
   noupaletdolent
   reixadolents.Refresh
End Sub
Sub noupaletdolent()
   Dim palet As Double
   Dim bobina As Double
   Dim metres As Double
   Dim obs As String
   Dim rstdolents As Recordset
   palet = cadbl(InputBox("Entra el numero de PALET:", "Entrada metres dolents"))
   If cadbl(palet) = 0 Then GoTo fi
   bobina = cadbl(InputBox("Entra el numero de BOBINA:", "Entrada metres dolents"))
   If cadbl(bobina) = 0 Then GoTo fi
   Set rstdolents = dbstocks.OpenRecordset("select * from bobines where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
   If rstdolents.EOF Then MsgBox "Aquest palet i bobina no existeixen.": Exit Sub
   Set rstdolents = Nothing
   metres = InputBox("Entra el numero de METRES:", "Entrada metres dolents")
   If cadbl(metres) = 0 Then GoTo fi
   obs = InputBox("Entra la observació:" + Chr(10) + Chr(13) + "(50 Caracters màxim)", "Entrada metres dolents")
   obs = Mid(obs, 1, 49)
   obs = treure_apostruf(obs)
   datadolents.Recordset.AddNew
   datadolents.Recordset!idcomanda = cadbl(Form1.comanda)
   datadolents.Recordset!operari = numop
   datadolents.Recordset!palet = palet
   datadolents.Recordset!bobina = bobina
   datadolents.Recordset!metres = metres
   datadolents.Recordset!comentari = obs
   datadolents.Recordset.Update
   actualitzamtrsdolents palet, bobina, metres
   Exit Sub
fi:
    MsgBox "Valors no vàlids"
End Sub
Sub actualitzamtrsdolents(palet As Double, bobina As Double, metres As Double)
   Dim rstdolents As Recordset
   'Set rstdolents = dbstocks.OpenRecordset("select sum(metres) from parcials where comanda=400 and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
   'If rstdolents.EOF Then
   dbstocks.Execute "insert into parcials (idpalet,idbobina,metres,comanda,data,seccio,utilitzada,orcomassignacio,operari) values (" + atrim(palet) + "," + atrim(bobina) + "," + atrim((metres)) + ",'400',now,'" + lletraseccio + "',true,'" + atrim(cadbl(Form1.comanda)) + "'," + atrim(numop) + ")"
   'End If
End Sub

Private Sub eliminar_Click()
  Dim palet As String
  Dim bobina As String
  If datadolents.Recordset.EOF Then Exit Sub
  If MsgBox("Segur que vols eliminar aquest registre?", vbCritical + vbYesNo, "Atenció") = vbNo Then Exit Sub
  palet = datadolents.Recordset!palet
  bobina = datadolents.Recordset!bobina
  dbstocks.Execute "delete * from parcials where comanda='400' and idpalet=" + palet + " and idbobina=" + bobina + " and orcomassignacio='" + Form1.comanda + "'"
  datadolents.Recordset.Delete
  reixadolents.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 110 Then KeyCode = 188
End Sub

Private Sub Form_Load()
  datadolents.DatabaseName = Form1.impresores.DatabaseName
  datadolents.RecordSource = "select * from impresores_mtrsdolents where idcomanda=" + atrim(Form1.comanda)
  datadolents.Refresh
  obrestocks
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.calcular_metresdolents
End Sub

Private Sub reixadolents_KeyDown(KeyCode As Integer, Shift As Integer)
  If reixadolents.col = 3 And KeyCode > 46 Then
      If (Len(reixadolents.text)) > 49 Then reixadolents.text = Mid(reixadolents.text, 1, 49)
  End If
End Sub

Private Sub reixadolents_OnAddNew()
'   datadolents.Recordset!idcomanda = cadbl(Form1.comanda)
End Sub

Private Sub sortir_Click()
Form1.calcular_metresdolents
dbstocks.Close
Unload metresdolents
End Sub
