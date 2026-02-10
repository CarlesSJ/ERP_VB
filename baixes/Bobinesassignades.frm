VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Bobinesassignades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bobines Assignades"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   Icon            =   "Bobinesassignades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Sel. Totes"
      Height          =   225
      Left            =   4080
      TabIndex        =   9
      Top             =   15
      Width           =   1170
   End
   Begin VB.Data parcials 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "W:\progcomandes\dades\Palets.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   -30
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Parcials"
      Top             =   4065
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Passar totes a utilitzades"
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4110
      Width           =   1470
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "Bobinesassignades.frx":058A
      Height          =   3765
      Left            =   105
      OleObjectBlob   =   "Bobinesassignades.frx":059D
      TabIndex        =   0
      Top             =   255
      Width           =   5130
   End
   Begin VB.Frame Frame1 
      Height          =   720
      Left            =   690
      TabIndex        =   2
      Top             =   4005
      Width           =   2730
      Begin VB.ComboBox seccions 
         Height          =   315
         ItemData        =   "Bobinesassignades.frx":1476
         Left            =   2115
         List            =   "Bobinesassignades.frx":1489
         TabIndex        =   7
         Top             =   285
         Width           =   540
      End
      Begin VB.TextBox data 
         Height          =   375
         Left            =   900
         TabIndex        =   5
         Top             =   285
         Width           =   1140
      End
      Begin VB.TextBox operari 
         Height          =   375
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   255
         Left            =   1215
         TabIndex        =   6
         Top             =   105
         Width           =   540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Operari"
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   90
         Width           =   825
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fes doble clic a sobre de la bobina per sel.leccionar-la."
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   30
      Width           =   4575
   End
End
Attribute VB_Name = "Bobinesassignades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Command1_Click()
  If Not IsDate(data) Then MsgBox "Data no valida": Exit Sub
  If Len(seccions) <> 1 Then MsgBox "Falta escullir la seccio.": Exit Sub
  parcials.Refresh
  If parcials.Recordset.EOF Then MsgBox "No hi ha bobines per modificar.": Exit Sub
  If parcials.Recordset.EditMode > 0 Then parcials.Recordset.Update
  parcials.Refresh
  While Not parcials.Recordset.EOF
    If parcials.Recordset!seccio = "S" Then
     parcials.Recordset.Edit
     parcials.Recordset!operari = cadbl(operari)
     parcials.Recordset!data = data
     parcials.Recordset!seccio = seccions
     parcials.Recordset!utilitzada = True
     parcials.Recordset.Update
    End If
    parcials.Recordset.MoveNext
  Wend
  parcials.Refresh
  Unload Bobinesassignades
End Sub

Private Sub Command2_Click()
  parcials.Refresh
  While Not parcials.Recordset.EOF
     parcials.Recordset.Edit
     parcials.Recordset!seccio = "S"
     parcials.Recordset.Update
    parcials.Recordset.MoveNext
  Wend
End Sub

Private Sub Form_Load()
  Dim rstc As Recordset
  Dim comandes As String
  Set rstc = dbtmp.OpenRecordset("select linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(entradabaixes.comanda))
  If Not rstc.EOF Then
      comandes = IIf(cadbl(rstc!linkcomanda1) > 0, " or comanda='" + atrim(rstc!linkcomanda1) + "'", "") + IIf(cadbl(rstc!linkcomanda2) > 0, " or comanda='" + atrim(rstc!linkcomanda2) + "'", "")
  End If
  comandes = comandes + ")"
  parcials.DatabaseName = rutadelfitxer(camicomandes) + "palets.mdb"
  parcials.RecordSource = "select * from parcials where  not utilitzada and (comanda='" + entradabaixes.comanda + "' " + comandes
  parcials.Refresh
  data = Format(Now, "dd/mm/yy")
  operari = "10"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  parcials.Refresh
  While Not parcials.Recordset.EOF
    If parcials.Recordset!seccio = "S" Then
      parcials.Recordset.Edit
      parcials.Recordset!seccio = ""
      parcials.Recordset.Update
    End If
    parcials.Recordset.MoveNext
  Wend
End Sub

Private Sub reixa_DblClick()
   reixa.Columns("sel") = IIf(reixa.Columns("sel") = "S", "", "S")
End Sub
