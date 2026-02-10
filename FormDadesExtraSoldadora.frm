VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormExtresSoldadores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dades Extres Soldadores"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9210
   Icon            =   "FormDadesExtraSoldadora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data dataaccessoris 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\baixes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4065
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "soldadores_accessorisutilitzats"
      Top             =   60
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H006BEBB1&
      Caption         =   "+  Afegir Accessori"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5055
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   525
      Width           =   1860
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "-  Eliminar Accessori"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5055
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1380
      Width           =   1860
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "FormDadesExtraSoldadora.frx":048A
      Height          =   4785
      Left            =   315
      OleObjectBlob   =   "FormDadesExtraSoldadora.frx":04A3
      TabIndex        =   0
      Top             =   405
      Width           =   4710
   End
End
Attribute VB_Name = "FormExtresSoldadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Dim vidaccessori As Double
  Dim vnomaccessori As String
  Dim vrefinplacsa As String
  Dim vcontrolLot As Boolean
  escullir_accessori vidaccessori, vnomaccessori, vrefinplacsa, vcontrolLot
  If vidaccessori > 0 Then
        dbbaixes.Execute "insert into soldadores_accessorisutilitzats (comanda,nomaccessori,idaccessori,refinplacsa,lottraçabilitat) values (" + atrim(formcomandes.Text1) + ",'" + atrim(treure_apostruf(vnomaccessori)) + "'," + atrim(vidaccessori) + ",'" + atrim(vrefinplacsa) + "'," + IIf(vcontrolLot, "'-'", "''") + ")"
        dataaccessoris.Refresh
  End If
End Sub

Private Sub Command3_Click()
If MsgBox("Segur que vols eliminar l'accessori " + vbNewLine + atrim(dataaccessoris.Recordset!nomaccessori) + "?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
  dbbaixes.Execute "delete * from  soldadores_accessorisutilitzats where id=" + atrim(dataaccessoris.Recordset!ID)
  dataaccessoris.Refresh
End Sub

Private Sub Form_Load()
  dataaccessoris.RecordSource = "Select * from soldadores_accessorisutilitzats where comanda=" + atrim(formcomandes.Text1)
  dataaccessoris.Refresh
End Sub
Sub escullir_accessori(vidaccessori As Double, vnomaccessori As String, vrefinplacsa As String, vcontrolLot As Boolean)
  Load formseleccio
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select numaccessori,familia,subfamilia,descripcio_curta,control_traçabilitat from accessoris_soldadora order by descripcio_curta"
  formseleccio.Caption = "Triar Accessori"
  formseleccio.refrescar
  formseleccio.Width = 12000
  formseleccio.DBGrid2.Columns(0).Width = 0
  formseleccio.DBGrid2.Columns(1).Width = 2700
  formseleccio.DBGrid2.Columns(2).Width = 2900
  formseleccio.DBGrid2.Columns(3).Width = 5000
  formseleccio.DBGrid2.Font.Size = 12
  formseleccio.Command3.Tag = "filtre"
  formseleccio.DBGrid2.col = 2
  formseleccio.Show 1
  If seleccioret = 1 Then
    vidaccessori = cadbl(atrim(formseleccio.Data1.Recordset!numaccessori))
    vnomaccessori = atrim(formseleccio.Data1.Recordset!descripcio_curta)
    vrefinplacsa = ""
    If formseleccio.Data1.Recordset!control_traçabilitat = True Then vcontrolLot = True
  End If
  Unload formseleccio
    
End Sub

