VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formtintessemblants 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tintes semblants"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12885
   Icon            =   "formtintessemblants.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datasemblants 
      Caption         =   "datasemblants"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2910
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   180
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Frame Frame8 
      Height          =   2445
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12795
      Begin VB.CommandButton Command30 
         Height          =   330
         Left            =   435
         Picture         =   "formtintessemblants.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Eliminar registre"
         Top             =   195
         Width           =   345
      End
      Begin VB.CommandButton Command29 
         Height          =   330
         Left            =   75
         Picture         =   "formtintessemblants.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Nova tinta"
         Top             =   195
         Width           =   345
      End
      Begin MSDBGrid.DBGrid reixasemblants 
         Bindings        =   "formtintessemblants.frx":109E
         Height          =   1770
         Left            =   60
         OleObjectBlob   =   "formtintessemblants.frx":10B6
         TabIndex        =   3
         Top             =   525
         Width           =   12600
      End
   End
End
Attribute VB_Name = "formtintessemblants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function escullir_tinta(vcodi As String, vdescripcio As String) As String
  Load formseleccio
  formseleccio.caption = "Selecciona una tinta"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "select codi,descripcio,referenciacolor from tintes order by descripcio"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = True
  formseleccio.DBGrid2.Columns(1).width = 4500
  formseleccio.DBGrid2.Columns(2).width = 800
  formseleccio.width = 10000
  formseleccio.Show 1
  If seleccioret = 1 Then
   vcodi = atrim(formseleccio.Data1.Recordset!codi)
   vdescripcio = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Function
Private Sub Command29_Click()
   'afegir tinta nova
   Dim vdescripcio As String
   Dim vcodi As String
   Dim vobservacions As String
   escullir_tinta vcodi, vdescripcio
   If vcodi <> "" Then
     vobservacions = UCase(InputBox("Si vols pots escriure una observació sobre aquesta relació", "Observacions"))
     datasemblants.Recordset.AddNew
     datasemblants.Recordset!coditintarelacio = atrim(formtintessemblants.tag)
     datasemblants.Recordset!coditinta = vcodi
     datasemblants.Recordset!nomdelatinta = vdescripcio
     datasemblants.Recordset!observacions = vobservacions
     datasemblants.Recordset.Update
   End If
   
End Sub

Private Sub Command30_Click()
  If Not datasemblants.Recordset.EOF And Not datasemblants.Recordset.BOF Then
      If MsgBox("Segur que vols borrar aquest registre?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
           datasemblants.Recordset.Delete
           datasemblants.Refresh
      End If
  End If
End Sub

Private Sub Form_Activate()
  datasemblants.DatabaseName = camitintes
  datasemblants.RecordSource = "SELECT * from tintes_semblants where coditintarelacio='" + atrim(formtintessemblants.tag) + "' order by nomdelatinta"
  datasemblants.Refresh
End Sub

Private Sub reixasemblants_DblClick()
  Dim vobservacions As String
  If reixasemblants.col = 2 Then
     vobservacions = UCase(InputBox("Si vols pots escriure una observació sobre aquesta relació" + Chr(10) + "PER NO POSAR CAP FES UN ESPAI EN BLANC", "Observacions", reixasemblants.Text))
     If vobservacions <> "" Then
        datasemblants.Recordset.Edit
        datasemblants.Recordset!observacions = atrim(vobservacions)
        datasemblants.Recordset.Update
     End If
  End If
End Sub
