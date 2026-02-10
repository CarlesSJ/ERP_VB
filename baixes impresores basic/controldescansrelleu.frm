VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formdescansirelleu 
   BackColor       =   &H00FDDECE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Descans i Relleu"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   Icon            =   "controldescansrelleu.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ctotes 
      BackColor       =   &H00FDDECE&
      Caption         =   "Totes"
      Height          =   195
      Left            =   5055
      TabIndex        =   5
      Top             =   960
      Width           =   1125
   End
   Begin VB.Data datacontroldescansirelleu 
      Caption         =   "datacontroldescansirelleu"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2895
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2670
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Acabar"
      Height          =   555
      Left            =   3135
      Picture         =   "controldescansrelleu.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   450
      Width           =   1770
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nou "
      Height          =   555
      Left            =   930
      Picture         =   "controldescansrelleu.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   450
      Width           =   1770
   End
   Begin MSDBGrid.DBGrid reixa 
      Align           =   2  'Align Bottom
      Bindings        =   "controldescansrelleu.frx":109E
      Height          =   2625
      Left            =   0
      OleObjectBlob   =   "controldescansrelleu.frx":10C2
      TabIndex        =   1
      ToolTipText     =   "Si vols fer canvi de dades fes dos clicks sobre el camp."
      Top             =   1200
      Width           =   6165
   End
   Begin VB.Label etnonou 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   60
      TabIndex        =   4
      Top             =   1005
      Width           =   4005
   End
   Begin VB.Label etnomoperari 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOM OPERARI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   6060
   End
   Begin VB.Menu mnou 
      Caption         =   "Nou"
      Visible         =   0   'False
      Begin VB.Menu mnoudescans 
         Caption         =   "Nou Descans"
      End
      Begin VB.Menu mnourelleu 
         Caption         =   "Nou Relleu"
      End
   End
End
Attribute VB_Name = "formdescansirelleu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Me.PopupMenu mnou, , Command1.Left + Command1.Width, Command1.Top
End Sub

Sub carregardades(Optional instsql As String)
   If instsql = "" Then instsql = "select * from controldescansrelleu where nummaq=" + atrim(nummaq) + " and operari=" + etnomoperari.Tag + " and seccio='" + atrim(lletraseccio) + "' and (hores=0 or hores=null) order by datainici,horainici"
   datacontroldescansirelleu.RecordSource = instsql
'   datacontroldescansirelleu.RecordSource = "controldescansrelleu"
   datacontroldescansirelleu.Refresh
   
End Sub

Private Sub Command2_Click()
   datacontroldescansirelleu.Recordset.FindFirst "hores=0 or hores=null"
   If datacontroldescansirelleu.Recordset.NoMatch Then MsgBox "No hi ha cap entrada sense finalitzar.", vbCritical, "Error": Exit Sub
   If MsgBox("Vols donar per acabat aquest " + atrim(datacontroldescansirelleu.Recordset!tipus) + "?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   datacontroldescansirelleu.Recordset.Edit
   datacontroldescansirelleu.Recordset!datafi = Date
   datacontroldescansirelleu.Recordset!horafi = Time
   datacontroldescansirelleu.Recordset!hores = calcularhores(datacontroldescansirelleu.Recordset)
   datacontroldescansirelleu.Recordset!comandafi = ncomanda
   datacontroldescansirelleu.Recordset.Update
   datacontroldescansirelleu.Recordset.Bookmark = datacontroldescansirelleu.Recordset.LastModified
   If datacontroldescansirelleu.Recordset!hores = 0 Then MsgBox "La diferencia entre inici i fi es massa curta" + Chr(10) + "Eliminaré aquesta entrada.", vbCritical, "Error": datacontroldescansirelleu.Recordset.Delete
End Sub
Function calcularhores(rst As Recordset) As Double
   With rst
   calcularhores = DateDiff("n", CVDate(atrim(!datainici) + " " + atrim(!horainici)), CVDate(atrim(!datafi) + " " + atrim(!horafi)))
   calcularhores = Redondejar(calcularhores / 60, 2)
   End With
End Function

Private Sub ctotes_Click()
   If ctotes.Value = 1 Then
      carregardades "select * from controldescansrelleu where nummaq=" + atrim(nummaq) + " and operari=" + etnomoperari.Tag + " and seccio='" + atrim(lletraseccio) + "' order by datainici,horainici"
     Else: carregardades
   End If
End Sub

Private Sub Form_Activate()
   carregardades
End Sub
Function hihaalgunsenseacabar() As Boolean
   datacontroldescansirelleu.Recordset.FindFirst "hores=0 or hores=null"
   If Not datacontroldescansirelleu.Recordset.NoMatch Then hihaalgunsenseacabar = True
End Function
Function escullintmenu(v As String) As Boolean
  If mnou.Tag = "escullint" Then
     escullintmenu = True: mnou.Tag = ""
     datacontroldescansirelleu.Recordset.Edit
     datacontroldescansirelleu.Recordset!tipus = v
     datacontroldescansirelleu.Recordset.Update
  End If
End Function

Private Sub mnoudescans_Click()
  If escullintmenu("Descans") Then Exit Sub
  If hihaalgunsenseacabar Then
     MsgBox "Hi ha una entrada sense finalitzar", vbCritical, "Error"
     Exit Sub
  End If
  datacontroldescansirelleu.Recordset.AddNew
  datacontroldescansirelleu.Recordset!nummaq = nummaq
  datacontroldescansirelleu.Recordset!operari = etnomoperari.Tag
  datacontroldescansirelleu.Recordset!seccio = lletraseccio
  datacontroldescansirelleu.Recordset!tipus = "Descans"
  datacontroldescansirelleu.Recordset!datainici = Date
  datacontroldescansirelleu.Recordset!horainici = Time
  datacontroldescansirelleu.Recordset!comanda = ncomanda
  datacontroldescansirelleu.Recordset.Update
  carregardades
End Sub

Private Sub mnourelleu_Click()
 If escullintmenu("Relleu") Then Exit Sub
 If hihaalgunsenseacabar Then
     MsgBox "Hi ha una entrada sense finalitzar", vbCritical, "Error"
     Exit Sub
  End If
  datacontroldescansirelleu.Recordset.AddNew
  datacontroldescansirelleu.Recordset!nummaq = nummaq
  datacontroldescansirelleu.Recordset!operari = etnomoperari.Tag
  datacontroldescansirelleu.Recordset!seccio = lletraseccio
  datacontroldescansirelleu.Recordset!tipus = "Relleu"
  datacontroldescansirelleu.Recordset!datainici = Date
  datacontroldescansirelleu.Recordset!horainici = Time
  datacontroldescansirelleu.Recordset!comanda = ncomanda
  datacontroldescansirelleu.Recordset.Update
  carregardades
End Sub
Sub canviartipus()
   mnou.Tag = "escullint"
   Me.PopupMenu mnou, , reixa.Left + reixa.Columns(reixa.col).Left, reixa.Top + reixa.Columns(reixa.col).Top
End Sub
Private Sub reixa_DblClick()
  Dim v As String
  Dim id As Long
  If datacontroldescansirelleu.Recordset.EOF Then Exit Sub
  If reixa.Columns(reixa.col).DataField = "hores" Then Exit Sub
  If reixa.Columns(reixa.col).DataField = "tipus" Then canviartipus: Exit Sub
  v = InputBox("Entra el valor de " + UCase(reixa.Columns(reixa.col).DataField), "Modificació", reixa.Text)
  If v <> "" Then
    On Error Resume Next
    id = datacontroldescansirelleu.Recordset!id
    datacontroldescansirelleu.Recordset.Edit
    datacontroldescansirelleu.Recordset.Fields(reixa.Columns(reixa.col).DataField) = v
    datacontroldescansirelleu.Recordset!hores = calcularhores(datacontroldescansirelleu.Recordset)
    datacontroldescansirelleu.Recordset.Update
    datacontroldescansirelleu.Recordset.FindFirst "id=" + atrim(id)
    If datacontroldescansirelleu.Recordset!horafi <> "" Then
       If datacontroldescansirelleu.Recordset!hores = 0 Then
         MsgBox "La diferencia entre inici i fi es massa curta" + Chr(10) + "Eliminaré aquesta entrada.", vbCritical, "Error"
         datacontroldescansirelleu.Recordset.Delete
         datacontroldescansirelleu.Refresh
       End If
    End If
    On Error GoTo 0
  End If
End Sub
