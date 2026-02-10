VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form mantenimentprestatges 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniment de Prestatges"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   Icon            =   "mantenimentprestatges.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data prestatges 
      Caption         =   "prestatges"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3405
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from prestatges order by mid(numlleixa,1,1),cdbl(mid(numlleixa,2))"
      Top             =   780
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Caption         =   "Prestatges"
      Height          =   7785
      Left            =   90
      TabIndex        =   1
      Top             =   960
      Width           =   4140
      Begin MSDBGrid.DBGrid reixa 
         Bindings        =   "mantenimentprestatges.frx":058A
         Height          =   7365
         Left            =   90
         OleObjectBlob   =   "mantenimentprestatges.frx":059F
         TabIndex        =   2
         Top             =   270
         Width           =   3930
      End
   End
   Begin VB.Frame Frame1 
      Height          =   765
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   4185
      Begin VB.CommandButton Command3 
         Height          =   345
         Left            =   1665
         Picture         =   "mantenimentprestatges.frx":1148
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   315
         Width           =   360
      End
      Begin VB.TextBox cforatsxrlleixa 
         Height          =   300
         Left            =   2100
         TabIndex        =   9
         Top             =   285
         Width           =   570
      End
      Begin VB.CommandButton Command2 
         Height          =   360
         Left            =   1290
         Picture         =   "mantenimentprestatges.frx":16D2
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   360
      End
      Begin VB.CommandButton sortir 
         Height          =   495
         Left            =   3510
         Picture         =   "mantenimentprestatges.frx":1C5C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Sortir"
         Top             =   195
         Width           =   540
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Afegir"
         Height          =   480
         Left            =   2685
         Picture         =   "mantenimentprestatges.frx":21E6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   810
      End
      Begin VB.CheckBox alt 
         Caption         =   "Alt?"
         Height          =   195
         Left            =   1260
         TabIndex        =   5
         Top             =   120
         Width           =   600
      End
      Begin VB.TextBox lleixa 
         Height          =   300
         Left            =   75
         TabIndex        =   3
         Top             =   345
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Forats xr lleixa"
         Height          =   240
         Left            =   1920
         TabIndex        =   10
         Top             =   60
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lleixa"
         Height          =   240
         Left            =   465
         TabIndex        =   4
         Top             =   135
         Width           =   1035
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ex: B199  (Prestatge B Fila 1 Nº 99)"
      Height          =   285
      Left            =   810
      TabIndex        =   11
      Top             =   765
      Width           =   2220
   End
End
Attribute VB_Name = "mantenimentprestatges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If cadbl(cforatsxrlleixa) = 0 Then MsgBox "Has de possar quants forats per lleixa hi ha.", vbCritical, "Error": Exit Sub
  If lleixa <> "" Then
    lleixa = UCase(lleixa)
    prestatges.Recordset.FindFirst "numlleixa='" + atrim(lleixa) + "'"
    If prestatges.Recordset.NoMatch Then
     prestatges.Recordset.AddNew
     prestatges.Recordset!numlleixa = lleixa
     prestatges.Recordset!prestatgealt = cabool(alt.Value)
     prestatges.Recordset!foratsperlleixa = cadbl(cforatsxrlleixa)
     prestatges.Recordset.Update
     prestatges.Refresh
       Else: MsgBox "Aquesta lleixa ja existeix.": Exit Sub
    End If
  End If
End Sub

Private Sub command2_click()
  prestatges.Recordset.FindFirst "numlleixa='" + atrim(lleixa) + "'"
End Sub

Private Sub Command3_Click()
 Dim llistat As CrystalReport
 Dim vlleixa As String
 Set llistat = Form1.llistat
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetaestanteria.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = ""
 vlleixa = UCase(lleixa) + "          "
 llistat.DiscardSavedData = True
 llistat.Formulas(1) = "estanteria='" + Mid(vlleixa, 1, 1) + "'"
 llistat.Formulas(0) = "fila='" + Mid(vlleixa, 4, 1) + "'"
 llistat.Formulas(2) = "columna='" + Trim(Mid(vlleixa, 2, 2)) + "'"
 llistat.Formulas(3) = ""
 llistat.Formulas(4) = ""
 llistat.Formulas(5) = ""
 llistat.Formulas(6) = ""
 llistat.Formulas(7) = ""
 llistat.Formulas(8) = ""
 llistat.Formulas(9) = ""
 llistat.Formulas(10) = ""
 llistat.Formulas(11) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If Form1.mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
End Sub

Private Sub Form_Load()
  prestatges.DatabaseName = Form1.palets.DatabaseName
  prestatges.Refresh
End Sub

Private Sub lleixa_LostFocus()
  lleixa = UCase(lleixa)
End Sub

Private Sub reixa_BeforeDelete(Cancel As Integer)
   If MsgBox("Segur que vols eliminar aquesta lleixa?", vbCritical + vbDefaultButton2 + vbYesNo, "Borrar lleixa") = vbNo Then Cancel = 1
End Sub

Private Sub reixa_DblClick()
   Dim v As String
   If prestatges.Recordset.EditMode = 0 Then prestatges.Recordset.Edit
   If reixa.Columns(reixa.col).DataField = "prestatgealt" Then
      prestatges.Recordset!prestatgealt = IIf(MsgBox("Vols que aquest prestage sigui alt?", vbInformation + vbYesNo + vbDefaultButton2, "Prestatge Alt?") = vbYes, True, False)
   End If
   If reixa.Columns(reixa.col).DataField = "foratsperlleixa" Then
      v = InputBox("Entra quants forats de aquesta lleixa", "Forats", 3)
      If StrPtr(v) = Empty Then Exit Sub
      prestatges.Recordset!foratsperlleixa = cadbl(v)
   End If
   prestatges.Recordset.Update
End Sub

Private Sub sortir_Click()
    Unload mantenimentprestatges
End Sub
