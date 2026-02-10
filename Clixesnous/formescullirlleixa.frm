VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formescullirlleixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Escullir Lleixa"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   3195
   ControlBox      =   0   'False
   Icon            =   "formescullirlleixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton bnotriarcap 
      Height          =   495
      Left            =   105
      Picture         =   "formescullirlleixa.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "No escullir-ne cap"
      Top             =   6495
      Width           =   810
   End
   Begin VB.Frame Frame3 
      Height          =   945
      Left            =   45
      TabIndex        =   7
      Top             =   5490
      Width           =   3090
      Begin VB.TextBox cxl 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1095
         MaxLength       =   5
         TabIndex        =   11
         Top             =   225
         Width           =   1755
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "XL-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   690
         Left            =   135
         TabIndex        =   8
         Top             =   150
         Width           =   975
      End
   End
   Begin VB.CommandButton sortir 
      Height          =   495
      Left            =   2295
      Picture         =   "formescullirlleixa.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Sortir"
      Top             =   6495
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   1455
      Picture         =   "formescullirlleixa.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Escullir"
      Top             =   6495
      Width           =   810
   End
   Begin VB.Frame Frame2 
      Height          =   945
      Left            =   45
      TabIndex        =   2
      Top             =   4515
      Width           =   3090
      Begin VB.TextBox ctaula 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   750
         MaxLength       =   5
         TabIndex        =   10
         Top             =   225
         Width           =   1755
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "T-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   690
         Left            =   135
         TabIndex        =   6
         Top             =   150
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   45
      TabIndex        =   1
      Top             =   3600
      Width           =   3090
      Begin VB.TextBox cpalet 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   765
         MaxLength       =   5
         TabIndex        =   9
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "P-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   690
         Left            =   135
         TabIndex        =   5
         Top             =   150
         Width           =   675
      End
   End
   Begin VB.Data dataubicacions 
      Caption         =   "dataubicacions"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   435
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select descripcio from tipificacions_xl"
      Top             =   2580
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "formescullirlleixa.frx":1628
      Height          =   3600
      Left            =   0
      OleObjectBlob   =   "formescullirlleixa.frx":1641
      TabIndex        =   0
      Top             =   0
      Width           =   3195
   End
   Begin VB.Menu mtipificacions 
      Caption         =   "Manteniment Tipificacions"
   End
End
Attribute VB_Name = "formescullirlleixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bnotriarcap_Click()
bnotriarcap.tag = "1"
  Me.Hide
End Sub

Private Sub Command1_Click()
  seleccioret = 1
  Me.Hide
End Sub
Function valorescullit() As String
  If bnotriarcap.tag = "1" Then valorescullit = "": Exit Function
  If cadbl(cpalet) > 0 Then valorescullit = "P-" + cpalet
  If atrim(ctaula) <> "" Then valorescullit = "T-" + ctaula
  If cadbl(cxl) > 0 Then valorescullit = "XL-" + cxl
  If valorescullit <> "" Then Exit Function
  If Not dataubicacions.Recordset.EOF Then
      valorescullit = atrim(dataubicacions.Recordset!descripcio)
  End If
End Function

Private Sub Command2_Click()
  
End Sub

Private Sub cpalet_Change()
  If Screen.ActiveControl.Name = "cpalet" Then
      cxl = ""
      ctaula = ""
  End If
End Sub

Private Sub ctaula_Change()
  If Screen.ActiveControl.Name = "ctaula" Then
      cxl = ""
      cpalet = ""
  End If
End Sub

Private Sub cxl_Change()
   If Screen.ActiveControl.Name = "cxl" Then
      cpalet = ""
      ctaula = ""
  End If
End Sub

Private Sub Form_Activate()
  If cxl.tag <> "1" Then cxl = ""
  ctaula = ""
  cpalet = ""
  cxl.SetFocus
End Sub

Private Sub Form_Load()
   dataubicacions.DatabaseName = rutadelfitxer(cami) + "clixesnous.mdb"
   bnotriarcap.tag = ""
End Sub

Private Sub mtipificacions_Click()
 Load formaltarep
  formaltarep.caption = "Manteniment Tipificacions"
'  formaltarep.autonum = "transportistes"
  formaltarep.Data1.DatabaseName = dataubicacions.DatabaseName
  formaltarep.Data1.RecordSource = "select * from tipificacions_xl"

  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).visible = False
  formaltarep.DBGrid1.Columns(1).width = formaltarep.DBGrid1.Columns(1).width * 2
  formaltarep.width = formaltarep.width - 1800
  'formaltarep.DBGrid1.width = formaltarep.DBGrid1.width + 700
  formaltarep.Show 1
End Sub

Private Sub sortir_Click()
  seleccioret = 9
  Me.Hide
End Sub
