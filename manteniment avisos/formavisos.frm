VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formavisos 
   Caption         =   "Manteniment d'Avisos"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10560
   Icon            =   "formavisos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   630
      Left            =   45
      TabIndex        =   15
      Top             =   -45
      Width           =   10455
      Begin VB.CommandButton sortir 
         Height          =   465
         Left            =   9840
         Picture         =   "formavisos.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Sortir a Menú"
         Top             =   120
         Width           =   540
      End
      Begin VB.CommandButton alta 
         Height          =   450
         Left            =   60
         Picture         =   "formavisos.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   120
         Width           =   450
      End
      Begin VB.CommandButton eliminar 
         Height          =   450
         Left            =   1065
         Picture         =   "formavisos.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   120
         Width           =   465
      End
      Begin VB.CommandButton modificar 
         Height          =   450
         Left            =   510
         Picture         =   "formavisos.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Modificar Registres"
         Top             =   120
         Width           =   540
      End
      Begin VB.CommandButton Command11 
         Height          =   450
         Left            =   1545
         Picture         =   "formavisos.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Acceptar els canvis (F1)."
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.Data dataavisos 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\avisosseccions.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3510
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "avisos"
      Top             =   2925
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Llista d'avisos"
      Height          =   4095
      Left            =   60
      TabIndex        =   1
      Top             =   3330
      Width           =   10425
      Begin MSDBGrid.DBGrid reixa 
         Bindings        =   "formavisos.frx":213C
         Height          =   3645
         Left            =   135
         OleObjectBlob   =   "formavisos.frx":2151
         TabIndex        =   12
         Top             =   270
         Width           =   10110
      End
   End
   Begin VB.Frame fdades 
      BackColor       =   &H00FAF1F1&
      Enabled         =   0   'False
      Height          =   2655
      Left            =   45
      TabIndex        =   0
      Top             =   510
      Width           =   10455
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FAF1F1&
         Caption         =   "Inactiu"
         DataField       =   "inactiu"
         DataSource      =   "dataavisos"
         Height          =   195
         Left            =   9435
         TabIndex        =   13
         Top             =   195
         Width           =   810
      End
      Begin VB.TextBox cmissatge 
         DataField       =   "missatge"
         DataSource      =   "dataavisos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   195
         TabIndex        =   10
         Top             =   1815
         Width           =   9150
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAD9CE&
         Height          =   1155
         Left            =   180
         TabIndex        =   2
         Top             =   375
         Width           =   2610
         Begin VB.ComboBox comboseccio 
            DataField       =   "seccio"
            DataSource      =   "dataavisos"
            Height          =   315
            ItemData        =   "formavisos.frx":3036
            Left            =   300
            List            =   "formavisos.frx":3046
            TabIndex        =   3
            Top             =   450
            Width           =   2175
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Secció"
            Height          =   240
            Left            =   915
            TabIndex        =   4
            Top             =   180
            Width           =   1620
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAD9CE&
         Height          =   1125
         Left            =   3390
         TabIndex        =   5
         Top             =   390
         Width           =   6810
         Begin VB.ComboBox combovalor 
            DataField       =   "valor"
            DataSource      =   "dataavisos"
            Height          =   315
            Left            =   3510
            TabIndex        =   7
            Top             =   480
            Width           =   3105
         End
         Begin VB.ComboBox combocondicionant 
            DataField       =   "condicionant"
            DataSource      =   "dataavisos"
            Height          =   315
            ItemData        =   "formavisos.frx":307E
            Left            =   345
            List            =   "formavisos.frx":308E
            TabIndex        =   6
            Top             =   465
            Width           =   2175
         End
         Begin VB.Label vvalor 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF8080&
            Height          =   180
            Left            =   1290
            TabIndex        =   14
            Top             =   840
            Width           =   5265
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   2865
            Picture         =   "formavisos.frx":30B7
            Top             =   495
            Width           =   240
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
            Height          =   240
            Left            =   4020
            TabIndex        =   9
            Top             =   195
            Width           =   1620
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Condicionant"
            Height          =   240
            Left            =   750
            TabIndex        =   8
            Top             =   195
            Width           =   1620
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Missatge"
         Height          =   240
         Left            =   585
         TabIndex        =   11
         Top             =   1575
         Width           =   1710
      End
   End
End
Attribute VB_Name = "formavisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
  nou_avis
End Sub
Sub nou_avis()
If dataavisos.Recordset.EditMode > 0 Then MsgBox "Estas editant un registre, primer finalitza els canvis.", vbCritical, "Atenció": Exit Sub
  fdades.Enabled = True
  dataavisos.Recordset.AddNew
  comboseccio.SetFocus
  SendKeys "%{DOWN}"
End Sub
Sub possarvvalor()
   Dim rst As Recordset
   vvalor = ""
   If combocondicionant = "Client" Then
       Set rst = dbcomandes.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(combovalor)))
       If rst.EOF Then vvalor = "": Exit Sub
       vvalor = atrim(rst!nom)
   End If
   Set rst = Nothing
End Sub

Private Sub combocondicionant_Click()
possarvvalor
End Sub

Private Sub combocondicionant_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
End Sub

Private Sub combocondicionant_KeyPress(KeyAscii As Integer)
  KeyPress = 0
End Sub

Private Sub combovalor_DropDown()
   If combocondicionant = "Client" Then escullir_client
   If combocondicionant = "Marca" Then escullir_marca
End Sub
Sub escullir_marca()
 Load formseleccio
   formseleccio.Command3.Tag = "filtre"
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "select Marca from marques"
   formseleccio.refrescar
   formseleccio.Show 1
   
    If seleccioret = 1 Then
            combovalor = formseleccio.DBGrid2.Columns("marca")
    End If
     If seleccioret = 9 Then
         combovalor = ""
         
    End If
    formseleccio.Data1.RecordSource = ""
    formseleccio.Data1.Refresh
    Unload formseleccio
    SendKeys "{TAB}"
End Sub
Sub escullir_client()
   Load formseleccio
   formseleccio.Command3.Tag = "filtre"
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select * from clients"
   formseleccio.refrescar
   formseleccio.Show 1
   
    If seleccioret = 1 Then
            vvalor = formseleccio.DBGrid2.Columns("nom")
            combovalor = formseleccio.DBGrid2.Columns("codi")
    End If
     If seleccioret = 9 Then
         combovalor = "0"
         
    End If
    formseleccio.Data1.RecordSource = ""
    formseleccio.Data1.Refresh
    Unload formseleccio
    SendKeys "{TAB}"
    
 
End Sub

Private Sub Command11_Click()
  gravar_canvis
  
End Sub
Sub gravar_canvis()
  If dataavisos.Recordset.EditMode = 0 Then MsgBox "No pots gravar sense estar editant.", vbCritical, "Atenció": Exit Sub
  If Not comprovarcamps Then MsgBox "Falta emplenar algun camp.", vbCritical, "Atenció": Exit Sub
  fdades.Enabled = False
  dataavisos.Recordset.Update
End Sub
Function comprovarcamps() As Boolean
   comprovarcamps = True
   If comboseccio = "" Or combocondicionant = "" Or combovalor = "" Or cmissatge = "" Then comprovarcamps = False
End Function

Private Sub dataavisos_Reposition()
   If Not dataavisos.Recordset.EOF Then possarvvalor
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 112 Then gravar_canvis
End Sub

Private Sub Form_Load()
  fitxerini = "comandes.ini"
  cami = llegir_ini("General", "cami", fitxerini)
  
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
  '"c:\misdoc~1\commandes\comandes.mdb"
  If existeix("c:\ordprog.ini") Then cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
  centerscreen Me
  camiclixes = rutadelfitxer(cami) + "clixesnous.mdb"
  Set dbclixes = DBEngine.OpenDatabase(camiclixes)
  Set dbcomandes = DBEngine.OpenDatabase(cami)
  dataavisos.DatabaseName = rutadelfitxer(cami) + "avisosseccions.mdb"
  dataavisos.RecordSource = "select * from avisos order by id desc"
  dataavisos.Refresh
End Sub

Private Sub modificar_Click()
  If dataavisos.Recordset.EditMode > 0 Then MsgBox "Estas editant un registre, primer finalitza els canvis.", vbCritical, "Atenció": Exit Sub
  fdades.Enabled = True
  dataavisos.Recordset.Edit
  comboseccio.SetFocus
End Sub

Private Sub sortir_Click()
   End
End Sub
