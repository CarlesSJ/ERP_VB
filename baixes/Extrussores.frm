VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Extrussores 
   Caption         =   "Baixes Extrussores"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   Icon            =   "Extrussores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   30
      Top             =   3150
   End
   Begin VB.CommandButton eliminar 
      Height          =   300
      Left            =   195
      Picture         =   "Extrussores.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   3435
      Width           =   300
   End
   Begin VB.Data bobines 
      Caption         =   "bobines"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "extrussores"
      Top             =   2910
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Data data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "y:\comandes\comandes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6900
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "comandes"
      Top             =   -135
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Frame Frame3 
      Caption         =   "Desguas de Feina"
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   3180
      Width           =   9210
      Begin VB.CommandButton detall 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Detall"
         Height          =   240
         Left            =   8355
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   495
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.ListBox List1 
         Height          =   645
         ItemData        =   "Extrussores.frx":0754
         Left            =   3195
         List            =   "Extrussores.frx":0761
         TabIndex        =   34
         Top             =   705
         Visible         =   0   'False
         Width           =   1275
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Extrussores.frx":0782
         Height          =   3450
         Left            =   75
         OleObjectBlob   =   "Extrussores.frx":0794
         TabIndex        =   33
         Top             =   330
         Width           =   9090
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   9645
         Picture         =   "Extrussores.frx":208B
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1740
         Width           =   315
      End
      Begin MSMask.MaskEdBox Text31 
         DataField       =   "mesuracantex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   9630
         TabIndex        =   6
         Top             =   540
         Visible         =   0   'False
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text23 
         DataField       =   "mesuraesp"
         DataSource      =   "data1"
         Height          =   285
         Left            =   9270
         TabIndex        =   7
         Top             =   540
         Visible         =   0   'False
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "oberturaex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   9585
         TabIndex        =   8
         Top             =   1125
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   503
         _Version        =   327681
         MaxLength       =   1
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         DataField       =   "micropex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   9585
         TabIndex        =   9
         Top             =   840
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   503
         _Version        =   327681
         MaxLength       =   1
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Totals"
      Height          =   765
      Left            =   150
      TabIndex        =   1
      Top             =   2220
      Width           =   9225
      Begin VB.TextBox kiloshora 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   8100
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   405
         Width           =   840
      End
      Begin VB.TextBox tmetres 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   7020
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   405
         Width           =   840
      End
      Begin VB.TextBox tkilos 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6015
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   405
         Width           =   840
      End
      Begin VB.TextBox tbobines 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4830
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   405
         Width           =   840
      End
      Begin VB.TextBox hfunc 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   405
         Width           =   840
      End
      Begin VB.TextBox havaria 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox hcanvi 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   390
         Width           =   840
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos/ Hora"
         Height          =   210
         Left            =   8100
         TabIndex        =   55
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Metres"
         Height          =   210
         Left            =   7020
         TabIndex        =   48
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Kilos"
         Height          =   210
         Left            =   6030
         TabIndex        =   46
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bob."
         Height          =   195
         Left            =   4830
         TabIndex        =   44
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Hores Func."
         Height          =   210
         Left            =   3690
         TabIndex        =   42
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Hores Avaria"
         Height          =   210
         Left            =   2730
         TabIndex        =   40
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Hores Canvi"
         Height          =   210
         Left            =   1695
         TabIndex        =   38
         Top             =   165
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Capçalera"
      Height          =   1965
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   9255
      Begin VB.TextBox ruta 
         Height          =   285
         Left            =   4170
         TabIndex        =   56
         Top             =   300
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "simulteneitat"
         DataSource      =   "data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Extrussores.frx":2451
         Left            =   8460
         List            =   "Extrussores.frx":2464
         TabIndex        =   27
         Top             =   945
         Width           =   720
      End
      Begin MSMask.MaskEdBox Text24 
         DataField       =   "colorex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   1395
         TabIndex        =   10
         Top             =   930
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text25 
         DataField       =   "materialex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   1395
         TabIndex        =   13
         Top             =   1245
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text26 
         DataField       =   "aditiuex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   1395
         TabIndex        =   16
         Top             =   1545
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text18 
         DataField       =   "ampleesq"
         DataSource      =   "data1"
         Height          =   285
         Left            =   7035
         TabIndex        =   19
         Top             =   615
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text19 
         DataField       =   "plegatesq"
         DataSource      =   "data1"
         Height          =   285
         Left            =   7995
         TabIndex        =   21
         Top             =   630
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text20 
         DataField       =   "solapa"
         DataSource      =   "data1"
         Height          =   285
         Left            =   7020
         TabIndex        =   23
         Top             =   930
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text21 
         DataField       =   "espessor"
         DataSource      =   "data1"
         Height          =   285
         Left            =   7020
         TabIndex        =   25
         Top             =   1245
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text30 
         DataSource      =   "data1"
         Height          =   285
         Left            =   8460
         TabIndex        =   29
         ToolTipText     =   "Mesura de la Quantitat"
         Top             =   1275
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         DataField       =   "comanda"
         DataSource      =   "data1"
         Height          =   285
         Left            =   1380
         TabIndex        =   31
         Top             =   630
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         DataField       =   "datacomanda"
         DataSource      =   "data1"
         Height          =   285
         Left            =   4425
         TabIndex        =   32
         Top             =   630
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text22 
         DataField       =   "micropex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   8685
         TabIndex        =   57
         ToolTipText     =   "Mesura de l'Espessor"
         Top             =   1605
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         DataField       =   "cantitatex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   7005
         TabIndex        =   58
         Top             =   1605
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "MicroP:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   26
         Left            =   7965
         TabIndex        =   60
         Top             =   1635
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Quantitat:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   3
         Left            =   6225
         TabIndex        =   59
         Top             =   1650
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Client"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   1245
         TabIndex        =   52
         Top             =   120
         Width           =   765
      End
      Begin VB.Label nomclient 
         Caption         =   "Nom del client"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   225
         MouseIcon       =   "Extrussores.frx":2477
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   330
         Width           =   5640
      End
      Begin VB.Label Label1 
         Caption         =   "Producte:"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   1
         Left            =   6630
         TabIndex        =   50
         Top             =   120
         Width           =   765
      End
      Begin VB.Label nomproducte 
         Caption         =   "Descripcio del producte"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   240
         Left            =   5985
         TabIndex        =   49
         Top             =   345
         Width           =   3450
      End
      Begin VB.Label Label1 
         Caption         =   "Obert?"
         DataSource      =   "data1"
         Height          =   255
         Index           =   19
         Left            =   7905
         TabIndex        =   30
         Top             =   1305
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Simult."
         DataSource      =   "data1"
         Height          =   255
         Index           =   33
         Left            =   7905
         TabIndex        =   28
         Top             =   1005
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Espessor:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   18
         Left            =   6255
         TabIndex        =   26
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Solapa:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   16
         Left            =   6225
         TabIndex        =   24
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "/"
         DataSource      =   "data1"
         Height          =   255
         Index           =   14
         Left            =   7845
         TabIndex        =   22
         Top             =   690
         Width           =   165
      End
      Begin VB.Label Label1 
         Caption         =   "Ample/Pleg:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   13
         Left            =   6120
         TabIndex        =   20
         Top             =   675
         Width           =   915
      End
      Begin VB.Label nomadditiu 
         BackStyle       =   0  'Transparent
         Caption         =   "Additiu:"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   23
         Left            =   2070
         TabIndex        =   18
         Top             =   1620
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Additiu:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   22
         Left            =   360
         TabIndex        =   17
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label nommaterial 
         BackStyle       =   0  'Transparent
         Caption         =   "Material:"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   23
         Left            =   2070
         TabIndex        =   15
         Top             =   1320
         Width           =   4065
      End
      Begin VB.Label Label1 
         Caption         =   "Material:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   21
         Left            =   345
         TabIndex        =   14
         Top             =   1305
         Width           =   915
      End
      Begin VB.Label nomcolor 
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   23
         Left            =   2055
         TabIndex        =   12
         Top             =   1005
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Color:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   20
         Left            =   330
         TabIndex        =   11
         Top             =   1005
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Data Comanda:"
         Height          =   165
         Left            =   3210
         TabIndex        =   4
         Top             =   675
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Comanda:"
         Height          =   165
         Index           =   0
         Left            =   225
         TabIndex        =   3
         Top             =   660
         Width           =   1095
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Prem F2 per sel.leccionar Taules..."
      Height          =   225
      Left            =   150
      TabIndex        =   53
      Top             =   7035
      Width           =   9120
   End
End
Attribute VB_Name = "Extrussores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub comodi_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 39 Then KeyCode = 0: DBGrid1.SetFocus: SendKeys "{RIGHT}"
  If KeyCode = 37 Then KeyCode = 0: DBGrid1.SetFocus: SendKeys "{LEFT}"
End Sub

Private Sub Data1_Reposition()
  carregar_lookups
'ensenyar_totalstotals
End Sub
Sub carregar_lookups()
lookupde "colorants", Text24, nomcolor(23)
lookupde "materials", Text25, nommaterial(23)
lookupde "aditius", Text26, nomadditiu(23)
 'LOOKUP DE producte
  Set rsttmp = dbtmp.OpenRecordset("select descripcio,ruta from productes where codi='" + atrim((data1.Recordset!producte)) + "'")
  If Not rsttmp.EOF Then
     nomproducte.Caption = atrim(data1.Recordset!producte) + " - " + atrim(rsttmp!descripcio)
     ruta.Text = atrim(rsttmp!ruta)
    Else: nomproducte.Caption = "": ruta = ""
  End If
   'LOOKUP DE client
  Set rsttmp = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(data1.Recordset!client)))
  If Not rsttmp.EOF Then
     nomclient.Caption = atrim(data1.Recordset!client) + " - " + atrim(rsttmp!nom)
    Else: nomclient.Caption = ""
  End If
End Sub

Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

 If bobines.Recordset.EditMode = 0 And Not bobines.Recordset.EOF Then
  bobines.Recordset.Edit
 End If
 On Error Resume Next
 bobines.Recordset!comanda = data1.Recordset!comanda
 On Error GoTo 0
 DBGrid1_RowColChange DBGrid1.Row, DBGrid1.col
End Sub

Private Sub DBGrid1_BeforeUpdate(Cancel As Integer)
'bobines.Recordset!datahorainici = DBGrid1.Columns(1).Text + " " + DBGrid1.Columns(2).Text
End Sub

Private Sub DBGrid1_ButtonClick(ByVal ColIndex As Integer)
 If ColIndex = 5 Then
   List1.Visible = True
   List1.Width = DBGrid1.Columns(ColIndex).Width
   List1.Top = DBGrid1.RowTop(DBGrid1.Row) + DBGrid1.Top + DBGrid1.RowHeight
   List1.Left = DBGrid1.Columns(ColIndex).Left + DBGrid1.Left
   List1.SetFocus
 End If

End Sub

Private Sub DBGrid1_Change()
'  Me.Caption = DBGrid1.Text
End Sub

Private Sub DBGrid1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
colocardetall
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    KeyCode = 0
    SendKeys "{TAB}"
  End If
 If KeyCode = 113 And DBGrid1.col = 0 Then
   triarextrussora
 End If
  If (KeyCode = Asc("D") Or KeyCode = Asc("d")) And Shift = 2 Then
    detall_Click
  End If
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
  If DBGrid1.col = 5 Then
    If InStr(1, "CAF", UCase$(Chr$(KeyAscii))) = 0 Then
       KeyAscii = 0
      Else: KeyAscii = Asc(UCase$(Chr$(KeyAscii))): DBGrid1.Text = ""
    End If
  End If
  
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Dim datatmp As String
  Dim col As Integer
   Dim valtmp As String
   
     'comprova si hem escrit el numero amb separat per .
  If LastCol >= 0 Then
   If IsNumeric(DBGrid1.Columns(LastCol)) Then
      If InStr(1, DBGrid1.Columns(LastCol), ".") Then
         DBGrid1.Columns(LastCol) = Mid(DBGrid1.Columns(LastCol), 1, InStr(1, DBGrid1.Columns(LastCol), ".") - 1) + "," + Mid(DBGrid1.Columns(LastCol), InStr(1, DBGrid1.Columns(LastCol), ".") + 1)
      End If
   End If
  End If
  
  'coloca el boto de detall al final de la reixa
  colocardetall
   'COLOCO LES DATES PER DEFECTE
  If DBGrid1.col = 1 Then
    If atrim(DBGrid1.Text) = "" Then DBGrid1.Text = Format(DateAdd("d", -1, Now), "dd/mm/yy")
  End If
  
  If DBGrid1.col = 3 Then
    If atrim(DBGrid1.Text) = "" Then DBGrid1.Text = Format(DBGrid1.Columns(1).Text, "dd/mm/yy")
  End If
  
  
  '-------
  
  If LastCol = 1 Or LastCol = 2 Then
  valtmp = DBGrid1.Columns(LastCol).Text
  If LastCol = 1 Then
      
      If InStr(1, valtmp, "/") = 0 Then valtmp = Mid(valtmp, 1, 2) + "/" + Mid(valtmp, 3, 2) + "/" + Mid(valtmp, 5, 2)
      If Not IsDate(valtmp) Then valtmp = ""
  End If
  If LastCol = 2 Then
    If InStr(1, valtmp, ":") = 0 Then valtmp = Mid(valtmp, 1, 2) + ":" + Mid(valtmp, 3, 2)
      If Not IsDate(Format(valtmp, "hh:nn")) Then valtmp = "00:00"

  End If
  DBGrid1.Columns(LastCol) = valtmp
  End If
  
  If LastCol = 3 Or LastCol = 4 Then
  valtmp = DBGrid1.Columns(LastCol).Text
  If LastCol = 3 Then
      
      If InStr(1, valtmp, "/") = 0 Then valtmp = Mid(valtmp, 1, 2) + "/" + Mid(valtmp, 3, 2) + "/" + Mid(valtmp, 5, 2)
      If Not IsDate(valtmp) Then valtmp = ""
  End If
  If LastCol = 4 Then
    If InStr(1, valtmp, ":") = 0 Then valtmp = Mid(valtmp, 1, 2) + ":" + Mid(valtmp, 3, 2)
      If Not IsDate(Format(valtmp, "hh:nn")) Then valtmp = "00:00"

  End If
  DBGrid1.Columns(LastCol) = valtmp
  End If
  'comprovo si la extrussora entrada es correcte
  If LastCol = 0 Then
   If cadbl(DBGrid1.Columns(0)) <> 0 Then
     Set rsttmp = dbtmp.OpenRecordset("select codi from maquines where maquina='E' and codi=" + atrim(cadbl(DBGrid1.Columns(0))))
     If rsttmp.EOF Then MsgBox "Aquesta Extrussora no Existeix": DBGrid1.Columns(0) = "": DBGrid1.col = 0
   End If
  End If
  
  calcular_totals
End Sub

Sub colocardetall()
 If Not bobines.Recordset.EOF Then
  If DBGrid1.Columns(10).Left > 0 Then
   If bobines.Recordset!tipus = "F" Then
     detall.Visible = True
     detall.Width = DBGrid1.Columns(10).Width
     detall.Top = DBGrid1.RowTop(DBGrid1.Row) + DBGrid1.Top
     detall.Left = DBGrid1.Columns(10).Left + DBGrid1.Left
    Else: detall.Visible = False
   End If
    Else: detall.Visible = False
  End If
 End If
End Sub

Private Sub DBGrid1_RowResize(Cancel As Integer)
colocardetall
End Sub

Private Sub detall_Click()
'  MsgBox "obrir un formulari de detall de bobines"
  On Error Resume Next
  Unload detallbobext
  On Error GoTo 0
  detallbobext.Show 1
  calcular_totals
  DBGrid1.Row = 0
  DBGrid1.SetFocus
  End Sub
Sub calcular_totals()
  Dim total As Double
  Dim hores As Double
  If bobines.Recordset.EOF Then Exit Sub
  If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
  Set rsttmp = dbtmpb.OpenRecordset("select count(*) as elgran from bobinesext where controlid=" + atrim(bobines.Recordset!ID))
  If Not rsttmp.EOF Then bobines.Recordset!totalbobines = rsttmp!elgran
  
  Set rsttmp = dbtmpb.OpenRecordset("select sum(kilos) as elgran from bobinesext where controlid=" + atrim(bobines.Recordset!ID))
  If Not rsttmp.EOF Then bobines.Recordset!totalkilos = rsttmp!elgran
  
  Set rsttmp = dbtmpb.OpenRecordset("select sum(metres) as elgran from bobinesext where controlid=" + atrim(bobines.Recordset!ID))
  If Not rsttmp.EOF Then bobines.Recordset!totalmetres = rsttmp!elgran
  
  With bobines.Recordset
  total = 0
  On Error Resume Next
  total = DateDiff("n", CVDate(atrim(!datainici) + " " + atrim(!horainici)), CVDate(atrim(!datafi) + " " + atrim(!horafi)))
  total = Format(total / 60, "#,##0.00")
  End With
  
  If Not rsttmp.EOF Then bobines.Recordset!totalhores = total
  bobines.Recordset.Update
  On Error GoTo 0
  ensenyar_totalstotals
  Set rstmp = Nothing
End Sub
Sub ensenyar_totalstotals()
'total bobines
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalbobines) as elgran from extrussores totalbobines where comanda=" + atrim(cadbl(data1.Recordset!comanda)))
  If Not rsttmp.EOF Then tbobines = cadbl(rsttmp!elgran)

  
'hores func
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from extrussores totalbobines where comanda=" + atrim(cadbl(data1.Recordset!comanda)) + " and tipus='F'")
  If Not rsttmp.EOF Then hfunc = cadbl(rsttmp!elgran)
  
'hores canvi
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from extrussores totalbobines where comanda=" + atrim(cadbl(data1.Recordset!comanda)) + " and tipus='C'")
  If Not rsttmp.EOF Then hcanvi = cadbl(rsttmp!elgran)

'hores avaria
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from extrussores totalbobines where comanda=" + atrim(cadbl(data1.Recordset!comanda)) + " and tipus='A'")
  If Not rsttmp.EOF Then havaria = cadbl(rsttmp!elgran)

'total kilos
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalkilos) as elgran from extrussores totalbobines where comanda=" + atrim(cadbl(data1.Recordset!comanda)))
  If Not rsttmp.EOF Then tkilos = cadbl(rsttmp!elgran)


'total metres
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalmetres) as elgran from extrussores totalbobines where comanda=" + atrim(cadbl(data1.Recordset!comanda)))
  If Not rsttmp.EOF Then tmetres = cadbl(rsttmp!elgran)
  
  
End Sub
Private Sub eliminar_Click()
'If cadbl(bobines.Recordset!totalbobines) > 0 Then MsgBox "No es pot borrar aquest registre conte detall de bobines.": Exit Sub
Set rst = dbtmpb.OpenRecordset("select count(ID) as fs from extrussores where tipus='F' and comanda=" + atrim(cadbl(entradabaixes.comanda.Text)))
If rst.EOF Then
   Exit Sub
   Else
     If rst!fs < 2 Then MsgBox "No es pot borrar l'ultima linia tipus F", vbCritical + vbOKOnly, "Atenció": Exit Sub
End If
If MsgBox("Segur que vols borrar aquest registre [També borraras totes les linies]?", vbCritical + 4, "Atenció") = vbYes Then
     If Not bobines.Recordset.EOF Then
        dbtmpb.Execute "delete * from bobinesext where  controlid=" + atrim(bobines.Recordset!ID)
        bobines.Recordset.Delete
     End If
     bobines.Refresh
     DBGrid1.Refresh
  End If
End Sub

Private Sub Form_Activate()
ensenyar_totalstotals
DBGrid1.SetFocus
End Sub
Sub triarextrussora()
  Load formseleccio
  formseleccio.Caption = "Triar Màquina Extrussora"
  formseleccio.data1.DatabaseName = camicomandes
  formseleccio.data1.RecordSource = "select * from maquines where donadadebaixa=null and maquina='E' order by codi"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   DBGrid1.Text = atrim(formseleccio.data1.Recordset!codi)
  ' nomextrussora(0).Caption = atrim(formseleccio.data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 112 Then
    If Not bobines.Recordset.EOF Then
      If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
      bobines.Recordset.Update
    End If
    ensenyar_totalstotals
    bobines.Refresh
    bobines.Recordset.MoveLast
  End If
  If KeyCode = 27 Then
     If bobines.Recordset.EditMode > 0 Then
        bobines.Recordset.CancelUpdate
       Else: Unload Extrussores
     End If
  End If
 
      
End Sub

Private Sub Form_Load()
centerscreen Me
data1.DatabaseName = camicomandes
data1.RecordSource = "select * from comandes where comanda=" + atrim(cadbl(entradabaixes.comanda.Text))
bobines.DatabaseName = cami
bobines.RecordSource = "select * from extrussores where comanda=" + atrim(cadbl(entradabaixes.comanda.Text))
Set dbtmp = OpenDatabase(data1.DatabaseName)
Set dbtmpb = OpenDatabase(bobines.DatabaseName)
data1.Refresh
bobines.Refresh

End Sub

Sub lookupde(taula As String, Optional control1 As Control, Optional control2 As Control, Optional camp As String, Optional altres As String)
If camp = "" Then camp = "descripcio"
If altres = "clientsextres" Then camp = camp + ",observacions1,observacions2,obsext1,obsext2,obsimp1,obsimp2,obslam1,obslam2,obsreb1,obsreb2,obssol1,obssol2"
If Len(taula) < 20 Then
    Set rsttmp = dbtmp.OpenRecordset("select " + camp + " from " + taula + " where codi=" + atrim(cadbl(control1.Text)))
   Else: Set rsttmp = dbtmp.OpenRecordset(taula)
End If
If Not rsttmp.EOF Then
     control2 = atrim(rsttmp.Fields(0))
     If altres = "clientsextres" Then
      If atrim(Text32) = "" Then Text32 = atrim(rsttmp.Fields(1))
      If atrim(Text12) = "" Then Text12 = atrim(rsttmp.Fields(2))
      If atrim(Text34) = "" Then Text34 = atrim(rsttmp.Fields(3))
      If atrim(Text35) = "" Then Text35 = atrim(rsttmp.Fields(4))
      If atrim(Text77) = "" Then Text77 = atrim(rsttmp.Fields(5))
      If atrim(Text76) = "" Then Text76 = atrim(rsttmp.Fields(6))
      If atrim(Text93) = "" Then Text93 = atrim(rsttmp.Fields(7))
      If atrim(Text94) = "" Then Text94 = atrim(rsttmp.Fields(8))
      If atrim(Text108) = "" Then Text108 = atrim(rsttmp.Fields(9))
      If atrim(Text110) = "" Then Text110 = atrim(rsttmp.Fields(10))
      If atrim(Text17) = "" Then Text17 = atrim(rsttmp.Fields(11))
      If atrim(Text88) = "" Then Text88 = atrim(rsttmp.Fields(12))
     End If
    Else: control2 = ""
End If

End Sub

Sub possarvalordcamps(Optional tamany As Byte)
Dim t As Double
If cadbl(tamany) = 0 Then t = tamany
On Error Resume Next
 For Each objecte In formcomandes
    If TypeOf objecte Is Label Then objecte.BackStyle = 0
    If TypeOf objecte Is TextBox Or TypeOf objecte Is MaskEdBox Then
      If objecte.DataField <> "" Then
         If cadbl(tamany) = 0 Then t = tamany_camp(data1.Recordset.Fields(objecte.DataField))
         
        ' objecte.Name
         
          'assigno el format standard a tots els controls
         If TypeOf objecte Is MaskEdBox Then
          If objecte.Format = "" Then
              'objecte.Mask = mascara_camp(data1.Recordset.Fields(objecte.DataField))
              objecte.Format = format_camp(data1.Recordset.Fields(objecte.DataField))
          End If
         End If
         objecte.MaxLength = t
      End If
    End If
Next

End Sub

Private Sub Form_LostFocus()
   If Not canvissortirseccio Then End
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Me.Name = Screen.ActiveForm.Name Then
    actualitza_totals_ext
    Set rst = dbtmpb.OpenRecordset("select id from extrussores where tipus='F' and comanda=" + atrim(cadbl(entradabaixes.comanda)))
    While Not rst.EOF
     Set rst2 = dbtmpb.OpenRecordset("select controlid from bobinesext where controlid=" + atrim(cadbl(rst!ID)))
     If Not rst2.EOF Then GoTo sortir
     rst.MoveNext
    Wend
sortir:
    controlar_fiseccio "E", ruta, IIf(rst.EOF, False, True)
    Set rst = Nothing
  End If
    
  
End Sub
Sub actualitza_totals_ext()
 If bobines.Recordset.EOF And bobines.Recordset.BOF Then Exit Sub
  bobines.Recordset.MoveLast
  While atrim(bobines.Recordset!tipus) <> "F" And Not (bobines.Recordset.EOF Or bobines.Recordset.BOF)
    bobines.Recordset.MovePrevious
  Wend
  Set rsttmp = dbtmpb.OpenRecordset("select * from extrussorestot where comanda=" + atrim(cadbl(entradabaixes.comanda)))
  If rsttmp.EOF Then
      rsttmp.AddNew
    Else: rsttmp.Edit
  End If
  With rsttmp
    !comanda = cadbl(entradabaixes.comanda)
    !hcanvi = cadbl(hcanvi)
    !havaria = cadbl(havaria)
    !hfuncio = cadbl(hfunc)
    !tbobines = cadbl(tbobines)
    !tkilos = cadbl(tkilos)
    !tmetres = cadbl(tmetres)
    !kiloshora = cadbl(kiloshora)
    If Not (bobines.Recordset.EOF Or bobines.Recordset.BOF) Then
      If Not IsNull(bobines.Recordset!datafi) Then !dataextrussio = bobines.Recordset!datafi
      !extrussora = cadbl(bobines.Recordset!numeromaquina)
    End If
   .Update
  End With

End Sub
Private Sub hfunc_Change()
On Error Resume Next
  kiloshora = Format(cadbl(tkilos) / cadbl(hfunc), "#.00")
End Sub

Private Sub List1_Click()
  DBGrid1.Text = Mid(List1.Text, 1, 1)
  List1.Visible = False
  DBGrid1.SetFocus
End Sub

Private Sub List1_LostFocus()
  List1.Visible = False
End Sub

Private Sub Timer1_Timer()
colocardetall
End Sub

Private Sub tkilos_Change()
  On Error Resume Next
  kiloshora = Format(cadbl(tkilos) / cadbl(hfunc), "#.00")
End Sub
