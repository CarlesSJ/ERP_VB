VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Laminadores 
   Caption         =   "Baixes Laminadores"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   Icon            =   "Laminadores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   30
      Top             =   3450
   End
   Begin VB.CommandButton eliminar 
      Height          =   300
      Left            =   210
      Picture         =   "Laminadores.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   4140
      Width           =   300
   End
   Begin VB.Data bobines 
      Caption         =   "bobines"
      Connect         =   "Access"
      DatabaseName    =   "Y:\comandes\baixes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6330
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Laminadores"
      Top             =   3000
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
      Height          =   3270
      Left            =   120
      TabIndex        =   2
      Top             =   3945
      Width           =   9210
      Begin VB.CommandButton detall 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Detall"
         Height          =   240
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   540
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.ListBox List1 
         Height          =   840
         ItemData        =   "Laminadores.frx":0754
         Left            =   3195
         List            =   "Laminadores.frx":0761
         TabIndex        =   34
         Top             =   705
         Visible         =   0   'False
         Width           =   1275
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Laminadores.frx":0782
         Height          =   2910
         Left            =   105
         OleObjectBlob   =   "Laminadores.frx":0794
         TabIndex        =   33
         Top             =   255
         Width           =   9045
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   9645
         Picture         =   "Laminadores.frx":2583
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
      Left            =   135
      TabIndex        =   1
      Top             =   3150
      Width           =   9225
      Begin VB.TextBox mtrsmin 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   8145
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   375
         Width           =   840
      End
      Begin VB.TextBox tbob 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5100
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   405
         Width           =   840
      End
      Begin VB.TextBox tenduridor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox grmsmtr2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   375
         Width           =   840
      End
      Begin VB.TextBox tmetres 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6030
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   375
         Width           =   840
      End
      Begin VB.TextBox tresina 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox hfunc 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox havaria 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox hcanvi 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   390
         Width           =   840
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Metres/Min"
         Height          =   210
         Left            =   8115
         TabIndex        =   87
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "T. Enduridor"
         Height          =   210
         Left            =   4095
         TabIndex        =   56
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Grms/m2"
         Height          =   210
         Left            =   7095
         TabIndex        =   54
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Metres"
         Height          =   210
         Left            =   6030
         TabIndex        =   47
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bob."
         Height          =   210
         Left            =   5085
         TabIndex        =   45
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Hores Func."
         Height          =   195
         Left            =   2145
         TabIndex        =   44
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "T. Resina"
         Height          =   210
         Left            =   3210
         TabIndex        =   42
         Top             =   195
         Width           =   840
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "H. Avaria"
         Height          =   210
         Left            =   1245
         TabIndex        =   40
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "H. Canvi"
         Height          =   210
         Left            =   240
         TabIndex        =   38
         Top             =   165
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Capçalera"
      Enabled         =   0   'False
      Height          =   3015
      Left            =   135
      TabIndex        =   0
      Top             =   60
      Width           =   9255
      Begin VB.TextBox ruta 
         Height          =   285
         Left            =   3825
         TabIndex        =   81
         Top             =   270
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox grcm2 
         Height          =   285
         Left            =   3255
         TabIndex        =   80
         Top             =   2640
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox grcm1 
         Height          =   285
         Left            =   2790
         TabIndex        =   79
         Top             =   2700
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "simulteneitatlam"
         DataSource      =   "data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Laminadores.frx":2949
         Left            =   8340
         List            =   "Laminadores.frx":295C
         TabIndex        =   66
         Top             =   2070
         Width           =   675
      End
      Begin VB.TextBox Text142 
         DataField       =   "texteimpressio"
         DataSource      =   "data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   675
         TabIndex        =   58
         ToolTipText     =   "Texte d'Impressió"
         Top             =   1935
         Width           =   4395
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "simulteneitat"
         DataSource      =   "data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Laminadores.frx":296F
         Left            =   8460
         List            =   "Laminadores.frx":2982
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
         Top             =   975
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
         Top             =   1290
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
         Top             =   1590
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
      Begin MSMask.MaskEdBox Text63 
         DataField       =   "numerotintes"
         DataSource      =   "data1"
         Height          =   285
         Left            =   5910
         TabIndex        =   59
         Top             =   1920
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text91 
         DataField       =   "camisa"
         DataSource      =   "data1"
         Height          =   285
         Left            =   8355
         TabIndex        =   68
         Top             =   2370
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox grmt2 
         DataField       =   "grmt2"
         DataSource      =   "data1"
         Height          =   285
         Left            =   6525
         TabIndex        =   70
         Top             =   2310
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text81 
         DataField       =   "lotmatdesb2"
         DataSource      =   "data1"
         Height          =   285
         Left            =   5445
         TabIndex        =   72
         Top             =   2655
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox vadhesiu 
         DataField       =   "tipusadhesiu"
         DataSource      =   "data1"
         Height          =   285
         Left            =   2085
         TabIndex        =   74
         Top             =   195
         Visible         =   0   'False
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text80 
         DataField       =   "lotmatdesb1"
         DataSource      =   "data1"
         Height          =   285
         Left            =   5445
         TabIndex        =   77
         Top             =   2355
         WhatsThisHelpID =   3
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text22 
         DataField       =   "micropex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   8685
         TabIndex        =   82
         ToolTipText     =   "Mesura de l'Espessor"
         Top             =   1575
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
         TabIndex        =   83
         Top             =   1575
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
         TabIndex        =   85
         Top             =   1605
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Quantitat:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   3
         Left            =   6225
         TabIndex        =   84
         Top             =   1620
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Lot Desb 1:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   65
         Left            =   4530
         TabIndex        =   78
         Top             =   2400
         Width           =   825
      End
      Begin VB.Label adhesiu 
         Caption         =   "DESCRIPCIO DE L'ADHESIU"
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
         ForeColor       =   &H00FF8080&
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   76
         Top             =   2745
         Width           =   4245
      End
      Begin VB.Label Label1 
         Caption         =   "Descripció Adhesiu i Enduridor"
         DataSource      =   "data1"
         Height          =   255
         Index           =   75
         Left            =   120
         TabIndex        =   75
         Top             =   2535
         Width           =   2610
      End
      Begin VB.Label Label1 
         Caption         =   "Lot Desb 2:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   66
         Left            =   4530
         TabIndex        =   73
         Top             =   2715
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Cola Gr/mt2:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   81
         Left            =   6375
         TabIndex        =   71
         Top             =   2115
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Camisa:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   76
         Left            =   7470
         TabIndex        =   69
         Top             =   2430
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Simult.Lam:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   73
         Left            =   7410
         TabIndex        =   67
         Top             =   2115
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Laminadora:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   67
         Left            =   105
         TabIndex        =   65
         Top             =   2355
         Width           =   1005
      End
      Begin VB.Label nomlaminadora 
         Caption         =   "nomlaminadora"
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
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   0
         Left            =   1185
         TabIndex        =   64
         Top             =   2340
         Width           =   3180
      End
      Begin VB.Label Label1 
         Caption         =   "Texte:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   63
         ToolTipText     =   "Texte d'Impressió"
         Top             =   1965
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "NºTinters:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   50
         Left            =   5175
         TabIndex        =   62
         Top             =   1980
         Width           =   750
      End
      Begin VB.Label nomimpressora 
         Caption         =   "nomimpressora"
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
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   1
         Left            =   7290
         TabIndex        =   61
         Top             =   1875
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Impressora:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   60
         Left            =   6405
         TabIndex        =   60
         Top             =   1890
         Width           =   795
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
         TabIndex        =   51
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
         MouseIcon       =   "Laminadores.frx":2995
         MousePointer    =   99  'Custom
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         Caption         =   "Sim.Ext:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   33
         Left            =   7875
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
         Top             =   1665
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
         Top             =   1665
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
         Top             =   1365
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
         Top             =   1350
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
         Top             =   1050
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
         Top             =   1050
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
      Left            =   195
      TabIndex        =   52
      Top             =   7215
      Width           =   9120
   End
End
Attribute VB_Name = "Laminadores"
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
  'carrega el nom de la impressora
  lookupde "select descripcio from maquines where maquina='I' and codi=" + atrim(cadbl(data1.Recordset!impressora)), , nomimpressora(1)
  'laminadora
  lookupde "select descripcio from maquines where maquina='L' and codi=" + atrim(cadbl(data1.Recordset!laminadora)), , nomlaminadora(0)
  possar_noms_adhesius True
  
  
End Sub
Function possar_metres_min() As Double
  Dim v As Double
  v = cadbl(hfunc)
  f = (Int(v) * 60) + (((v - Int(v)) * 100) * 60 / 100)
 On Error Resume Next
  possar_metres_min = Format(cadbl(tmetres) / (f), "#.00")
End Function


Sub possar_noms_adhesius(Optional lookup As Boolean)
  Set rsttmp = dbtmp.OpenRecordset("select * from adhesius where codi=" + atrim(cadbl(data1.Recordset!tipusadhesiu)))
  If Not rsttmp.EOF Then
    adhesiu(1) = atrim(rsttmp!resina)
    adhesiu(1) = adhesiu(1) + " + " + atrim(rsttmp!enduridor)
    grcm1 = cadbl(rsttmp!grmcm3_resina)
    grcm2 = cadbl(rsttmp!grmcm3_ENDURIDOR)
    
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
 If ColIndex = 6 Then
   List1.Visible = True
   'List1.Width = DBGrid1.Columns(ColIndex).Width
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
   triarlaminadora
 End If
 If KeyCode = 113 And DBGrid1.col = 1 Then
   triaroperaris
 End If
  If (KeyCode = Asc("D") Or KeyCode = Asc("d")) And Shift = 2 Then
    detall_Click
  End If
 
End Sub

Sub triaroperaris()
  Load formseleccio
  formseleccio.Caption = "Triar Operaris"
  formseleccio.data1.DatabaseName = camicomandes
  formseleccio.data1.RecordSource = "select * from operaris where maquina='L'"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   DBGrid1.Text = atrim(formseleccio.data1.Recordset!codi)
  End If
  Unload formseleccio
  
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
  If DBGrid1.col = 6 Then
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
  If DBGrid1.col = 2 Then
    If atrim(DBGrid1.Text) = "" Then DBGrid1.Text = Format(DateAdd("d", -1, Now), "dd/mm/yy")
  End If
  
  If DBGrid1.col = 4 Then
    If atrim(DBGrid1.Text) = "" Then DBGrid1.Text = Format(DBGrid1.Columns(2).Text, "dd/mm/yy")
  End If
  
  
  '-------
  
  If LastCol = 2 Or LastCol = 3 Then
  valtmp = DBGrid1.Columns(LastCol).Text
  If LastCol = 2 Then
      
      If InStr(1, valtmp, "/") = 0 Then valtmp = Mid(valtmp, 1, 2) + "/" + Mid(valtmp, 3, 2) + "/" + Mid(valtmp, 5, 2)
      If Not IsDate(valtmp) Then valtmp = ""
  End If
  If LastCol = 3 Then
    If InStr(1, valtmp, ":") = 0 Then valtmp = Mid(valtmp, 1, 2) + ":" + Mid(valtmp, 3, 2)
      If Not IsDate(Format(valtmp, "hh:nn")) Then valtmp = "00:00"

  End If
  DBGrid1.Columns(LastCol) = valtmp
  End If
  
  If LastCol = 4 Or LastCol = 5 Then
  valtmp = DBGrid1.Columns(LastCol).Text
  If LastCol = 4 Then
      
      If InStr(1, valtmp, "/") = 0 Then valtmp = Mid(valtmp, 1, 2) + "/" + Mid(valtmp, 3, 2) + "/" + Mid(valtmp, 5, 2)
      If Not IsDate(valtmp) Then valtmp = ""
  End If
  If LastCol = 5 Then
    If InStr(1, valtmp, ":") = 0 Then valtmp = Mid(valtmp, 1, 2) + ":" + Mid(valtmp, 3, 2)
      If Not IsDate(Format(valtmp, "hh:nn")) Then valtmp = "00:00"

  End If
  DBGrid1.Columns(LastCol) = valtmp
  End If
  
  'comprovo si la laminadora entrada es correcte
  If LastCol = 0 Then
   If cadbl(DBGrid1.Columns(0)) <> 0 Then
     Set rsttmp = dbtmp.OpenRecordset("select codi from maquines where maquina='L' and codi=" + atrim(cadbl(DBGrid1.Columns(0))))
     If rsttmp.EOF Then MsgBox "Aquesta Laminadora no Existeix": DBGrid1.Columns(0) = "": DBGrid1.col = 0
   End If
  End If
  
  'comprovo si l'operari entrat es correcte
  If LastCol = 1 Then
   If cadbl(DBGrid1.Columns(1)) <> 0 Then
     Set rsttmp = dbtmp.OpenRecordset("select codi from operaris where maquina='L' and codi=" + atrim(cadbl(DBGrid1.Columns(1))))
     If rsttmp.EOF Then MsgBox "Aquest Operari no Existeix": DBGrid1.Columns(1) = "": DBGrid1.col = 1
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
  Unload detallboblam
  On Error GoTo 0
  detallboblam.Show 1
  calcular_totals
  DBGrid1.Row = 0
  DBGrid1.SetFocus
  End Sub
Sub calcular_totals()
  Dim total As Double
  Dim hores As Double
  If bobines.Recordset.EOF Then Exit Sub
  If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
  Set rsttmp = dbtmpb.OpenRecordset("select count(*) as elgran from bobineslam where controlid=" + atrim(bobines.Recordset!ID))
  If Not rsttmp.EOF Then bobines.Recordset!totalbobines = rsttmp!elgran
  
 ' Set rsttmp = dbtmpb.OpenRecordset("select sum(kilos) as elgran from bobinesimp where controlid=" + atrim(bobines.Recordset!id))
 ' If Not rsttmp.EOF Then bobines.Recordset!totalkilos = rsttmp!elgran
  
  Set rsttmp = dbtmpb.OpenRecordset("select sum(metres) as elgran from bobineslam where controlid=" + atrim(bobines.Recordset!ID))
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
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalbobines) as elgran from Laminadores  where comanda=" + atrim(cadbl(data1.Recordset!comanda)))
  If Not rsttmp.EOF Then tbob = cadbl(rsttmp!elgran)

  
'hores func
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from Laminadores  where comanda=" + atrim(cadbl(data1.Recordset!comanda)) + " and tipus='F'")
  If Not rsttmp.EOF Then hfunc = cadbl(rsttmp!elgran)
  
'hores canvi
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from Laminadores  where comanda=" + atrim(cadbl(data1.Recordset!comanda)) + " and tipus='C'")
  If Not rsttmp.EOF Then hcanvi = cadbl(rsttmp!elgran)

'hores avaria
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from Laminadores  where comanda=" + atrim(cadbl(data1.Recordset!comanda)) + " and tipus='A'")
  If Not rsttmp.EOF Then havaria = cadbl(rsttmp!elgran)

'total resina
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totallitresresina) as elgran from Laminadores  where comanda=" + atrim(cadbl(data1.Recordset!comanda)))
  If Not rsttmp.EOF Then tresina = cadbl(rsttmp!elgran)
  
'total enduridor
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totallitresenduridor) as elgran from Laminadores  where comanda=" + atrim(cadbl(data1.Recordset!comanda)))
  If Not rsttmp.EOF Then tenduridor = cadbl(rsttmp!elgran)
  
'total metres
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalmetres) as elgran from Laminadores  where comanda=" + atrim(cadbl(data1.Recordset!comanda)))
  If Not rsttmp.EOF Then tmetres = cadbl(rsttmp!elgran)
  
  
'total grmsmtr2
  grmsmtr2.Text = calcular_grmsmtr2(cadbl(tresina), cadbl(tenduridor), cadbl(tmetres), cadbl(Text91), cadbl(grcm1), cadbl(grcm2))
  'aixo es un exemple per veure si funciona
  'grmsmtr2.Text = calcular_grmsmtr2(33.9, 15.6, 28950, 89, 1.16, 1.16)
  
  If cadbl(hfunc) > 0 Then mtrsmin = possar_metres_min
End Sub
Function calcular_grmsmtr2(tresina As Double, tenduridor As Double, tmetres As Double, camisa As Double, gramsresina As Double, gramsenduridor As Double) As Double
  Dim result1 As Double
  Dim result2 As Double
  On Error Resume Next
  result1 = (tresina * 1000 * gramsresina) / (tmetres * (camisa / 100))
  result2 = (tenduridor * 1000 * gramsenduridor) / (tmetres * (camisa / 100))
  
  calcular_grmsmtr2 = cadbl(Format(result1 + result2, "#,##0.00"))
End Function

Private Sub eliminar_Click()
Set rst = dbtmpb.OpenRecordset("select count(ID) as fs from laminadores where tipus='F' and comanda=" + atrim(cadbl(entradabaixes.comanda.Text)))
If rst.EOF Then
   Exit Sub
   Else
     If rst!fs < 2 Then MsgBox "No es pot borrar l'ultima linia tipus F", vbCritical + vbOKOnly, "Atenció": Exit Sub
End If
If cadbl(bobines.Recordset!totalbobines) > 0 Then MsgBox "No es pot borrar aquest registre conte detall de bobines.": Exit Sub
If MsgBox("Segur que vols borrar aquest registre. [També borraras totes les Bobines]?", vbCritical + 4, "Atenció") = vbYes Then
     If Not bobines.Recordset.EOF Then
        dbtmpb.Execute "delete * from bobineslam where  controlid=" + atrim(bobines.Recordset!ID)
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
Sub triarlaminadora()
  Load formseleccio
  formseleccio.Caption = "Triar Màquina Laminadora"
  formseleccio.data1.DatabaseName = camicomandes
  formseleccio.data1.RecordSource = "select * from maquines where donadadebaixa=null and maquina='L' order by codi"
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
       Else: Unload Laminadores
     End If
  End If
 
      
End Sub

Private Sub Form_Load()
centerscreen Me
data1.DatabaseName = camicomandes
data1.RecordSource = "select * from comandes where comanda=" + atrim(cadbl(entradabaixes.comanda.Text))
bobines.DatabaseName = cami
bobines.RecordSource = "select * from Laminadores where comanda=" + atrim(cadbl(entradabaixes.comanda.Text))
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

Private Sub hclixe_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Me.Name = Screen.ActiveForm.Name Then
  actualitza_totals_lam
  Set rst = dbtmpb.OpenRecordset("select id from laminadores where tipus='F' and comanda=" + atrim(cadbl(entradabaixes.comanda)))
  While Not rst.EOF
    Set rst2 = dbtmpb.OpenRecordset("select controlid from bobineslam where controlid=" + atrim(cadbl(rst!ID)))
    If Not rst2.EOF Then GoTo sortir
    rst.MoveNext
  Wend
sortir:
  controlar_fiseccio "L", ruta, IIf(rst.EOF, False, True)
End If
End Sub
Sub actualitza_totals_lam()
  If bobines.Recordset.EOF And bobines.Recordset.BOF Then Exit Sub
  bobines.Recordset.MoveLast
  While atrim(bobines.Recordset!tipus) <> "F" And Not (bobines.Recordset.EOF Or bobines.Recordset.BOF)
    bobines.Recordset.MovePrevious
  Wend
  Set rsttmp = dbtmpb.OpenRecordset("select * from laminadorestot where comanda=" + atrim(cadbl(entradabaixes.comanda)))
  If rsttmp.EOF Then
      rsttmp.AddNew
    Else: rsttmp.Edit
  End If
  With rsttmp
    !comanda = cadbl(entradabaixes.comanda)
    !hcanvi = cadbl(hcanvi)
    !havaria = cadbl(havaria)
    !hfuncio = cadbl(hfunc)
    !tbobines = cadbl(tbob)
    !tresina = cadbl(tresina)
    !tenduridor = cadbl(tenduridor)
    !tmetres = cadbl(tmetres)
    '!metresmin = cadbl(bobines.Recordset!mtrsminut)
    !metresmin = cadbl(possar_metres_min)
    !grmmtr2 = cadbl(grmsmtr2)
    If Not (bobines.Recordset.EOF Or bobines.Recordset.BOF) Then
      If Not IsNull(bobines.Recordset!datafi) Then !datalaminacio = bobines.Recordset!datafi
      !operari = cadbl(bobines.Recordset!operari)
      !laminadora = cadbl(bobines.Recordset!numeromaquina)
    End If
   .Update
  End With

End Sub

Private Sub hfunc_Change()
On Error Resume Next
  kiloshora = Format(cadbl(tkilos) / cadbl(hfunc), "#.00")
End Sub

Private Sub kiloshora_Change()

End Sub

Private Sub List1_Click()
  DBGrid1.Text = Mid(List1.Text, 1, 1)
  List1.Visible = False
  DBGrid1.SetFocus
End Sub

Private Sub List1_LostFocus()
  List1.Visible = False
End Sub

Private Sub tkilos_Change()
  On Error Resume Next
  kiloshora = Format(cadbl(tkilos) / cadbl(hfunc), "#.00")
End Sub

Private Sub Timer1_Timer()
colocardetall
End Sub
