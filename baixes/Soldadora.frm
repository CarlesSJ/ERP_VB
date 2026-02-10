VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Soldadores 
   Caption         =   "Baixes Soldadores"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   Icon            =   "Soldadora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ccomandaacabada 
      Caption         =   "Comanda Acabada?"
      Enabled         =   0   'False
      Height          =   210
      Left            =   7500
      TabIndex        =   98
      Top             =   7965
      Width           =   1815
   End
   Begin VB.CommandButton botoacabada 
      BackColor       =   &H0080FF80&
      Caption         =   "Comanda Acabada"
      Height          =   390
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   7530
      Width           =   1845
   End
   Begin VB.CommandButton botonoacabada 
      BackColor       =   &H008080FF&
      Caption         =   "No Acabada"
      Height          =   390
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   7545
      Width           =   1845
   End
   Begin VB.ComboBox combosacsocaixes 
      BackColor       =   &H00FF80FF&
      Height          =   315
      ItemData        =   "Soldadora.frx":0442
      Left            =   7515
      List            =   "Soldadora.frx":044C
      TabIndex        =   94
      TabStop         =   0   'False
      Text            =   "Caixes"
      ToolTipText     =   "Sacs o Caixes"
      Top             =   4155
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   45
      Top             =   3795
   End
   Begin VB.CommandButton eliminar 
      Height          =   300
      Left            =   255
      Picture         =   "Soldadora.frx":045E
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   4695
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
      Left            =   6780
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Soldadores"
      Top             =   1545
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Data data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
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
      Height          =   2850
      Left            =   120
      TabIndex        =   2
      Top             =   4515
      Width           =   9210
      Begin VB.CommandButton detall 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Detall"
         Height          =   240
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   540
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.ListBox List1 
         Height          =   840
         ItemData        =   "Soldadora.frx":0770
         Left            =   3195
         List            =   "Soldadora.frx":0780
         TabIndex        =   17
         Top             =   705
         Visible         =   0   'False
         Width           =   1275
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Soldadora.frx":07A9
         Height          =   2505
         Left            =   105
         OleObjectBlob   =   "Soldadora.frx":07BB
         TabIndex        =   16
         Top             =   270
         Width           =   9045
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   9645
         Picture         =   "Soldadora.frx":258A
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
      Height          =   735
      Left            =   135
      TabIndex        =   1
      Top             =   3345
      Width           =   9225
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Bobines Assignades"
         Height          =   495
         Left            =   8175
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   135
         Width           =   990
      End
      Begin VB.TextBox hparada 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   360
         Width           =   840
      End
      Begin VB.TextBox tsacs 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4695
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   360
         Width           =   840
      End
      Begin VB.TextBox mtrsmin 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   375
         Width           =   840
      End
      Begin VB.TextBox tunitats 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   375
         Width           =   840
      End
      Begin VB.TextBox hfunc 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3660
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   360
         Width           =   840
      End
      Begin VB.TextBox havaria 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   360
         Width           =   840
      End
      Begin VB.TextBox hcanvi 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   570
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Hores Parada"
         Height          =   195
         Left            =   2520
         TabIndex        =   89
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Unitats/Hora"
         Height          =   210
         Left            =   7215
         TabIndex        =   35
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sacs"
         Height          =   210
         Left            =   4680
         TabIndex        =   28
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Hores Func."
         Height          =   195
         Left            =   3645
         TabIndex        =   27
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Unitats"
         Height          =   210
         Left            =   5670
         TabIndex        =   25
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "H. Avaria"
         Height          =   210
         Left            =   1665
         TabIndex        =   23
         Top             =   150
         Width           =   990
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "H. Canvi"
         Height          =   210
         Left            =   540
         TabIndex        =   21
         Top             =   135
         Width           =   990
      End
   End
   Begin VB.Frame NomRebobinadora 
      Caption         =   "Capçalera"
      Enabled         =   0   'False
      Height          =   3360
      Left            =   135
      TabIndex        =   0
      Top             =   -15
      Width           =   9255
      Begin VB.TextBox ruta 
         Height          =   285
         Left            =   3870
         TabIndex        =   90
         Top             =   180
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "simulteneitatlam"
         DataSource      =   "data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Soldadora.frx":2950
         Left            =   5175
         List            =   "Soldadora.frx":2963
         TabIndex        =   43
         Top             =   1635
         Width           =   675
      End
      Begin VB.TextBox Text142 
         DataField       =   "texteimpressio"
         DataSource      =   "data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   675
         TabIndex        =   37
         ToolTipText     =   "Texte d'Impressió"
         Top             =   1350
         Width           =   4395
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         DataField       =   "comanda"
         DataSource      =   "data1"
         Height          =   285
         Left            =   1380
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   38
         Top             =   1335
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text81 
         DataField       =   "lotmatdesb2"
         DataSource      =   "data1"
         Height          =   285
         Left            =   3090
         TabIndex        =   45
         Top             =   1665
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
         TabIndex        =   47
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
         Left            =   1095
         TabIndex        =   48
         Top             =   1680
         WhatsThisHelpID =   3
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text98 
         DataField       =   "amplereb"
         DataSource      =   "data1"
         Height          =   285
         Left            =   5445
         TabIndex        =   50
         Top             =   1995
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
         _Version        =   327681
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text104 
         DataField       =   "diamextbob"
         DataSource      =   "data1"
         Height          =   285
         Left            =   8535
         TabIndex        =   54
         Top             =   2010
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text103 
         DataField       =   "mtrslinbob"
         DataSource      =   "data1"
         Height          =   285
         Left            =   7860
         TabIndex        =   55
         Top             =   2010
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text102 
         DataField       =   "simulteneitatreb"
         DataSource      =   "data1"
         Height          =   285
         Left            =   6645
         TabIndex        =   56
         Top             =   2010
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         DataField       =   "kilosbob"
         DataSource      =   "data1"
         Height          =   285
         Left            =   7185
         TabIndex        =   57
         Top             =   2010
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text127 
         DataSource      =   "data1"
         Height          =   285
         Left            =   6810
         TabIndex        =   58
         Top             =   2520
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text126 
         DataField       =   "espessorsol"
         DataSource      =   "data1"
         Height          =   285
         Left            =   5985
         TabIndex        =   59
         Top             =   2535
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text125 
         DataField       =   "fuellebocasol"
         DataSource      =   "data1"
         Height          =   285
         Left            =   5160
         TabIndex        =   60
         Top             =   2535
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text124 
         DataField       =   "fuellebasesol"
         DataSource      =   "data1"
         Height          =   285
         Left            =   4350
         TabIndex        =   61
         Top             =   2535
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text123 
         DataField       =   "solapasol"
         DataSource      =   "data1"
         Height          =   285
         Left            =   3540
         TabIndex        =   62
         Top             =   2535
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text122 
         DataField       =   "longitudsol"
         DataSource      =   "data1"
         Height          =   285
         Left            =   2730
         TabIndex        =   63
         Top             =   2535
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text121 
         DataField       =   "amplesol"
         DataSource      =   "data1"
         Height          =   285
         Left            =   1095
         TabIndex        =   64
         Top             =   2535
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text119 
         DataField       =   "ampleplegsol"
         DataSource      =   "data1"
         Height          =   285
         Left            =   1905
         TabIndex        =   65
         Top             =   2535
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tipussoldadura 
         DataSource      =   "data1"
         Height          =   285
         Left            =   4305
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   3015
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   327681
         BackColor       =   16777215
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text132 
         DataField       =   "unitatsxcaixa"
         DataSource      =   "data1"
         Height          =   285
         Left            =   6990
         TabIndex        =   79
         Top             =   3045
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text131 
         DataField       =   "unitatsxpaquet"
         DataSource      =   "data1"
         Height          =   285
         Left            =   6090
         TabIndex        =   80
         Top             =   3045
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox cunitatscomanda 
         DataField       =   "cantitatsol"
         DataSource      =   "data1"
         Height          =   285
         Left            =   8070
         TabIndex        =   84
         Top             =   3030
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text25 
         DataField       =   "materialex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   0
         TabIndex        =   85
         Top             =   315
         Visible         =   0   'False
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
         Left            =   0
         TabIndex        =   86
         Top             =   615
         Visible         =   0   'False
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         DataField       =   "cantitatex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   7170
         TabIndex        =   91
         Top             =   615
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text24 
         DataField       =   "COLORANTex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   0
         TabIndex        =   100
         Top             =   0
         Visible         =   0   'False
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Unitats Sol"
         DataSource      =   "data1"
         Height          =   270
         Index           =   5
         Left            =   8100
         TabIndex        =   99
         Top             =   2835
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Quantitat Reb:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   3
         Left            =   6060
         TabIndex        =   92
         Top             =   645
         Width           =   1080
      End
      Begin VB.Label nomsold 
         Caption         =   "NomSoldadora"
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
         Left            =   990
         TabIndex        =   87
         Top             =   3045
         Width           =   3075
      End
      Begin VB.Label nomrebo 
         Caption         =   "NomRebobinadora"
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
         Left            =   1260
         TabIndex        =   83
         Top             =   2055
         Width           =   3075
      End
      Begin VB.Label Label1 
         Caption         =   "Un. Paquet:"
         DataSource      =   "data1"
         Height          =   270
         Index           =   106
         Left            =   5940
         TabIndex        =   82
         Top             =   2820
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Un. Caixa"
         DataSource      =   "data1"
         Height          =   270
         Index           =   126
         Left            =   6990
         TabIndex        =   81
         Top             =   2820
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Tipus Soldadura:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   125
         Left            =   4380
         TabIndex        =   78
         Top             =   2805
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "Soldadora:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   115
         Left            =   105
         TabIndex        =   76
         Top             =   3045
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Rebobinadora:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   90
         Left            =   105
         TabIndex        =   75
         Top             =   2025
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Plegat:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   114
         Left            =   2055
         TabIndex        =   74
         Top             =   2310
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Ample:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   117
         Left            =   1245
         TabIndex        =   73
         Top             =   2295
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "B/L/F/BB:"
         DataSource      =   "data1"
         Height          =   270
         Index           =   118
         Left            =   180
         TabIndex        =   72
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Longitud:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   119
         Left            =   2775
         TabIndex        =   71
         Top             =   2310
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Solapa:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   120
         Left            =   3690
         TabIndex        =   70
         Top             =   2310
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Fuelle Ba:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   121
         Left            =   4395
         TabIndex        =   69
         Top             =   2310
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Fuelle Bo:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   122
         Left            =   5220
         TabIndex        =   68
         Top             =   2310
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Espessor:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   123
         Left            =   6045
         TabIndex        =   67
         Top             =   2310
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Mesura:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   124
         Left            =   6930
         TabIndex        =   66
         Top             =   2295
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "Kg/Mtrs/Diam:"
         DataSource      =   "data1"
         Height          =   270
         Index           =   4
         Left            =   7680
         TabIndex        =   53
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Sim:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   89
         Left            =   6300
         TabIndex        =   52
         Top             =   2010
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Ample Reb:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   88
         Left            =   4545
         TabIndex        =   51
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Lot Desb 1:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   65
         Left            =   180
         TabIndex        =   49
         Top             =   1725
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Lot Desb 2:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   66
         Left            =   2175
         TabIndex        =   46
         Top             =   1725
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Simult.Lam:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   73
         Left            =   4245
         TabIndex        =   44
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Texte:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   42
         ToolTipText     =   "Texte d'Impressió"
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "NºTinters:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   50
         Left            =   5175
         TabIndex        =   41
         Top             =   1395
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
         TabIndex        =   40
         Top             =   1350
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Impressora:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   60
         Left            =   6405
         TabIndex        =   39
         Top             =   1365
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
         TabIndex        =   32
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
         MouseIcon       =   "Soldadora.frx":2976
         MousePointer    =   99  'Custom
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   345
         Width           =   3450
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
         Left            =   6690
         TabIndex        =   13
         Top             =   1035
         Width           =   2370
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
         Left            =   4230
         TabIndex        =   12
         Top             =   1050
         Width           =   2295
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
         Left            =   1740
         TabIndex        =   11
         Top             =   1050
         Width           =   2340
      End
      Begin VB.Label Label1 
         Caption         =   "Color/Mat/Aditiu:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   20
         Left            =   330
         TabIndex        =   10
         Top             =   1050
         Width           =   1380
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
   Begin VB.Label Label11 
      Caption         =   "Tipus d'embalatge:"
      Height          =   270
      Left            =   5700
      TabIndex        =   95
      Top             =   4170
      Width           =   1650
   End
   Begin VB.Label Label9 
      Caption         =   "Prem F2 per sel.leccionar Taules..."
      Height          =   225
      Left            =   150
      TabIndex        =   33
      Top             =   7380
      Width           =   9120
   End
End
Attribute VB_Name = "Soldadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub comodi_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 39 Then KeyCode = 0: DBGrid1.SetFocus: SendKeys "{RIGHT}"
  If KeyCode = 37 Then KeyCode = 0: DBGrid1.SetFocus: SendKeys "{LEFT}"
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub botoacabada_Click()
  Dim vquant As Double
  vquant = cadbl(cunitatscomanda)
  If Not canvissortirseccio Then Exit Sub
   If cadbl(tunitats) > (vquant + (vquant / 10)) Or cadbl(tunitats) < (vquant - (vquant / 10)) Then
        MsgBox "LA QUANTITAT FABRICADA ESTÀ PER SOBRE O PER SOTA DEL 10% QUE S'HAVIA DEMANAT." + Chr(10) + "SISPLAU REVISEU QUE TOT SIGUI CORRECTE.", vbCritical, "A T E N C I Ó"
   End If
   botoacabada.Tag = "1"
   ccomandaacabada.Value = 1
   actualitzar_bobinesent cadbl(entradabaixes.comanda), atrim(ruta.Text)
   Unload Me
End Sub

Private Sub botonoacabada_Click()
   If Not canvissortirseccio Then Exit Sub
   botoacabada.Tag = "1"
   ccomandaacabada.Value = 0
   actualitzar_bobinesent cadbl(entradabaixes.comanda), atrim(ruta.Text)
   Unload Me
   entradabaixes.Visible = True
End Sub

Private Sub combosacsocaixes_Click()
  
    If combosacsocaixes <> "" Then DBGrid1.SetFocus
End Sub

Private Sub combosacsocaixes_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub combosacsocaixes_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub combosacsocaixes_LostFocus()
   If combosacsocaixes = "" Then MsgBox "Escull si son sacs o caixes.", vbCritical, "Atenció": combosacsocaixes.SetFocus
End Sub

Private Sub Command1_Click()
   Bobinesassignades.Show 1
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
  'lookupde "select descripcio from maquines where maquina='L' and codi=" + atrim(cadbl(data1.Recordset!laminadora)), , nomlaminadora(0)
  'possar_noms_adhesius True
  
   'carrega el nom de la rebobinadora
   lookupde "select descripcio from maquines where maquina='R' and codi=" + atrim(cadbl(data1.Recordset!rebobinadora)), , nomrebo
 
  
  'carrega el nom de la soldadora
  lookupde "select descripcio from maquines where maquina='s' and codi=" + atrim(cadbl(data1.Recordset!soldadora)), , nomsold
  
 'lookup de tipussoldadura
  Set rsttmp = dbtmp.OpenRecordset("select descripcio from tipussoldadura where codi='" + atrim((data1.Recordset!tipusoldadura)) + "'")
  If Not rsttmp.EOF Then
     tipussoldadura = atrim(rsttmp!descripcio)
    Else: tipussoldadura = ""
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
   triarSoldadora
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
  formseleccio.data1.RecordSource = "select * from operaris where maquina='S'"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   DBGrid1.Text = atrim(formseleccio.data1.Recordset!codi)
  End If
  Unload formseleccio
  
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
  If DBGrid1.col = 6 Then
    If InStr(1, "CAFP", UCase$(Chr$(KeyAscii))) = 0 Then
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
  
  'comprovo si la soldadora entrada es correcte
  If LastCol = 0 Then
   If cadbl(DBGrid1.Columns(0)) <> 0 Then
     Set rsttmp = dbtmp.OpenRecordset("select codi from maquines where maquina='S' and codi=" + atrim(cadbl(DBGrid1.Columns(0))))
     If rsttmp.EOF Then MsgBox "Aquesta Soldadora no Existeix": DBGrid1.Columns(0) = "": DBGrid1.col = 0
   End If
  End If
  
  'comprovo si l'operari entrat es correcte
  If LastCol = 1 Then
   If cadbl(DBGrid1.Columns(1)) <> 0 Then
     Set rsttmp = dbtmp.OpenRecordset("select codi from operaris where maquina='S' and codi=" + atrim(cadbl(DBGrid1.Columns(1))))
     If rsttmp.EOF Then MsgBox "Aquest Operari no Existeix": DBGrid1.Columns(1) = "": DBGrid1.col = 1
   End If
  End If
  
  
  calcular_totals
End Sub
Sub colocardetall()
 If Not bobines.Recordset.EOF Then
  If DBGrid1.Columns(10).Left > 0 And DBGrid1.Row >= 0 Then
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
  Unload detallbobsol
  On Error GoTo 0
  detallbobsol.Show 1
  calcular_totals
  DBGrid1.Row = 0
  DBGrid1.SetFocus
  End Sub
Sub calcular_totals()
  Dim total As Double
  Dim hores As Double
  If bobines.Recordset.EOF Then Exit Sub
  If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
  Set rsttmp = dbtmpb.OpenRecordset("select count(*) as elgran from bobinessol where controlid=" + atrim(bobines.Recordset!ID))
  If Not rsttmp.EOF Then bobines.Recordset!totalsacs = rsttmp!elgran
  
 ' Set rsttmp = dbtmpb.OpenRecordset("select sum(kilos) as elgran from bobinesimp where controlid=" + atrim(bobines.Recordset!id))
 ' If Not rsttmp.EOF Then bobines.Recordset!totalkilos = rsttmp!elgran
  
  Set rsttmp = dbtmpb.OpenRecordset("select sum(unitatsxsac) as elgran from bobinessol where controlid=" + atrim(bobines.Recordset!ID))
  If Not rsttmp.EOF Then bobines.Recordset!totalunitats = rsttmp!elgran
  
  
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
'total sacs
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalsacs) as elgran from Soldadores  where comanda=" + atrim(cadbl(data1.Recordset!comanda)))
  If Not rsttmp.EOF Then tsacs = cadbl(rsttmp!elgran)
'total unitats
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalunitats) as elgran from Soldadores  where comanda=" + atrim(cadbl(data1.Recordset!comanda)))
  If Not rsttmp.EOF Then tunitats = cadbl(rsttmp!elgran)

  
'hores func
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from Soldadores  where comanda=" + atrim(cadbl(data1.Recordset!comanda)) + " and tipus='F'")
  If Not rsttmp.EOF Then hfunc = cadbl(rsttmp!elgran)
  
'hores canvi
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from Soldadores  where comanda=" + atrim(cadbl(data1.Recordset!comanda)) + " and tipus='C'")
  If Not rsttmp.EOF Then hcanvi = cadbl(rsttmp!elgran)

'hores avaria
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from Soldadores  where comanda=" + atrim(cadbl(data1.Recordset!comanda)) + " and tipus='A'")
  If Not rsttmp.EOF Then havaria = cadbl(rsttmp!elgran)

'hores parada
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from Soldadores  where comanda=" + atrim(cadbl(data1.Recordset!comanda)) + " and tipus='P'")
  If Not rsttmp.EOF Then hparada = cadbl(rsttmp!elgran)
  
  
'total mtrs/minut
  If cadbl(hfunc) > 0 Then mtrsmin = IIf(cadbl(hfunc) > 0, Format(cadbl(tunitats) / cadbl(hfunc), "#,##0.00"), 0)
'acabada o no
  Set rsttmp = dbtmpb.OpenRecordset("select * from soldadorestot where comanda=" + atrim(cadbl(entradabaixes.comanda)))
  If Not rsttmp.EOF Then ccomandaacabada.Value = cadbl(rsttmp!acavada)
  
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
Set rst = dbtmpb.OpenRecordset("select count(ID) as fs from soldadores where tipus='F' and comanda=" + atrim(cadbl(entradabaixes.comanda.Text)))
If rst.EOF Then
   Exit Sub
   Else
     If rst!fs < 2 And atrim(bobines.Recordset!tipus) = "F" Then MsgBox "No es pot borrar l'ultima linia tipus F", vbCritical + vbOKOnly, "Atenció": Exit Sub
End If
If cadbl(bobines.Recordset!totalsacs) > 0 Then MsgBox "No es pot borrar aquest registre conte detall de bobines.": Exit Sub
If MsgBox("Segur que vols borrar aquest registre?  [També borraras totes les Bobines]", vbCritical + 4, "Atenció") = vbYes Then
     If Not bobines.Recordset.EOF Then
        dbtmpb.Execute "delete * from bobinessol where  controlid=" + atrim(bobines.Recordset!ID)
        bobines.Recordset.Delete
     End If
     bobines.Refresh
     DBGrid1.Refresh
  End If
End Sub

Private Sub Form_Activate()
ensenyar_totalstotals
comprovarsitepreuassignatosinoenviarunmail cadbl(entradabaixes.comanda)
DBGrid1.SetFocus
demanarsacsocaixes
End Sub
Sub triarSoldadora()
  Load formseleccio
  formseleccio.Caption = "Triar Màquina Soldadora"
  formseleccio.data1.DatabaseName = camicomandes
  formseleccio.data1.RecordSource = "select * from maquines where donadadebaixa=null and maquina='S' order by codi"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   DBGrid1.Text = atrim(formseleccio.data1.Recordset!codi)
  ' nomextrussora(0).Caption = atrim(formseleccio.data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub
Sub demanarsacsocaixes()
   Dim rsttmp As Recordset
   Set rsttmp = dbtmpb.OpenRecordset("select * from soldadorestot where comanda=" + atrim(cadbl(entradabaixes.comanda)))
   If Not rsttmp.EOF Then combosacsocaixes = atrim(rsttmp!sacsocaixes)
   If combosacsocaixes = "" Then
       combosacsocaixes.SetFocus
       SendKeys "%{DOWN}"
   End If
   Set rsttmp = Nothing
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
       Else: Unload Soldadores
     End If
  End If
 
      
End Sub

Private Sub Form_Load()
centerscreen Me
data1.DatabaseName = camicomandes
data1.RecordSource = "select * from comandes where comanda=" + atrim(cadbl(entradabaixes.comanda.Text))
bobines.DatabaseName = cami
bobines.RecordSource = "select * from Soldadores where comanda=" + atrim(cadbl(entradabaixes.comanda.Text)) + " order by datainici,horainici"
Set dbtmp = OpenDatabase(data1.DatabaseName)
Set dbtmpb = OpenDatabase(bobines.DatabaseName)
data1.Refresh
bobines.Refresh
Set dbstocks = OpenDatabase(rutadelfitxer(camicomandes) + "palets.mdb")
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
   If Me.Name <> Screen.ActiveForm.Name Then Exit Sub
   If canvissortirseccio Then
     If botoacabada.Tag = "" Then
        MsgBox "Per tancar la baixa escull comanda Acabada o No Acabada.", vbCritical, "Error": Cancel = 1
         Else: rutinadesurtidadesoldadores
     End If
       Else: rutinadesurtidadesoldadores
   End If
   
End Sub
Sub rutinadesurtidadesoldadores()
    If Me.Name = Screen.ActiveForm.Name Then
      actualitza_totals_sol
      Set rst = dbtmpb.OpenRecordset("select id from soldadores where tipus='F' and comanda=" + atrim(cadbl(entradabaixes.comanda)))
      While Not rst.EOF
        Set rst2 = dbtmpb.OpenRecordset("select controlid from bobinessol where controlid=" + atrim(cadbl(rst!ID)))
        If Not rst2.EOF Then GoTo sortir
        rst.MoveNext
      Wend
sortir:
      If ccomandaacabada.Value = 1 Then
         controlar_fiseccio "S", ruta, IIf(rst.EOF, False, True)
           Else: entradabaixes.Visible = True
      End If
      Set dbstocks = Nothing
    End If
End Sub
Sub actualitza_totals_sol()
If bobines.Recordset.EOF And bobines.Recordset.BOF Then Exit Sub
  Set rsttmp = dbtmpb.OpenRecordset("select * from soldadorestot where comanda=" + atrim(cadbl(entradabaixes.comanda)))
  If rsttmp.EOF Then
      rsttmp.AddNew
    Else: rsttmp.Edit
  End If
  With rsttmp
    !comanda = cadbl(entradabaixes.comanda)
    !hcanvi = cadbl(hcanvi)
    !havaria = cadbl(havaria)
    !hparada = cadbl(hparada)
    !tsacs = cadbl(tsacs)
    !tunitats = cadbl(tunitats)
    !unitatshora = cadbl(mtrsmin)
    !sacsocaixes = atrim(combosacsocaixes)
    !acavada = atrim(ccomandaacabada)
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
  If Not (bobines.Recordset.EOF And bobines.Recordset.BOF) Then
    If bobines.Recordset.RecordCount = 0 Then
      avisarquelacomandasestaacabant data1.Recordset!comanda, "S"
    End If
  End If
End Sub
Sub avisarquelacomandasestaacabant(vnumc As Double, vseccioactual As String)
  Dim rst As Recordset
  Dim vruta As String
  Set rst = dbtmp.OpenRecordset("SELECT comandes.direnvio,comandes.comanda, comandes.producte, productes.ruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi where comanda=" + atrim(vnumc))
  If rst.EOF Then GoTo fi
  vruta = atrim(rst!ruta)
  If vseccioactual = Mid(vruta, Len(vruta), 1) Then
      Set rst = dbtmp.OpenRecordset("select * from clients_envios where id=" + atrim(cadbl(rst!direnvio)))
      If rst.EOF Then GoTo fi
         If atrim(rst!avisfiproduccio) <> "" Then
             avisarfiproduccio "La comanda " + atrim(vnumc) + " està acabant la producció.", atrim(rst!avisfiproduccio)
         End If
  End If
fi:
  Set rst = Nothing
End Sub
Sub avisarfiproduccio(assumpte As String, cos As String)
   Dim rutamdb As String
   Dim dbavisos As Database
   Dim rsta As Recordset
   Dim destinatari As String
   
   destinatari = "avisfiproduccio"
   rutamdb = rutadelfitxer(cami) + "avisosincidencies.mdb"
   Set dbavisos = DBEngine.OpenDatabase(rutamdb)
   Set rsta = dbavisos.OpenRecordset("select * from envios_mails where assumpte='" + atrim(assumpte) + "'")
   If rsta.EOF Then
      dbavisos.Execute "insert into envios_mails (data,destinatari,assumpte,cos) values (now,'" + destinatari + "','" + atrim(assumpte) + "','" + atrim(cos) + "')"
   End If
   Set rst = Nothing
   dbavisos.Close
   Set dbavisos = Nothing
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

