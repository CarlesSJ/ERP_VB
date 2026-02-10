VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Entrega 
   Caption         =   "Baixes Entrega"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   Icon            =   "Entrega.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080FF80&
      Caption         =   "Bobines Assignades"
      Height          =   495
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   2385
      Width           =   990
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   600
      Left            =   60
      TabIndex        =   78
      Top             =   2130
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entrada de dades"
      Height          =   885
      Left            =   135
      TabIndex        =   71
      Top             =   3345
      Width           =   9240
      Begin VB.TextBox ccalloff 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5865
         TabIndex        =   101
         Top             =   480
         Width           =   1650
      End
      Begin VB.TextBox ccontracte 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   99
         Top             =   480
         Width           =   1650
      End
      Begin VB.TextBox cnumalb 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   660
         TabIndex        =   98
         Top             =   510
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Dia -"
         Height          =   270
         Left            =   3060
         TabIndex        =   82
         Top             =   165
         Width           =   525
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Dia +"
         Height          =   270
         Left            =   2520
         TabIndex        =   79
         Top             =   165
         Width           =   525
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Entregar"
         Height          =   345
         Left            =   7815
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   135
         Width           =   1350
      End
      Begin VB.ComboBox Transportista 
         Height          =   315
         ItemData        =   "Entrega.frx":0442
         Left            =   4395
         List            =   "Entrega.frx":0444
         TabIndex        =   74
         Top             =   135
         Width           =   2595
      End
      Begin VB.TextBox dataentrega 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1245
         TabIndex        =   72
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label12 
         Caption         =   "Call Off:"
         Height          =   210
         Left            =   5235
         TabIndex        =   102
         Top             =   510
         Width           =   765
      End
      Begin VB.Label Label11 
         Caption         =   "Contracte:"
         Height          =   210
         Left            =   2520
         TabIndex        =   100
         Top             =   510
         Width           =   1035
      End
      Begin VB.Label Label10 
         Caption         =   "Albarà:"
         Height          =   210
         Left            =   90
         TabIndex        =   97
         Top             =   540
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "Transport:"
         Height          =   210
         Left            =   3645
         TabIndex        =   75
         Top             =   195
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Data d'Entrega:"
         Height          =   210
         Left            =   60
         TabIndex        =   73
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   8820
      Top             =   3630
   End
   Begin VB.Data bobines 
      Caption         =   "bobines"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6825
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "bobinesent"
      Top             =   2970
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
      Caption         =   "Desguas de Bobines"
      Height          =   3000
      Left            =   150
      TabIndex        =   1
      Top             =   4305
      Width           =   9210
      Begin VB.CommandButton comandanoacabada 
         Caption         =   "Comanda No Acabada"
         Height          =   315
         Left            =   7275
         TabIndex        =   95
         Top             =   2580
         Width           =   1815
      End
      Begin VB.TextBox entregatk 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   8340
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   1410
         Width           =   840
      End
      Begin VB.TextBox pendentk 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   8340
         TabIndex        =   90
         Top             =   1740
         Width           =   840
      End
      Begin VB.TextBox pendentm 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   7455
         TabIndex        =   87
         Top             =   1755
         Width           =   840
      End
      Begin VB.TextBox entregatm 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   7455
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   1410
         Width           =   840
      End
      Begin VB.CheckBox retornmaterial 
         Caption         =   "Retorn Material"
         Height          =   210
         Left            =   7590
         TabIndex        =   85
         Top             =   780
         Width           =   1560
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selec. Tot o DesSeleccionar"
         Height          =   480
         Left            =   7470
         TabIndex        =   81
         Top             =   225
         Width           =   1635
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Entrega.frx":0446
         Height          =   2715
         Left            =   105
         OleObjectBlob   =   "Entrega.frx":0458
         TabIndex        =   77
         Top             =   210
         Width           =   7110
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   9645
         Picture         =   "Entrega.frx":1D5F
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1740
         Width           =   315
      End
      Begin MSMask.MaskEdBox Text31 
         DataField       =   "mesuracantex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   9630
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
         Top             =   840
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   503
         _Version        =   327681
         MaxLength       =   1
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin VB.Label lseccioenruta 
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
         Height          =   240
         Left            =   6795
         TabIndex        =   96
         Top             =   2250
         Width           =   2370
      End
      Begin VB.Label Label8 
         Caption         =   "Kilos"
         Height          =   240
         Left            =   8415
         TabIndex        =   93
         Top             =   1110
         Width           =   435
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Mtrs o Un."
         Height          =   240
         Left            =   7425
         TabIndex        =   92
         Top             =   1110
         Width           =   960
      End
      Begin VB.Line Line2 
         X1              =   8295
         X2              =   8295
         Y1              =   1170
         Y2              =   2265
      End
      Begin VB.Line Line1 
         X1              =   7350
         X2              =   9075
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Label Label6 
         Caption         =   "E:"
         Height          =   270
         Left            =   7305
         TabIndex        =   89
         Top             =   1425
         Width           =   285
      End
      Begin VB.Label Label5 
         Caption         =   "P:"
         Height          =   270
         Left            =   7305
         TabIndex        =   88
         Top             =   1785
         Width           =   285
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
         TabIndex        =   70
         Top             =   180
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "simulteneitatlam"
         DataSource      =   "data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Entrega.frx":2125
         Left            =   5175
         List            =   "Entrega.frx":2138
         TabIndex        =   25
         Top             =   1635
         Width           =   675
      End
      Begin VB.TextBox Text142 
         DataField       =   "texteimpressio"
         DataSource      =   "data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   675
         TabIndex        =   19
         ToolTipText     =   "Texte d'Impressió"
         Top             =   1350
         Width           =   4395
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         DataField       =   "comanda"
         DataSource      =   "data1"
         Height          =   285
         Left            =   1380
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   20
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
         TabIndex        =   27
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
         TabIndex        =   29
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
         TabIndex        =   30
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
         TabIndex        =   32
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
         TabIndex        =   36
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
         TabIndex        =   37
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
         TabIndex        =   38
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
         TabIndex        =   39
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
         TabIndex        =   40
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
         TabIndex        =   41
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
         TabIndex        =   42
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
         TabIndex        =   43
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
         TabIndex        =   44
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
         TabIndex        =   45
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
         TabIndex        =   46
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
         TabIndex        =   47
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
         TabIndex        =   59
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
         TabIndex        =   61
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
         TabIndex        =   62
         Top             =   3045
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   327681
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text24 
         DataField       =   "colorex"
         DataSource      =   "data1"
         Height          =   285
         Left            =   0
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
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
         Left            =   0
         TabIndex        =   67
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
         TabIndex        =   68
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
         Left            =   7080
         TabIndex        =   83
         Top             =   630
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         _Version        =   327681
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Quantitat Reb:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   3
         Left            =   5955
         TabIndex        =   84
         Top             =   675
         Width           =   1095
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
         Left            =   5730
         TabIndex        =   80
         Top             =   345
         Width           =   3450
      End
      Begin VB.Label nomsold 
         Caption         =   "NomEntrega"
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
         TabIndex        =   69
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
         TabIndex        =   65
         Top             =   2055
         Width           =   3075
      End
      Begin VB.Label Label1 
         Caption         =   "Un. Paquet:"
         DataSource      =   "data1"
         Height          =   270
         Index           =   106
         Left            =   5940
         TabIndex        =   64
         Top             =   2820
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Un. Caixa"
         DataSource      =   "data1"
         Height          =   270
         Index           =   126
         Left            =   6990
         TabIndex        =   63
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
         TabIndex        =   60
         Top             =   2805
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "Entrega:"
         DataSource      =   "data1"
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   115
         Left            =   105
         TabIndex        =   58
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
         TabIndex        =   57
         Top             =   2025
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Plegat:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   114
         Left            =   2055
         TabIndex        =   56
         Top             =   2310
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Ample:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   117
         Left            =   1245
         TabIndex        =   55
         Top             =   2295
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "B/L/F/BB:"
         DataSource      =   "data1"
         Height          =   270
         Index           =   118
         Left            =   180
         TabIndex        =   54
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Longitud:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   119
         Left            =   2775
         TabIndex        =   53
         Top             =   2310
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Solapa:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   120
         Left            =   3690
         TabIndex        =   52
         Top             =   2310
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Fuelle Ba:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   121
         Left            =   4395
         TabIndex        =   51
         Top             =   2310
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Fuelle Bo:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   122
         Left            =   5220
         TabIndex        =   50
         Top             =   2310
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Espessor:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   123
         Left            =   6045
         TabIndex        =   49
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
         TabIndex        =   48
         Top             =   2295
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "Kg/Mtrs/Diam:"
         DataSource      =   "data1"
         Height          =   270
         Index           =   4
         Left            =   7680
         TabIndex        =   35
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Sim:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   89
         Left            =   6300
         TabIndex        =   34
         Top             =   2010
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Ample Reb:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   88
         Left            =   4545
         TabIndex        =   33
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
         TabIndex        =   31
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
         TabIndex        =   28
         Top             =   1725
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Simult.Lam:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   73
         Left            =   4245
         TabIndex        =   26
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Texte:"
         DataSource      =   "data1"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   17
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
         MouseIcon       =   "Entrega.frx":214B
         MousePointer    =   99  'Custom
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   120
         Width           =   765
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         Left            =   1725
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   1050
         Width           =   1380
      End
      Begin VB.Label Label2 
         Caption         =   "Data Comanda:"
         Height          =   165
         Left            =   3210
         TabIndex        =   3
         Top             =   675
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Comanda:"
         Height          =   165
         Index           =   0
         Left            =   225
         TabIndex        =   2
         Top             =   660
         Width           =   1095
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Prem F2 per sel.leccionar Taules..."
      Height          =   225
      Left            =   90
      TabIndex        =   18
      Top             =   7365
      Width           =   9120
   End
End
Attribute VB_Name = "Entrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub comodi_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 39 Then KeyCode = 0: DBGrid1.SetFocus: SendKeys "{RIGHT}"
  If KeyCode = 37 Then KeyCode = 0: DBGrid1.SetFocus: SendKeys "{LEFT}"
End Sub


Private Sub cnumalb_LostFocus()
  If Len(cnumalb) > 0 And Len(cnumalb) < 12 Then
     If MsgBox("Aquest numero es mes petit de lo habitual" + Chr(10) + "ES CORRECTE?", vbInformation + vbYesNo, "Atenció") = vbNo Then cnumalb.SetFocus
  End If
End Sub

Private Sub comandanoacabada_Click()
  comandanoacabada.Tag = "1"
  Unload Me
End Sub

Private Sub Command1_Click()
If bobines.Recordset.EOF Then MsgBox "Encara no s'han fet les bobines de la última secció": Exit Sub
 bobines.Recordset.MoveFirst
  While Not bobines.Recordset.EOF
    If atrim(bobines.Recordset!data) = "" And atrim(bobines.Recordset!entregat) = "S" Then
      bobines.Recordset.Edit
      bobines.Recordset!data = dataentrega
      bobines.Recordset!Transportista = Transportista.ItemData(Transportista.ListIndex)
     ' bobines.Recordset!numalbara = atrim(cnumalb)
      bobines.Recordset.Update
    End If
    bobines.Recordset.MoveNext
  Wend
  bobines.Recordset.MoveFirst
  sumar_totals
End Sub
Function totentregat()
  Dim tot As Double
  tot = True
  bobines.Refresh
  If bobines.Recordset.EOF Then tot = 0
  While Not bobines.Recordset.EOF
     If atrim(bobines.Recordset!data) <> "" Then tot = tot + 1
     bobines.Recordset.MoveNext
  Wend
  totentregat = tot
End Function
Private Sub Command3_Click()
  actualitzar_bobinesent cadbl(entradabaixes.comanda), atrim(ruta.Text)
  bobines.Refresh
  'Clipboard.Clear
  'Clipboard.SetText bobines.RecordSource
End Sub



Private Sub Command4_Click()
  dataentrega = Format(DateAdd("d", 1, dataentrega), "dd/mm/yy")
End Sub

Private Sub Command5_Click()
If bobines.Recordset.EOF Then MsgBox "Encara no s'han fet les bobines de la última secció": Exit Sub
  bobines.Recordset.MoveFirst
  While Not bobines.Recordset.EOF
    If atrim(bobines.Recordset!data) = "" Then
      bobines.Recordset.Edit
      If atrim(bobines.Recordset!entregat) = "S" Then
         bobines.Recordset!entregat = "N"
          Else: bobines.Recordset!entregat = "S"
      End If
      bobines.Recordset.Update
    End If
    bobines.Recordset.MoveNext
  Wend
  bobines.Recordset.MoveFirst
End Sub

Private Sub Command6_Click()
  dataentrega = Format(DateAdd("d", -1, dataentrega), "dd/mm/yy")
End Sub

Private Sub Command7_Click()
  Bobinesassignades.Show 1
  Unload Bobinesassignades
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
 
  
  'carrega el nom de la Entrega
  lookupde "select descripcio from maquines where maquina='s' and codi=" + atrim(cadbl(data1.Recordset!soldadora)), , nomsold
  
 'lookup de tipussoldadura
  Set rsttmp = dbtmp.OpenRecordset("select descripcio from tipussoldadura where codi='" + atrim((data1.Recordset!tipusoldadura)) + "'")
  If Not rsttmp.EOF Then
     tipussoldadura = atrim(rsttmp!descripcio)
    Else: tipussoldadura = ""
  End If
End Sub



Private Sub dataentrega_LostFocus()
  Static contad As Byte
  valtmp = dataentrega
  If InStr(1, valtmp, "/") = 0 Then valtmp = Mid(valtmp, 1, 2) + "/" + Mid(valtmp, 3, 2) + "/" + Mid(valtmp, 5, 2)
  If Not IsDate(valtmp) Then
        If cadbl(contad) < 1 Then
           valtmp = "": contad = cadbl(contad) + 1: dataentrega.SetFocus
           MsgBox "Error data equivocada", vbCritical, "Reintenta-ho"
         Else: contad = 0
        End If
         Else: contad = 0
  End If
  dataentrega = valtmp
End Sub

Private Sub DBGrid1_BeforeUpdate(Cancel As Integer)
'bobines.Recordset!datahorainici = DBGrid1.Columns(1).Text + " " + DBGrid1.Columns(2).Text
End Sub

Private Sub DBGrid1_Change()
'  Me.Caption = DBGrid1.Text
End Sub

Private Sub DBGrid1_GotFocus()
  DBGrid1_RowColChange 1, 1
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    KeyCode = 0
    SendKeys "{TAB}"
  End If
 'If KeyCode = 113 And DBGrid1.col = 0 Then
 '  triarEntrega
 'End If
 'If KeyCode = 113 And DBGrid1.col = 1 Then
 '  triaroperaris
 'End If
 
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
'  If DBGrid1.col = 6 Then
'    If InStr(1, "CAFP", UCase$(Chr$(KeyAscii))) = 0 Then
'       KeyAscii = 0
'      Else: KeyAscii = Asc(UCase$(Chr$(KeyAscii))): DBGrid1.Text = ""
'    End If
'  End If
'
End Sub

Private Sub DBGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If X > DBGrid1.Columns("Entregat").Left And X < (DBGrid1.Columns("Entregat").Left + DBGrid1.Columns("Entregat").Width) Then
      Screen.MousePointer = 10
    Else: Screen.MousePointer = 0
  End If
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 If DBGrid1.Columns(DBGrid1.col).DataField = "entregat" Then
   If DBGrid1.Columns("Data Alb.") = "" Or retornmaterial.Value = 1 Then
    If DBGrid1.Columns("Entr.") = "S" Then
        DBGrid1.Text = "N"
        DBGrid1.Columns("Alb.") = Null
        If DBGrid1.Columns("Data Alb.") <> "" Then DBGrid1.Columns("Data Alb.") = Null
      Else: DBGrid1.Text = "S"
    End If
   End If
   SendKeys ("{Right}")
  ' DBGrid1.Columns(2).SelStart = 0
  End If
  If DBGrid1.Columns("Data Alb.") = "" Then DBGrid1.Columns("Data Alb.") = Null
 sumar_totals
End Sub
Sub sumar_totals()
  Dim ventregatm As Double
  Dim vpendentm As Double
  Dim ventregatk As Double
  Dim vpendentk As Double
  Dim rsttmpt As Recordset
  Set rsttmpt = dbtmpb.OpenRecordset("select metresisacs,data,kilosiunitats from bobinesent where comanda=" + atrim(cadbl(entradabaixes.comanda.Text)))
 ' rsttmpt.MoveLast
  While Not rsttmpt.EOF
    
     If rsttmpt!data <> "" Then
       ventregatm = ventregatm + cadbl(rsttmpt!metresisacs)
       ventregatk = ventregatk + cadbl(rsttmpt!kilosiunitats)
      Else:
         vpendentm = vpendentm + cadbl(rsttmpt!metresisacs)
         vpendentk = vpendentk + cadbl(rsttmpt!kilosiunitats)
    End If
    rsttmpt.MoveNext
  Wend

  entregatm = Format(ventregatm, "#,##0.00")
  pendentm = Format(vpendentm, "#,##0.00")
  entregatk = Format(ventregatk, "#,##0.00")
  pendentk = Format(vpendentk, "#,##0.00")

  Set rsttmpt = Nothing
End Sub
Private Sub Form_Load()
centerscreen Me
data1.DatabaseName = camicomandes
data1.RecordSource = "select * from comandes where comanda=" + atrim(cadbl(entradabaixes.comanda.Text))
bobines.DatabaseName = cami
bobines.RecordSource = "SELECT transportistes.descripcio AS NomTransport, * FROM bobinesent LEFT JOIN transportistes ON bobinesent.transportista = transportistes.codi WHERE comanda=" + atrim(cadbl(entradabaixes.comanda.Text)) + " order by numbob asc"
'Clipboard.Clear
'Clipboard.SetText bobines.RecordSource
Set dbtmp = OpenDatabase(data1.DatabaseName)
Set dbtmpb = OpenDatabase(bobines.DatabaseName)
data1.Refresh
bobines.Refresh
dataentrega = Format(Now, "dd/mm/yy")
'DBGrid1.Columns(2).ValueItems.Presentation = 4

'coloco el combo de transportistes
Set rsttmp = dbtmp.OpenRecordset("select * from transportistes order by codi")
While Not rsttmp.EOF
 Transportista.AddItem atrim(rsttmp!codi) + " --> " + atrim(rsttmp!descripcio)
 Transportista.ItemData(Transportista.NewIndex) = rsttmp!codi
 rsttmp.MoveNext
Wend
Command3_Click
On Error Resume Next
Transportista.ListIndex = 0
sumar_totals
Entrega.Tag = pendentm
Set dbstocks = OpenDatabase(rutadelfitxer(camicomandes) + "palets.mdb")
comprovarlaseccioenruta
End Sub
Sub comprovarlaseccioenruta()
  Dim posicioruta As String
    Set rsttmp = dbtmp.OpenRecordset("SELECT comandes.proximaseccio, productes.ruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (comandes.comanda)=" + atrim(entradabaixes.comanda))
    posicioruta = posicioenlaruta(cadbl(entradabaixes.comanda), rsttmp!proximaseccio, rsttmp!ruta)
    If posicioruta <> "" Then
        lseccioenruta = "Seccio no acabada: " + posicioruta
      Else: lseccioenruta = ""
    End If
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Screen.MousePointer = 10 Then Screen.MousePointer = 0
End Sub
Function sihihanbobinesassignades(numc As Double) As Boolean
   Dim rstb As Recordset
   Dim bobines As String
   Set rstb = dbstocks.OpenRecordset("select * from parcials where not utilitzada and comanda='" + atrim(numc) + "'")
   bobines = ""
   While Not rstb.EOF
     bobines = bobines + "[" + atrim(rstb!idpalet) + "/" + atrim(rstb!idbobina) + "] "
     rstb.MoveNext
   Wend
   If atrim(bobines) <> "" Then
      MsgBox "No es pot passar la comanda a acabada fins que les següents bobines es passin a disponibles o gastades." + Chr(10) + Chr(13) + bobines, vbInformation, "Atenció"
      sihihanbobinesassignades = True
     Else: sihihanbobinesassignades = False
   End If
End Function

Function posicioenlaruta(numc As Double, seccioactual As String, laruta As String) As String
  Dim rstp As Recordset
  If InStr(1, "VPT", seccioactual) = 0 Then Exit Function
  Set rstp = dbtmpb.OpenRecordset("SELECT comandes.comanda, rebobinadorestot.acavada as acavadar, laminadorestot.acavada as acavadal, impressorestot.acavada as acavadai FROM ((comandes LEFT JOIN rebobinadorestot ON comandes.comanda = rebobinadorestot.comanda) LEFT JOIN laminadorestot ON comandes.comanda = laminadorestot.comanda) LEFT JOIN impressorestot ON comandes.comanda = impressorestot.comanda WHERE (((comandes.comanda)=" + atrim(numc) + "));")
  
  If Not rstp.EOF Then
     If InStr(1, laruta, "R") > 0 And cadbl(rstp!acavadar) = 0 Then posicioenlaruta = "R"
     If InStr(1, laruta, "L") > 0 And cadbl(rstp!acavadal) = 0 Then posicioenlaruta = "L"
     If InStr(1, laruta, "I") > 0 And cadbl(rstp!acavadai) = 0 Then posicioenlaruta = "I"
  End If
  
  Set rstp = Nothing
End Function
Sub passarlaseccioaentregada()
   Dim resp As Long
   If Entrega.Tag = pendentm Then Exit Sub
   resp = MsgBox("Està tot entregat d'aquesta comanda? (T)" + Chr(10) + " si poses No serà Parcial (P)", vbInformation + vbYesNo + vbDefaultButton1, "Entrega Total o Parcial")
   If resp = vbYes Then
    dbtmp.Execute "update comandes set proximaseccio='T' where comanda=" + atrim(entradabaixes.comanda)
    Set rsttmp = dbtmp.OpenRecordset("select linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(entradabaixes.comanda))
    If cadbl(rsttmp!linkcomanda1) > 0 Then dbtmp.Execute "update comandes set proximaseccio='T' where comanda=" + atrim(rsttmp!linkcomanda1)
    If cadbl(rsttmp!linkcomanda2) > 0 Then dbtmp.Execute "update comandes set proximaseccio='T' where comanda=" + atrim(rsttmp!linkcomanda2)
    Set rsttmp = Nothing
   End If
   If resp = vbNo Then dbtmp.Execute "update comandes set proximaseccio='P' where comanda=" + atrim(entradabaixes.comanda)
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Dim posicioruta As String
'If Me.Name = Screen.ActiveForm.Name Then controlar_fiseccio "T", ruta
If Me.Name = Screen.ActiveForm.Name Then
 mirarlesquenotenendata
 actualitzatot_entrega
 i = totentregat
 If i = bobines.Recordset.RecordCount - 1 Then
    mirarlesquenotenendata
    If sihihanbobinesassignades(cadbl(entradabaixes.comanda)) Then
      If MsgBox("Vols surtir d'aquesta baixa?", vbCritical + vbYesNo, "Atenció") = vbNo Then
          Cancel = 1
      End If
      GoTo fi
    End If
    If comandanoacabada.Tag = "" Then
     Set rsttmp = dbtmp.OpenRecordset("SELECT comandes.proximaseccio, productes.ruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (comandes.comanda)=" + atrim(entradabaixes.comanda))
     posicioruta = posicioenlaruta(cadbl(entradabaixes.comanda), rsttmp!proximaseccio, rsttmp!ruta)
     If posicioruta = "" Then
        passarlaseccioaentregada
       Else
         MsgBox "Aquesta comanda no està acabada en la seccio de " + posicioruta + Chr(10) + "SI ES UN ERROR PASSA LA SECCIO A ACABADA I TORNA A LA BAIXA D'ENTREGA."
         If i > 0 Then dbtmp.Execute "update comandes set proximaseccio='P' where comanda=" + atrim(entradabaixes.comanda)
     End If
    End If
   ' modificar_estat_comanda entradabaixes.comanda, ruta, "T", 1
      Else:
         If cadbl(i) > 0 Then
             dbtmp.Execute "update comandes set proximaseccio='P' where comanda=" + atrim(entradabaixes.comanda)
           Else: If comandanoacabada.Tag = "" Then modificar_estat_comanda entradabaixes.comanda, ruta, "V", 0
         End If
 End If
End If
entradabaixes.Visible = True
fi:

'Set dbstocks = Nothing
If Not canvissortirseccio Then End

End Sub
Sub mirarlesquenotenendata()
    Dim avis As Boolean
    bobines.Refresh
    While Not bobines.Recordset.EOF
       If bobines.Recordset!entregat = "S" And Not IsDate(bobines.Recordset!data) Then
          If Not avis Then MsgBox "Hi han bobines marcades amb entregades sense data es possaran a no entregades": avis = True
          bobines.Recordset.Edit
          bobines.Recordset!entregat = "N"
          bobines.Recordset.Update
       End If
       bobines.Recordset.MoveNext
    Wend
End Sub

Sub actualitzatot_entrega()
' bobines.Recordset.MoveLast
' While atrim(bobines.Recordset!tipus) <> "F" And Not bobines.Recordset.EOF
'    bobines.Recordset.MovePrevious
' Wend
'  Set rsttmp = dbtmpb.OpenRecordset("select * from impressorestot where comanda=" + atrim(cadbl(entradabaixes.comanda)))
'  If rsttmp.EOF Then
'      rsttmp.AddNew
'    Else: rsttmp.Edit
'  End If
'  With rsttmp
 '   !comanda = cadbl(entradabaixes.comanda)
 '   !hclixe = cadbl(hclixe)
 '   !hmaquina = cadbl(hmaquina)
 '   !hajusts = cadbl(hajusts)
 '   !hfuncio = cadbl(hfunc)
 '   !tbobines = cadbl(tbob)
 '   !tprova = cadbl(tprova)
 '   !tkilos = cadbl(tkilos)
 '   !tmetres = cadbl(tmetres)
 '   !metresmin = cadbl(kiloshora)
 '   !kilostinta = cadbl(bobines.Recordset!kgtinta)
 '   If Not IsNull(bobines.Recordset!datafi) Then !dataimpressio = bobines.Recordset!datafi
 '   !impressora = cadbl(bobines.Recordset!numeromaquina)
 '   !operari = cadbl(bobines.Recordset!operari)
 '  .Update
 ' End With

End Sub


Private Sub retornmaterial_Click()
   If retornmaterial.Value = 1 Then
      MsgBox "Clica sobre de la S d'entregat de cada bobina per escullir quines bobines es retornen", vbInformation, "Atenció"
   End If
End Sub

