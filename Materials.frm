VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fmaterials 
   Caption         =   "Manteniment de Materials (Productes)"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9780
   Icon            =   "Materials.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9780
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameestoc 
      BackColor       =   &H00EEE4D7&
      Caption         =   "Control d'estoc"
      Height          =   3450
      Left            =   7965
      TabIndex        =   142
      Top             =   4290
      Width           =   9585
      Begin VB.CommandButton Command30 
         Height          =   360
         Left            =   705
         Picture         =   "Materials.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   147
         ToolTipText     =   "Eliminacio Registres"
         Top             =   315
         Width           =   420
      End
      Begin VB.CommandButton Command29 
         Height          =   360
         Left            =   270
         Picture         =   "Materials.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   146
         ToolTipText     =   "Alta  Registres"
         Top             =   315
         Width           =   420
      End
      Begin VB.CommandButton Command28 
         Height          =   390
         Left            =   9135
         Picture         =   "Materials.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   145
         ToolTipText     =   "Tornar"
         Top             =   165
         Width           =   390
      End
      Begin VB.Data Dataestoc 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   315
         Left            =   4365
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "materials_estoc"
         Top             =   300
         Visible         =   0   'False
         Width           =   1290
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "Materials.frx":1628
         Height          =   2325
         Left            =   255
         OleObjectBlob   =   "Materials.frx":163C
         TabIndex        =   143
         Top             =   720
         Width           =   9165
      End
      Begin VB.Label ettotalestoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2745
         TabIndex        =   144
         Top             =   225
         Width           =   6180
      End
   End
   Begin VB.CommandButton Command23 
      Height          =   300
      Left            =   555
      Picture         =   "Materials.frx":238F
      Style           =   1  'Graphical
      TabIndex        =   127
      TabStop         =   0   'False
      ToolTipText     =   "Filtre per totes les families"
      Top             =   2310
      Width           =   330
   End
   Begin VB.CommandButton Command22 
      Height          =   300
      Left            =   1275
      Picture         =   "Materials.frx":2919
      Style           =   1  'Graphical
      TabIndex        =   123
      TabStop         =   0   'False
      ToolTipText     =   "Filtre per tots els materials d'aquest proveidor."
      Top             =   1245
      Width           =   330
   End
   Begin VB.Frame framefamilies 
      BackColor       =   &H00F3B378&
      Caption         =   "                        Escandall de preus"
      Height          =   855
      Left            =   2940
      TabIndex        =   101
      Top             =   4890
      Visible         =   0   'False
      Width           =   8310
      Begin VB.CommandButton Command13 
         Height          =   285
         Left            =   7815
         Picture         =   "Materials.frx":2EA3
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Eliminacio Registres"
         Top             =   315
         Width           =   300
      End
      Begin VB.CommandButton Command12 
         Height          =   285
         Left            =   7515
         Picture         =   "Materials.frx":342D
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Alta  Registres"
         Top             =   315
         Width           =   300
      End
      Begin VB.ComboBox combofamilies 
         Height          =   315
         Left            =   1035
         TabIndex        =   102
         Top             =   300
         Width           =   6435
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Families:"
         Height          =   180
         Index           =   5
         Left            =   105
         TabIndex        =   105
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.Frame Frameescandall 
      BackColor       =   &H00F1B75F&
      Caption         =   "                               Escandall de preus"
      Height          =   5085
      Left            =   2940
      TabIndex        =   73
      Top             =   5730
      Visible         =   0   'False
      Width           =   8310
      Begin VB.ComboBox combomesura 
         Height          =   315
         ItemData        =   "Materials.frx":39B7
         Left            =   5400
         List            =   "Materials.frx":39C1
         TabIndex        =   111
         Text            =   "Kg"
         Top             =   1155
         Width           =   750
      End
      Begin VB.OptionButton cmicresoamplada 
         BackColor       =   &H00F1B75F&
         Caption         =   "Amplada"
         Height          =   195
         Index           =   1
         Left            =   1095
         TabIndex        =   110
         Top             =   1005
         Width           =   1050
      End
      Begin VB.OptionButton cmicresoamplada 
         BackColor       =   &H00F1B75F&
         Caption         =   "Micres"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   109
         Top             =   1005
         Width           =   855
      End
      Begin VB.CheckBox Checkactiu 
         BackColor       =   &H00F1B75F&
         Caption         =   "Actiu per pressupostos."
         Height          =   240
         Left            =   3195
         TabIndex        =   106
         Top             =   285
         Width           =   2370
      End
      Begin VB.Frame frameseleccio 
         BackColor       =   &H0000FFFF&
         Caption         =   "Escullir material"
         Height          =   2265
         Left            =   1215
         TabIndex        =   89
         Top             =   1935
         Visible         =   0   'False
         Width           =   5925
         Begin VB.ComboBox famad 
            Height          =   315
            Left            =   705
            TabIndex        =   97
            Top             =   1170
            Width           =   2580
         End
         Begin VB.ComboBox subfamad 
            Height          =   315
            Left            =   3315
            TabIndex        =   96
            Tag             =   "famad"
            Top             =   1170
            Width           =   2490
         End
         Begin VB.ComboBox famcol 
            Height          =   315
            Left            =   705
            TabIndex        =   95
            Top             =   840
            Width           =   2580
         End
         Begin VB.ComboBox subfamcol 
            Height          =   315
            Left            =   3315
            TabIndex        =   94
            Tag             =   "famcol"
            Top             =   840
            Width           =   2490
         End
         Begin VB.ComboBox subfammat 
            Height          =   315
            Left            =   3315
            TabIndex        =   93
            Tag             =   "fammat"
            Top             =   510
            Width           =   2490
         End
         Begin VB.ComboBox fammat 
            Height          =   315
            Left            =   705
            TabIndex        =   92
            Top             =   510
            Width           =   2580
         End
         Begin VB.CommandButton Command18 
            Height          =   360
            Left            =   4035
            Picture         =   "Materials.frx":39CE
            Style           =   1  'Graphical
            TabIndex        =   91
            ToolTipText     =   "Acceptar canvis"
            Top             =   1635
            Width           =   840
         End
         Begin VB.CommandButton Command17 
            Height          =   360
            Left            =   4890
            Picture         =   "Materials.frx":3F58
            Style           =   1  'Graphical
            TabIndex        =   90
            ToolTipText     =   "Cancelar"
            Top             =   1635
            Width           =   840
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Mat:         Col:          Ad:"
            Height          =   1020
            Left            =   390
            TabIndex        =   100
            Top             =   450
            Width           =   345
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Subfamilia "
            Height          =   285
            Index           =   3
            Left            =   3780
            TabIndex        =   99
            Top             =   285
            Width           =   2115
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Familia "
            Height          =   285
            Left            =   2205
            TabIndex        =   98
            Top             =   300
            Width           =   750
         End
      End
      Begin VB.CommandButton Command16 
         Height          =   285
         Left            =   3060
         Picture         =   "Materials.frx":44E2
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Eliminacio Registres"
         Top             =   1200
         Width           =   300
      End
      Begin VB.CommandButton Command15 
         Height          =   285
         Left            =   2760
         Picture         =   "Materials.frx":4A6C
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Alta  Registres"
         Top             =   1200
         Width           =   300
      End
      Begin VB.Data datapreukg 
         Caption         =   "datapreukg"
         Connect         =   "Access"
         DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2685
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tarifesproveidorsescandall"
         Top             =   4275
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.TextBox cnumoferta 
         Height          =   300
         Left            =   6810
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   630
         Width           =   1425
      End
      Begin VB.ListBox llistamicres 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00ED823A&
         Height          =   2700
         Left            =   90
         TabIndex        =   81
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox dataoferta 
         Height          =   315
         Left            =   1155
         TabIndex        =   80
         Top             =   630
         Width           =   1425
      End
      Begin VB.TextBox cvigencia 
         Height          =   300
         Left            =   4500
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   615
         Width           =   1080
      End
      Begin VB.CommandButton Command7 
         Height          =   285
         Left            =   2055
         Picture         =   "Materials.frx":4FF6
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Alta  Registres"
         Top             =   1215
         Width           =   300
      End
      Begin VB.CommandButton Command8 
         Height          =   285
         Left            =   2055
         Picture         =   "Materials.frx":5580
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Eliminacio Registres"
         Top             =   1500
         Width           =   300
      End
      Begin VB.CommandButton Command10 
         Height          =   285
         Left            =   2580
         Picture         =   "Materials.frx":5B0A
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Alta  Registres"
         Top             =   645
         Width           =   300
      End
      Begin VB.CommandButton Command11 
         Height          =   285
         Left            =   2880
         Picture         =   "Materials.frx":6094
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Eliminacio Registres"
         Top             =   645
         Width           =   300
      End
      Begin MSDBGrid.DBGrid reixapreukg 
         Bindings        =   "Materials.frx":661E
         Height          =   2790
         Left            =   2745
         OleObjectBlob   =   "Materials.frx":6633
         TabIndex        =   78
         Top             =   1485
         Width           =   2865
      End
      Begin VB.TextBox ckgminim 
         Height          =   300
         Left            =   4605
         Locked          =   -1  'True
         TabIndex        =   107
         Top             =   1170
         Width           =   765
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Compra mínim:"
         Height          =   225
         Index           =   3
         Left            =   3540
         TabIndex        =   108
         Top             =   1245
         Width           =   1245
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Oferta:"
         Height          =   225
         Index           =   4
         Left            =   6105
         TabIndex        =   85
         Top             =   690
         Width           =   885
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Data oferta:"
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   83
         Top             =   630
         Width           =   885
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Vigencia:"
         Height          =   225
         Index           =   2
         Left            =   3795
         TabIndex        =   82
         Top             =   675
         Width           =   885
      End
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "Materials.frx":7026
      Height          =   2610
      Left            =   30
      OleObjectBlob   =   "Materials.frx":703A
      TabIndex        =   42
      Top             =   4050
      Width           =   4740
   End
   Begin VB.Frame framematerials 
      Caption         =   "Dades dels Materials"
      Enabled         =   0   'False
      Height          =   3435
      Left            =   60
      TabIndex        =   17
      Top             =   585
      Width           =   9585
      Begin VB.CommandButton bestoc 
         BackColor       =   &H00F3B378&
         Caption         =   "Stock"
         Height          =   300
         Left            =   8340
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   615
         Width           =   1125
      End
      Begin VB.Frame framesubstancies 
         BackColor       =   &H006BEBB1&
         Caption         =   "Substancies"
         Height          =   3150
         Left            =   4185
         TabIndex        =   116
         Top             =   3090
         Visible         =   0   'False
         Width           =   5700
         Begin VB.CommandButton Command26 
            Height          =   360
            Left            =   4185
            Picture         =   "Materials.frx":8AE7
            Style           =   1  'Graphical
            TabIndex        =   140
            ToolTipText     =   "Modificar data anàlisis"
            Top             =   195
            Width           =   420
         End
         Begin VB.TextBox cdataanalisis 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2805
            Locked          =   -1  'True
            TabIndex        =   138
            Top             =   255
            Width           =   1365
         End
         Begin VB.CommandButton Command25 
            Height          =   360
            Left            =   5100
            Picture         =   "Materials.frx":9071
            Style           =   1  'Graphical
            TabIndex        =   137
            ToolTipText     =   "Copiar substancies a un altra Material."
            Top             =   165
            Width           =   420
         End
         Begin VB.CommandButton Command21 
            Height          =   360
            Left            =   60
            Picture         =   "Materials.frx":95FB
            Style           =   1  'Graphical
            TabIndex        =   119
            ToolTipText     =   "Alta  Registres"
            Top             =   210
            Width           =   570
         End
         Begin VB.CommandButton Command20 
            Height          =   360
            Left            =   645
            Picture         =   "Materials.frx":9B85
            Style           =   1  'Graphical
            TabIndex        =   118
            ToolTipText     =   "Eliminacio Registres"
            Top             =   210
            Width           =   420
         End
         Begin VB.Data datasubstancies 
            Caption         =   "datasubstancies"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   4560
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   750
            Visible         =   0   'False
            Width           =   1140
         End
         Begin MSDBGrid.DBGrid reixasubstancies 
            Bindings        =   "Materials.frx":A10F
            Height          =   2385
            Left            =   135
            OleObjectBlob   =   "Materials.frx":A129
            TabIndex        =   117
            Top             =   615
            Width           =   5325
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Data anàlisis:"
            Height          =   285
            Left            =   1770
            TabIndex        =   139
            Top             =   285
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FDDECE&
         Caption         =   "Tractat cares"
         Height          =   930
         Left            =   5580
         TabIndex        =   132
         Top             =   2460
         Width           =   3735
         Begin VB.ComboBox Combocara2 
            Height          =   315
            Left            =   555
            TabIndex        =   136
            Top             =   570
            Width           =   3135
         End
         Begin VB.ComboBox Combocara1 
            Height          =   315
            Left            =   555
            TabIndex        =   135
            Top             =   195
            Width           =   3135
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Cara2:"
            Height          =   180
            Left            =   45
            TabIndex        =   134
            Top             =   615
            Width           =   495
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Cara1:"
            Height          =   180
            Left            =   45
            TabIndex        =   133
            Top             =   255
            Width           =   495
         End
      End
      Begin VB.CommandButton Command24 
         BackColor       =   &H0080FFFF&
         Caption         =   "Detall descripció"
         Height          =   450
         Left            =   6510
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   180
         Width           =   870
      End
      Begin VB.ComboBox Combocolorrec 
         DataField       =   "colorreciclatge"
         DataSource      =   "materials"
         Height          =   315
         ItemData        =   "Materials.frx":ACC1
         Left            =   8460
         List            =   "Materials.frx":ACCE
         TabIndex        =   129
         Top             =   1245
         Width           =   1080
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "tanpercentimpostenvasos"
         DataSource      =   "materials"
         Height          =   315
         Index           =   14
         Left            =   8460
         TabIndex        =   125
         Top             =   930
         Width           =   450
      End
      Begin VB.ComboBox CombotipusCQ 
         Height          =   315
         ItemData        =   "Materials.frx":ACE7
         Left            =   6555
         List            =   "Materials.frx":ACF4
         TabIndex        =   120
         Top             =   2070
         Width           =   1875
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H006BEBB1&
         Caption         =   "Substancies"
         Height          =   450
         Left            =   8340
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   150
         Width           =   1140
      End
      Begin VB.TextBox cmatcompatible 
         DataField       =   "subfamiliacompatible"
         DataSource      =   "materials"
         Height          =   285
         Left            =   555
         TabIndex        =   114
         Top             =   3060
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox combosubfamcompatible 
         BackColor       =   &H00FFC0C0&
         DataSource      =   "materials"
         Height          =   315
         ItemData        =   "Materials.frx":AD30
         Left            =   2100
         List            =   "Materials.frx":AD32
         TabIndex        =   112
         Top             =   3015
         Width           =   2835
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00F3B378&
         Caption         =   "€ Tarifes"
         Height          =   390
         Left            =   7530
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   630
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox Text1 
         DataField       =   "descripciopressupost"
         DataSource      =   "materials"
         Height          =   315
         Index           =   13
         Left            =   825
         MaxLength       =   40
         TabIndex        =   70
         Top             =   990
         Width           =   4110
      End
      Begin VB.Frame frameFT 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Valors Fitxa Tècnica"
         Height          =   2040
         Left            =   8160
         TabIndex        =   54
         Top             =   2475
         Visible         =   0   'False
         Width           =   3390
         Begin VB.CommandButton Command5 
            Height          =   285
            Left            =   3045
            Picture         =   "Materials.frx":AD34
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Eliminar el PDF vinculat i la data"
            Top             =   705
            Width           =   300
         End
         Begin VB.CommandButton botopdf 
            DisabledPicture =   "Materials.frx":B2BE
            DownPicture     =   "Materials.frx":C8A8
            Height          =   690
            Left            =   2235
            OLEDropMode     =   1  'Manual
            Picture         =   "Materials.frx":DE92
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   315
            Width           =   780
         End
         Begin VB.TextBox Text7 
            DataField       =   "Ftvigencia"
            DataSource      =   "materials"
            Height          =   285
            Left            =   1650
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   65
            Top             =   1140
            Width           =   1230
         End
         Begin VB.TextBox Text6 
            DataField       =   "FTcodi"
            DataSource      =   "materials"
            Height          =   285
            Left            =   1110
            MaxLength       =   20
            TabIndex        =   63
            Top             =   1590
            Width           =   2160
         End
         Begin VB.TextBox Text5 
            DataField       =   "FTgrmsm2_max"
            DataSource      =   "materials"
            Height          =   285
            Left            =   1125
            TabIndex        =   61
            Top             =   735
            Width           =   345
         End
         Begin VB.TextBox Text4 
            DataField       =   "FTgrmsm2_min"
            DataSource      =   "materials"
            Height          =   285
            Left            =   600
            TabIndex        =   59
            Top             =   735
            Width           =   345
         End
         Begin VB.TextBox Text3 
            DataField       =   "FTmicres_max"
            DataSource      =   "materials"
            Height          =   285
            Left            =   1125
            TabIndex        =   58
            Top             =   435
            Width           =   345
         End
         Begin VB.TextBox Text2 
            DataField       =   "FTmicres_min"
            DataSource      =   "materials"
            Height          =   285
            Left            =   600
            TabIndex        =   55
            Top             =   435
            Width           =   345
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Arrastrar PDF FT"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   390
            Left            =   2100
            TabIndex        =   67
            Top             =   90
            Width           =   1245
         End
         Begin VB.Label Label12 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Vigència FT:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   64
            Top             =   1170
            Width           =   1395
         End
         Begin VB.Label Label11 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Codi FT:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   45
            TabIndex        =   62
            Top             =   1605
            Width           =   1095
         End
         Begin VB.Label Label10 
            BackColor       =   &H00EAD9CE&
            Caption         =   "g/m2       -"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   60
            TabIndex        =   60
            Top             =   750
            Width           =   1860
         End
         Begin VB.Label Label9 
            BackColor       =   &H00EAD9CE&
            Caption         =   "%Min-%Max"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   555
            TabIndex        =   57
            Top             =   225
            Width           =   1110
         End
         Begin VB.Label Label8 
            BackColor       =   &H00EAD9CE&
            Caption         =   "µm           -"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   45
            TabIndex        =   56
            Top             =   450
            Width           =   1710
         End
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Fitxes Tècniques"
         Height          =   450
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   165
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "micresdelsgrm2"
         DataSource      =   "materials"
         Height          =   315
         Index           =   12
         Left            =   5715
         TabIndex        =   50
         ToolTipText     =   "Nomes possar si són grms/m2 per saber les micres teòriques."
         Top             =   630
         Width           =   675
      End
      Begin VB.CheckBox material2cares 
         Caption         =   "Tractat 2 Cares"
         DataField       =   "material2cares"
         DataSource      =   "materials"
         Height          =   195
         Left            =   5730
         TabIndex        =   49
         Top             =   1815
         Width           =   1575
      End
      Begin VB.ComboBox mesespcompra 
         BackColor       =   &H00FFC0C0&
         DataField       =   "mesuarespcompra"
         DataSource      =   "materials"
         Height          =   315
         ItemData        =   "Materials.frx":F47C
         Left            =   4290
         List            =   "Materials.frx":F492
         TabIndex        =   46
         Top             =   1560
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "grmm2"
         DataSource      =   "materials"
         Height          =   315
         Index           =   11
         Left            =   5715
         TabIndex        =   43
         Top             =   1380
         Width           =   675
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "grmcm3"
         DataSource      =   "materials"
         Height          =   315
         Index           =   10
         Left            =   5715
         TabIndex        =   4
         Top             =   1005
         Width           =   675
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "subfamiliaad"
         DataSource      =   "materials"
         Height          =   315
         Index           =   9
         Left            =   1275
         TabIndex        =   10
         Top             =   2610
         Width           =   390
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "subfamiliacol"
         DataSource      =   "materials"
         Height          =   315
         Index           =   8
         Left            =   1275
         TabIndex        =   8
         Top             =   2310
         Width           =   390
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "subfamilia"
         DataSource      =   "materials"
         Height          =   315
         Index           =   7
         Left            =   1275
         TabIndex        =   6
         Top             =   1995
         Width           =   390
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "familiaad"
         DataSource      =   "materials"
         Height          =   315
         Index           =   6
         Left            =   855
         TabIndex        =   9
         Top             =   2610
         Width           =   390
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "familiacol"
         DataSource      =   "materials"
         Height          =   315
         Index           =   5
         Left            =   855
         TabIndex        =   7
         Top             =   2310
         Width           =   390
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "familia"
         DataSource      =   "materials"
         Height          =   315
         Index           =   4
         Left            =   855
         TabIndex        =   5
         Top             =   1995
         Width           =   390
      End
      Begin VB.TextBox Text1 
         DataField       =   "refproducte"
         DataSource      =   "materials"
         Height          =   315
         Index           =   3
         Left            =   810
         TabIndex        =   3
         Top             =   1395
         Width           =   3300
      End
      Begin VB.TextBox Text1 
         DataField       =   "descripcio"
         DataSource      =   "materials"
         Height          =   315
         Index           =   2
         Left            =   1365
         TabIndex        =   1
         Top             =   270
         Width           =   5010
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "proveidor"
         DataSource      =   "materials"
         Height          =   315
         Index           =   1
         Left            =   840
         TabIndex        =   2
         Top             =   645
         Width           =   390
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         DataField       =   "codi"
         DataSource      =   "materials"
         Height          =   315
         Index           =   0
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   270
         Width           =   480
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Material delicat (Especial)"
         DataField       =   "materialdelicat"
         DataSource      =   "materials"
         Height          =   405
         Left            =   7320
         TabIndex        =   124
         Top             =   1710
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Color Reciclat:"
         Height          =   300
         Index           =   14
         Left            =   7335
         TabIndex        =   128
         Top             =   1305
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "% Impost Envasos:"
         Height          =   315
         Index           =   13
         Left            =   7065
         TabIndex        =   126
         Top             =   990
         Width           =   1815
      End
      Begin VB.Label etdataqualitat 
         BackStyle       =   0  'Transparent
         DataField       =   "dataCQ"
         DataSource      =   "materials"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   8460
         TabIndex        =   122
         Top             =   2100
         Width           =   1080
      End
      Begin VB.Label Label17 
         Caption         =   "Tipus Qualitat:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   121
         Top             =   2145
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Subfam Material Compatible Mat. Específic:"
         Height          =   390
         Index           =   12
         Left            =   90
         TabIndex        =   113
         Top             =   2925
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Desc Pressupost"
         Height          =   450
         Index           =   11
         Left            =   30
         TabIndex        =   71
         Top             =   900
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Espessor:"
         Height          =   300
         Index           =   10
         Left            =   4980
         TabIndex        =   52
         Top             =   660
         Width           =   690
      End
      Begin VB.Label Label7 
         Caption         =   "Micres                                         (Només materials amb grm2)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6435
         TabIndex        =   51
         Top             =   660
         Width           =   1890
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mesura per comprar:"
         Height          =   300
         Index           =   9
         Left            =   4215
         TabIndex        =   45
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Grm/m2                                           (Només material  xr metres)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6435
         TabIndex        =   44
         Top             =   1425
         Width           =   1890
      End
      Begin VB.Label Label5 
         Caption         =   "Grm/cm3"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6420
         TabIndex        =   41
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Densitat:"
         Height          =   300
         Index           =   3
         Left            =   4980
         TabIndex        =   40
         Top             =   1035
         Width           =   690
      End
      Begin VB.Label Label3 
         Height          =   285
         Index           =   2
         Left            =   1710
         TabIndex        =   32
         Top             =   2670
         Width           =   4080
      End
      Begin VB.Label Label3 
         Height          =   285
         Index           =   1
         Left            =   1710
         TabIndex        =   31
         Top             =   2340
         Width           =   5880
      End
      Begin VB.Label Label3 
         Height          =   285
         Index           =   0
         Left            =   1710
         TabIndex        =   30
         Top             =   1995
         Width           =   3885
      End
      Begin VB.Label Label1 
         Caption         =   "Fam  Sub"
         Height          =   300
         Index           =   8
         Left            =   915
         TabIndex        =   29
         Top             =   1725
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Aditiu"
         Height          =   300
         Index           =   7
         Left            =   90
         TabIndex        =   28
         Top             =   2655
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Colorant"
         Height          =   300
         Index           =   6
         Left            =   90
         TabIndex        =   27
         Top             =   2340
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Material"
         Height          =   300
         Index           =   5
         Left            =   90
         TabIndex        =   26
         Top             =   2025
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Ref.Mat."
         Height          =   300
         Index           =   4
         Left            =   75
         TabIndex        =   25
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label nomproveidor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1650
         TabIndex        =   24
         Top             =   690
         Width           =   3195
      End
      Begin VB.Label Label1 
         Caption         =   "Proveidor"
         Height          =   300
         Index           =   1
         Left            =   45
         TabIndex        =   22
         Top             =   675
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Codi/Desc"
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   21
         Top             =   285
         Width           =   945
      End
      Begin VB.Label Label18 
         Caption         =   "(Material verge)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7440
         TabIndex        =   131
         Top             =   1155
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   9555
      Begin VB.CommandButton Command14 
         BackColor       =   &H00ED823A&
         Caption         =   "Tarifes Families"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6885
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   195
         Width           =   1200
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   960
         Picture         =   "Materials.frx":F4BC
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   420
      End
      Begin Crystal.CrystalReport llistat 
         Left            =   2625
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton Command9 
         Height          =   360
         Index           =   1
         Left            =   8115
         Picture         =   "Materials.frx":FA46
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir llistat de materials >500"
         Top             =   195
         Width           =   420
      End
      Begin VB.CommandButton Command3 
         Height          =   360
         Left            =   1410
         Picture         =   "Materials.frx":FFD0
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   150
         Width           =   420
      End
      Begin VB.Data materials 
         Caption         =   "Materials"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   360
         Left            =   3570
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from materials where codi>499"
         Top             =   180
         Width           =   3000
      End
      Begin VB.CommandButton consultar 
         Height          =   360
         Left            =   8625
         Picture         =   "Materials.frx":1055A
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Busqueda de Registres"
         Top             =   180
         Width           =   420
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "Materials.frx":10AE4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   520
         Picture         =   "Materials.frx":1106E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Consulta Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   9075
         Picture         =   "Materials.frx":115F8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Sortir"
         Top             =   150
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Height          =   390
         Left            =   4290
         Picture         =   "Materials.frx":11B82
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   150
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label estattaula 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1995
         TabIndex        =   19
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Espesors"
      Height          =   3030
      Left            =   270
      TabIndex        =   33
      Top             =   825
      Visible         =   0   'False
      Width           =   4500
      Begin VB.TextBox micres 
         Height          =   285
         Left            =   1035
         TabIndex        =   36
         Top             =   420
         Width           =   585
      End
      Begin VB.TextBox grmm2 
         Height          =   285
         Left            =   2535
         TabIndex        =   35
         Top             =   420
         Width           =   690
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "Afegir"
         Height          =   285
         Left            =   3375
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   390
         Width           =   750
      End
      Begin VB.Data espesors 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   885
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "materials_espesors"
         Top             =   2820
         Visible         =   0   'False
         Width           =   1710
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Materials.frx":11E94
         Height          =   1995
         Left            =   660
         OleObjectBlob   =   "Materials.frx":11EA7
         TabIndex        =   37
         Top             =   810
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Micres"
         Height          =   255
         Left            =   375
         TabIndex        =   39
         Top             =   465
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Grms/m2"
         Height          =   270
         Left            =   1815
         TabIndex        =   38
         Top             =   435
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   300
      Index           =   2
      Left            =   90
      TabIndex        =   23
      Top             =   1605
      Width           =   405
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   135
      TabIndex        =   16
      Top             =   5985
      Width           =   4470
   End
   Begin VB.Label autonum 
      Height          =   135
      Left            =   0
      TabIndex        =   15
      Top             =   1335
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "fmaterials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ruta_FT As String
Private Sub colsbloc_Change()

End Sub

Private Sub alta_Click()
Dim gran As Long
Dim rst As Recordset
framematerials.Enabled = True
Set rst = dbtmp.OpenRecordset("select max(codi) as elgran from materials ")
If rst.EOF Then
   gran = 0
     Else: gran = cadbl(rst!elgran)
 End If
'materials.RecordSource = "select * from materials order by codi"
'materials.Recordset.MoveLast
'If Not materials.Recordset.EOF Then gran = materials.Recordset!codi
gran = gran + 1
materials.Recordset.AddNew
framematerials.Enabled = True
materials.Recordset!codi = gran
materials.Recordset!tipusCQ = "L"
Text1(0) = gran
Text1(2).SetFocus
Set rst = Nothing
End Sub

Private Sub bestoc_Click()
 Dataestoc.DatabaseName = materials.DatabaseName
  Dataestoc.RecordSource = "select * from materials_estoc where codi=" + atrim(materials.Recordset!codi) + " order by data desc"
  Dataestoc.Refresh
  frameestoc.Top = framematerials.Top
  frameestoc.Left = framematerials.Left
  calcular_total_estoc
  frameestoc.Visible = True
End Sub

Private Sub botopdf_Click()
  obrir_document ruta_FT + "\" + atrim(materials.Recordset!codi) + "\FT_" + treuresimbols(materials.Recordset!descripcio) + ".PDF"
End Sub

Private Sub botopdf_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  agafar_elPDF Data
End Sub
Sub agafar_elPDF(Data As DataObject)
  netejartemporals
  copiarpdfFTaltemporal Data
  copiarpdfFTdeltemporaladefinitiuidemanardata
End Sub
Sub copiarpdfFTdeltemporaladefinitiuidemanardata()
  guardarelpdf "c:\temp\tmpFT\dragover.pdf"
  materials.Recordset.Move 0
End Sub
Function demanardatapdf() As String
   
    demanardatapdf = InputBox("Entra la data de validesa del pdf." + Chr(10) + "   SI FAS CANCELAR O NO POSES DATA NO ES GUARDARÀ EL PDF." + Chr(10) + "Format dd/mm/yy o escriu ddmmyy", "Data PDF")
    If Len(demanardatapdf) = 6 Then demanardatapdf = Mid(demanardatapdf, 1, 2) + "/" + Mid(demanardatapdf, 3, 2) + "/" + Mid(demanardatapdf, 5, 2)
    If IsDate(demanardatapdf) Then
       demanardatapdf = "#" + Format(demanardatapdf, "mm/dd/yy") + "#"
      Else: demanardatapdf = "null"
    End If
    
End Function
Sub guardarelpdf(rutaorigenpdf As String, Optional tipuspdf As String, Optional vnomissatges As Boolean)
  Dim rutadesti As String
  Dim sobreescriure As Boolean
  Dim datapdf As String
  
  
  On Error GoTo erroricontinua
  If Not existeix(rutaorigenpdf) Then
     If Not vnomissatges Then MsgBox "Error... no trobo el fitxer PDF"
     Exit Sub
  End If
  rutadesti = ruta_FT + "\" + atrim(materials.Recordset!codi) + "\FT_" + treuresimbols(materials.Recordset!descripcio) + ".PDF"
  If existeix(rutadesti) Then
      If UCase(InputBox("Aquest material ja te un PDF de fitxa tècnica, VOLS SOBREESCRIURE'L?" + Chr(10) + " Escriu el [Si] per comfirmar-ho", "Sobrescriure el PDF")) = "SI" Then
              sobreescriure = True
         Else:
            Kill (rutaorigenpdf): Exit Sub
      End If
    Else: If Not existeix(ruta_FT + "\" + atrim(materials.Recordset!codi)) Then MkDir ruta_FT + "\" + atrim(materials.Recordset!codi)
  End If
  datapdf = demanardatapdf
  If datapdf = "null" Then Exit Sub
  If sobreescriure Then Kill (rutadesti)
  Copiar_Fitxer rutaorigenpdf, rutadesti
  dbtmp.Execute "update materials set Ftvigencia=" + datapdf + " where codi=" + atrim(materials.Recordset!codi)
  Exit Sub
erroricontinua:

End Sub
Sub copiarpdfFTaltemporal(Data As DataObject)
    Copiar_Fitxer Data.Files(1), "c:\temp\tmpFT\dragover.pdf"
End Sub
Sub netejartemporals()
  If existeix("c:\temp\tmpFT") Then CreateObject("Scripting.FileSystemObject").DeleteFolder "c:\temp\tmpFT"          ' Eliminamos carpeta.
  On Error Resume Next
  If Not existeix("c:\temp\tmpFT") Then MkDir "c:\temp\tmpFT"
  Kill "c:\temp\tmpFT\*.*"
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Checkactiu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Set rst = dbtmp.OpenRecordset("select * from tarifesproveidors where id=" + atrim(cadbl(dataoferta.ItemData(dataoferta.ListIndex))))
   If rst.EOF Then Exit Sub
   rst.Edit
   rst.actiupressupostos = IIf(Checkactiu.Value = 1, True, False)
   rst.Update
   carregar_oferta
End Sub

Private Sub ckgminim_DblClick()
   Dim v As String
   v = InputBox("Entra el valor de Kg mínim.", "Canvi Kg minim.", ckgminim)
   If cadbl(v) = 0 Then Exit Sub
   cambiarquantitatminima IIf(combomesura = "Kg", cadbl(v), cadbl(v) * -1)
  carregar_baremtarifa
End Sub
Sub cambiarquantitatminima(v As Double)
   datapreukg.Recordset.MoveFirst
   While Not datapreukg.Recordset.EOF
      datapreukg.Recordset.Edit
      datapreukg.Recordset!kgminim = v
      datapreukg.Recordset.Update
      datapreukg.Recordset.MoveNext
   Wend
End Sub
Private Sub cmicresoamplada_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If cmicresoamplada(Index) Then Exit Sub
   If MsgBox("Segur que vols canviar el sistema de càlcul de la oferta?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then
      Index = -1
     Else:
       cmicresoamplada(Index).Value = True
       dbtmp.Execute "update tarifesproveidors set micresoamplada='" + IIf(cmicresoamplada(0), "M", "A") + "' where id=" + atrim(cadbl(dataoferta.ItemData(dataoferta.ListIndex)))
   End If
   carregar_oferta
End Sub

Private Sub Combocara1_DropDown()
   Dim vnom As String
   Dim vcodi As Double
   escullir_tractamentcara vcodi, vnom
   Combocara1.Tag = vcodi
   Combocara1 = vnom
   materials.Recordset!codidescmatcara1 = vcodi
End Sub

Private Sub Combocara1_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub Combocara1_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Combocara2_DropDown()
   Dim vnom As String
   Dim vcodi As Double
   escullir_tractamentcara vcodi, vnom
   Combocara2.Tag = vcodi
   Combocara2 = vnom
   materials.Recordset!codidescmatcara2 = vcodi
End Sub
Sub escullir_tractamentcara(vcodi As Double, vnom As String)
  Load formseleccio
  formseleccio.Caption = "Selecciona el tractament de la cara"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codi,descripcio from tractamentcares order by descripcio"
  formseleccio.refrescar
  formseleccio.Width = 7000
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Width = 5000
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Show 1
  If seleccioret = 1 Then
   vnom = atrim(formseleccio.Data1.Recordset!descripcio)
   vcodi = formseleccio.Data1.Recordset!codi
  End If
  If combofamilies.Text = "" Then netejarcampsofertes
  Unload formseleccio
  SendKeys "{tab}"
End Sub


Private Sub Combocara2_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub Combocara2_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub combofamilies_DropDown()
  Load formseleccio
  formseleccio.Caption = "Selecciona families"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select descripciofamilia from tarifesproveidors where descripciofamilia<>'' order by descripciofamilia"
  formseleccio.refrescar
  formseleccio.Width = 9500
  formseleccio.Command3.Tag = "filtre"
  formseleccio.DBGrid2.RowHeight = formseleccio.DBGrid2.RowHeight * 2
  formseleccio.Show 1
  If seleccioret = 1 Then
   combofamilies.Text = atrim(formseleccio.Data1.Recordset!descripciofamilia)
   carregarcomboofertes
   If dataoferta.ListCount > 0 Then
       dataoferta.ListIndex = 0
       carregar_oferta
   End If
  End If
  If combofamilies.Text = "" Then netejarcampsofertes
  Unload formseleccio
  SendKeys "{tab}"
End Sub

Private Sub combomesura_Click()
   If MsgBox("Segur que vols cambiar la mesura de la quantitat mínima?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then
        carregar_baremtarifa
        Exit Sub
   End If
   cambiarquantitatminima IIf(combomesura = "Kg", cadbl(ckgminim), cadbl(ckgminim) * -1)
End Sub

Private Sub combosubfamcompatible_DropDown()
   cmatcompatible = triar_familia("select * from subfamiliesmaterials where codifam=" + atrim(cadbl(Text1(4))))
   descfamilies
   SendKeys "{TAB}"
End Sub

Private Sub CombotipusCQ_Click()
  If Not fmaterials.Visible Then Exit Sub
  If Screen.ActiveControl.Name = "CombotipusCQ" Then
   materials.Recordset!tipusCQ = IIf(CombotipusCQ.ListIndex = 0, "C", materials.Recordset!tipusCQ)
   materials.Recordset!tipusCQ = IIf(CombotipusCQ.ListIndex = 1, "L", materials.Recordset!tipusCQ)
   materials.Recordset!tipusCQ = IIf(CombotipusCQ.ListIndex = 2, "N", materials.Recordset!tipusCQ)
   materials.Recordset!tipusCQ = IIf(CombotipusCQ.ListIndex = 3, "", materials.Recordset!tipusCQ)
   If CombotipusCQ.ListIndex = 0 Then
      demanar_dataCQ
        Else: etdataqualitat = ""
   End If
  End If
End Sub

Sub demanar_dataCQ()
   Dim v As String
   If CombotipusCQ.ListIndex <> 0 Then Exit Sub
   v = InputBox("Entra la data de caducitat de la 'Qualitat concertada'.", "Data", etdataqualitat)
   If StrPtr(v) = 0 Then Exit Sub
   If v = "" Then etdataqualitat = ""
   If IsDate(v) Then
       If DateDiff("d", Now, v) < 0 Then MsgBox "Aquesta data ja està passada.", vbCritical, "Error": Exit Sub
       etdataqualitat = v
   End If
End Sub

Private Sub CombotipusCQ_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
End Sub

Private Sub Command10_Click()
   nova_oferta
End Sub
Sub nova_oferta()
   Dim rst As Recordset
   Dim vdata As String
   Dim vvigencia As String
   Dim vnumoferta As String
   Dim rstclone As Recordset
   Dim vdescfam As String
   
   'demano els valors de la capçalera de l'oferta
   vdata = InputBox("Entra la data de la oferta. (dd/mm/yy)", "Data oferta")
   If Not IsDate(vdata) Then MsgBox "Aquesta data no es vàlida", vbCritical, "Error": Exit Sub
   vvigencia = InputBox("Entra la vigència d'aquesta oferta." + Chr(10) + "Pots posar la data (dd/mm/yy) o els mesos que vols que duri.", "Data vigència")
   If cadbl(vvigencia) > 0 Then vvigencia = DateAdd("m", cadbl(vvigencia), vdata)
   If atrim(vvigencia) <> "" Then
        If Not IsDate(vvigencia) Then MsgBox "Aquesta data de vigència no es vàlida", vbCritical, "Error": Exit Sub
   End If
   vnumoferta = InputBox("Entra el numero d'oferta del proveidor.", "Oferta")
   vnumoferta = Mid(vnumoferta, 1, 15)
   
   'creo la oferta a la BD
   Set rst = dbtmp.OpenRecordset("select * from tarifesproveidors")
   rst.AddNew
   rst!dataoferta = vdata
   rst!datavenciment = vvigencia
   rst!numerooferta = vnumoferta
   If Not framefamilies.Visible Then
      rst!materialrelacionat = materials.Recordset!codi
        Else:
          If fammat = "" Then
                Set rstclone = dbtmp.OpenRecordset("select * from tarifesproveidors where descripciofamilia='" + atrim(combofamilies.Text) + "'")
                If Not rstclone.EOF Then
                      rst!descripciofamilia = combofamilies.Text
                      rst!familia = rstclone!familia
                      rst!familiacol = rstclone!familiacol
                      rst!familiaad = rstclone!familiaad
                      rst!subfamilia = rstclone!familia
                      rst!subfamiliacol = rstclone!subfamiliacol
                      rst!subfamiliaad = rstclone!subfamiliaad
                       Else: MsgBox "Error": rst.CancelUpdate: Exit Sub
                End If
                  Else
                      rst!descripciofamilia = fammat + "+" + subfammat + "+" + famcol + "+" + subfamcol + "+" + famad + "+" + subfamad
                      combofamilies = atrim(rst!descripciofamilia)
                      rst!familia = fammat.ItemData(fammat.ListIndex)
                      rst!familiacol = subfammat.ItemData(subfammat.ListIndex)
                      rst!familiaad = famad.ItemData(famad.ListIndex)
                      rst!subfamilia = subfamad.ItemData(subfamad.ListIndex)
                      rst!subfamiliacol = famcol.ItemData(famcol.ListIndex)
                      rst!subfamiliaad = subfamcol.ItemData(subfamcol.ListIndex)
          End If
   End If
   rst.Update
   'recarrego les ofertes
   carregarcomboofertes vdata
   
   Set rst = Nothing
   Set rstclone = Nothing
End Sub

Private Sub Command11_Click()
   If MsgBox("Segur que vols eliminar tota aquesta oferta i els escandall relacionats amb les micres?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   dbtmp.Execute "delete * from tarifesproveidorsescandall where id=" + atrim(cadbl(dataoferta.ItemData(dataoferta.ListIndex)))
   dbtmp.Execute "delete * from tarifesproveidors where id=" + atrim(cadbl(dataoferta.ItemData(dataoferta.ListIndex)))
   carregarcomboofertes
   If dataoferta.ListCount > 0 Then
        dataoferta.ListIndex = 0
        carregar_oferta
   End If
End Sub

Private Sub Command12_Click()
   frameseleccio.Visible = True
   netejarcombos
   frameseleccio.Top = 570
   frameseleccio.Left = 1380
End Sub
Sub netejarcombos()
  fammat.ListIndex = -1: subfammat.ListIndex = -1
  famcol.ListIndex = -1: subfamcol.ListIndex = -1
  famad.ListIndex = -1: subfamad.ListIndex = -1
End Sub

Private Sub Command13_Click()
  Dim i As Integer
  If MsgBox("Segur que vols eliminar TOTES les ofertes i els escandall relacionats amb les micres?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
  While i < dataoferta.ListCount
         dataoferta.ListIndex = i
         dbtmp.Execute "delete * from tarifesproveidorsescandall where id=" + atrim(cadbl(dataoferta.ItemData(dataoferta.ListIndex)))
         dbtmp.Execute "delete * from tarifesproveidors where id=" + atrim(cadbl(dataoferta.ItemData(dataoferta.ListIndex)))
         i = i + 1
  Wend
  
  carregarfamiliesiofertes
  
End Sub

Private Sub Command14_Click()
    Frameescandall.Visible = Not Frameescandall.Visible
    framefamilies.Visible = Frameescandall.Visible
    Frameescandall.Left = 60
    Frameescandall.Top = 1575
    framefamilies.Left = 30
    framefamilies.Top = 780
    netejarcampsofertes
    'framefamilies.Visible = True
    
    If Frameescandall.Visible Then
       carregarfamiliesiofertes
    End If
End Sub
Sub carregarfamiliesiofertes()
  carregarcomboofertesFAMILIES
  carregarcomboofertes
  If dataoferta.ListCount > 0 Then
     dataoferta.ListIndex = 0
     carregar_oferta
  End If
End Sub
Sub carregarcomboofertesFAMILIES()
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select descripciofamilia from tarifesproveidors order by dataoferta desc")
   If Not rst.EOF Then combofamilies.Text = atrim(rst!descripciofamilia)
   Set rst = Nothing
End Sub
Private Sub Command15_Click()
   Dim vfinsa As Double
   Dim vpreukg As Double
   Dim rst As Recordset
   Dim vkggranbarem As Double
   
   'poso el valor de minim de kg per no deixar entrar un valor inferior
   If Not datapreukg.Recordset.EOF Then
        datapreukg.Recordset.MoveLast
        vkggranbarem = CDbl(datapreukg.Recordset!finsaxkg)
   End If
   If vkggranbarem < cadbl(ckgminim) Then vkggranbarem = cadbl(ckgminim)
   '------------
   
   vfinsa = cadbl(InputBox("Entra fins a quants KG serà aquest preu." + Chr(10) + " Ex: 1000 (aquest valor inclós)", "Kg de l'escalat."))
   If cadbl(vfinsa) <= 0 Then Exit Sub
   If cadbl(vfinsa) <= cadbl(vkggranbarem) Then MsgBox "Els Kg han de ser mes gran que els Kg anteriors del barem.", vbCritical, "Error": Exit Sub
   vpreukg = cadbl(InputBox("Entra el preu que et faran fins a " + atrim(vfinsa) + " Kgs", "Preu Kg"))
   If cadbl(vpreukg) <= 0 Then Exit Sub
   Set rst = datapreukg.Recordset.Clone
   If datapreukg.Recordset!finsaxkg = 0 Then
      datapreukg.Recordset.Edit
        Else
          datapreukg.Recordset.AddNew
          datapreukg.Recordset!ID = rst!ID
          datapreukg.Recordset!dexmicres = rst!dexmicres
          datapreukg.Recordset!axmicres = rst!axmicres
          datapreukg.Recordset!kgminim = rst!kgminim
   End If
   datapreukg.Recordset!finsaxkg = vfinsa
   datapreukg.Recordset!preuxrkg = vpreukg
   datapreukg.Recordset.Update
   datapreukg.Refresh
   
End Sub

Private Sub Command16_Click()
   If MsgBox("Segur que vols eliminar aquest valor de l'escandall?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   If datapreukg.Recordset.RecordCount = 1 Then
       datapreukg.Recordset.Edit
       datapreukg.Recordset!finsaxkg = 0
       datapreukg.Recordset!preuxrkg = 0
       datapreukg.Recordset.Update
      Else
         datapreukg.Recordset.Delete
   End If
   datapreukg.Refresh
End Sub

Private Sub Command17_Click()
  frameseleccio.Visible = False
End Sub

Private Sub Command18_Click()
  frameseleccio.Visible = False
  nova_oferta
End Sub

Private Sub Command19_Click()
  framesubstancies.Left = 2640
  framesubstancies.Top = 165
  framesubstancies.ZOrder 0
  framesubstancies.Visible = Not framesubstancies.Visible
End Sub

Private Sub Command2_Click()
  If Not existeixrang Then
      espesors.Recordset.AddNew
       espesors.Recordset!codi = materials.Recordset!codi
       espesors.Recordset!micres = cadbl(micres)
       espesors.Recordset!grmsm2 = cadbl(grmm2)
      espesors.Recordset.Update
     Else: MsgBox "Aquest espesor ja existeix."
  End If
End Sub
Function existeixrang() As Boolean
  existeixrang = False
  If cadbl(micres) > 0 And espesors.Recordset.RecordCount > 0 Then
   espesors.Recordset.FindFirst "micres=" + atrim(cadbl(micres))
   If Not espesors.Recordset.NoMatch Then
      existeixrang = True
   End If
  End If
  If (cadbl(micres) = 0 And cadbl(grmm2) > 0) And espesors.Recordset.RecordCount > 0 Then
   espesors.Recordset.FindFirst "grmsm2=" + atrim(cadbl(grmm2))
   If Not espesors.Recordset.NoMatch Then
      existeixrang = True
   End If
  End If
  
End Function

Private Sub Command20_Click()
  If Not datasubstancies.Recordset.EOF And Not datasubstancies.Recordset.BOF Then
      If MsgBox("Estas segur que vols treure aquesta substancia?" + vbNewLine + datasubstancies.Recordset!descripcio, vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
      dbtmp.Execute "delete * from materials_substancies where codimaterial=" + atrim(materials.Recordset!codi) + " and codisubstancia='" + atrim(datasubstancies.Recordset!codi) + "'"
      datasubstancies.Refresh
  End If
End Sub

Private Sub Command21_Click()
  Dim vmgkg As String
  Dim rst As Recordset
  If cdataanalisis = "" Then MsgBox "Primer has de posar la data de l'analisis.", vbCritical, "Atenció": Exit Sub
  Load formseleccio
  formseleccio.Data1.DatabaseName = materials.DatabaseName
  formseleccio.Data1.RecordSource = "select * from substancies where codiref<>'' and codiref<>null order by descripcio"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 800
  formseleccio.DBGrid2.Columns(1).Width = 1300
  formseleccio.DBGrid2.Columns(2).Width = 6000
  formseleccio.DBGrid2.Columns(3).Width = 800
  formseleccio.Width = 10000
  'formseleccio.Command3.Tag = "filtre"
  formseleccio.Text1.Tag = "0"
  formseleccio.Show 1
  If seleccioret = 1 Then
      Set rst = dbtmp.OpenRecordset("select * from materials_substancies")
      vmgkg = InputBox("Entra el valor de mg/Kg d'aquesta substancia:", "Valor mg/Kg")
      vmgkg = substituir(vmgkg, ".", ",")
      rst.AddNew
      rst!codisubstancia = atrim(formseleccio.Data1.Recordset!codiref)
      rst!codimaterial = materials.Recordset!codi
      rst![mg/kg] = cadbl(vmgkg)
      rst.Update
      datasubstancies.Refresh
      Set rst = Nothing
  End If
  Unload formseleccio
End Sub

Private Sub Command22_Click()
   If Not materials.Recordset.EOF Then
       materials.RecordSource = "select * from materials where proveidor=" + atrim(materials.Recordset!proveidor)
       materials.Refresh
   End If
End Sub

Private Sub Command23_Click()
    Dim vfam As String
    For i = 5 To 9
       vfam = vfam + IIf(vfam <> "", " and ", "") + Text1(i).DataField + "=" + atrim(cadbl(Text1(i)))
    Next i
    If Not materials.Recordset.EOF Then
       materials.RecordSource = "select * from materials where " + vfam
       materials.Refresh
   End If
End Sub

Private Sub Command24_Click()
   Load formMaterialsdetalldescripcio
   formMaterialsdetalldescripcio.Tag = atrim(materials.Recordset!codi)
   formMaterialsdetalldescripcio.carregar_combo_refinplacsa formMaterialsdetalldescripcio.Tag
   formMaterialsdetalldescripcio.actualitzardades
   
   
   formMaterialsdetalldescripcio.Show 1
End Sub

Private Sub Command25_Click()
  Dim vcodimat As String
  Dim rst As Recordset
  Dim rst2 As Recordset
  vcodimat = InputBox("Escriu el codi de material A QUI VOLS COPIAR AQUESTES SUBSTANCIES.", "COPIAR SUBSTANCIES A UN ALTRA MATERIAL")
  If cadbl(vcodimat) = 0 Then Exit Sub
  Set rst = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(vcodimat))
  If rst.EOF Then GoTo fi
  If MsgBox("Segur que vols passar aquestes substancies al material " + vbNewLine + UCase(rst!descripcio), vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then GoTo fi
  Set rst = dbtmp.OpenRecordset("select * from materials_substancies where codimaterial=" + atrim(materials.Recordset!codi))
  Set rst2 = dbtmp.OpenRecordset("select * from materials_substancies")
  dbtmp.Execute "delete * from materials_substancies where codimaterial=" + atrim(vcodimat)
  dbtmp.Execute "update  materials set dataanalisissubstancies=#" + Format(materials.Recordset!dataanalisissubstancies, "mm/dd/yy") + "# where codi=" + atrim(vcodimat)
  rst.MoveLast
  rst.MoveFirst
  
  While Not rst.EOF
     rst2.AddNew
     rst2!codimaterial = vcodimat
     rst2!codisubstancia = rst!codisubstancia
     rst2![mg/kg] = rst![mg/kg]
     rst2.Update
     rst.MoveNext
    
  Wend
  If MsgBox("Vols anar al registre que acabes de modificar?", vbDefaultButton2 + vbYesNoCancel, "Atenció") = vbYes Then
      materials.Recordset.FindFirst "codi=" + atrim(vcodimat)
  End If
fi:
  Set rst = Nothing
  Set rst2 = Nothing
End Sub

Private Sub Command26_Click()
  possar_data_analisis
End Sub
Sub possar_data_analisis()
  Dim v As String
  v = InputBox("Entra la data de l'analisis: " + vbNewLine + "Ex: 01/01/25")
  If IsDate(v) Then
     If DateDiff("d", v, Now) < 0 Then MsgBox "La data entrada no pot ser superior a la data d 'avui", vbCritical, "Error": GoTo fi
     'If materials.Recordset.EditMode = 0 Then materials.Recordset.Edit
     'materials.Recordset!dataanalisissubstancies = v
     'materials.Recordset.Update: materials.Recordset.Edit
     dbtmp.Execute "update  materials set dataanalisissubstancies=#" + Format(v, "mm/dd/yy") + "# where codi=" + atrim(materials.Recordset!codi)
     cdataanalisis = Format(v, "dd/mm/yy")
  End If
fi:
End Sub


Sub calcular_total_estoc()
  Dim vtotal As Double
  Dim rst As Recordset
  ettotalestoc = ""
  Set rst = dbtmp.OpenRecordset("select * from materials_estoc where codi=" + atrim(materials.Recordset!codi) + " order by data desc")
  rst.FindFirst "tipus='R'"
  If rst.NoMatch Then ettotalestoc = "No hi ha cap reguralització.": GoTo fi
  While Not rst.BOF
     If rst!tipus = "R" Or rst!tipus = "A" Then vtotal = vtotal + cadbl(rst!quantitat)
     If rst!tipus = "T" Then vtotal = vtotal - cadbl(rst!quantitat)
     rst.MovePrevious
  Wend
  ettotalestoc = "Total estoc: " + atrim(vtotal)
fi:
  Set rst = Nothing
End Sub

Private Sub Command28_Click()
frameestoc.Visible = False
End Sub

Private Sub Command29_Click()
  
  Dim vquantitat As Double
  vquantitat = InputBox("Escriu la quanitat a Reguralitzar l'estoc." + vbNewLine + "D'AQUI NOMÉS POTS REGURALITZAR AFEGIR O TREURE ESTOC ES FA AMB COMPRES.", "Reguralitzar")
  If cadbl(vquantitat) = 0 Then Exit Sub
  Dataestoc.Recordset.AddNew
  Dataestoc.Recordset!codi = materials.Recordset!codi
  Dataestoc.Recordset!tipus = "R"
  Dataestoc.Recordset!Data = Now
  Dataestoc.Recordset!quantitat = cadbl(vquantitat)
  Dataestoc.Recordset!numerodocument = nomordinador
  Dataestoc.Recordset.Update
  Dataestoc.Refresh
  calcular_total_estoc
End Sub

Private Sub Command3_Click()
  
  gravar_registre
End Sub

Private Sub Command30_Click()
  If Dataestoc.Recordset.EOF Then MsgBox "Primer escull el registre que vols eliminar.", vbCritical, "Atenció": Exit Sub
  If Dataestoc.Recordset!tipus <> "R" Then MsgBox "El registre escullit ha de ser R (Reguralització).", vbCritical, "Error": Exit Sub
  If MsgBox("Segur que vols eliminar aquesta reguralització?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
      Dataestoc.Recordset.Delete
      Dataestoc.Refresh
      calcular_total_estoc
  End If
End Sub

Private Sub Command4_Click()
   frameFT.Visible = Not frameFT.Visible
   frameFT.Left = 4905
   frameFT.Top = 600
   frameFT.ZOrder 0
End Sub

Private Sub Command5_Click()
  Dim vpdfFT As String
  vpdfFT = ruta_FT + "\" + atrim(materials.Recordset!codi) + "\FT_" + treuresimbols(materials.Recordset!descripcio) + ".PDF"
  If existeix(vpdfFT) Then
    If UCase(InputBox("Estàs segur que vols eliminar aquest PDF?" + Chr(10) + "Escriu [SI] si estàs d'acord.", "Eliminar PDF")) = "SI" Then
      Kill vpdfFT
      dbtmp.Execute "update materials set Ftvigencia=null where codi=" + atrim(materials.Recordset!codi)
      wait 1
      materials.Recordset.Move 0
    End If
  End If
End Sub

Private Sub Command6_Click()
    Frameescandall.Visible = Not Frameescandall.Visible
    Frameescandall.Left = 60
    Frameescandall.Top = 1575
    framefamilies.Visible = False
    If Frameescandall.Visible Then
       carregarcomboofertes
       If dataoferta.ListCount > 0 Then
        dataoferta.ListIndex = 0
        carregar_oferta
       End If
    End If
End Sub
Sub carregarcomboofertes(Optional vdata As String)
   Dim rst As Recordset
   Dim v As String
   Dim vindex As Long
   vindex = -1
   If Not framefamilies.Visible Then
        v = " materialrelacionat=" + atrim(materials.Recordset!codi)
          Else: v = " descripciofamilia='" + combofamilies + "' "
   End If
   Set rst = dbtmp.OpenRecordset("select * from tarifesproveidors where " + v + " order by dataoferta desc")
   dataoferta.Clear
   While Not rst.EOF
     dataoferta.AddItem Format(rst!dataoferta, "dd/mm/yy")
     dataoferta.ItemData(dataoferta.NewIndex) = rst!ID
     If Format(rst!dataoferta, "dd/mm/yy") = Format(vdata, "dd/mm/yy") Then vindex = dataoferta.NewIndex
     rst.MoveNext
   Wend
   
   If vindex <> -1 Then dataoferta.ListIndex = vindex
   Set rst = Nothing
End Sub

Private Sub Command7_Click()
   Dim vde As String
   Dim va As String
   Dim rst As Recordset
   Dim kgmin As String
   
   vde = InputBox("Entra el valor de la micra que vols fer l'escandall." + Chr(10) + "Si es un rang de micres separa-les per guió. Ex: 20-100", "Micres")
   If InStr(1, vde, "-") > 0 Then
       va = cadbl(Mid(vde, InStr(1, vde, "-") + 1))
       vde = cadbl(Mid(vde, 1, InStr(1, vde, "-") - 1))
      Else: va = vde
   End If
   kgmin = InputBox("Entra els Kgs mínim d'aquest barem.", "Kg mínim")
   If cadbl(kgmin) = 0 Then GoTo fi
   Set rst = dbtmp.OpenRecordset("select* from tarifesproveidorsescandall  where id=" + atrim(cadbl(dataoferta.ItemData(dataoferta.ListIndex))) + " order by finsaxkg")
   rst.FindFirst "dexmicres=" + atrim(vde)
   If Not rst.NoMatch Then MsgBox "Aquestes micres ja estan donades d'alta si vols utilitzar-les de nou hauras de borrar les anteriors.", vbExclamation, "Error": GoTo fi
   rst.AddNew
   rst!dexmicres = cadbl(vde)
   rst!axmicres = cadbl(va)
   rst!kgminim = cadbl(kgmin)
   rst!ID = cadbl(dataoferta.ItemData(dataoferta.ListIndex))
   rst.Update
   
fi:
   Set rst = Nothing
   carregar_micresofertades cadbl(vde)
End Sub

Private Sub Command8_Click()
   If MsgBox("Segur que vols eliminar aquest valor micres i tot l'escandall relacionat?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   dbtmp.Execute "delete * from tarifesproveidorsescandall where id=" + atrim(cadbl(dataoferta.ItemData(dataoferta.ListIndex))) + " and dexmicres=" + atrim(cadbl(llistamicres.ItemData(llistamicres.ListIndex)))
   carregar_oferta
End Sub

Private Sub Command9_Click(Index As Integer)
  If MsgBox("Vols imprimir agrupant per proveidor?", vbInformation + vbYesNo, "Atenció") = vbYes Then
     llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "llistatdematerials.rpt"
    Else
      llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "llistatdematerialssensegrups.rpt"
  End If
  llistat.DataFiles(0) = cami
  llistat.Destination = crptToWindow
  llistat.Action = 1
End Sub

Private Sub consultar_Click()
   Dim b As String
   b = InputBox("Entra la Descripcio/RefProducte a buscar o el Codi" + Chr(10) + " No escriguis res per treure els filtres", "Busqueda")
   b = treure_apostruf(b)
   If cadbl(b) > 0 Then
     materials.RecordSource = "select * from materials where codi>499 and codi=" + atrim(cadbl(b)) + ""
     materials.Refresh
     b = ""
      Else
       If b <> "" Then
        materials.RecordSource = "select * from materials where codi>499 and descripcio like '*" + b + "*' or refproducte like '*" + b + "*'"
        materials.Refresh
          Else
             materials.RecordSource = "select * from materials where codi>499 "
             materials.Refresh
       End If
   End If
End Sub

Private Sub dataoferta_Click()
   carregar_oferta
End Sub
Sub carregar_oferta()
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from tarifesproveidors where id=" + atrim(cadbl(dataoferta.ItemData(dataoferta.ListIndex))))
   If Not rst.EOF Then
        cvigencia = rst!datavenciment
        cnumoferta = rst!numerooferta
        Checkactiu = IIf(rst!actiupressupostos, 1, 0)
        cmicresoamplada(0).Value = True
        If atrim(rst!micresoamplada) = "A" Then cmicresoamplada(1).Value = True
        posarvalorkgminim 0
        carregar_micresofertades
   End If
   Set rst = Nothing
End Sub
Sub carregar_micresofertades(Optional vde As Double)
   Dim rst As Recordset
   Dim vultim As Double
   Dim v As String
   Dim vindex As Long
   vindex = -1
   datapreukg.RecordSource = "select * from tarifesproveidorsescandall where 1=2"
   datapreukg.Refresh
   Set rst = dbtmp.OpenRecordset("select * from tarifesproveidorsescandall where id=" + atrim(cadbl(dataoferta.ItemData(dataoferta.ListIndex))) + " order by dexmicres")
   llistamicres.Clear
   While Not rst.EOF
      If vultim <> rst!dexmicres Then
         v = IIf(rst!dexmicres = rst!axmicres, atrim(rst!dexmicres), atrim(rst!dexmicres) + " ~ " + atrim(rst!axmicres))
         llistamicres.AddItem v + IIf(cmicresoamplada(0).Value, "µ", "mm")
         llistamicres.ItemData(llistamicres.NewIndex) = rst!dexmicres
         If rst!dexmicres = vde Then vindex = llistamicres.NewIndex
         vultim = rst!dexmicres
      End If
      rst.MoveNext
   Wend
   If llistamicres.ListCount > 0 Then
       llistamicres.ListIndex = IIf(vindex <> -1, vindex, 0)
       datapreukg.RecordSource = "select * from tarifesproveidorsescandall where dexmicres=" + atrim(llistamicres.ItemData(0)) + " and id=" + atrim(cadbl(dataoferta.ItemData(dataoferta.ListIndex))) + " order by finsaxkg"
       datapreukg.Refresh
       carregar_baremtarifa
   End If
   Set rst = Nothing
End Sub

Private Sub dataoferta_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub dataoferta_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub eliminar_Click()
  Dim vcodi As String
  vcodi = cadbl(materials.Recordset!codi)
  If MsgBox("Segur que vols borrar aquest material?", vbCritical + vbYesNo + vbDerfaultButton2, "Atenció") = vbYes Then
     If InputBox("Escriu la paraula [ELIMINAR] el material " + atrim(vcodi) + "-" + atrim(materials.Recordset!descripcio) + vbNewLine + " per fer efectiu l'eliminació", "Control de seguretat") = "ELIMINAR" Then
         dbtmp.Execute ("delete * from materials_espesors where codi=" + atrim(cadbl(vcodi)))
         dbtmp.Execute ("delete * from materials where codi=" + atrim(cadbl(vcodi)))
        ' materials.Recordset.Delete
         enviaremailgeneric "miquel.inplacsa@gmail.com", "Material eliminat: " + nomordinador, "Material:rr " + atrim(materials.Recordset!codi) + " - " + atrim(materials.Recordset!descripcio)
         materials.Refresh
     End If
  End If
  
End Sub

Private Sub etdataqualitat_DblClick()
   demanar_dataCQ
End Sub

Private Sub Form_Click()
   'While Not materials.Recordset.BOF
   '  modificar_Click
   '  DoEvents
   '  gravar_registre
   '  materials.Recordset.MovePrevious
   'Wend
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then cancelar_registre
If KeyCode = 112 Then gravar_registre
End Sub
Sub gravar_registre()
  Dim i As Byte
  Dim nogravar As Boolean
  Command3.SetFocus
  DoEvents
  'MsgBox materials.Recordset!codidescmatcara1
  If cadbl(Text1(10)) = 0 And cadbl(Text1(14)) > 0 And mesespcompra <> "Grm/m2" Then MsgBox "Per poder gravar els canvis has d'entrar un valor a Grm/m3.", vbCritical, "Atenció": Exit Sub
  nogravar = False
  If materials.Recordset.EditMode > 0 Then
    framesubstancies.Visible = False
    If atrim(mesespcompra) = "" Then MsgBox "No hi ha mesura del producte entrada", vbCritical, "Error": Exit Sub
    If mesespcompra = "Grm/m2" And cadbl(Text1(12)) = 0 Then MsgBox "Has d'entrar l'equivalencia en micres al camp d'Espessor quan es Grm/m2", vbInformation, "Atenció": Exit Sub
    If mesespcompra <> "Unitats" Then
        For i = 4 To 9
           If cadbl(Text1(i)) = 0 Then nogravar = True
        Next i
        If nogravar Then MsgBox "Falta possar alguna familia o subfamilia abans de guardar les dades": Exit Sub
        If cadbl(Text1(4)) < 500 Or cadbl(Text1(5).Text) < 500 Or cadbl(Text1(6).Text) < 500 Then MsgBox "El codi de les families ha de ser mes gran de 500": Exit Sub
    End If
    If Not comprovar_families(cadbl(Text1(4)), cadbl(Text1(7)), cadbl(Text1(5)), cadbl(Text1(8)), cadbl(Text1(6)), cadbl(Text1(9))) Then
        MsgBox "Falta possar alguna familia o subfamilia abans de guardar les dades": Exit Sub
    End If
    generar_id_families cadbl(Text1(4)), cadbl(Text1(7)), cadbl(Text1(5)), cadbl(Text1(8)), cadbl(Text1(6)), cadbl(Text1(9))
    dbtmp.Execute "update materials set codidescmatcara1=" + atrim(materials.Recordset!codidescmatcara1) + " where codi=" + atrim(materials.Recordset!codi)
    dbtmp.Execute "update materials set codidescmatcara2=" + atrim(materials.Recordset!codidescmatcara2) + " where codi=" + atrim(materials.Recordset!codi)
    materials.Recordset.Update
    framematerials.Enabled = False
    frameFT.Visible = False
    framesubstancies.Visible = False
    materials.Recordset.Bookmark = materials.Recordset.LastModified
  End If
End Sub
Function comprovar_families(fam As Double, subfam As Double, famcol As Double, subfamcol As Double, famad As Double, subfamad As Double) As Boolean
   comprovar_families = True
   If fam = 0 Or subfam = 0 Or famcol = 0 Or subfamcol = 0 Or famad = 0 Or subfamad = 0 Then comprovar_families = False
End Function
Sub generar_id_families(fam As Double, subfam As Double, famcol As Double, subfamcol As Double, famad As Double, subfamad As Double)
   Dim rstf As Recordset
   Set rstf = dbtmp.OpenRecordset("select id_familia from materials where id_familia<>null and familia=" + atrim(fam) + " and subfamilia=" + atrim(subfam) + " and familiacol=" + atrim(famcol) + " and subfamiliacol=" + atrim(subfamcol) + " and familiaad=" + atrim(famad) + " and subfamiliaad=" + atrim(subfamad))
   'MsgBox "select id_familia from materials where familia=" + atrim(fam) + " and subfamilia=" + atrim(subfam) + " and familiacol=" + atrim(famcol) + " and subfamiliacol=" + atrim(subfamcol) + " and familiaad=" + atrim(famad) + " and subfamiliaad=" + atrim(subfamad)
   If Not rstf.EOF Then
    If cadbl(rstf!id_familia) > 0 Then
       materials.Recordset!id_familia = rstf!id_familia
       Exit Sub
    End If
   End If
   Set rstf = dbtmp.OpenRecordset("select max(id_familia) as gran from materials")
   materials.Recordset!id_familia = 1
   If Not rstf.EOF Then materials.Recordset!id_familia = cadbl(rstf!gran) + 1
End Sub
Sub cancelar_registre()
 If espesors.EditMode > 0 Then espesors.Recordset.CancelUpdate
 If materials.EditMode > 0 Then materials.Recordset.CancelUpdate
 framematerials.Enabled = False
End Sub

Private Sub Form_Load()
  materials.DatabaseName = cami
  espesors.DatabaseName = cami
  datapreukg.DatabaseName = cami
  datasubstancies.DatabaseName = cami
  frameFT.Visible = False
  framesubstancies.Visible = False
  ruta_FT = llegir_ini("ruta", "ruta_documentacio_FitxesTecniques", rutadelfitxer(cami) + "valorsprograma.ini")
  carregar_combo_families
End Sub

Private Sub Form_Resize()
  reixa.Width = fmaterials.Width - reixa.Left - 200
  
End Sub

Private Sub llistamicres_Click()
   carregar_baremtarifa
End Sub
Sub carregar_baremtarifa()
    datapreukg.RecordSource = "select * from tarifesproveidorsescandall where dexmicres=" + atrim(llistamicres.ItemData(llistamicres.ListIndex)) + " and id=" + atrim(cadbl(dataoferta.ItemData(dataoferta.ListIndex))) + " order by finsaxkg asc"
   datapreukg.Refresh
   ckgminim = ""
   If Not datapreukg.Recordset.EOF Then
      posarvalorkgminim cadbl(datapreukg.Recordset!kgminim)
   End If
End Sub
Sub posarvalorkgminim(vvalor As Double)
      If vvalor < 0 Then
           ckgminim = atrim(vvalor * -1)
           combomesura = "Mts"
            Else
              ckgminim = atrim(vvalor)
              combomesura = "Kg"
      End If
End Sub
Private Sub materials_Reposition()
  framematerials.Enabled = False
  Frameescandall.Visible = False
  frameseleccio.Visible = False
  framesubstancies.Visible = False
  frameestoc.Visible = False
  'framesubstancies.Visible = False
  carregar_camps
  If Not materials.Recordset.EOF Then materials.Caption = "Mat. " + atrim(1 + cadbl(materials.Recordset.AbsolutePosition)) + "/" + atrim(cadbl(materials.Recordset.RecordCount))
  botopdf.Picture = botopdf.DisabledPicture
  If existeix(ruta_FT + "\" + atrim(materials.Recordset!codi) + "\FT_" + treuresimbols(atrim(materials.Recordset!descripcio)) + ".PDF") Then botopdf.Picture = botopdf.DownPicture
  If mesespcompra = "Un" Then bestoc.Visible = True Else bestoc.Visible = False
End Sub
Sub carregar_camps()
  Dim rstp As Recordset
  nomproveidor = ""
  Label3(0) = ""
  Label3(1) = ""
  Label3(2) = ""
  Combocara1 = ""
  Combocara2 = ""
  
  
  If materials.Recordset.EOF Then
    espesors.RecordSource = ""
    espesors.Refresh
    datasubstancies.RecordSource = ""
    datasubstancies.Refresh
    Exit Sub
  End If
  If cadbl(materials.Recordset!codi) > 0 Then
   espesors.RecordSource = "select * from materials_espesors where codi=" + atrim(cadbl(materials.Recordset!codi))
   datasubstancies.RecordSource = "SELECT materials_substancies.codimaterial,materials_substancies.codisubstancia as Codi, substancies.descripcio as Descripcio,materials_substancies.[mg/kg] FROM materials_substancies LEFT JOIN substancies ON materials_substancies.codisubstancia = substancies.codiref where codimaterial = " + atrim(cadbl(materials.Recordset!codi))
     Else: espesors.RecordSource = "": datasubstancies.RecordSource = ""
  End If
  espesors.Refresh
  datasubstancies.Refresh
  Set rstp = materials.Database.OpenRecordset("select * from proveidors where codi=" + atrim(cadbl(materials.Recordset!proveidor)))
  If Not rstp.EOF And cadbl(materials.Recordset!proveidor) > 0 Then
     nomproveidor = atrim(rstp!nom)
       Else: nomproveidor = ""
  End If
  descfamilies
  netejarcampsofertes
  Combocara1 = descripciocara(cadbl(materials.Recordset!codidescmatcara1))
  Combocara2 = descripciocara(cadbl(materials.Recordset!codidescmatcara2))
  CombotipusCQ.ListIndex = -1
  cdataanalisis = ""
  If IsDate(materials.Recordset!dataanalisissubstancies) Then cdataanalisis = Format(materials.Recordset!dataanalisissubstancies, "dd/mm/yy")
  If atrim(materials.Recordset!tipusCQ) = "L" Then CombotipusCQ.ListIndex = 1
  If atrim(materials.Recordset!tipusCQ) = "C" Then CombotipusCQ.ListIndex = 0
  If atrim(materials.Recordset!tipusCQ) = "N" Then CombotipusCQ.ListIndex = 2
End Sub
Function descripciocara(vcodi As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select descripcio from tractamentcares where codi=" + atrim(vcodi))
   If Not rst.EOF Then
       descripciocara = rst!descripcio
   End If
   Set rst = Nothing
End Function
Sub netejarcampsofertes()
   dataoferta.Clear
   dataoferta = ""
   cvigencia = ""
   cnumoferta = ""
   ckgminim = ""
   llistamicres.Clear
   datapreukg.RecordSource = "select * from tarifesproveidorsescandall where 1=2"
   datapreukg.Refresh
   
End Sub
Sub descfamilies()
  Dim rstp As Recordset
  Dim rstp2 As Recordset
  Dim l As String
  'families materials
  Set rstp = materials.Database.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(cadbl(Text1(4))))
  Set rstp2 = materials.Database.OpenRecordset("select descripcio from subfamiliesmaterials where codi=" + atrim(cadbl(Text1(7))))
  If Not rstp.EOF Then
     l = atrim(rstp!descripcio)
     If Not rstp2.EOF Then l = l + " - " + atrim(rstp2!descripcio)
     Label3(0) = l
  End If
  
  'families colorants
  Set rstp = materials.Database.OpenRecordset("select descripcio from familiescolorants where codi=" + atrim(cadbl(Text1(5))))
  Set rstp2 = materials.Database.OpenRecordset("select descripcio from subfamiliescolorants where codi=" + atrim(cadbl(Text1(8))))
  If Not rstp.EOF Then
     l = atrim(rstp!descripcio)
     If Not rstp2.EOF Then l = l + " - " + atrim(rstp2!descripcio)
     Label3(1) = l
  End If
  
  'families aditius
  Set rstp = materials.Database.OpenRecordset("select descripcio from familiesaditius where codi=" + atrim(cadbl(Text1(6))))
  Set rstp2 = materials.Database.OpenRecordset("select descripcio from subfamiliesaditius where codi=" + atrim(cadbl(Text1(9))))
  If Not rstp.EOF Then
     l = atrim(rstp!descripcio)
     If Not rstp2.EOF Then l = l + " - " + atrim(rstp2!descripcio)
     Label3(2) = l
  End If
  
  'subfamilia compatible material especific
  combosubfamcompatible = ""
  Set rstp = materials.Database.OpenRecordset("select descripcio from subfamiliesmaterials where codi=" + atrim(cadbl(cmatcompatible)))
  If Not rstp.EOF Then
     l = atrim(rstp!descripcio)
     combosubfamcompatible = l
  End If
  
  Set rstp = Nothing
  Set rstp2 = Nothing
End Sub

Private Sub materials_Validate(Action As Integer, Save As Integer)
   If Save = -1 And cadbl(Text1(10)) = 0 And cadbl(Text1(14)) > 0 And mesespcompra <> "Grm/m2" Then MsgBox "Per poder gravar els canvis has d'entrar un valor a Grm/m3.", vbCritical, "Atenció": Save = False
    If Save = -1 And mesespcompra = "Micres" And Combocolorrec = "" Then MsgBox "Per poder gravar els canvis has d'entrar el color del material per RECICLAR.", vbCritical, "Atenció": Save = False
    If Not Save And materials.Recordset.EditMode > 0 Then materials.Recordset.CancelUpdate
End Sub

Private Sub mesespcompra_LostFocus()
   If mesespcompra = "Grm/m2" And cadbl(Text1(12)) = 0 Then MsgBox "Pensa a entrar l'equivalencia en micres al camp d'Espessor", vbInformation, "Atenció"
End Sub

Private Sub modificar_Click()
   If Not materials.Recordset.EOF Then
     materials.Recordset.Edit
     framematerials.Enabled = True
     Text1(2).SetFocus
   End If
End Sub

Private Sub sortir_Click()
 Unload fmaterials
End Sub

Private Sub Timer1_Timer()

End Sub
Sub triar_proveidor()
  Load formseleccio
  formseleccio.Data1.DatabaseName = materials.DatabaseName
  formseleccio.Data1.RecordSource = "select * from proveidors"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text1(1).Text = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   materials.Recordset!proveidor = Text1(1).Text
   nomproveidor.Caption = atrim(formseleccio.Data1.Recordset!nom)
  End If
  Unload formseleccio
  
End Sub



Private Sub subfamad_DropDown()
carregar_subfamilies
End Sub

Private Sub subfamcol_DropDown()
carregar_subfamilies
End Sub

Private Sub subfammat_DropDown()
carregar_subfamilies
End Sub

Private Sub Text1_Change(Index As Integer)
   If Index = 12 And materials.Recordset.EditMode > 0 Then If cadbl(Text1(12)) = 0 Then Text1(12) = "0"
End Sub

Private Sub Text1_GotFocus(Index As Integer)
  If materials.Recordset.Fields(Text1(Index).DataField).Type = 10 Then
     Text1(Index).MaxLength = materials.Recordset.Fields(Text1(Index).DataField).Size
  End If
  If Index = 0 Then
    If espesors.RecordSource = "" Then Exit Sub
    If Not espesors.Recordset.EOF Then
       Text1(1).SetFocus
       Text1(0).Locked = True
       MsgBox "No pots editar aquest camp si hi ha micres asignades"
      Else: Text1(0).Locked = False
    End If
  End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then
      Select Case Index
        Case 1
           triar_proveidor
        Case 4
          Text1(Index) = triar_familia("select * from familiesmaterials where codi>499")
        Case 5
          Text1(Index) = triar_familia("select * from familiescolorants where codi>499")
        Case 6
          Text1(Index) = triar_familia("select * from familiesaditius where codi>499")
        Case 7
           Text1(Index) = triar_familia("select * from subfamiliesmaterials where codifam=" + atrim(cadbl(Text1(4))))
        Case 8
           Text1(Index) = triar_familia("select * from subfamiliescolorants where codifam=" + atrim(cadbl(Text1(5))))
        Case 9
           Text1(Index) = triar_familia("select * from subfamiliesaditius where codifam=" + atrim(cadbl(Text1(6))))
      End Select
      descfamilies
  End If
  
End Sub
Function triar_familia(seleccio As String) As String
   Load formseleccio
   formseleccio.Caption = "Triar Familia o Subfamilia"
   formseleccio.Data1.DatabaseName = materials.DatabaseName
   formseleccio.Data1.RecordSource = seleccio
   formseleccio.refrescar
   formseleccio.Show 1
   If seleccioret = 1 Then
     triar_familia = atrim(cadbl(formseleccio.Data1.Recordset!codi))
      Else: triar_familia = "0"
   End If
  Unload formseleccio
End Function
Sub comprovarmesde500(Index As Integer)
   If cadbl(Text1(Index)) = 0 Then Exit Sub
   Select Case Index
        Case 4
          If cadbl(Text1(Index)) < 500 Then Text1(Index) = "0": MsgBox "La familia ha de ser superior a 500"
        Case 5
          If cadbl(Text1(Index)) < 500 Then Text1(Index) = "0": MsgBox "La familia ha de ser superior a 500"
        Case 6
          If cadbl(Text1(Index)) < 500 Then Text1(Index) = "0": MsgBox "La familia ha de ser superior a 500"
   End Select
   
   
End Sub
Private Sub Text1_LostFocus(Index As Integer)

  comprovarmesde500 Index

  If Index = 4 Or Index = 5 Or Index = 6 Or Index = 7 Or Index = 8 Or Index = 9 Then
      descfamilies
  End If
  If Index = 11 Then
      If cadbl(Text1(11)) > 0 And cadbl(Text1(10)) > 0 Then
         If MsgBox("Hi ha un valor entrat a Grm/cm3 que no pot coexistir amb els Grm/m2." + Chr(10) + Chr(13) + "Vols eliminar els Grm/cm3?", vbYesNo, "Atenció") = vbYes Then
           Text1(10) = 0
             Else: Text1(11) = 0
         End If
      End If
  End If
  If Index = 10 Then
      If cadbl(Text1(10)) > 0 And cadbl(Text1(11)) > 0 Then
        If MsgBox("Hi ha un valor entrat a Grm/m2 que no pot coexistir amb els Grm/cm3." + Chr(10) + Chr(13) + "Vols eliminar els Grm/m2?", vbYesNo, "Atenció") = vbYes Then
           Text1(11) = 0
             Else:
                If MsgBox("Vols conservar igualment els Grm/cm3?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Text1(10) = 0
         End If
      End If
  End If
  If Index = 14 Then
     Text1(14) = cadbl(Text1(14))
     If cadbl(Text1(14)) > 100 Then MsgBox "Aquest camp no pot ser mes de 100%", vbCritical, "Error": Text1(14) = "100": Exit Sub
     MsgBox "Si poses aqui el (" + Text1(14) + "% es de material VERGE." + vbNewLine + "Per tan el " + atrim(100 - cadbl(Text1(14))) + "% restant es de material RECICLAT." + vbNewLine + "ES CORRECTE?", vbExclamation + vbDefaultButton2 + vbYesNo, "A T E N C I Ó"
  End If
End Sub

Private Sub Text2_LostFocus()
   Text2 = atrim(cadbl(Text2))
End Sub

Private Sub Text3_LostFocus()
Text3 = atrim(cadbl(Text3))
End Sub

Private Sub Text4_LostFocus()
Text4 = atrim(cadbl(Text4))
End Sub

Private Sub Text5_LostFocus()
Text5 = atrim(cadbl(Text5))
End Sub

Private Sub Text7_LostFocus()
'   On Error Resume Next
'   If Not IsDate(Text7) Then MsgBox "Data incorrecte": Text7.SetFocus
End Sub

Private Sub Text8_Change()
  
End Sub

Sub carregar_combo_families()
  Dim rstfam As Recordset
  
  Set rstfam = dbtmp.OpenRecordset("select * from familiesmaterials where codi>499")
  fammat.Clear
  While Not rstfam.EOF
    fammat.AddItem atrim(rstfam!descripcio)
    fammat.ItemData(fammat.NewIndex) = cadbl(rstfam!codi)
    rstfam.MoveNext
  Wend
  Set rstfam = dbtmp.OpenRecordset("select * from familiescolorants where codi>499")
  famcol.Clear
  While Not rstfam.EOF
    famcol.AddItem atrim(rstfam!descripcio)
    famcol.ItemData(famcol.NewIndex) = cadbl(rstfam!codi)
    rstfam.MoveNext
  Wend
  Set rstfam = dbtmp.OpenRecordset("select * from familiesaditius where codi>499")
  famad.Clear
  While Not rstfam.EOF
    famad.AddItem atrim(rstfam!descripcio)
    famad.ItemData(famad.NewIndex) = cadbl(rstfam!codi)
    rstfam.MoveNext
  Wend
End Sub
Sub carregar_subfamilies(Optional combof As Control)
  Dim rstsub As Recordset
  Dim combo As Control
  Dim subfamilia As String
  
  Set combo = fmaterials.ActiveControl
  If Not combof Is Nothing Then Set combo = combof
  If fmaterials.Controls(combo.Tag).ListIndex = -1 And combof Is Nothing Then MsgBox "Primer has d'escullir la familia": Exit Sub
  'If combo.ListIndex = -1 Then combo.Clear: Exit Sub
  If combo.Name = "subfammat" And fammat.ListIndex <> -1 Then r = " codifam=" + atrim(cadbl(fammat.ItemData(fammat.ListIndex))): subfamilia = "subfamiliesmaterials"
  If combo.Name = "subfamcol" And famcol.ListIndex <> -1 Then r = " codifam=" + atrim(cadbl(famcol.ItemData(famcol.ListIndex))): subfamilia = "subfamiliescolorants"
  If combo.Name = "subfamad" And famad.ListIndex <> -1 Then r = " codifam=" + atrim(cadbl(famad.ItemData(famad.ListIndex))): subfamilia = "subfamiliesaditius"
    combo.Clear

  If subfamilia <> "" Then
     Set rstsub = dbtmp.OpenRecordset("select codi,descripcio from " + subfamilia + " where " + r) '+ " and descripcio like '*" + treure_apostrof(subfammat.Text) + "*'")
    Else: Exit Sub
  End If
  
  While Not rstsub.EOF
    combo.AddItem atrim(rstsub!descripcio)
    combo.ItemData(combo.NewIndex) = cadbl(rstsub!codi)
    rstsub.MoveNext
  Wend
  
  
End Sub
