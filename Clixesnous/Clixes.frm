VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form formclixes 
   Caption         =   "Manteniment de Clixes"
   ClientHeight    =   10830
   ClientLeft      =   6675
   ClientTop       =   4095
   ClientWidth     =   15360
   Icon            =   "Clixes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10830
   ScaleWidth      =   15360
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   0
      Top             =   690
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   1275
      Left            =   4485
      TabIndex        =   122
      Top             =   9510
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Passar els noms dels colors dels treballs a codis de tinta"
      Height          =   480
      Left            =   9720
      TabIndex        =   121
      Top             =   10155
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Convertir nom dels colors de les tintes"
      Height          =   495
      Left            =   12585
      TabIndex        =   119
      Top             =   10110
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.CommandButton arxiudocumentaciorelacionada 
      BackColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   14445
      OLEDropMode     =   1  'Manual
      Picture         =   "Clixes.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   103
      ToolTipText     =   "Documents i mails relacionats amb el treball."
      Top             =   5985
      Width           =   705
   End
   Begin VB.CommandButton sortir 
      Height          =   390
      Left            =   14775
      Picture         =   "Clixes.frx":1834
      Style           =   1  'Graphical
      TabIndex        =   100
      ToolTipText     =   "Sortir"
      Top             =   120
      Width           =   390
   End
   Begin VB.ListBox opcionscombo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF1F1&
      Height          =   225
      Left            =   1440
      TabIndex        =   99
      Top             =   1815
      Visible         =   0   'False
      Width           =   6045
   End
   Begin VB.ListBox llistadecomandespendents 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF1F1&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   12540
      TabIndex        =   98
      Top             =   1380
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton botopressupost 
      Caption         =   "Pressupost"
      Height          =   570
      Left            =   6405
      OLEDropMode     =   1  'Manual
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   6870
      Width           =   1320
   End
   Begin VB.CommandButton comandespendents 
      BackColor       =   &H000000FF&
      Caption         =   "Comandes pendents..."
      Height          =   330
      Left            =   11730
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   1050
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.CommandButton botopdf 
      DisabledPicture =   "Clixes.frx":1DBE
      DownPicture     =   "Clixes.frx":33A8
      Height          =   690
      Left            =   14100
      OLEDropMode     =   1  'Manual
      Picture         =   "Clixes.frx":4992
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   2595
      Width           =   780
   End
   Begin VB.CommandButton Command11 
      Height          =   360
      Left            =   1980
      Picture         =   "Clixes.frx":5F7C
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Alta  modificacio"
      Top             =   2265
      Width           =   420
   End
   Begin VB.CommandButton Command10 
      Height          =   360
      Left            =   2415
      Picture         =   "Clixes.frx":6506
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Eliminacio de la modificacio"
      Top             =   2265
      Width           =   420
   End
   Begin VB.Frame framehistorialmodificacions 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Historial de Modificacions"
      Height          =   1770
      Left            =   105
      TabIndex        =   16
      Top             =   7560
      Width           =   15090
      Begin VB.Data modificacions 
         Caption         =   "modificacions"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5715
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Modificacions"
         Top             =   1305
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSDBGrid.DBGrid reixamodificacions 
         Bindings        =   "Clixes.frx":6A90
         Height          =   1500
         Left            =   90
         OleObjectBlob   =   "Clixes.frx":6AA8
         TabIndex        =   17
         Top             =   210
         Width           =   14760
      End
   End
   Begin VB.CommandButton botorepasclixes 
      Caption         =   "Repas Clixés"
      Height          =   555
      Left            =   9375
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6855
      Width           =   1485
   End
   Begin VB.CommandButton botocomandesfotogravador 
      Caption         =   "Comandes Fotogravador"
      Height          =   555
      Left            =   7785
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6870
      Width           =   1485
   End
   Begin VB.CommandButton bototintes 
      Caption         =   "BotóTintes"
      Height          =   555
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6885
      Width           =   1485
   End
   Begin VB.CommandButton botoclientsvinculats 
      Caption         =   " Clients Vinculats"
      Height          =   555
      Left            =   3270
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6885
      Width           =   1485
   End
   Begin VB.CommandButton botoliniesalbarans 
      Caption         =   "Linies Albarans"
      Height          =   555
      Left            =   1650
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6885
      Width           =   1485
   End
   Begin VB.CommandButton botoliniesmodificacions 
      Caption         =   "Linies Modificacions"
      Height          =   555
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6900
      Width           =   1485
   End
   Begin VB.Frame framemodificacions 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Dades de la Modificació                     "
      Enabled         =   0   'False
      Height          =   4530
      Left            =   75
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Top             =   2265
      Width           =   15225
      Begin VB.CommandButton Command19 
         Caption         =   "Observacions repàs clixes"
         Height          =   255
         Left            =   7710
         TabIndex        =   139
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox cobservaciorepasclixes 
         DataField       =   "observacionsrepasclixes"
         DataSource      =   "modificacions"
         Height          =   600
         Left            =   7680
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   138
         Top             =   3855
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.ComboBox Combo6 
         DataField       =   "digimarc"
         DataSource      =   "modificacions"
         Height          =   315
         ItemData        =   "Clixes.frx":822A
         Left            =   10245
         List            =   "Clixes.frx":8234
         TabIndex        =   137
         Top             =   765
         Width           =   645
      End
      Begin VB.TextBox Text13 
         DataField       =   "valordeltamaxim"
         DataSource      =   "modificacions"
         Height          =   285
         Left            =   7365
         TabIndex        =   134
         Top             =   2895
         Width           =   525
      End
      Begin VB.CommandButton Command18 
         Height          =   330
         Left            =   1020
         Picture         =   "Clixes.frx":8240
         Style           =   1  'Graphical
         TabIndex        =   132
         ToolTipText     =   "Assignar linia d'impresió"
         Top             =   750
         Width           =   570
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00E0E0E0&
         DataField       =   "numerodelinia"
         DataSource      =   "modificacions"
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   131
         Top             =   750
         Width           =   345
      End
      Begin VB.CommandButton btestcolor 
         BackColor       =   &H00FFFFFF&
         Height          =   705
         Left            =   13605
         OLEDropMode     =   1  'Manual
         Picture         =   "Clixes.frx":87CA
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Test de color del Fotogravador"
         Top             =   3720
         Width           =   705
      End
      Begin VB.TextBox Text7 
         DataField       =   "amplesang"
         DataSource      =   "modificacions"
         Height          =   285
         Left            =   8595
         TabIndex        =   117
         Top             =   2235
         Width           =   510
      End
      Begin VB.ComboBox comboimpresiocentrada 
         DataField       =   "impresiocentrada"
         DataSource      =   "modificacions"
         Height          =   315
         ItemData        =   "Clixes.frx":A7B0
         Left            =   13785
         List            =   "Clixes.frx":A7BA
         TabIndex        =   115
         Top             =   2265
         Width           =   780
      End
      Begin VB.TextBox Text6 
         DataField       =   "amplebandaseguiment"
         DataSource      =   "modificacions"
         Height          =   285
         Left            =   12780
         TabIndex        =   113
         Top             =   2265
         Width           =   510
      End
      Begin VB.ComboBox Combo5 
         DataField       =   "bandaseguiment"
         DataSource      =   "modificacions"
         Height          =   315
         ItemData        =   "Clixes.frx":A7C6
         Left            =   11250
         List            =   "Clixes.frx":A7D3
         TabIndex        =   111
         Top             =   2250
         Width           =   1515
      End
      Begin VB.ComboBox Combo4 
         DataField       =   "macula"
         DataSource      =   "modificacions"
         Height          =   315
         ItemData        =   "Clixes.frx":A7F2
         Left            =   9480
         List            =   "Clixes.frx":A805
         TabIndex        =   109
         Top             =   2235
         Width           =   1515
      End
      Begin VB.ComboBox Combo3 
         DataField       =   "portasang"
         DataSource      =   "modificacions"
         Height          =   315
         ItemData        =   "Clixes.frx":A838
         Left            =   7065
         List            =   "Clixes.frx":A845
         TabIndex        =   107
         Top             =   2205
         Width           =   1515
      End
      Begin VB.TextBox Text10 
         DataField       =   "bandes"
         DataSource      =   "modificacions"
         Height          =   285
         Left            =   5385
         TabIndex        =   92
         Top             =   795
         Width           =   405
      End
      Begin VB.TextBox Text11 
         DataField       =   "gruixpolimer"
         DataSource      =   "modificacions"
         Height          =   285
         Left            =   5985
         TabIndex        =   91
         Top             =   795
         Width           =   750
      End
      Begin VB.TextBox camplelamina 
         DataField       =   "amplelamina"
         DataSource      =   "modificacions"
         Height          =   285
         Left            =   6900
         TabIndex        =   90
         Top             =   795
         Width           =   690
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H00C0C0C0&
         DataField       =   "tinters"
         DataSource      =   "modificacions"
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   795
         Width           =   510
      End
      Begin VB.ComboBox sistemaimpresio 
         DataField       =   "sistemadimpresio"
         DataSource      =   "modificacions"
         Height          =   315
         ItemData        =   "Clixes.frx":A86A
         Left            =   8880
         List            =   "Clixes.frx":A886
         TabIndex        =   88
         Top             =   780
         Width           =   1305
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0C0C0&
         DataField       =   "desarroll"
         DataSource      =   "modificacions"
         Enabled         =   0   'False
         Height          =   285
         Left            =   7725
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   795
         Width           =   510
      End
      Begin VB.TextBox Text1 
         DataField       =   "observacions"
         DataSource      =   "modificacions"
         Height          =   285
         Left            =   195
         MaxLength       =   250
         TabIndex        =   75
         Top             =   3870
         Width           =   7440
      End
      Begin VB.Frame framededates 
         BackColor       =   &H00EEE4D7&
         Height          =   555
         Left            =   6495
         TabIndex        =   68
         Top             =   1800
         Visible         =   0   'False
         Width           =   1350
         Begin VB.CommandButton bdatapdf 
            BackColor       =   &H008080FF&
            Caption         =   "Data Pdf"
            Height          =   195
            Left            =   15
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   120
            Width           =   1290
         End
         Begin VB.CommandButton bdatacromalin 
            BackColor       =   &H0080FFFF&
            Caption         =   "Data Cromalin"
            Height          =   195
            Left            =   15
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   315
            Width           =   1290
         End
      End
      Begin VB.TextBox cdatacromalin 
         DataField       =   "dataprovacolor"
         DataSource      =   "modificacions"
         Height          =   285
         Left            =   13455
         TabIndex        =   56
         Top             =   1395
         Width           =   1305
      End
      Begin VB.CommandButton Command9 
         Height          =   285
         Left            =   14805
         Picture         =   "Clixes.frx":A8E0
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         ToolTipText     =   "Borrar la data de Baixa del Treball"
         Top             =   1395
         Width           =   270
      End
      Begin VB.ComboBox comboformaimpresio 
         Height          =   315
         ItemData        =   "Clixes.frx":AE6A
         Left            =   3750
         List            =   "Clixes.frx":AE74
         TabIndex        =   60
         Top             =   795
         Width           =   1515
      End
      Begin VB.ComboBox nomproveidor 
         Height          =   315
         Left            =   9030
         Locked          =   -1  'True
         TabIndex        =   55
         Tag             =   "proveidor"
         Top             =   1410
         Width           =   2895
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EEE4D7&
         Caption         =   "Canvis a comprovar en el clixé"
         Height          =   1515
         Left            =   165
         TabIndex        =   44
         Top             =   1935
         Width           =   6660
         Begin VB.ComboBox Combo2 
            DataField       =   "textevalidaciocolors"
            DataSource      =   "modificacions"
            Height          =   315
            ItemData        =   "Clixes.frx":AE8F
            Left            =   690
            List            =   "Clixes.frx":AE9C
            TabIndex        =   73
            Tag             =   "proveidor"
            Top             =   1020
            Width           =   3975
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "textevalidaciomides"
            DataSource      =   "modificacions"
            Height          =   315
            ItemData        =   "Clixes.frx":AEC3
            Left            =   690
            List            =   "Clixes.frx":AED0
            TabIndex        =   72
            Tag             =   "proveidor"
            Top             =   705
            Width           =   3975
         End
         Begin VB.ComboBox ctextevalidacio 
            DataField       =   "textevalidaciotexte"
            DataSource      =   "modificacions"
            Height          =   315
            ItemData        =   "Clixes.frx":AEEC
            Left            =   690
            List            =   "Clixes.frx":AEF9
            TabIndex        =   71
            Tag             =   "proveidor"
            Top             =   375
            Width           =   3975
         End
         Begin VB.CommandButton Command8 
            Height          =   285
            Left            =   6045
            Picture         =   "Clixes.frx":AF15
            Style           =   1  'Graphical
            TabIndex        =   54
            TabStop         =   0   'False
            ToolTipText     =   "Borrar la data"
            Top             =   1020
            Width           =   270
         End
         Begin VB.TextBox cdatacolors 
            DataField       =   "datavalidaciocolors"
            DataSource      =   "modificacions"
            Height          =   285
            Left            =   4710
            TabIndex        =   53
            Top             =   1020
            Width           =   1305
         End
         Begin VB.CommandButton Command7 
            Height          =   285
            Left            =   6045
            Picture         =   "Clixes.frx":B49F
            Style           =   1  'Graphical
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "Borrar la data"
            Top             =   705
            Width           =   270
         End
         Begin VB.TextBox cdatamides 
            DataField       =   "datavalidaciomides"
            DataSource      =   "modificacions"
            Height          =   285
            Left            =   4710
            TabIndex        =   51
            Top             =   705
            Width           =   1305
         End
         Begin VB.CommandButton Command6 
            Height          =   285
            Left            =   6060
            Picture         =   "Clixes.frx":BA29
            Style           =   1  'Graphical
            TabIndex        =   50
            TabStop         =   0   'False
            ToolTipText     =   "Borrar la data"
            Top             =   375
            Width           =   270
         End
         Begin VB.TextBox cdatatextes 
            DataField       =   "datavalidaciotexte"
            DataSource      =   "modificacions"
            Height          =   285
            Left            =   4725
            TabIndex        =   48
            Top             =   375
            Width           =   1305
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Control"
            Height          =   360
            Left            =   4860
            TabIndex        =   49
            Top             =   135
            Width           =   1125
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Colors"
            Height          =   225
            Left            =   75
            TabIndex        =   47
            Top             =   1050
            Width           =   675
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Mides"
            Height          =   225
            Left            =   90
            TabIndex        =   46
            Top             =   697
            Width           =   675
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Textes"
            Height          =   225
            Left            =   105
            TabIndex        =   45
            Top             =   345
            Width           =   675
         End
      End
      Begin VB.TextBox Text9 
         DataField       =   "descripcio"
         DataSource      =   "modificacions"
         Height          =   285
         Left            =   165
         TabIndex        =   41
         Top             =   1410
         Width           =   7440
      End
      Begin VB.CommandButton Command5 
         Height          =   285
         Left            =   3375
         Picture         =   "Clixes.frx":BFB3
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Borrar la data de Baixa del Treball"
         Top             =   780
         Width           =   270
      End
      Begin VB.TextBox dataobertura 
         DataField       =   "dataobertura"
         DataSource      =   "modificacions"
         Height          =   285
         Left            =   2265
         TabIndex        =   38
         Top             =   780
         Width           =   1050
      End
      Begin VB.TextBox ordre 
         BackColor       =   &H00C0C0C0&
         DataField       =   "ordre"
         DataSource      =   "modificacions"
         Height          =   285
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox formaimpresio 
         DataField       =   "formaimpresio"
         DataSource      =   "modificacions"
         Height          =   285
         Left            =   3405
         TabIndex        =   34
         Top             =   75
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "     PCC                                     (Punt de control crític)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   9660
         TabIndex        =   141
         Top             =   2805
         Width           =   3570
      End
      Begin VB.Label etcodidelinia 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005C31DD&
         Height          =   315
         Left            =   180
         TabIndex        =   140
         Top             =   1125
         Width           =   1980
      End
      Begin VB.Label lerrorpressupost 
         BackStyle       =   0  'Transparent
         Caption         =   "ATENCIÓ, ELS VALORS DE PRESSUPOST I TREBALL NO COINCIDEIXEN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   105
         TabIndex        =   127
         Top             =   4215
         Visible         =   0   'False
         Width           =   7605
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "DigiMarc"
         Height          =   285
         Left            =   10230
         TabIndex        =   136
         Top             =   555
         Width           =   1290
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor delta Max:"
         Height          =   225
         Left            =   7080
         TabIndex        =   135
         Top             =   2655
         Width           =   1200
      End
      Begin VB.Label etdesarrollerroni 
         BackStyle       =   0  'Transparent
         Caption         =   "Avisar al client desarroll diferent que el de comanda."
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
         Height          =   285
         Left            =   7425
         TabIndex        =   133
         Top             =   1110
         Visible         =   0   'False
         Width           =   6000
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Línia d'Impresió"
         Height          =   225
         Left            =   390
         TabIndex        =   130
         Top             =   555
         Width           =   1725
      End
      Begin VB.Label etreprint 
         BackStyle       =   0  'Transparent
         Caption         =   "Reprint"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   13320
         TabIndex        =   129
         Top             =   3135
         Width           =   1860
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Impresió Centrada"
         Height          =   225
         Left            =   13560
         TabIndex        =   116
         Top             =   2025
         Width           =   1440
      End
      Begin VB.Label materialultimacomanda 
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3150
         TabIndex        =   114
         Top             =   315
         Width           =   8400
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Banda Seguiment i ample en mm"
         Height          =   225
         Left            =   11100
         TabIndex        =   112
         Top             =   2010
         Width           =   2430
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Màcula:"
         Height          =   225
         Left            =   9705
         TabIndex        =   110
         Top             =   1995
         Width           =   1200
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Porta Sang?   i ample en mm"
         Height          =   225
         Left            =   7290
         TabIndex        =   108
         Top             =   1980
         Width           =   2640
      End
      Begin VB.Label etestatclixemod 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CLIXES ENTRATS"
         DataField       =   "estatclixe"
         DataSource      =   "clixes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   10410
         TabIndex        =   101
         Top             =   90
         Width           =   4755
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Gruix Polimer"
         Height          =   225
         Left            =   5835
         TabIndex        =   97
         Top             =   555
         Width           =   1200
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Ample Lam"
         Height          =   225
         Left            =   6840
         TabIndex        =   96
         Top             =   555
         Width           =   870
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Tinters"
         Height          =   225
         Left            =   8310
         TabIndex        =   95
         Top             =   555
         Width           =   705
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema d'impresió"
         Height          =   285
         Left            =   8850
         TabIndex        =   94
         Top             =   555
         Width           =   1575
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Desaroll"
         Height          =   225
         Left            =   7695
         TabIndex        =   93
         Top             =   555
         Width           =   705
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Observacions:"
         Height          =   270
         Left            =   615
         TabIndex        =   76
         Top             =   3630
         Width           =   2325
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Cromalin:"
         Height          =   360
         Left            =   12360
         TabIndex        =   67
         Top             =   1425
         Width           =   1170
      End
      Begin VB.Label etdatapdf 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "99/99/99"
         Height          =   240
         Left            =   14085
         TabIndex        =   65
         ToolTipText     =   "Data de validesa del PDF"
         Top             =   1215
         Width           =   765
      End
      Begin VB.Label missatgenoupdf 
         BackStyle       =   0  'Transparent
         Caption         =   "Arrastra un pdf a la pantalla per linkar-lo--->"
         ForeColor       =   &H00008080&
         Height          =   255
         Left            =   10935
         TabIndex        =   63
         Top             =   750
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.Label nomclient 
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2940
         TabIndex        =   62
         Top             =   135
         Width           =   8400
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Fotogravador:"
         Height          =   240
         Left            =   7980
         TabIndex        =   57
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Bandes"
         Height          =   225
         Left            =   5220
         TabIndex        =   43
         Top             =   555
         Width           =   1200
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripció de la Modificació"
         Height          =   270
         Left            =   2235
         TabIndex        =   42
         Top             =   1185
         Width           =   2325
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Obertura Modif."
         Height          =   360
         Left            =   2190
         TabIndex        =   39
         Top             =   555
         Width           =   1875
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Versió"
         Height          =   225
         Left            =   180
         TabIndex        =   37
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Forma Impresió"
         Height          =   225
         Left            =   3885
         TabIndex        =   35
         Top             =   555
         Width           =   1200
      End
   End
   Begin VB.Frame frameclixes 
      BackColor       =   &H00FAF1F1&
      Caption         =   "Capçalera del Clixé "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   60
      TabIndex        =   8
      Top             =   570
      Width           =   15225
      Begin VB.TextBox Text8 
         DataField       =   "descripcioquantitatlinia"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   13785
         TabIndex        =   125
         ToolTipText     =   "Descripció de la quantitat de l'envàs a la Linia"
         Top             =   885
         Width           =   1305
      End
      Begin VB.CheckBox cportareduccio 
         Caption         =   "Porta reducció (Es tapa)"
         Height          =   285
         Left            =   9210
         TabIndex        =   124
         Top             =   1275
         Width           =   2160
      End
      Begin VB.CommandButton bxl 
         Height          =   255
         Left            =   1890
         Picture         =   "Clixes.frx":C53D
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   555
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CommandButton imprimirbossaclixe 
         Height          =   360
         Left            =   2220
         Picture         =   "Clixes.frx":CAC7
         Style           =   1  'Graphical
         TabIndex        =   120
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir l'etiqueta de la bossa del Clixé"
         Top             =   450
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Height          =   330
         Left            =   2610
         Picture         =   "Clixes.frx":D051
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Modificar Registres"
         Top             =   1260
         Width           =   405
      End
      Begin VB.TextBox reducciopermetre 
         DataField       =   "reduccioxmetre"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   1875
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   104
         Top             =   1275
         Width           =   735
      End
      Begin VB.TextBox reducciocilindref2 
         DataField       =   "redcilindref2"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   8145
         MaxLength       =   25
         TabIndex        =   84
         Top             =   1275
         Width           =   735
      End
      Begin VB.TextBox reducciocilindrefw 
         DataField       =   "redcilindrefw"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   5505
         MaxLength       =   25
         TabIndex        =   83
         Top             =   1290
         Width           =   705
      End
      Begin VB.ComboBox nomclienttemporal 
         DataField       =   "nomclienttemporal"
         DataSource      =   "clixes"
         Height          =   315
         Left            =   8445
         Locked          =   -1  'True
         TabIndex        =   27
         Tag             =   "proveidor"
         Top             =   525
         Width           =   3135
      End
      Begin VB.CommandButton Command4 
         Height          =   285
         Left            =   7995
         Picture         =   "Clixes.frx":D5DB
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Borrar la data de Baixa del Treball"
         Top             =   540
         Width           =   270
      End
      Begin VB.ComboBox liniaproducte 
         DataField       =   "linia"
         DataSource      =   "clixes"
         Height          =   315
         Left            =   7335
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   900
         Width           =   5655
      End
      Begin VB.ComboBox marcaproducte 
         DataField       =   "marca"
         DataSource      =   "clixes"
         Height          =   315
         Left            =   735
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   915
         Width           =   5865
      End
      Begin VB.TextBox Text3 
         DataField       =   "databaixaclixe"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   6660
         TabIndex        =   26
         Top             =   540
         Width           =   1305
      End
      Begin VB.TextBox codidebarres 
         BackColor       =   &H0080C0FF&
         DataField       =   "codidebarres"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   3630
         TabIndex        =   24
         Top             =   540
         Width           =   2325
      End
      Begin VB.TextBox Text4 
         DataField       =   "ubicacio"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   2715
         TabIndex        =   22
         Top             =   540
         Width           =   885
      End
      Begin VB.TextBox Text2 
         DataField       =   "arxiu"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   1305
         TabIndex        =   20
         Top             =   540
         Width           =   885
      End
      Begin VB.TextBox campid_treball 
         BackColor       =   &H00C0C0C0&
         DataField       =   "id_treball"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   540
         Width           =   885
      End
      Begin VB.Label Label32 
         Caption         =   "Quantitat:"
         Height          =   180
         Left            =   13050
         TabIndex        =   126
         Top             =   930
         Width           =   915
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   105
         Top             =   1245
         Width           =   11460
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Distorsió per metre:"
         Height          =   300
         Left            =   150
         TabIndex        =   105
         Top             =   1320
         Width           =   1785
      End
      Begin VB.Label etestatclixe 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         DataField       =   "estatclixe"
         DataSource      =   "clixes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   9645
         TabIndex        =   64
         Top             =   135
         Width           =   4755
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Reduccio Cilindre   F2:"
         Height          =   300
         Left            =   6420
         TabIndex        =   86
         Top             =   1320
         Width           =   1785
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Factor Cilindre   FW:"
         Height          =   300
         Left            =   3750
         TabIndex        =   85
         Top             =   1320
         Width           =   1785
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Client Temporal"
         Height          =   360
         Left            =   9075
         TabIndex        =   81
         Top             =   300
         Width           =   1875
      End
      Begin VB.Label nomclientclixe 
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3015
         TabIndex        =   74
         Top             =   120
         Width           =   9330
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Linia:"
         Height          =   285
         Left            =   6885
         TabIndex        =   32
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         Height          =   300
         Left            =   210
         TabIndex        =   30
         Top             =   930
         Width           =   660
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Baixa del Treball"
         Height          =   360
         Left            =   6525
         TabIndex        =   28
         Top             =   315
         Width           =   1875
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Codi de Barres"
         Height          =   270
         Left            =   4215
         TabIndex        =   25
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicació"
         Height          =   225
         Left            =   2835
         TabIndex        =   23
         Top             =   315
         Width           =   1200
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Arxiu del Clixé"
         Height          =   225
         Left            =   1200
         TabIndex        =   21
         Top             =   315
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NºTreball"
         Height          =   225
         Left            =   315
         TabIndex        =   19
         Top             =   315
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   -45
      Width           =   15225
      Begin VB.CommandButton botoavisosliniesalbarans 
         Height          =   360
         Left            =   13605
         Picture         =   "Clixes.frx":DB65
         Style           =   1  'Graphical
         TabIndex        =   128
         ToolTipText     =   "Treballs amb clixes entrats sense albarans del fotogravador."
         Top             =   180
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   315
         Left            =   13380
         TabIndex        =   102
         Top             =   150
         Visible         =   0   'False
         Width           =   660
      End
      Begin Crystal.CrystalReport llistat 
         Left            =   4335
         Top             =   180
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton Command14 
         Caption         =   "importar 1 Treball versio anterior"
         Height          =   435
         Left            =   11340
         TabIndex        =   79
         Top             =   135
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.CommandButton Command13 
         Height          =   360
         Left            =   1905
         Picture         =   "Clixes.frx":E0EF
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Buscar un treball"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton Command12 
         Caption         =   "importar treballs versio anterior"
         Height          =   435
         Left            =   9300
         TabIndex        =   78
         Top             =   135
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.CommandButton botoavisos 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   14040
         Picture         =   "Clixes.frx":E679
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Comandes que no tenen numero de treball"
         Top             =   180
         Width           =   420
      End
      Begin VB.Timer timerdrag 
         Enabled         =   0   'False
         Interval        =   900
         Left            =   3135
         Top             =   135
      End
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   1446
         Picture         =   "Clixes.frx":EC03
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Actualitzar Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   532
         Picture         =   "Clixes.frx":F18D
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Modificar Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   989
         Picture         =   "Clixes.frx":F717
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio del treball"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "Clixes.frx":FCA1
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Alta treball"
         Top             =   150
         Width           =   420
      End
      Begin VB.Data clixes 
         Caption         =   "Clixes"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   480
         Left            =   5205
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from Clixes order by id_treball Desc"
         Top             =   105
         Width           =   4080
      End
      Begin VB.CommandButton consultar 
         Height          =   360
         Left            =   2355
         Picture         =   "Clixes.frx":1022B
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Busqueda avançada"
         Top             =   150
         Width           =   420
      End
      Begin VB.Label estatedicio 
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
         Left            =   2955
         TabIndex        =   7
         Top             =   195
         Width           =   2025
      End
   End
   Begin VB.Menu nmanteniment 
      Caption         =   "Manteniment"
      Begin VB.Menu mfotogravadors 
         Caption         =   "Fotogravadors"
      End
      Begin VB.Menu mestatsclixes 
         Caption         =   "Estats Clixes"
      End
      Begin VB.Menu mavisoclixes 
         Caption         =   "Avisos clixes per client"
      End
   End
   Begin VB.Menu mllistats 
      Caption         =   "Llistats"
      Begin VB.Menu mpendentsdefacturar 
         Caption         =   "Pendents de Facturar"
      End
      Begin VB.Menu mclixesimpresio 
         Caption         =   "Clixes amb ultima data d'impresio"
      End
      Begin VB.Menu mclixesperpalet 
         Caption         =   "Llistat de clixes per palet"
      End
      Begin VB.Menu mquanbossesarxiu 
         Caption         =   "Llistat de quantitat de bosses per arxiu"
      End
      Begin VB.Menu mtotallliureslleixes 
         Caption         =   "Llistat total de lliures a les lleixes"
      End
      Begin VB.Menu mclixesfotogravador 
         Caption         =   "Clixes per fotogravador"
      End
      Begin VB.Menu mcomandesafotogravadorspendents 
         Caption         =   "Comandes a fotogravadors pendents"
      End
      Begin VB.Menu mimprimirbossasoldadores 
         Caption         =   "Imprimir bossa de soldadores"
      End
   End
   Begin VB.Menu mestattreballs 
      Caption         =   "Estat treballs"
   End
   Begin VB.Menu submlinies 
      Caption         =   "Linies d'impresió"
      Begin VB.Menu mlinies 
         Caption         =   "Linies d'impresió visuals"
      End
      Begin VB.Menu mliniesperpantone 
         Caption         =   "Linies d'impresió per Pantones"
      End
   End
   Begin VB.Menu mutils 
      Caption         =   "Utilitats"
      Begin VB.Menu mactualitzaciodadesimpresors 
         Caption         =   "Actualitzacio dades dels impresors"
      End
   End
End
Attribute VB_Name = "formclixes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim modificacioautomatica As Boolean
Dim llistatclixesvells As Boolean
Dim baixaclixes As Boolean
Dim imprimirbossasoldadores As Boolean

Private Const COLORBOTOSENSEDADES = &H8000000F
Private Const COLORBOTOAMBDADES = &HFF8080
Function ObtenerLíneaComando(Optional MaxArgs)
    'Declara las variables.
    Dim c, LíneaComando, LonLínComando, ArgIn, i, NúmArgs
    'Ver si MaxArgs está.
    If IsMissing(MaxArgs) Then MaxArgs = 10
    'Crea una matriz del tamaño correcto.
    ReDim ArgArray(MaxArgs)
    NúmArgs = 0: ArgIn = False
    'Obtiene los argumentos de la línea de comandos.
    LíneaComando = Command()
    LonLínComando = Len(LíneaComando)
    'Recorre la línea de comando carácter a carácter
    'a la vez.

For i = 1 To LonLínComando
        c = Mid(LíneaComando, i, 1)
        'Comprueba espacio o tabulación.
        If (c <> " " And c <> vbTab) Then
            'Ningún espacio o tabulación.
            'Comprueba si está en el argumento.
            If Not ArgIn Then
            'Empieza el nuevo argumento.
            'Comprueba para más argumentos.
                If NúmArgs = MaxArgs Then Exit For
                    NúmArgs = NúmArgs + 1
                    ArgIn = True
                End If
            'Agrega el carácter al argumento actual.

ArgArray(NúmArgs) = ArgArray(NúmArgs) + c
        Else
            'Encontró un espacio o tabulador.
            'Establece ArgIn a False.
            ArgIn = False
        End If
    Next i
    'Redimensiona la matriz lo suficiente para contener los argumentos.
    'ReDim Preserve ArgArray(NúmArgs)
    'Devuelve la matriz en nombre de la función.
    ObtenerLíneaComando = ArgArray()
End Function








Private Sub arxiudocumentaciorelacionada_Click()
  Dim rutaarxiu As String
  
   rutaarxiu = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Arxiu_documentacio_relacionada" + "\v" + atrim(ordremodificacio)
   
   obrircarpetaarxiu rutaarxiu, id_treball
End Sub

Private Sub arxiudocumentaciorelacionada_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
   Static inicidragover As Date
   Dim rutaarxiu As String
   If arxiudocumentaciorelacionada.tag = "1" Then Exit Sub
   If inicidragover = "0:00:00" Then inicidragover = Now
   rutaarxiu = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Arxiu_documentacio_relacionada" + "\v" + atrim(ordremodificacio)
   If DateDiff("s", inicidragover, Now) > 4 Then inicidragover = Now
   If DateDiff("s", inicidragover, Now) = 1 Then
      arxiudocumentaciorelacionada.BackColor = &HFFFF&
      inicidragover = DateAdd("s", 1, Now): arxiudocumentaciorelacionada.tag = "1": timerdrag.Enabled = True
      
   End If
End Sub
Sub obrircarpetaarxiu(ruta As String, idt As Long)
  Dim idp As Long
  ratoli "espera"
  DoEvents
  crearruta ruta_documentacio_clixes + "\" + Format(idt, "00000")
  crearruta ruta_documentacio_clixes + "\" + Format(idt, "00000") + "\Arxiu_documentacio_relacionada"
  crearruta ruta
  idp = ShellExecute(Me.hWnd, "Open", "c:\windows\explorer.exe", " " + ruta, "", 1)
  ratoli "normal"
End Sub
Sub crearruta(ruta As String)
   On Error Resume Next
   If Not existeix(ruta) Then MkDir ruta
End Sub


Private Sub bdatacromalin_Click()
  If framededates.tag <> "" Then Me.Controls(framededates.tag) = cdatacromalin
  framededates.visible = False
           framededates.tag = ""
End Sub

Private Sub bdatapdf_Click()
  If framededates.tag <> "" Then Me.Controls(framededates.tag) = etdatapdf
  framededates.visible = False
           framededates.tag = ""
End Sub

Private Sub botoavisos_Click()
   arreglarlesquejashanfet
   ensenyalespendents
   comprovaravisos
End Sub

Private Sub botoeditarlinia_Click()
  botoeditarlinia.tag = "editant"
  liniaproducte.SetFocus
End Sub

Private Sub botoeditarmarca_Click()
  botoeditarmarca.tag = "editant"
  
  marcaproducte.SetFocus
End Sub

Private Sub botoavisosliniesalbarans_Click()
   
   
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "SELECT pressupostos.id_treball, pressupostos.ordremodificacio,Clixes.estatclixe FROM (pressupostos LEFT JOIN Clixes_albarans ON (pressupostos.id_treball = Clixes_albarans.id_treball) AND (pressupostos.ordremodificacio = Clixes_albarans.ordremodificacio)) LEFT JOIN Clixes ON pressupostos.id_treball = Clixes.id_treball WHERE (((Clixes_albarans.num_alb) Is Null) AND ((pressupostos.enviat)=True) AND ((Clixes.estatclixe)='CLIXES ENTRATS'));"
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns("id_treball").width = 1300
   formseleccio.DBGrid2.Columns("ordremodificacio").width = 330
   formseleccio.DBGrid2.Columns("estatclixe").width = 1800
   formseleccio.width = 5000
   formseleccio.DBGrid2.width = 4785
   formseleccio.Show 1
    If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
              clixes.Recordset.FindFirst "id_treball=" + atrim(cadbl(formseleccio.DBGrid2.Columns("Id_treball")))
              modificacions.Recordset.FindFirst "ordre=" + atrim(cadbl(formseleccio.DBGrid2.Columns("ordremodificacio")))
        End If
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   comprovartreballsambpressupostisenseliniesdalbara
End Sub

Private Sub botocomandesfotogravador_Click()
   If Not existeix(formclixes.rutapdftreball) Then MsgBox "Encara no hi ha el PDF d'aquest treball i no podré generar la comanda al proveidor", vbCritical, "Error": Exit Sub
   formcomandaclixes.Show 1
   possarcolorbotocomandafotogravador
   possarestatclixe
End Sub

Private Sub botopdf_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
 If data.GetFormat(15) Then
     inicidragover = DateAdd("s", 1, inicidragover)
     obrirtemporalclixes True
     Copiar_Fitxer data.Files(1), "c:\temp\tmpclixes\dragover.pdf"
     timerdrag.Enabled = True
       Else: inicidragover = 0
 End If
 
End Sub

Private Sub botopdf_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
 If data.GetFormat(1) Then MsgBox "No arrastris el pdf desde el gestor de correu. Primer arrastra'l a l'escriptori i despres al programa, GRACIES.", vbCritical, "Atenció": Exit Sub
 If inicidragover = 0 Or DateDiff("s", inicidragover, Now) > 10 Then inicidragover = Now
 ' If data.GetFormat(1) And DateDiff("s", inicidragover, Now) = 2 Then
 '    inicidragover = DateAdd("s", 1, inicidragover)
 '    obrirtemporalclixes False
 '    timerdrag.Enabled = True
 ' End If
End Sub

Function valorspressupostcorrectes(ample As Double, bandes As Byte, cilindre As Double, desarroll As Double, tinters As Byte) As Boolean
   Dim rst As Recordset
   Dim rstm As Recordset
   valorspressupostcorrectes = True
   Set rstm = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio))
   Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
   If rstm.EOF Or rstm.EOF Then Exit Function
   valorspressupostcorrectes = True
   If bandes = 0 And ample = 0 And desarroll = 0 And tinters = 0 And cilindre = 0 Then Exit Function
   If bandes <> cadbl(rstm!bandes) Then valorspressupostcorrectes = False
   If ample <> cadbl(rstm!amplelamina) Then valorspressupostcorrectes = False
   If desarroll <> cadbl(rstm!desarroll) Then valorspressupostcorrectes = False
   If tinters <> cadbl(rstm!tinters) Then valorspressupostcorrectes = False
   If cilindre <> cadbl(rst!cilindre) Then valorspressupostcorrectes = False
   Set rst = Nothing
   Set rstm = Nothing
End Function


Private Sub botopressupost_Click()

    formpressupost.Show 1
  possarcolorbotopressupost
  revisarpreualbaransipressupost
End Sub

Private Sub btestcolor_Click()
  formtestcolors.Show 1
End Sub

Private Sub bxl_Click()
  Dim vcontrol As Control
  If bxl.tag = "text2" Then Set vcontrol = Text2
  If bxl.tag = "text4" Then Set vcontrol = Text4
  Load formescullirlleixa
  formescullirlleixa.Top = (vcontrol.Top * 3) + formclixes.Top + frameclixes.Top
  formescullirlleixa.Left = vcontrol.Left + formclixes.Left + frameclixes.Left
  formescullirlleixa.Show 1
  If seleccioret = 1 Then formclixes.Controls(bxl.tag) = formescullirlleixa.valorescullit
  bxl.visible = False
End Sub

Private Sub cdatacolors_GotFocus()
 colocarframedates cdatacolors, True
End Sub

Private Sub cdatacolors_LostFocus()
 colocarframedates cdatacolors, False
End Sub

Private Sub cdatamides_GotFocus()
 colocarframedates cdatamides, True
End Sub

Private Sub cdatamides_LostFocus()
 colocarframedates cdatamides, False
End Sub

Private Sub cdatatextes_GotFocus()
   colocarframedates cdatatextes, True
End Sub
Sub colocarframedates(c As Control, visible As Boolean)
   Dim nomproximboto As String
   If visible Then
      framededates.tag = c.Name
      framededates.visible = True
     Else:
        nomproximboto = Screen.ActiveControl.Name
        If nomproximboto <> "bdatapdf" And nomproximboto <> "bdatacromalin" Then
           framededates.visible = False
           framededates.tag = ""
        End If
   End If
   framededates.Top = c.Top + Frame5.Top - framededates.Height
   framededates.Left = c.Left + Frame5.Left
   
   
   
End Sub

Private Sub cdatatextes_LostFocus()

 colocarframedates cdatatextes, False
End Sub

Private Sub codidebarres_LostFocus()
  If hihaalgunsimbolextrany(codidebarres) Then MsgBox "Hi ha algun espai o symbol extrany en aquest codi de barres revisa-ho sisplau o pot donar problemes.", vbCritical, "Atenció"
End Sub
Function hihaalgunsimbolextrany(vcdb) As Boolean
   Dim i As Byte
   Dim vcascii As Long
   For i = 1 To Len(vcdb)
      vcascii = Asc(Mid(vcdb, i, 1))
      If vcascii < 32 Or vcascii > 90 Then hihaalgunsimbolextrany = True
   Next i
End Function
Private Sub comandespendents_Click()
   'If llistadecomandespendents.visible Then llistadecomandespendents.visible = False
   'If llistadecomandespendents.ListCount = 1 Then cridarcomandes llistadecomandespendents.List(0): Exit Sub
   'If llistadecomandespendents.ListCount > 1 Then
   llistadecomandespendents.visible = Not llistadecomandespendents.visible
      
End Sub
Sub cridarcomandes(comanda As Double)
 On Error GoTo obrircomandes
  escriure_ini "Planificacio", "comandaxrobrir", atrim(comanda), "comandes.ini"
  AppActivate "Manteniment de Comandes"
  
  On Error Resume Next
  Exit Sub
obrircomandes:
  On Error Resume Next
   Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", "comandes.ini")) + "comandes.exe - comandes", vbNormalFocus
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Sub emplenar_llistadesemblants()
   Dim rstc As Recordset
   combosemblants.Clear
   'If atrim(liniaproducte) = "" Then opcionscombo.visible = False: Exit Sub
   Set rstc = dbclixes.OpenRecordset("select distinct agrupaciosemblants from modificacions ")
   While Not rstc.EOF
     If atrim(rstc!agrupaciosemblants) <> "" Then combosemblants.AddItem atrim(rstc!agrupaciosemblants)
     rstc.MoveNext
   Wend
End Sub

Private Sub comboimpresiocentrada_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub combosemblants_DropDown()
End Sub

Private Sub Command10_Click()
  Dim resp As String
  resp = UCase(InputBox("Estas segur que vols eliminar aquesta modificació?" + Chr(10) + "AIXÒ TAMBÉ ELIMINARÀ ELS IMPS I PDF'S RELACIONATS I CLIXES COMPARTITS." + Chr(10) + "Escriu [eliminarmodificacio] per eliminar-la.", "Eliminar modificació"))
  If resp = "ELIMINARMODIFICACIO" Then
      eliminarmodificacio cadbl(id_treball), cadbl(ordremodificacio)
      modificacions.Refresh
      'modificacions.Recordset.FindFirst "ordre<" + atrim(ordremodificacio)
  End If
End Sub
Sub eliminarmodificacio(treball As Double, ordre As Integer)
  Dim rutapdf As String
  Dim rutaimp As String
  rutapdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\PDF" + Format(treball, "00000") + "-" + Format(ordre, "000") + ".pdf"
  rutaimp = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\IMP" + Format(treball, "00000") + "-" + Format(ordre, "000") + "-*.doc"
  dbclixes.Execute "delete * from clientsvinculats WHERE id_Treball=" + atrim(treball) + " and ordremodificacio=" + atrim(ordre)
  dbclixes.Execute "delete * from clixes_albarans WHERE id_Treball=" + atrim(treball) + " and ordremodificacio=" + atrim(ordre)
  dbclixes.Execute "delete * from clixes_modifi WHERE id_Treball=" + atrim(treball) + " and ordremodificacio=" + atrim(ordre)
  dbclixes.Execute "delete * from pressupostos WHERE id_Treball=" + atrim(treball) + " and ordremodificacio=" + atrim(ordre)
  mirarsicomparteixclixes_iactualitzarhoabansdeborrar treball, cadbl(ordre)
  dbclixes.Execute "delete * from tintes WHERE id_Treball=" + atrim(treball) + " and ordremodificacio=" + atrim(ordre)
  dbclixes.Execute "delete * from modificacions WHERE id_Treball=" + atrim(treball) + " and ordre=" + atrim(ordre)
  On Error Resume Next
  If existeix(rutapdf) Then Kill rutapdf
  Kill rutaimp
End Sub
Sub mirarsicomparteixclixes_iactualitzarhoabansdeborrar(vtreball As Double, vordre As Double)
  Dim rst As Recordset
  Set rst = dbclixes.OpenRecordset("select * from tintes where  id_treball=" + atrim(vtreball) + " and ordremodificacio=" + atrim(vordre))
  While Not rst.EOF
     dbclixes.Execute "update tintes set tinterlinkambid_treball=" + atrim(cadbl(rst!id_tinter_anterior)) + " where tinterlinkambid_treball=" + atrim(rst!id_tinter)
     rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
Private Sub Command12_Click()
    'per crear o actualitzar les marques i linies
    'crearidsmarcailinia
    importarclixesjacreats
    
End Sub
Function busca_nomdirectori_codiclient(ByVal ruta As String, codiclient As String) As String
On Error Resume Next
ruta = ruta + "\"
minombre = Dir(ruta + codiclient + "*", vbDirectory)  ' Recupera la primera entrada.
Do While minombre <> "" And busca_nomdirectori_codiclient = "" ' Inicia el bucle.
    ' Ignora el directorio actual y el que lo abarca.
    If minombre <> "." And minombre <> ".." Then

' Utiliza comparación a nivel de bits para asegurarse de que MiNombre es un directorio.

    If (GetAttr(ruta & minombre) And vbDirectory) = vbDirectory Then
            If Mid(minombre, 1, 6) = codiclient Then busca_nomdirectori_codiclient = minombre
        End If  ' solamente si representa un directorio.
    End If
    minombre = Dir  ' Obtiene siguiente entrada.
Loop
On Error GoTo 0
End Function
Function rutaifitxerpdf(valorcamppdf As String, codiclient As Long) As String
 Dim ru As String
 If valorcamppdf = "" Then Exit Function
 ru = ""
 If InStr(1, valorcamppdf, "\pdfs") = 0 Then
    ru = busca_nomdirectori_codiclient(ruta_relativa_docs, Format(codiclient, "000000"))
    ru = ru + "\pdfs\"
 End If
 rutaifitxerpdf = ruta_relativa_docs + "\" + ru + valorcamppdf  '+ "" + Chr$(34) + Chr$(34)
'MsgBox r + Chr$(34) + ruta_relativa_docs + "\" + Text4.Text + Chr$(34)

 
End Function
Sub importacio_colocarelpdfalseulloc(rstvells As Recordset)
   Dim origenpdf As String
   origenpdf = rutaifitxerpdf(atrim(rstvells!link_pdf), cadbl(rstvells!id_client))
   If existeix(origenpdf) Then
       Copiar_Fitxer origenpdf, rutapdftreball
       
       modificacions.Recordset.Edit
       modificacions.Recordset!pdfvalid = True
       modificacions.Recordset.Update
   End If
End Sub
Sub importarclixesjacreats()
  Dim rstvells As Recordset
  Dim rstmarca As Recordset
  Dim rstnou As Recordset
  '4411  3984  2786
  Set rstvells = dbclixesvells.OpenRecordset("select * from clixes ")
  rstvells.MoveLast
  rstvells.MoveFirst
  While Not rstvells.EOF
    Set rstnou = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(rstvells!id_treball))
    If Not rstnou.EOF Then GoTo proxim
    importacio_crearnouclixe rstvells
    crear_modificacio_nova
    importacio_dadesnovamodificacio rstvells
    importacio_colocarelpdfalseulloc rstvells
    importacio_tintesiclientsvinculats
    importacio_copiarlesliniesdemodificacioialbarans rstvells
    clixes.Recordset.Bookmark = clixes.Recordset.LastModified
    possarestatclixe
    Me.caption = atrim(rstvells.AbsolutePosition) + " / " + atrim(rstvells.RecordCount)
proxim:
    DoEvents
    rstvells.MoveNext
  Wend

End Sub
Sub importacio_tintesiclientsvinculats()
    Dim ultimacomanda As Double
    Dim rstc As Recordset
    Dim rstt As Recordset
    Dim ultimclient As Double
    Dim clientprincipal As Double
    'amb la ultima comanda que s'ha fet d'aquest treball s'hauria de
    
    'importartintes
    ultimacomanda = buscarultimacomandadaquesttreball(atrim(clixes.Recordset!codidebarres))
    If ultimacomanda = 0 Then Exit Sub
    Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(ultimacomanda))
    If rstc.EOF Then Exit Sub
    clixes.Recordset.Edit
     clixes.Recordset!redcilindrefw = atrim(rstc!cmaquina)
    clixes.Recordset.Update
    dbclixes.Execute "delete * from tintes where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio)
    formtintes.crear_tintes
    Set rstt = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio) + " order by ordretinter asc")
    While Not rstt.EOF
       rstt.Edit
        rstt!color = rstc.Fields("tinta" + atrim(rstt!ordretinter) + "a")
       ' If atrim(rstt!Color) <> "" Then
         rstt!observacions = rstc.Fields("tinta" + atrim(rstt!ordretinter) + "b")
         rstt!anilox = cadbl(rstc.Fields("lin" + atrim(rstt!ordretinter)))
         rstt!cilindre = cadbl(rstc!cilindres)
         rstt!desarroll = cadbl(rstc!dessarroll)
         rstt!clixeosleeve = "Clixé"
         rstt!numpolimers = 0
       ' End If
       rstt.Update
       rstt.MoveNext
    Wend
    comptartinters
    Unload formtintes
    
    'importarclientsvinculats
    ultimacomanda = 0
    clientprincipal = 0
    Set rstc = dbcomandes.OpenRecordset("select * from comandes where numtreball=" + atrim(id_treball) + " order by client, comanda DESC")
    While Not rstc.EOF
      If ultimacomanda < rstc!comanda Then ultimacomanda = rstc!comanda: clientprincipal = cadbl(rstc!direnvio)
      If ultimclient <> rstc!direnvio Then
         ultimclient = rstc!direnvio
         importarclientvinculat rstc
      End If
      rstc.MoveNext
    Wend
    If cadbl(clientprincipal) > 0 Then
      dbclixes.Execute "update clientsvinculats set principal=true where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio) + " and direnvio=" + atrim(clientprincipal)
    End If
End Sub
Sub importarclientvinculat(rstc As Recordset)
    Dim rstv As Recordset
    Dim rstcli As Recordset
    ruta_relativa_client = carpeta_del_client(rstc!client)
    If ruta_relativa_client = "" Then Exit Sub 'MsgBox "No s'ha trobat la ruta de la carpeta del client: " + atrim(rstc!client):
    Set rstv = dbclixes.OpenRecordset("select * from clientsvinculats where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio) + " and codiclient=" + atrim(rstc!client))
    If Not rstv.EOF Then Exit Sub ' vol dir que ja exiteix i surt
    Set rstcli = dbcomandes.OpenRecordset("SELECT Clients_envios.id, clients.nom, Clients_envios.poblacioe FROM Clients_envios INNER JOIN clients ON Clients_envios.codi = clients.codi WHERE Clients_envios.id=" + atrim(cadbl(rstc!direnvio)) + ";")
    If rstcli.EOF Then Exit Sub 'sinotrobo el client també surtu
    rstv.AddNew
    rstv!id_treball = id_treball
    rstv!ordremodificacio = ordremodificacio
    
    rstv!codiclient = rstc!client
    rstv!direnvio = rstc!direnvio
    rstv!nomclient = atrim(rstcli!nom)
    rstv!nomdirenvio = atrim(rstcli!poblacioe)
    rstv!codimuntadora = atrim(rstc!arxiumontadora)
    rstv!refclient = atrim(rstc!refclient)
    rstv!refclientalternatives = atrim(rstc!refclialt)
    'importar arxiu imp al seu lloc i si n'hi ha possar el rstv!arxiuimp=true
    rstv!arxiuimp = False
    If copiarfitxerimpalseulloc(rstc) Then rstv!arxiuimp = True
    rstv.Update
End Sub
Function crearfitxerimpacopiar(cli As Long, direnvio As Double) As String
   On Error Resume Next
   MkDir ruta_documentacio_clixes + "\" + Format(campid_treball, "00000")
   crearfitxerimpacopiar = ruta_documentacio_clixes + "\" + Format(campid_treball, "00000") + "\IMP" + Format(campid_treball, "00000") + "-" + Format(ordremodificacio, "000") + "-" + Format(cli, "000000") + "_" + atrim(cadbl(direnvio)) + ".doc"
End Function
Function copiarfitxerimpalseulloc(rstc As Recordset) As Boolean
  Dim fitxer_imp_aoncopiar As String
  Dim fitxer_imp_origen As String
  Dim nomfitxer As String
  Dim ruta As String
  Dim numcarpetaclient As String
  copiarfitxerimpalseulloc = False
  fitxer_imp_aoncopiar = crearfitxerimpacopiar(rstc!client, cadbl(rstc!direnvio))
  
  numcarpetaclient = Mid(ruta_relativa_client, 1, 6)
  nomfitxer = atrim(rstc!arxiuimpressora)
  If cadbl(Mid(nomfitxer, 1, 6)) = 0 Then nomfitxer = numcarpetaclient + " " + Trim(nomfitxer)
  ruta = ruta_relativa_docs + "\" + nomfitxer
  If existeix(ruta) Then
     fitxer_imp_origen = ruta
    Else: Exit Function
  End If
  If existeix(fitxer_imp_aoncopiar) Then
    If MsgBox("Aquest client ja te creat un fitxer IMP per aquest treball i modificacio." + Chr(10) + fitxer_imp_origen + Chr(10) + fitxer_imp_aoncopiar + Chr(10) + "Vols sobrescriure'l?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Function
  End If
  Copiar_Fitxer fitxer_imp_origen, fitxer_imp_aoncopiar
  copiarfitxerimpalseulloc = True
  '...
End Function

Function buscarultimacomandadaquesttreball(vcodidebarres As String) As Double
    Dim rst As Recordset
    'Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda, InStr(1,[ruta],'I',1) AS hihaimpresora FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((InStr(1,[ruta],'I',1))>0)) and numtreball=" + atrim(id_treball) + " order by comanda DESC")
    Set rst = dbcomandes.OpenRecordset("SELECT comanda FROM comandes where numtreball=" + atrim(id_treball) + " and producte in (SELECT productes.codi from productes WHERE (((InStr(1,[ruta],'I',1))>0));)" + " order by comanda DESC")
  '  If rst.EOF Then
  '     Set rst = dbcomandes.OpenRecordset("select comanda from comandes where codibarras='" + atrim(vcodidebarres) + "' order by comanda DESC")
  '  End If
    If rst.EOF Then
       buscarultimacomandadaquesttreball = 0
      Else: buscarultimacomandadaquesttreball = cadbl(rst!comanda)
   End If
End Function

Sub importacio_copiarlesliniesdemodificacioialbarans(rstvells As Recordset)
     Dim nous As Recordset
     Dim vells As Recordset
     Set nous = dbclixes.OpenRecordset("select * from clixes_modifi where id_treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio))
     Set vells = dbclixesvells.OpenRecordset("select * from clixes_modifi where id_treball=" + atrim(id_treball))
     While Not vells.EOF
        nous.AddNew
        For i = 0 To vells.Fields.Count - 1
           nous.Fields(vells.Fields(i).Name) = vells.Fields(i)
        Next i
        If IsNull(nous!data_inici) Then nous!data_inici = CVDate("01/01/1900")
        If IsNull(nous!data_prevista) Then nous!data_prevista = nous!data_inici
        nous!ordremodificacio = ordremodificacio
        nous.Update
        vells.MoveNext
     Wend
     
     Set nous = dbclixes.OpenRecordset("select * from clixes_albarans where id_treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio))
     Set vells = dbclixesvells.OpenRecordset("select * from clixes_albarans where id_treball=" + atrim(id_treball))
     While Not vells.EOF
        nous.AddNew
        For i = 0 To vells.Fields.Count - 1
           nous.Fields(vells.Fields(i).Name) = vells.Fields(i)
        Next i
        nous!ordremodificacio = ordremodificacio
        nous.Update
        vells.MoveNext
     Wend
     Set nous = Nothing
     Set vells = Nothing
End Sub
Sub importacio_dadesnovamodificacio(rstvells As Recordset)
    Dim rstc As Recordset
    modificacions.Recordset!dataobertura = rstvells!datainicitreball
    modificacions.Recordset!formaimpresio = atrim(rstvells!forma_imprimir)
    modificacions.Recordset!bandes = cadbl(rstvells!bandesclixes)
    modificacions.Recordset!fotograbador = cadbl(rstvells!proveidor)
    modificacions.Recordset!sistemadimpresio = atrim(rstvells!sistemadimpresio)
    modificacions.Recordset!observacions = atrim(rstvells!observacio)
    
    Set rstc = dbcomandes.OpenRecordset("SELECT comandes.* FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi where InStr([ruta],'I')<>0 and numtreball=" + atrim(rstvells!id_treball) + ";")
    If rstc.EOF Then GoTo sortir
    
    
    modificacions.Recordset!gruixpolimer = cadbl(rstc!gruixpol)
    modificacions.Recordset!amplelamina = amplecomanda(rstc)
    modificacions.Recordset!desarroll = cadbl(rstc!dessarroll)
    
    'modificacions.UpdateControls
    'dataobertura = ""
sortir:
    gravarcanvis
    Set rstc = Nothing
End Sub
Function amplecomanda(rstvells As Recordset) As Double
   Dim rstp As Recordset
   Set rstp = dbcomandes.OpenRecordset("select ruta from productes where codi='" + atrim(rstvells!producte) + "'")
   If rstp.EOF Then Exit Function
   If InStr(1, rstp!ruta, "L") > 0 Then amplecomanda = cadbl(rstvells!ampleutil): Exit Function
   If InStr(1, rstp!ruta, "R") > 0 Then amplecomanda = cadbl(rstvells!amplereb): Exit Function
   If InStr(1, rstp!ruta, "S") > 0 Then amplecomanda = cadbl(rstvells!ampleesq): Exit Function
   'If InStr(1, rstp!ruta, "S") > 0 Then amplecomanda = cadbl(rstvells!amplesol): Exit Function
   amplecomanda = cadbl(rstvells!ampleesq)
End Function
Function bandescomanda(rstvells As Recordset) As Double
   Dim rstp As Recordset
   Set rstp = dbcomandes.OpenRecordset("select ruta from productes where codi='" + atrim(rstvells!producte) + "'")
   If rstp.EOF Then Exit Function
   If InStr(1, rstp!ruta, "R") > 0 Then bandescomanda = cadbl(rstvells!simulteneitatreb): Exit Function
   If InStr(1, rstp!ruta, "S") > 0 Then bandescomanda = cadbl(rstvells!simulteneitatsol): Exit Function
   bandescomanda = 0
End Function
Sub importacio_crearnouclixe(rstvell As Recordset)
   
   clixes.Recordset.AddNew
   clixes.Recordset!id_treball = rstvell!id_treball
   clixes.Recordset!arxiu = rstvell!arxiuclixe
   clixes.Recordset!ubicacio = rstvell!arxiuclixe
   clixes.Recordset!codidebarres = rstvell!codibarres
   clixes.Recordset!marca = rstvell!id_marca
   clixes.Recordset!linia = rstvell!id_liniaproducte
   clixes.Recordset.Update
   clixes.Recordset.FindFirst "id_treball=" + atrim(rstvell!id_treball)
   
End Sub
Sub crearidsmarcailinia()
  Dim rstvells As Recordset
  Dim rstnous As Recordset
  Dim rstmarca As Recordset
  'Set rstvells = dbclixesvells.OpenRecordset("select distinct(id_liniaproducte) as linia,first(id_marca) as marca from clixes group by id_liniaproducte")
  Set rstvells = dbclixesvells.OpenRecordset("select id_liniaproducte as linia,id_marca as marca from clixes where exists (SELECT distinct [id_marca] & '  ' & [id_liniaproducte] AS liniaimarca FROM Clixes);")
  Set rstnous = dbclixes.OpenRecordset("select * from linies")
  rstvells.MoveLast
  rstvells.MoveFirst
  While Not rstvells.EOF
    dbclixes.Execute "insert into marques (marca) values ('" + atrim(rstvells!marca) + "')"
    r = atrim(rstvells!linia)
    Set rstmarca = dbclixes.OpenRecordset("select * from marques where marca='" + atrim(rstvells!marca) + "'")
    If Not rstmarca.EOF Then
       rstnous.AddNew
       rstnous!id_marca = cadbl(rstmarca!id_marca)
       rstnous!linia = atrim(rstvells!linia)
       rstnous.Update
       'dbclixes.Execute "insert into linies (id_marca,linia) values (" + atrim(rstmarca!id_marca) + ",'" + atrim(rstvells!linia) + "')"
       
         ' Else: Stop
    End If
    DoEvents
    Me.caption = atrim(rstvells.AbsolutePosition) + " / " + atrim(rstvells.RecordCount)
    rstvells.MoveNext
  Wend
End Sub
  Sub ensenyalespendents()
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "select comanda from   Avisoscomandessenseidtreball order by comanda"
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns("comanda").width = 1000
   formseleccio.Show 1
   
  End Sub
Sub arreglarlesquejashanfet()
   Dim rst As Recordset
   Dim rstc As Recordset
   Set rst = dbclixes.OpenRecordset("Avisoscomandessenseidtreball")
   While Not rst.EOF
      Set rstc = dbcomandes.OpenRecordset("Select numtreball,proximaseccio from comandes where comanda=" + atrim(rst!comanda))
      If Not rstc.EOF Then
         If cadbl(rstc!numtreball) > 0 Or InStr(1, "TPV", atrim(rstc!proximaseccio)) > 0 Then rst.Delete
      End If
      rst.MoveNext
   Wend
   
   Set rstc = Nothing
   Set rst = Nothing
   
End Sub

Private Sub Command13_Click()
   Dim treball As String
   treball = InputBox("Entra el treball que vols buscar." + Chr(10) + "Si poses un numero de comanda buscara el Treball relacionat", "Busqueda de número de treball")
   If cadbl(treball) > 100000 Then treball = buscartreballdelacomanda(cadbl(treball))
   If cadbl(treball) > 0 Then
       clixes.Recordset.FindFirst "id_treball=" + atrim(cadbl(treball))
   End If
End Sub
Function buscartreballdelacomanda(vnumc As Double) As String
  Dim rst As Recordset
  buscartreballdelacomanda = "0"
  Set rst = dbcomandes.OpenRecordset("select numtreball from comandes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then buscartreballdelacomanda = atrim(rst!numtreball)
End Function

Private Sub Command14_Click()
  Dim rstvells As Recordset
  Dim rstmarca As Recordset
  Dim rstnou As Recordset
  Dim numt As String
  numt = cadbl(InputBox("Entra el numero de treball a importar", "Importar un treball"))
  If cadbl(numt) = 0 Then Exit Sub
  Set rstvells = dbclixesvells.OpenRecordset("select * from clixes where id_treball=" + numt)
  rstvells.MoveLast
  rstvells.MoveFirst
  While Not rstvells.EOF
    Set rstnou = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(rstvells!id_treball))
    If Not rstnou.EOF Then GoTo proxim
    importacio_crearnouclixe rstvells
    crear_modificacio_nova
    importacio_dadesnovamodificacio rstvells
    importacio_colocarelpdfalseulloc rstvells
    importacio_tintesiclientsvinculats
    importacio_copiarlesliniesdemodificacioialbarans rstvells
    clixes.Recordset.Bookmark = clixes.Recordset.LastModified
    possarestatclixe
    Me.caption = atrim(rstvells.AbsolutePosition) + " / " + atrim(rstvells.RecordCount)
proxim:
    DoEvents
    rstvells.MoveNext
  Wend

End Sub

Private Sub Command15_Click()
   Dim rst As Recordset
   Dim rstt As Recordset
   Dim rstc As Recordset
   Dim resp As String
   Dim r As String
   If UCase(InputBox("entra la contrasenya")) <> "INPLACSA" Then Exit Sub
   resp = InputBox("Entra el treball que vols fer de prova o bé * per tots", "Treball a provar")
   If resp = "*" Then
      Set rstm = dbclixes.OpenRecordset("select * from modificacions order by id_treball Desc ,ordre Desc")
        Else
          If cadbl(resp) = 0 Then Exit Sub
          Set rstm = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(cadbl(resp)) + " order by id_treball Desc ,ordre Desc")
   End If
   While Not rstm.EOF
      
      Set rstt = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(rstm!id_treball) + " and ordremodificacio=" + atrim(rstm!ordre))
      Set rstc = dbcomandes.OpenRecordset("select * from comandes where numtreball=" + atrim(rstm!id_treball) + " and (numordremodificacio=" + atrim(rstm!ordre) + " or numordremodificacio=0) order by comanda Desc")
      While Not rstt.EOF
          Me.caption = "Traspasant tintes del treball: " + atrim(rstm!id_treball)
          If Not rstc.EOF Then Me.caption = Me.caption + " i actualitzant la comanda: " + atrim(rstc!comanda) + " " + atrim(i)
          DoEvents
          r = ""
          If InStr(1, Mid(atrim(rstt!color), 1, 4), "º") > 0 Then
              
              r = Mid(atrim(rstt!color), 3)
              rstt.Edit
              rstt!color = atrim(r)
              rstt.Update
              'If Not rstc.EOF Then
              '  rstc.Edit
              '  rstc.Fields("tinta" + atrim(rstt!ordretinter) + "a") = atrim(r)
              '  rstc.Update
              'End If
           Else: r = atrim(rstt!color)
          End If
          If Not rstc.EOF Then
              If rstt!ordretinter > 0 Then
               If InStr(1, Mid(atrim(rstc.Fields("tinta" + atrim(rstt!ordretinter) + "a")), 1, 4), "º") > 0 Then
                 i = i + 1
                 rstc.Edit
                 rstc.Fields("tinta" + atrim(rstt!ordretinter) + "a") = atrim(r)
                 rstc.Update
               End If
              End If
          End If
          rstt.MoveNext
      Wend
      rstm.MoveNext
   Wend
   MsgBox "Procès acavat. " + atrim(i)
End Sub
Sub possarcolorbotopressupost()
  Dim rst As Recordset
  botopressupost.BackColor = COLORBOTOSENSEDADES
  
  Set rst = dbclixes.OpenRecordset("select * from pressupostos where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
  If Not rst.EOF Then
      botopressupost.BackColor = &H6BEBB1
      If rst!enviat Then botopressupost.BackColor = QBColor(12)
      If rst!comfirmat Then botopressupost.BackColor = QBColor(10)
  End If
End Sub
Sub possarcolorbotocomandafotogravador()
  Dim rst As Recordset
  botocomandesfotogravador.BackColor = COLORBOTOSENSEDADES
  
  Set rst = dbclixes.OpenRecordset("select * from comandesfotogravador where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
  If Not rst.EOF Then
      botocomandesfotogravador.BackColor = COLORBOTOAMBDADES
      If rst!okenviat Then botocomandesfotogravador.BackColor = QBColor(12)
      If IsDate(rst!datarecepcio) Then botocomandesfotogravador.BackColor = QBColor(10)
  End If
End Sub

Private Sub Command16_Click()
  Dim rst As Recordset
   Dim rstt As Recordset
   Dim rstc As Recordset
   Dim rsttras As Recordset
   Dim rsttintes As Recordset
   Dim dbtintes As Database
   Dim resp As String
   Dim r As String
   If UCase(InputBox("entra la contrasenya")) <> "INPLACSA" Then Exit Sub
   resp = InputBox("Entra el treball que vols fer de prova o bé * per tots", "Treball a provar")
   If resp = "*" Then
      Set rstm = dbclixes.OpenRecordset("select * from modificacions order by id_treball Desc ,ordre Desc")
        Else
          If cadbl(resp) = 0 Then Exit Sub
          Set rstm = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(cadbl(resp)) + " order by id_treball Desc ,ordre Desc")
   End If
   Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
   Set rsttras = dbtintes.OpenRecordset("select * from tmp_colorsdelstreballs")
   While Not rstm.EOF
      
      Set rstt = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(rstm!id_treball) + " and ordremodificacio=" + atrim(rstm!ordre))
      Me.caption = "Traspasant tintes del treball: " + atrim(rstm!id_treball)
      DoEvents
      While Not rstt.EOF
          
          'If Not rstc.EOF Then Me.caption = Me.caption + " i actualitzant la comanda: " + atrim(rstc!comanda) + " " + atrim(i)
          
          r = ""
          
          If InStr(1, Mid(atrim(rstt!color), 1, 4), "º") > 0 Then
              r = Mid(atrim(rstt!color), 3)
              
           Else: r = atrim(rstt!color)
          End If
          If r <> "" Then
                'buscar el color
                rsttras.FindFirst "nomcolor='" + atrim(r) + "'"
                If Not rsttras.NoMatch Then
                    Set rsttintes = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(rsttras!coditinta) + "'")
                    If Not rsttintes.EOF Then
                          rstt.Edit
                          rstt!coloranterior = atrim(rstt!color)
                          rstt!color = atrim(rsttintes!descripcio)
                          rstt!detalltinter = atrim(rsttras!detalldeltinter)
                          rstt!coditinta = rsttras!coditinta
                          rstt.Update
                    End If
                End If
          End If
          rstt.MoveNext
      Wend
      rstm.MoveNext
   Wend
   Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
   Set rstm = dbbaixes.OpenRecordset("select distinct tinta_comanda from impresores_aniloxos")
   While Not rstm.EOF
     rsttras.FindFirst "nomcolor='" + atrim(rstm!tinta_comanda) + "'"
     If Not rsttras.NoMatch Then
        Set rsttintes = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(rsttras!coditinta) + "'")
        If Not rsttintes.EOF Then
          dbbaixes.Execute "update impresores_aniloxos set coditinta_comanda='" + atrim(rsttras!coditinta) + "', tinta_comanda='" + atrim(rsttintes!descripcio) + "',detalltinter_comanda='" + atrim(rsttras!detalldeltinter) + "' where tinta_comanda='" + atrim(rstm!tinta_comanda) + "'"
        End If
     End If
     rstm.MoveNext
   Wend
   
   MsgBox "Procès acavat. " + atrim(i)
End Sub

Private Sub Command17_Click()
   '  Dim rst As Recordset
   '  Set rst = dbclixes.OpenRecordset("select * from clixes where mid(arxiu,1,2)='G-'")
   '  While Not rst.EOF
   '     rst.Edit
   '     rst!ubicacio = generarnumpalet(rst!arxiu)
   '     rst.Update
   '     rst.MoveNext
   '  Wend
   '  MsgBox "ja esta"
 
     
End Sub
Function generarnumpalet(arxiu As String) As String
    Dim numg As Double
    numg = cadbl(Mid(arxiu, 3))
    generarnumpalet = "Palet-0"
    If numg >= 1 And numg <= 70 Then generarnumpalet = "Palet-1"
    If numg >= 71 And numg <= 145 Then generarnumpalet = "Palet-2"
    If numg >= 146 And numg <= 224 Then generarnumpalet = "Palet-3"
    If numg >= 225 And numg <= 299 Then generarnumpalet = "Palet-4"
    If numg >= 300 And numg <= 340 Then generarnumpalet = "Palet-5"
    If numg >= 341 And numg <= 403 Then generarnumpalet = "Palet-6"
    If numg >= 484 And numg <= 628 Then generarnumpalet = "Palet-7"
    If numg >= 404 And numg <= 453 Then generarnumpalet = "Palet-7"
    If numg >= 543 And numg <= 552 Then generarnumpalet = "Palet-6"
    If numg >= 485 And numg <= 542 Then generarnumpalet = "Palet-7"
    
End Function

Private Sub Command18_Click()
   If marcaproducte = "" Or liniaproducte = "" Then MsgBox "La Marca i Linia han d'estar posades per poder assignar la linia d'impressió.", vbCritical, "Error": Exit Sub
   Command18.tag = marcaproducte
   vincularliniesimpresio.Show
   vincularliniesimpresio.SetFocus
   Command18.tag = ""
   While Command18.tag = ""
     DoEvents
   Wend
   Command18.tag = ""
   gravarcanvis True
   modificacions.Recordset.Move 0
   posaraobservaciotintes_eltreballrelacionatsicorrespon
   
End Sub
Sub posaraobservaciotintes_eltreballrelacionatsicorrespon()
  Dim rst As Recordset
  Dim rstm As Recordset
  Dim valoractual As Double
  Dim valoranterior As Double
  Set rst = dbclixes.OpenRecordset("select * from tintes_observacions where id_Treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio) + " order by id")
  Set rstm = dbclixes.OpenRecordset("select numerodelinia,ordre from modificacions where id_Treball=" + atrim(id_treball) + " order by ordre asc")
  rstm.FindFirst "ordre=" + atrim(ordremodificacio)
  If Not rstm.NoMatch Then
      valoractual = cadbl(rstm!numerodelinia)
      If ordremodificacio > 1 Then
          rstm.MovePrevious
          valoranterior = cadbl(rstm!numerodelinia)
          If valoractual = valoranterior Then valoractual = 0: GoTo fi
      End If
  End If
  If rst.EOF Then
      'possar la observació
      Set rstm = dbclixes.OpenRecordset("select arxiu,clixes.id_treball,ordre from clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball where marca='" + atrim(clixes.Recordset!marca) + "' and clixes.id_treball<>" + atrim(id_treball) + " And numerodelinia = " + atrim(valoractual))
      If Not rstm.EOF Then
          dbclixes.Execute "Insert into tintes_observacions (id_treball,ordre,observacio) values (" + atrim(id_treball) + "," + atrim(ordremodificacio) + ",'Té la mateixa línia de disseny que el treball " + atrim(rstm!id_treball) + ". " + atrim(rstm!arxiu) + "')"
      End If
  End If
fi:
  Set rstm = Nothing
  Set rst = Nothing
End Sub


Private Sub Command19_Click()
   cobservaciorepasclixes.visible = Not cobservaciorepasclixes.visible
End Sub

Private Sub Command2_Click()
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim rstm As Recordset
   Dim rstv As Recordset
   Set rst = dbclixes.OpenRecordset("select * from clixes")
   While Not rst.EOF
     Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=(select max(comanda) from comandes where numtreball=" + atrim(rst!id_treball) + ")")
     If Not rstc.EOF Then
        'MsgBox atrim(rstc!comanda) + "    Clixe: " + atrim(rst!arxiu) + "   Comanda: " + atrim(rstc!arxiu)
        
        If atrim(rst!arxiu) = "" Then
           rst.Edit
           rst!arxiu = atrim(rstc!arxiu)
           rst.Update
        End If
        'Set rstm = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(rst!id_treball) + " order by ordre DESC")
        'If Not rstm.EOF Then
        '   If atrim(rstc!formaimp) <> "" Then
        '     rstm.Edit
        '     rstm!formaimpresio = atrim(rstc!formaimp)
        '     rstm.Update
        '   End If
        '   Set rstv = dbclixes.OpenRecordset("select * from clientsvinculats where id_treball=" + atrim(rst!id_treball) + " and ordremodificacio=" + atrim(rstm!ordre) + " and codiclient=" + atrim(rstc!client) + " and direnvio=" + atrim(cadbl(rstc!direnvio)))
        '   If Not rstv.EOF Then
        '    rstv.Edit
        '    rstv!codimuntadora = atrim(rstc!arxiumontadora)
        '    rstv!refclient = atrim(rstc!refclient)
        '    rstv!refclientalternatives = atrim(rstc!refclialt)
        '    rstv.Update
        '   End If
        'End If
        
     End If
     Me.caption = rst.AbsolutePosition
     DoEvents
     rst.MoveNext
   Wend
End Sub

Private Sub Command9_Click()
  borrardatamodificacio "dataprovacolor"
End Sub

Private Sub alta_Click()
   Dim estapa As Boolean
   Dim treballacopiartapa As Double
   Dim rsttapa As Recordset
   If clixes.Recordset.EditMode > 0 Then MsgBox "Estas editant primer finalitza la operació i despres afegeix.", vbCritical, "Atenció": Exit Sub
   estapa = IIf(MsgBox("Aquest treball que comences s'utilitzarà com a tapa?", vbInformation + vbYesNo, "Atenció") = vbYes, True, False)
   If estapa Then treballacopiartapa = cadbl(InputBox("Entra el treball d'on vols copiar els valors de distorsió per la tapa.", "Copiar distorsió d'impresió"))
   clixes.Recordset.AddNew
   campid_treball = clixeidgran + 1
   cportareduccio.Value = IIf(estapa, 1, 0)
   If treballacopiartapa > 0 Then
        Set rsttapa = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(treballacopiartapa))
        If Not rsttapa.EOF Then
            reducciopermetre = cadbl(rsttapa!reduccioxmetre)
            reducciocilindref2 = cadbl(rsttapa!redcilindref2)
            reducciocilindrefw = cadbl(rsttapa!redcilindrefw)
        End If
   End If
   clixes.Recordset!id_treball = cadbl(campid_treball)
   framesactivats True
   
   codidebarres.SetFocus
End Sub
Sub framesactivats(activats As Boolean)
   frameclixes.Enabled = activats
   framemodificacions.Enabled = activats
   framehistorialmodificacions.Enabled = Not activats
End Sub
Function clixeidgran() As Long
   Dim rst As Recordset
   clixeidgran = 0
   Set rst = dbclixes.OpenRecordset("select max(id_treball) as gran from clixes")
   If Not rst.EOF Then clixeidgran = cadbl(rst!gran)

   Set rst = Nothing
End Function
Function ordremesgran() As Long
   Dim rst As Recordset
   ordremesgran = 0
   Set rst = dbclixes.OpenRecordset("select max(ordre) as gran from modificacions where id_treball=" + atrim(cadbl(clixes.Recordset!id_treball)))
   If Not rst.EOF Then ordremesgran = cadbl(rst!gran)
   Set rst = Nothing
End Function


Private Sub botoclientsvinculats_Click()
   formclientsvinculats.Show 1
   carregar_modificacio
   comprovardiferenciescomandesafectades
End Sub

Private Sub botoliniesalbarans_Click()
  formliniesalbarans.Show 1
  revisarpreualbaransipressupost
End Sub
Sub revisarpreualbaransipressupost()
  comprovarsialbaransmesquepressupost
  If albaransmesquepressupost Then MsgBox "Atenció el cost dels clixes supera el pressupostat.", vbCritical, "Atenció"
End Sub
Private Sub botoliniesmodificacions_Click()
   Dim vestatanterior As String
   vestatanterior = etestatclixemod
   formliniesmodificacio.Show 1
'   vestatanterior = "04 REPOSICIÓ DEL CLIXE"
   possarestatclixe
   If InStr(1, etestatclixemod, "CLIXES ENTRATS") > 0 And InStr(1, vestatanterior, "CLIXES ENTRATS") = 0 Then
       mirarCOSESquanhihaCLIXESENTRATS vestatanterior
   End If
End Sub
Sub mirarCOSESquanhihaCLIXESENTRATS(vestatanterior As String)
   Dim i As Integer
   Dim numcomanda As String
   numcomanda = ""
   For i = 0 To llistadecomandespendents.ListCount - 1
       numcomanda = Mid(llistadecomandespendents.List(i), 1, 7)
       imprimirbossessoldadores (cadbl(numcomanda))
   Next i
   If cadbl(numcomanda) > 0 Then comprovarsihihaamuntadoraeltreballiavisar cadbl(numcomanda)
   passarrepasalasecciodIMPRESORES cadbl(id_treball), cadbl(ordremodificacio), vestatanterior
End Sub
Function demanarnumerodetaula() As Byte
   Dim v As String
   While cadbl(v) = 0
      v = InputBox("Escriu el numero de taula on vols fer la revisió dels CLIXES ENTRATS.", "NUMERO DE TAULA")
   Wend
   demanarnumerodetaula = cadbl(v)
End Function
Function possar_numtaula_1o2() As String
   Dim v As String
   v = llegir_ini("General", "ultim_numtaula", rutadelfitxer(cami) + "valorsprograma.ini")
   If v = "1" Then v = "2" Else v = "1"
   possar_numtaula_1o2 = v
   escriure_ini "General", "ultim_numtaula", v, rutadelfitxer(cami) + "valorsprograma.ini"
End Function
Sub passarrepasalasecciodIMPRESORES(cidtreball As Double, cversio As Double, vestatanterior As String)
   Dim rst As Recordset
   Dim rstclixes As Recordset
   Dim vnumtaula As String
   Dim vesreposicio As Boolean
   
   If InStr(1, vestatanterior, "REPOSICIÓ DEL CLIXE") > 0 Then vesreposicio = True
   Set rst = dbclixes.OpenRecordset("select * from clixesentrats_control where datafet=null and numtreball=" + atrim(cidtreball) + " and versio=" + atrim(cversio))
   If Not rst.EOF Then MsgBox "Aquests clixes ja estan passats a REVISAR I TALLAR i no es poden tornar entrar fins que s'hagin revisat.", vbCritical, "Atenció": GoTo fi
   Set rstclixes = dbclixes.OpenRecordset("SELECT Clixes.ubicacio, modificacions.desarroll FROM Clixes RIGHT JOIN modificacions ON Clixes.id_treball = modificacions.id_treball Where modificacions.id_treball = " + atrim(cidtreball) + " And ordre = " + atrim(cversio))
   If rstclixes.EOF Then GoTo fi
   If vesreposicio Then
     vnumtaula = "N-10"
       Else:
           vnumtaula = atrim(rstclixes!ubicacio)
           If vnumtaula = "" Then vnumtaula = "T-" + atrim(possar_numtaula_1o2) 'atrim(demanarnumerodetaula)
   End If
   rst.AddNew
   rst!numtreball = cadbl(cidtreball)
   rst!versio = cadbl(cversio)
   If Not rstclixes.EOF Then rst!desarroll = cadbl(rstclixes!desarroll)
   rst!dataentrada = Now
   If vesreposicio Then rst!datarepas = Now: rst!reposicio = True
   rst!numtaula = vnumtaula
   rst.Update
   If clixes.Recordset.EditMode = 0 Then clixes.Recordset.Edit
   clixes.Recordset!ubicacio = vnumtaula
   clixes.Recordset.Update
   
   If vesreposicio Then MsgBox "S'ha passat a pendent de revisar aquesta reposició.", vbInformation, "Reposició"
   If Not vesreposicio Then enviar_email_a_tintesperavisar cidtreball, cversio
fi:
   Set rst = Nothing
   Set rstclixes = Nothing
End Sub
Sub enviar_email_a_tintesperavisar(cidtreball As Double, cversio As Double)
   Dim v As String
   Dim vcomandes As String
   For i = 0 To llistadecomandespendents.ListCount - 1
      vcomandes = vcomandes + " " + llistadecomandespendents.List(i)
   Next i
   v = "Han entrat els clixes del treball " + atrim(cidtreball) + "/" + atrim(cversio) + " de la comanda " + vcomandes
   v = v + vbNewLine + "Texte impresio: " + atrim(marcaproducte) + " - " + atrim(liniaproducte)
   enviaremailgeneric "tintes@inplacsa.com", "CLIXES ENTRATS a disseny. " + atrim(cidtreball) + "/" + atrim(cversio), treure_apostruf(v)
End Sub
Sub comprovarsihihaamuntadoraeltreballiavisar(vcomanda As Double)
    Dim rst As Recordset
    Dim vresp As String
    Dim vcos As String
    Dim vnumtreball As Double
    
    Set rst = dbcomandes.OpenRecordset("select numtreball from comandes where comanda=" + atrim(vcomanda))
    If rst.EOF Then GoTo fi
    Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
    vnumtreball = cadbl(rst!numtreball)
    Set rst = dbbaixes.OpenRecordset("SELECT muntadora_ordremuntatge.comanda, comandes.numtreball FROM muntadora_ordremuntatge INNER JOIN comandes ON muntadora_ordremuntatge.comanda = comandes.comanda where numtreball=" + atrim(vnumtreball))
    If Not rst.EOF Then
       While vresp <> "TREBALL A MUNTADORA"
         vresp = UCase(InputBox(vbCr & vbCr & vbCr & vbCr & vbCr + "Hi ha una comanda entrada a muntadora amb aquest mateix treball." & vbCr & "Escriu [TREBALL A MUNTADORA] per continuar." & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr, "ATENCIÓ"))
       Wend
       'avisar per email
       vcos = "ATENCIÓ!!! s'ha activat la comanda " + atrim(vcomanda) + " i el mateix treball està entrat a ordre de muntatge de muntadora."
       enviaremailgeneric "tintes@inplacsa.com;impresores@inplacsa.com", "URGENT!!! Comanda activada amb el treball a muntadora.", treure_apostruf(vcos)
       'tintes@inplacsa.com;impresores@inplacsa.com
    End If
fi:
    Set rst = Nothing
    Set dbbaixes = Nothing
 End Sub

Sub imprimirbossessoldadores(vnumc As Double, Optional vnocrearcodi As Boolean)
  Dim rst As Recordset
  Dim rstc As Recordset
  Dim vnumbossa As String
  Set rst = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then Set rstc = dbcomandes.OpenRecordset("select * from comandes_extres where comanda=" + atrim(vnumc))
  If Not rst.EOF And Not rstc.EOF Then
     If (atrim(rstc!numerobossasoldadores) = "" Or vnocrearcodi) And larutahiha(atrim(rst!producte), "S") Then
       If Not vnocrearcodi Then dbcomandes.Execute "update  comandes_extres set numerobossasoldadores='" + atrim(rst!numtreball) + "' where comanda=" + atrim(vnumc)
       MsgBox "Comanda " + atrim(vnumc) + Chr(10) + "Impresió de la Bossa de Soldadora i/o la seva VQ." + Chr(10) + "Possa-ho junt amb la comanda.", vbInformation, "Atenció"
       vnumbossa = atrim(rst!numtreball)
       Set rst = dbcomandes.OpenRecordset("select * from comandes_extres where numerobossasoldadores='" + vnumbossa + "' and comanda<>" + atrim(vnumc))
       If rst.EOF Then imprimiretiquetabossasoldadores vnumc, llistat, False
       Set rst = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
       imprimir_VQ_soldadores rst
     End If
  End If
  Set rst = Nothing
  Set rsc = Nothing
End Sub
Sub possarestatclixe()
  Dim rst As Recordset
  etestatclixemod = estatclixemod(id_treball, ordremodificacio)
  Set rst = dbclixes.OpenRecordset("SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, CLIXES_MODIFI.data_prevista,Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE Clixes_modifi.id_treball=" + atrim(id_treball) + " AND Clixes_modifi.ordremodificacio=(select max(ordre) from modificacions where id_treball=" + atrim(id_treball) + ")  And clixes_modifi.ordre = (select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(id_treball) + " AND Clixes_modifi.ordremodificacio=(select max(ordre) from modificacions where id_treball=" + atrim(id_treball) + "));")
  etestatclixe = ""
  If Not rst.EOF Then
    etestatclixe = IIf(Not IsDate(rst!data_fi), Format(rst!data_prevista, "dd/mm") + " - ", "") + atrim(rst!descrip)
    dbclixes.Execute "update clixes set estatclixe='" + atrim(etestatclixe) + "' where id_treball=" + atrim(id_treball)
    
  End If
End Sub
Function estatclixemod(ByVal ntreball As Double, ByVal ordrem As Double) As String
  Dim rst As Recordset
  If ordrem = 0 Then ordrem = 1
  Set rst = dbclixes.OpenRecordset("SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, CLIXES_MODIFI.data_prevista,Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE Clixes_modifi.id_treball=" + atrim(ntreball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordrem) + " AND clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(ntreball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordrem) + ");")
  ' CONSULTA ULTIMA MODIFICACIO AMB DATA FI  VINCULANT "SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE (((Clixes_modifi.id_treball)=" + atrim(id_treball) + ") AND ((Clixes_modifi.ordremodificacio)=" + atrim(ordremodificacio) + ") AND ((Clixes_estats.vinculant)=True and isdate(clixes_modifi.data_fi)) and clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(id_treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + "));"
  ' CONSULTA ULTIMA MODIFICACIO AMB DATA FI SENSE VINCULANT "SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE (((Clixes_modifi.id_treball)=" + atrim(id_treball) + ") AND ((Clixes_modifi.ordremodificacio)=" + atrim(ordremodificacio) + ") AND (isdate(clixes_modifi.data_fi)) and clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(id_treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + "));"
  If Not rst.EOF Then
     estatclixemod = IIf(Not IsDate(rst!data_fi), Format(rst!data_prevista, "dd/mm") + " - ", "") + atrim(rst!descrip)
       Else: estatclixemod = ""
  End If
End Function
Function nomfitxerpdfeditable(vnompdf As String) As String
    nomfitxerpdfeditable = Mid(vnompdf, 1, InStr(1, vnompdf, ".pdf") - 1) + "_Editable.pdf"
End Function
Sub eliminar_elpdf(v As String)
    Dim pdfeditable As String
    Dim pdfmini As String
    If InStr(1, v, "pdf") > 0 Then
       If MsgBox("Segur que vols eliminar el PDF?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then GoTo fi
       pdfeditable = nomfitxerpdfeditable(rutapdftreball)
       pdfmini = substituir(pdfeditable, "_Editable.pdf", "_Mini.pdf")
       If existeix(pdfeditable) Then Kill pdfeditable
       If existeix(rutapdftreball) Then Kill rutapdftreball
       If existeix(pdfmini) Then Kill pdfmini
       dbclixes.Execute "update modificacions set pdfvalid=false where id_treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio)
    End If
    If InStr(1, v, "separacio") > 0 Then
       If MsgBox("Segur que vols eliminar el PDF de separació de colors?", vbCritical + vbYesNo, "Atenció") = vbNo Then GoTo fi
       If existeix(rutapdftreball(, True)) Then Kill rutapdftreball(, True)
    End If
    If InStr(1, v, "cingular") > 0 Then
       If MsgBox("Segur que vols eliminar el PDF Cingular?", vbCritical + vbYesNo, "Atenció") = vbNo Then GoTo fi
       If existeix(rutapdftreball(, True, True)) Then Kill rutapdftreball(, True, True)
       eliminar_elpng cadbl(id_treball), cadbl(ordremodificacio)
    End If
    If InStr(1, v, "previ") > 0 And InStr(1, v, "previSC") = 0 Then
       If MsgBox("Segur que vols eliminar el PDF Prèvi?", vbCritical + vbYesNo, "Atenció") = vbNo Then GoTo fi
       If existeix(rutapdftreball(, True, True, True)) Then Kill rutapdftreball(, True, True, True)
    End If
    If InStr(1, v, "previSC") > 0 Then
       If MsgBox("Segur que vols eliminar el PDF Prèvi de SEPARACIÓ DE COLORS?", vbCritical + vbYesNo, "Atenció") = vbNo Then GoTo fi
       If existeix(rutapdftreball(, True, True, True, True)) Then Kill rutapdftreball(, True, True, True, True)
    End If
fi:
    
End Sub
Sub eliminar_elpng(id_treball As Double, ordremodificacio As Double)
   Dim vrutapng As String
   vrutapng = llegir_ini("ruta", "ruta_pdf_a_png", rutadelfitxer(cami) + "valorsprograma.ini")
   vrutapng = vrutapng + "\" + Format(id_treball, "00000") + "-" + Format(ordremodificacio, "00") + ".png"
   If existeix(vrutapng) Then Kill vrutapng
End Sub

Private Sub botopdf_Click()
   carregar_veure_pdfs
End Sub
Sub carregar_veure_pdfs()
   Dim tipusdepdf As String
  ' If modificacions.Recordset.EditMode > 0 Then MsgBox "Primer guarda canvis abans d'obrir el pdf": Exit Sub
   tipusdepdf = demanarcomplertoseparaciocolor
   If InStr(1, tipusdepdf, "Borrar_") > 0 Then eliminar_elpdf tipusdepdf: GoTo fi
   If tipusdepdf = "N" Then
        If existeix(rutapdftreball) Then
           If existeix(nomfitxerpdfeditable(rutapdftreball)) Then
              obrir_document nomfitxerpdfeditable(rutapdftreball)
                   Else: obrir_document rutapdftreball
           End If
           If Not modificacions.Recordset!pdfvalid Then
             dbclixes.Execute "update modificacions set pdfvalid=true where id_treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio)
             'If modificacions.Recordset.EditMode = 0 Then modificacions.Recordset.Edit
             'modificacions.Recordset!pdfvalid = True
             'modificacions.Recordset.Update
             'modificacions.Recordset.Move 0
           End If
          Else
            If ordremodificacio > 1 Then
             If MsgBox("No hi ha pdf assignat, VOLS COPIAR EL DE LA MODIFICACIO ANTERIOR?", vbInformation + vbYesNo + vbDefaultButton2, "Copiar PDF") = vbYes Then
               copiarpdfanterior
             End If
            End If
        End If
   End If
   If tipusdepdf = "SC" Then If existeix(rutapdftreball(, True)) Then obrir_document rutapdftreball(, True)
   If tipusdepdf = "CR" Then If existeix(rutapdftreball(, True, True)) Then obrir_document rutapdftreball(, True, True)
   If tipusdepdf = "PR" Then If existeix(rutapdftreball(, True, True, True)) Then obrir_document rutapdftreball(, True, True, True)
   If tipusdepdf = "PRSC" Then If existeix(rutapdftreball(, True, True, True, True)) Then obrir_document rutapdftreball(, True, True, True, True)
fi:
End Sub
Sub copiarpdfanterior()
  guardarelpdf rutapdftreball(True), "N", True
  guardarelpdf rutapdftreball(True, True), "SC", True
  guardarelpdf rutapdftreball(True, True), "CR", True
  modificacions.Recordset.Move 0
End Sub

Private Sub bototintes_Click()
   Dim vtinters As Double
   Dim vdesarroll As Double
   Dim vmsg As String
   If modificacions.Recordset.EOF Then Exit Sub
   vtinters = cadbl(modificacions.Recordset!tinters)
   vdesarroll = cadbl(modificacions.Recordset!desarroll)
   
   formtintes.Show 1
   modificacions.Recordset.Move 0
   If formtintes.tag = "reprint" Then
     Unload formtintes
     ordremodificacio = ordremodificacio * IIf(ordremodificacio > 0, -1, 1)
     formtintes.Show 1
     ordremodificacio = ordremodificacio * IIf(ordremodificacio < 0, -1, 1)
   End If
   comptartinters
  ' If modificartintes Then
  '     If arguments(6) = "" Or arguments(6) = "+TIN" Then comprovardiferenciescomandesafectades
  '      Else: comprovardiferenciescomandesafectades
  ' End If
   comprovardiferenciescomandesafectades
   posarcolor_bototintes
   Unload formtintes
   If llistadecomandespendents.ListCount > 0 Then
     If vdesarroll <> cadbl(modificacions.Recordset!desarroll) Then vmsg = "Desarroll: " + atrim(vdesarroll) + "-> " + atrim(modificacions.Recordset!desarroll) + vbNewLine
     If vtinters <> cadbl(modificacions.Recordset!tinters) Then vmsg = vmsg + "Tinters: " + atrim(vtinters) + "-> " + atrim(modificacions.Recordset!tinters) + vbNewLine
     If vmsg <> "" Then
       enviar_canvis_tintersidesarrolls vmsg
     End If
   End If
End Sub
Sub posarcolor_bototintes()
   Dim vestat As String
   If bototintes.BackColor = COLORBOTOSENSEDADES Then Exit Sub
   vestat = UCase(atrim(modificacions.Recordset!estatrevisiotintes))
   
   If vestat = "DISSENY" Then bototintes.BackColor = &H5C31DD: GoTo fi        'vermell xulu
   If InStr(1, vestat, "OK DISSENY") > 0 Then bototintes.BackColor = &H17D062: GoTo fi        'verd xulu
   If InStr(1, vestat, "+IMP") > 0 And InStr(1, vestat, "+TIN") > 0 Then bototintes.BackColor = &HFFFF&: GoTo fi     'GROG XULU  Quan està a punt per donar OK DISSENY
   If vestat = "" Then GoTo fi
   bototintes.BackColor = &H80FF&    ' taronja    'si no es cap dels anteriors O SIGUI FALT O  +IMP O  +TIN
fi:
End Sub
Sub enviar_canvis_tintersidesarrolls(vmsg As String)
   Dim rst As Recordset
   Dim vcomandes As String
   Dim vcap As String
   Dim i As Byte
   Dim rstc As Recordset
   Dim vnomclient As String
   
   
   If llistadecomandespendents.ListCount = 0 Then Exit Sub
   For i = 0 To llistadecomandespendents.ListCount - 1
       vcomandes = vcomandes + " " + atrim(cadbl(Mid(llistadecomandespendents.List(i), 1, 7)))
   Next i
   
   Set rst = dbclixes.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(Mid(llistadecomandespendents.List(0), 1, 7))))
   If rst.EOF Then Exit Sub
       Set rstc = dbclixes.OpenRecordset("select nom from clients where codi=" + atrim(rst!client))
       If Not rstc.EOF Then vnomclient = atrim(rstc!nom)
       vcap = "Codi client: " + atrim(rst!client) + "-" + vnomclient + vbNewLine
       vcap = vcap + "Ref.Client: " + atrim(rst!refclient) + IIf(atrim(rst!comandaclient) <> "", " Com.Cli: " + atrim(rst!comandaclient), "") + vbNewLine
       vcap = vcap + "Texte Imp: " + atrim(rst!marcailinia) + vbNewLine + vbNewLine + "S'ha canviat:" + vbNewLine
       enviaremailgeneric "comandesrevisarpreus@inplacsa.com", "La comanda " + atrim(rst!comanda) + " revisar preu s'ha modificat paràmetres", vcap + Chr(13) + Chr(10) + Chr(13) + Chr(10) + vmsg
End Sub
Sub comptartinters()
  Dim rst As Recordset
  Dim ntintes As Long
  Dim desarroll As Double
  Dim vsql As String
  
  Set rst = dbclixes.OpenRecordset("SELECT Max(Tintes.desarroll) AS desarrollgran, Count(Tintes.id_tinter) AS tintes From Tintes WHERE Tintes.id_treball=" + atrim(id_treball) + " AND Tintes.ordremodificacio=" + atrim(ordremodificacio) + " AND (Tintes.color<>'' OR Tintes.tinterlinkambid_treball>0);")
  If Not rst.EOF Then
     ntintes = cadbl(rst!tintes)
     desarroll = cadbl(rst!desarrollgran)
     If desarroll = 0 Then
       vsql = "SELECT Tintes.desarroll from Tintes WHERE (((Tintes.desarroll)>0) AND "
       vsql = vsql + " ((Tintes.id_tinter) In (SELECT Tintes.tinterlinkambid_treball  from Tintes WHERE "
       vsql = vsql + " (((Tintes.id_treball)=" + atrim(id_treball) + ") AND ((Tintes.ordremodificacio)=" + atrim(ordremodificacio) + ") AND ((Tintes.tinterlinkambid_treball)>0));)));"
       Set rst = dbclixes.OpenRecordset(vsql)
       If Not rst.EOF Then desarroll = cadbl(rst!desarroll)
     End If
       Else: ntintes = 0: desarroll = 0
  End If
  
  modificacions.Recordset.Edit
  modificacions.Recordset!tinters = ntintes
  modificacions.Recordset!desarroll = desarroll
  modificacions.Recordset.Update
End Sub

Private Sub clixes_Reposition()
  carregar_clixe
End Sub
Sub netejaretiquetes()
   nomproveidor = ""
   id_treball = 0
   ordremodificacio = 0
   comandespendents.visible = False
   llistadecomandespendents.visible = False
   nomclient = ""
   nomclientclixe = ""
End Sub
Sub carregar_clixe()
   netejaretiquetes
   If Not clixes.Recordset.EOF Then
      cportareduccio.Value = IIf(clixes.Recordset!portareduccio, 1, 0)
      id_treball = cadbl(clixes.Recordset!id_treball)
      modificacions.RecordSource = "select * from modificacions where id_treball=" + atrim(cadbl(clixes.Recordset!id_treball)) + " ordeR by ordre DESC"
      modificacions.Refresh
      posarfotograbador
      mirarsicomandespendentsialtres
      If UCase(atrim(marcaproducte)) = "TEST" Then
           btestcolor.visible = True
            Else: btestcolor.visible = False
      End If
   End If
   
End Sub
Function escullirordredesdecomandespendents(treball As Double) As Integer
   comandabuscada = 0
   escullirordredesdecomandespendents = 0
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "select comanda ,numtreball as Treball,numordremodificacio as Versió,proximaseccio as Estat from comandes where (producte<>'PC' and producte<>'PCP' and producte<>'PC2') and numtreball=" + atrim(cadbl(treball)) + " and (proximaseccio <>'T') order by comanda Desc"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.Show 1
   
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
              escullirordredesdecomandespendents = formseleccio.DBGrid2.Columns("Versió")
              If escullirordredesdecomandespendents = 0 Then escullirordredesdecomandespendents = 1
              comandabuscada = cadbl(formseleccio.DBGrid2.Columns("Comanda"))
        End If
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   
End Function
Sub mirarsicomandespendentsialtres()
   Dim rstc As Recordset
   Dim rstcextres As Recordset
   Dim hihacomplexes As Boolean
   comandespendents.visible = False
   llistadecomandespendents.visible = False
   llistadecomandespendents.Clear
   camplelamina.BackColor = QBColor(15)
   Text10.BackColor = QBColor(15)
   Text5.BackColor = &HC0C0C0
   etdesarrollerroni.visible = False
   Set rstc = dbcomandes.OpenRecordset("select * from comandes where (producte<>'PC' and producte<>'PCP' and producte<>'PC2') and numtreball=" + atrim(cadbl(clixes.Recordset!id_treball)) + " and (proximaseccio <>'T') order by comanda Desc")
   If Not rstc.EOF Then
     rstc.MoveLast
     rstc.MoveFirst
     comandespendents.visible = True
     comandespendents.caption = "Comandes pendents. #" + atrim(rstc.RecordCount)
   End If
   hihacomplexes = False
   While Not rstc.EOF
       Set rstcextres = dbcomandes.OpenRecordset("select * from comandes_extres where comanda=" + atrim(rstc!comanda), , ReadOnly)
       llistadecomandespendents.AddItem atrim(rstc!comanda) + " v" + atrim(rstc!numordremodificacio) + " " + atrim(rstc!proximaseccio)
       If cadbl(rstc!linkcomanda1) > 0 Then hihacomplexes = True
       If amplecomanda(rstc) <> cadbl(modificacions.Recordset!amplelamina) Then camplelamina.BackColor = &H80C0FF
       If bandescomanda(rstc) <> cadbl(modificacions.Recordset!bandes) Then Text10.BackColor = &H80C0FF
       If Not rstcextres.EOF Then
          If cadbl(rstcextres!desarrollclient) > 0 And cadbl(rstcextres!desarrollclient) <> cadbl(modificacions.Recordset!desarroll) Then
             Text5.BackColor = QBColor(15)
             etdesarrollerroni.visible = True
          End If
       End If
       rstc.MoveNext
   Wend
   llistadecomandespendents.tag = IIf(hihacomplexes, "hihacomplexes", "")
   codidebarres.BackColor = QBColor(15)
   codidebarres.ToolTipText = ""
   If atrim(clixes.Recordset!codidebarres) <> "" Then
        Set rstc = dbclixes.OpenRecordset("select codidebarres from clixes where codidebarres='" + atrim(clixes.Recordset!codidebarres) + "'")
        If Not rstc.EOF Then
           rstc.MoveLast
           If rstc.RecordCount > 1 Then
               codidebarres.BackColor = &H80C0FF
               codidebarres.ToolTipText = "Hi ha codis de barres igual en altres treballs."
           End If
        End If
   End If
   Set rstc = Nothing
   Set rstcextres = Nothing
End Sub
Sub posarfotograbador()
   Dim rstf As Recordset
   Set rstf = dbclixes.OpenRecordset("select * from fotogravadors where codi=" + atrim(cadbl(modificacions.Recordset!fotograbador)))
   If Not rstf.EOF Then
      nomproveidor = atrim(rstf!nomfotogravador)
   End If
End Sub
Sub carregarnomclient()
   Dim rst As Recordset
   Dim nomc As String
   nomclient = ""
   Set rst = dbclixes.OpenRecordset("SELECT clientsvinculats.codiclient,Clientsvinculats.nomclient From Clientsvinculats where (((Clientsvinculats.id_treball)=" + atrim(id_treball) + ") AND ((Clientsvinculats.ordremodificacio)=" + atrim(ordremodificacio) + ")) and principal;")
   If Not rst.EOF Then nomc = atrim(rst!codiclient) + " - " + atrim(rst!nomclient)
   Set rst = dbclixes.OpenRecordset("SELECT Clientsvinculats.id_treball, Clientsvinculats.ordremodificacio, First(Clientsvinculats.nomclient) AS pnomclient, Count(Clientsvinculats.codiclient) AS quants From Clientsvinculats GROUP BY Clientsvinculats.id_treball, Clientsvinculats.ordremodificacio HAVING (((Clientsvinculats.id_treball)=" + atrim(id_treball) + ")) order by ordremodificacio;") ' AND ((Clientsvinculats.ordremodificacio)=" + atrim(ordremodificacio) + "))
   
   If Not rst.EOF Then
     rst.MoveLast
     rst.MoveFirst
     nomclient = "#" + atrim(rst.quants) + "# -> " + nomc   'atrim(rst!pnomclient)
   End If
   
   
   nomclientclixe = ""
   Set rst = dbclixes.OpenRecordset("SELECT Clientsvinculats.id_treball, First(Clientsvinculats.nomclient) AS pnomclient,first(clientsvinculats.codiclient) as pcodiclient, Count(Clientsvinculats.codiclient) AS quants From Clientsvinculats GROUP BY Clientsvinculats.id_treball HAVING Clientsvinculats.id_treball=" + atrim(id_treball) + ";")
   If Not rst.EOF Then
     rst.MoveLast
     rst.MoveFirst
     nomclientclixe = "#" + atrim(rst.quants) + "# -> " + atrim(rst!pnomclient) ' nomc
     If cadbl(clixes.Recordset!codiclienttemporal) = 0 Then
        clixes.Recordset.Edit
        clixes.Recordset!codiclienttemporal = cadbl(rst!pcodiclient)
        clixes.Recordset!nomclienttemporal = atrim(rst!pnomclient)
        clixes.Recordset.Update
     End If
   End If
   If nomclientclixe = "" Then nomclientclixe = nomclienttemporal
End Sub
Sub possarcoloralsbotonsquetenendades()
    Dim rst As Recordset
    possarcolorbotopressupost
    possarcolorbotocomandafotogravador
    botoliniesmodificacions.BackColor = COLORBOTOSENSEDADES
    botoliniesalbarans.BackColor = COLORBOTOSENSEDADES
    botoclientsvinculats.BackColor = COLORBOTOSENSEDADES
    bototintes.BackColor = COLORBOTOSENSEDADES
    'botocomandesfotogravador.BackColor = COLORBOTOSENSEDADES
    botorepasclixes.BackColor = COLORBOTOSENSEDADES
    If modificacions.Recordset.EOF Or modificacions.Recordset.BOF Then Exit Sub
    Set rst = dbclixes.OpenRecordset("select id_treball from clixes_modifi where id_treball=" + atrim(clixes.Recordset!id_treball) + " and ordremodificacio=" + atrim(modificacions.Recordset!ordre))
    If Not rst.EOF Then botoliniesmodificacions.BackColor = COLORBOTOAMBDADES
    Set rst = dbclixes.OpenRecordset("select id_treball from clixes_albarans where id_treball=" + atrim(clixes.Recordset!id_treball) + " and ordremodificacio=" + atrim(modificacions.Recordset!ordre))
    If Not rst.EOF Then botoliniesalbarans.BackColor = COLORBOTOAMBDADES
    Set rst = dbclixes.OpenRecordset("select id_treball from tintes where id_treball=" + atrim(clixes.Recordset!id_treball) + " and ordremodificacio=" + atrim(modificacions.Recordset!ordre))
    If Not rst.EOF Then bototintes.BackColor = COLORBOTOAMBDADES
    Set rst = dbclixes.OpenRecordset("select id_treball from clixes_modifi where id_treball=" + atrim(clixes.Recordset!id_treball) + " and ordremodificacio=" + atrim(modificacions.Recordset!ordre))
    If Not rst.EOF Then botoliniesmodificacions.BackColor = COLORBOTOAMBDADES
    Set rst = dbclixes.OpenRecordset("select id_treball from clientsvinculats where id_treball=" + atrim(clixes.Recordset!id_treball) + " and ordremodificacio=" + atrim(modificacions.Recordset!ordre))
    If Not rst.EOF Then botoclientsvinculats.BackColor = COLORBOTOAMBDADES
   ' Set rst = dbclixes.OpenRecordset("select id_treball from comandesfotogravador where id_treball=" + atrim(clixes.Recordset!id_treball) + " and ordremodificacio=" + atrim(modificacions.Recordset!ordre))
   ' If Not rst.EOF Then botocomandesfotogravador.BackColor = COLORBOTOAMBDADES
    
End Sub



Private Sub comandesfotogravador_Click()

End Sub

Private Sub comboformaimpresio_Click()
   formaimpresio.Text = Mid(comboformaimpresio, 1, 1)
   If formaimpresio.Text = "N" And llistadecomandespendents.ListCount > 0 Then
      If llistadecomandespendents.tag = "hihacomplexes" Then MsgBox "Hi ha comandes complexes pendents per aquest treball..." + Chr(10) + "SEGUR QUE VOLS ESCULLIR IMPRESIO NORMAL?", vbCritical + vbOKOnly, "AVÍS"
   End If
   If formaimpresio.Text = "T" And llistadecomandespendents.ListCount > 0 Then
      If llistadecomandespendents.tag = "" Then MsgBox "No hi ha comandes complexes pendents per aquest treball..." + Chr(10) + "SEGUR QUE VOLS ESCULLIR IMPRESIO PER TRANSPARENCIA?", vbCritical + vbOKOnly, "AVÍS"
   End If
End Sub

Private Sub comboformaimpresio_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
End Sub

Private Sub comboformaimpresio_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Command1_Click()
   gravarcanvis
End Sub
Sub gravarcanvis(Optional nocomprovardiferencies As Boolean)
  Dim bk As Long
  Dim bkclixe As Long
   framesactivats False
  If clixes.Recordset.EditMode = 0 Then Exit Sub
    clixes.Recordset!portareduccio = IIf(cportareduccio.Value = 1, True, False)
    If Not comprovarreduccions Then MsgBox "Falta algun valor dels camps de reduccio cilindre.", vbCritical, "Atenció"
 ' If Not clixes.Recordset.BOF Then bkclixe = modificacions.Recordset!id_treball
   If modificacions.Recordset.EditMode > 0 Then
     bk = modificacions.Recordset!ordre
     modificacions.Recordset.Update
     modificacions.Refresh
     reixamodificacions.Refresh
   End If
   clixes.Recordset.Update
   clixes.Recordset.Bookmark = clixes.Recordset.LastModified
   possarelcodimuntadora
'   clixes.Recordset.FindFirst "id_treball=" + atrim(bkclixe)
   modificacions.Recordset.FindFirst "ordre=" + atrim(bk)
   If Not nocomprovardiferencies Then comprovardiferenciescomandesafectades
End Sub
Sub possarelcodimuntadora()
   Dim vmuntadora As String
   vmuntadora = buscarcodimuntadora(cadbl(id_treball))
   If vmuntadora = "" Then vmuntadora = Format(id_treball, "00000000")
   dbclixes.Execute "update clientsvinculats set codimuntadora='" + atrim(vmuntadora) + "' where id_Treball=" + atrim(id_treball)
End Sub
Function buscarcodimuntadora(idtreball As Double) As String
    Dim rst As Recordset
    Set rst = dbclixes.OpenRecordset("select codimuntadora from clientsvinculats where id_Treball=" + atrim(idtreball) + " order by codimuntadora Desc")
    If rst.EOF Then Exit Function
    buscarcodimuntadora = atrim(rst!codimuntadora)
End Function
Function comprovarreduccions() As Boolean
    comprovarreduccions = True
    If cadbl(reducciopermetre) = 0 And (cadbl(reducciocilindrefw) > 0 Or cadbl(reducciocilindref2) > 0) Then
       comprovarreduccions = False
    End If
    If cadbl(reducciopermetre) > 0 And (cadbl(reducciocilindrefw) = 0 And cadbl(reducciocilindref2) = 0) Then
       comprovarreduccions = False
    End If
End Function
Private Sub Command11_Click()
   crear_modificacio_nova
End Sub
Sub crear_modificacio_nova()
  Dim numclixe As Integer
  Dim rstmodant As Recordset
  If clixes.Recordset.EditMode > 0 Then MsgBox "Estas editant primer finalitza la operació i despres afegeix.", vbCritical, "Atenció": Exit Sub
  If etestatclixe <> "CLIXES ENTRATS" And Not modificacions.Recordset.BOF Then
     If Not modificacioautomatica Then
        MsgBox "No pots donar d'alta una modificació nova si la anterior encara no té els CLIXES ENTRATS.", vbCritical, "Atenció"
     End If
     Exit Sub
  End If
   numclixe = clixes.Recordset!id_treball
   clixes.RecordSource = "select * from clixes order by id_treball DESC"
   clixes.Refresh
   clixes.Recordset.FindFirst "id_Treball=" + atrim(numclixe)
   clixes.Recordset.Edit
   Set rstmodant = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(numclixe) + " order by ordre desc")
   comprovarsihihaavisosperaquestclient clixes.Recordset!codiclienttemporal, "C"
   modificacions.Recordset.AddNew
   ordre = ordremesgran + 1
   modificacions.Recordset!id_treball = cadbl(clixes.Recordset!id_treball)
   modificacions.Recordset!ordre = cadbl(ordre)
   dataobertura = Format(Now, "dd/mm/yy")
   vnumerodeliniaimpresioanterior = cadbl(rstmodant!numerodelinia)
   If vnumerodeliniaimpresioanterior > 0 Then
       If MsgBox("Vols conservar la Linia d'Impresió assignada a la modificació anterior?", vbInformation + vbYesNo + vbDefaultButton1, "Atenció") = vbNo Then
         vnumerodeliniaimpresioanterior = 0
       End If
   End If
   modificacions.Recordset!formaimpresio = atrim(rstmodant!formaimpresio)
   modificacions.Recordset!bandes = cadbl(rstmodant!bandes)
   modificacions.Recordset!gruixpolimer = cadbl(rstmodant!gruixpolimer)
   modificacions.Recordset!amplelamina = cadbl(rstmodant!amplelamina)
   modificacions.Recordset!sistemadimpresio = atrim(rstmodant!sistemadimpresio)
   modificacions.Recordset!fotograbador = cadbl(rstmodant!fotograbador)
   modificacions.Recordset!numerodelinia = vnumerodeliniaimpresioanterior
   modificacions.Recordset!valordeltamaxim = cadbl(rstmodant!valordeltamaxim)
   If cadbl(modificacions.Recordset!valordeltamaxim) = 0 Then modificacions.Recordset!valordeltamaxim = 2.5
   framesactivats True
   modificacions.Recordset.Update
   possarestatclixe
   modificacions.Recordset.Bookmark = modificacions.Recordset.LastModified
   modificacions.Recordset.Edit
   If Not modificacioautomatica Then
      dataobertura.SetFocus
      If llistadecomandespendents.ListCount > 0 Then MsgBox "ATENCIÓ QUE HI HA COMANDES PENDENTS." + Chr(10) + "S'HAURIA DE REVISAR SI ALGUNA D'ELLES QUEDA AFECTADA PER AQUESTA VERSIÓ NOVA.", vbCritical, "ATENCIÓ"
   End If
   Set rstmodant = Nothing
   
End Sub
Sub comprovarsihihaavisosperaquestclient(vnumclient As Double, vllocavis As String)
    Dim rst As Recordset
    Set rst = dbclixes.OpenRecordset("select * from avisos_clixes where llocavis='" + atrim(vllocavis) + "' and NUMCLIENT=" + atrim(vnumclient))
    If rst.EOF Then GoTo fi
    MsgBox UCase(rst!descripcio_avis), vbExclamation, "A T E N C I Ó"
    'If rst!tipus = "NT" Then MsgBox UCase(rst!descripcio_avis), vbExclamation, "A T E N C I Ó"
fi:
    Set rst = Nothing
End Sub

Private Sub Command3_Click()
     Dim resp As String
     resp = InputBox("Entra el valor de Distorsió per metre:" + Chr(10) + "RECORDA QUE SI EL VALOR ES POSITIU SERA AUGMENT I NEGATIU REDUCCIÓ", "DISTORSIÓ DEL CILINDRE")
     resp = atrim(cadbl(resp))
     reducciopermetre = resp
     If cadbl(resp) >= 0 And cadbl(reducciocilindrefw) < 0 Then MsgBox "El valor de factor per la FW no es correcte", vbCritical, "Error FW": reducciocilindrefw = "0"
     If cadbl(resp) < 0 And cadbl(reducciocilindrefw) > 0 Then MsgBox "El valor de factor per la FW no es correcte", vbCritical, "Error FW": reducciocilindrefw = "0"
     If cadbl(resp) >= 0 And cadbl(reducciocilindref2) < 0 Then MsgBox "El valor de factor per la F2 no es correcte", vbCritical, "Error F2": reducciocilindref2 = "0"
     If cadbl(resp) < 0 And cadbl(reducciocilindref2) > 0 Then MsgBox "El valor de factor per la F2 no es correcte", vbCritical, "Error F2": reducciocilindref2 = "0"
End Sub
Sub afegirlinianova(resp As String)
     resp = treure_apostruf(resp)
     Set rst = dbclixes.OpenRecordset("select * from linies where  id_marca=" + atrim(clixes.Recordset!id_marca) + " and linia='" + resp + "'")
     If Not rst.EOF Then Exit Sub 'MsgBox "Aquesta linia ja està donada d'alta.", vbExclamation, "Error"
     dbclixes.Execute "insert into linies (id_marca,linia) values (" + atrim(clixes.Recordset!id_marca) + ",'" + resp + "')"
     Set rst = dbclixes.OpenRecordset("select * from linies where  id_marca=" + atrim(clixes.Recordset!id_marca) + " and linia='" + resp + "'")
     If Not rst.EOF Then  'busco la nova linia i la posso al box
         liniaproducte = rst!linia
         clixes.Recordset!id_linia = rst!id_linia
     End If
End Sub
Private Sub Command4_Click()
   borrardataclixe "databaixaclixe"
End Sub
Sub borrardataclixe(camp As String)
  gravarcanvis
  dbclixes.Execute "update clixes set " + camp + "=null where id_treball=" + atrim(clixes.Recordset!id_treball)
  clixes.UpdateControls
  modificar_Click
  
End Sub
Sub borrardatamodificacio(camp As String)
  gravarcanvis
  dbclixes.Execute "update modificacions set " + camp + "=null where id_treball=" + atrim(clixes.Recordset!id_treball) + " and ordre=" + atrim(cadbl(modificacions.Recordset!ordre))
  clixes.UpdateControls
  modificacions.UpdateControls
  modificar_Click
  
End Sub

Private Sub Command5_Click()
  borrardatamodificacio "dataobertura"
End Sub

Private Sub Command6_Click()
borrardatamodificacio "datavalidaciotexte"
End Sub

Private Sub Command7_Click()
borrardatamodificacio "datavalidaciomides"
End Sub

Private Sub Command8_Click()
   borrardatamodificacio "datavalidaciocolors"
End Sub

Private Sub controldrag_Timer()
End Sub

Private Sub consultar_Click()
  fbuscar.Show 1
End Sub

Private Sub cportareduccio_Click()
   If cadbl(reducciocilindref2) <> 0 Or cadbl(reducciocilindrefw) <> 0 Then MsgBox "Hi ha valors entrats a reducció vigila que sigui correcte.", vbCritical, "Atenció"
End Sub

Private Sub eliminar_Click()
Dim resp As String
  resp = UCase(InputBox("Estas segur que vols eliminar aquest CLIXE o TREBALL?" + Chr(10) + "AIXÒ TAMBÉ ELIMINARÀ ELS IMPS I PDF'S RELACIONATS I CLIXES COMPARTITS DE TOTES LES MODIFICACIONS." + Chr(10) + "Escriu [eliminartreballimodificacions] per eliminar-la.", "Eliminar CLIXE I MODIFICACIONS"))
  If resp = "ELIMINARTREBALLIMODIFICACIONS" Then
     resp = UCase(InputBox("Estas segur que vols eliminar aquest CLIXE o TREBALL?" + Chr(10) + "AIXÒ TAMBÉ ELIMINARÀ ELS IMPS I PDF'S RELACIONATS I CLIXES COMPARTITS DE TOTES LES MODIFICACIONS." + Chr(10) + "Escriu [ELIMINARPDFSIELSIMPS] per eliminar-la.", "Eliminar CLIXE I MODIFICACIONS"))
     If UCase(resp) = "ELIMINARPDFSIELSIMPS" Then eliminartreball cadbl(id_treball)
  End If
End Sub
Sub eliminartreball(treball As Integer)
   Dim rst As Recordset
   
   Set rst = dbclixes.OpenRecordset("select ordre from modificacions where id_treball=" + atrim(treball))
   While Not rst.EOF
      eliminarmodificacio cadbl(treball), rst!ordre
      rst.MoveNext
   Wend
   dbclixes.Execute "delete * from clixes where id_treball=" + atrim(treball)
   clixes.Refresh
      clixes.Recordset.FindFirst "id_treball<" + atrim(treball)
End Sub

Private Sub etestatclixe_Change()
    If IsNumeric(Mid(etestatclixe, 1, 1)) Then
          etestatclixe.ForeColor = QBColor(12)
         Else: etestatclixe.ForeColor = &H8000000D
    End If
    If InStr(1, etestatclixe, "CLIXES REBUTS") > 0 Then etestatclixe.ForeColor = QBColor(10)
End Sub


Sub comprovardiferenciescomandesafectades()
   Dim i As Integer
   Dim numcomanda As String
   For i = 0 To llistadecomandespendents.ListCount - 1
       numcomanda = Mid(llistadecomandespendents.List(i), 1, 7)
       If mirardiferenciescomandaitreball(cadbl(numcomanda)) Then
           imprimirdiferenciescomandaitreball cadbl(numcomanda)
           If UCase(InputBoxEx("La comanda " + numcomanda + " te diferencies amb aquest treball." + Chr(10) + "VOLS ACTUALITZAR LA COMANDA AMB LES DADES D'AQUEST TREBALL? (Escriu [si] per fer-ho.)", "Atenció")) = "SI" Then
               posardiferenciesacomandadeltreball cadbl(numcomanda)
           End If
       End If
   Next i
     
End Sub

Private Sub Form_Activate()
  Static jaheentrat As Boolean
  If jaheentrat Then Exit Sub

  If Command13.Enabled And Command13.visible Then Command13.SetFocus
  
  jaheentrat = True
  If modificacioautomatica Then crearmodificacioautomatica cadbl(arguments(2))
  If llistatclixesvells Then
    mclixesimpresio_Click
    End
  End If
  If baixaclixes Then actualitzardadesimpresors cadbl(arguments(2)): End
  If imprimirbossasoldadores Then formclixes.visible = False: imprimiretiquetabossasoldadores atrim(arguments(4)), llistat, False: End
  If modificartintes Then
     formclixes.visible = False
     clixes.Recordset.FindFirst "id_treball=" + atrim(arguments(4))
     If clixes.Recordset.NoMatch Then Exit Sub
     modificacions.Recordset.FindFirst "ordre=" + atrim(arguments(5))
     If modificacions.Recordset.NoMatch Then Exit Sub
     bototintes_Click
     End
  End If
  DoEvents
  dbclixes.Execute "UPDATE tintes SET tintes.desarroll = 0, tintes.cilindre = 0 WHERE (((tintes.color)='') AND ((tintes.desarroll)>0));"
  If Not existeix("c:\ordprog.ini") Then comprovarlotsclixessenseimp
  comprovarcomandessenseenviar
  comprovartreballsambpressupostisenseliniesdalbara
  comprovaravisosreposicionscaducades
  jaheentrat = False
    If Command13.Enabled And Command13.visible Then Command13.SetFocus
End Sub
Sub comprovaravisosreposicionscaducades()
   Dim rst As Recordset
   If cadbl(llegir_ini("General", "comprovarreposicionscaducades", "comandes.ini")) = Day(Now) Then Exit Sub
   Set rst = dbclixes.OpenRecordset("SELECT Clixes.marca, Clixes.linia, Clixes.estatclixe, reposicionsfotogravador.id_treball, reposicionsfotogravador.dataenviament, reposicionsfotogravador.datateoricarecepcio FROM Clixes INNER JOIN reposicionsfotogravador ON Clixes.id_treball = reposicionsfotogravador.id_treball  WHERE (((Clixes.estatclixe)='REPOSICIÓ DEL CLIXE') AND ((reposicionsfotogravador.dataenviament)>=Date()));")
   If rst.EOF Then GoTo fi
   'creo el fitxer de cos de missatge
   Open "c:\temp\cosmissatge.txt" For Output As #2
   Print #2, " "
   Print #2, " "
   Print #2, "Relació de reposicions que ja haurien d'haver arribat."
   Print #2, " "
   Print #2, " "
   While Not rst.EOF
     Print #2, atrim(rst!id_treball) + "  " + atrim(rst!marca) + " - " + atrim(rst!linia) + "            Data teorica: " + atrim(rst!datateoricarecepcio)
     rst.MoveNext
   Wend
   Close #2
   enviar_reposicions
fi:
   escriure_ini "General", "comprovarreposicionscaducades", atrim(Day(Now)), "comandes.ini"
End Sub
Sub enviar_reposicions()
  Dim usuarim As String
  Dim contrasenyam As String
  usuarim = llegir_ini("Enviomails", "usuari", "comandes.ini")
  contrasenyam = llegir_ini("Enviomails", "contrasenya", "comandes.ini")
  If usuarim = "{[}]" Or contrasenyam = "{[}]" Then Exit Sub
   
   
    Set objMessage = CreateObject("CDO.Message")
    objMessage.Subject = "Reposicions amb plaç d'entrega passat."
    objMessage.from = usuarim
    objMessage.To = "mkinplacsa@inplacsa.com"
    '    objmessage.To = "miquel.inplacsa@gmail.com"
    objMessage.TextBody = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\temp\cosmissatge.txt", 1).ReadAll
    'objMessage.AddAttachment formenviomails.nomfitxeradjunt
    
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = usuarim
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = contrasenyam
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    objMessage.Configuration.Fields.Update
    
    objMessage.Send
    '==End remote SMTP server configuration section==
    'If cadbl(objMessage.Send) = 0 Then enviaremail2 = True
    'ratoli "normal"
    'fenviant.visible = False
    'formcomandaclixes.Enabled = True
End Sub


Sub comprovartreballsambpressupostisenseliniesdalbara()
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("SELECT pressupostos.id_treball, pressupostos.ordremodificacio, pressupostos.numpressupost, Clixes_albarans.num_alb, pressupostos.enviat, Clixes.estatclixe FROM (pressupostos LEFT JOIN Clixes_albarans ON (pressupostos.id_treball = Clixes_albarans.id_treball) AND (pressupostos.ordremodificacio = Clixes_albarans.ordremodificacio)) LEFT JOIN Clixes ON pressupostos.id_treball = Clixes.id_treball WHERE (((Clixes_albarans.num_alb) Is Null) AND ((pressupostos.enviat)=True) AND ((Clixes.estatclixe)='CLIXES ENTRATS'));")
   If Not rst.EOF Then
      botoavisosliniesalbarans.visible = True
        Else: botoavisosliniesalbarans.visible = False
   End If
   Set rst = Nothing
   End Sub
Sub comprovarcomandessenseenviar()
   Dim rst As Recordset
   Dim msg As String
   Set rst = dbclixes.OpenRecordset("select * from comandesfotogravador where okenviat=false")
   While Not rst.EOF
      msg = msg + IIf(msg <> "", ",", "") + atrim(rst!id_treball) + "/" + atrim(rst!ordremodificacio)
      rst.MoveNext
   Wend
   If msg <> "" Then MsgBox "Comandes a fotogravador sense enviar:" + Chr(10) + msg, vbInformation, "Atenció"
   Set rst = Nothing
End Sub
Sub comprovarlotsclixessenseimp()
   Static jahaviaentrat As Boolean
   Dim rstc As Recordset
   Dim rstt As Recordset
   Dim numordre As String
   Dim msgcomandes As String
   If jahaviaentrat Then Exit Sub
   ratoli "espera"
   Set rstc = dbcomandes.OpenRecordset("SELECT direnvio,numordremodificacio,producte,comandes.comanda, comandes.numtreball, comandes.proximaseccio from comandes WHERE (((comandes.numtreball)>0) AND ((comandes.proximaseccio)<>'P' And (comandes.proximaseccio)<>'T' And (comandes.proximaseccio)<>'V'));")
   While Not rstc.EOF
    If larutahiha(atrim(rstc!producte), "I") Then
     If estatclixemod(rstc!numtreball, cadbl(rstc!numordremodificacio)) = "CLIXES ENTRATS" Then
      numordre = cadbl(rstc!numordremodificacio)
      If numordre = 0 Then numordre = 1
      Set rstt = dbclixes.OpenRecordset("select clixes.id_treball,clixes.marca,clixes.linia,clientsvinculats.arxiuimp,clientsvinculats.nomclient from clientsvinculats INNER JOIN Clixes ON Clientsvinculats.id_treball = Clixes.id_treball where clientsvinculats.id_treball=" + atrim(rstc!numtreball) + " and ordremodificacio=" + atrim(numordre) + " and direnvio=" + atrim(cadbl(rstc!direnvio)))
      If Not rstt.EOF Then
        If Not rstt!arxiuimp Then msgcomandes = msgcomandes + atrim(rstc!comanda) + "---> Nº Treball: " + atrim(rstt!id_treball) + Chr(10) + atrim(rstt!nomclient) + Chr(10) + atrim(rstt!marca) + " - " + atrim(rstt!linia) + Chr(10) + " -------------------------" + Chr(10)
          'Else: Stop
      End If
     End If
    End If
     rstc.MoveNext
   Wend
   If atrim(msgcomandes) <> "" Then MsgBox "Hi han comandes pendents amd CLIXES ENTRATS i que no tenen IMP entrat." + Chr(10) + msgcomandes
   jahaviaentrat = True
   ratoli "normal"
End Sub


Private Sub Form_Click()
  Dim v1fw As Double
  Dim v2fw As Double
  Dim v1f2 As Double
  Dim v2f2 As Double
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim vult As Integer
  Dim vtreballsrepes As String
  
  'passar_avis_reposicio_a_inplacsa
  Exit Sub
  Set rst = dbclixes.OpenRecordset("select * from modificacions")
  While Not rst.EOF
    Set rst2 = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(rst!id_treball) + " and ordremodificacio=" + atrim(rst!ordre) + " order by ordretinter")
    vult = 0
    While Not rst2.EOF
       If vult = cadbl(rst2!ordretinter) Then vtreballsrepes = vtreballsrepes + " , " + atrim(rst2!id_treball) + "/" + atrim(rst2!ordremodificacio): rst2.MoveLast
       vult = cadbl(rst2!ordretinter)
       rst2.MoveNext
    Wend
    rst.MoveNext
  Wend
'  Clipboard.Clear
'  Clipboard.SetText vtreballsrepes
  Set rst = Nothing
  Set rst2 = Nothing
 ' Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
 '  calcular_mtrsminut id_treball, ordremodificacio, v1fw, v2fw, v1f2, v2f2
 ' MsgBox atrim(v1fw) + " - " + atrim(v2fw) + " / " + atrim(v1f2) + " - " + atrim(v2f2)
End Sub
Sub calcular_mtrsminut(vidtreball As Long, vordre As Long, v1fw As Double, v2fw As Double, v1f2 As Double, v2f2 As Double)
  Dim rst As Recordset
  Dim vsql As String
  vsql = "SELECT impressores.numeromaquina as nummaq, Avg(impressores.mtrsminut) AS mitjana, Max(impressores.mtrsminut) AS maxim FROM comandes INNER JOIN impressores ON comandes.comanda = impressores.comanda Where "
  vsql = vsql + " (((comandes.numtreball) = " + atrim(vidtreball) + ") And ((comandes.numordremodificacio) = " + atrim(vordre) + ") And ((impressores.tipus) = 'F')) GROUP BY impressores.numeromaquina HAVING (((Avg(impressores.mtrsminut))<>0));"
  Set rst = dbbaixes.OpenRecordset(vsql)
  While Not rst.EOF
    If rst!nummaq = 7 Then v1fw = cadbl(rst!mitjana): v2fw = cadbl(rst!maxim)
    If rst!nummaq = 9 Then v1f2 = cadbl(rst!mitjana): v2f2 = cadbl(rst!maxim)
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then
   gravarcanvis
 End If
 If KeyCode = 27 Then
      cancelar_canvis
 End If
End Sub
Sub cancelar_canvis()
If modificacions.Recordset.EditMode > 0 Then modificacions.Recordset.CancelUpdate
      If clixes.Recordset.EditMode > 0 Then clixes.Recordset.CancelUpdate
      framesactivats False
End Sub
Function rutapdftreball(Optional anterior As Boolean, Optional separacio As Boolean, Optional cingularreal2 As Boolean, Optional previ As Boolean, Optional previSC As Boolean)
   Dim ordre As Integer
   On Error Resume Next
   ordre = ordremodificacio
   If anterior And ordre > 1 Then ordre = ordre - 1
    MkDir ruta_documentacio_clixes + "\" + Format(id_treball, "00000")
    MkDir ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF"
    rutapdftreball = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordre, "000") + IIf(separacio, "_SC", "") + ".pdf"
    If cingularreal2 Then rutapdftreball = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordre, "000") + "_CR" + ".pdf"
    If previ Then rutapdftreball = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordre, "000") + "_PR" + ".pdf"
    If previSC Then rutapdftreball = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordre, "000") + "_PRSC" + ".pdf"
End Function
Sub guardarelpdf(rutaorigenpdf As String, Optional tipuspdf As String, Optional vnomissatges As Boolean)
  Dim rutadesti As String
  Dim sobreescriure As Boolean
  Dim datapdf As String
  On Error GoTo erroricontinua
  timerdrag.Enabled = False
  If Not existeix(rutaorigenpdf) Then
     If Not vnomissatges Then MsgBox "Error... no trobo el fitxer PDF"
     Exit Sub
  End If
  rutadesti = rutapdftreball(False, IIf(tipuspdf = "SC", True, False), IIf(tipuspdf = "CR", True, False), IIf(tipuspdf = "PR", True, False), IIf(tipuspdf = "PRSC", True, False))
  If existeix(rutadesti) Then
      If cadbl(InputBox("Aquest treball ja te un PDF, VOLS SOBREESCRIURE'L?" + Chr(10) + " Escriu el numero de treball per comfirmar-ho", "Sobrescriure el PDF")) = id_treball Then
              sobreescriure = True
         Else:
            Kill (rutaorigenpdf): inicidragover = 0: timerdrag.Enabled = False: Exit Sub
      End If
    Else: crearcarpeta ruta_documentacio_clixes + "\" + Format(id_treball, "00000")
  End If
  obrir_document rutaorigenpdf
  wait 2
  If tipuspdf = "SC" Or tipuspdf = "CR" Or tipuspdf = "PR" Or tipuspdf = "PRSC" Then
     If MsgBox("Es correcte aquest pdf que vols assignar a aquest treball?" + Chr(10) + "ASSEGURA QUE HAS TANCAT EL PDF ABANS DE RESPONDRE", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then GoTo fi
     GoTo sensedata
  End If
  'If MsgBox("Es correcte aquest pdf que has visualitzat?", vbInformation + vbYesNo + vbDefaultButton2, "Comfirmació") = vbNo Then timerdrag.Enabled = False: Exit Sub
  datapdf = demanardatapdf
  If datapdf = "null" Then Exit Sub
sensedata:
  If tipuspdf = "PR" Then passar_estatrevisiotintes_a_DISSENY
  If sobreescriure Then
     Kill (rutadesti)
     If tipuspdf = "CR" Then eliminar_elpng cadbl(id_treball), cadbl(ordremodificacio)
  End If
  Copiar_Fitxer rutaorigenpdf, rutadesti
  If tipuspdf = "N" Then
     
       'copio el pdf amb els dos noms de fitxer
        If existeix(nomfitxerpdfeditable(rutaorigenpdf)) Then Kill nomfitxerpdfeditable(rutaorigenpdf)
        Copiar_Fitxer rutaorigenpdf, nomfitxerpdfeditable(rutadesti)
        'miro si supera 5Mb i el marco per convertir-lo perquè no pesi tan
        If (FileLen(nomfitxerpdfeditable(rutadesti)) / 1000) >= 5000 Then
           dbclixes.Execute "update modificacions set pdfperconvertir=true where id_treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio)
        End If
           
       '------
        dbclixes.Execute "update modificacions set datapdf=" + datapdf + " where id_treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio)
        dbclixes.Execute "update modificacions set pdfvalid=true where id_treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio)
  End If
fi:
  timerdrag.Enabled = False:
  Exit Sub
erroricontinua:
timerdrag.Enabled = False:
End Sub
Sub passar_estatrevisiotintes_a_DISSENY()
   dbclixes.Execute "update modificacions set estatrevisiotintes='DISSENY' WHERE id_treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio)
End Sub
Function demanardatapdf() As String
   
    demanardatapdf = InputBox("Entra la data de validesa del pdf." + Chr(10) + "   Es la data que surt dins al pdf a la casella de data." + Chr(10) + "   SI FAS CANCELAR O NO POSES DATA NO ES GUARDARÀ EL PDF." + Chr(10) + "Format dd/mm/yy o escriu ddmmyy", "Data PDF")
    If Len(demanardatapdf) = 6 Then demanardatapdf = Mid(demanardatapdf, 1, 2) + "/" + Mid(demanardatapdf, 3, 2) + "/" + Mid(demanardatapdf, 5, 2)
    If IsDate(demanardatapdf) Then
       demanardatapdf = "#" + Format(demanardatapdf, "mm/dd/yy") + "#"
      Else: demanardatapdf = "null"
    End If
    
End Function

Sub crearcarpeta(carpetac As String)
   On Error Resume Next
   MkDir carpetac
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
End Sub

Private Sub Form_Load()

arguments = ObtenerLíneaComando
fitxerini = "comandes.ini"
If atrim(arguments(1)) <> "" Then fitxerini = atrim(arguments(1))
  cami = llegir_ini("General", "cami", fitxerini)
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
  
  '"c:\misdoc~1\commandes\comandes.mdb"
  If existeix("c:\ordprog.ini") Then cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
  inicidragover = 0
  hora = Now
  centerscreen Me
  camiclixes = rutadelfitxer(cami) + "clixesnous.mdb"
  clixes.DatabaseName = camiclixes
  Set dbclixesvells = OpenDatabase(rutadelfitxer(cami) + "clixes.mdb")
  modificacions.DatabaseName = camiclixes
  modificacions.RecordSource = ""
  Set dbclixes = DBEngine.OpenDatabase(camiclixes)
  Set dbcomandes = DBEngine.OpenDatabase(cami)
  ultimtinter = 99
  If arguments(3) = "llistatclixesvells" Then llistatclixesvells = True
  If arguments(3) = "novamodificacio" Then modificacioautomatica = True
  If arguments(3) = "baixaclixes" Then baixaclixes = True
  If arguments(3) = "modificartintes" Then modificartintes = True
  If arguments(3) = "imprimirbossasoldadores" Then imprimirbossasoldadores = True:  Form_Activate
  If cadbl(arguments(2)) > 0 And (arguments(3) = "" Or arguments(3) = "baixaclixes") Then
     clixes.RecordSource = "select * from clixes where id_treball=" + atrim(cadbl(arguments(2)))
     Frame3.Enabled = False
     'm_manteniments.Enabled = False
     'mllistats.Enabled = False
      Else: clixes.RecordSource = "select * from clixes order by id_treball DESC"
  End If
  If Not existeix("c:\ordprog.ini") Then arreglarlesquejashanfet
    comprovaravisos
 ' MsgBox "ENCARA HAS DE FER SERVIR EL PROGRAMA DE CLIXES VELLS"
 ' End
 
fi:
End Sub

Sub crearmodificacioautomatica(ntreball As Double)
   formclixes.visible = False
   
   escriure_ini "General", "creantmodificacio", "si", "clixes.ini"
   clixes.RecordSource = "select * from clixes where id_treball=" + atrim(cadbl(arguments(2)))
   clixes.Refresh
   wait (1)
   crear_modificacio_nova
   escriure_ini "General", "creantmodificacio", atrim(modificacions.Recordset!ordre), "clixes.ini"
   End
End Sub
Sub comprovaravisos()
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("Avisoscomandessenseidtreball")
   If Not rst.EOF Then
       botoavisos.BackColor = QBColor(12)
        Else: botoavisos.BackColor = QBColor(15)
   End If
End Sub
Function demanarcomplertoseparaciocolor() As String
   Unload formtipuspdf
   formtipuspdf.Show 1
   demanarcomplertoseparaciocolor = formtipuspdf.tag
   Unload formtipuspdf
End Function
Sub esperar10segonsaveuresientraelfitxer()
   Dim fitxer As String
   Dim tipuspdf As String
   If DateDiff("s", inicidragover, Now) < 12 Then
      fitxer = mirarsihihaalgualtemp
         Else: inicidragover = 0: timerdrag.Enabled = False
   End If
   If fitxer <> "" Then
      AppActivate Me.caption
      tipuspdf = demanarcomplertoseparaciocolor
      guardarelpdf "c:\temp\tmpclixes\" + fitxer, tipuspdf
      If existeix("c:\temp\tmpclixes\" + fitxer) Then
        On Error GoTo erroralborrar
        If tipuspdf = "CR" Then eliminar_elpng cadbl(id_treball), cadbl(ordremodificacio)
        If tipuspdf = "PRSC" Or tipuspdf = "PR" Then
          If existeix(ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordremodificacio, "000") + "_PR" + ".pdf") Then
            If existeix(ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordremodificacio, "000") + "_PRSC" + ".pdf") Then
              enviar_revisio_previ
            End If
          End If
        End If
        Kill "c:\temp\tmpclixes\" + fitxer
      End If
      carregar_modificacio
   End If
   Exit Sub
erroralborrar:
   If MsgBox("Error al borrar el pdf temporal potser el tens obert." + Chr(10) + "Vols tornar a provar?", vbCritical + vbDefaultButton1 + vbYesNo, "Error") = vbYes Then
      Resume
       Else: Resume Next
   End If
End Sub
Function mirarsihihaalgualtemp() As String
   Dim fitxer As String
   fitxer = Dir("c:\temp\tmpclixes\*.*")
   Do While fitxer <> ""   ' Start the loop
     mirarsihihaalgualtemp = fitxer
    DoEvents
    fitxer = Dir()   ' Get next entry.
   Loop
   DoEvents
End Function
Sub obrirtemporalclixes(noobrirexplorer As Boolean)
  Dim idp As Long
  If existeix("c:\temp\tmpclixes") Then
    CreateObject("Scripting.FileSystemObject").DeleteFolder "c:\temp\tmpclixes"           ' Eliminamos carpeta.
  End If
  On Error Resume Next
  If Not existeix("c:\temp\tmpclixes") Then MkDir "c:\temp\tmpclixes"
  Kill "c:\temp\tmpclixes\*.*"
  If noobrirexplorer Then Exit Sub
  idp = ShellExecute(Me.hWnd, "Open", "c:\windows\explorer.exe", " c:\temp\tmpclixes", "", 1)  'Shell("explorer c:\temp\tmpclixes", vbNormalFocus)
  'hwd = GetWinHandle(idp)
  'hwd = InstanceToWnd(idp)
 ' MsgBox hwd
End Sub

Private Sub formaimpresio_Change()
   If Not (formaimpresio = "T" Or formaimpresio = "N") Then comboformaimpresio = "": Exit Sub
   comboformaimpresio = IIf(formaimpresio = "T", "Transparencia", "Normal")
End Sub


Private Sub id_treball_Change()

End Sub
Sub imprimir_VQ_soldadores(rst As Recordset)
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim rstc As Recordset
  Dim rstcli As Recordset
  Dim vnumtreballiversio As String
  Dim vtexteimpresio As String
  Set rstc = dbcomandes.OpenRecordset("select numerobossasoldadores from comandes_extres where comanda=" + atrim(rst!comanda))
  Set rstcli = dbcomandes.OpenRecordset("select * from clients where codi=" + atrim(rst!client))
  If Mid(atrim(rstc!numerobossasoldadores) + " ", 1, 1) <> "C" Then
       vnumtreballiversio = atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio)
       vtexteimpresio = atrim(rst!marcailinia)
  End If
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "verificacioqualitatVQsoldadores.rpt", 1)
  oreport.FormulaFields.GetItemByName("numbossa").Text = "'NºBossa: " + atrim(rstc!numerobossasoldadores) + "'"
  oreport.FormulaFields.GetItemByName("nomclient").Text = "'" + atrim(rstcli!nom) + "'"
  oreport.FormulaFields.GetItemByName("treballversio").Text = "'" + atrim(vnumtreballiversio) + "'"
  oreport.FormulaFields.GetItemByName("texteimpresio").Text = "'" + atrim(vtexteimpresio) + "'"
  oreport.FormulaFields.GetItemByName("refclient").Text = "'" + atrim(rst!refclient) + "'"
  oreport.FormulaFields.GetItemByName("comanda").Text = "'" + atrim(rst!comanda) + "'"
   oreport.DiscardSavedData
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   If Not existeix("c:\ordprog.ini") Then
          veurereport.CRViewer.PrintReport
          oreport.PrintOut False
         Else:
           veurereport.CRViewer.ViewReport
           veurereport.WindowState = 2
           veurereport.Show 1
   End If
  Set rstc = Nothing

End Sub

Private Sub Frame3_Click()
 'imprimiretiquetabossasoldadores 188315, llistat, False
 
End Sub

Private Sub frameclixes_Click()
  'Dim rst As Recordset
  'Set rst = dbclixes.OpenRecordset("select * from modificacions where datapdf>#" + Format(DateAdd("yyyy", -2, Now), "mm/dd/yy") + "#")
  'rst.MoveLast
  'rst.MoveFirst
  'While Not rst.EOF
  '   id_treball = rst!id_treball
  '   ordre = rst!ordre
  '   If existeix(rutapdftreball(, True, True)) Then
  '     rst.Edit: rst!observacionsfacturaclixes = "1": rst.Update
  '   End If
  '   Me.caption = "  " + atrim(rst.AbsolutePosition) + "/" + atrim(rst.RecordCount)
  '   rst.MoveNext
  'Wend
  'Set rst = Nothing
End Sub

Private Sub imprimirbossaclixe_Click()
   Dim vXL As String
   If comandespendents.visible Then demanariposarXL: clixes.Recordset.Update: wait 1
   If atrim(Text2) = "" Then If MsgBox("Aquest treball no te XL assignat." + vbNewLine + "VOLS IMPRIMIR L'ETIQUETA IGUALMENT?", vbExclamation + vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbNo Then Exit Sub
   imprimiretiquetabossaclixesdemanantimpresora clixes.Recordset!id_treball, modificacions.Recordset!ordre, llistat, False
   
End Sub
Sub demanariposarXL()
If modificacions.Recordset.EOF Then Exit Sub
   If Text2 = "" Then
     vXL = demanararxiu(True)
     If vXL <> "" Then
        Text2 = atrim(vXL)
         Else: Exit Sub
     End If
   End If
End Sub
Function demanararxiu(Optional vsuggerirXL As Boolean) As String
  Dim vcontrol As Control
  Dim vnumXLsuggerit As String
  vnumXLsuggerit = suggerirXL
  Unload formescullirlleixa
  Load formescullirlleixa
  formescullirlleixa.cxl = vnumXLsuggerit
  formescullirlleixa.cxl.tag = "1"
  formescullirlleixa.Top = formclixes.Top + ((formclixes.Height - formescullirlleixa.Height) / 2)
  formescullirlleixa.Left = formclixes.Left + ((formclixes.width - formescullirlleixa.width) / 2)
  'formescullirlleixa.Left = formscrooll.Left + formcomandes.Left
  formescullirlleixa.Show 1
  If seleccioret = 1 Then demanararxiu = formescullirlleixa.valorescullit
  
End Function
Function suggerirXL() As String
  Dim rst As Recordset
  Dim vamplada As Double
  Dim vXL As String
  Dim vQtreballsXRbossa As Integer
  Dim vcriteri As String
  vQtreballsXRbossa = 7
  While vamplada <> 55 And vamplada <> 64 And vamplada <> 80
    vamplada = cadbl(InputBox("Entra la emplada de la bossa amb CMs (55cm, 64cm o 80cm)", "Ampla de la bossa XL"))
  Wend
  vXL = " instr(1,trim(Clixes.arxiu),'XL')>0 and isnull(Clixes.databaixaclixe) and instr(1,trim(Clixes.ubicacio),'P')=0 "
  
  'BOSSA PETITA
  If vamplada = 55 Then vcriteri = "(NumXL>=55 and NumXL<=72) or (NumXL>=136 and NumXL<=262)"
  'BOSSA GRAN
  If vamplada = 80 Then vcriteri = "(NumXL>=1 and NumXL<=54) or (NumXL>=73 and NumXL<=90) or (NumXL>=484 and NumXL<=492)"
  'BOSSA NORMAL
  If vamplada = 64 Then vcriteri = "(not (NumXL>=55 and NumXL<=72) and not ((NumXL>=1 and NumXL<=54) or (NumXL>=73 and NumXL<=90) or (NumXL>=484 and NumXL<=492)))"
  vcriteri = " HAVING " + substituirtot(vcriteri, "NumXL", "(CDbl(Mid(Trim(First(Clixes.arxiu)),4)))")
  Set rst = dbclixes.OpenRecordset("SELECT count(Clixes.arxiu) as Q,first(clixes.arxiu) as Parxiu, cdbl(mid(trim(first(Clixes.arxiu)),4)) as NumXL FROM Clixes WHERE " + vXL + " GROUP BY Clixes.arxiu " + vcriteri + " ORDER BY Count(trim(arxiu)) asc;")
  'Clipboard.Clear
  'Clipboard.SetText "SELECT count(Clixes.arxiu) as Q,first(clixes.arxiu) as Parxiu, cdbl(mid(trim(first(Clixes.arxiu)),4)) as NumXL FROM Clixes WHERE " + vXL + " GROUP BY Clixes.arxiu " + vcriteri + " ORDER BY Count(Clixes.arxiu) asc;"
  
  If Not rst.EOF Then suggerirXL = rst!NumXL
fi:
  Set rst = Nothing
End Function

Private Sub liniaproducte_Change()
  If clixes.Recordset.EditMode > 0 Then
     If Screen.ActiveControl.Name = "liniaproducte" Then emplenarllistalinies
  End If
End Sub

Private Sub liniaproducte_DropDown()
  Dim idmarca As String
  
   
   If marcaproducte = "" Then MsgBox "Primer has d'escollir una marca.", vbCritical, "Atenció": Exit Sub
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "select distinct linia from clixes where marca='" + treure_apostruf(marcaproducte) + "' order by linia"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar

   'formseleccio.DBGrid2.Columns(0).Width = 0
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           liniaproducte = formseleccio.DBGrid2.Columns("linia")
        End If
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Sub emplenarllistamarques(client As Double)
   Dim rstc As Recordset
   opcionscombo.Clear
   If atrim(liniaproducte) = "" Then opcionscombo.visible = False: Exit Sub
   Set rstc = dbclixes.OpenRecordset("select distinct marca from clixes where marca like '" + treure_apostruf(marcaproducte) + "*' " + IIf(cadbl(codiclient) > 0, "and codiclienttemporal=" + atrim(cadbl(codiclient)), ""))
   While Not rstc.EOF
     opcionscombo.AddItem rstc!marca
     rstc.MoveNext
   Wend
   opcionscombo.Left = frameclixes.Left + marcaproducte.Left
   opcionscombo.Top = marcaproducte.Top + marcaproducte.Height + frameclixes.Top
   opcionscombo.width = marcaproducte.width
   opcionscombo.Height = 2000
   If opcionscombo.ListCount < 11 Then opcionscombo.Height = opcionscombo.ListCount * 200
   
   opcionscombo.visible = True
End Sub


Sub emplenarllistalinies()
   Dim rstc As Recordset
   opcionscombo.Clear
   If atrim(liniaproducte) = "" Then opcionscombo.visible = False: Exit Sub
   Set rstc = dbclixes.OpenRecordset("select distinct linia from clixes where marca='" + atrim(cadbl(clixes.Recordset!marca)) + "' AND linia like '" + treure_apostruf(liniaproducte) + "*' order by linia")
   While Not rstc.EOF
     opcionscombo.AddItem rstc!linia
     rstc.MoveNext
   Wend
   opcionscombo.Left = frameclixes.Left + liniaproducte.Left
   opcionscombo.Top = liniaproducte.Top + liniaproducte.Height + frameclixes.Top
   opcionscombo.width = liniaproducte.width
      opcionscombo.Height = 2000
   If opcionscombo.ListCount < 11 Then opcionscombo.Height = opcionscombo.ListCount * 200

   
   opcionscombo.visible = True
End Sub

Private Sub liniaproducte_LostFocus()
   opcionscombo.visible = False
End Sub

Private Sub llistadecomandespendents_Click()
  If llistadecomandespendents.ListIndex = -1 Then llistadecomandespendents.visible = False: Exit Sub
  cridarcomandes cadbl(Mid(llistadecomandespendents.List(llistadecomandespendents.ListIndex), 1, 7))
  llistadecomandespendents.visible = False
End Sub

Private Sub mactualitzaciodadesimpresors_Click()
  actualitzardadesimpresors
End Sub
Sub actualitzardadesimpresors(Optional treballbaixa)
 Dim treball As Double
   Dim ordre As Double
   Dim arxiu As String
   Dim nummontadora As String
   Dim rstvinculats As Recordset
   If baixaclixes Then
      If cadbl(clixes.Recordset!id_treball) <> treballbaixa Then MsgBox "No s'ha trobat aquest treball", vbCritical, "Error": End
      'FormClixes.visible = False
   End If
   If cadbl(treballbaixa) = 0 Then
      treball = cadbl(InputBox("Entra el treball que vols editar.", "Atenció"))
        Else: treball = treballbaixa
   End If
   If treball < 1 Then Exit Sub
   'ordre = cadbl(InputBox("Entra la modificacio que vols editar.", "Atenció"))
   ordre = escullirordredesdecomandespendents(treball)
   If ordre < 1 Then Exit Sub
   clixes.Recordset.FindFirst "id_treball=" + atrim(treball)
   If clixes.Recordset!id_treball <> treball Then Exit Sub
   modificacions.Recordset.FindFirst "ordre=" + atrim(ordre)
   If modificacions.Recordset!ordre <> ordre Then Exit Sub
   modificacions.Recordset.FindFirst "ordre=" + atrim(ordre)
   arxiu = InputBox("Entra el numero d'arxiu correcte.", "Arxiu", atrim(clixes.Recordset!arxiu))
   If atrim(arxiu) <> "" Then
        clixes.Recordset.Edit
        clixes.Recordset!arxiu = arxiu
        clixes.Recordset.Update
    End If
   'canvio tots els arxius de muntadora
   Set rstvinculats = dbclixes.OpenRecordset("select * from clientsvinculats where id_treball=" + atrim(treball) + " and ordremodificacio=" + atrim(ordre))
   If Not rstvinculats.EOF Then nummontadora = rstvinculats!codimuntadora
   nummontadora = atrim(InputBox("Entra el numero de montadora", "Nº Montadora", nummontadora))
   If nummontadora <> "" Then
     While Not rstvinculats.EOF
      rstvinculats.Edit
      rstvinculats!codimuntadora = atrim(nummontadora)
      rstvinculats.Update
      rstvinculats.MoveNext
     Wend
   End If
   '------------------------
   formtintes.Show 1
   'formclientsvinculats.Show 1
   
   comptartinters
   carregar_modificacio
   comprovardiferenciescomandesafectades
   comandabuscada = 0
   If baixaclixes Then End
End Sub
Private Sub marcaproducte_Change()
 If clixes.Recordset.EditMode > 0 Then
     If Screen.ActiveControl.Name = "marcaproducte" Then emplenarllistamarques clixes.Recordset!codiclienttemporal
  End If
End Sub

Private Sub marcaproducte_DropDown()
  Dim were As String
   ' were = IIf(cadbl(clixes.Recordset!codiclienttemporal) > 0, " codiclienttemporal=" + atrim(clixes.Recordset!codiclienttemporal), "")
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "select distinct marca from clixes " + were + " order by marca"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   'formseleccio.DBGrid2.Columns(0).Width = 0
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           marcaproducte = formseleccio.DBGrid2.Columns("marca")
        End If
   End If
    If seleccioret = 9 Then
        marcaproducte = ""
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub marcaproducte_LostFocus()
opcionscombo.visible = False
End Sub
Sub escullirfotogravador(vcodi As Double, vnom As String)
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "select codi,nomfotogravador from fotogravadors"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.sortirs.tag = "filtre"
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           vnom = formseleccio.DBGrid2.Columns("nomfotogravador")
           vcodi = formseleccio.DBGrid2.Columns("CODI")
        End If
   End If
    If seleccioret = 9 Then
        vnom = ""
        vcodi = 0
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub mavisoclixes_Click()
  MsgBox "LLOC AVÍS VALDRIA UNA C PER AL CREAR CLIXÉ I MODIFICACIO I UNA R PER REPOSICIÓ", vbInformation, "ATENCIÓ"
  Load formaltarep
  formaltarep.caption = "Manteniment avisos per client"
'  formaltarep.autonum = "transportistes"
  formaltarep.Data1.DatabaseName = clixes.DatabaseName
  formaltarep.Data1.RecordSource = "select * from avisos_clixes"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).visible = False
  formaltarep.DBGrid1.Columns(3).width = 1000
  formaltarep.DBGrid1.Columns(1).width = 1300
  formaltarep.DBGrid1.Columns(2).width = 5000
  formaltarep.width = 9000
  formaltarep.status = "LLOC AVIS (C,R)"
  'formaltarep.DBGrid1.width = formaltarep.DBGrid1.width + 700
  formaltarep.Show 1
End Sub

Private Sub mclixesfotogravador_Click()
     Dim datainici As Date
     Dim datafi As Date
     Dim resp As String
     Dim oapp As CRAXDDRT.Application
     Dim oreport As CRAXDDRT.Report
     Dim subbusqueda As String
     Dim vcodifotogravador As Double
     Dim vnomfotogravador As String
     escullirfotogravador vcodifotogravador, vnomfotogravador
     If vcodifotogravador = 0 Then Exit Sub
     resp = InputBox("Desde el dia: " + Chr(10) + "Data inici... dd/mm/yy", "Llistat de clixes no utilitzats desde...")
     If Not IsDate(resp) Then MsgBox "Data no vàlida.", vbCritical, "Error": Exit Sub
     datainici = CVDate(resp)
     resp = InputBox("Fins el dia: " + Chr(10) + "Data fi... dd/mm/yy", "Llistat de clixes no utilitzats desde...")
     If Not IsDate(resp) Then MsgBox "Data no vàlida.", vbCritical, "Error": Exit Sub
     datafi = CVDate(resp)
     Set oapp = New CRAXDDRT.Application
     Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "clixesllistatperfotogravador.rpt", 1)
     oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "clixesnous.mdb"
     oreport.RecordSelectionFormula = "{treballsdunfotograbador.dataobertura}>=#" + Format(datainici, "mm/dd/yy") + "# and {treballsdunfotograbador.dataobertura}<=#" + Format(datafi, "mm/dd/yy") + "# and {treballsdunfotograbador.fotograbador}=" + atrim(vcodifotogravador)
     oreport.DiscardSavedData
     oreport.FormulaFields.GetItemByName("titols").Text = "'Treballs del fotogravador : " + vnomfotogravador + "'"
     oreport.FormulaFields.GetItemByName("limits").Text = "'Data inici: " + Format(datainici, "dd/mm/yyyy") + " i Data fi: " + Format(datafi, "dd/mm/yyyy") + "'"
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
   
End Sub

Private Sub mclixesimpresio_Click()
     Dim datainici As Date
     Dim datafi As Date
     Dim resp As String
     Dim oapp As CRAXDDRT.Application
     Dim oreport As CRAXDDRT.Report
     Dim subbusqueda As String
     
     resp = InputBox("Clixes que no s'utilitzen desde el dia: " + Chr(10) + "Data inici... dd/mm/yy", "Llistat de clixes no utilitzats desde...")
     If Not IsDate(resp) Then MsgBox "Data no vàlida.", vbCritical, "Error": Exit Sub
     datainici = CVDate(resp)
     resp = InputBox("Clixes que no s'utilitzen fins el dia: " + Chr(10) + "Data fi... dd/mm/yy", "Llistat de clixes no utilitzats desde...")
     If Not IsDate(resp) Then MsgBox "Data no vàlida.", vbCritical, "Error": Exit Sub
     datafi = CVDate(resp)
     Set oapp = New CRAXDDRT.Application
     Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatclixesiultimaimpresio.rpt", 1)
     oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "clixesnous.mdb"
     oreport.RecordSelectionFormula = "{dadesclixesiultimacomanda.MáxDedatacomanda}>=#" + Format(datainici, "mm/dd/yy") + "# and {dadesclixesiultimacomanda.MáxDedatacomanda}<=#" + Format(datafi, "mm/dd/yy") + "#"
     oreport.DiscardSavedData
     oreport.FormulaFields.GetItemByName("rangdates").Text = "'Data inici: " + Format(datainici, "dd/mm/yyyy") + " i Data fi: " + Format(datafi, "dd/mm/yyyy") + "'"
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
   resp = InputBox("Vols donar de baixa aquests treballs?" + Chr(10) + "Entra la data que vols utilitzar per donar-los de baixa.", "Baixa de clixes")
   If IsDate(resp) Then
       If MsgBox("Segur que vols donar de baixa tots aquests clixes amb data " + atrim(resp), vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
           subbusqueda = "SELECT comandes.numtreball from comandes GROUP BY comandes.numtreball HAVING (((Max(comandes.datacomanda))>=#" + Format(datainici, "mm/dd/yy") + "# And (Max(comandes.datacomanda))<=#" + Format(datafi, "mm/dd/yy") + "#));"
           dbclixes.Execute "update clixes set databaixaclixe=#" + Format(CVDate(resp), "mm/dd/yy") + "# where id_treball in (" + subbusqueda + ")"
       End If
   End If
End Sub

Private Sub mclixesperpalet_Click()
    Dim oapp As CRAXDDRT.Application
    Dim oreport As CRAXDDRT.Report
    Dim vpalet As Double
    
    vpalet = cadbl(InputBox("Entra el numero de palet que vols llistar", "Llistar contingut de clixes d'un palet"))
    If vpalet = 0 Then MsgBox "Has d'escriure un numero de palet vàlid.", vbCritical, "Error": Exit Sub
    Set oapp = New CRAXDDRT.Application
    Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatclixesperpalet.rpt", 1)
    oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "clixesnous.mdb"
    
    oreport.RecordSelectionFormula = "{Clixes.ubicacio}='P-" + Format(vpalet, "000") + "'"
    oreport.DiscardSavedData
    oreport.FormulaFields.GetItemByName("titol").Text = "'Llistat de clixes dins el palet:  P-" + Format(vpalet, "000") + "'"
    
    Load veurereport
    veurereport.CRViewer.ReportSource = oreport
    veurereport.CRViewer.DisplayGroupTree = False
    veurereport.CRViewer.ViewReport
    veurereport.WindowState = 2
    veurereport.Show 1
    
End Sub

Private Sub mcomandesafotogravadorspendents_Click()
Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "clixesllistatcomandespendents.rpt", 1)
  
  oreport.Database.Tables.Item(1).Location = camiclixes
  oreport.DiscardSavedData
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.ViewReport
   veurereport.width = formclixes.width - 200
   veurereport.Height = formclixes.Height - 300
   
   veurereport.Show 1, Me
  ratoli "normal"
End Sub

Private Sub mestatsclixes_Click()
 Load formaltarep
  formaltarep.caption = "Estat de clixés"
'  formaltarep.autonum = "transportistes"
  formaltarep.Data1.DatabaseName = clixes.DatabaseName
  formaltarep.Data1.RecordSource = "select * from clixes_estats"

  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  
  formaltarep.width = formaltarep.width + 700
  formaltarep.DBGrid1.width = formaltarep.DBGrid1.width + 700
  formaltarep.Show
End Sub

Private Sub mestattreballs_Click()
  Formconsultaestats.Show
End Sub

Private Sub mfotogravadors_Click()
   fFotogravadors.Show 1
End Sub

Private Sub mimprimirbossasoldadores_Click()
   Dim vnumc As String
   vnumc = InputBox("Entra el numero de comanda que vols imprimir la bossa de soldadores", "Bossa soldadores")
   If cadbl(vnumc) > 0 Then imprimirbossessoldadores cadbl(vnumc), True
End Sub

Private Sub mlinies_Click()
  vincularliniesimpresio.Show
End Sub

Private Sub mliniesperpantone_Click()
  Load Formagrupartreballs
  Formagrupartreballs.Command3.visible = False
  Formagrupartreballs.Show 1
End Sub

Private Sub modificacions_Reposition()
   carregar_modificacio
End Sub
Sub possarnommaterialultimacomanda()
   Dim rst As Recordset
   Dim rstmat As Recordset
   Dim subseleccio As String
   materialultimacomanda = ""
   Set rst = dbcomandes.OpenRecordset("select * from comandes where numtreball=" + atrim(id_treball) + " and numordremodificacio=" + atrim(ordremodificacio) + " order by comanda DESC")
   If rst.EOF Then Exit Sub
   Set rst = dbcomandes.OpenRecordset("select materialex from comandes where comanda=" + atrim(rst!comanda) + "or comanda=" + atrim(cadbl(rst!linkcomanda1)) + " or comanda=" + atrim(cadbl(rst!linkcomanda2)))
   While Not rst.EOF
        Set rstmat = dbcomandes.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rst!materialex)))
        materialultimacomanda = materialultimacomanda + "  " + descripciomaterial_curta(rstmat)
        rst.MoveNext
   Wend
   Set rst = Nothing
   Set rstmat = Nothing
End Sub
Function descripciomaterial_curta(rstmat As Recordset) As String
  Dim desc As String
  Dim rstfam As Recordset
  If rstmat.EOF Then Exit Function
  Set rstfam = dbcomandes.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(cadbl(rstmat!familia)))
  If Not rstfam.EOF Then desc = desc + atrim(rstfam!descripcio)
 ' Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliesmaterials where codi=" + atrim(cadbl(rstmat!subfamilia)))
 ' If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbcomandes.OpenRecordset("select descripcio from familiescolorants where codi=" + atrim(cadbl(rstmat!familiacol)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  'Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliescolorants where codi=" + atrim(cadbl(rstmat!subfamiliacol)))
  'If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  'Set rstfam = dbtmpb.OpenRecordset("select descripcio from familiesaditius where codi=" + atrim(cadbl(rstmat!familiaad)))
  'If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  'Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliesaditius where codi=" + atrim(cadbl(rstmat!subfamiliaad)))
  'If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  descripciomaterial_curta = desc
End Function

Function af(v As Variant) As String
  v = atrim(v)
  If Len(v) > 1 Then
     v = " + " + v
    Else: v = ""
  End If
  af = v
End Function
Function albaransmesquepressupost() As Boolean
   Dim rstalb As Recordset
   Dim rstpress As Recordset
   albaransmesquepressupost = False
   Set rstpress = dbclixes.OpenRecordset("select  preu  from pressupostos where id_Treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
   Set rstalb = dbclixes.OpenRecordset("select sum(import) as total from clixes_albarans where id_Treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
   If Not rstalb.EOF And Not rstpress.EOF Then
       If cadbl(rstalb!total) > cadbl(rstpress!preu) Then albaransmesquepressupost = True
   End If
   Set rstalb = Nothing
   Set rstpress = Nothing
End Function
Sub comprovarsialbaransmesquepressupost()
   If albaransmesquepressupost Then
        botoliniesalbarans.BackColor = QBColor(12)
   End If
End Sub
Sub carregar_modificacio()
   ordremodificacio = 0
   If modificacions.Recordset.EOF Or modificacions.Recordset.BOF Then Exit Sub
   possarcolorbotopdf
   ordremodificacio = cadbl(modificacions.Recordset!ordre)
   carregarnomclient
   posarfotograbador
   possarcoloralsbotonsquetenendades
   posarcolor_bototintes
   comprovarsialbaransmesquepressupost
   possarestatclixe
   possarnommaterialultimacomanda
   etcodidelinia = ""
   If cadbl(modificacions.Recordset!codidelinia) > 0 Then etcodidelinia = "Codi de linia: " + Format(cadbl(modificacions.Recordset!codidelinia), "000") + "#" + Format(cadbl(modificacions.Recordset!codideliniav), "0")
   If IsDate(modificacions.Recordset!datapdf) Then
       etdatapdf = Format(modificacions.Recordset!datapdf, "dd/mm/yy")
      Else: etdatapdf = ""
   End If
   If modificacions.Recordset!reimpres Then
       etreprint.visible = True
         Else: etreprint.visible = False
   End If
   If modificacions.Recordset!ordre < ordremesgran Then
     framemodificacions.BackColor = &HC0C0FF
       Else
         framemodificacions.BackColor = &HEAD9CE
   End If
   revisarelsvalorspressupostcorrectes
   If cobservaciorepasclixes = "" Then
      cobservaciorepasclixes.visible = False
       Else: cobservaciorepasclixes.visible = True
   End If
   
End Sub
Sub revisarelsvalorspressupostcorrectes()
   Dim rstpress As Recordset
   Set rstpress = dbclixes.OpenRecordset("select  *  from pressupostos where id_Treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
   lerrorpressupost.visible = False
   If rstpress.EOF Then Exit Sub
   If Not valorspressupostcorrectes(cadbl(rstpress!amplelamina), cadbl(rstpress!bandes), cadbl(rstpress!cilindre), cadbl(rstpress!desarroll), cadbl(rstpress!tinters)) Then lerrorpressupost.visible = True
End Sub

Sub possarcolorbotopdf()
    If modificacions.Recordset!pdfvalid Then
        botopdf.Picture = botopdf.DownPicture
        missatgenoupdf.visible = False
          Else:
              botopdf.Picture = botopdf.DisabledPicture
              missatgenoupdf.visible = True
    End If
End Sub

Private Sub modificar_Click()
  If clixes.Recordset.EditMode > 0 Then MsgBox "Estas editant primer finalitza la operació i despres afegeix.", vbCritical, "Atenció": Exit Sub
   clixes.Recordset.Edit
   If Not (modificacions.Recordset.EOF And modificacions.Recordset.BOF) Then modificacions.Recordset.Edit
   
   framesactivats True
   
   codidebarres.SetFocus
End Sub

Private Sub nompdf_Change()
  If atrim(nompdf) <> "" Then
     botopdf.Picture = botopdf.DownPicture
    Else: botopdf.Picture = botopdf.DisabledPicture
  End If
End Sub

Private Sub º_Click()

End Sub

Private Sub repasclixes_Click()

End Sub

Sub crear_taula_tmp_llistatpntfact()
  Dim taula_tmp As String
  Dim camps(100, 2) As String
   'creo la taula de linies d'albarans
  taula_tmp = "tmp_albarans_pendents"
  On Error Resume Next
   dbclixes.Execute "drop table " + taula_tmp
  On Error GoTo 0
  i = 1
  camps(i, 1) = "id_treball": camps(i, 2) = "integer": i = i + 1
  camps(i, 1) = "data": camps(i, 2) = "date": i = i + 1
  camps(i, 1) = "client": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "producte": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "total": camps(i, 2) = "double": i = i + 1
  
  dbclixes.Execute ("create table " + taula_tmp + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbclixes.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
  
End Sub


Private Sub mpendentsdefacturar_Click()
 Dim rstlli As Recordset
   Dim rstmod As Recordset
   Dim rsttreball As Recordset
   Dim rstclient
   Dim nomclient As String
   crear_taula_tmp_llistatpntfact
   Set rstlli = dbclixes.OpenRecordset("tmp_albarans_pendents")
   r = "SELECT Clixes_albarans.id_treball, Clixes_albarans.ordremodificacio,Sum(Clixes_albarans.import) AS total From Clixes_albarans"
   r = r + " GROUP BY Clixes_albarans.id_treball,Clixes_albarans.ordremodificacio, Clixes_albarans.facturat"
   r = r + " HAVING (((Clixes_albarans.facturat)=False));"
   
   

   
   
   Set rstmod = dbclixes.OpenRecordset(r)
   While Not rstmod.EOF
     r = "SELECT Clixes.id_treball, Clixes_albarans.ordremodificacio, Clixes.linia, Clixes_albarans.data, Clixes_albarans.quantitat, clients.nom, Clixes.nomclienttemporal "
     r = r + " FROM clients RIGHT JOIN ((Clixes INNER JOIN Clixes_albarans ON Clixes.id_treball = Clixes_albarans.id_treball) LEFT JOIN Clientsvinculats ON Clixes.id_treball = Clientsvinculats.id_treball) ON clients.codi = Clientsvinculats.codiclient"
     r = r + " WHERE (((Clixes.id_treball)=" + atrim(rstmod!id_treball) + ") AND ((Clixes_albarans.ordremodificacio)=" + atrim(rstmod!ordremodificacio) + "));"
     Set rsttreball = dbclixes.OpenRecordset(r)
     If Not rsttreball.EOF Then
         'Set rstclient = dbtmpb.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(rsttreball!id_client)))
         nomclient = atrim(rsttreball!nom)
         If nomclient = "" Then nomclient = atrim(rsttreball!nomclienttemporal)
         rstlli.AddNew
           rstlli!id_treball = rsttreball!id_treball
           rstlli!data = rsttreball!data
           rstlli!client = nomclient
           rstlli!producte = rsttreball!linia
           rstlli!total = cadbl(rstmod!total)
         rstlli.Update
     End If
     rstmod.MoveNext
     nomclient = ""
   Wend
   Set rstlli = Nothing
   wait (2)  'faig una espera perque a vegades falten registres
   'llenço el llistat
   llistat.Formulas(0) = ""
   llistat.Formulas(1) = ""
   llistat.Formulas(2) = ""
  llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "clixespendents.rpt"
  llistat.DiscardSavedData = True
 llistat.DataFiles(0) = camiclixes
 llistat.Destination = crptToWindow
 llistat.Action = 1
   
End Sub

Private Sub mquanbossesarxiu_Click()
     Dim resp As String
     Dim oapp As CRAXDDRT.Application
     Dim oreport As CRAXDDRT.Report
     
     Set oapp = New CRAXDDRT.Application
     Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistatquantitatsbossesperarxiu.rpt", 1)
     oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "clixesnous.mdb"
     'oreport.RecordSelectionFormula = "mid({Clixes.ubicacio},1,5)<>'Palet' and {Clixes.arxiu}<>'' and isnull({Clixes.databaixaclixe}) and {clixes.estatclixe}<>'RETORNEM CLIXES'"
     oreport.RecordSelectionFormula = "{@arxiusenseXL}>0 and (trim({Clixes.arxiu})<>'' and isnull({Clixes.databaixaclixe}) and {clixes.estatclixe}<>'RETORNEM CLIXES')"
     oreport.DiscardSavedData
     'oreport.FormulaFields.GetItemByName("rangdates").Text = "'Data inici: " + Format(datainici, "dd/mm/yyyy") + " i Data fi: " + Format(datafi, "dd/mm/yyyy") + "'"
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1


End Sub

Private Sub mtotallliureslleixes_Click()
  Dim resp As String
     Dim oapp As CRAXDDRT.Application
     Dim oreport As CRAXDDRT.Report
     
     Set oapp = New CRAXDDRT.Application
     Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistatotalsbossesperarxiu.rpt", 1)
     oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "clixesnous.mdb"
     'oreport.RecordSelectionFormula = "mid({Clixes.ubicacio},1,5)<>'Palet' and {Clixes.arxiu}<>'' and isnull({Clixes.databaixaclixe}) and {clixes.estatclixe}<>'RETORNEM CLIXES'"
     oreport.RecordSelectionFormula = "{@arxiusenseXL}>0 and (trim({Clixes.arxiu})<>'' and isnull({Clixes.databaixaclixe}) and {clixes.estatclixe}<>'RETORNEM CLIXES')"
     oreport.DiscardSavedData
     'oreport.FormulaFields.GetItemByName("rangdates").Text = "'Data inici: " + Format(datainici, "dd/mm/yyyy") + " i Data fi: " + Format(datafi, "dd/mm/yyyy") + "'"
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1

End Sub

Private Sub nomclienttemporal_DropDown()
  'If nomclientclixe = "" Then
   Load formseleccio
   formseleccio.sortirs.tag = "filtre"
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select * from clients"
   formseleccio.refrescar
   formseleccio.Show 1
   
    If seleccioret = 1 Then
            clixes.Recordset!codiclienttemporal = formseleccio.DBGrid2.Columns("codi")
            nomclienttemporal = formseleccio.DBGrid2.Columns("nom")
            comprovarsihihaavisosperaquestclient clixes.Recordset!codiclienttemporal, "C"
    End If
     If seleccioret = 9 Then
         clixes.Recordset!codiclienttemporal = "0"
         nomclienttemporal = ""
    End If
    formseleccio.Data1.RecordSource = ""
    formseleccio.Data1.Refresh
    Unload formseleccio
    SendKeys "{TAB}"
   '   Else: MsgBox "Ja hi ha un client vinculat no pots possar un temporal", vbCritical, "Atenció": Exit Sub
   'End If
 
End Sub

Private Sub nomproveidor_DropDown()
   If modificacions.Recordset.EditMode = 0 Then MsgBox "No pots modificar, primer prem el botó d'editar.", vbCritical, "Atenció": Exit Sub
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "select codi,nomfotogravador from fotogravadors where actiu"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.sortirs.tag = "filtre"
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           nomproveidor = formseleccio.DBGrid2.Columns("nomfotogravador")
           modificacions.Recordset!fotograbador = formseleccio.DBGrid2.Columns("CODI")
        End If
   End If
    If seleccioret = 9 Then
        nomproveidor = ""
        modificacions.Recordset!fotograbador = Null
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   If modificacions.Recordset.EditMode > 0 Then
      modificacions.Recordset.Update
      modificacions.Recordset.Edit
   End If
End Sub

Private Sub opcionscombo_DblClick()
   If opcionscombo.ListIndex = -1 Then Exit Sub
   liniaproducte = opcionscombo.Text
   clixes.Recordset!id_linia = opcionscombo.ItemData(opcionscombo.ListIndex)
   opcionscombo.visible = False
End Sub

Private Sub reducciocilindref2_LostFocus()
If cadbl(reducciopermetre) >= 0 And cadbl(reducciocilindref2) < 0 Then MsgBox "El valor de factor per la F2 no es correcte", vbCritical, "Error F2": reducciocilindref2 = "0"
     If cadbl(reducciopermetre) < 0 And cadbl(reducciocilindref2) > 0 Then MsgBox "El valor de factor per la F2 no es correcte", vbCritical, "Error F2": reducciocilindref2 = "0"
End Sub

Private Sub reducciocilindrefw_LostFocus()
   If cadbl(reducciopermetre) >= 0 And cadbl(reducciocilindrefw) < 0 Then MsgBox "El valor de factor per la FW no es correcte", vbCritical, "Error FW": reducciocilindrefw = "0"
   If cadbl(reducciopermetre) < 0 And cadbl(reducciocilindrefw) > 0 Then MsgBox "El valor de factor per la FW no es correcte", vbCritical, "Error FW": reducciocilindrefw = "0"
End Sub

Private Sub sistemaimpresio_Click()
   If sistemaimpresio = "Offset" Then MsgBox "Aquest sistema no es pot escollir.", vbCritical, "Atenció": sistemaimpresio = "": Exit Sub
End Sub

Private Sub sistemaimpresio_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub sistemaimpresio_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub sortir_Click()
  End
End Sub
Sub crear_taules_tmp()
  Dim camps(100, 2) As String
  taula_tmp = "tmp_clixes_capcalera"
  On Error Resume Next
   dbclixes.Execute "drop table " + taula_tmp
  On Error GoTo 0
  i = 1
  camps(i, 1) = "id_treball": camps(i, 2) = "integer": i = i + 1
  camps(i, 1) = "numordremodificacio": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "arxiuclixe": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "datainici": camps(i, 2) = "date": i = i + 1
  camps(i, 1) = "formaimp": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "estatclixe": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "client": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "marca": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "linia": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "representant": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "proveidor": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "montadora": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "codibarres": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "dataentrega": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "observacions": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "sistemaimpresio": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "bandesclixes": camps(i, 2) = "integer": i = i + 1
  camps(i, 1) = "ample": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "desarroll": camps(i, 2) = "double": i = i + 1
  
  dbclixes.Execute ("create table " + taula_tmp + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbclixes.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
  dbclixes.Execute "CREATE INDEX ordre ON " + taula_tmp + " ([numordremodificacio]);"
  'creo la taula de linies d'albarans
  taula_tmp = "tmp_clixes_albarans_linies"
  On Error Resume Next
   dbclixes.Execute "drop table " + taula_tmp
  On Error GoTo 0
  i = 1
  camps(i, 1) = "id_treball": camps(i, 2) = "integer": i = i + 1
  camps(i, 1) = "data": camps(i, 2) = "date": i = i + 1
  camps(i, 1) = "numalb": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "quantitat": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "descripcio": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "import": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "facturat": camps(i, 2) = "string": i = i + 1
  
  dbclixes.Execute ("create table " + taula_tmp + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbclixes.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
  
  'creo la taula de linies de tintes
  taula_tmp = "tmp_clixes_tintes_linies"
  On Error Resume Next
   dbclixes.Execute "drop table " + taula_tmp
  On Error GoTo 0
  i = 1
  camps(i, 1) = "id_treball": camps(i, 2) = "integer": i = i + 1
  camps(i, 1) = "numordremodificacio": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "color": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "anilox": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "cilindre": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "desarroll": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "continuu": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "facturat": camps(i, 2) = "string": i = i + 1
  
  dbclixes.Execute ("create table " + taula_tmp + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbclixes.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
  
  'creo la taula de linies
  taula_tmp = "tmp_clixes_modifis_linies"
  On Error Resume Next
   dbclixes.Execute "drop table " + taula_tmp
  On Error GoTo 0
  i = 1
  camps(i, 1) = "descripcio": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "inici": camps(i, 2) = "date": i = i + 1
  camps(i, 1) = "fi": camps(i, 2) = "date": i = i + 1
  camps(i, 1) = "id_treball": camps(i, 2) = "integer": i = i + 1
  
  dbclixes.Execute ("create table " + taula_tmp + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbclixes.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
  
End Sub





Private Sub Text16_Change()

End Sub

Private Sub Text2_GotFocus()
  bxl.visible = True
  bxl.Left = Text2.Left + Text2.width - bxl.width
  bxl.Top = Text2.Top
  bxl.tag = "text2"
End Sub

Private Sub Text2_LostFocus()
    If Screen.ActiveControl.Name <> "bxl" Then bxl.visible = False
End Sub

Private Sub Text4_GotFocus()
  bxl.visible = True
  bxl.Left = Text4.Left + Text4.width - bxl.width
  bxl.Top = Text4.Top
  bxl.tag = "text4"
End Sub

Private Sub Text4_LostFocus()
  If Screen.ActiveControl.Name <> "bxl" Then bxl.visible = False
End Sub

Private Sub Timer1_Timer()
  ferpampallugadebotons
End Sub

Private Sub timerdrag_Timer()
  Dim rutaarxiu As String
  
  If modificacioautomatica Then controlartempsobert
  rutaarxiu = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Arxiu_documentacio_relacionada"
  If arxiudocumentaciorelacionada.tag = "1" Then
     arxiudocumentaciorelacionada.tag = ""
     timerdrag.Enabled = False
     
     obrircarpetaarxiu rutaarxiu, id_treball
     arxiudocumentaciorelacionada.BackColor = QBColor(15)
  End If
  esperar10segonsaveuresientraelfitxer
     

End Sub
Sub ferpampallugadebotons()
   If botoavisosliniesalbarans.visible Then
      If botoavisosliniesalbarans.BackColor = QBColor(12) Then
           botoavisosliniesalbarans.BackColor = &H8000000F
         Else: botoavisosliniesalbarans.BackColor = QBColor(12)
      End If
   End If
End Sub
Sub controlartempsobert()
   Static dataentrada As Date
   If DateDiff("s", dataentrada, Now) > 100 Then dataentrada = Now
   If DateDiff("s", dataentrada, Now) > 8 Then End
End Sub
Function ProcIDFromWnd(ByVal hWnd As Long) As Long
   Dim idProc As Long
   
   ' Get PID for this HWnd
   GetWindowThreadProcessId hWnd, idProc
   
   ' Return PID
   ProcIDFromWnd = idProc
End Function
      
Function GetWinHandle(hInstance As Long) As Long
   Dim tempHwnd As Long
   
   ' Grab the first window handle that Windows finds:
   tempHwnd = FindWindow(vbNullString, vbNullString)
   
   ' Loop until you find a match or there are no more window handles:
   Do Until tempHwnd = 0
      ' Check if no parent for this window
      If GetParent(tempHwnd) = 0 Then
         ' Check for PID match
         If hInstance = ProcIDFromWnd(tempHwnd) Then
            ' Return found handle
            GetWinHandle = tempHwnd
            ' Exit search loop
            Exit Do
         End If

      End If
   
      ' Get the next window handle
      tempHwnd = GetWindow(tempHwnd, 2)
   Loop
End Function


Sub guardar_i_enviar_previs()
   Dim vnomfitxerdocx As String
   Dim vnomfitxerpdf1 As String
   Dim vnomfitxerpdf2 As String
   If MsgBox("VOLS ENVIAR EL FORMULARI OMPLERT I ELS PDF PRÈVIS A QUI CORRESPONGUI?", vbExclamation + vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbNo Then Exit Sub
   If Not existeix(ruta_documentacio_clixes) Then MsgBox "No hi ha acces a la carpeta " + ruta_documentacio_clixes, vbCritical, "Error": Exit Sub
   vnomfitxerdocx = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Arxiu_documentacio_relacionada\v" + atrim(ordremodificacio) + "\Revisió_Prèvi_" + atrim(id_treball) + "-" + atrim(ordremodificacio) + ".docx"
   If Not existeix(ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Arxiu_documentacio_relacionada") Then MkDir ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Arxiu_documentacio_relacionada"
   If Not existeix(ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Arxiu_documentacio_relacionada\v" + atrim(ordremodificacio)) Then MkDir ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Arxiu_documentacio_relacionada\v" + atrim(ordremodificacio)
   
   
   If existeix(vnomfitxerdocx) Then Kill vnomfitxerdocx
  ' MsgBox vnomfitxerdocx
   FileCopy "c:\temp\TEMP_PLANTILLA REVISIÓ PREVI DE CLIXES.docx", vnomfitxerdocx
   vnomfitxerpdf1 = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordremodificacio, "000") + "_PRSC" + ".pdf"
   vnomfitxerpdf2 = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordremodificacio, "000") + "_PR" + ".pdf"
   'enviaremail "miquel.inplacsa@gmail.com", "Revisió PRÈVI del treball: " + atrim(id_treball) + "/" + atrim(ordremodificacio), "Adjunto pdfs previs i document de revisió del prèvi a l'espera del definitiu.", vnomfitxerdocx, , vnomfitxerpdf1, vnomfitxerpdf2
   enviaremail "RevisioTintesTreballs", "Revisió PRÈVI del treball: " + atrim(id_treball) + "/" + atrim(ordremodificacio), "Adjunto pdfs previs i document de revisió del prèvi a l'espera del definitiu.", vnomfitxerdocx, , vnomfitxerpdf1, vnomfitxerpdf2
End Sub

Function pesMbdelsPDFprèvis() As Double
   Dim vfitxer1 As String
   Dim vfitxer2 As String
   Dim vpes As Double
   vfitxer1 = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordremodificacio, "000") + "_PR" + ".pdf"
   vfitxer2 = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordremodificacio, "000") + "_PRSC" + ".pdf"
   If existeix(vfitxer1) Then vpes = FileLen(vfitxer1)
   If existeix(vfitxer2) Then vpes = vpes + FileLen(vfitxer2)
   vpes = vpes / 1000000
   pesMbdelsPDFprèvis = vpes
End Function
Sub enviar_revisio_previ()
  Dim oDoc As Word.Document
  Dim oWord As Word.Application
  Set oWord = CreateObject("Word.Application")
  If pesMbdelsPDFprèvis > 25 Then MsgBox "Els PDFs prèvis pesen mes de 25Mb i no es poden enviar.", vbCritical, "Error": Exit Sub
  comprovarsihihaavisosperaquestclient clixes.Recordset!codiclienttemporal, "R"
  If MsgBox("S'obrirà un document de word on has de omplir els camps marcats, quan acabis s'enviarà juntament amb els PDF prèvis a qui correspongui." + vbNewLine + "VOLS CONTINUAR?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
  
  If Not assignardocumentword Then MsgBox "No s'ha pogut obrir el document de Plantilla mira que no tinguis cap document de word obert.", vbCritical, "Error": Exit Sub
  oWord.visible = True
  
  
  oWord.Documents.Add llegir_ini("General", "rutallistats", "comandes.ini") + "\PLANTILLA REVISIÓ PREVI DE CLIXES.docx"

  Set oDoc = oWord.Documents(1)
  
  'Clipboard.Clear
  'Clipboard.SetData Formplantillarevisio.Image, vbCFBitmap
  If Not possarvalorsdelsclixesalaplantilla(oDoc) Then
    oDoc.Close False
    GoTo tancat
  End If
  oDoc.SaveAs "c:\temp\TEMP_PLANTILLA REVISIÓ PREVI DE CLIXES.docx"
  oWord.WindowState = wdWindowStateMaximize
  oDoc.Activate
  
  AppActivate "TEMP_PLANTILLA REVISIÓ PREVI DE CLIXES"
  On Error GoTo tancat
  While oDoc.Name <> ""
    wait 1
    DoEvents
  Wend
tancat:
  If existeix("c:\temp\TEMP_PLANTILLA REVISIÓ PREVI DE CLIXES.docx") Then guardar_i_enviar_previs
  assignardocumentword 'elimino el temporal
  Set oWord = Nothing
  Set oDoc = Nothing
End Sub
Function possarvalorsdelsclixesalaplantilla(oDoc As Word.Document) As Boolean
  possarvalorsdelsclixesalaplantilla = True
  On Error GoTo errors
  oDoc.Bookmarks("client").Range.Text = substituir(formclixes.nomclienttemporal, "´", "'")
  oDoc.Bookmarks("marcailinia").Range.Text = substituir(formclixes.marcaproducte, "´", "'") + " - " + substituir(formclixes.liniaproducte, "´", "'")
  oDoc.Bookmarks("numtreball").Range.Text = id_treball
  If formclixes.llistadecomandespendents.ListCount > 0 Then oDoc.Bookmarks("numlot").Range.Text = Mid(formclixes.llistadecomandespendents.List(0), 1, 6)
  Exit Function
errors:
   MsgBox "Hi ha hagut un error al intentar possar els valors del treball a la plantilla." + vbNewLine + "Assegura que estigui habilitat l'us de macros en el Word.", vbCritical, "Error"
   possarvalorsdelsclixesalaplantilla = False
   
End Function

Function assignardocumentword() As Boolean
  On Error GoTo fi
  If existeix("c:\temp\TEMP_PLANTILLA REVISIÓ PREVI DE CLIXES.docx") Then
    Kill "c:\temp\TEMP_PLANTILLA REVISIÓ PREVI DE CLIXES.docx"
  End If
  assignardocumentword = True
  Exit Function
fi:
  assignardocumentword = False
End Function
