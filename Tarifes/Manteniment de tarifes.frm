VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form form_tarifes 
   Caption         =   "Manteniment de tarifes."
   ClientHeight    =   10950
   ClientLeft      =   315
   ClientTop       =   645
   ClientWidth     =   17070
   Icon            =   "Manteniment de tarifes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   17070
   Begin VB.Timer Timercadasegon 
      Interval        =   950
      Left            =   10605
      Top             =   150
   End
   Begin VB.Data datatarifes 
      Caption         =   "datatarifes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   465
      Left            =   6285
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   30
      Width           =   3105
   End
   Begin VB.Timer Timer_escullirclient 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6540
      Top             =   210
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00EAD9CE&
      Height          =   540
      Left            =   105
      TabIndex        =   1
      Top             =   -75
      Width           =   6090
      Begin VB.ComboBox comboclient 
         Height          =   315
         Left            =   660
         TabIndex        =   2
         Top             =   165
         Width           =   5310
      End
      Begin VB.Label Label2 
         BackColor       =   &H00EAD9CE&
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   240
         Left            =   90
         TabIndex        =   3
         Top             =   195
         Width           =   540
      End
   End
   Begin TabDlg.SSTab Pestanyes 
      Height          =   9750
      Left            =   150
      TabIndex        =   0
      Top             =   945
      Width           =   16650
      _ExtentX        =   29369
      _ExtentY        =   17198
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Dades generals de la Tarifa"
      TabPicture(0)   =   "Manteniment de tarifes.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "etversio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fcapcalera1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fcapcalera2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ctarifa"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Escalat de preus"
      TabPicture(1)   =   "Manteniment de tarifes.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "botopegar"
      Tab(1).Control(1)=   "Frameescalat"
      Tab(1).Control(2)=   "cbotoafegirbarem"
      Tab(1).Control(3)=   "Command2"
      Tab(1).Control(4)=   "Framecondicionant"
      Tab(1).Control(5)=   "reixa"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Condicionants"
      TabPicture(2)   =   "Manteniment de tarifes.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "reixacond"
      Tab(2).Control(1)=   "beliminarcond"
      Tab(2).Control(2)=   "Frame4"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Transports"
      TabPicture(3)   =   "Manteniment de tarifes.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "datadestins"
      Tab(3).Control(1)=   "reixadestins"
      Tab(3).Control(2)=   "Command5"
      Tab(3).Control(3)=   "Command4"
      Tab(3).Control(4)=   "reixaports"
      Tab(3).Control(5)=   "Label21"
      Tab(3).Control(6)=   "etcalcultamanyfont"
      Tab(3).ControlCount=   7
      Begin VB.CommandButton botopegar 
         Height          =   315
         Left            =   -74820
         Picture         =   "Manteniment de tarifes.frx":05FA
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Pegar els valors del portapapers a la reixa."
         Top             =   1380
         Width           =   390
      End
      Begin VB.Frame Frameescalat 
         BackColor       =   &H00EEE4D7&
         Enabled         =   0   'False
         Height          =   990
         Left            =   -74820
         TabIndex        =   106
         Top             =   375
         Width           =   16230
         Begin VB.TextBox cdesarrollpcs 
            Alignment       =   2  'Center
            DataField       =   "desarroll_pcs"
            DataSource      =   "datatarifes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4530
            TabIndex        =   119
            Top             =   585
            Width           =   765
         End
         Begin VB.ComboBox combounitatfacturacio 
            DataField       =   "unitat_escalat"
            DataSource      =   "datatarifes"
            Height          =   315
            Left            =   1635
            TabIndex        =   110
            Top             =   315
            Width           =   1230
         End
         Begin VB.ComboBox combounitatescalat 
            DataField       =   "unitat_facturacio"
            DataSource      =   "datatarifes"
            Height          =   315
            Left            =   105
            TabIndex        =   109
            Top             =   330
            Width           =   1230
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            DataField       =   "quantitatminima"
            DataSource      =   "datatarifes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4530
            TabIndex        =   108
            Top             =   195
            Width           =   750
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            DataField       =   "tamanybobina"
            DataSource      =   "datatarifes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6825
            TabIndex        =   107
            Top             =   255
            Width           =   780
         End
         Begin VB.Label etdesarrollpcs 
            BackColor       =   &H00FDDECE&
            BackStyle       =   0  'Transparent
            Caption         =   "Desarroll pcs m/m:"
            Height          =   240
            Left            =   3180
            TabIndex        =   120
            Top             =   645
            Width           =   1560
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Unitat de l'escalat"
            Height          =   210
            Left            =   105
            TabIndex        =   115
            Top             =   105
            Width           =   1710
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Unitat de facturació"
            Height          =   210
            Left            =   1575
            TabIndex        =   114
            Top             =   105
            Width           =   1710
         End
         Begin VB.Label Label29 
            BackColor       =   &H00FDDECE&
            BackStyle       =   0  'Transparent
            Caption         =   "Quantitat mínima:"
            Height          =   240
            Left            =   3195
            TabIndex        =   113
            Top             =   270
            Width           =   1350
         End
         Begin VB.Label Label31 
            BackColor       =   &H00FDDECE&
            BackStyle       =   0  'Transparent
            Caption         =   "Tamany Bobina:                    m"
            Height          =   240
            Left            =   5565
            TabIndex        =   112
            Top             =   315
            Width           =   2355
         End
         Begin VB.Label etpesnet 
            BackStyle       =   0  'Transparent
            Caption         =   "---"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   135
            TabIndex        =   111
            Top             =   675
            Width           =   1935
         End
      End
      Begin VB.Data datadestins 
         Caption         =   "datadestins"
         Connect         =   "Access"
         DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\Tarifes.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   -72345
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   540
         Visible         =   0   'False
         Width           =   2625
      End
      Begin MSDBGrid.DBGrid reixadestins 
         Bindings        =   "Manteniment de tarifes.frx":0B84
         Height          =   7575
         Left            =   -74925
         OleObjectBlob   =   "Manteniment de tarifes.frx":0B9A
         TabIndex        =   84
         Top             =   1440
         Width           =   3990
      End
      Begin VB.TextBox ctarifa 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   390
         Width           =   600
      End
      Begin VB.Frame fcapcalera2 
         BackColor       =   &H00EAD9CE&
         Caption         =   "      Materials"
         Enabled         =   0   'False
         Height          =   1875
         Left            =   8685
         TabIndex        =   48
         Top             =   750
         Width           =   7785
         Begin VB.ComboBox Combomat1 
            DataField       =   "nommat1"
            DataSource      =   "datatarifes"
            Height          =   315
            Left            =   765
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   390
            Width           =   4620
         End
         Begin VB.ComboBox combomat2 
            DataField       =   "nommat2"
            DataSource      =   "datatarifes"
            Height          =   315
            Left            =   780
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   825
            Width           =   4605
         End
         Begin VB.ComboBox combomat3 
            DataField       =   "nommat3"
            DataSource      =   "datatarifes"
            Height          =   315
            Left            =   780
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   1275
            Width           =   4605
         End
         Begin VB.TextBox mat1m_de 
            DataField       =   "mat1_esp_de"
            DataSource      =   "datatarifes"
            Height          =   315
            Left            =   5520
            TabIndex        =   57
            Top             =   390
            Width           =   435
         End
         Begin VB.TextBox mat1m_a 
            DataField       =   "mat1_esp_a"
            DataSource      =   "datatarifes"
            Height          =   315
            Left            =   6225
            TabIndex        =   56
            Top             =   390
            Width           =   435
         End
         Begin VB.TextBox mat2m_de 
            DataField       =   "mat2_esp_de"
            DataSource      =   "datatarifes"
            Height          =   315
            Left            =   5505
            TabIndex        =   55
            Top             =   855
            Width           =   435
         End
         Begin VB.TextBox mat2m_a 
            DataField       =   "mat2_esp_a"
            DataSource      =   "datatarifes"
            Height          =   315
            Left            =   6210
            TabIndex        =   54
            Top             =   855
            Width           =   435
         End
         Begin VB.TextBox mat3m_de 
            DataField       =   "mat3_esp_de"
            DataSource      =   "datatarifes"
            Height          =   315
            Left            =   5505
            TabIndex        =   53
            Top             =   1275
            Width           =   435
         End
         Begin VB.TextBox mat3m_a 
            DataField       =   "mat3_esp_a"
            DataSource      =   "datatarifes"
            Height          =   315
            Left            =   6210
            TabIndex        =   52
            Top             =   1275
            Width           =   435
         End
         Begin VB.TextBox tmat1 
            DataField       =   "mat1"
            DataSource      =   "datatarifes"
            Height          =   285
            Left            =   15
            TabIndex        =   51
            Top             =   210
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.TextBox tmat2 
            DataField       =   "mat2"
            DataSource      =   "datatarifes"
            Height          =   285
            Left            =   15
            TabIndex        =   50
            Top             =   675
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.TextBox tmat3 
            BackColor       =   &H00FFFFFF&
            DataField       =   "mat3"
            DataSource      =   "datatarifes"
            Height          =   285
            Left            =   30
            TabIndex        =   49
            Top             =   1080
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Mat. 1a:"
            Height          =   285
            Left            =   150
            TabIndex        =   67
            Top             =   435
            Width           =   735
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Mat. 2a:"
            Height          =   285
            Left            =   150
            TabIndex        =   66
            Top             =   870
            Width           =   735
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Mat. 3a:"
            Height          =   285
            Left            =   165
            TabIndex        =   65
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "a"
            Height          =   180
            Left            =   6030
            TabIndex        =   64
            Top             =   435
            Width           =   195
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Espessor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5760
            TabIndex        =   63
            Top             =   150
            Width           =   1080
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "a"
            Height          =   180
            Left            =   6015
            TabIndex        =   62
            Top             =   900
            Width           =   195
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "a"
            Height          =   180
            Left            =   6015
            TabIndex        =   61
            Top             =   1320
            Width           =   195
         End
      End
      Begin VB.Frame fcapcalera1 
         BackColor       =   &H00EAD9CE&
         Enabled         =   0   'False
         Height          =   7245
         Left            =   135
         TabIndex        =   38
         Top             =   840
         Width           =   8385
         Begin VB.Frame Frame3 
            BackColor       =   &H00FDDECE&
            Caption         =   "Mida de les Peces/Unitats"
            Height          =   1320
            Index           =   1
            Left            =   5685
            TabIndex        =   122
            Top             =   1740
            Width           =   2190
            Begin VB.TextBox Text3 
               Alignment       =   2  'Center
               DataField       =   "llargpeces"
               DataSource      =   "datatarifes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   780
               TabIndex        =   125
               Top             =   750
               Width           =   645
            End
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               DataField       =   "amplepeces"
               DataSource      =   "datatarifes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   780
               TabIndex        =   123
               Top             =   270
               Width           =   660
            End
            Begin VB.Label Label18 
               BackColor       =   &H00FDDECE&
               Caption         =   "Llarg:                   mm"
               Height          =   240
               Left            =   270
               TabIndex        =   126
               Top             =   810
               Width           =   1680
            End
            Begin VB.Label Label17 
               BackColor       =   &H00FDDECE&
               Caption         =   "Ample:                 mm"
               Height          =   240
               Left            =   270
               TabIndex        =   124
               Top             =   330
               Width           =   1695
            End
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H0000FFFF&
            Caption         =   "Rappel"
            Height          =   420
            Left            =   6630
            Style           =   1  'Graphical
            TabIndex        =   121
            Top             =   165
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.ComboBox combogrupdetarifa 
            DataField       =   "grupdetarifes"
            DataSource      =   "datatarifes"
            Height          =   315
            ItemData        =   "Manteniment de tarifes.frx":15A2
            Left            =   3705
            List            =   "Manteniment de tarifes.frx":15A4
            Sorted          =   -1  'True
            TabIndex        =   117
            Top             =   180
            Width           =   1995
         End
         Begin VB.ComboBox combotipusdeports 
            DataField       =   "tipusdeports"
            DataSource      =   "datatarifes"
            Height          =   315
            Left            =   1770
            TabIndex        =   104
            Top             =   5640
            Width           =   2850
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            DataField       =   "datatarifa"
            DataSource      =   "datatarifes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1005
            TabIndex        =   102
            Top             =   150
            Width           =   1125
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Clixes inclosos"
            DataField       =   "clixesinclosos"
            DataSource      =   "datatarifes"
            Height          =   240
            Left            =   2160
            TabIndex        =   97
            Top             =   4425
            Width           =   1485
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            DataField       =   "preuclixes"
            DataSource      =   "datatarifes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3000
            TabIndex        =   95
            Top             =   4755
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Transport inclòs"
            DataField       =   "transportinclos"
            DataSource      =   "datatarifes"
            Height          =   240
            Left            =   285
            TabIndex        =   94
            Top             =   5655
            Width           =   1605
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            DataField       =   "valordolar"
            DataSource      =   "datatarifes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1125
            TabIndex        =   92
            Top             =   4800
            Width           =   735
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            DataField       =   "platts"
            DataSource      =   "datatarifes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1125
            TabIndex        =   90
            Top             =   4380
            Width           =   735
         End
         Begin VB.TextBox Text5 
            DataField       =   "descripcioproducteclient"
            DataSource      =   "datatarifes"
            Height          =   345
            Left            =   2640
            TabIndex        =   86
            Top             =   1335
            Width           =   5100
         End
         Begin VB.TextBox ccodiproducte 
            DataField       =   "codiproducte"
            DataSource      =   "datatarifes"
            Height          =   285
            Left            =   2370
            TabIndex        =   82
            Top             =   795
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.ComboBox comboproducte 
            DataField       =   "descripcioproducte"
            DataSource      =   "datatarifes"
            Height          =   315
            Left            =   2640
            TabIndex        =   81
            Top             =   765
            Width           =   5160
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FDDECE&
            Height          =   735
            Left            =   150
            TabIndex        =   70
            Top             =   3615
            Width           =   6540
            Begin VB.ComboBox Combo4 
               DataField       =   "macrop"
               DataSource      =   "datatarifes"
               Height          =   315
               ItemData        =   "Manteniment de tarifes.frx":15A6
               Left            =   5670
               List            =   "Manteniment de tarifes.frx":15B0
               TabIndex        =   80
               Text            =   "N"
               Top             =   255
               Width           =   540
            End
            Begin VB.ComboBox Combo3 
               DataField       =   "microp"
               DataSource      =   "datatarifes"
               Height          =   315
               ItemData        =   "Manteniment de tarifes.frx":15BA
               Left            =   4050
               List            =   "Manteniment de tarifes.frx":15C4
               TabIndex        =   78
               Text            =   "N"
               Top             =   255
               Width           =   540
            End
            Begin VB.ComboBox Combo1 
               DataField       =   "obert"
               DataSource      =   "datatarifes"
               Height          =   315
               ItemData        =   "Manteniment de tarifes.frx":15CE
               Left            =   2775
               List            =   "Manteniment de tarifes.frx":15DB
               TabIndex        =   76
               Text            =   "N"
               Top             =   255
               Width           =   540
            End
            Begin VB.TextBox Text8 
               Alignment       =   2  'Center
               DataField       =   "solapa"
               DataSource      =   "datatarifes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1785
               TabIndex        =   72
               Top             =   255
               Width           =   420
            End
            Begin VB.TextBox Text7 
               Alignment       =   2  'Center
               DataField       =   "plegat"
               DataSource      =   "datatarifes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   645
               TabIndex        =   71
               Top             =   255
               Width           =   405
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "MACROp:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   2
               Left            =   4845
               TabIndex        =   79
               Top             =   315
               Width           =   900
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "MicroP:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   1
               Left            =   3450
               TabIndex        =   77
               Top             =   315
               Width           =   675
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Obert:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   19
               Left            =   2280
               TabIndex        =   75
               Top             =   315
               Width           =   1095
            End
            Begin VB.Label Label24 
               BackColor       =   &H00FDDECE&
               Caption         =   "Plegat:"
               Height          =   240
               Left            =   105
               TabIndex        =   74
               Top             =   315
               Width           =   585
            End
            Begin VB.Label Label23 
               BackColor       =   &H00FDDECE&
               Caption         =   "Solapa:"
               Height          =   285
               Left            =   1170
               TabIndex        =   73
               Top             =   315
               Width           =   600
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FDDECE&
            Caption         =   "Validesa de la Tarifa"
            Height          =   1050
            Left            =   150
            TabIndex        =   42
            Top             =   600
            Width           =   1830
            Begin VB.TextBox cinicivalidesa 
               Alignment       =   2  'Center
               DataField       =   "valid_inici"
               DataSource      =   "datatarifes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   420
               TabIndex        =   44
               Top             =   315
               Width           =   1275
            End
            Begin VB.TextBox cfivalidesa 
               Alignment       =   2  'Center
               DataField       =   "valid_fi"
               DataSource      =   "datatarifes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   405
               TabIndex        =   43
               Top             =   705
               Width           =   1275
            End
            Begin VB.Label Label3 
               BackColor       =   &H00FDDECE&
               Caption         =   "a:"
               Height          =   180
               Left            =   195
               TabIndex        =   46
               Top             =   750
               Width           =   315
            End
            Begin VB.Label Label4 
               BackColor       =   &H00FDDECE&
               Caption         =   "De:"
               Height          =   180
               Left            =   90
               TabIndex        =   45
               Top             =   360
               Width           =   315
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FDDECE&
            Caption         =   "Colors"
            Height          =   1890
            Index           =   0
            Left            =   3075
            TabIndex        =   41
            Top             =   1740
            Width           =   2580
            Begin VB.CommandButton Command7 
               Height          =   285
               Left            =   180
               Picture         =   "Manteniment de tarifes.frx":15E8
               Style           =   1  'Graphical
               TabIndex        =   101
               ToolTipText     =   "Eliminacio Registres"
               Top             =   225
               Width           =   345
            End
            Begin VB.Data Datacolors 
               Caption         =   "Data1"
               Connect         =   "Access"
               DatabaseName    =   ""
               DefaultCursorType=   0  'DefaultCursor
               DefaultType     =   2  'UseODBC
               Exclusive       =   0   'False
               Height          =   345
               Left            =   405
               Options         =   0
               ReadOnly        =   0   'False
               RecordsetType   =   1  'Dynaset
               RecordSource    =   ""
               Top             =   1230
               Visible         =   0   'False
               Width           =   1590
            End
            Begin MSDBGrid.DBGrid reixacolors 
               Bindings        =   "Manteniment de tarifes.frx":1B72
               Height          =   1320
               Left            =   150
               OleObjectBlob   =   "Manteniment de tarifes.frx":1B87
               TabIndex        =   99
               Top             =   525
               Width           =   2280
            End
         End
         Begin VB.TextBox cobservacions 
            DataField       =   "observacions"
            DataSource      =   "datatarifes"
            Height          =   555
            Left            =   135
            MultiLine       =   -1  'True
            TabIndex        =   40
            Top             =   6405
            Width           =   8070
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FDDECE&
            Caption         =   "Ample m/m"
            Height          =   1860
            Left            =   135
            TabIndex        =   39
            Top             =   1740
            Width           =   2790
            Begin VB.CommandButton Command6 
               Height          =   285
               Left            =   120
               Picture         =   "Manteniment de tarifes.frx":255E
               Style           =   1  'Graphical
               TabIndex        =   100
               ToolTipText     =   "Eliminacio Registres"
               Top             =   195
               Width           =   345
            End
            Begin VB.Data Dataamples 
               Caption         =   "Data1"
               Connect         =   "Access"
               DatabaseName    =   ""
               DefaultCursorType=   0  'DefaultCursor
               DefaultType     =   2  'UseODBC
               Exclusive       =   0   'False
               Height          =   345
               Left            =   210
               Options         =   0
               ReadOnly        =   0   'False
               RecordsetType   =   1  'Dynaset
               RecordSource    =   ""
               Top             =   1230
               Visible         =   0   'False
               Width           =   1590
            End
            Begin MSDBGrid.DBGrid reixaamples 
               Bindings        =   "Manteniment de tarifes.frx":2AE8
               Height          =   1320
               Left            =   105
               OleObjectBlob   =   "Manteniment de tarifes.frx":2AFD
               TabIndex        =   98
               Top             =   495
               Width           =   2550
            End
         End
         Begin VB.Label label1 
            BackColor       =   &H00EAD9CE&
            BackStyle       =   0  'Transparent
            Caption         =   "Grup de tarifa:"
            Height          =   255
            Index           =   3
            Left            =   2610
            TabIndex        =   118
            Top             =   225
            Width           =   1200
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FDDECE&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipus de ports"
            Height          =   240
            Left            =   2415
            TabIndex        =   105
            Top             =   5445
            Width           =   1050
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FDDECE&
            BackStyle       =   0  'Transparent
            Caption         =   "Data tarifa:"
            Height          =   180
            Left            =   180
            TabIndex        =   103
            Top             =   195
            Width           =   1215
         End
         Begin VB.Label Label30 
            BackColor       =   &H00FDDECE&
            BackStyle       =   0  'Transparent
            Caption         =   "Preu clixés:"
            Height          =   240
            Left            =   2100
            TabIndex        =   96
            Top             =   4815
            Width           =   1050
         End
         Begin VB.Label Label28 
            BackColor       =   &H00FDDECE&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Dollar:"
            Height          =   240
            Left            =   225
            TabIndex        =   93
            Top             =   4860
            Width           =   1050
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FDDECE&
            BackStyle       =   0  'Transparent
            Caption         =   "Platt's:  "
            Height          =   240
            Left            =   240
            TabIndex        =   91
            Top             =   4440
            Width           =   585
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripció de producte del client (com en diu ell)"
            Height          =   300
            Left            =   3000
            TabIndex        =   87
            Top             =   1125
            Width           =   4275
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripció del producte"
            Height          =   300
            Left            =   3060
            TabIndex        =   83
            Top             =   555
            Width           =   2850
         End
         Begin VB.Label Label7 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Observacions"
            Height          =   240
            Left            =   225
            TabIndex        =   47
            Top             =   6135
            Width           =   2790
         End
      End
      Begin VB.CommandButton cbotoafegirbarem 
         Height          =   320
         Left            =   -74415
         Picture         =   "Manteniment de tarifes.frx":34D4
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1365
         Width           =   930
      End
      Begin VB.CommandButton Command2 
         Height          =   975
         Left            =   -74835
         Picture         =   "Manteniment de tarifes.frx":3A5E
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1695
         Width           =   360
      End
      Begin VB.Frame Framecondicionant 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   -73455
         TabIndex        =   32
         Top             =   1335
         Visible         =   0   'False
         Width           =   1440
         Begin VB.CommandButton bcondicionantbarem 
            Height          =   285
            Left            =   30
            Picture         =   "Manteniment de tarifes.frx":3FE8
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   45
            Width           =   315
         End
         Begin VB.Label Label22 
            Caption         =   "Condicionant"
            Height          =   195
            Left            =   375
            TabIndex        =   34
            Top             =   90
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EEE4D7&
         Height          =   1170
         Left            =   -74700
         TabIndex        =   17
         Top             =   600
         Width           =   15660
         Begin VB.Frame Frame5 
            BackColor       =   &H00FDDECE&
            Height          =   990
            Left            =   60
            TabIndex        =   24
            Top             =   105
            Width           =   3060
            Begin VB.OptionButton csumaresta 
               BackColor       =   &H00FDDECE&
               Caption         =   " Caracteristica sense cost."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   2
               Left            =   105
               TabIndex        =   89
               Top             =   690
               Width           =   2910
            End
            Begin VB.OptionButton csumaresta 
               BackColor       =   &H00FDDECE&
               Caption         =   "-  Reducció de preu."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   1
               Left            =   105
               TabIndex        =   26
               Top             =   420
               Width           =   2475
            End
            Begin VB.OptionButton csumaresta 
               BackColor       =   &H00FDDECE&
               Caption         =   "+ Augment de preu."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   0
               Left            =   105
               TabIndex        =   25
               Top             =   150
               Width           =   2475
            End
         End
         Begin VB.ComboBox combocondicionants 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            ItemData        =   "Manteniment de tarifes.frx":4572
            Left            =   3270
            List            =   "Manteniment de tarifes.frx":45A0
            TabIndex        =   23
            Top             =   480
            Width           =   4560
         End
         Begin VB.TextBox ceuros 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   12765
            TabIndex        =   22
            Top             =   450
            Width           =   855
         End
         Begin VB.TextBox tv1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   8220
            TabIndex        =   21
            Top             =   450
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox tv2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   9255
            TabIndex        =   20
            Top             =   450
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox tv3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   10290
            TabIndex        =   19
            Top             =   450
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Acceptar"
            Height          =   645
            Left            =   14250
            Picture         =   "Manteniment de tarifes.frx":4715
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   255
            Width           =   1140
         End
         Begin VB.Label etvalortexte 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8055
            TabIndex        =   88
            Top             =   540
            Width           =   4605
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Condicionant"
            Height          =   300
            Left            =   3300
            TabIndex        =   31
            Top             =   240
            Width           =   3780
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Euros"
            Height          =   285
            Left            =   12960
            TabIndex        =   30
            Top             =   225
            Width           =   885
         End
         Begin VB.Label etv1 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor 1"
            Height          =   285
            Left            =   8475
            TabIndex        =   29
            Top             =   210
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label etv2 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor2"
            Height          =   285
            Left            =   9435
            TabIndex        =   28
            Top             =   225
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label etv3 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor3"
            Height          =   285
            Left            =   10470
            TabIndex        =   27
            Top             =   225
            Visible         =   0   'False
            Width           =   885
         End
      End
      Begin VB.CommandButton beliminarcond 
         Height          =   600
         Left            =   -75000
         Picture         =   "Manteniment de tarifes.frx":4C9F
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2115
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Kg"
         Height          =   870
         Left            =   -70875
         Picture         =   "Manteniment de tarifes.frx":5229
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Nou nivell d'escalat (fila)"
         Top             =   1605
         Width           =   420
      End
      Begin VB.CommandButton Command4 
         Height          =   360
         Left            =   -70425
         Picture         =   "Manteniment de tarifes.frx":57B3
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Nova tarifa de ports"
         Top             =   1110
         Width           =   1050
      End
      Begin MSFlexGridLib.MSFlexGrid reixaports 
         Height          =   7245
         Left            =   -70455
         TabIndex        =   12
         Top             =   1470
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   12779
         _Version        =   393216
         BackColor       =   2486189
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid reixacond 
         Height          =   5850
         Left            =   -74625
         TabIndex        =   16
         Top             =   2085
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   10319
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid reixa 
         Height          =   7575
         Left            =   -74490
         TabIndex        =   36
         Top             =   1665
         Width           =   15810
         _ExtentX        =   27887
         _ExtentY        =   13361
         _Version        =   393216
         BackColor       =   16777215
      End
      Begin VB.Label etversio 
         BackStyle       =   0  'Transparent
         Height          =   270
         Left            =   1335
         MouseIcon       =   "Manteniment de tarifes.frx":5D3D
         MousePointer    =   99  'Custom
         TabIndex        =   85
         Top             =   465
         Width           =   600
      End
      Begin VB.Label label1 
         BackColor       =   &H00EAD9CE&
         BackStyle       =   0  'Transparent
         Caption         =   "Tarifa:"
         Height          =   165
         Index           =   0
         Left            =   180
         TabIndex        =   69
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarifes de transports per aquest client:"
         Height          =   240
         Left            =   -73740
         TabIndex        =   14
         Top             =   405
         Width           =   3165
      End
      Begin VB.Label etcalcultamanyfont 
         AutoSize        =   -1  'True
         Caption         =   "no borrar"
         Height          =   195
         Left            =   -66465
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   630
      End
   End
   Begin VB.Frame Frame9 
      Height          =   510
      Left            =   105
      TabIndex        =   4
      Top             =   405
      Width           =   1755
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   450
         Picture         =   "Manteniment de tarifes.frx":62C7
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Consulta Registres"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   30
         Picture         =   "Manteniment de tarifes.frx":6851
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Alta  Registres"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton Command3 
         Height          =   360
         Left            =   1290
         Picture         =   "Manteniment de tarifes.frx":6DDB
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   870
         Picture         =   "Manteniment de tarifes.frx":7365
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Eliminacio Registres"
         Top             =   105
         Width           =   420
      End
   End
   Begin VB.Label etestat 
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
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   1980
      TabIndex        =   9
      Top             =   585
      Width           =   1680
   End
End
Attribute VB_Name = "form_tarifes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vultimcolorcol As Double

Private Sub alta_Click()
  Dim vtarifa As String
  If datatarifes.Recordset.EditMode > 0 Then MsgBox "Ja s'està editant una tarifa.", vbCritical, "Error": Exit Sub
  vtarifa = proximnumerodetarifa(IIf(cadbl(comboclient.tag) = 0, comboclient, cadbl(comboclient.tag)))
  vtarifa = InputBox("Escriu el numero de tarifa que vols utilitzar.", "Nova tarifa", vtarifa)
  If StrPtr(vtarifa) = 0 Then Exit Sub
   fcapcalera1.Enabled = True
   fcapcalera2.Enabled = True
   Frameescalat.Enabled = True
   datatarifes.Recordset.AddNew
   datatarifes.Recordset!numerotarifa = cadbl(vtarifa)
   datatarifes.Recordset!versio = 1
   ctarifa = vtarifa
   etversio = " v" + atrim(datatarifes.Recordset!versio)
   datatarifes.Recordset.Update
   datatarifes.Recordset.Bookmark = datatarifes.Recordset.LastModified
   etestat = "Afegint..."
   datatarifes.Recordset.Edit
   datatarifes.Recordset!client = cadbl(comboclient.tag)
   If cadbl(comboclient.tag) = 0 Then datatarifes.Recordset!grupclients = (comboclient)
   Pestanyes.Tab = 0

End Sub
Function proximnumerodetarifa(vcodiclient As String) As Double
   Dim rst As Recordset
   Dim vsql As String
   If cadbl(vcodiclient) > 0 Then
        vsql = "select max(numerotarifa) as maxnumero from tarifes_capcalera where client=" + atrim(vcodiclient)
         Else: vsql = "select max(numerotarifa) as maxnumero from tarifes_capcalera where grupclients='" + atrim(vcodiclient) + "'"
   End If
   Set rst = dbtarifes.OpenRecordset(vsql)
   proximnumerodetarifa = cadbl(rst!maxnumero) + 1
   Set rst = Nothing
End Function

Private Sub bcondicionantbarem_Click()
   Dim v1 As String
   Dim v2 As String
   Dim vcolactiva As Long
   Dim vCONDICIO As String
   Dim vOPERACIO  As String
   bcondicionantbarem.tag = ""
   form_escullircondicionant.Show 1
   If atrim(bcondicionantbarem.tag) <> "" Then
      vOPERACIO = Mid(atrim(bcondicionantbarem.tag) + "    ", 1, 1)
      vCONDICIO = Mid(atrim(bcondicionantbarem.tag) + "    ", 2, 4)
      If vCONDICIO = "[NC]" Or vCONDICIO = "[EM]" Then demanar_v1_v2 v1, v2, Mid(atrim(bcondicionantbarem.tag) + "  ", 2)
      vcolactiva = reixa.col
      crear_elcondicionant reixa.ColData(reixa.col), Mid(atrim(bcondicionantbarem.tag) + "  ", 2), v1, v2, vOPERACIO
      possar_dades_tarifes
      reixa.col = vcolactiva
   End If
End Sub
Sub demanar_v1_v2(v1 As String, v2 As String, vcriteri)
    v1 = InputBox("Entra el valor DE:" + vbNewLine + vcriteri, "Valor de")
    v2 = InputBox("Entra el valor A:" + vbNewLine + vcriteri, "Valor a")
End Sub
Sub crear_elcondicionant(vidlinia As Long, vcondicionant As String, v1 As String, v2 As String, vOPERACIO As String)
    Dim rst As Recordset
    Dim rstn As Recordset
    vcondicionant = atrim(Mid(vcondicionant + "   ", 1, 4))
    If vcondicionant = "" Then Exit Sub
    Set rst = dbtarifes.OpenRecordset("select * from tarifes_barem where idliniabarem=" + atrim(vidlinia))
    If rst.EOF Then GoTo fi
    Set rst = dbtarifes.OpenRecordset("select * from tarifes_barem where idtarifa=" + atrim(rst!idtarifa) + " and valor1=" + atrim(rst!valor1) + " and valorcondicionant='" + atrim(rst!valorcondicionant) + "'")
    Set rstn = dbtarifes.OpenRecordset("select * from tarifes_barem")
    While Not rst.EOF
      rstn.AddNew
      rstn!idtarifa = rst!idtarifa
      rstn!linkvalor1condicionant = rst!valor1
      rstn!linkvalorcondicionant = atrim(rst!valorcondicionant)
      rstn!valorcondicionant = vcondicionant
      rstn!mes_menys = vOPERACIO
      rstn!valor1 = cadbl(v1)
      rstn!valor2 = cadbl(v2)
      rstn!desde_kg = rst!desde_kg
      rstn.Update
      rst.MoveNext
    Wend
fi:
    Set rst = Nothing
    Set rstn = Nothing
End Sub

Private Sub beliminarcond_Click()
  If MsgBox("Segur que vols eliminar aquest condicionant?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then GoTo fi
  dbtarifes.Execute "delete * from tarifes_condicionants where id=" + atrim(reixacond.TextMatrix(reixacond.Row, 2))
  carregar_condicionants
fi:
  beliminarcond.visible = False
End Sub

Private Sub botopegar_Click()
   Dim i As Byte
   Dim v As String
   Dim vcanvis As Long
   Dim vector(100) As String
   Dim vaccio As String
   
   If Not Clipboard.GetFormat(vbCFText) Then MsgBox "Valors del portapapers no vàlids.", vbCritical, "Error": Exit Sub
   If reixa.Cols = 0 Then
      If MsgBox("NO HI HA ESCALAT, ES COPIARAN ELS valors que estiguin al portapapers com a ESCALAT." + vbNewLine + "Vols continuar?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
      vaccio = "ESC"
        Else:
          If MsgBox("Es copiarant tots els valors que estiguin al portapapers a la reixa a partir de la casella sel.lecionada ara mateix." + vbNewLine + "Vols continuar?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
          vaccio = "VAL"
   End If
   v = Clipboard.GetText
   passarvalorsalvector v, vector
   While vector(i) <> vbNewLine
     If vaccio = "VAL" Then
        posar_valor_casella cadbl(vector(i))
        If reixa.Row + 1 < reixa.Rows - 1 Then
            reixa.Row = reixa.Row + 1
            Else: If vector(i + 1) <> vbNewLine Then MsgBox "No hi caben tots els valors.", vbCritical, "Error": GoTo fi
        End If
     End If
     If vaccio = "ESC" Then
       posarvalorescalat cadbl(vector(i))
       If reixa.Row + 1 = reixa.Rows Then reixa.Rows = reixa.Rows + 1
       If reixa.Row + 1 < reixa.Rows Then reixa.Row = reixa.Row + 1
     End If
     i = i + 1
   Wend
fi:
   possar_dades_tarifes
End Sub
Sub passarvalorsalvector(v As String, vector() As String)
   Dim valor As String
   Dim i As Long
   valor = v
   i = 0
   While InStr(1, valor, vbNewLine) > 0
      vector(i) = Mid(valor, 1, InStr(1, valor, vbNewLine) - 1)
      valor = Mid(valor, InStr(1, valor, vbNewLine) + 1)
      i = i + 1
   Wend
   vector(i) = vbNewLine
End Sub
Private Sub cbotoafegirbarem_Click()
 Dim vref As Double
 Dim vvalor2 As String
  Dim rst As Recordset
  Dim vkg As Double
  Dim vunitat As String
  Set rst = dbtarifes.OpenRecordset("select * from tarifes_barem where idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa)) + " order by desde_kg")
  If rst.EOF Then MsgBox "Primer hi ha d'haver entrat almenys un valor de KG.", vbCritical, "Error": GoTo fi
  vkg = rst!desde_kg
  vunitat = escullir_unitatpetita
  If vunitat = "" Then Exit Sub
  vref = cadbl(InputBox("Entra el rang [DE->A]  De que vols.", "Entra els Referència"))
  If vref = 0 Then Exit Sub
  rst.FindFirst "valor1=" + atrim(vref) + " and valorcondicionant='" + vunitat + "'"
  If Not rst.NoMatch Then MsgBox "Aquestes referències ja està entrada.", vbCritical, "Error": GoTo fi
  vvalor2 = InputBox("Entra el valor del rang A que vols.", "Entra els Referència")
  If StrPtr(vvalor2) = 0 Then Exit Sub
  rst.AddNew
  rst!desde_kg = vkg
  rst!preu_kg = 0
  rst!valor1 = vref
  rst!valor2 = cadbl(vvalor2)
  rst!idtarifa = datatarifes.Recordset!idtarifa
  rst!valorcondicionant = Mid(vunitat, 3)
  rst.Update
  carregar_reixa_barem
fi:
  Set rst = Nothing
End Sub
Function escullir_unitatpetita() As String

  Load formseleccio
  formseleccio.caption = "Escull la unitat"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "tarifes.mdb"
  formseleccio.Data1.RecordSource = "select unitat,unitatpetita from barem_unitats order by unitat"
  
  formseleccio.refrescar
    formseleccio.dbgrid1_clickcapçalera (0)
  formseleccio.DBGrid2.Columns(0).width = 2000
  formseleccio.DBGrid2.Columns(1).visible = False
  
  formseleccio.Show 1
  If seleccioret = 1 Then
   escullir_unitatpetita = atrim(formseleccio.Data1.Recordset!unitat)
  End If
  Unload formseleccio
  
End Function
Private Sub cbotoafegirdesti_Click()

End Sub

Private Sub cfivalidesa_GotFocus()
   cfivalidesa.tag = cfivalidesa
End Sub

Private Sub cfivalidesa_LostFocus()
   If Not IsDate(cfivalidesa) Then
        MsgBox "Aquesta data no es correcte", vbCritical, "Error"
        cfivalidesa = cfivalidesa.tag
   End If
End Sub

Private Sub Check1_Click()
  If Check1.Value = 1 Then Pestanyes.Tab = 0: Pestanyes.TabEnabled(3) = False
  If Check1.Value = 0 Then Pestanyes.Tab = 0: Pestanyes.TabEnabled(3) = True
End Sub

Private Sub cinicivalidesa_GotFocus()
   cinicivalidesa.tag = cinicivalidesa
End Sub

Private Sub cinicivalidesa_LostFocus()
   If Not IsDate(cinicivalidesa) Then
        MsgBox "Aquesta data no es correcte", vbCritical, "Error"
        cinicivalidesa = cinicivalidesa.tag
   End If
End Sub

Private Sub Combo2_Change()

End Sub

Private Sub comboclient_Click()
  If comboclient = "- Escullir client -" And Screen.ActiveControl.Name = "comboclient" Then
    ctarifa.SetFocus
    Timer_escullirclient.Enabled = True
    
   ' datatarifes.Recordset.Move 0
     Else: comboclient.tag = comboclient
  End If
  carregar_tarifesclient
End Sub
Function crearllistaidtarifes(vsql As String) As String
  Dim rst As Recordset
  Dim vtarifa As Double
  Dim v As String
  Set rst = dbtarifes.OpenRecordset(vsql)
  While Not rst.EOF
    vtarifa = rst!numerotarifa
    crearllistaidtarifes = crearllistaidtarifes + IIf(crearllistaidtarifes <> "", "," + atrim(rst!idtarifa), atrim(rst!idtarifa))
    While vtarifa = rst!numerotarifa
      rst.MoveNext
      If rst.EOF Then GoTo fi
    Wend
    crearllistaidtarifes = crearllistaidtarifes + IIf(crearllistaidtarifes <> "", "," + atrim(rst!idtarifa), atrim(rst!idtarifa))
    rst.MoveNext
  Wend
fi:
  Set rst = Nothing
End Function
Sub carregar_tarifesclient(Optional vversio As String)
  Dim vsql As String
  datatarifes.RecordSource = "select * from tarifes_capcalera where idtarifa=0"
If cadbl(comboclient.tag) > 0 Then
   vsql = crearllistaidtarifes("SELECT * From Tarifes_capcalera where client=" + comboclient.tag + " order by Tarifes_capcalera.numerotarifa,versio desc;")
   vsql = " and idtarifa in (" + vsql + ")"
   If InStr(1, vsql, "()") = 0 Then
     If vversio <> "" Then vsql = vversio
     datatarifes.RecordSource = "select * from tarifes_capcalera where client=" + comboclient.tag + vsql
     
   End If
    Else:
       vsql = crearllistaidtarifes("SELECT * From Tarifes_capcalera where grupclients='" + comboclient.tag + "' order by Tarifes_capcalera.numerotarifa,versio desc;")
       vsql = " and idtarifa in (" + vsql + ")"
       If InStr(1, vsql, "()") = 0 Then
         If vversio <> "" Then vsql = vversio
         datatarifes.RecordSource = "select * from tarifes_capcalera where grupclients='" + comboclient.tag + "' " + vsql
        
       End If
  End If
  datatarifes.Refresh
   
End Sub
Sub escullir_client()

  Load formseleccio
  
  formseleccio.caption = "Escull el client"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "tarifes.mdb"
  formseleccio.Data1.RecordSource = "select codi,nom from clients order by nom"
  
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 800
  formseleccio.DBGrid2.Columns(1).width = 3500

  
  formseleccio.Show 1
  If seleccioret = 1 Then
   comboclient = atrim(formseleccio.Data1.Recordset!nom)
   comboclient.tag = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   'datatarifes.Recordset.Edit
   'datatarifes.Recordset!client = formseleccio.Data1.Recordset!codi
   'datatarifes.Recordset.Update
  End If
  Unload formseleccio
  
End Sub

Sub ensenya_amaga_valorscondicionants(vnumvalor As Byte, vcaptionetiqueta As String, vamaga As Boolean)
   etvalortexte = ""
   If vnumvalor = 1 Then
      tv1.visible = Not vamaga
      etv1.visible = Not vamaga
      etv1.caption = vcaptionetiqueta
   End If
   
   If vnumvalor = 2 Then
      tv2.visible = Not vamaga
      etv2.visible = Not vamaga
      etv2.caption = vcaptionetiqueta
   End If
   
   If vnumvalor = 3 Then
      tv3.visible = Not vamaga
      etv3.visible = Not vamaga
      etv3.caption = vcaptionetiqueta
   End If
   
   
End Sub

Private Sub comboclient_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub combocondicionants_Click()
  ensenya_amaga_valorscondicionants 1, "", True
  ensenya_amaga_valorscondicionants 2, "", True
  ensenya_amaga_valorscondicionants 3, "", True
  If Mid(combocondicionants, 1, 4) = "[EM]" Then ensenya_amaga_valorscondicionants 1, "Micres", False
  If Mid(combocondicionants, 1, 4) = "[NC]" Then
     ensenya_amaga_valorscondicionants 1, "de Colors", False
     ensenya_amaga_valorscondicionants 2, "a Colors", False
  End If
  If Mid(combocondicionants, 1, 4) = "[MK]" Then
      etvalortexte = escullir_marca
  End If
  tv1 = ""
  tv2 = ""
  tv3 = ""
  ceuros = ""
  
End Sub
Function escullir_marca() As String
  Load formseleccio
  formseleccio.caption = "Escull la marca"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "comandes.mdb"
  formseleccio.Data1.RecordSource = "select distinct marca from clixes order by marca"
  formseleccio.sortirs.tag = "filtre"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 5500
  formseleccio.width = 6500

  
  formseleccio.Show 1
  If seleccioret = 1 Then
   escullir_marca = atrim(formseleccio.Data1.Recordset!marca)
  End If
  Unload formseleccio
  SendKeys "{TAB}"
   
End Function

Private Sub combogrupdetarifa_LostFocus()
  If Len(combogrupdetarifa) > 20 Then combogrupdetarifa = Mid(combogrupdetarifa, 1, 20)
End Sub

Private Sub Combomat1_DropDown()
   escullir_familia tmat1, Combomat1
End Sub
Sub escullir_familia(cmat1 As TextBox, ccombomat As ComboBox)
   Load formseleccio
  
  formseleccio.caption = "Escull familia"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "comandes.mdb"
  formseleccio.Data1.RecordSource = "select codi,descripcio from familiesmaterials order by descripcio"
  formseleccio.Data1.RecordSource = "select primercodi,descripcio_mat from llistatmaterialsdetallat order by descripcio_mat"
  formseleccio.sortirs.tag = "filtre"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 0
  formseleccio.DBGrid2.Columns(1).width = 7500
  formseleccio.width = 8800

  
  formseleccio.Show 1
  If seleccioret = 1 Then
   ccombomat.Text = atrim(formseleccio.Data1.Recordset!descripcio_mat)
   cmat1 = atrim(cadbl(formseleccio.Data1.Recordset!primercodi))
   'datatarifes.Recordset.Edit
   'datatarifes.Recordset!client = formseleccio.Data1.Recordset!codi
   'datatarifes.Recordset.Update
  End If
  Unload formseleccio
  SendKeys "{TAB}"
End Sub

Private Sub combomat2_DropDown()
escullir_familia tmat2, combomat2
End Sub

Private Sub combomat3_DropDown()
escullir_familia tmat3, combomat3
End Sub

Sub carregar_tarifes()
   Dim rst As Recordset
   Dim vsql As String
   Dim vestat As String
   
   vsql = IIf(cadbl(comboclient.tag) > 0, " codiclient=" + atrim(cadbl(comboclient.tag)), " grupclients='" + atrim(comboclient) + "'")
   Set rst = dbtarifes.OpenRecordset("select * from tarifes_ports_capcalera where " + vsql + " order by datainici desc")
   combotarifesports.Clear
   While Not rst.EOF
     vestat = IIf(Now > rst!datainici And Now < rst!datafi, " Vigent.", "Caducada.")
     combotarifesports.AddItem Format(rst!datainici, "dd/mm/yy") + "~" + Format(rst!datafi, "dd/mm/yy") + " " + vestat
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Sub escullir_producte()
  Load formseleccio
 
  formseleccio.caption = "Escull el producte"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "comandes.mdb"
  formseleccio.Data1.RecordSource = "select codi,descripcio from productes order by descripcio"
  
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 800
  formseleccio.DBGrid2.Columns(1).width = 3500

  
  formseleccio.Show 1
  If seleccioret = 1 Then
   comboproducte = atrim(formseleccio.Data1.Recordset!descripcio)
   ccodiproducte = atrim(formseleccio.Data1.Recordset!codi)
  End If
  Unload formseleccio
End Sub

Private Sub comboproducte_DropDown()
    escullir_producte
End Sub

Private Sub combotarifesports_DropDown()
   carregar_tarifes
End Sub

Private Sub combotipusdeports_LostFocus()
   If Len(combotipusdeports) > 30 Then combotipusdeports = Mid(combotipusdeports, 1, 30)
End Sub

Private Sub combounitatescalat_Click()
  If datatarifes.Recordset.EditMode = 0 Then MsgBox "No pots canviar sense editar la tarifa.", vbCritical, "Error": Exit Sub
  possar_dades_tarifes
  If combounitatfacturacio = "" Then
    combounitatfacturacio = combounitatescalat
  End If
  datatarifes.Recordset!pesnet = False
  If InStr(1, UCase(combounitatescalat + " "), "K ") > 0 Then
        If clientespesnet Then datatarifes.Recordset!pesnet = True
  End If
  comprovar_si_peces
End Sub
Sub comprovar_si_peces()
  If InStr(1, UCase(combounitatescalat + " "), "U ") > 0 Then
        cdesarrollpcs.visible = True
        etdesarrollpcs.visible = True
          Else
            cdesarrollpcs.visible = False
            etdesarrollpcs.visible = False
  End If
End Sub
Function clientespesnet() As Boolean
    Dim rst As Recordset
    'pesnetbrut
    vcodis = llistaclientsdelatarifa
    If vcodis = "" Then Exit Function
    Set rst = dbtarifes.OpenRecordset("select pesnetbrut from clients_envios where codi in (" + vcodis + ")")
    If Not rst.EOF Then If rst!pesnetbrut Then clientespesnet = True
    Set rst = Nothing
End Function
Function llistaclientsdelatarifa() As String
   Dim rst As Recordset
   If datatarifes.Recordset.EOF Then Exit Function
   If atrim(datatarifes.Recordset!grupclients) <> "" Then
      Set rst = dbtarifes.OpenRecordset("select codi from clients where grupdeclient='" + atrim(datatarifes.Recordset!grupclients) + "'")
       Else: Set rst = dbtarifes.OpenRecordset("select codi from clients where codi=" + atrim(datatarifes.Recordset!client))
   End If
   While Not rst.EOF
      llistaclientsdelatarifa = llistaclientsdelatarifa + IIf(llistaclientsdelatarifa <> "", ",", "") + atrim(rst!codi)
      rst.MoveNext
   Wend
   Set rst = Nothing
End Function
Private Sub Command1_Click()
  guardar_condicionant
  carregar_condicionants
  
End Sub
Sub netejar_condicionants()
  tv1 = ""
  tv2 = ""
  tv3 = ""
  tv1.visible = False
  tv2.visible = False
  tv3.visible = False
  etvalortexte = ""
  ceuros = ""
  combocondicionants = ""
  csumaresta(0) = False
  csumaresta(1) = False
End Sub
Function falten_condicionants() As Boolean
  If csumaresta(0).Value = False And csumaresta(1).Value = False And csumaresta(2).Value = False Then MsgBox "S'ha d'escullir Augment, Reducció o caracteristica", vbCritical, "Error": falten_condicionants = True
  If combocondicionants = "" Then MsgBox "S'ha de escullir un condicionant.", vbCritical, "Error": falten_condicionants = True
  If tv1.visible And cadbl(tv1) = 0 Then MsgBox "Falta el valor1.", vbCritical, "Error": falten_condicionants = True
  If tv2.visible And cadbl(tv2) = 0 Then MsgBox "Falta el valor2.", vbCritical, "Error": falten_condicionants = True
  If tv3.visible And cadbl(tv3) = 0 Then MsgBox "Falta el valor3.", vbCritical, "Error": falten_condicionants = True
  If cadbl(ceuros) = 0 And csumaresta(2).Value = 0 Then MsgBox "Falta el valor del EUROS.", vbCritical, "Error": falten_condicionants = True
End Function
Sub guardar_condicionant()
  Dim rst As Recordset
  If falten_condicionants Then Exit Sub
  Set rst = dbtarifes.OpenRecordset("select * from tarifes_condicionants")
  rst.AddNew
  rst!idtarifa = datatarifes.Recordset!idtarifa
  rst!condicionant = combocondicionants
  rst!mes_menys = IIf(csumaresta(0).Value = True, "A", IIf(csumaresta(1).Value = True, "R", IIf(csumaresta(2).Value = True, "C", "")))
  rst!valorc1 = cadbl(tv1)
  rst!valorc2 = cadbl(tv2)
  rst!valorc3 = cadbl(tv3)
  rst!valortexte = atrim(etvalortexte)
  rst!valoreuros = cadbl(ceuros)
  rst.Update
  netejar_condicionants
  Set rst = Nothing
End Sub
Private Sub Command2_Click()
  posarvalorescalat
End Sub
Sub posarvalorescalat(Optional vkg As Double)
  Dim rst As Recordset
  Dim vvalor1 As Double
  Dim vhapassatelvalor As Boolean
  
  If vkg > 0 Then vhapassatelvalor = True
  If combounitatfacturacio = "" Then MsgBox "Primer has d'escullir la unitat de facturació", vbCritical, "Error": Exit Sub
  If combounitatescalat = "" Then MsgBox "Primer has d'escullir la unitat de l'escalat", vbCritical, "Error": Exit Sub
  If Not vhapassatelvalor Then vkg = cadbl(InputBox("Entra els Kg que vols.", "Entra els kg"))
  If vkg = 0 Then Exit Sub
  Set rst = dbtarifes.OpenRecordset("select * from tarifes_barem where idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa)) + " and linkvalor1condicionant=0 order by valor1")
  rst.FindFirst "desde_kg=" + passaradecimalpunt(atrim(vkg))
  If Not rst.NoMatch Then MsgBox "Aquest valor d'escalat (" + atrim(vkg) + " ja està entrat.", vbCritical, "Error": GoTo fi
  If Not rst.EOF Then
     vvalor1 = rst!valor1
      Else: vvalor1 = 1
  End If
  rst.AddNew
  rst!desde_kg = vkg
  rst!preu_kg = 0
  'rst!valor1 = vvalor1
  rst!valor1 = 0
  rst!idtarifa = datatarifes.Recordset!idtarifa
  rst!valorcondicionant = "*Esc"
  rst.Update
  If Not vhapassatelvalor Then carregar_reixa_barem
fi:
  Set rst = Nothing
End Sub

Private Sub Command3_Click()
   fcapcalera1.Enabled = False
   fcapcalera2.Enabled = False
   Frameescalat.Enabled = False
   etestat = ""
   If datatarifes.Recordset.EditMode > 0 Then
     datatarifes.Recordset.Update
     possar_dades_tarifes
   End If
End Sub

Private Sub Command4_Click()
  crear_tarifadeports
  possar_dades_ports
End Sub
Sub crear_tarifadeports()
  Dim vinici As String
  Dim vfi As String
  Dim rst As Recordset
  Dim rstdesti As Recordset
  Dim vsql As String
  Dim vnomtarifa As String
  Dim vidtarifa As Long
  
  vsql = IIf(cadbl(comboclient.tag) > 0, " codiclient=" + atrim(cadbl(comboclient.tag)), " grupclients='" + atrim(comboclient) + "'")
  vnomtarifa = UCase(atrim(InputBox("Escriu el nom que vols possar a aquesta tarifa.", "Nom de la tarifa")))
  If vnomtarifa = "" Then Exit Sub
  If Len(vnomtarifa) > 20 Then MsgBox "El nom no pot tenir mes de 20 caràcters.", vbCritical, "Error": Exit Sub
  vnomtarifa = treure_apostruf(vnomtarifa)
  Set rst = dbtarifes.OpenRecordset("select * from tarifes_ports_capcalera where ucase(nomdelatarifa)='" + vnomtarifa + "' and " + vsql)
  If Not rst.EOF Then MsgBox "Aquest nom de tarifa ja existeix", vbCritical, "Error": Exit Sub
  vinici = InputBox("Entra la data d'inici de validesa de la tarifa de ports.", "Inici Validesa")
  If StrPtr(vinici) = 0 Then Exit Sub
  If Not IsDate(vinici) Then MsgBox "Aquesta data no ès vàlida.", vbCritical, "Error": Exit Sub
  vfi = InputBox("Entra la data de fi de validesa de la tarifa de ports.", "Fi Validesa")
  If StrPtr(vfi) = 0 Then Exit Sub
  If Not IsDate(vfi) Then MsgBox "Aquesta data no ès vàlida.", vbCritical, "Error": Exit Sub
  If Now < vinici And Now < vfi Then If MsgBox("Aquestes dates son inferiors a la data actual, la tarifa no serà vàlida." + vbNewLine + " VOLS UTILITZAR AQUESTES DATES IGUALMENT?", vbExclamation + vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbNo Then Exit Sub
  
  
  
  'If Not rst.EOF Then
      rst.AddNew
      vidtarifa = rst!idtarifaports
      rst!datainici = vinici
      rst!datafi = vfi
      rst!codiclient = cadbl(comboclient.tag)
      rst!nomdelatarifa = vnomtarifa
      rst!grupclients = " "
      If rst!codiclient = 0 Then rst!grupclients = comboclient
      rst.Update
      rst.FindFirst "idtarifaports=" + atrim(vidtarifa)
      vsql = "SELECT DISTINCT Clients_envios.provinciae FROM clients INNER JOIN Clients_envios ON clients.codi = Clients_envios.codi WHERE "
      vsql = vsql + IIf(rst!codiclient > 0, " codi=" + atrim(rst!codiclient), " grupdeclient='" + atrim(rst!grupclients) + "'")
      'Set rstdesti = dbtarifes.OpenRecordset(vsql)
      'While Not rstdesti.EOF
      '   dbtarifes.Execute "insert into tarifes_ports (idtarifaports,desde_kg,preu_kg) values (" + atrim(vidtarifa) + ",0,0)"
      '   rstdesti.MoveNext
      'Wend
 ' End If
  Set rstdesti = Nothing
  Set rst = Nothing
  
End Sub

Private Sub Command5_Click()
  Dim vkg As Double
  Dim rst As Recordset
  If reixaports.Cols = 1 Then MsgBox "Primer has d'afegir una tarifa i despres el barem.", vbCritical, "Error": Exit Sub
  vkg = cadbl(InputBox("Entra els Kg que vols.", "Entra els kg"))
  If vkg = 0 Then Exit Sub
  Set rst = dbtarifes.OpenRecordset("select * from tarifes_ports where idtarifaports=" + atrim(cadbl(datatarifes.Recordset!idtarifaports)) + " order by desde_kg")
  rst.FindFirst "desde_kg=" + atrim(vkg)
  If Not rst.NoMatch Then MsgBox "Aquest valor de Kg ja està entrada.", vbCritical, "Error": GoTo fi
  rst.AddNew
  rst!desde_kg = vkg
  rst!preu_kg = 0
  rst!idtarifaports = reixaports.ColData(1)
  rst.Update
  carregar_reixa_ports
fi:
  Set rst = Nothing
End Sub

Private Sub Command7_Click()
   If Datacolors.Recordset.EOF Then MsgBox "Escull un color per eliminar", vbCritical, "Error": Exit Sub
   If MsgBox("Segur que vols eliminar aquest color?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
        Datacolors.Recordset.Delete
        Datacolors.Refresh
   End If
End Sub

Private Sub Command6_Click()
If Dataamples.Recordset.EOF Then MsgBox "Escull una amplada per eliminar", vbCritical, "Error": Exit Sub
   If MsgBox("Segur que vols eliminar aquesta amplada?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
        Dataamples.Recordset.Delete
        Dataamples.Refresh
   End If
End Sub

Private Sub Command8_Click()
  If datatarifes.Recordset.EOF Then Exit Sub
  ensenyar_rappel IIf(atrim(datatarifes.Recordset!grupclients) <> "", "grupdeclients='" + atrim(datatarifes.Recordset!grupclients) + "'", "codiclient=" + atrim(cadbl(datatarifes.Recordset!client))), True

End Sub
Sub ensenyar_rappel(vsql As String, vensenyar As Boolean)
  Dim rst As Recordset
  Dim vmsg As String
  Set rst = dbtarifes.OpenRecordset("select * from clients_rappels where " + vsql)
  If Not vensenyar Then
     If Not rst.EOF Then
           Command8.visible = True
          Else: Command8.visible = False
     End If
     If vsql = "codiclient=0" Then Command8.visible = False
     GoTo fi
  End If
  While Not rst.EOF
    vmsg = vmsg + "Fins " + justificar(Format(cadbl(rst!fins), "#,##0") + "€", 10, "D") + "  <->   " + justificar(atrim(rst!tanx100) + "%", 5, "D") + vbNewLine
    rst.MoveNext
  Wend
  If vmsg <> "" Then MsgBox vmsg, vbInformation, "R A P P E L"
fi:
  Set rst = Nothing
End Sub

Private Sub datatarifes_Reposition()
  carregar_combotipusdeports
  carregar_combogruptarifes
  possar_dades_tarifes
  possar_dades_ports
  carregar_condicionants
End Sub
Sub carregar_combotipusdeports()
    Dim rst As Recordset
    Dim vsql As String
    Dim v As String
    v = combotipusdeports
    combotipusdeports.Clear
    'vsql = IIf(cadbl(comboclient.tag) > 0, " codiclient=" + atrim(cadbl(comboclient.tag)), " grupclients='" + atrim(comboclient) + "'")
    Set rst = dbtarifes.OpenRecordset("select distinct tipusdeports as Tipus from tarifes_capcalera")
    While Not rst.EOF
       If atrim(rst!tipus) <> "" Then combotipusdeports.AddItem atrim(rst!tipus)
       rst.MoveNext
    Wend
    Set rst = Nothing
    combotipusdeports = v
End Sub
Sub carregar_combogruptarifes()
    Dim rst As Recordset
    Dim vsql As String
    Dim v As String
    v = combogrupdetarifa
    combogrupdetarifa.Clear
    vsql = IIf(cadbl(comboclient.tag) > 0, " client=" + atrim(cadbl(comboclient.tag)), " grupclients='" + atrim(comboclient) + "'")
    
    Set rst = dbtarifes.OpenRecordset("select distinct grupdetarifes as Gtarifa from tarifes_capcalera where " + vsql)
    While Not rst.EOF
       If atrim(rst!Gtarifa) <> "" Then combogrupdetarifa.AddItem atrim(rst!Gtarifa)
       rst.MoveNext
    Wend
    Set rst = Nothing
    combogrupdetarifa = v
End Sub
Sub possar_dades_tarifes()
   Dataamples.RecordSource = "select * from barem_tipusx where tipus=''"
   Dataamples.Refresh
   Datacolors.RecordSource = "select * from barem_tipusx where tipus='' "
   Datacolors.Refresh
   reixa.Clear
   reixa.Cols = 0
   reixa.Rows = 1
   
   ctarifa = ""
   If datatarifes.Recordset.EOF Then
     Exit Sub
   End If
   
   Dataamples.RecordSource = "select * from barem_tipusx where tipus='A' and idtarifa=" + atrim(datatarifes.Recordset!idtarifa) + " order by de"
   Dataamples.Refresh
   
   Datacolors.RecordSource = "select * from barem_tipusx where tipus='C' and idtarifa=" + atrim(datatarifes.Recordset!idtarifa) + " order by de"
   Datacolors.Refresh
   
   ctarifa = Format(datatarifes.Recordset!numerotarifa, "000")
   etversio = " v" + atrim(datatarifes.Recordset!versio)
   If cadbl(datatarifes.Recordset!client) > 0 Then comboclient = nomdelclient(cadbl(datatarifes.Recordset!client))
   carregar_reixa_barem
   If Not datatarifes.Recordset.EOF Then ensenyar_rappel IIf(atrim(datatarifes.Recordset!grupclients) <> "", "grupdeclients='" + atrim(datatarifes.Recordset!grupclients) + "'", "codiclient=" + atrim(cadbl(datatarifes.Recordset!client))), False

End Sub
Function nomdelclient(vcodicli As Double) As String
   Dim rst As Recordset
   Set rst = dbtarifes.OpenRecordset("select nom from clients where codi=" + atrim(vcodicli))
   If Not rst.EOF Then
       nomdelclient = atrim(rst!nom)
   End If
   Set rst = Nothing
End Function

Private Sub etversio_DblClick()
   Dim v As String
   Dim vnumtarifaactual As String
   vnumtarifaactual = ctarifa
   v = InputBox("Escriu la versio que vols veure d'aquesta tarifa.", "Versió de la tarifa")
   If cadbl(v) > 0 Then
       carregar_tarifesclient " and numerotarifa=" + ctarifa + " and versio=" + v
         Else: carregar_tarifesclient
   End If
   If datatarifes.Recordset.EOF Then
     MsgBox "Aquesta versió no existeix", vbCritical, "Error"
     carregar_tarifesclient
   End If
   If ctarifa <> vnumtarifaactual Then datatarifes.Recordset.FindFirst "numerotarifa=" + vnumtarifaactual
End Sub

Private Sub Form_Load()
  fitxerini = "comandes.ini"
  
  cami = llegir_ini("General", "cami", fitxerini)
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  centerscreen Me
  Set dbtarifes = OpenDatabase(rutadelfitxer(cami) + "Tarifes.mdb", , False)
  datatarifes.DatabaseName = rutadelfitxer(cami) + "tarifes.mdb"
  Dataamples.DatabaseName = rutadelfitxer(cami) + "tarifes.mdb"
  Datacolors.DatabaseName = rutadelfitxer(cami) + "tarifes.mdb"
  datatarifes.RecordSource = "select * from tarifes_capcalera where idtarifa=0"
  datatarifes.Refresh
  Pestanyes.Tab = 0
  emplenar_combo_client
  emplenar_combo_unitatfacturacio
  etpesnet = ""
  
End Sub
Sub emplenar_combo_client()
    Dim rst As Recordset
   comboclient.Clear
   comboclient.AddItem "- Escullir client -"
   Set rst = dbtarifes.OpenRecordset("select distinct(grupdeclient) as grup from clients  ")
   While Not rst.EOF
      If atrim(rst!grup) <> "" Then comboclient.AddItem UCase(rst!grup)
      rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Sub emplenar_combo_unitatfacturacio()
    Dim rst As Recordset
   combounitatfacturacio.Clear
   combounitatescalat.Clear
   Set rst = dbtarifes.OpenRecordset("select unitatinterna from mesures order by unitatinterna")
   While Not rst.EOF
      combounitatfacturacio.AddItem UCase(rst!unitatinterna)
      combounitatescalat.AddItem UCase(rst!unitatinterna)
      rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Function descripciomesurabarem(v As String) As String
   If v = "REF" Then descripciomesurabarem = "Ref"
   descripciomesurabarem = v
End Function
Sub carregar_reixa_barem()
  Dim rst As Recordset
  Dim rstcol As Recordset
  Dim vrow As Double
  Dim vcol As Double
  Dim rstcond As Recordset
  vultimcolorcol = &HFDDECE
   Framecondicionant.visible = False
  reixa.Clear
  reixa.Cols = 0
  reixa.Rows = 0
  reixa.Redraw = False
  vrow = 0
  vcol = 0
  'PER CALCULAR L'AMPLADA DE LA REIXA APROX
   etcalcultamanyfont.FontName = reixa.Font.Name
   etcalcultamanyfont.FontSize = reixa.FontSize + 2
  etpesnet = ""
  etpesnet = IIf(datatarifes.Recordset!pesnet, "!!! PES NET !!!", "")
  Set rstcol = dbtarifes.OpenRecordset("select distinct VALOR1 from tarifes_barem where  linkvalor1condicionant=0 and idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa)))
  If rstcol.EOF Then GoTo fi
  rstcol.MoveLast: rstcol.MoveFirst
  
  reixa.Cols = 1
  reixa.Rows = 1
  reixa.col = 0
  reixa.Row = 0
  reixa.CellAlignment = 3
  reixa.TextMatrix(0, 0) = Mid(combounitatescalat, 3)
  While Not rstcol.EOF
        Set rst = dbtarifes.OpenRecordset("select * from tarifes_barem where idtarifa=" + atrim(datatarifes.Recordset!idtarifa) + " and valor1=" + atrim(rstcol!valor1) + " and linkvalor1condicionant=0 order by desde_kg")
        If Not rst.EOF Then rst.MoveLast: rst.MoveFirst
        emplenar_lafila_reixa rst, vrow, vcol
        vcol = vcol + 1
        If rstcol!valor1 > 0 Then
         Set rstcond = dbtarifes.OpenRecordset("select distinct valorcondicionant&valor1 as valorcondicionantmesvalor1 from tarifes_barem where idtarifa=" + atrim(datatarifes.Recordset!idtarifa) + " and linkvalor1condicionant=" + atrim(rstcol!valor1))
         If Not rst.EOF Then rst.MoveLast: rst.MoveFirst
         While Not rstcond.EOF
           Set rst = dbtarifes.OpenRecordset("select * from tarifes_barem where idtarifa=" + atrim(datatarifes.Recordset!idtarifa) + " and linkvalor1condicionant=" + atrim(rstcol!valor1) + " and valorcondicionant&trim(valor1)='" + atrim(rstcond!valorcondicionantmesvalor1) + "' order by desde_kg")
           emplenar_lafila_reixa rst, vrow, vcol
           rstcond.MoveNext
            vcol = vcol + 1
         Wend
        End If
        rstcol.MoveNext
  Wend
fi:
 ' reixa.Cols = reixa.Cols + 1
  If reixa.Cols > 0 Then
    reixa.Rows = reixa.Rows + 1
    reixa.Cols = reixa.Cols + 1
    If reixa.Rows > 1 Then reixa.FixedRows = 1
    If reixa.Cols > 1 Then
        reixa.FixedCols = 1
            Else
                reixa.Cols = 2
                reixa.FixedCols = 1
                reixa.Cols = 1
    End If
    reixa.RowHeight(reixa.Rows - 1) = 1
    reixa.ColWidth(reixa.Cols - 1) = 1
  End If
  Set rst = Nothing
  Set rstcol = Nothing
  comprovar_si_peces
  reixa.Redraw = True
End Sub
Function colorcondicionantmesmenys(vmes_menys As String) As Double
    If vmes_menys = "+" Then colorcondicionantmesmenys = &H25EFAD   'verd
    If vmes_menys = "-" Then colorcondicionantmesmenys = &H5C31DD  'vermell
    If vmes_menys = "F" Then colorcondicionantmesmenys = &HF8FDB5     'blau
End Function
Sub emplenar_lafila_reixa(rst As Recordset, vrow As Double, vcol As Double)
    If rst!valorcondicionant = "*Esc" Then
      reixa.Rows = 1
      While Not rst.EOF
        If rst.AbsolutePosition + 1 >= reixa.Rows Then reixa.Rows = reixa.Rows + 1
        reixa.TextMatrix(rst.AbsolutePosition + 1, 0) = Format(rst!desde_kg, "0")
        rst.MoveNext
      Wend
      GoTo fi
    End If
    If vcol >= reixa.Cols Then
       reixa.Cols = reixa.Cols + 2
       reixa.ColWidth(reixa.Cols - 2) = 700
       reixa.ColWidth(reixa.Cols - 1) = 1
       End If
     'If vcol = 1 Then reixa.Rows = rst.RecordCount + 2
     reixa.ColData(vcol) = rst!idliniabarem
     
     reixa.TextMatrix(0, vcol) = atrim(rst!valor1) + IIf(cadbl(rst!valor2) > 0, "-" + atrim(rst!valor2), "") + " " + descripciomesurabarem(rst!valorcondicionant)
     If rst!linkvalor1condicionant > 0 Then
        reixa.TextMatrix(0, vcol) = "<-Cond:" + atrim(rst!mes_menys) + " " + descripciomesurabarem(rst!valorcondicionant) + IIf(rst!valor1 > 0, atrim(rst!valor1), "") + IIf(cadbl(rst!valor2) > 0, "-" + atrim(rst!valor2), "")
        vcolorcol = colorcondicionantmesmenys(atrim(rst!mes_menys))
         Else
           vultimcolorcol = IIf(vultimcolorcol = &HFDDECE, &HED823A, &HFDDECE)
           vcolorcol = vultimcolorcol
     End If
     etcalcultamanyfont = reixa.TextMatrix(0, vcol) + "  "
     reixa.ColWidth(vcol) = etcalcultamanyfont.width
     
     vrow = 1
     While Not rst.EOF
        vpreu = 0
        If rst.AbsolutePosition + 1 >= reixa.Rows Then reixa.Rows = reixa.Rows + 1
        If (vcol) = 1 Then reixa.TextMatrix(rst.AbsolutePosition + 1, 0) = Format(rst!desde_kg, "0")
        If reixa.TextMatrix(vrow, 0) = Format(rst!desde_kg, "0") Then vpreu = rst!preu_kg
        If vpreu > 0 Then reixa.TextMatrix(vrow, vcol) = Format(vpreu, "0.00€")
        If reixa.TextMatrix(vrow, 0) = Format(rst!desde_kg, "0") Then
              rst.MoveNext
            Else
              If vrow + 1 >= reixa.Rows Then rst.MoveNext: vrow = 0
        End If
        'poso el color de la casella i si no hi ha preu en NEGRE
          reixa.col = vcol
          reixa.Row = vrow
          reixa.CellBackColor = vcolorcol 'IIf(vpreu = 0, &H80000008, vcolorcol)
        vrow = vrow + 1
     Wend
     While vrow < reixa.Rows
          reixa.col = vcol
          reixa.Row = vrow
          reixa.CellBackColor = vcolorcol
          vrow = vrow + 1
     Wend
fi:
End Sub
Sub carregar_reixa_ports()
  Dim rst As Recordset
  Dim rstcol As Recordset
  Dim rstt As Recordset
  Dim vsql As String
  Dim vrow As Double
  Dim vkgant As Double
  Dim vampladaminimacol As Long
  
  'reixaports.Redraw = False
  reixaports.Clear
  reixaports.Rows = 2
  reixaports.col = 0
  reixaports.Row = 0
  reixaports.Cols = 1
  reixaports.ColWidth(0) = 1200
  reixaports.CellAlignment = 3
  reixaports.TextMatrix(0, 0) = "KG"
  DoEvents
  etcalcultamanyfont.FontName = reixaports.Font.Name
  etcalcultamanyfont.FontSize = reixaports.FontSize + 2
  etcalcultamanyfont = "OOOOOO"
  vampladaminimacol = etcalcultamanyfont.width
  vsql = IIf(cadbl(comboclient.tag) > 0, " codiclient=" + atrim(cadbl(comboclient.tag)), " grupclients='" + atrim(comboclient) + "'")
  Set rstcol = dbtarifes.OpenRecordset("select nomdelatarifa as Tnomdelatarifa, datafi as Tdatafi from tarifes_ports_capcalera where " + vsql)
  If rstcol.EOF Then Exit Sub
  rstcol.MoveLast: rstcol.MoveFirst
  While Not rstcol.EOF
        Set rstt = dbtarifes.OpenRecordset("select * from tarifes_ports_capcalera where nomdelatarifa='" + atrim(rstcol!Tnomdelatarifa) + "' and datafi=#" + atrim(rstcol!Tdatafi) + "# and " + vsql)
        If rstt.EOF Then GoTo proxim
        Set rst = dbtarifes.OpenRecordset("select * from tarifes_ports where idtarifaports=" + atrim(rstt!idtarifaports) + " order by desde_kg")
        If Not rst.EOF Then rst.MoveLast: rst.MoveFirst
        If rstcol.AbsolutePosition + 1 >= reixaports.Cols Then
          reixaports.Cols = reixaports.Cols + 1
          reixaports.col = reixaports.Cols - 1
          reixaports.ColWidth(reixaports.Cols - 1) = 700
          reixaports.ColData(reixaports.Cols - 1) = rstt!idtarifaports
          reixaports.Row = 0
          reixaports.CellAlignment = 3
          End If
        If (rstcol.AbsolutePosition + 1) = 1 Then reixaports.Rows = IIf(rst.RecordCount + 1 < 3, 3, rst.RecordCount + 1)
        reixaports.TextMatrix(0, rstcol.AbsolutePosition + 1) = atrim(rstt!nomdelatarifa)
        etcalcultamanyfont = rstt!nomdelatarifa
        reixaports.ColWidth(rstcol.AbsolutePosition + 1) = IIf(etcalcultamanyfont.width < vampladaminimacol, vampladaminimacol, etcalcultamanyfont.width)
        reixaports.col = rstcol.AbsolutePosition + 1: reixaports.Row = 1: reixaports.CellBackColor = QBColor(14)
        reixaports.TextMatrix(1, rstcol.AbsolutePosition + 1) = Format(rstt!datainici, "dd/mm/yy")
        reixaports.col = rstcol.AbsolutePosition + 1: reixaports.Row = 2: reixaports.CellBackColor = QBColor(14)
        reixaports.TextMatrix(2, rstcol.AbsolutePosition + 1) = Format(rstt!datafi, "dd/mm/yy")
        
        vrow = 3
        While Not rst.EOF
           vpreu = 0
           If vrow >= reixaports.Rows Then reixaports.Rows = reixaports.Rows + 1
           If (rstcol.AbsolutePosition + 1) = 1 Then
              'reixaports.TextMatrix(rst.AbsolutePosition + 1, 0) = atrim(vkgant) + " - " + Format(rst!desde_kg, "0") + "Kg"
              reixaports.TextMatrix(vrow, 0) = Format(rst!desde_kg, "0")
              vkgant = cadbl(rst!desde_kg)
              reixaports.TextMatrix(1, 0) = "Vàlid_De"
              reixaports.TextMatrix(2, 0) = "Vàlid_A"
           End If
           If reixaports.TextMatrix(vrow, 0) = Format(rst!desde_kg, "0") Then vpreu = rst!preu_kg
           If vpreu > 0 Then reixaports.TextMatrix(vrow, rstcol.AbsolutePosition + 1) = Format(vpreu, "0.00€")
           If reixaports.TextMatrix(vrow, 0) = Format(rst!desde_kg, "0") Then
                 rst.MoveNext
               Else
                 If vrow + 1 >= reixaports.Rows Then rst.MoveNext: vrow = 0
           End If
           If vpreu = 0 And vrow > 0 Then
              reixaports.col = rstcol.AbsolutePosition + 1
              reixaports.Row = vrow
              'reixaports.CellBackColor = &H80000008
           End If
           vrow = vrow + 1
        Wend
        'While vrow < reixaports.Rows
        '     reixaports.col = rstcol.AbsolutePosition + 1
        '     reixaports.Row = vrow
        '     reixaports.CellBackColor = &H80000008
        '     vrow = vrow + 1
        'Wend
proxim:
        rstcol.MoveNext
  Wend
  If reixaports.Cols = 2 Then
     reixaports.Cols = 3
     reixaports.ColWidth(2) = 0
  End If
  Set rst = Nothing
  Set rstcol = Nothing
  reixaports.Redraw = True
End Sub

Private Sub mat1m_a_GotFocus()
  If cadbl(mat1m_a) = 0 Then mat1m_a = mat1m_de
  seleccionartotelcontrol
End Sub

Private Sub mat1m_de_LostFocus()
If cadbl(mat1m_a) = 0 Then mat1m_a = mat1m_de
End Sub

Private Sub mat2m_a_GotFocus()
   If cadbl(mat2m_a) = 0 Then mat2m_a = mat2m_de
   seleccionartotelcontrol
End Sub
Sub seleccionartotelcontrol()
   Screen.ActiveControl.SelStart = 0
   Screen.ActiveControl.SelLength = Len(Screen.ActiveControl)
End Sub

Private Sub mat2m_de_LostFocus()
If cadbl(mat2m_a) = 0 Then mat2m_a = mat2m_de
End Sub

Private Sub mat3m_a_GotFocus()
   If cadbl(mat3m_a) = 0 Then mat3m_a = mat3m_de
   seleccionartotelcontrol
End Sub

Private Sub mat3m_de_LostFocus()
   If cadbl(mat3m_a) = 0 Then mat3m_a = mat3m_de
End Sub

Private Sub modificar_Click()
   If datatarifes.Recordset.EOF Then Exit Sub
   If datatarifes.Recordset.EditMode > 0 Then MsgBox "Ja s'està editant una tarifa.", vbCritical, "Error": Exit Sub
   fcapcalera1.Enabled = True
   fcapcalera2.Enabled = True
   Frameescalat.Enabled = True
   datatarifes.Recordset.Edit
   etestat = "Editant..."
  ' Pestanyes.Tab = 0
End Sub

Private Sub Pestanyes_Click(PreviousTab As Integer)
  If comboclient = "" Then Exit Sub
  If Pestanyes.caption = "Condicionants" Then carregar_condicionants
  If Pestanyes.caption = "Transports" Then possar_dades_ports
  If Pestanyes.caption = "Escalat de preus" Then possar_dades_tarifes
End Sub
Sub possar_dades_ports()
   
   'If datatarifes.Recordset.EOF Then
   '  reixaports.Clear
   '  Exit Sub
   'End If
   carregar_reixa_destins
   carregar_reixa_ports
End Sub
Sub carregar_reixa_destins()
   Dim vsql As String
   Dim rst As Recordset
   datadestins.DatabaseName = rutadelfitxer(cami) + "tarifes.mdb"
   datadestins.RecordSource = "SELECT * from tarifes_destins "
   vsql = IIf(cadbl(comboclient.tag) > 0, " codiclient=" + atrim(cadbl(comboclient.tag)), " grupclients='" + atrim(comboclient) + "'")
   datadestins.RecordSource = datadestins.RecordSource + " where " + vsql + " order by nomdeldesti"
   datadestins.Refresh
   vsql = IIf(cadbl(comboclient.tag) > 0, " clients.codi=" + atrim(cadbl(comboclient.tag)), " grupdeclient='" + atrim(comboclient) + "'")
   Set rst = dbtarifes.OpenRecordset("SELECT DISTINCT Clients_envios.provinciae FROM clients INNER JOIN Clients_envios ON clients.codi = Clients_envios.codi WHERE " + vsql)
   While Not rst.EOF
     datadestins.Recordset.FindFirst "nomdeldesti='" + atrim(rst!provinciae) + "'"
     If datadestins.Recordset.NoMatch Then dbtarifes.Execute "insert into tarifes_destins (" + IIf(cadbl(comboclient.tag) > 0, "codiclient", "grupclients") + ",nomdeldesti) values (" + IIf(cadbl(comboclient.tag) > 0, atrim(cadbl(comboclient.tag)), "'" + atrim(comboclient) + "'") + ",'" + atrim(rst!provinciae) + "')"
     rst.MoveNext
   Wend
   datadestins.Refresh
   Set rst = Nothing
End Sub
Sub carregar_condicionants()
  Dim rst As Recordset
  Dim vfila As Integer
  reixacond.Clear
  If datatarifes.Recordset.EOF Then Exit Sub
  reixacond.Rows = 2
  reixacond.Cols = 3
  reixacond.FixedCols = 0
  reixacond.FixedRows = 1
  reixacond.TextMatrix(0, 0) = "Condició"
  reixacond.TextMatrix(0, 1) = "€/Kg"
  reixacond.ColWidth(0) = 10000
  reixacond.ColWidth(1) = 1000
  reixacond.ColWidth(2) = 0
  vfila = 1
  Set rst = dbtarifes.OpenRecordset("Select * from tarifes_condicionants where idtarifa=" + atrim(datatarifes.Recordset!idtarifa))
  While Not rst.EOF
    reixacond.Rows = vfila + 1
    reixacond.TextMatrix(vfila, 1) = Format(rst!valoreuros, "0.00") + "€"
    reixacond.TextMatrix(vfila, 0) = generar_condicio(rst)
    reixacond.TextMatrix(vfila, 2) = rst!ID
    vfila = vfila + 1
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
Function generar_condicio(rst As Recordset) As String
   Dim vcond As String
   Dim vcodi As String
   vcond = UCase(rst!condicionant)
   vcodi = Mid(vcond + "   ", 1, 4)
   generar_condicio = IIf(rst!mes_menys = "A", "Augment de preu per ", IIf(rst!mes_menys = "R", "Reducció de preu per ", IIf(rst!mes_menys = "C", "", "")))
   generar_condicio = generar_condicio + Mid(vcond, 5)
   If vcodi = "[NC]" Then generar_condicio = generar_condicio + " " + atrim(rst!valorc1) + "/" + atrim(rst!valorc2)
   If vcodi = "[EM]" Then generar_condicio = generar_condicio + " " + atrim(rst!valorc1) + " µ"
   If vcodi = "[MK]" Then generar_condicio = generar_condicio + "->" + atrim(rst!valortexte)
   
End Function
Sub posar_valor_casella(Optional v As String)
   Dim vkg As Double
   Dim rstcol As Recordset
   Dim rst As Recordset
   
   If v = "" Then v = InputBox("Escriu el preu per aquesta casella.", "Valor casella")
   If StrPtr(v) = 0 Then Exit Sub
   vkg = cadbl(reixa.TextMatrix(reixa.Row, 0))
   Set rstcol = dbtarifes.OpenRecordset("select * from tarifes_barem where idliniabarem=" + atrim(reixa.ColData(reixa.col)))
   If rstcol.EOF Then Exit Sub
   'Clipboard.Clear
   'Clipboard.SetText "select * from tarifes_barem where desde_kg=" + atrim(vkg) + " and  idtarifa=" + atrim(rstcol!idtarifa) + " and valorcondicionant='" + atrim(rstcol!valorcondicionant) + "' and linkvalor1condicionant=" + atrim(rstcol!linkvalor1condicionant) + " and linkvalorcondicionant='" + atrim(rstcol!linkvalorcondicionant) + "'"
   
   Set rst = dbtarifes.OpenRecordset("select * from tarifes_barem where valor1=" + atrim(rstcol!valor1) + " and desde_kg=" + atrim(vkg) + " and  idtarifa=" + atrim(rstcol!idtarifa) + " and valorcondicionant='" + atrim(rstcol!valorcondicionant) + "' and linkvalor1condicionant=" + atrim(rstcol!linkvalor1condicionant) + " and linkvalorcondicionant=" + IIf(IsNull(rstcol!linkvalorcondicionant), "null", "'" + atrim(rstcol!linkvalorcondicionant) + "'"))
   If rst.EOF Then
        'fer un registre nou
         rst.AddNew
         rst!idtarifa = rstcol!idtarifa
         rst!valor1 = rstcol!valor1
         rst!valor2 = rstcol!valor2
         rst!desde_kg = vkg
         rst!valorcondicionant = rstcol!valorcondicionant
         rst!linkvalorcondicionant = rstcol!linkvalorcondicionant
         rst!linkvalor1condicionant = rstcol!linkvalor1condicionant
         rst!preu_kg = cadbl(v)
         rst.Update
        Else
          rst.Edit
          rst!preu_kg = cadbl(v)
          rst.Update
   End If
End Sub
Private Sub reixa_DblClick()
   Dim rst As Recordset
   Dim rstcol As Recordset
   Dim vkg As String
   Dim vref As String
   Dim vvalor2 As String
   Dim vcolref As Double
   Dim vguardarcol As Double
   vguardarcol = reixa.col
   If reixa.col = 0 And (reixa.Row = 1 And reixa.RowSel = reixa.Rows - 1) And reixa.Cols = 1 Then
      If MsgBox("Borrar tot l'escalat?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
          dbtarifes.Execute "delete * from  tarifes_barem  where idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa))
      End If
      GoTo fi
   End If
   'posar el valor d'una casella
   If reixa.TextMatrix(reixa.Row, 0) = "" Then Exit Sub
   If reixa.RowSel = reixa.Row And reixa.ColSel = reixa.col Then
       If reixa.col > 0 Then
           posar_valor_casella
       End If
   End If
   If (reixa.ColSel = reixa.Cols - 1 And reixa.col = 1) And (reixa.RowSel = reixa.Rows - 1 And reixa.Row = 1) Then
      If MsgBox("Vols eliminar tots els valors de la reixa?", vbDefaultButton2 + vbYesNo + vbCritical, "Atenció") = vbYes Then
          dbtarifes.Execute "delete* from tarifes_barem where idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa))
      End If
      GoTo fi
   End If
   'si vol canviar els Kg d'una fila
   If reixa.ColSel = reixa.Cols - 1 And reixa.col = 1 Then
       vkg = InputBox("Entra el nou valor de la FILA." + vbNewLine + "SI VOLS ELIMINAR-LA ESCRIU UN ZERO.", combounitatescalat)
       If StrPtr(vkg) = 0 Then Exit Sub
       If cadbl(vkg) > 0 Then
         Set rst = dbtarifes.OpenRecordset("select * from tarifes_barem where idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa)) + " order by valor1")
         rst.FindFirst "desde_kg=" + atrim(vkg)
         If Not rst.NoMatch Then MsgBox "Aquest valor ja està entrat.", vbCritical, "Error": GoTo fi
         dbtarifes.Execute "update tarifes_barem set desde_kg=" + atrim(vkg) + " where idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa)) + " and desde_kg=" + atrim(reixa.TextMatrix(reixa.Row, 0))
           Else
             If MsgBox("Segur que vols ELIMINAR aquesta fila de Kg de la tarifa?", vbCritical + vbDefaultButton2 + vbYesNo, "Eliminar") = vbYes Then
               dbtarifes.Execute "delete * from tarifes_barem where desde_kg=" + atrim(reixa.TextMatrix(reixa.Row, 0))
             End If
       End If
   End If
   
   'si vols canviar el valor de la columna
   If reixa.RowSel = reixa.Rows - 1 And reixa.Row = 1 Then
       Set rstcol = dbtarifes.OpenRecordset("select * from tarifes_barem where idliniabarem=" + atrim(reixa.ColData(reixa.col)))
       If rstcol.EOF Then Exit Sub
       vref = InputBox("Entra el nou valor de la columna. DE " + vbNewLine + "SI VOLS ELIMINAR-LA ESCRIU UN ZERO.", "Columna")
       If StrPtr(vref) = 0 Then Exit Sub
       If cadbl(rstcol!valor1) = 0 And vref <> "0" Then vref = -1
       If cadbl(rstcol!valor2) > 0 And cadbl(vref) > 0 Then
           vvalor2 = InputBox("Entra el segon valor deliminador A " + rstcol!valorcondicionant + vbNewLine + "SI VOLS ELIMINAR-LA ESCRIU UN ZERO.", "Columna")
           If cadbl(vvalor2) = 0 Then Exit Sub
       End If
       If cadbl(vref) > 0 Then
         Set rst = dbtarifes.OpenRecordset("select * from tarifes_barem where idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa)) + " and valorcondicionant='" + atrim(rstcol!valorcondicionant) + "' and linkvalor1condicionant=" + atrim(rstcol!valor1) + " order by desde_kg")
         If vref <> rstcol!valor1 Then
            rst.FindFirst "valor1=" + atrim(vref)
            If Not rst.NoMatch Then MsgBox "Aquest valor de columna ja està entrada.", vbCritical, "Error": GoTo fi
         End If
         If rstcol!linkvalor1condicionant > 0 Then
            'dbtarifes.Execute "update tarifes_barem set linkvalor1condicionant=" + atrim(vref) + " where idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa)) + " and linkvalor1condicionant=" + atrim(rstcol!linkvalor1condicionant) + " and linkvalorcondicionant='" + atrim(rstcol!linkvalorcondicionant) + "'"
            dbtarifes.Execute "update tarifes_barem set valor1=" + atrim(vref) + ",valor2=" + atrim(cadbl(vvalor2)) + " where idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa)) + " and valor1=" + atrim(rstcol!valor1) + " and linkvalor1condicionant=" + atrim(rstcol!linkvalor1condicionant) + " and linkvalorcondicionant='" + atrim(rstcol!linkvalorcondicionant) + "'"
             Else
               dbtarifes.Execute "update tarifes_barem set linkvalor1condicionant=" + atrim(vref) + " where idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa)) + " and linkvalor1condicionant=" + atrim(rstcol!valor1) + " and linkvalorcondicionant='" + atrim(rstcol!valorcondicionant) + "'"
               dbtarifes.Execute "update tarifes_barem set valor1=" + atrim(vref) + ",valor2=" + atrim(cadbl(vvalor2)) + " where idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa)) + " and valor1=" + atrim(rstcol!valor1) + " and valorcondicionant='" + atrim(rstcol!valorcondicionant) + "'"
               
         End If
           Else
             If vref <> "-1" Then
                If MsgBox("Segur que vols ELIMINAR aquesta columna de la tarifa?", vbCritical + vbDefaultButton2 + vbYesNo, "Eliminar") = vbYes Then
                    dbtarifes.Execute "delete * from tarifes_barem where idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa)) + " and linkvalor1condicionant=" + atrim(rstcol!valor1) + " and linkvalorcondicionant='" + atrim(rstcol!valorcondicionant) + "'"
                    dbtarifes.Execute "delete * from tarifes_barem where idtarifa=" + atrim(cadbl(datatarifes.Recordset!idtarifa)) + " and  valor1=" + atrim(rstcol!valor1) + " and linkvalor1condicionant=" + atrim(rstcol!linkvalor1condicionant) + " and  valorcondicionant='" + atrim(rstcol!valorcondicionant + "'")
                End If
             End If
       End If
   End If
fi:
   Set rst = Nothing
   Set rstcol = Nothing
   carregar_reixa_barem
   If reixa.Cols > vguardarcol Then reixa.col = vguardarcol
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub reixa_SelChange()
  If Mid(reixa.TextMatrix(0, reixa.col) + "      ", 1, 2) <> "<-" And Trim(reixa.TextMatrix(0, reixa.col)) <> "" And reixa.col > 0 Then
      Framecondicionant.visible = True
      Framecondicionant.Left = reixa.Left + reixa.CellLeft
        Else: Framecondicionant.visible = False
  End If
End Sub

Private Sub reixaamples_BeforeUpdate(Cancel As Integer)
   Dataamples.Recordset!idtarifa = datatarifes.Recordset!idtarifa
   Dataamples.Recordset!tipus = "A"
End Sub

Private Sub reixacolors_BeforeUpdate(Cancel As Integer)
   Datacolors.Recordset!idtarifa = datatarifes.Recordset!idtarifa
   Datacolors.Recordset!tipus = "C"
End Sub

Private Sub reixacond_GotFocus()
  beliminarcond.visible = True
End Sub

Private Sub reixacond_LostFocus()
  If Screen.ActiveControl.Name <> "beliminarcond" Then beliminarcond.visible = False
End Sub

Private Sub reixadestins_ButtonClick(ByVal ColIndex As Integer)
  If ColIndex = 1 Then
      escullir_tarifaportsperaquestdesti
  End If
End Sub
Sub escullir_tarifaportsperaquestdesti()
  Dim vsql As String
  Dim vnomtarifa As String
  Dim vidtarifa As Long
  vsql = IIf(cadbl(comboclient.tag) > 0, " codiclient=" + atrim(cadbl(comboclient.tag)), " grupclients='" + atrim(comboclient) + "'")
  vsql = "select nomdelatarifa as Tnomdelatarifa, datafi as Tdatafi,idtarifaports as idtarifap from tarifes_ports_capcalera where " + vsql
  
  Load formseleccio
  formseleccio.caption = "Escull el nom de la tarifa"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "tarifes.mdb"
  formseleccio.Data1.RecordSource = vsql
  
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 2000
  formseleccio.DBGrid2.Columns(1).width = 1500
  formseleccio.DBGrid2.Columns(2).visible = False

  
  formseleccio.Show 1
  If seleccioret = 1 Then
   vnomtarifa = atrim(formseleccio.Data1.Recordset!Tnomdelatarifa)
   vidtarifa = cadbl(formseleccio.Data1.Recordset!idtarifap)
   'dbtarifes.Execute "update tarifes_destins set idtarifa=" + atrim(vidtarifa) + ",nomdelatarifa='" + treure_apostruf(vnomtarifa) + "' where id=" + atrim(datatarifesdestins.Recordset!id)
   
   datadestins.Recordset.Edit
   datadestins.Recordset!idtarifa = vidtarifa
   datadestins.Recordset!nomdelatarifa = vnomtarifa
   datadestins.Recordset.Update
   
  End If
  Unload formseleccio
    
End Sub

Sub posar_valor_casella_ports()
   Dim rst As Recordset
   Dim v As String
   v = InputBox("Escriu el valor que vols per aquesta casella.", "Preu")
   If StrPtr(v) = 0 Then Exit Sub
   Set rst = dbtarifes.OpenRecordset("select * from tarifes_ports where idtarifaports=" + atrim(reixaports.ColData(reixaports.col)) + " and desde_kg=" + atrim(reixaports.TextMatrix(reixaports.Row, 0)))
   If rst.EOF Then
        'fer un registre nou
         rst.AddNew
         rst!idtarifaports = reixaports.ColData(reixaports.col)
         rst!desde_kg = atrim(reixaports.TextMatrix(reixaports.Row, 0))
         rst!preu_kg = cadbl(v)
         rst.Update
        Else
          rst.Edit
          rst!preu_kg = cadbl(v)
          rst.Update
   End If
   reixaports.Text = Format(v, "0.00€")
   Set rst = Nothing
End Sub
Private Sub reixaports_DblClick()
   Dim vnomtarifa As String
   Dim vkg As String
   Dim vsql As String
   Dim rst As Recordset
   
   
     'si vols canviar el valor de la columna
   If reixaports.RowSel = reixaports.Rows - 1 And reixaports.Row = 1 Then
        vsql = IIf(cadbl(comboclient.tag) > 0, " codiCLIENT=" + atrim(cadbl(comboclient.tag)), " grupclients='" + atrim(comboclient) + "'")
        vnomtarifa = UCase(atrim(InputBox("Escriu el nom que vols possar a aquesta tarifa." + vbNewLine + "ESCRIU [ELIMINAR] PER ELIMINAR AQUESTA TARIFA.", "Nom de la tarifa")))
        If vnomtarifa = "" Then Exit Sub
        If vnomtarifa = "ELIMINAR" Then
          If MsgBox("Segur que vols eliminar aquesta tarifa?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
               dbtarifes.Execute "delete * from tarifes_ports_capcalera where idtarifaports=" + atrim(reixaports.ColData(reixaports.col))
               dbtarifes.Execute "delete * from tarifes_ports where idtarifaports=" + atrim(reixaports.ColData(reixaports.col))
               dbtarifes.Execute "update tarifes_destins set nomdelatarifa=null,idtarifa=0 where idtarifa=" + atrim(reixaports.ColData(reixaports.col))
               possar_dades_ports
          End If
          Exit Sub
        End If
        If Len(vnomtarifa) > 20 Then MsgBox "El nom no pot tenir mes de 20 caràcters.", vbCritical, "Error": Exit Sub
        vnomtarifa = treure_apostruf(vnomtarifa)
        Set rst = dbtarifes.OpenRecordset("select * from tarifes_ports_capcalera where ucase(nomdelatarifa)='" + vnomtarifa + "' and " + vsql)
        If Not rst.EOF Then MsgBox "Aquest nom de tarifa ja existeix", vbCritical, "Error": Exit Sub
        Set rst = dbtarifes.OpenRecordset("select * from tarifes_ports_capcalera where idtarifaports=" + atrim(reixaports.ColData(reixaports.col)))
        If Not rst.EOF Then
            rst.Edit
            rst!nomdelatarifa = vnomtarifa
            rst.Update
              'canvio el nom dels destins també
            dbtarifes.Execute "update tarifes_destins set nomdelatarifa='" + vnomtarifa + "' where idtarifa=" + atrim(reixaports.ColData(reixaports.col))
            possar_dades_ports
        End If
        Exit Sub
   End If
   
   'si vol canviar el valor de la capçalera d'una fila
   If reixaports.ColSel = reixaports.Cols - 1 And reixaports.col = 1 Then
       vkg = InputBox("Entra el nou valor de la fila." + vbNewLine + "SI VOLS ELIMINAR-LA ESCRIU UN ZERO.", combounitatescalat)
       If StrPtr(vkg) = 0 Then Exit Sub
       vsql = IIf(cadbl(comboclient.tag) > 0, " codiclient=" + atrim(cadbl(comboclient.tag)), " grupclients='" + atrim(comboclient) + "'")
       If cadbl(vkg) > 0 Then
           dbtarifes.Execute "update Tarifes_ports INNER JOIN Tarifes_ports_capcalera ON Tarifes_ports.idtarifaports = Tarifes_ports_capcalera.idtarifaports set desde_kg=" + atrim(vkg) + " where desde_kg=" + reixaports.TextMatrix(reixaports.Row, 0) + " and " + vsql
            Else
              Set rst = dbtarifes.OpenRecordset("select * from Tarifes_ports INNER JOIN Tarifes_ports_capcalera ON Tarifes_ports.idtarifaports = Tarifes_ports_capcalera.idtarifaports where desde_kg=" + reixaports.TextMatrix(reixaports.Row, 0) + " and " + vsql)
              While Not rst.EOF
                rst.Delete
                rst.MoveNext
              Wend
       End If
       possar_dades_ports
       Exit Sub
   End If
   
    'posar el valor d'una casella
   If reixaports.ColSel = reixaports.col And reixaports.RowSel = reixaports.Row Then
       If reixaports.Row > 2 Then
           posar_valor_casella_ports
           Exit Sub
       End If
   End If
End Sub

Private Sub reixarappels_BeforeUpdate(Cancel As Integer)
   
End Sub

Private Sub Timer_escullirclient_Timer()
    Timer_escullirclient.Enabled = False
    escullir_client
    carregar_tarifesclient
    
End Sub

Private Sub Timercadasegon_Timer()
   etpesnet.visible = IIf(Second(Now) Mod 2 = 0, True, False)
End Sub
