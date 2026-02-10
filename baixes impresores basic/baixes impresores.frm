VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form form1 
   BackColor       =   &H80000005&
   Caption         =   "Baixes Comandes (Impressores)"
   ClientHeight    =   11430
   ClientLeft      =   4965
   ClientTop       =   2115
   ClientWidth     =   11700
   ClipControls    =   0   'False
   Icon            =   "baixes impresores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11430
   ScaleWidth      =   11700
   Begin VB.Frame framebobentrada 
      Caption         =   "Bobines Entrada"
      Height          =   3885
      Left            =   6945
      TabIndex        =   73
      Top             =   3990
      Visible         =   0   'False
      Width           =   3435
      Begin MSDBGrid.DBGrid bobentrada 
         Bindings        =   "baixes impresores.frx":058A
         Height          =   3075
         Left            =   60
         OleObjectBlob   =   "baixes impresores.frx":059F
         TabIndex        =   74
         Top             =   210
         Width           =   3315
      End
      Begin VB.CommandButton bdesb2 
         BackColor       =   &H00F1B75F&
         Height          =   480
         Left            =   780
         Picture         =   "baixes impresores.frx":0F85
         Style           =   1  'Graphical
         TabIndex        =   139
         ToolTipText     =   "Bobina desbobinador2"
         Top             =   2790
         Width           =   645
      End
      Begin VB.CommandButton bdesb1 
         BackColor       =   &H0017D062&
         Height          =   480
         Left            =   105
         Picture         =   "baixes impresores.frx":14D5
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Bobina desbobinador1"
         Top             =   2790
         Width           =   645
      End
      Begin VB.CommandButton Command34 
         Height          =   330
         Left            =   2970
         Picture         =   "baixes impresores.frx":1A29
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Ubicació d'una bobina a magatzem."
         Top             =   3510
         Width           =   420
      End
      Begin VB.CommandButton veuregrupsdestoc 
         Height          =   480
         Left            =   1410
         Picture         =   "baixes impresores.frx":1FB3
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Llista de bobines del grup"
         Top             =   3315
         Width           =   645
      End
      Begin VB.CommandButton eliminarbobentrada 
         Height          =   480
         Left            =   2055
         Picture         =   "baixes impresores.frx":253D
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Eliminar bobina d'entrada"
         Top             =   3315
         Width           =   645
      End
      Begin VB.CommandButton Command20 
         Height          =   480
         Left            =   765
         Picture         =   "baixes impresores.frx":2AC7
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Afegir manualment el Palet/Bobina d'entrada"
         Top             =   3315
         Width           =   645
      End
      Begin VB.CheckBox veuretotesbobent 
         Caption         =   "Totes"
         Height          =   255
         Left            =   2700
         TabIndex        =   99
         Top             =   3300
         Width           =   705
      End
      Begin VB.CommandButton botoensenyarpacking 
         Height          =   480
         Left            =   105
         Picture         =   "baixes impresores.frx":3051
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Sel.lecciona la bobina del Packinglist"
         Top             =   3315
         Width           =   645
      End
      Begin VB.CommandButton Command37 
         Height          =   480
         Left            =   2700
         Picture         =   "baixes impresores.frx":35DB
         Style           =   1  'Graphical
         TabIndex        =   140
         ToolTipText     =   "Modificar Bobines desbobinadors"
         Top             =   2790
         Width           =   645
      End
      Begin VB.Label etmaterialexacte 
         BackStyle       =   0  'Transparent
         Caption         =   "     "
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   90
         TabIndex        =   123
         Top             =   3090
         Width           =   3255
      End
   End
   Begin VB.Frame frameempalmes 
      Caption         =   "Senyals"
      Height          =   3795
      Left            =   6975
      TabIndex        =   70
      Top             =   3975
      Visible         =   0   'False
      Width           =   3435
      Begin MSDBGrid.DBGrid reixaempalmes 
         Bindings        =   "baixes impresores.frx":3BAB
         Height          =   3570
         Left            =   60
         OleObjectBlob   =   "baixes impresores.frx":3BBE
         TabIndex        =   71
         Top             =   165
         Width           =   3330
      End
   End
   Begin VB.CommandButton Command38 
      Height          =   360
      Left            =   1680
      Picture         =   "baixes impresores.frx":45B7
      Style           =   1  'Graphical
      TabIndex        =   142
      ToolTipText     =   "Llista de comandes pendents."
      Top             =   180
      Width           =   375
   End
   Begin VB.TextBox comanda 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005C31DD&
      Height          =   360
      Left            =   270
      TabIndex        =   141
      Top             =   180
      Width           =   1410
   End
   Begin VB.CommandButton Command36 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10410
      Picture         =   "baixes impresores.frx":4B41
      Style           =   1  'Graphical
      TabIndex        =   137
      ToolTipText     =   "Bobines portades a Impresores"
      Top             =   5385
      Width           =   1245
   End
   Begin VB.TextBox cobservacionsoperari 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   5850
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   134
      Top             =   7860
      Width           =   6150
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   1665
      Top             =   2370
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Avaria"
      Enabled         =   0   'False
      Height          =   570
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox cpostitcomanda 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   285
      Locked          =   -1  'True
      MouseIcon       =   "baixes impresores.frx":52CE
      TabIndex        =   114
      Text            =   "Per entrar la comanda prem la fletxa i escull la comanda."
      Top             =   -180
      Visible         =   0   'False
      Width           =   4380
   End
   Begin VB.CommandButton botodescansrelleu 
      Height          =   390
      Left            =   10215
      Picture         =   "baixes impresores.frx":5858
      Style           =   1  'Graphical
      TabIndex        =   113
      ToolTipText     =   "Control Descans i Relleu"
      Top             =   900
      Width           =   1365
   End
   Begin VB.CommandButton Command13 
      Height          =   690
      Left            =   10410
      Picture         =   "baixes impresores.frx":5DE2
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   6690
      Width           =   1260
   End
   Begin VB.ComboBox combocomanda 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   660
      TabIndex        =   109
      Top             =   180
      Width           =   1425
   End
   Begin VB.CommandButton bobsajust 
      BackColor       =   &H0080FF80&
      Caption         =   "Bobs Ajust"
      Height          =   435
      Left            =   6930
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   870
      Visible         =   0   'False
      Width           =   1155
   End
   Begin Crystal.CrystalReport llistatbob 
      Left            =   -120
      Top             =   1665
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame panellavis 
      BackColor       =   &H00C0C0FF&
      Height          =   7185
      Left            =   10545
      TabIndex        =   95
      Top             =   11295
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CommandButton Command19 
         BackColor       =   &H000000FF&
         Caption         =   "D'acord"
         Height          =   825
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   5670
         Width           =   2295
      End
      Begin VB.Label missatgeavis 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   26.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1860
         Left            =   225
         TabIndex        =   96
         Top             =   1920
         Width           =   9360
      End
   End
   Begin VB.CommandButton imprimir 
      BackColor       =   &H00FF8080&
      Caption         =   "Imprimir"
      Height          =   330
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   465
      Width           =   765
   End
   Begin VB.CommandButton maquina 
      BackColor       =   &H00FF8080&
      Caption         =   "Maq: 0"
      Height          =   390
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   60
      Width           =   765
   End
   Begin VB.CommandButton Command9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   10410
      Picture         =   "baixes impresores.frx":62D1
      Style           =   1  'Graphical
      TabIndex        =   89
      ToolTipText     =   "Ensenya Pantones utilitzats (Apretat x modificar)"
      Top             =   5385
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Data lots 
      Caption         =   "dblots"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   10500
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "lots"
      Top             =   3555
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H008080FF&
      Caption         =   "No Acabada"
      Height          =   660
      Left            =   10755
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   30
      Width           =   915
   End
   Begin VB.CommandButton command15 
      BackColor       =   &H0080FF80&
      Caption         =   "Acabar Comanda"
      Height          =   645
      Left            =   9870
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   45
      Width           =   885
   End
   Begin VB.Frame calculant 
      Height          =   2580
      Left            =   7155
      TabIndex        =   76
      Top             =   11475
      Visible         =   0   'False
      Width           =   5730
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calculant..."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1170
         TabIndex        =   77
         Top             =   990
         Width           =   3795
      End
   End
   Begin VB.CommandButton Command14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6480
      Picture         =   "baixes impresores.frx":7DCB
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   870
      Width           =   435
   End
   Begin VB.Data bobinesent 
      Caption         =   "bobinesentimp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "bobinesentimp"
      Top             =   6975
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Data empalmes 
      Caption         =   "empalmes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7185
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CommandButton Command11 
      Caption         =   "T"
      Height          =   405
      Left            =   5940
      Picture         =   "baixes impresores.frx":8355
      TabIndex        =   68
      ToolTipText     =   "Recalcular Totals"
      Top             =   900
      Width           =   540
   End
   Begin VB.Data imppantones 
      Caption         =   "imppantones"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10425
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1335
      Visible         =   0   'False
      Width           =   2415
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   -105
      Top             =   1620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Fulla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10410
      Picture         =   "baixes impresores.frx":902B
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   525
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Data bobines 
      Caption         =   "bobines"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10710
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "bobinesimp"
      Top             =   7470
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Frame Frame2 
      Caption         =   "Totals"
      Height          =   885
      Left            =   105
      TabIndex        =   8
      Top             =   10545
      Width           =   11580
      Begin VB.CommandButton Command22 
         Height          =   315
         Left            =   3615
         Picture         =   "baixes impresores.frx":AC2D
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Metres impresos dolents "
         Top             =   360
         Width           =   270
      End
      Begin VB.TextBox mtrsdolents 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   103
         Top             =   390
         Width           =   630
      End
      Begin VB.TextBox canvienfilada 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   9585
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   345
         Width           =   345
      End
      Begin VB.TextBox trentats 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   8355
         MaxLength       =   2
         TabIndex        =   82
         Top             =   360
         Width           =   420
      End
      Begin VB.TextBox pclixers 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   9015
         MaxLength       =   2
         TabIndex        =   81
         Top             =   345
         Width           =   390
      End
      Begin VB.CheckBox comandaacavada 
         Caption         =   "Acabada"
         Enabled         =   0   'False
         Height          =   225
         Left            =   10470
         TabIndex        =   80
         Top             =   120
         Width           =   1050
      End
      Begin VB.TextBox hclixe 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   390
         Width           =   555
      End
      Begin VB.TextBox hmaquina 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   885
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   390
         Width           =   615
      End
      Begin VB.TextBox hfunc 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2265
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   390
         Width           =   690
      End
      Begin VB.TextBox hajusts 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   390
         Width           =   645
      End
      Begin VB.TextBox tkilos 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6450
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   840
      End
      Begin VB.TextBox tmetres 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   375
         Width           =   840
      End
      Begin VB.TextBox kiloshora 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   7395
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   840
      End
      Begin VB.TextBox tprova 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox tbob 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4875
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   390
         Width           =   570
      End
      Begin VB.Label etestadistica 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005C31DD&
         Height          =   210
         Left            =   300
         TabIndex        =   144
         Top             =   675
         Width           =   10980
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Mtrs Dolents"
         Height          =   210
         Left            =   2985
         TabIndex        =   104
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "PClixers"
         Height          =   210
         Left            =   8940
         TabIndex        =   86
         Top             =   150
         Width           =   990
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Canvi Enf."
         Height          =   210
         Left            =   9540
         TabIndex        =   85
         Top             =   150
         Width           =   990
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Tint Rent"
         Height          =   210
         Left            =   8190
         TabIndex        =   84
         Top             =   150
         Width           =   990
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "H. Clixe"
         Height          =   210
         Left            =   225
         TabIndex        =   26
         Top             =   165
         Width           =   645
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "H. Màquina"
         Height          =   210
         Left            =   825
         TabIndex        =   25
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "H. Ajusts"
         Height          =   210
         Left            =   1665
         TabIndex        =   24
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "H. Func."
         Height          =   195
         Left            =   2340
         TabIndex        =   23
         Top             =   180
         Width           =   750
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bob."
         Height          =   210
         Left            =   4740
         TabIndex        =   22
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Metres"
         Height          =   210
         Left            =   5505
         TabIndex        =   21
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Metres/Min"
         Height          =   210
         Left            =   7365
         TabIndex        =   20
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Mtrs Prova"
         Height          =   210
         Left            =   3915
         TabIndex        =   19
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Kilos"
         Height          =   210
         Left            =   6465
         TabIndex        =   18
         Top             =   165
         Width           =   990
      End
   End
   Begin VB.TextBox nomoperari 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Escull Operari"
      Top             =   885
      Width           =   3375
   End
   Begin VB.Timer rellotge 
      Interval        =   900
      Left            =   210
      Top             =   480
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2070
      TabIndex        =   4
      Top             =   165
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Funcionament"
      Enabled         =   0   'False
      Height          =   570
      Left            =   4995
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ajust Imp."
      Enabled         =   0   'False
      Height          =   570
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Màquina"
      Enabled         =   0   'False
      Height          =   570
      Left            =   2955
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1035
   End
   Begin VB.Data impresores 
      Caption         =   "impresores"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8010
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -165
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "baixes impresores.frx":B1B7
      Height          =   2160
      Left            =   30
      OleObjectBlob   =   "baixes impresores.frx":B1CC
      TabIndex        =   5
      Top             =   1395
      Width           =   11520
   End
   Begin VB.Frame framebobines 
      Caption         =   "Bobines"
      Height          =   3840
      Left            =   120
      TabIndex        =   27
      Top             =   3810
      Width           =   11595
      Begin VB.CommandButton Command12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10290
         Picture         =   "baixes impresores.frx":D2ED
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Ensenya els empalmes"
         Top             =   2265
         Width           =   1260
      End
      Begin VB.CommandButton Command5 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10290
         TabIndex        =   29
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   10305
         Picture         =   "baixes impresores.frx":E337
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   885
         Width           =   1245
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11130
         TabIndex        =   30
         Top             =   270
         Width           =   375
      End
      Begin MSDBGrid.DBGrid reixabobines 
         Bindings        =   "baixes impresores.frx":EC21
         Height          =   3570
         Left            =   150
         OleObjectBlob   =   "baixes impresores.frx":EC33
         TabIndex        =   28
         Top             =   240
         Width           =   10080
      End
      Begin VB.Label barraestat 
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   330
         TabIndex        =   69
         Top             =   2820
         Width           =   6315
      End
   End
   Begin VB.Frame framepantones 
      Caption         =   "Pantones"
      Height          =   3945
      Left            =   4980
      TabIndex        =   34
      Top             =   3810
      Visible         =   0   'False
      Width           =   5415
      Begin MSDBGrid.DBGrid dblots 
         Bindings        =   "baixes impresores.frx":101BF
         Height          =   3675
         Left            =   4560
         OleObjectBlob   =   "baixes impresores.frx":101CE
         TabIndex        =   87
         Top             =   3525
         Visible         =   0   'False
         Width           =   5310
      End
      Begin VB.Frame Framereprint 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Reprint"
         Height          =   3540
         Left            =   750
         TabIndex        =   125
         Top             =   3045
         Visible         =   0   'False
         Width           =   5325
         Begin VB.CommandButton Command30 
            Caption         =   "Total KG"
            Height          =   435
            Left            =   2385
            TabIndex        =   131
            Top             =   1620
            Width           =   855
         End
         Begin VB.CommandButton Command32 
            Height          =   375
            Left            =   2400
            Picture         =   "baixes impresores.frx":10BA0
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   1185
            Width           =   480
         End
         Begin VB.CommandButton Command31 
            Height          =   375
            Left            =   2400
            Picture         =   "baixes impresores.frx":1112A
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   810
            Width           =   480
         End
         Begin VB.ListBox llistallaunesreprint 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2625
            Left            =   345
            TabIndex        =   126
            Top             =   795
            Width           =   1965
         End
         Begin VB.Label ettotalkgreprint 
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
            Height          =   285
            Left            =   2730
            TabIndex        =   133
            Top             =   2070
            Width           =   690
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Kg: "
            Height          =   270
            Left            =   2430
            TabIndex        =   132
            Top             =   2085
            Width           =   360
         End
         Begin VB.Label etvernis 
            BackStyle       =   0  'Transparent
            Caption         =   "_______________________________"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   60
            TabIndex        =   130
            Top             =   225
            Width           =   5220
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Llaunes utilitzades al REPRINT"
            Height          =   180
            Left            =   255
            TabIndex        =   127
            Top             =   555
            Width           =   2610
         End
      End
      Begin VB.CommandButton botollaunesreprint 
         BackColor       =   &H008080FF&
         Caption         =   "Reprint"
         Height          =   240
         Left            =   1110
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   3690
         Width           =   3375
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot12"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   11
         Left            =   3450
         MaxLength       =   30
         TabIndex        =   118
         Tag             =   "888"
         Top             =   3390
         Width           =   1350
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone12"
         DataSource      =   "imppantones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   255
         MaxLength       =   40
         TabIndex        =   116
         Tag             =   "888"
         Top             =   3405
         Width           =   3195
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone11"
         DataSource      =   "imppantones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   255
         MaxLength       =   40
         TabIndex        =   115
         Tag             =   "888"
         Top             =   3120
         Width           =   3195
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone10"
         DataSource      =   "imppantones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   255
         MaxLength       =   40
         TabIndex        =   62
         Tag             =   "888"
         Top             =   2835
         Width           =   3195
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg10"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   9
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   64
         Tag             =   "1"
         Top             =   2835
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot10"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   9
         Left            =   3450
         MaxLength       =   30
         TabIndex        =   63
         Tag             =   "888"
         Top             =   2835
         Width           =   1350
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg9"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   8
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   61
         Tag             =   "1"
         Top             =   2565
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot9"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   8
         Left            =   3450
         MaxLength       =   30
         TabIndex        =   60
         Tag             =   "888"
         Top             =   2565
         Width           =   1350
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone9"
         DataSource      =   "imppantones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   255
         MaxLength       =   40
         TabIndex        =   59
         Tag             =   "888"
         Top             =   2565
         Width           =   3195
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg8"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   7
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   58
         Tag             =   "1"
         Top             =   2310
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot8"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   7
         Left            =   3450
         MaxLength       =   30
         TabIndex        =   57
         Tag             =   "888"
         Top             =   2310
         Width           =   1350
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone8"
         DataSource      =   "imppantones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   255
         MaxLength       =   40
         TabIndex        =   56
         Tag             =   "888"
         Top             =   2310
         Width           =   3195
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg7"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   6
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   55
         Tag             =   "1"
         Top             =   2025
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot7"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   6
         Left            =   3450
         MaxLength       =   30
         TabIndex        =   54
         Tag             =   "888"
         Top             =   2025
         Width           =   1350
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone7"
         DataSource      =   "imppantones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   255
         MaxLength       =   40
         TabIndex        =   53
         Tag             =   "888"
         Top             =   2025
         Width           =   3195
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg6"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   5
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   52
         Tag             =   "1"
         Top             =   1755
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot6"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   5
         Left            =   3450
         MaxLength       =   30
         TabIndex        =   51
         Tag             =   "888"
         Top             =   1755
         Width           =   1350
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone6"
         DataSource      =   "imppantones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   255
         MaxLength       =   40
         TabIndex        =   50
         Tag             =   "888"
         Top             =   1755
         Width           =   3195
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg5"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   4
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   49
         Tag             =   "1"
         Top             =   1470
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot5"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   4
         Left            =   3450
         MaxLength       =   30
         TabIndex        =   48
         Tag             =   "888"
         Top             =   1470
         Width           =   1350
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone5"
         DataSource      =   "imppantones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   255
         MaxLength       =   40
         TabIndex        =   47
         Tag             =   "888"
         Top             =   1470
         Width           =   3195
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg4"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   3
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   46
         Tag             =   "1"
         Top             =   1200
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot4"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   3
         Left            =   3450
         MaxLength       =   30
         TabIndex        =   45
         Tag             =   "888"
         Top             =   1200
         Width           =   1350
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone4"
         DataSource      =   "imppantones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   255
         MaxLength       =   40
         TabIndex        =   44
         Tag             =   "888"
         Top             =   1200
         Width           =   3195
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg3"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   2
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   43
         Tag             =   "1"
         Top             =   930
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot3"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   2
         Left            =   3450
         MaxLength       =   30
         TabIndex        =   42
         Tag             =   "888"
         Top             =   930
         Width           =   1350
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone3"
         DataSource      =   "imppantones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   255
         MaxLength       =   40
         TabIndex        =   41
         Tag             =   "888"
         Top             =   930
         Width           =   3195
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg2"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   1
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   40
         Tag             =   "1"
         Top             =   660
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot2"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   1
         Left            =   3450
         MaxLength       =   30
         TabIndex        =   39
         Tag             =   "888"
         Top             =   660
         Width           =   1350
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone2"
         DataSource      =   "imppantones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   255
         MaxLength       =   40
         TabIndex        =   38
         Tag             =   "888"
         Top             =   660
         Width           =   3195
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg1"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   0
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   37
         Tag             =   "1"
         Top             =   375
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot1"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   0
         Left            =   3450
         MaxLength       =   30
         TabIndex        =   36
         Tag             =   "888"
         Top             =   375
         Width           =   1350
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone1"
         DataSource      =   "imppantones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   255
         MaxLength       =   40
         TabIndex        =   35
         Tag             =   "888"
         Top             =   375
         Width           =   3195
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot11"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   10
         Left            =   3450
         MaxLength       =   30
         TabIndex        =   117
         Tag             =   "888"
         Top             =   3120
         Width           =   1350
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg11"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   10
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   119
         Tag             =   "1"
         Top             =   3120
         Width           =   550
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg12"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   11
         Left            =   4800
         MaxLength       =   8
         TabIndex        =   120
         Tag             =   "1"
         Top             =   3405
         Width           =   550
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "11 12"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   15
         TabIndex        =   121
         Top             =   3135
         Width           =   285
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1 2 3 4 5 6 7 8 9 10"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         Left            =   45
         TabIndex        =   66
         Top             =   390
         Width           =   195
      End
      Begin VB.Label Label2 
         Caption         =   "NOM                                             LOT                        KG"
         Height          =   255
         Left            =   1200
         TabIndex        =   65
         Top             =   150
         Width           =   4080
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Eines"
      Height          =   1755
      Left            =   135
      TabIndex        =   145
      Top             =   7725
      Width           =   3420
      Begin VB.CheckBox Checkescanerendollat 
         Caption         =   "Check1"
         Height          =   255
         Left            =   3135
         TabIndex        =   160
         Top             =   690
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Agrupa Treballs"
         Height          =   600
         Left            =   2385
         Picture         =   "baixes impresores.frx":116B4
         Style           =   1  'Graphical
         TabIndex        =   157
         ToolTipText     =   "Agrupar Treballs"
         Top             =   300
         Width           =   690
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FF8080&
         Caption         =   "Llistat Pantones"
         Height          =   705
         Left            =   1665
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   915
         Width           =   885
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00FF8080&
         Height          =   705
         Left            =   2550
         Picture         =   "baixes impresores.frx":134A6
         Style           =   1  'Graphical
         TabIndex        =   151
         ToolTipText     =   "Manteniments Varis"
         Top             =   915
         Width           =   525
      End
      Begin VB.CommandButton Command21 
         Height          =   660
         Left            =   975
         Picture         =   "baixes impresores.frx":13A30
         Style           =   1  'Graphical
         TabIndex        =   150
         ToolTipText     =   "Calcul diametre"
         Top             =   240
         Width           =   645
      End
      Begin VB.CommandButton Command25 
         Height          =   660
         Left            =   255
         Picture         =   "baixes impresores.frx":13FBA
         Style           =   1  'Graphical
         TabIndex        =   149
         ToolTipText     =   "Impresió de comanda, baixes muntadora, etc..."
         Top             =   240
         Width           =   645
      End
      Begin VB.CommandButton Command26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   285
         MouseIcon       =   "baixes impresores.frx":14544
         Picture         =   "baixes impresores.frx":1520E
         Style           =   1  'Graphical
         TabIndex        =   148
         ToolTipText     =   "Canviar dades anilox i tintes."
         Top             =   930
         Width           =   1320
      End
      Begin VB.CommandButton Command33 
         Height          =   300
         Left            =   3030
         Picture         =   "baixes impresores.frx":15B80
         Style           =   1  'Graphical
         TabIndex        =   147
         ToolTipText     =   "Observacions del treball operari."
         Top             =   1335
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton Command35 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1710
         Picture         =   "baixes impresores.frx":1610A
         Style           =   1  'Graphical
         TabIndex        =   146
         ToolTipText     =   "Reimprimir VQ"
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comunicació"
      Height          =   885
      Left            =   180
      TabIndex        =   153
      Top             =   9585
      Width           =   3390
      Begin VB.CommandButton Command24 
         BackColor       =   &H00FFC0C0&
         Height          =   495
         Left            =   2655
         Picture         =   "baixes impresores.frx":17018
         Style           =   1  'Graphical
         TabIndex        =   159
         ToolTipText     =   "Calendari"
         Top             =   240
         Width           =   630
      End
      Begin VB.CommandButton Command23 
         BackColor       =   &H00F8FDB5&
         Height          =   495
         Left            =   2010
         Picture         =   "baixes impresores.frx":1761A
         Style           =   1  'Graphical
         TabIndex        =   158
         ToolTipText     =   "Calendari"
         Top             =   240
         Width           =   630
      End
      Begin VB.CommandButton Command28 
         Height          =   495
         Left            =   120
         Picture         =   "baixes impresores.frx":17773
         Style           =   1  'Graphical
         TabIndex        =   156
         ToolTipText     =   "Enviar un email per dir alguna cosa a oficines."
         Top             =   240
         Width           =   630
      End
      Begin VB.CommandButton Command39 
         BackColor       =   &H0017D062&
         Height          =   495
         Left            =   750
         Picture         =   "baixes impresores.frx":17CFD
         Style           =   1  'Graphical
         TabIndex        =   155
         ToolTipText     =   "Chat maquinista"
         Top             =   240
         Width           =   630
      End
      Begin VB.CommandButton Command40 
         BackColor       =   &H006BEBB1&
         Height          =   495
         Left            =   1380
         Picture         =   "baixes impresores.frx":17D6E
         Style           =   1  'Graphical
         TabIndex        =   154
         ToolTipText     =   "Chat ajudant."
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.Frame framesegonapantalla 
      Height          =   11445
      Left            =   0
      TabIndex        =   161
      Top             =   -15
      Visible         =   0   'False
      Width           =   11700
      Begin VB.Image Image2 
         Height          =   3855
         Left            =   2610
         Picture         =   "baixes impresores.frx":17DDF
         Stretch         =   -1  'True
         Top             =   405
         Width           =   5790
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Continua a la pantalla Tàctil"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   870
         TabIndex        =   162
         Top             =   4800
         Width           =   9930
      End
   End
   Begin VB.Shape reciclarmaterial1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   15
      Shape           =   3  'Circle
      Top             =   195
      Width           =   225
   End
   Begin VB.Label nomoperari2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   1980
      TabIndex        =   143
      Top             =   1200
      Width           =   3525
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Observacions per l'operari"
      Height          =   210
      Left            =   5355
      TabIndex        =   135
      Top             =   7650
      Width           =   3690
   End
   Begin VB.Label texteimpresio 
      BackColor       =   &H00FFFFFF&
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
      Height          =   240
      Left            =   255
      TabIndex        =   72
      Top             =   3570
      Width           =   5280
   End
   Begin VB.Label ettoleranciaample 
      Alignment       =   1  'Right Justify
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
      Height          =   240
      Left            =   5700
      TabIndex        =   112
      Top             =   3585
      Width           =   3900
   End
   Begin VB.Label codidebarres 
      Caption         =   "Label18"
      Height          =   60
      Left            =   15
      TabIndex        =   111
      Top             =   420
      Width           =   75
   End
   Begin VB.Label stockopacking 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   2595
      TabIndex        =   107
      ToolTipText     =   "Si es 'E' son bobines d'Estoc i si es 'P' de Packing-list"
      Top             =   90
      Width           =   330
   End
   Begin VB.Label etmetresajust 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   690
      TabIndex        =   102
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label avisapantalla 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   8070
      TabIndex        =   94
      Top             =   825
      Width           =   3450
   End
   Begin VB.Label controlstock 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   30
      TabIndex        =   91
      Top             =   0
      Width           =   555
   End
   Begin VB.Label ettreball 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   9675
      TabIndex        =   90
      Top             =   3585
      Width           =   1725
   End
   Begin VB.Label firmat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7215
      TabIndex        =   88
      ToolTipText     =   "Codi operari que ha firmat"
      Top             =   45
      Width           =   1245
   End
   Begin VB.Label client 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   195
      TabIndex        =   33
      Top             =   615
      Visible         =   0   'False
      Width           =   525
      WordWrap        =   -1  'True
   End
   Begin VB.Label hora 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   45
      TabIndex        =   6
      Top             =   870
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Comanda"
      Height          =   255
      Left            =   450
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu memail 
      Caption         =   "email"
      Visible         =   0   'False
      Begin VB.Menu moficines 
         Caption         =   "OFICINES"
      End
      Begin VB.Menu mencarregat 
         Caption         =   "ENCARREGAT"
      End
   End
   Begin VB.Menu mllistat 
      Caption         =   "Llistats"
      Visible         =   0   'False
      Begin VB.Menu mvisualitzacomanda 
         Caption         =   "Visualitzar la comanda."
      End
      Begin VB.Menu mvisualitzarlabaixademuntadora 
         Caption         =   "Visualitzar la baixa de muntadora."
      End
      Begin VB.Menu mveurepdf 
         Caption         =   "Visualitzar PDF."
      End
      Begin VB.Menu mvisualitzarIMP 
         Caption         =   "Visualitzar IMP."
      End
      Begin VB.Menu mveureMODIFICACIONS 
         Caption         =   "Visualitzar MODIFICACIONS."
      End
   End
   Begin VB.Menu mmenuarrancariVQ 
      Caption         =   "menuarrancariVQ"
      Visible         =   0   'False
      Begin VB.Menu marrancar 
         Caption         =   "Arrancar"
      End
      Begin VB.Menu metiquetavq 
         Caption         =   "Etiqueta VQ"
      End
      Begin VB.Menu m_impvdelta 
         Caption         =   "Imprimir Valor Delta"
         Begin VB.Menu m_DELTANEGRE 
            Caption         =   "Delta NEGRE"
         End
         Begin VB.Menu m_DELTACYAN 
            Caption         =   "Delta CYAN"
         End
         Begin VB.Menu m_DELTAMAGENTA 
            Caption         =   "Delta MAGENTA"
         End
         Begin VB.Menu m_DELTAGROC 
            Caption         =   "Delta GROC"
         End
      End
   End
   Begin VB.Menu mmenucalculadores 
      Caption         =   "menucalculadores"
      Visible         =   0   'False
      Begin VB.Menu mbuscarbobinaamagatzem 
         Caption         =   "Buscar Bobina a magatzem"
         Visible         =   0   'False
      End
      Begin VB.Menu mcalculadora 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu mcalculadoradiametre 
         Caption         =   "Calculadora Diàmetre"
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = 2
Private Const HWND_TOP = 0
Private Const SWP_NOZORDER = &H4
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim rstbobinesdentrada As Recordset
Dim nomfitxertemporalbobent As String
Dim direnvio As Long
Dim vtubbase As Double
Dim vvalidaciocodidebarres As String
Dim vdigimarc As Boolean
Dim vhihadigimarc As Boolean
Dim vnopreguntar As Boolean
Dim vavispeu As String
Dim vtotalmetres As Double
Dim vgrupmaterialcompatible As Double
Dim vestemfentfingerprint As Boolean
Dim vmetresarrancada As Double
Dim vescanerendollat As Boolean

Sub m_DELTANEGRE_click()
   etiqueta_DELTA "NEGRE"
End Sub
Sub m_DELTAMAGENTA_click()
   etiqueta_DELTA "MAGENTA"
End Sub
Sub m_DELTAGROC_click()
   etiqueta_DELTA "GROC"
End Sub
Sub m_DELTACYAN_click()
   etiqueta_DELTA "CYAN"
End Sub
Sub etiqueta_DELTA(vcolor As String)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vvalorcolor As String
  vvalorcolor = InputBox("Escriu el valor DELTA del color " + vcolor + ".", "Etiqueta Delta")
  If cadbl(vvalorcolor) = 0 Then Exit Sub
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "Etiqueta DeltaColors.rpt", 1)
  oreport.FormulaFields.GetItemByName("color").text = "'" + vcolor + "'"
  oreport.FormulaFields.GetItemByName("valorColor").text = "'" + vvalorcolor + "'"
 Set vprinter = triarimpresoratickets
  'MsgBox vprinter.DeviceName
  If InStr(1, UCase(triarimpresoratickets.DeviceName), "TICKETS") > 0 Or InStr(1, UCase(triarimpresoratickets.DeviceName), "80 PRINTER") > 0 Then
     oreport.SelectPrinter vprinter.DriverName, vprinter.DeviceName, vprinter.Port
       Else: MsgBox "No s'ha trobat la impresora Tickets instal.lada al sistema", vbCritical, "Error": Exit Sub
  End If
  oreport.PaperOrientation = crDefaultPaperOrientation
  oreport.DisplayProgressDialog = False
  oreport.PrintOut False, 1

 

   
End Sub

Private Sub Command17_Click()
  Shell "\\SERVERPRODU\Dades\progcomandes\aplicacio\Manteniment tintes.exe agrupartreballs", vbNormalFocus
End Sub

Private Sub Command23_Click()
  Dim vrutaThunderbird As String
  ratoli "espera"
  vrutaThunderbird = "C:\Program Files\Mozilla Thunderbird\thunderbird.exe"
  If Not existeix(vrutaThunderbird) Then vrutaThunderbird = "\\serverprodu\Dades\progcomandes\aplicacio\CalendariThunderbird\ThunderbirdPortable\ThunderbirdPortable.exe"
  Shell vrutaThunderbird, vbMaximizedFocus
  wait 2
  ratoli "normal"
End Sub

Private Sub Command24_Click()
  obrir_document "https://docs.google.com/spreadsheets/d/1JsSjjQ6-qyosT-yB1ex5NRkHOfYnHGvt/edit?gid=1610146904#gid=1610146904"
End Sub

Private Sub Command39_Click()
'  If existeix("c:\ordprog.ini") Then
    Load formCHAT
    formCHAT.carregar_missatges_operari "I", cadbl(numop)
    formCHAT.Show 1
    comprovarsihihamissatgesCHAT
 ' End If
End Sub

Private Sub Command40_Click()
'   If existeix("c:\ordprog.ini") Then
    Load formCHAT
    formCHAT.carregar_missatges_operari "I", cadbl(numop2)
    formCHAT.Show 1
    comprovarsihihamissatgesCHAT
 ' End If
End Sub

Private Sub Form_Activate()
 'If existeix("c:\ordprog.ini") Then
 'obrirprogramalecturaCB
 'enviaremailgeneric "miquel.inplacsa@gmail.com", "444-Ajust de " + vbobina + " " + atrim(cadbl(vmetres - vmetresanteriors)) + " metres per diametre. " + atrim(Now), "Prova"
 If nummaq = 7 Then Checkescanerendollat.Value = 1
 If nummaq = 9 Then Checkescanerendollat.Value = 1
 'form1.Hide: formannex.Hide: formrevisarCQ.Show 1: End
  ' form1.Hide: formannex.Hide: formrevisarCQ.Show 1: End

  'MsgBox mirarsihihamaterialLAM(221507)
' Dim v As Boolean

'    Form1.Hide
'    formannex.Hide
'    Load formCHAT
'    formCHAT.carregar_missatges_operari "I", cadbl(numop)
'    formCHAT.Show 1
'    comprovarsihihamissatgesCHAT
'    End
'  End If

End Sub

Sub mbuscarbobinaamagatzem_click()
     escriure_ini "Baixes", "numcomanda", comanda, "comandes.ini"
     Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "palets.exe comandes.ini FiltrarBobinesImpresores", vbNormalFocus
End Sub
Sub mcalculadoradiametre_click()
Load calculdiametre
  calculdiametre.micres = micrescomanda
  calculdiametre.Show 1
End Sub
Sub mcalculadora_click()
 Static id As Double
 
 On Error GoTo cridar
 AppActivate id
 Exit Sub
cridar:
 id = Shell("C:\WINDOWS\SYSTEM32\CALC.EXE", vbNormalFocus)
End Sub

Sub metiquetavq_click()
   imprimir_controlqualitatVQ cadbl(comanda), True, True
End Sub
Sub marrancar_click()
    imprimir_full_arrancar_rentar cadbl(comanda)
End Sub
Sub mvisualitzacomanda_click()

  escriure_ini "Baixes", "imprimircomanda", cadbl(comanda), "comandes.ini"
  Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", "comandes.ini")) + "comandes.exe - imprimir", vbHide
   missatgevist.Show 1
End Sub
Sub mvisualitzarlabaixademuntadora_click()
   Dim vnumc As Double
   Dim vrutaPDF As String
   vnumc = cadbl(comanda)
   If vnumc = 0 Then Exit Sub
   vrutaPDF = "Les_" + atrim(atrim(Int(cadbl(vnumc) / 1000)) + "000")
   vrutaPDF = llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\" + vrutaPDF + "\" + atrim(numc)
   vrutaPDF = vrutaPDF + atrim(vnumc) + "\" + atrim(vnumc) + "_BaixaMuntadora.pdf"
   If existeix(vrutaPDF) Then obrir_document vrutaPDF
End Sub
Sub mveureMODIFICACIONS_click()
obrir_fitxer_modificacions
End Sub
Sub mveurepdf_CLICK()
  veureelpdf
End Sub
Sub mvisualitzarIMP_CLICK()
   veureelimp
End Sub
Sub moficines_click()
     enviar_email_oficines
End Sub
Sub mencarregat_click()
     enviar_email_encarregat
End Sub
Sub passarlotsaprincipal()
Dim i As Byte
  Dim vordre As Double
  If form1.impresores.Recordset.EOF Then Exit Sub
  form1.imppantones.Recordset.Edit
  For i = 10 To 7
    form1.Controls("pantone")(i) = "": form1.Controls("compantone")(i) = "": form1.Controls("kbpantone")(i) = 0
  Next i
  For i = 0 To 7
     ' If atrim(compantone(i)) <> "" Then
       vordre = cadbl(formaniloxos.ordre(i)) - 1
       If cadbl(formaniloxos.ordre(i)) < 1 Then vordre = i + 1
       form1.Controls("pantone")(vordre) = atrim(formaniloxos.tintacomanda(i))
       form1.Controls("compantone")(vordre) = atrim(formaniloxos.compantone(i))
       'If cadbl(Form1.Controls("kbpantone")(atrim(i))) = 0 Then
       form1.Controls("kbpantone")(vordre) = atrim(cadbl(formaniloxos.kbpantone(i)))
      'End If
   Next i
   'EN MIRALLES A DATA 01/07/2021 EM FA TREURE 8-ETOXI I CANVIAR 9-R25 PER 80/20
    If form1.Controls("compantone")(8) = "" Then
           form1.Controls("pantone")(8) = ""
           form1.Controls("compantone")(8) = ""
           form1.Controls("kbpantone")(8) = "0"
         Else
             form1.Controls("pantone")(8) = formaniloxos.Label8
             form1.Controls("compantone")(8) = atrim(formaniloxos.compantone(8))
             form1.Controls("kbpantone")(8) = atrim(cadbl(formaniloxos.kbpantone(8)))
   End If
   form1.Controls("pantone")(9) = formaniloxos.Label9
   form1.Controls("compantone")(9) = atrim(formaniloxos.compantone(9))
   form1.Controls("kbpantone")(atrim(9)) = atrim(cadbl(formaniloxos.kbpantone(9)))
   form1.imppantones.Recordset.Update
End Sub

Sub calcular_totals()
  Dim total As Double
  Dim hores As Double
  Dim bkimp As Double
  Dim bkbob As Double
  barraestat.caption = "Calculant els totals..."
  'calculant.Visible = True
  fcalculant.Show 0, Me
  calculant.Top = 2222
  DoEvents
  
  
  On Error GoTo fi
  
  If impresores.Recordset.EOF Or cadbl(impresores.Recordset!id) = 0 Then
    barraestat.caption = ""
    GoTo fi
  End If
  
  '---- guardo la posicio de linies imp i de bobina x recuperarlames avall
  If impresores.Recordset!tipus = "F" Then bkimp = atrim(cadbl(impresores.Recordset!id))
  If Not bobines.Recordset.EOF Then bkbob = atrim(cadbl(bobines.Recordset!numerodebobina))
  '------
  If impresores.Recordset.EditMode > 0 Then impresores.Recordset.Update
  reixa.EditActive = False
  reixabobines.EditActive = False
  command15.tag = ""
 ' On Error Resume Next
  impresores.Recordset.MoveFirst
  While Not impresores.Recordset.EOF
   If impresores.Recordset.EditMode = 0 Then impresores.Recordset.Edit
   Set rsttmp = dbtmpb.OpenRecordset("select count(*) as elgran from bobinesimp where controlid=" + atrim(impresores.Recordset!id))
   If Not rsttmp.EOF Then impresores.Recordset!totalbobines = rsttmp!elgran
  
   Set rsttmp = dbtmpb.OpenRecordset("select sum(kilos) as elgran from bobinesimp where controlid=" + atrim(impresores.Recordset!id))
   If Not rsttmp.EOF Then impresores.Recordset!totalkilos = rsttmp!elgran
  
   Set rsttmp = dbtmpb.OpenRecordset("select sum(metres) as elgran from bobinesimp where controlid=" + atrim(impresores.Recordset!id))
   If Not rsttmp.EOF Then impresores.Recordset!totalmetres = rsttmp!elgran
  
   Set rsttmp = dbtmpb.OpenRecordset("select id,metres from bobinesimp where metres=0 and controlid=" + atrim(impresores.Recordset!id))
   If Not rsttmp.EOF Then
    If rsttmp!id <> bobines.Recordset!id Then MsgBox "Hi ha bobines sense metres"
   End If
   impresores.Recordset.Update
   With impresores.Recordset
    total = 0
    On Error Resume Next
    total = DateDiff("n", CVDate(atrim(!datainici) + " " + atrim(!horainici)), CVDate(atrim(!datafi) + " " + atrim(!horafi)))
    If total < 0 Then MsgBox "Hi ha una data d'inici mes gran que la de fi a les linies de TIPUS": command15.tag = "Error"
    If total > 999 Then MsgBox "Hi ha un TIPUS que passa de 999 minuts es considera un error d'entrada REVISEU SI LES DATES D'INICI I FINAL SON CORRECTES": command15.tag = "Error"
    total = Redondejar(total / 60, 2)
    If impresores.Recordset.EditMode = 0 Then impresores.Recordset.Edit
    impresores.Recordset!totalhores = total
    impresores.Recordset.Update
   End With
  impresores.Recordset.MoveNext
 Wend
  'If Not (impresores.Recordset.EOF And impresores.Recordset.BOF) Then impresores.UpdateRecord
  'impresores.UpdateRecord
  'reixa.Refresh
  impresores.Refresh
  On Error GoTo 0
  ensenyar_totalstotals
  possar_metres_min
  If cadbl(tmetres) > metrescomanda Then
    avisapantalla = "     ATENCIO!!!!!!" + Chr(10) + Chr(13) + "Comanda de " + atrim(metrescomanda) + " Metres"
     Else: If InStr(1, avisapantalla, "ATENCIO!!!!!!") > 0 Then avisapantalla = ""
  End If
  Set rstmp = Nothing
  barraestat.caption = ""
  
  '---recupero la pocisio de linis imp i de bob
   If bkimp > 0 Then
     impresores.Recordset.FindFirst "id=" + atrim(bkimp)
     bobines.Recordset.FindFirst "numerodebobina=" + atrim(bkbob)
   Else: impresores.Recordset.MoveLast
  End If
  '---
fi:
'calculant.Visible = False
'If err.Description <> "" Then MsgBox err.Description

barraestat.caption = ""

Unload fcalculant
form1.SetFocus
Set rsttmp = Nothing

End Sub

Sub possar_metres_min()
  Dim v As Double
  Dim mtrsmin As Double
  mtrsmin = 0
  DoEvents
  v = cadbl(hfunc)
  f = (Int(v) * 60) + (((v - Int(v)) * 100) * 60 / 100)
  If f > 0 Then
     mtrsmin = cadbl(tmetres) / (f)
     kiloshora = Redondejar(mtrsmin, 2)
    Else: kiloshora = "0"
  End If
  If mtrsmin > 340 And maquina = 7 Then MsgBox "Els metres/min superen els 340 mtrs/min, es considera impossible"
  If mtrsmin > 440 And maquina = 9 Then MsgBox "Els metres/min superen els 440 mtrs/min, es considera impossible"
End Sub

Sub ensenyar_totalstotals()
 tbob = 0: hfunc = 0: hclixe = 0: hmaquina = 0: hajusts = 0: tkilos = 0: tmetres = 0: tprova = 0:
'total bobines
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalbobines) as elgran from Impressores totalbobines where comanda=" + atrim(cadbl(comanda.text)))
  If Not rsttmp.EOF Then tbob = cadbl(rsttmp!elgran)

  
'hores func
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from Impressores totalbobines where comanda=" + atrim(cadbl(comanda.text)) + " and tipus='F'")
  If Not rsttmp.EOF Then hfunc = cadbl(rsttmp!elgran)
  
'hores clixe
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from Impressores totalbobines where comanda=" + atrim(cadbl(comanda.text)) + " and tipus='C'")
  If Not rsttmp.EOF Then hclixe = cadbl(rsttmp!elgran)

'hores maquina
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from Impressores totalbobines where comanda=" + atrim(cadbl(comanda.text)) + " and tipus='M'")
  If Not rsttmp.EOF Then hmaquina = cadbl(rsttmp!elgran)

'hores ajusts
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from Impressores totalbobines where comanda=" + atrim(cadbl(comanda.text)) + " and tipus='A'")
  If Not rsttmp.EOF Then hajusts = cadbl(rsttmp!elgran)

'total kilos
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalkilos) as elgran from Impressores totalbobines where comanda=" + atrim(cadbl(comanda.text)))
  If Not rsttmp.EOF Then tkilos = cadbl(rsttmp!elgran)
  
'total metres
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalmetres) as elgran from Impressores totalbobines where comanda=" + atrim(cadbl(comanda.text)))
  If Not rsttmp.EOF Then tmetres = cadbl(rsttmp!elgran)
'total metres arrancada
  vmetresarrancada = 0
  Set rsttmp = dbtmpb.OpenRecordset("select paletbobprova  from Impressores where comanda=" + atrim(cadbl(comanda.text)) + " and mid(paletbobprova,1,3)='[A]'")
  While Not rsttmp.EOF
    vmetresarrancada = vmetresarrancada + cadbl(Mid(Mid(rsttmp!paletbobprova, InStr(1, rsttmp!paletbobprova, "-") + 1), 1, 3))
    rsttmp.MoveNext
  Wend
'total prova
  Set rsttmp = dbtmpb.OpenRecordset("select sum(mtrsprova) as elgran from Impressores totalbobines where comanda=" + atrim(cadbl(comanda.text)))
  If Not rsttmp.EOF Then tprova = cadbl(rsttmp!elgran) + vmetresarrancada
  
  
  guarda_totals
  ensenya_totals
End Sub

Sub guarda_totals()
Dim rsttinta As Recordset
Set rsttinta = dbtmpb.OpenRecordset("Select max(numeromaquina) as maq,max(operari) as op,sum(kgtinta) as kg from impressores where comanda=" + atrim(cadbl(comanda)))
Set rsttmp = dbtmpb.OpenRecordset("select * from impressorestot where comanda=" + atrim(cadbl(comanda)))
  If rsttmp.EOF Then
      rsttmp.AddNew
    Else: rsttmp.Edit
  End If
  With rsttmp
    !firmat = atrim(firmat.caption)
    !comanda = cadbl(comanda)
    !hclixe = cadbl(hclixe)
    !hmaquina = cadbl(hmaquina)
    !hajusts = cadbl(hajusts)
    !hfuncio = cadbl(hfunc)
    !tbobines = cadbl(tbob)
    !tprova = cadbl(tprova)
    !tkilos = cadbl(tkilos)
     !tmetresdolents = cadbl(mtrsdolents)
    !tmetres = cadbl(tmetres)
    !metresmin = cadbl(kiloshora)
    !tintersrentats = cadbl(trentats)
    !portaclixers = cadbl(pclixers)
    !canvienfilada = canvienfilada
    !acavada = cadbl(comandaacavada.Value)
    If Not rsttinta.EOF Then
      !kilostinta = cadbl(rsttinta!kg)
      !impressora = cadbl(rsttinta!maq)
      !operari = cadbl(rsttinta!op)
    End If
    'If Not (bobines.Recordset.EOF Or bobines.Recordset.BOF) Then
    ' !kilostinta = cadbl(bobines.Recordset!kgtinta)
    ' If Not IsNull(bobines.Recordset!datafi) Then !dataimpressio = bobines.Recordset!datafi
     '!impressora = cadbl(impresores.Recordset!numeromaquina)
     '!operari = cadbl(bobines.Recordset!operari)
    'End If
   .Update
  End With
  Set rsttinta = Nothing
  Set rsttmp = Nothing
End Sub



Sub ensenya_totals()
  Set rsttmp = dbtmpb.OpenRecordset("select * from impressorestot where comanda=" + atrim(cadbl(comanda)))
  If rsttmp.EOF Then Exit Sub
  If cadbl(rsttmp!comandafingerprintoriginal) > 0 Then
     form1.BackColor = &HFFFF&
     vestemfentfingerprint = True
     Else: form1.BackColor = &H80000005
  End If
  With rsttmp
    'comanda = atrim(!comanda)
    firmat = atrim(!firmat)
    hclixe = atrim(!hclixe)
    hmaquina = atrim(!hmaquina)
    hajusts = atrim(!hajusts)
    hfunc = atrim(!hfuncio)
    tbob = atrim(!tbobines)
    tprova = atrim(!tprova)
    mtrsdolents = atrim(!tmetresdolents)
    tkilos = atrim(!tkilos)
    tmetres = atrim(!tmetres)
    kiloshora = atrim(!metresmin)
    trentats = cadbl(!tintersrentats)
    pclixers = cadbl(!portaclixers)
    canvienfilada = atrim(!canvienfilada)
    comandaacavada.Value = IIf(cadbl(!acavada) = 0, 0, 1)
    vvalidaciocodidebarres = atrim(rsttmp!validaciocodidebarres)
    vdigimarc = cabool(rsttmp!validaciodigimarc)
    If codidebarres = "" Then vvalidaciocodidebarres = "-"
    'If Not (bobines.Recordset.EOF Or bobines.Recordset.BOF) Then
    ' !kilostinta = cadbl(bobines.Recordset!kgtinta)
    ' If Not IsNull(bobines.Recordset!datafi) Then !dataimpressio = bobines.Recordset!datafi
     '!impressora = cadbl(impresores.Recordset!numeromaquina)
     '!operari = cadbl(bobines.Recordset!operari)
    'End If
  
  End With

End Sub


Private Sub bobentrada_Error(ByVal DataError As Integer, Response As Integer)
  If DataError = 16389 Then
   MsgBox "Error de conversió de dades. Pel anular prem <Esc> "
      bobentrada.col = 1
      Response = 0
  End If
  
End Sub

Private Sub bobentrada_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 46 Then KeyCode = 0
End Sub

Private Sub bobentrada_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub bobentrada_KeyUp(KeyCode As Integer, Shift As Integer)
If bobentrada.col = 0 And Len(bobentrada.text) = 5 And KeyCode > 46 Then bobentrada.col = 1
End Sub

Private Sub bobentrada_LostFocus()
On Error Resume Next
  bobinesent.UpdateRecord
  botoensenyarpacking.tag = "bobentrada"
End Sub

Private Sub bobentrada_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim temps As Byte
   If Shift = 2 Then
     temps = cadbl(InputBox("Entra el temps d'espera en segons", "Temps abans d'imprimir", 2))
     imprimir_controlqualitatbobinaentrada cadbl(bobinesent.Recordset!palet), cadbl(bobinesent.Recordset!bobina), 0, temps
   End If
End Sub

Private Sub bobentrada_OnAddNew()
bobinesent.Recordset!id = bobines.Recordset!id
 bobentrada.col = 0
End Sub

Private Sub bobines_Reposition()
On Error Resume Next
 empalmes.UpdateRecord
 If empalmes.Recordset.EditMode = 0 Then
   empalmes.RecordSource = "select * from impempalmes where id=" + atrim(cadbl(bobines.Recordset!id))
   empalmes.Refresh
 End If
 bobinesent.UpdateRecord
 If bobinesent.Recordset.EditMode = 0 Then
   bobinesent.RecordSource = "select * from bobinesentimp where id=" + atrim(cadbl(bobines.Recordset!id))
   bobinesent.Refresh
 End If
 
End Sub

Private Sub clixes_Click()
 
End Sub
Sub finalitza_seccio()
  On Error GoTo fi
  If impresores.Recordset.EOF Then Exit Sub
  On Error Resume Next
  impresores.Recordset.MoveLast
  If IsDate(impresores.Recordset!datafi) Then r = "no": Exit Sub
  'comprovo si fa menys de 30segons i si es aixi borro la linia
  'If DateDiff("s", CVDate(Str(impresores.Recordset!datainici) + " " + Str(impresores.Recordset!horainici)), Now) < 30 Then
  '   impresores.Recordset.Delete
  '   impresores.Recordset.MoveLast
  '   Exit Sub
  '     Else:
  '       If MsgBox("Es finalitzarà la seccio actual.", vbCritical + vbOKCancel, "Atenció") = vbOK Then
  '           r = "si": Exit Sub
  '       End If
  'End If
  'fins aqui
  impresores.Recordset.Edit
  impresores.Recordset!datafi = Date: impresores.Recordset!horafi = Time
  Select Case impresores.Recordset!tipus
   Case "C"
   Case "M"
   Case "A"
   Case "V"
   Case "F"
  End Select
  impresores.Recordset.Update
calcular_totals
fi:
End Sub



Private Sub bobsajust_Click()
  paletsajust.Show 1
End Sub


Sub obrir_fitxer_modificacions()
   Dim vpdfmodifi As String
  ' carregaravisosmanteniment False
  ' avisosxrseccio.Show 1
  
  vpdfmodifi = rutamodifispdftreball(id_treball, ordremodificacio)
  If existeix(vpdfmodifi) Then obrir_document vpdfmodifi
End Sub

Function rutamodifispdftreball(vidtreball As Double, vordre As Double) As String
   On Error Resume Next
   MkDir ruta_documentacio_clixes + "\" + Format(vidtreball, "00000")
   rutamodifispdftreball = ruta_documentacio_clixes + "\" + Format(vidtreball, "00000") + "\MODIFI" + Format(vidtreball, "00000") + "-" + Format(vordre, "000") + ".pdf"
   
End Function

Private Sub botodescansrelleu_Click()
   Load formdescansirelleu
   If Not impresores.Recordset.EOF Then
        impresores.Recordset.MoveLast
        If Not impresores.Recordset.EOF Then
           If Not IsDate(impresores.Recordset!datafi) And Not IsDate(impresores.Recordset!horafi) Then
             formdescansirelleu.Command1.Enabled = False
             formdescansirelleu.etnonou = "Funcionament activat a baixes"
           End If
        End If
   End If
   formdescansirelleu.datacontroldescansirelleu.DatabaseName = cami
   formdescansirelleu.etnomoperari = nomoperari
   formdescansirelleu.etnomoperari.tag = atrim(numop)
   formdescansirelleu.Show 1
   comprovarsiacabarfuncionament
End Sub
Sub comprovarsiacabarfuncionament()
  
End Sub
Private Sub botoensenyarpacking_Click()
 Dim palet As Double
 Dim bobina As Double
 Dim utilitzades As String
 If bobines.Recordset.EOF Then Exit Sub
 utilitzades = "noutilitzades"
 If veuretotesbobent.Value = 1 Then utilitzades = ""
 botoensenyarpacking.tag = ""
 carregar_bobinesdentrada "ensenyar" + utilitzades, 1, palet, bobina, cadbl(comanda), , , , True
 
 ratoli "espera"
 If palet > 0 And bobina > 0 And Not bobines.Recordset.EOF Then
    'bobentrada.Columns("Palet") = atrim(palet): bobentrada.Columns("Bobina") = atrim(bobina)
    If etmaterialexacte <> "" Then
      obrestocks
      If Not comprovar_materialexacte(palet, cadbl(form1.etmaterialexacte.tag)) Then MsgBox "Aquest material no es exactament el que demana el client.", vbCritical, "Error": GoTo fi
    End If
    afegir_labobinadentrada palet, bobina
    If palet < 120000 And palet > 0 Then
      If Not vestemfentfingerprint Then
        demanar_verificacio_espesoritractat palet, bobina
        imprimir_controlqualitatbobinaentrada palet, bobina, 0
      End If
    End If
 End If
 botoensenyarpacking.tag = ""
 bobinesent.UpdateRecord
fi:
 ratoli "normal"
End Sub
Sub demanar_verificacio_espesoritractat(palet As Double, bobina As Double)
   Dim rstb As Recordset
   Dim rstp As Recordset
   Dim espesor As Double
   Dim tractat As Boolean
   Dim correcte As Boolean
   Dim rstm As Recordset
   Dim sonmicres As Boolean
   Dim espesorcorrecte As Double
   Dim pregunta As String
   Dim colormat As String
   Dim vtotok As Boolean
   ratoli "normal"
   obrestocks
   Set rstb = bobines.Database.OpenRecordset("select * from bobinesentimp where id=" + atrim(cadbl(bobines.Recordset!id)) + " and palet=" + atrim(palet) + " and bobina=" + atrim(bobina))
   If Not rstb.EOF Then
      Set rstp = dbstocks.OpenRecordset("select codimatprognou,grmsm2,micres from palets where idpalet=" + atrim(palet))
      If rstp.EOF Then Exit Sub
      Set rstm = dbtmp.OpenRecordset("SELECT materials.codi, familiescolorants.descripcio FROM familiescolorants RIGHT JOIN materials ON familiescolorants.codi = materials.familiacol where materials.codi=" + atrim(rstp!codimatprognou))
      If rstm.EOF Then Exit Sub
      colormat = UCase(atrim(rstm!descripcio))
      
      If cadbl(rstp!micres) Then espesorcorrecte = cadbl(rstp!micres): sonmicres = True
      If cadbl(rstp!grmsm2) Then espesorcorrecte = degramsamicres(rstp!codimatprognou): sonmicres = True 'espesorcorrecte = cadbl(rstp!grmsm2):
      pregunta = "Entra el valor de l'espesor del micrometre." + Chr(10) + " +-10% de " + atrim(espesorcorrecte) + " Micres. VALOR MICROMETRE --> " + atrim(espesorcorrecte * 4)
tornarhi:
      carregarvalorsformbobentrada
      espesor = cadbl(formcomprovacionsbobentrada.cespessor)
      vtotok = True
      'If sonmicres Then
      correcte = ((espesor <= (espesorcorrecte + (espesorcorrecte * 10 / 100))) And (espesor >= (espesorcorrecte - (espesorcorrecte * 10 / 100))))
      '         Else
      '              espesor = Redondejar(espesor / 4, 1)
      '              correcte = ((espesor <= (espesorcorrecte + (espesorcorrecte * 10 / 100))) And (espesor >= (espesorcorrecte - (espesorcorrecte * 10 / 100))))
      ' End If
      tractat = IIf(formcomprovacionsbobentrada.cverificaciotractat.Value = 0, False, True)
      If Not correcte Then MsgBox "L'espessor no es el correcte.", vbCritical, "Atenció": vtotok = False
      If atrim(formcomprovacionsbobentrada.ccolormaterial) <> UCase(colormat) Then MsgBox "El color escullit no es el correcte.", vbCritical, "Atenció": vtotok = False
      If Not tractat Then MsgBox "Has de verificar el tractat.", vbCritical, "Atenció": vtotok = False
      If Not vtotok Then
         If MsgBox("Hi ha error de verificació vols revisar els valors?" + Chr(10) + " Si no els revises es guardaran amb aquests valors i podras continuar", vbCritical + vbYesNo + vbDefaultButton1, "Error") = vbYes Then GoTo tornarhi
      End If
      Unload formcomprovacionsbobentrada
      rstb.Edit
      'si son grm/m2 ho passo amb negatiu
      If Not sonmicres Then espesorcorrecte = espesorcorrecte * -1
      rstb!espesorteoric = espesorcorrecte
      rstb!verificacioespesor = IIf(sonmicres, espesor, espesor * -1)
      rstb!verificaciotractat = tractat
      rstb!verificaciocolor = Mid(colormat, 1, 15)
      rstb.Update
   End If
   ratoli "normal"
End Sub
Sub carregarvalorsformbobentrada()
    Dim rstm As Recordset
    Set rstm = dbtmp.OpenRecordset("select descripcio from familiescolorants where codi>499 order by descripcio")
    Load formcomprovacionsbobentrada
    While Not rstm.EOF
       formcomprovacionsbobentrada.ccolormaterial.AddItem UCase(rstm!descripcio)
       rstm.MoveNext
    Wend
    If Not vestemfentfingerprint Then formcomprovacionsbobentrada.Show 1
End Sub

Sub imprimir_controlqualitatbobinaentrada(palet As Double, bobina As Double, desb As Byte, Optional tempsdespera As Byte)
   Dim ultimalinia As String
   Dim esp As Double
  If bobinesent.Recordset.EOF Then Exit Sub
  bobinesent.Recordset.FindFirst "palet=" + atrim(palet) + " and bobina=" + atrim(bobina)
  llistat.DataFiles(0) = ""
   llistat.DataFiles(1) = ""
   ultimalinia = "Imp-" + atrim(nummaq) + " Op: " + atrim(numop) + " Comanda: " + atrim(comanda) + " Fecha: " + Format(Now, "dd/mm/yy")
   For i = 0 To 100
     llistat.Formulas(i) = ""
   Next i
   id = " +"
   
   
   esp = cadbl(bobinesent.Recordset!espesorteoric)
   llistat.Formulas(0) = "lot='" + atrim(palet) + "/" + atrim(bobina) + "'"
   llistat.Formulas(1) = "ultimalinia='" + atrim(ultimalinia) + "'"
   llistat.Formulas(2) = "caratractada='" + IIf(bobinesent.Recordset!verificaciotractat, "X", "") + " '"
   llistat.Formulas(3) = "espesormaterial='" + IIf(cadbl(bobinesent.Recordset!verificacioespesor) > 0, atrim(bobinesent.Recordset!verificacioespesor) + " Micres", atrim(cadbl(bobinesent.Recordset!verificacioespesor)) + " Grm/m2") + "'"
   llistat.Formulas(4) = "colormaterial='" + atrim(bobinesent.Recordset!verificaciocolor) + "'"
   llistat.Formulas(5) = "valorvalidespesor='Marge: >=" + atrim(Redondejar(esp - (esp / 10), 1)) + " i <=" + atrim(Redondejar(esp + (esp / 10), 1)) + "'"
   
   llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "verificacioqualitatlaminadoresbobinesentrada.rpt"
   llistat.Destination = crptToPrinter
    llistat.CopiesToPrinter = 1
   llistat.DiscardSavedData = True
' llistat.PrinterName = llegir_ini("Impressores", "nomfulla", "baixesimpressora.ini")
' llistat.PrinterPort = llegir_ini("Impressores", "portfulla", "baixesimpressora.ini")
' llistat.PrinterDriver = llegir_ini("Impressores", "driverfulla", "baixesimpressora.ini")
   DoEvents
   If tempsdespera = 0 Then tempsdespera = 2
   wait tempsdespera
   If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
   llistat.Action = 1
   MsgBox "ATENCIÓ CONTROL DE VERIFICACIÓ DE QUALITAT." + Chr(10) + "VERIFICA LA IMPRESIÓ AMB L'ETIQUETA IMPRESA", vbInformation, "VERIFICACIÓ QUALITAT"
   llistat.SelectionFormula = ""
   llistat.DataFiles(0) = ""

End Sub

Private Sub botollaunesreprint_Click()
 
   If estemfentreprint Then
         Framereprint.visible = True
         carregarllistadellaunesreprint
         Framereprint.Left = 45
         Framereprint.Top = 165
       Else
           Framereprint.visible = False
   End If
End Sub



Private Sub canvienfilada_DblClick()
If canvienfilada = "Si" Then
   canvienfilada = "No"
 Else: canvienfilada = "Si"
End If
End Sub

Private Sub cobservacionsoperari_DblClick()
Command33_Click
End Sub

Private Sub comanda_DropDown()
  'carregar_llista_ordremuntadora
  carregar_llista_ordreimpressio
End Sub
Sub carregar_llista_ordreimpressio(Optional vnumc As String, Optional vnoobrirla As Boolean)
  'Dim vnumc As String
  Dim vcomandaactual As Double
  Dim vcomandafingerprint As Double
  
  
  vcomandaactual = cadbl(comanda)
  vcomandafingerprint = vcomandaactual
  Load formordreimpresio
  formordreimpresio.Show
  formordreimpresio.reixa.row = 1
  formordreimpresio.reixa_refrescarfila

  While Screen.ActiveForm.Name = "formordreimpresio" Or Screen.ActiveForm.Name = "obsidtreball" Or Screen.ActiveForm.Name = "veurereport" Or Screen.ActiveForm.Name = "formcanvisanilox"
     vnumc = formordreimpresio.reixa.TextMatrix(formordreimpresio.reixa.row, 0)
     If Not IsNumeric(Mid(vnumc, 1, 1)) Then vnumc = Mid(vnumc, 2)
     If cadbl(vnumc) > 0 Then
       If cadbl(vnumc) <> vcomandaactual Then
          'carrego l'annex
           formannex.carregarcomanda cadbl(vnumc)
           vcomandaactual = vnumc
           DoEvents
           If isvisible("formordreimpresio") Then formordreimpresio.SetFocus
       End If
       If seleccioret = 5 Then
        imprimir_packinglistTICKET cadbl(vnumc), IIf(formordreimpresio.Command6.visible, True, False)
        seleccioret = 0
       ' vnumc = 0
       End If
     End If
     esperar 600
    DoEvents
  Wend
senseescullir:
  Unload veurereport
  If seleccioret = 1 Or seleccioret = 5 Or seleccioret = 2 Then
   vnumc = formordreimpresio.reixa.TextMatrix(formordreimpresio.reixa.row, 0)
   If Not IsNumeric(Mid(vnumc, 1, 1)) Then vnumc = Mid(vnumc, 2)
   If seleccioret = 5 Then
        imprimir_packinglistTICKET cadbl(vnumc)
        vnumc = 0
   End If
   If seleccioret = 1 Then
        vnopreguntar = False
        form1.BackColor = &H80000005
        vestemfentfingerprint = False
        Unload formseleccio
        formannex.carregarcomanda cadbl(comanda)
        'vnumc = cadbl(InputBox("Entra la comanda manualment.", "Comanda"))
   End If
   If seleccioret = 2 Then
        If MsgBox("Vols fer la comanda " + atrim(vcomandaactual) + " com a copia de la " + atrim(vcomandafingerprint) + "?", vbExclamation + vbDefaultButton2 + vbYesNo, "Finger Print") = vbYes Then
          vestemfentfingerprint = True
          form1.BackColor = &HFFFF&
          vnopreguntar = True
          Unload formseleccio
          vcomandaactual = cadbl(vnumc)
          formannex.carregarcomanda cadbl(vcomandaactual)
           Else: vnumc = 0
        End If
   End If
'     Else: vnumc = cadbl(InputBox("Entra la comanda manualment.", "Comanda"))
      Else:
           If seleccioret = 99 Or InStr(1, UCase(Environ("computername")), "IMPRESSORS") > 0 Or existeix("c:\ordprog.ini") Then
               vnumc = cadbl(InputBox("Entra la comanda manualment.", "Comanda"))
                Else: vnumc = 0
           End If
  End If
  Unload formordreimpresio
  Unload formcanvisanilox
  If cadbl(vnumc) = 0 Then formannex.carregarcomanda cadbl(comanda): Exit Sub
  If vnoobrirla Then GoTo fi
  comanda.text = vnumc
  mirar_missatgeXroperaris cadbl(vnumc)
  Command4_Click
  If vestemfentfingerprint Then
     ensenyar_formanilox_i_tancar
     copiarvalorsdelacomandaoriginal vcomandafingerprint, cadbl(vnumc)
     vnopreguntar = False
     formannex.carregarcomanda cadbl(comanda)
  End If
fi:
  
  Unload formordreimpresio
  
End Sub
Sub mirar_missatgeXroperaris(vnumc As Double)
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("select * from Impresores_ObsXoperaris where numcomanda=" + atrim(vnumc))
   If Not rst.EOF Then
       If atrim(rst!obscomanda) <> "" Then
         Load avis
         avis.caption = "MISSATGE PER L'OPERARI"
         avis.BackColor = &HFF80FF
         avis.missatge.BackColor = &HFF80FF
         avis.missatge.FontBold = True
         avis.missatge = rst!obscomanda
         avis.Show 1
       End If
   End If
   Set rst = Nothing
End Sub
Sub copiarvalorsdelacomandaoriginal(vcomandafingerprint As Double, vcomandadesti As Double)
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim i As Integer
  Dim vcomanda As Double
  Set rst = dbtmpb.OpenRecordset("select comandafingerprintoriginal from impressorestot where comanda=" + atrim(vcomandafingerprint))
  If Not rst.EOF Then
     vcomanda = cadbl(rst!comandafingerprintoriginal)
     If vcomanda = 0 Then
        rst.Edit
        rst!comandafingerprintoriginal = vcomandafingerprint
        rst.Update
     End If
  End If
  If vcomanda = 0 Then vcomanda = vcomandafingerprint
  
  dbtmpb.Execute "update impressorestot set comandafingerprintoriginal=" + atrim(vcomanda) + " where comanda=" + atrim(vcomandadesti)
'copia les llaunes gastades
  Set rst = dbtmpb.OpenRecordset("select * from impresores_llaunesgastades where comanda=" + atrim(vcomanda))
  Set rst2 = dbtmpb.OpenRecordset("select * from impresores_llaunesgastades where comanda=" + atrim(vcomandadesti))
  If rst2.EOF Then
        While Not rst.EOF
          rst2.AddNew
          For i = 0 To rst.Fields.Count - 1
               rst2.Fields(i) = rst.Fields(i)
          Next i
          rst2!comanda = vcomandadesti
          rst2.Update
          rst.MoveNext
        Wend
  End If
'copia el impresores pantones
  Set rst = dbtmpb.OpenRecordset("select * from impresorespantones where comanda=" + atrim(vcomanda))
  Set rst2 = dbtmpb.OpenRecordset("select * from impresorespantones where comanda=" + atrim(vcomandadesti))
  If rst2.EOF Then
      rst2.AddNew
      For i = 0 To rst.Fields.Count - 1
               rst2.Fields(i) = rst.Fields(i)
      Next i
      rst2!comanda = vcomandadesti
      rst2.Update
  End If
  
'copia els consums de tinta per cada anilox
  Set rst = dbtmpb.OpenRecordset("select * from impresores_aniloxos where comanda=" + atrim(vcomanda))
  Set rst2 = dbtmpb.OpenRecordset("select * from impresores_aniloxos where comanda=" + atrim(vcomandadesti))
  While Not rst2.EOF And Not rst.EOF
     rst.FindFirst "ordretinter=" + atrim(rst2!ordretinter)
     If Not rst.NoMatch Then
      rst2.Edit
      For i = 5 To rst.Fields.Count - 1
            rst2.Fields(i) = rst.Fields(i)
      Next i
      rst2.Update
     End If
     rst2.MoveNext
  Wend
'copio la linia de temps de canvi anilox
  Set rst = dbtmp.OpenRecordset("select * from aniloxtimeline where comanda=" + atrim(vcomanda))
  Set rst2 = dbtmp.OpenRecordset("select * from aniloxtimeline where comanda=" + atrim(vcomandadesti))
  If rst2.EOF Then
     If Not rst.EOF Then
      rst2.AddNew
      For i = 1 To rst.Fields.Count - 1
               rst2.Fields(i) = rst.Fields(i)
      Next i
      rst2!comanda = vcomandadesti
      rst2!nummaquina = nummaq
      rst2!numoperari = numop
      rst2!Data = Now
      rst2.Update
    End If
  End If
  Set rst = Nothing
  Set rst2 = Nothing
End Sub
Sub carregar_llista_ordremuntadora()
  Dim vnumc As String
  ratoli "espera"
  Unload formseleccio
  Load formseleccio
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "SELECT muntadoratot.comanda as Num_Comanda FROM comandes INNER JOIN muntadoratot ON comandes.comanda = muntadoratot.comanda WHERE (((muntadoratot.acabada)=True) AND ((comandes.proximaseccio)='I'));"
  formseleccio.caption = "Clixes Muntats"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 3000
  formseleccio.bimprimir.visible = True
  formseleccio.bimprimir.tag = "comprovarimpresio"
  formseleccio.Left = 1000
  formseleccio.Top = 1000
  ratoli "normal"
  formseleccio.Show
  While Screen.ActiveForm.Name = "formseleccio"
     If Not formseleccio.Data1.Recordset.EOF Then
       If cadbl(vnumc) <> cadbl(formseleccio.Data1.Recordset!num_comanda) Then
          'carrego l'annex
           formannex.carregarcomanda cadbl(formseleccio.Data1.Recordset!num_comanda)
           vnumc = cadbl(formseleccio.Data1.Recordset!num_comanda)
       End If
     End If
    DoEvents
  Wend
senseescullir:
  If seleccioret = 1 Or seleccioret = 5 Then
   vnumc = cadbl(formseleccio.Data1.Recordset!num_comanda)
   If seleccioret = 5 Then
        imprimir_packinglistTICKET cadbl(vnumc)
        vnumc = 0
   End If
     Else:
        Unload formseleccio
        formannex.carregarcomanda cadbl(comanda)
        vnumc = cadbl(InputBox("Entra la comanda manualment.", "Comanda"))
  End If
  Unload formseleccio
  If vnumc = 0 Then formannex.carregarcomanda cadbl(comanda): Exit Sub
  
  comanda.text = vnumc
  Command4_Click
End Sub
Function descripciogrup(vnumc As Double) As String
   Dim rst As Recordset
   Set rst = dbstocks.OpenRecordset("SELECT opcionsdajust.comanda, grupsdepalets.nomdelgrup fROM opcionsdajust INNER JOIN grupsdepalets ON opcionsdajust.grupdestoc = grupsdepalets.numerogrup WHERE (((opcionsdajust.comanda)=" + atrim(vnumc) + "));")
   If Not rst.EOF Then
       descripciogrup = atrim(rst!nomdelgrup)
   End If
   Set rst = Nothing
End Function

Sub prepararlesbobinespelllistat(vnumc As Double, vnommaterial As String)
   Dim rst As Recordset
   Dim vtipusbobina As String
   Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb", , True)
   Set rst = dbstocks.OpenRecordset("SELECT Palets.Idpalet, Bobines.Idbobina, materials.descripcio, Bobines.Numcomrev, Parcials.comanda FROM materials RIGHT JOIN (Palets LEFT JOIN (Bobines LEFT JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)) ON Palets.Idpalet = Bobines.Idpalet) ON materials.codi = Palets.codimatprognou WHERE (((Parcials.comanda)='" + atrim(vnumc) + "'));")
   If rst.EOF Then vnommaterial = descripciogrup(vnumc)
   While Not rst.EOF
      vtipusbobina = IIf(bobinesdentrada.esparcial(rst!idpalet, rst!idbobina), "P", "")
      vtipusbobina = IIf(bobinesdentrada.esrestu(rst!idpalet, rst!idbobina), "R", vtipusbobina)
      If vtipusbobina = "" Then vtipusbobina = "Q"
      rst.Edit
      rst!Numcomrev = vtipusbobina
      rst.Update
      If vnommaterial = "" Then vnommaterial = atrim(rst!descripcio)
      rst.MoveNext
   Wend
   Set rst = Nothing
   Unload bobinesdentrada
End Sub
Sub imprimir_packinglistTICKET(vnumc As Double, Optional vperpantalla As Boolean)
 
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vnummaq As Byte
  Dim vnommaterial As String
  Dim vnopackinglist As String

  prepararlesbobinespelllistat vnumc, vnommaterial
  If Not vperpantalla Then wait 2
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatPACKINGLIST_TICKET.rpt", 1) '"etiqueta_llaunes.rpt"
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "palets.mdb"
  oreport.RecordSelectionFormula = "{parcials.comanda}='" + atrim(vnumc) + "' "
  'oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
 ' oreport.Sections("D").ReportObjects.Item("serie2").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
  oreport.DiscardSavedData
  'oreport.Sections("D").ReportObjects.Item("recuperador").Suppress = True
  oreport.FormulaFields.GetItemByName("titol").text = "'PackingList:  " + atrim(vnumc) + "'"
  oreport.FormulaFields.GetItemByName("ajust").text = "'" + descripcioajustcomanda(cadbl(vnumc), vnopackinglist) + "'"
  oreport.FormulaFields.GetItemByName("tipusdematerial").text = "'" + vnommaterial + "'"
  'si no hi ha nom de material es que es estoc i despres passo la linia de descripció del grup d'estoc
  If InStr(1, vnopackinglist, "Estoc") > 0 And vnopackinglist <> "" Then oreport.FormulaFields.GetItemByName("nopackinglist").text = "'" + vnopackinglist + "'"
  If vperpantalla Then
        Load veurereport
        
        veurereport.caption = "Packinglist"
        veurereport.CRViewer.ReportSource = oreport
        veurereport.CRViewer.DisplayGroupTree = False
        veurereport.CRViewer.DisplayBorder = False
        veurereport.CRViewer.DisplayToolbar = False
        veurereport.width = 5500
        veurereport.Height = formordreimpresio.Height
        veurereport.CRViewer.ViewReport
        'veurereport.WindowState = 0
        veurereport.Show
        If UCase(arguments(1)) <> "DESBOBINADORS" Then
            veurereport.Top = formordreimpresio.Top
            veurereport.Left = formordreimpresio.Left + formordreimpresio.width + 100
            formordreimpresio.SetFocus
        End If
     Else
      oreport.DisplayProgressDialog = False
      oreport.PrintOut False, 1
   End If
  
End Sub
Function descripcioajustcomanda(vnumc As Double, vnopackinglist As String) As String
  Dim rsta As Recordset
  Dim db As Database
  Dim rst As Recordset
  
  Set db = OpenDatabase(rutadelfitxer(cami) + "palets.mdb", , True)
  Set rsta = db.OpenRecordset("select* from opcionsdajust where comanda=" + atrim(cadbl(vnumc)))
   'possar ajustteoric
   If Not rsta.EOF Then
    vnopackinglist = atrim(rsta!grupdestoc)
    If rsta!sistemadajust = 1 Then descripcioajustcomanda = "Llençar: " + atrim(rsta!mtrsajust) + "m"
    If rsta!sistemadajust = 2 Then descripcioajustcomanda = "Estoc: " + atrim(rsta!mtrsajust) + "m"
    If rsta!sistemadajust = 3 Then descripcioajustcomanda = atrim(rsta!paletajust) + "/" + atrim(rsta!bobinaajust) + " " + atrim(rsta!mtrsajust) + "m"
   End If
   If vnopackinglist <> "" Then
     Set rst = dbtmp.OpenRecordset("select cantitatex from comandes where comanda=" + atrim(vnumc), , ReadOnly)
     If Not rst.EOF Then vnopackinglist = "Estoc " + vnopackinglist + " --> " + atrim(rst!cantitatex) + " Mtrs"
   End If
   Set rst = Nothing
   Set rsta = Nothing
  Set db = Nothing
End Function

Private Sub comanda_KeyDown(KeyCode As Integer, Shift As Integer)
  
  KeyCode = 0
End Sub

Private Sub comanda_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
  Exit Sub
  If KeyAscii = 13 Then Command4_Click
End Sub

Private Sub comanda_LostFocus()
'   escriure_ini "Baixes", "ultimacomanda", comanda, "comandes.ini"
   'Command4_Click
   cpostitcomanda.visible = False
End Sub

Private Sub comescanerbascula_OnComm()

End Sub

Private Sub Command1_Click()
  If Not comprovamaq Then Exit Sub
  If comprovarsidescansorelleu Then Exit Sub
 If Not impresores.Recordset.EOF Then
  impresores.Recordset.MoveLast
  If impresores.Recordset!tipus = "M" Then
      numop = escullir_operari
      nomoperari = UCase(r)
      numop2 = escullir_operari("Escullir AJUDANT d'OPERARI")
      nomoperari2 = "Ajudant: " + UCase(r)
  End If
 End If
 crearseccio "M"
 imprimir_etiquetallaunesdetintapercomanda
End Sub
Function comprovarsidescansorelleu() As Boolean
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("select * from controldescansrelleu where (hores=0 or hores=null) and operari=" + atrim(numop) + " and seccio='" + atrim(lletraseccio) + "'")
  If rst.EOF Then Exit Function
  comprovarsidescansorelleu = True
  MsgBox UCase(nomoperari) + " en aquest moment està fent " + atrim(rst!tipus) + Chr(10) + "Primer dona per acabada la incidència.", vbExclamation, "Atenció"
End Function

Private Sub Command10_Click()
 If llegir_ini("Baixes", "programaamaquina", fitxerini) = 0 Then MsgBox "No pots ACABAR COMANDA si no estàs a màquina." + Chr(10) + "Si vols imprimir-la al costat hi ha el botó d'imprimir.", vbCritical, "Atenció": Exit Sub
 client.ToolTipText = client.caption
' If Not existeix(nomordinadorcomexi) Then
'    MsgBox "No es pot accedir a l'ordinador de Comexi." + Chr(10) + "No es guardarà el fitxer de temperatures.", vbCritical + vbOKOnly, "Atenció"
'    mantenimentbobina.passaravis 0, 0, "Fitxer Temperatures", comanda, "No s'ha pogut accedir a l'ordinador de comexi i no s'ha generat el fitxer de temperatures. TORNEU A PROVAR-HO MES TARD."
'    GoTo sensefitxertemperatures
' End If
 'If buscarcomandaacomexi(comanda) = "" Then
 '  If UCase(InputBox("No localitzo el fitxer de TEMPERATURES de COMEXI" + Chr(10) + "ABANS DE GUARDAR AQUESTA COMANDA HAS DE GUARDAR LA DE COMEXI." + Chr(10) + "Si vols continuar igualment escriu [continuar]?", "ATENCIÓ")) <> "CONTINUAR" Then Exit Sub
 '  generarfitxernoudetemperatures comanda
 '  mantenimentbobina.passaravis 0, 0, "Fitxer Temperatures", comanda, "No s'ha trobat el fitxer i s'ha generat un automàticament. Reviseu que tot sigui correcte."
 '   Else:
 '     resposta = guardar_fitxer_temperatures(cadbl(comanda))
 '     If resposta <> "" Then
 '       MsgBox resposta, vbCritical, "Error de temperatures impresora Comexi"
 '       mantenimentbobina.passaravis 0, 0, "Fitxer Temperatures", comanda, atrim(resposta)
 '       generarfitxernoudetemperatures comanda
 '     End If
 'End If
'sensefitxertemperatures:
 impresores.Recordset.MoveLast
 If Not impresores.Recordset.EOF Then
    If impresores.Recordset!tipus = "A" Then
      If cadbl(impresores.Recordset!mtrsprova) = 0 Then
         If MsgBox("No hi ha els metres d'ajust entrats vols entrar-los ara?", vbCritical + vbYesNo + vbDefaultButton1, "Metres de prova") = vbYes Then demanar_metres_dajust
      End If
    End If
    If Not IsDate(impresores.Recordset!datafi) And Not IsDate(impresores.Recordset!horafi) Then
        impresores.Recordset.Edit
        impresores.Recordset!datafi = Date
        impresores.Recordset!horafi = Time
        impresores.Recordset.Update
    End If
End If
passarcomanda_a_noacabada cadbl(comanda)
Command8_Click
imprimir_full_arrancar_rentar cadbl(comanda)
imprimirparcialsajustsinohihafuncionament
If command15.tag = "Error" Then Exit Sub
r = cadbl(InputBox("Entra la nova comanda", "Canvi de comanda"))
If cadbl(r) > 0 Then comanda.text = atrim(cadbl(r))
Command4_Click
End Sub
Sub passarcomanda_a_noacabada(vnumc As Double)
   dbtmp.Execute "update comandes set proximaseccio='I' where comanda=" + atrim(vnumc)
   dbtmp.Execute "update comandes set seccioactual='I' where comanda=" + atrim(vnumc)
   dbtmpb.Execute "update impressorestot set acavada=0 where comanda=" + atrim(vnumc)
   comandaacavada.Value = 0
End Sub
Sub imprimirparcialsajustsinohihafuncionament()
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("select * from impressores where comanda=" + atrim(comanda) + " order by datainici,horainici")
  If rst.EOF Then Exit Sub
  rst.MoveLast
  While Not rst.BOF
     If rst!tipus = "F" Then GoTo cont
     If rst!tipus = "A" Then GoTo cont
     rst.MovePrevious
  Wend
  If rst.BOF Then Exit Sub
cont:
 If rst!tipus = "A" Then
    obrestocks
    If cadbl(rst!paletprova) > 0 And cadbl(rst!bobinaprova) > 0 And cadbl(rst!paletprova) <> 11111 Then
        bobinesdentrada.imprimir_bobinaparcial cadbl(rst!paletprova), cadbl(rst!bobinaprova), , 1
    End If
    If cadbl(rst!paletprova2) > 0 And cadbl(rst!bobinaprova2) > 0 And cadbl(rst!paletprova2) <> 11111 Then
        bobinesdentrada.imprimir_bobinaparcial cadbl(rst!paletprova2), cadbl(rst!bobinaprova2), , 1
    End If
 End If
End Sub
Private Sub Command11_Click()
calcular_totals
End Sub

Private Sub Command12_Click()
If bobines.Recordset.EOF Then
   MsgBox "No hi ha bobina creada"
  Else
    frameempalmes.visible = Not frameempalmes.visible
    framepantones.visible = False
    framebobentrada.visible = False
End If
End Sub

Private Sub Command13_Click()
' If bobines.Recordset.EOF Then
'   MsgBox "No hi ha bobina creada"
'  Else
    framebobentrada.visible = Not framebobentrada.visible
    framepantones.visible = False
    frameempalmes.visible = False
 'End If
End Sub

Private Sub Command14_Click()
  Dim r2 As Double
  If Not comprovamaq Then Exit Sub
  If impresores.Recordset!tipus = "A" And atrim(impresores.Recordset!paletbobprova) <> "" And impresores.Recordset!paletbobprova <> "0-0/0-0" Then
      MsgBox "No pots eliminar una linia d'ajust... que ja tingui bobines d'ajust posades" + Chr(10) + Chr(13) + "Si has de fer un canvi notifica-ho a oficines.", vbCritical + vbOKOnly, "Atenció"
      Exit Sub
  End If
If MsgBox("Eliminar aquesta pot suposar eliminar informació de bobines.", vbCritical + vbYesNo, "Atenció") = vbYes Then
     'reixa_BeforeDelete 0
     If MsgBox("Segur que vols borrar aquesta linia i tot el seu contingut?", vbYesNo, "Atenció") = vbNo Then Cancel = 1
    If Cancel <> 1 Then
    r = 0
    r2 = cadbl(impresores.Recordset!id)
    If atrim(impresores.Recordset!tipus) = "F" Then r = atrim(cadbl(impresores.Recordset!id))
    dbtmpb.Execute "delete * from bobinesimp where controlid=" + r
'    dbtmpb.Execute "delete * from impressores where id=" + atrim(r2)
    impresores.Recordset.Delete
    Command4_Click
    If Not impresores.Recordset.EOF And Not impresores.Recordset.BOF Then impresores.Recordset.MoveLast
  End If
End If
End Sub


Sub mirar_bobinesdentrada_noacavades()
 Dim metres As Double
 Dim metresant As Double
 Dim palet As Double
 Dim bobina As Double
 Dim rstconsulta2 As Recordset
 Dim rst As Recordset
 
 noespota0 = True
   carregar_bobinesdentrada "carregarbobinesnoutilitzades", , , , cadbl(comanda)
   If Not rstconsulta.EOF Or Not rstconsulta.BOF Then rstconsulta.MoveFirst
   Set rstconsulta2 = rstconsulta.Clone
   mantenimentbobina.checknoimprimirparcial = 1
   While Not rstconsulta2.EOF
      palet = rstconsulta2!idpalet
      bobina = rstconsulta2!idbobina
      PoB = IIf(rstconsulta2!taula = "parcials", "p", "b")
            
      If palet > 0 And bobina > 0 And UCase(PoB) = "P" Then 'atrim(rstconsulta2!tipus) >= "O"
           'demanar_final_palet_bobina_stock palet, bobina
           estatdelabobina palet, bobina, 0, ncomanda
           'bobinesdentrada.imprimir_bobinaparcial palet, bobina
           'Set rst = dbtmpb.OpenRecordset("select * from parcials where idbobina=" + atrim(bobina) + " and idpalet=" + atrim(palet) + " and comanda='" + atrim(ncomanda) + "'")
           'metres = 0
           'If Not rst.EOF Then metres = cadbl(rst!metres)
           'If metres <> 0 Then
           '   metres = bobinesdentrada.calcular_mtrsdispreals(palet, bobina, True)
           '   If metres > 500 Then passar_bobina_pendentverificardiametre palet, bobina
           'End If
           passar_bobina_pendentverificardiametre palet, bobina
           Set rst = Nothing
         Else
            'es una bobina feta a inplacsa
              If UCase(PoB) = "B" Then
                  carregar_bobinesdentrada "marcarutilitzadademanar", , palet, bobina, cadbl(comanda), True
              End If
      End If
      rstconsulta2.MoveNext
   Wend
   comprovar_fi_bobsent cadbl(comanda)
   Set rstconsulta2 = Nothing
   mantenimentbobina.checknoimprimirparcial = 0
   Unload mantenimentbobina
   noespota0 = False
End Sub
Sub passar_bobina_pendentverificardiametre(vpalet As Double, vbobina As Double)
      Dim vvalues As String
   vvalues = "(" + atrim(vpalet) + "," + atrim(vbobina) + ",'I'," + atrim(nummaq) + ")"
   dbtmpb.Execute "delete * from bobines_pendent_revisar_diametre where palet=" + atrim(vpalet) + " and bobina=" + atrim(vbobina) + " and seccio='I' "
   dbtmpb.Execute "insert into bobines_pendent_revisar_diametre (palet,bobina,seccio,maquina) values " + vvalues
End Sub
Sub comprovar_bobinesnoacabadesAvisarpernoemportarseles(numc As Double)
   Dim rstpar As Recordset
   Dim rst As Recordset
   Dim vmsg As String
   Dim rstordre As Recordset
   Set rstordre = dbtmpb.OpenRecordset("select * from impresores_ordreimpresio")
   Set rstpar = dbstocks.OpenRecordset("select * from parcials where orcomassignacio<>'500' and comanda='" + atrim(numc) + "'")
   While Not rstpar.EOF
     Set rst = dbstocks.OpenRecordset("select * from parcials where data=null and idpalet=" + atrim(rstpar!idpalet) + " and idbobina=" + atrim(rstpar!idbobina))
     While Not rst.EOF
        rstordre.FindFirst "comanda=" + atrim(cadbl(rst!comanda))
        If Not rstordre.NoMatch Then vmsg = vmsg + " " + atrim(rst!idpalet) + "/" + atrim(rst!idbobina) + " [" + atrim(rst!comanda) + "]" + Chr(13) + Chr(10)
        rst.MoveNext
     Wend
     rstpar.MoveNext
   Wend
   Set rstpar = Nothing
   If vmsg <> "" Then
      MsgBox "Les bobines seguents no s'han de moure de impresores perquè s'utilitzaran en altres comandes." + Chr(13) + vmsg
      imprimir_ticket_bobines vmsg
      
   End If
End Sub
Sub imprimir_ticket_bobines(vmsg As String)

 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vlinia As String
  Dim vtotallinia As Double
  Exit Sub
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "ticket_impresio.rpt", 1)
  oreport.FormulaFields.GetItemByName("l1").text = "'BOBINES UTILITZADES'"
  oreport.FormulaFields.GetItemByName("l1.2").text = "'PER ALTRES COMANDES'"
  oreport.FormulaFields.GetItemByName("l1.3").text = "'-------------------'"
  'oreport.FormulaFields.GetItemByName("l2").Text = "'Kg a recuperar: " + atrim(kgxrecuperar(0)) + "'"
  oreport.FormulaFields.GetItemByName("l4").text = "'" + vmsg + "'"
  
  
  oreport.DiscardSavedData
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
    Else
      oreport.DisplayProgressDialog = False
      oreport.PrintOut False, 1
  End If
 


End Sub
Sub comprovar_fi_bobsent(numc As Double)
 Dim rstbobent As Recordset
 Dim rstpar As Recordset
 Dim palet As Double
 Dim bobina As Double
 Set rstbobent = dbtmpb.OpenRecordset("SELECT bobinesentimp.palet, bobinesentimp.bobina, impressores.comanda FROM (bobinesentimp INNER JOIN bobinesimp ON bobinesentimp.id = bobinesimp.Id) INNER JOIN impressores ON bobinesimp.controlid = impressores.Id WHERE (((impressores.comanda)=" + atrim(numc) + "));")
 While Not rstbobent.EOF
    palet = rstbobent!palet
    bobina = rstbobent!bobina
    Set rstpar = dbstocks.OpenRecordset("select * from parcials where orcomassignacio<>'500' and comanda='" + atrim(numc) + "' and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
    'If rstpar.EOF Then Set rstpar = dbstocks.OpenRecordset("select * from parcials where cadbl(comanda)>2000 and cadbl(comanda)<3000 and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
    If rstpar.EOF Then
      estatdelabobina palet, bobina, 0, numc
      'bobinesdentrada.imprimir_bobinaparcial palet, bobina
    End If
    rstbobent.MoveNext
 Wend
Set rstpar = Nothing
Set rstbobent = Nothing

End Sub
Sub passar_avis_a_produccio(comanda As String, metrestotal As Double, metrescomanda As Double, op As Double)
  Dim r As String
    r = App.Path + "\aviscampsmodificats.txt"
    If Not existeix(r) Then
              Open r For Output As 1
          Else: Open r For Append As 1
    End If
    Print #1, Trim(Now) + " MassaMtrs Com: " + comanda + " Mtrs: " + atrim(metrescomanda) + " MtrsFets: " + atrim(metrestotal) + " Op: " + Trim(op)
    Close 1
End Sub

Sub comprovar_quenoespassindemetres(vdiferenciareal_teoric As Double)
   Dim totalmetres As Double
   Dim metresprova As Double
   Dim metrescomanda As Double
   Dim mtrsajust As Double
   Dim metresgastats As Double
   Dim stock As Boolean
   Dim sistemaajust As Byte
   Dim marge As Double
   Dim rsta As Recordset
   Dim rstb As Recordset
   Dim rstc As Recordset
   Dim rsts As Recordset
   Dim rsto As Recordset
   Set rstb = dbtmpb.OpenRecordset("select sum(mtrsprova) as prova, max(operari) as op from impressores where comanda=" + atrim(cadbl(comanda)))
   'Set rstc = dbtmp.OpenRecordset("select cantitatex from comandes where comanda=" + atrim(cadbl(comanda)))
   Set rsta = dbstocks.OpenRecordset("select mtrsajust,sistemadajust from opcionsdajust where comanda=" + atrim(cadbl(comanda)))
   Set rsts = dbtmp.OpenRecordset("SELECT comandes_extres.assignarstock as estoc, comandes_extres.mtrsassignatsestock as mtrsestoc frOM comandes_extres WHERE comanda=" + atrim(cadbl(comanda)) + ";")
   Set rstc = dbstocks.OpenRecordset("select sum(metres) as summetres from historic_packinglist where comanda='" + atrim(comanda) + "'")
   
   metrescomanda = 0
   If Not rstc.EOF Then metrescomanda = cadbl(rstc!summetres)
   If Not rsts.EOF Then
      stock = rsts!estoc
      If cadbl(rsts!mtrsestoc) = 0 And stock Then Exit Sub
      If stock Then metrescomanda = cadbl(rsts!mtrsestoc)
   End If
   sistemaajust = 0
   If Not rsta.EOF Then sistemaajust = rsta!sistemadajust
   If Not rstb.EOF And Not rstc.EOF Then
      mtrsajust = 1000
      If sistemaajust = 1 Then
        mtrsajust = rsta!mtrsajust
      End If
      'If Not rsta.EOF Then mtrsajust = rsta!mtrsajust
      metresprova = cadbl(rstb!prova) + vmetresarrancada
      marge = (metrescomanda * 1.1) - metrescomanda '10% de marge... aixo ara mateix fem 2000 mtrs rodons mes avall assigno
      totalmetres = cadbl(tmetres) + metresprova + cadbl(mtrsdolents) 'els metres dolents els he sumat perquè l'alicia ha dit que s'han de sumar
      If stock Then
        marge = 0: metrescomanda = metrescomanda + mtrsajust ': mtrsajust = 0
        If Not rsta.EOF Then
           If sistemaajust = 2 Then
             mtrsajust = rsta!mtrsajust
             metrescomanda = metrescomanda + mtrsajust
           End If
        End If
      End If
       'If totalmetres > (metrescomanda + mtrsajust + marge) Then
       marge = 2000 'estava a 1000 pero l'alicia m'ho ha fet passar a 2000
      metresgastats = metrescomanda + marge
      If sistemaajust = 1 Then metresgastats = metresgastats + metresdajustgastats(cadbl(comanda))
      If totalmetres > metresgastats Then
          MsgBox "Hi ha hagut massa metres gastats en aquesta comanda" + Chr(10) + Chr(13) + "S'havien assignat " + IIf(stock, "d'ESTOC ", "") + atrim(metresgastats) + " metres i s'han consumit " + atrim(totalmetres) + " metres" + Chr(10) + Chr(13) + "  Es passarà un avís al departament de Producció.", vbCritical + vbOKOnly, "Avís d'accés de metres"
          ' passar_avis_a_produccio comanda, totalmetres, metrescomanda, cadbl(rstb!op)
          mantenimentbobina.passaravis 0, 0, "Hi ha hagut massa metres gastats en aquesta comanda. S'havien assignat " + IIf(stock, "d'ESTOC ", "") + atrim(metrescomanda) + " metres i s'han consumit " + atrim(totalmetres) + " metres", comanda, "DIF. REAL-TEORIC: " + atrim(vdiferenciareal_teoric) + " Mts."
      End If
   End If
   Set rstb = Nothing
   Set rstc = Nothing
   Set rsta = Nothing
End Sub
Function metresdajustgastats(numc As Double) As Double
  Dim rstb As Recordset
  Set rstb = dbtmpb.OpenRecordset("select sum(mtrsprova) as ajust11111 from impressores where (paletprova=11111 or paletprova2=11111) and comanda=" + atrim(numc))
  If Not rstb.EOF Then
     metresdajustgastats = cadbl(rstb!ajust11111)
  End If
  
End Function
Sub comprovar_quenoespassindemetres_vell()
   Dim totalmetres As Double
   Dim metresprova As Double
   Dim metrescomanda As Double
   Dim mtrsajust As Double
   Dim stock As Boolean
   Dim sistemaajust As Byte
   Dim marge As Double
   Dim rsta As Recordset
   Dim rstb As Recordset
   Dim rstc As Recordset
   Dim rsts As Recordset
   Dim rsto As Recordset
   Set rstb = dbtmpb.OpenRecordset("select sum(mtrsprova) as prova, max(operari) as op from impressores where comanda=" + atrim(cadbl(comanda)))
   'Set rstc = dbtmp.OpenRecordset("select cantitatex from comandes where comanda=" + atrim(cadbl(comanda)))
   Set rsta = dbstocks.OpenRecordset("select mtrsajust,sistemadajust from opcionsdajust where comanda=" + atrim(cadbl(comanda)))
   Set rsts = dbtmp.OpenRecordset("SELECT comandes_extres.assignarstock as estoc, comandes_extres.mtrsassignatsestock as mtrsestoc frOM comandes_extres WHERE comanda=" + atrim(cadbl(comanda)) + ";")
   Set rstc = dbstocks.OpenRecordset("select sum(metres) as summetres from historic_packinglist where comanda='" + atrim(comanda) + "'")
   
      
   
   metrescomanda = 0
   If Not rstc.EOF Then metrescomanda = cadbl(rstc!summetres)
   If Not rsts.EOF Then
      stock = rsts!estoc
      If cadbl(rsts!mtrsestoc) = 0 And stock Then Exit Sub
      metrescomanda = cadbl(rsts!mtrsestoc)
   End If
   sistemaajust = 0
   If Not rsta.EOF Then sistemaajust = rsta!sistemadajust
   If Not rstb.EOF And Not rstc.EOF Then
      mtrsajust = 1000
      If sistemaajust = 1 Then
        mtrsajust = rsta!mtrsajust
        metrescomanda = metrescomanda + mtrsajust
      End If
      If Not rsta.EOF Then mtrsajust = rsta!mtrsajust
      metresprova = cadbl(rstb!prova)
      marge = (metrescomanda * 1.1) - metrescomanda '10% de marge
      totalmetres = cadbl(tmetres) + metresprova
      If stock Then
        marge = 0: metrescomanda = cadbl(rsts!mtrsestoc) + mtrsajust ': mtrsajust = 0
        If Not rsta.EOF Then
           If sistemaajust = 2 Then
             mtrsajust = rsta!mtrsajust
             metrescomanda = metrescomanda + mtrsajust
           End If
        End If
      End If
       'If totalmetres > (metrescomanda + mtrsajust + marge) Then
       marge = 1000
      
      If totalmetres > (metrescomanda + marge) Then
          MsgBox "Hi ha hagut massa metres gastats en aquesta comanda" + Chr(10) + Chr(13) + "S'havien assignat " + IIf(stock, "d'ESTOC ", "") + atrim(metrescomanda) + " metres i s'han consumit " + atrim(totalmetres) + " metres" + Chr(10) + Chr(13) + "  Es passarà un avís al departament de Producció.", vbCritical + vbOKOnly, "Avís d'accés de metres"
          ' passar_avis_a_produccio comanda, totalmetres, metrescomanda, cadbl(rstb!op)
          mantenimentbobina.passaravis 0, 0, "Hi ha hagut massa metres gastats en aquesta comanda. S'havien assignat " + IIf(stock, "d'ESTOC ", "") + atrim(metrescomanda) + " metres i s'han consumit " + atrim(totalmetres) + " metres", comanda
      End If
   End If
Set rstb = Nothing
   Set rstc = Nothing
   Set rsta = Nothing
   
End Sub
Sub des_reservar(numc As String)
   Dim rstdesr As Recordset
   Dim rstres As Recordset
   
   Dim msgdesr As String
   Set rstdesr = dbstocks.OpenRecordset("SELECT percomandaoclient.numcomanda, percomandaoclient.numclient, Reserves.Ample,reserves.idreserva, percomandaoclient.metres,percomandaoclient.idcompra FROM Reserves INNER JOIN percomandaoclient ON Reserves.idreserva = percomandaoclient.idreserva WHERE (((percomandaoclient.numcomanda)=" + atrim(cadbl(numc)) + "));")
   If rstdesr.EOF Then
      Exit Sub
   End If
   While Not rstdesr.EOF
       dbstocks.Execute "update reserves set metresreservats=metresreservats-" + atrim(cadbl(rstdesr!metres)) + " where idreserva=" + atrim(cadbl(rstdesr!idreserva))
       
       rstdesr.MoveNext
   Wend
   dbstocks.Execute "delete * from percomandaoclient where numcomanda=" + atrim(cadbl(numc)) '+ " and (idcompra<1 or idcompra=null)"
End Sub

Sub des_reservar2(numc As String)
   Dim rstdesr As Recordset
   Dim rstres As Recordset
   Dim msgdesr As String
   Set rstdesr = dbstocks.OpenRecordset("SELECT percomandaoclient.numcomanda, percomandaoclient.numclient, Reserves.Ample,reserves.idreserva, percomandaoclient.metres,percomandaoclient.idcompra FROM Reserves INNER JOIN percomandaoclient ON Reserves.idreserva = percomandaoclient.idreserva WHERE (((percomandaoclient.numcomanda)=" + atrim(cadbl(numc)) + "));")
   If rstdesr.EOF Then
     If r <> "nopregunta" And Not vnopreguntar Then
      MsgBox "No hi ha cap reserva d'aquesta comanda."
     End If
      Exit Sub
   End If
   While Not rstdesr.EOF
      msgdesr = msgdesr + "Ample: " + Redondejar(cadbl(rstdesr!ample), 1) + " cm <---> " + Redondejar(cadbl(rstdesr!metres), 0) + " Mtrs" + Chr(13) + Chr(10)
      rstdesr.MoveNext
   Wend
   If r <> "Info" And r <> "nopregunta" And Not vnopreguntar Then
     r = InputBox("Entra el numero de comanda per des-reservar, ha de coincidir amb la consulta." + Chr(13) + Chr(10) + msgdesr, "Comfirmació Des-Reservar")
     msgdesr = ""
   End If
   If cadbl(r) = cadbl(numc) Or r = "Info" Or r = "nopregunta" Then
      's = "(select distinct(idreserva) from percomandaoclient where numcomanda=" + comanda + ")"
      'Set rstdesr = dbtmp.OpenRecordset("SELECT percomandaoclient.numcomanda, percomandaoclient.numclient, Reserves.Ample,reserves.idreserva, percomandaoclient.metres,percomandaoclient.idcompra FROM Reserves INNER JOIN percomandaoclient ON Reserves.idreserva = percomandaoclient.idreserva WHERE (((percomandaoclient.numcomanda)=" + atrim(cadbl(comanda)) + ") and percomandaoclient.idreserva in " + s + ");")
      rstdesr.MoveFirst
      msgdesr = msgdesr + "Compres afectades: " + Chr(10) + Chr(13)
      While Not rstdesr.EOF
       Set rstres = dbstocks.OpenRecordset("select * from compresmaterial where not entregada and idreserva=" + atrim(cadbl(rstdesr!idreserva)))
       While Not rstres.EOF
          msgdesr = msgdesr + atrim(rstres!codimat) + "-" + atrim(rstres!descmat) + "    --->  NºCompra: " + atrim(rstres!numcompra) + Chr(13) + Chr(10)
          rstres.MoveNext
       Wend
       If r <> "Info" And cadbl(rstdesr!idcompra) < 1 And cadbl(rstdesr!numclient) = 0 And cadbl(rstdesr!numcomanda) > 0 Then
         dbstocks.Execute "update reserves set metresreservats=metresreservats-" + atrim(cadbl(rstdesr!metres)) + " where idreserva=" + atrim(cadbl(rstdesr!idreserva))
         'ho desabilito ja k ho faig mes endavant ... dbstocks.Execute "delete * from percomandaoclient where numcomanda=" + atrim(numc)
       End If
       rstdesr.MoveNext
     Wend
     
      
     If r <> "Info" Then
         dbstocks.Execute "delete * from percomandaoclient where numcomanda=" + atrim(cadbl(numc)) '+ " and (idcompra<1 or idcompra=null)"
     End If
     
     If msgdesr <> "" And r <> "nopregunta" And Not vnopreguntar Then MsgBox "Aquestes compres queden afectades per la Des-Reserva." + Chr(10) + Chr(13) + msgdesr
   End If
End Sub

Sub mirar_excesdemetresdeprova()
  Dim rstprova As Recordset
  Dim explicacio As String
  Set rstprova = dbtmpb.OpenRecordset("select * from toleranciesmaquina where tintes=" + atrim(ntintes))
  
  If Not rstprova.EOF Then
      If cadbl(tprova) > (cadbl(rstprova!metresajust) + cadbl(rstprova!metrestolerancia)) Then
        explicacio = InputBox("Vols donar una explicació del perquè s'han gastat mes metres de prova del teòrics?", "Metres de prova teòrics")
        mantenimentbobina.passaravis 0, 0, "Tolerancia de metres de prova superada", comanda, explicacio
      End If
  End If
End Sub



Function metresfetsinferiorsacomanda(numc As Double) As Boolean
   Dim metresc As Double
   If cadbl(tmetres) < (cadbl(metrescomanda) - ((cadbl(metrescomanda) / 100) * 4)) Then
          If UCase(InputBox("Aquesta comanda es de " + tmetres.tag + " metres i tu has fet " + tmetres + " metres" + Chr(10) + "PASSARÉ LA COMANDA A NO ACABADA. ESCRIU ACABADA SI ESTÀ REALMENT ACABADA", "ATENCIÓ")) = "ACABADA" Then
              metresfetsinferiorsacomanda = False
               Else: metresfetsinferiorsacomanda = True
          End If
   End If
End Function

Function esunreprint() As Boolean
  Dim rst As Recordset
  Set rst = impresores.Recordset.Clone
  If rst.EOF Then Exit Function
  rst.MoveFirst
  While Not rst.EOF
    If rst!numeromaquina <> nummaq Then esunreprint = True
    rst.MoveNext
  Wend
End Function
Function comprovar_consums_tintes(numc As Double) As Boolean
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim v As Double
  Set rst = dbtmpb.OpenRecordset("select * from impresores_aniloxos where comanda=" + atrim(numc) + " order by ordretinter_original")
  Set rst2 = dbtmpb.OpenRecordset("select * from impresorespantones where comanda=" + atrim(numc))
  rst2.Edit
  While Not rst.EOF
     v = 0
     If cadbl(rst!kgconsumits) = 0 And atrim(rst!tinta_comanda) <> "" Then
        While v = 0
         v = cadbl(InputBox("Entra els consum per la tinta " + UCase(atrim(rst!tinta_comanda)), "Kg gastats de tinta"))
        Wend
        rst.Edit
        rst!kgconsumits = v
        rst.Update
     End If
     If atrim(rst!tinta_comanda) <> "" Then
        If atrim(rst2.Fields("lot" + atrim(rst!ordretinter))) = "" Then
          If MsgBox("Hi ha un tinter que no te llauna escanejada." + Chr(10) + "VOLS CONTINUAR O VOLS CANCELAR I ANAR A POSSAR-LA?", vbCritical + vbDefaultButton2 + vbOKCancel, "Falta un tinter") = vbOK Then
             enviaremailgeneric "missatgesgenericsimpresores", "Comanda " + atrim(numc) + " sense Llauna escanejada.  " + nommaq + " - " + nomoperari, treure_apostruf("Hi ha un tinter sense escanejar.")
               Else: comprovar_consums_tintes = True: Exit Function
          End If
        End If
     End If
     rst2.Fields("kg" + atrim(rst!ordretinter)) = cadbl(rst!kgconsumits)
     rst.MoveNext
  Wend
  rst2.Update
  Set rst2 = Nothing
  Set rst = Nothing
End Function
Function buscardadesbasiquescomanda(vnumc As String) As String
  Dim rstc As Recordset
  Set rstc = dbtmp.OpenRecordset("SELECT comandes.*, clients.nom FROM comandes INNER JOIN clients ON comandes.client = clients.codi where comanda = " + atrim(cadbl(vnumc)))
  If rstc.EOF Then Exit Function
  buscardadesbasiquescomanda = "Nom: " + atrim(rstc!client) + " - " + atrim(rstc!nom) + Chr(13) + Chr(10) + "Ref Client: " + atrim(rstc!refclient) + Chr(13) + Chr(10) + "Texte Imp.: " + atrim(rstc!marcailinia)
End Function
Private Sub Command15_Click()
  Dim hores As Double
  Dim resposta As String
  Dim vdiferenciareal_teoric As Double
  Dim vdiferenciaFulla_real As Double
  
  
  vtotalmetres = 0
  Me.caption = "Tancant comanda... "
  If llegir_ini("Baixes", "programaamaquina", fitxerini) = 0 Then MsgBox "No pots ACABAR COMANDA si no estàs a màquina." + Chr(10) + "Si vols imprimir-la al costat hi ha el botó d'imprimir.", vbCritical, "Atenció": Exit Sub
  If comprovarsilesbobinessoncorrelatives Then Exit Sub
  verificarsihihaentratelcanvidaniloxos
  carregarllistadellaunesreprint
  If estemfentreprint Then
    If llistallaunesreprint.ListCount = 0 Or cadbl(ettotalkgreprint) = 0 Then MsgBox "Estas fent REPRINT i no has possat les llaunes utilitzades o el TOTAL de KG.", vbCritical, "Error": Command26_Click: Exit Sub
      Else: ensenyar_formanilox_i_tancar
  End If
  If Not vestemfentfingerprint Then
    If Not existeix(nomordinadorcomexi) Then
        MsgBox "No es pot accedir a l'ordinador de Comexi." + Chr(10) + "No es guardarà el fitxer de temperatures.", vbCritical + vbOKOnly, "Atenció"
        'mantenimentbobina.passaravis 0, 0, "Fitxer Temperatures", comanda, "No s'ha pogut accedir a l'ordinador de comexi i no s'ha generat el fitxer de temperatures. TORNEU A PROVAR-HO MES TARD."
        enviaremailgeneric "calidad@inplacsa.com", atrim(comanda) + " - Fitxer Temperatures", "Nº màq: " + atrim(nummaq) + Chr(10) + "Comanda: " + atrim(comanda) + Chr(10) + "Observació= " + "No s'ha pogut accedir a l'ordinador de comexi i no s'ha generat el fitxer de temperatures. TORNEU A PROVAR-HO MES TARD." + Chr(13) + Chr(10) + buscardadesbasiquescomanda(comanda)
        'avisfitxerTemperaturesImpresores=calidad@inplacsa.com
        GoTo sensefitxertemperatures
    End If
  If esunreprint Or existeix("c:\ordprog.ini") Then GoTo sensefitxertemperatures
  Me.caption = "Tancant comanda... Temperatures"
  If buscarcomandaacomexi(comanda) = "" Then
   If UCase(InputBox("No localitzo el fitxer de TEMPERATURES de COMEXI" + Chr(10) + "ABANS DE GUARDAR AQUESTA COMANDA HAS DE GUARDAR LA DE COMEXI." + Chr(10) + "Si vols continuar igualment escriu [continuar]?", "ATENCIÓ")) <> "CONTINUAR" Then Exit Sub
   generarfitxernoudetemperatures comanda
'   mantenimentbobina.passaravis 0, 0, "Fitxer Temperatures", comanda, "No s'ha trobat el fitxer i s'ha generat un automàticament. Reviseu que tot sigui correcte."
   enviaremailgeneric "calidad@inplacsa.com", atrim(comanda) + " - Fitxer Temperatures", "Nº màq: " + atrim(nummaq) + Chr(10) + "Comanda: " + atrim(comanda) + Chr(10) + "Observació= " + "No s'ha trobat el fitxer i s'ha generat un automàticament. Reviseu que tot sigui correcte." + Chr(13) + Chr(10) + buscardadesbasiquescomanda(comanda)
    Else:
      Me.caption = "Tancant comanda... Guardant temperatures"
      resposta = guardar_fitxer_temperatures(cadbl(comanda))
      If resposta <> "" Then
         MsgBox resposta, vbCritical, "Error de temperatures impresora Comexi"
         'mantenimentbobina.passaravis 0, 0, "Fitxer Temperatures", comanda, atrim(resposta)
         enviaremailgeneric "calidad@inplacsa.com", atrim(comanda) + " - Fitxer Temperatures", "Nº màq: " + atrim(nummaq) + Chr(10) + "Comanda: " + atrim(comanda) + Chr(10) + "Observació= " + atrim(resposta) + Chr(13) + Chr(10) + buscardadesbasiquescomanda(comanda)
         generarfitxernoudetemperatures comanda
      End If
  End If
 End If
 Me.caption = "Tancant comanda... Passar Temperatures a directori final"
 passartemperaturestemporalsadirectorifinal
sensefitxertemperatures:
Me.caption = "Tancant comanda... comprovar consums tintes"
 If comprovar_consums_tintes(cadbl(comanda)) Then GoTo fi
 client.ToolTipText = client.caption
 Me.caption = "Tancant comanda... comprovar si falten camps"
 If Not vestemfentfingerprint Then
   If comprovarsifaltencamps Then ratoli "normal": Exit Sub
   verificacio_netejaidespeje
 End If
 If lacomandatereprint(cadbl(comanda)) Then
     If estemfentreprint Then
        comandaacavada.Value = 1
         Else: comandaacavada.Value = 0
     End If
    Else: comandaacavada.Value = 1
 End If
 Me.caption = "Tancant comanda... posso la data de fi"
 If Not impresores.Recordset.EOF Then impresores.Recordset.MoveLast
 If Not impresores.Recordset.EOF Then
    If Not IsDate(impresores.Recordset!datafi) And Not IsDate(impresores.Recordset!horafi) Then
        
        'impresores.Recordset.Edit
        'impresores.Recordset!datafi = Date
        'impresores.Recordset!horafi = Time
        'impresores.Recordset.Update
        'impresores.Refresh
        reixa.Columns("datafi") = Date
        reixa.Columns("horafi") = Time
        hores = DateDiff("n", CVDate(atrim(reixa.Columns("datainici")) + " " + atrim(reixa.Columns("horainici"))), CVDate(atrim(reixa.Columns("datafi")) + " " + atrim(reixa.Columns("horafi"))))
        hores = Redondejar(hores / 60, 2)
        reixa.Columns("totalhores") = hores
        reixa.CurrentCellModified = True
        reixa.EditActive = False
        wait 1
    End If
End If
Me.caption = "Tancant comanda... Calculant metres dolents"
calcular_metresdolents
enviaremailsimesde500metresdolents
Me.caption = "Tancant comanda... Calculant totals"
calcular_totals
If cadbl(kiloshora) > 300 Then
  If MsgBox("Els metres/min surten a mes de 300, es correcte aquest valor o hi ha algun error en el temps de funcionament?", vbCritical + vbYesNo, "Molts mtrs/min") = vbNo Then
     Exit Sub
  End If
End If

Me.caption = "Tancant comanda... mirar bobines no acabades"
mirar_bobinesdentrada_noacavades
If Not vestemfentfingerprint Then
    'comprovar_quenoespassindemetres   ' ho he mogut mes avall per poder calcular primer els metres direrencials
    If metresfetsinferiorsacomanda(cadbl(comanda)) Then Command10_Click: Exit Sub
    If nohihadadesqualitatentrades(cadbl(comanda)) Then Command26_Click: Exit Sub
End If
passar_comanda_a_acavada
comprovar_bobinesnoacabadesAvisarpernoemportarseles cadbl(comanda)
If Not vestemfentfingerprint Then imprimir_full_arrancar_rentar cadbl(comanda)
borrar_ubicacio_deltreball
r = "nopregunta"
des_reservar comanda
guarda_totals
Command4_Click
r = ""
ratoli "espera"

'Command8_Click
Me.caption = "Tancant comanda... Imprimint PackingList..."
guardar_totals_packinglist cadbl(comanda), vdiferenciareal_teoric, vdiferenciaFulla_real
If stockopacking = "E" Then imprimir_packinglist cadbl(comanda)
If Not vestemfentfingerprint Then comprovar_quenoespassindemetres vdiferenciareal_teoric
Me.caption = "Tancant comanda... Imprimint fulla..."
imprimir_fulla
' Aixo de guardar els anilox ho vam fer automatic al afegir l'entrada d'aniloxos timeline
' ja no tenia sentit controlar-ho. abans ho feia desde muntadora l'encarregat de impresores
 'ara ja no es guarden els canvis que es facin, nomes queda reflexat a la comanda
   ' si hi ha lagun canvi ja es fa desde tintes o l'encarregat
  'guardar_aniloxositintesutilitzadescomadefinitives
guardar_estadistica_aniloxos cadbl(comanda), vtotalmetres
guardar_metres_rasquetes vtotalmetres
borrar_bobines_aimpresoresquejanohison
If command15.tag = "Error" Then Me.caption = "Error Impresió...": Exit Sub
If vdiferenciaFulla_real >= 1000 Or vdiferenciaFulla_real <= -1000 Then
   demanar_explicacions_massametres
End If
wait (3)
ratoli "normal"
'r = cadbl(InputBox("Entra la nova comanda", "Fi de comanda"))
carregar_llista_ordreimpressio r, True
If cadbl(r) = cadbl(comanda.text) Or cadbl(r) = 0 Then Exit Sub
If cadbl(r) > 0 Then comanda.text = atrim(cadbl(r))
ratoli "espera"
Command4_Click
ratoli "normal"
If impresores.Recordset.EOF And impresores.Recordset.BOF Then Exit Sub
If cadbl(comanda.text) = 0 Then Exit Sub
trentats = InputBox("Quants tinters has rentat?", "Nova Comanda", atrim(trentats))
pclixers = InputBox("Quants portaclixers?", "Nova Comanda", atrim(pclixers))
canvienfilada = InputBox("Has fet canvi d'enfilada?   S o N ", "Nova Comanda", "N")
If Mid(canvienfilada, 1, 1) = "N" Then
   canvienfilada = "No"
    Else: canvienfilada = "Si"
End If
wait (2)
guarda_totals
fi:
 ratoli "normal"
 Me.caption = "Baixes Comandes (Impressores)"
End Sub
Sub enviaremailsimesde500metresdolents()
   'en principi era mes de 500 pero despres mes de 0 metres
     If cadbl(form1.mtrsdolents) > 0 Then
         mantenimentbobina.passaravis 0, 0, "IMPRESORES: Metres dolents llençats a la comanda " + atrim(comanda), atrim(comanda), "A la comanda " + atrim(comanda) + " s´han possat " + atrim(cadbl(form1.mtrsdolents)) + " metres com a dolents."
'          enviaremailgeneric "compres@inplacsa.com", , "A la comanda " + atrim(comanda) + " s´han possat " + atrim(cadbl(Form1.mtrsdolents)) + " metres com a dolents."
     End If
End Sub
Sub borrar_bobines_aimpresoresquejanohison()
  Dim rst As Recordset
  Dim vpalet As Double
  Dim vbob As Double
  Dim vbobina As String
  Dim vmtrs As Double
  Dim rstb As Recordset
  Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
  ' trec totes les bobines que ja no cal que estiguin dins la taula
  Set rst = dbtmpb.OpenRecordset("Select * from impresores_bobinesamaquina where data>#" + Format(DateAdd("m", -3, Now), "mm/dd/yy") + "# order by data desc")
  While Not rst.EOF
    vbobina = rst!numbobina
    vpalet = cadbl(Mid(" " + vbobina, 1, InStr(1, vbobina + "  ", "/")))
    vbob = cadbl(Mid(vbobina, InStr(1, vbobina + "  ", "/") + 1))
    Set rstb = dbtmpb.OpenRecordset("select sit from bobines where idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbob))
    If Not rstb.EOF Then
       vmtrs = bobinesdentrada.calcular_mtrsdispreals(vpalet, vbob)
       If vmtrs <= 500 Then
            rst.Delete
           Else: If rstb.sit <> "IMP" Then rst.Delete
       End If
    End If
    rst.MoveNext
  Wend
  Set rst = Nothing
  Set rstb = Nothing
  


End Sub

Sub demanar_explicacions_massametres()
   Dim vmsg As String
   While atrim(vmsg) = ""
      vmsg = InputBox("Hi ha hagut un desquadrament amb els metres Reals i els de la Fulla." + vbNewLine + "ESCRIU UNA EXPLICACIÓ DEL PERQUÈ PER TAL DE PODER-HO SOLUCIONAR CORRECTAMENT.", "DESQUADRAMENT")
   Wend
    enviaremailgeneric "missatgesgenericsimpresores", "Desquadrament de metres Reals i els de la Fulla. Comanda: " + form1.comanda + "  " + nommaq + " - " + nomoperari, "Missatge de l'operari: " + vbNewLine + treure_apostruf(vmsg)
End Sub
Sub guardar_metres_rasquetes(vmetres As Double)
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim i As Byte
   Set rst = dbtmpb.OpenRecordset("select * from aniloxtimeline where comanda=" + form1.comanda)
   If rst.EOF Then GoTo fi
   If rst!rasquetesactualitzades Then Exit Sub
   For i = 1 To 8
     If rst.Fields("color" + atrim(i)) <> "" Then
         Set rst2 = dbtmpb.OpenRecordset("select * from impresores_rasquetes where nummaquina=" + atrim(nummaq) + " and numrasqueta=" + atrim(i))
         If rst2.EOF Then
                rst2.AddNew
                rst2!nummaquina = nummaq
                rst2!numrasqueta = i
              Else: rst2.Edit
         End If
         rst2!metres = cadbl(rst2!metres) + cadbl(vmetres)
         rst2.Update
     End If
   Next i
   rst.Edit
   rst!rasquetesactualitzades = True
   rst.Update
fi:
   Set rst = Nothing
End Sub
Sub guardar_aniloxositintesutilitzadescomadefinitives()
   Dim i As Integer
   Unload formaniloxos
   Load formaniloxos
   formaniloxos.tag = Me.comanda
   formaniloxos.boto_nou(0).tag = Command26.tag
   formaniloxos.fbotonsok.tag = "activats"
   formaniloxos.Show
   formaniloxos.carregartintes cadbl(formaniloxos.tag)
   
   'formaniloxos.visible = False
   wait 2
   formaniloxos.guardar_totselscanvisdanilox
   
   Unload formaniloxos
End Sub
Sub guardar_estadistica_aniloxos(vcomanda As Double, vmetres As Double)
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("SELECT aniloxtimeline.* From aniloxtimeline Where (((aniloxtimeline.comanda) = " + atrim(vcomanda) + ")) ORDER BY aniloxtimeline.data DESC;")
   If Not rst.EOF Then
      rst.Edit
      rst!totalmetres1 = 0: rst!totalmetres2 = 0: rst!totalmetres3 = 0: rst!totalmetres4 = 0: rst!totalmetres5 = 0: rst!totalmetres6 = 0: rst!totalmetres7 = 0: rst!totalmetres8 = 0
      If atrim(rst!color1) <> "" Then rst!totalmetres1 = vmetres
      If atrim(rst!color2) <> "" Then rst!totalmetres2 = vmetres
      If atrim(rst!color3) <> "" Then rst!totalmetres3 = vmetres
      If atrim(rst!color4) <> "" Then rst!totalmetres4 = vmetres
      If atrim(rst!color5) <> "" Then rst!totalmetres5 = vmetres
      If atrim(rst!color6) <> "" Then rst!totalmetres6 = vmetres
      If atrim(rst!color7) <> "" Then rst!totalmetres7 = vmetres
      If atrim(rst!color8) <> "" Then rst!totalmetres8 = vmetres
      rst.Update
   End If
   Set rst = Nothing
   'enviar ordre al servidor de calcular els aniloxos
   escriure_ini "General", "calcularestadisticaaniloxos", "Si", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini"
End Sub
Function metresimpresosdelacomanda(vnumc As Double) As Double
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("select tmetresdolents,tprova,tmetres from impressorestot where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
       metresimpresosdelacomanda = cadbl(rst!tmetresdolents) + cadbl(rst!tprova) + cadbl(rst!tmetres)
   End If
   Set rt = Nothing
End Function
Sub imprimir_full_arrancar_rentar(vnumc As Double)
  Dim rst As Recordset
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vnumtreball As String
  Dim vidtinters As String
  Dim vmetresimpresos As Double
  
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  Set rst = dbtmp.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vnumc))
  If rst.EOF Then Exit Sub
  vmetresimpresos = metresimpresosdelacomanda(vnumc)
  vnumtreball = atrim(cadbl(rst!numtreball))
  vsql = "SELECT IIf([tinterlinkambid_treball]>0,[tinterlinkambid_treball],[id_tinter]) AS indinters From Tintes  WHERE Tintes.id_treball=" + vnumtreball + " and ordremodificacio=" + atrim(rst!numordremodificacio) + ";"
  Set rst = dbclixes.OpenRecordset(vsql)
  While Not rst.EOF
     vidtinters = vidtinters + IIf(vidtinters = "", "", " or ") + " id_tinter=" + atrim(rst!indinters)
     rst.MoveNext
  Wend
  vsql = "SELECT distinct Tintes.id_treball From Tintes WHERE " + vidtinters
  Set rst = dbclixes.OpenRecordset(vsql)
  If rst.EOF Then Exit Sub
  
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "etiqueta_arrencarirentar.rpt", 1) '"etiqueta_llaunes.rpt"

  
  'oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "tintes.mdb"
 ' oreport.RecordSelectionFormula = "{Llaunes.numllauna}='" + UCase(atrim(numllauna)) + "'"
  'oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
 ' oreport.Sections("D").ReportObjects.Item("serie2").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
  oreport.DiscardSavedData
  'oreport.Sections("D").ReportObjects.Item("recuperador").Suppress = True
  While Not rst.EOF
    vnumtreball = atrim(cadbl(rst!id_treball))
    oreport.FormulaFields.GetItemByName("numtreball").text = "'" + vnumtreball + "'"
    oreport.FormulaFields.GetItemByName("metresimpresos").text = "'" + atrim(vmetresimpresos) + "'"
    If estemfentreprint Then
       oreport.FormulaFields.GetItemByName("etvernisreprint").text = "'REPRINT_Vernis'"
        Else: oreport.FormulaFields.GetItemByName("etvernisreprint").text = "''"
    End If
    oreport.DisplayProgressDialog = False
    oreport.PrintOut False, 1
    rst.MoveNext
  Wend
  
  Set rst = Nothing
End Sub
Sub verificacio_netejaidespeje()
  'Dim v As String
  'Dim vcont As Byte
  'vcont = 9
'//tret per ordre de lencarregat i en miralles (Possar nomes OK)
  'While UCase(v) <> "NETEJA" And vcont > 0
  '  v = InputBox("Verificació de Neteja i despeje de línia." + Chr(10) + "Escriu [neteja] per acceptar", "Neteja i despeje (" + atrim(vcont) + ")")
  '  vcont = vcont - 1
  'Wend
  MsgBox "Verificació de Neteja i despeje de línia." + vbNewLine + "Verificació d'elements físics.", vbExclamation + vbOKOnly, "Neteja, despeje i elements físics."
End Sub
Function nohihadadesqualitatentrades(numc As Double) As Boolean
   Dim rst As Recordset
   nohihadadesqualitatentrades = True
   Set rst = dbtmpb.OpenRecordset("select * from  impressorestot where comanda=" + atrim(cadbl(numc)))
   If Not rst.EOF Then
      If cadbl(rst!qualitatimpresio) > 0 Then nohihadadesqualitatentrades = False
   End If
   Set rst = Nothing
End Function
Sub borrar_ubicacio_deltreball()
  Dim rstc As Recordset
  Dim dbclixes As Database
  Set rstc = dbtmp.OpenRecordset("select numtreball from comandes where comanda=" + atrim(cadbl(comanda)))
  If rstc.EOF Then Exit Sub
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  dbclixes.Execute "update  clixes set ubicacio='' where arxiu<>'' and id_treball=" + atrim(cadbl(rstc!numtreball))
  Set dbclixes = Nothing
  Set rstc = Nothing
End Sub
Function lacomandatereprint(vnumc As Double) As Boolean
   Dim rstc As Recordset
   Dim rstt As Recordset
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
   Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
   If rstc.EOF Then GoTo fi
   Set rstt = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordre=" + atrim(cadbl(rstc!numordremodificacio)))
   If rstt.EOF Then GoTo fi
   lacomandatereprint = IIf(Not IsNull(rstt!reimpres), rstt!reimpres, False)
   vhihadigimarc = IIf(rstt!digimarc = "Si", True, False)
fi:
   Set rstc = Nothing
   Set rstt = Nothing

End Function
Sub passar_comanda_a_acavada()
Dim estat As String
Dim ruta As String
Dim vresp As String
Dim vdata As String
Dim vidtreball As Double
Dim vordre As Double
'si hi ha alguna bobina passo l'estat de la comanda a la proxima seccio
  impresores.Recordset.MoveLast
  'posso la data als totals de seccio
  If IsDate(impresores.Recordset!datafi) Then
   vdata = atrim(impresores.Recordset!datafi) + " " + atrim(IIf(Format(impresores.Recordset!horafi, "hhnn") = "0000", Time, Format(impresores.Recordset!horafi, "hh:nn")))
     Else: vdata = Format(Now, "dd/mm/yy") + " " + Format(Now, "hh:nn")
  End If
  dbtmpb.Execute "update impressorestot set dataimpressio=#" + Format(vdata, "mm/dd/yy hh:nn:ss") + "# where comanda=" + atrim(cadbl(comanda))
   'SI ES REPRINT PREGUNTO PER AVANÇAR LA SECCIO O NO
   If lacomandatereprint(cadbl(comanda)) Then
      vresp = UCase(InputBox("Aquesta comanda porta REPRINT, ja has acabat el reprint?" + Chr(10) + "[SI] si has fet la segona passada [NO] si estàs fent l'impresió.", "Reprint"))
      If vresp = "SI" Then
            If llistallaunesreprint.ListCount = 0 Then
               MsgBox "No hi ha cap numero de llauna entrat en el reprint, comprova que sigui correcte", vbCritical, "Atenció"
               GoTo noavançarseccio
            End If
          Else: GoTo noavançarseccio
      End If
   End If
   'passo l'estat de comanda a la proxima
   Set rsttmp = dbtmp.OpenRecordset("select producte,proximaseccio,numtreball,numordremodificacio from comandes where comanda=" + atrim(comanda))
   If Not rsttmp.EOF Then
     vidtreball = rsttmp!numtreball
     vordre = rsttmp!numordremodificacio
     estat = atrim(rsttmp!proximaseccio)
     If estat = "" Or estat = "E" Then estat = "I"
   End If
   Set rsttmp = dbtmp.OpenRecordset("select ruta from productes where codi='" + rsttmp!producte + "'")
   ruta = rsttmp!ruta
   If estat = "I" Then
     seccio = Mid(ruta, 3, 1)
     If seccio = "" Then seccio = "V"
     dbtmp.Execute "update comandes set proximaseccio='" + seccio + "' where comanda=" + atrim(comanda)
     dbtmp.Execute "update comandes set seccioactual='I' where comanda=" + atrim(comanda)
     passaraRcomandespendentsambelmateixtreballiversio cadbl(comanda), cadbl(vidtreball), cadbl(vordre)
   End If
   Set rsttmp = Nothing
noavançarseccio:
End Sub
Sub passaraRcomandespendentsambelmateixtreballiversio(vnumc As Double, vidtreball As Double, vordre As Double)
   dbtmp.Execute "update comandes set impressio='R' where comanda<>" + atrim(vnumc) + " and (numtreball=" + atrim(vidtreball) + " and numordremodificacio=" + atrim(vordre) + ") and (proximaseccio='E' or proximaseccio='I')"
End Sub
Sub guardar_totals_packinglist(numc As Double, Optional vdiferenciareal_teoric As Double, Optional vdiferenciaFulla_real As Double)
   Dim rstt As Recordset
   Dim rstc As Recordset
   Dim rsto As Recordset
   Dim rsti As Recordset
   Dim rste As Recordset
   obrestocks
   Set rsto = dbstocks.OpenRecordset("select * from opcionsdajust where comanda=" + atrim(numc))
   Set rstt = dbstocks.OpenRecordset("select * from totals_full_packinglist where comanda=" + atrim(numc))
   Set rstc = dbtmp.OpenRecordset("select cantitatex from comandes where comanda=" + atrim(numc))
   If rstc.EOF Then Exit Sub
   If rstt.EOF Then
      rstt.AddNew
      rstt!comanda = numc
        Else: rstt.Edit
   End If
   rstt!mtrs_comanda = cadbl(rstc!cantitatex)
   rstt!t_mtrs_assignat = calcularmetrespackinglist(cadbl(comanda), "historic_packinglist", 0)
   rstt!r_mtrs_assignat = calcularmetrespackinglist(cadbl(comanda), "parcials", 1)
   Set rste = dbtmp.OpenRecordset("SELECT comandes_extres.assignarstock as estoc, comandes_extres.mtrsassignatsestock as metresestoc frOM comandes_extres WHERE comanda=" + atrim(cadbl(comanda)) + ";")
   If Not rste.EOF Then
       If rste!estoc Then
          rstt!t_mtrs_assignat = cadbl(rste!metresestoc)
          If rstt!t_mtrs_assignat = 0 Then rstt!t_mtrs_assignat = rstt!mtrs_comanda
          rstt!esestoc = True
            Else: rstt!esestoc = False
       End If
   End If
   
   'possar ajustteoric
   If Not rsto.EOF Then
    If rsto!sistemadajust = 1 Then rstt!t_ajust_llençar = rsto!mtrsajust
    If rsto!sistemadajust = 2 Then rstt!t_ajust_estoc = rsto!mtrsajust
    If rsto!sistemadajust = 3 Then
       rstt!t_ajust_paletbob = rsto!mtrsajust
       rstt!t_mtrs_assignat = rstt!t_mtrs_assignat - rsto!mtrsajust
       rstt!t_ajust_numpaletbob = atrim(rsto!paletajust) + "/" + atrim(rsto!bobinaajust)
    End If
   End If
   
   'possar ajustreal
   If Not rsto.EOF Then
    rstt!r_ajust_llençar = metresrealsllençar(numc)
    If rsto!sistemadajust = 2 Then rstt!r_ajust_estoc = calcularmetrespackinglist(cadbl(comanda), "parcials", 2)
    If rsto!sistemadajust <> 2 Then
       rstt!r_ajust_paletbob = calcularmetrespackinglist(cadbl(comanda), "parcials", 2)
    End If
   End If
   'metres dolents
   Set rsti = dbtmpb.OpenRecordset("select tmetresdolents from impressorestot where comanda=" + atrim(numc))
   If Not rsti.EOF Then rstt!mtrs_dolents = rsti!tmetresdolents
   'metres impresos
   Set rsti = dbtmpb.OpenRecordset("SELECT impressores.comanda, Sum(bobinesimp.metres) AS SumaDemetres FROM bobinesimp INNER JOIN impressores ON bobinesimp.controlid = impressores.Id GROUP BY impressores.comanda HAVING (((impressores.comanda)=" + atrim(numc) + "));")
   If Not rsti.EOF Then rstt!mtrs_impresos = rsti!sumademetres
   
   'actualitzo els valors de packinglist real
   If cadbl(rstt!mtrs_dolents) > 0 Then
   
      'HE SUBSTITUIT LA LINIA DE SOTA PER POSSAR ELS METRES DOLENTS PERÒ ABANS ESTAVA AIXÍ
      'L'ALICIA HO HA VOLGUT AIXÍ A DIA 26/04/2022
      'rstt!r_mtrs_dolents =rstt!r_mtrs_assignat - cadbl(rstt!r_ajust_estoc) - cadbl(rstt!r_ajust_paletbob) - cadbl(rstt!mtrs_impresos)
      rstt!r_mtrs_dolents = rstt!mtrs_dolents
      
      rstt!r_mtrs_assignat = cadbl(rstt!mtrs_impresos)
    Else:
       rstt!r_mtrs_dolents = 0
'27/01/22 he tret perquè els totals de real diuen que no es correcte no se si es així UNA PROVA
       ' rstt!r_mtrs_assignat = rstt!r_mtrs_assignat - cadbl(rstt!r_ajust_estoc) - cadbl(rstt!r_ajust_paletbob)
   End If
   
   rstt.Update
   If Not rstt.EOF Then
      vdiferenciareal_teoric = cadbl(rstt!r_mtrs_assignat) + cadbl(rstt!r_mtrs_dolents) + cadbl(rstt!r_ajust_estoc) + cadbl(rstt!r_ajust_paletbob) + cadbl(rstt!r_ajust_llençar)
      vdiferenciaFulla_real = vdiferenciareal_teoric
      vdiferenciareal_teoric = vdiferenciareal_teoric - (cadbl(rstt!t_mtrs_assignat) + cadbl(rstt!t_ajust_estoc) + cadbl(rstt!t_ajust_paletbob) + cadbl(rstt!t_ajust_llençar))
      vdiferenciaFulla_real = vdiferenciaFulla_real - (cadbl(rstt!mtrs_impresos) + cadbl(rstt!mtrs_dolents) + cadbl(rstt!r_ajust_paletbob) + cadbl(rstt!r_ajust_estoc) + cadbl(rstt!r_ajust_llençar))
   End If
   Set rsti = Nothing
   Set rstt = Nothing
   Set rstc = Nothing
   Set rsto = Nothing
   Set rste = Nothing
End Sub
Function metresrealsllençar(numc As Double)
  Dim rstc As Recordset
  Set rstc = dbtmpb.OpenRecordset("select sum(metresprova) as total from impressores where paletprova=11111 and comanda=" + atrim(numc))
  If Not rstc.EOF Then metresrealsllençar = cadbl(rstc!total)
  Set rstc = dbtmpb.OpenRecordset("select sum(metresprova2) as total from impressores where paletprova2=11111 and comanda=" + atrim(numc))
  If Not rstc.EOF Then metresrealsllençar = metresrealsllençar + cadbl(rstc!total)
  Set rstc = dbstocks.OpenRecordset("select sum(metres) as total from parcials where comanda='" + atrim(numc) + "' and orcomassignacio='500' and instr(1,[observacions],'#llençar')>0")
  If Not rstc.EOF Then metresrealsllençar = metresrealsllençar + cadbl(rstc!total)
  Set rstc = Nothing
End Function
Function calcularmetrespackinglist(numc As Double, nomtaula As String, Optional senseajust As Byte) As Double
  Dim rststock As Recordset
  calcularmetrespackinglist = 0
  'Clipboard.SetText "select sum(metres) as total from " + nomtaula + " where " + IIf(senseajust = 1, " orcomassignacio<>'500' and ", IIf(senseajust = 2, "(orcomassignacio='500' and instr(1,[observacions],'#llençar')=0) and ", "")) + " comanda='" + atrim(numc) + "'"
  Set rststock = dbstocks.OpenRecordset("select sum(metres) as total from " + nomtaula + " where " + IIf(senseajust = 1, " orcomassignacio<>'500' and ", IIf(senseajust = 2, "(orcomassignacio='500' and InStr(1,IIf([observacions] Is Null,'',[observacions]),'#llençar')=0) and ", "")) + " comanda='" + atrim(numc) + "'")
 ' Clipboard.Clear
 ' Clipboard.SetText "select sum(metres) as total from " + nomtaula + " where " + IIf(senseajust = 1, " orcomassignacio<>'500' and ", IIf(senseajust = 2, "(orcomassignacio='500' and InStr(1,IIf([observacions] Is Null,'',[observacions]),'#llençar')=0) and ", "")) + " comanda='" + atrim(numc) + "'"
  If Not rststock.EOF Then calcularmetrespackinglist = cadbl(rststock!total)
  Set rststock = Nothing
End Function

Function comprovarsifaltencamps() As Boolean
  Dim faltenpatones As Boolean
  Dim faltenmtrs As Boolean
  If cadbl(tmetres) < (metrescomanda - (metrescomanda / 10)) And metrescomanda <> 999999 Then
      If UCase(InputBox("No arrives al -10% dels " + atrim(metrescomanda) + Chr(10) + Chr(13) + "PER CONTINUAR ESCRIU [SI]", "-10% METRES COMANDA")) <> "SI" Then
        comprovarsifaltencamps = True
        Exit Function
      End If
  End If
  For i = 0 To 7
    If atrim(pantone(i)) <> "" And cadbl(kbpantone(i)) = 0 Then
       faltenpantones = True
    End If
  Next i
  impresores.Recordset.FindFirst "tipus='F'"
  While Not impresores.Recordset.NoMatch
    If cadbl(impresores.Recordset!mtrsminut) = 0 Then
      impresores.Recordset.Edit
      impresores.Recordset!mtrsminut = cadbl(InputBox("Falten els Mtrs/Min.", "Atenció"))
      impresores.Recordset.Update
      If cadbl(impresores.Recordset!mtrsminut) <= 0 Then MsgBox "Els Mtrs/Min han de ser >0": comprovarsifaltencamps = True
    End If
    impresores.Recordset.FindNext "tipus='F'"
  Wend
  If faltenpantones Then MsgBox "Falta entrar els Kg de tinta.": comprovarsifaltencamps = True
  
End Function

Private Sub Command16_Click()
  Dim r As Variant
  'llistatpantones.Show 1
  'Set rstpantones = Nothing
 'Set dbpantones = Nothing
 'Unload llistatpantones
   r = Shell(llegir_ini("General", "rutallistats", "comandes.ini") + "llistatpantones.exe", vbNormalFocus)
   
   
End Sub




Private Sub Command18_Click()
  manteniment.Show 1, form1
End Sub

Private Sub Command19_Click()
   panellavis.visible = False
End Sub

Sub demanar_metres_dajust()
  While cadbl(impresores.Recordset!mtrsprova) = 0
        bobsajust_Click
        If cadbl(impresores.Recordset!mtrsprova) Then MsgBox "Has d'entrar metres d'ajust per poder continuar.", vbCritical, "Atenció"
   Wend
End Sub

Private Sub Command2_Click()
  Dim secciocreada As Boolean
  If Not comprovamaq Then Exit Sub
  comprovarsidescansorelleu
  If Not impresores.Recordset.EOF Then
   impresores.Recordset.MoveLast
   If impresores.Recordset!tipus = "A" Then
      'mtrsprova = InputBox("Entra els Metres de prova.", "Atenció")
      
      'impresores.Recordset.Edit
      'impresores.Recordset!mtrsprova = cadbl(mtrsprova)
      'impresores.Recordset!paletbobprova = demanar_bobinaprova(impresores.Recordset!paletbobprova, True)
      'If cadbl(impresores.Recordset!paletprova2) > 0 Then impresores.Recordset!paletbobprova = "*" + atrim(impresores.Recordset!paletprova) + "-" + atrim(impresores.Recordset!bobinaprova)
      
     ' impresores.Recordset.Update
      demanar_metres_dajust
      numop = escullir_operari
      nomoperari = UCase(r)
      numop2 = escullir_operari("Escullir AJUDANT d'OPERARI")
      nomoperari2 = "Ajudant: " + UCase(r)
        Else
          If impresores.Recordset!tipus = "M" Then
            crearseccio "A"
            secciocreada = True
            
            'impresores.Recordset.Edit
            'impresores.Recordset!paletbobprova = demanar_bobinaprova(impresores.Recordset!paletbobprova, True)
            'If cadbl(impresores.Recordset!paletprova2) > 0 Then impresores.Recordset!paletbobprova = "*" + atrim(impresores.Recordset!paletprova) + "-" + atrim(impresores.Recordset!bobinaprova)
            'impresores.Recordset.Update
          End If
   End If
  End If
  'impresores.Recordset.FindLast "tipus='A'"
  'If impresores.Recordset.NoMatch Then
  '  numop = escullir_operari
  '  nomoperri = UCase(r)
  'End If
  verificar_tubbase
  imprimir_controlqualitatVQ cadbl(comanda)
  If Not secciocreada Then
     crearseccio "A": bobsajust_Click
      Else:
        If Not estemfentreprint Then
           obrir_llegirllaunes
             Else
                'ensenyo panel de llaunes reprint
                Command26_Click
                'apreto el botó dafegir llaunes vernis
                Command31_Click
        End If
  End If
  verificarsihihaentratelcanvidaniloxos
End Sub
Sub verificarsihihaentratelcanvidaniloxos()
  Dim rst As Recordset
  Dim vcont As Byte
  vcont = 0
  Do
    Set rst = dbtmpb.OpenRecordset("select * from aniloxtimeline where nummaquina=" + atrim(nummaq) + " and comanda=" + comanda.text + " order by data desc", , ReadOnly)
    If rst.EOF Then
       MsgBox "No hi ha entrat el canvi d'anilox per aquesta comanda, no pots continuar sense entrar-los", vbCritical, "Atenció"
       formcanvisanilox.Show 1: Unload formcanvisanilox
    End If
    vcont = vcont + 1
  Loop While rst.EOF And vcont < 3
End Sub
Function escullir_avaria() As String
  Dim opvell As Byte
  Dim r As String
   Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select  tipificacioavaria as [Tipus avaria] from impresores_tipificacionsavaria order by 1"
   formseleccio.caption = "Selecció Avaria"
   formseleccio.refrescar
   formseleccio.Show 1
   If seleccioret = 1 Then
    escullir_avaria = formseleccio.Data1.Recordset![Tipus avaria]
   End If
   Unload formseleccio
End Function

Sub crearseccio(tipus As String, Optional vobservacio As String)
 Dim com As Double
 Dim rsttmpcs As Recordset
 Dim vtipusavaria As String
 Dim vobservaciodelaavaria As String
 
  r = ""
  Set rsttmpcs = dbtmp.OpenRecordset("select comanda,texteimpressio from comandes where comanda=" + Trim(comanda))
  
  ' MsgBox "Baixa nova es començarà amb edició de Clixes."
   com = cadbl(comanda)
   If rsttmpcs.EOF Then MsgBox "No hi ha numero de comanda vàlida": com = 0
  If Not impresores.Recordset.EOF Then
      finalitza_seccio
      'If r <> "no" Then Exit Sub
      com = cadbl(impresores.Recordset!comanda)
  End If
  r = ""
  If com = 0 Then Exit Sub
  If tipus = "V" Then
    vtipusavaria = escullir_avaria
    If vtipusavaria = "" Then vtipusavaria = InputBox("Escriu la descripció de l'avaria", "Error")
    If vtipusavaria = "" Then MsgBox "S'ha d'escullir una avaria per continuar.", vbCritical, "Error": Exit Sub
    vobservaciodelaavaria = InputBox("Escriu una descripció de la avaria si correspon.", "Error")
    enviaremailsical_avaria vtipusavaria, vobservaciodelaavaria
  End If
  impresores.Recordset.AddNew
  impresores.Recordset!comanda = com
  impresores.Recordset!numeromaquina = nummaq
  impresores.Recordset!operari = numop
  impresores.Recordset!operari2 = numop2
  impresores.Recordset!tipus = tipus
  impresores.Recordset!datainici = Date
  impresores.Recordset!horainici = Time
  impresores.Recordset!tipificacioavaria = Mid(vtipusavaria, 1, 50)
  impresores.Recordset!texteimpresio = rsttmpcs!texteimpressio
  impresores.Recordset!paletbobprova = Mid(vobservacio, 1, 20)
  If tipus = "V" Then impresores.Recordset!observacio = Mid(atrim(vobservaciodelaavaria), 1, 100)
  r = impresores.Recordset!id
  impresores.Recordset.Update
  impresores.Recordset.MoveLast
  Set rsttmpcs = Nothing
  possar_peu_imprenta cadbl(direnvio), tipus
End Sub
Sub possar_peu_imprenta(denvio As Long, tipus As String)
  Dim rstd As Recordset
  vavispeu = ""
  If denvio > 0 Then
      Set rstd = dbtmp.OpenRecordset("SELECT Clients_envios.codi, peuimprenta.descripcio FROM Clients_envios LEFT JOIN peuimprenta ON Clients_envios.peuimprenta = peuimprenta.codi where clients_envios.id=" + atrim(denvio))
      If Not rstd.EOF Then vavispeu = atrim(rstd!descripcio)
      If vavispeu <> "" And tipus = "A" Then MsgBox """Oju"" amb el peu d'imprenta. " + Chr(10) + vavispeu, vbExclamation + vbOKOnly, "Atenció"
  End If
  Set rstd = Nothing
End Sub

Function comprovamaq() As Boolean
   comprovamaq = True
   If nummaq = 0 Then
       MsgBox "Escull primer un numero de màquina"
      comprovamaq = False
   End If
End Function
Sub comprovar_reducciocilindre()
  If InStr(1, avisapantalla, "Cilindre:") > 0 Then
     panellavis.visible = True
     panellavis.Top = 800
     panellavis.Left = 1400
     missatgeavis = avisapantalla
  End If
End Sub


Private Sub Command20_Click()
  Dim palet As Double
  Dim bobina As Double
  Dim rst As Recordset
  Dim inssql As String
  Dim jaexisteix As Boolean
  Dim metresreals As Double
  Dim observacio As String
  Dim vresp As String
  Dim vgrup As Double
  If bobines.Recordset.EOF Then GoTo fi
  demanar_paletibobina palet, bobina
  
  If palet > 0 And bobina > 0 Then
    obrestocks
    If etmaterialexacte <> "" Then
      If Not comprovar_materialexacte(palet, cadbl(form1.etmaterialexacte.tag)) Then
         MsgBox "Aquest material no es exactament el que demana el client." + vbNewLine + "PERO ASSEGURA SI ES CORRECTE ABANS D'UTILITZAR-LA", vbCritical, "Error"
      End If
    End If
    If stockopacking <> "E" Then
       vresp = comprovarsieselmateixmaterial(palet, bobina, cadbl(comanda), vgrup, vgrupmaterialcompatible)
       If InStr(1, vresp, "#materialerror") > 0 Then GoTo fi
       If InStr(1, vresp, "#noesdelpackinglist") > 0 Then
           vresp = InputBox("Aquesta bobina no es del packinglist, si vols continuar utilitzant-la escriu una explicació perquè l'utilitzes." + vbNewLine + "NO ESCRIGUIS RES SI NO VOLS UTILITZAR-LA O FES CANCELAR", "Bobina no Packinglist")
           If vresp = "" Then GoTo fi
           passaravisdebobinaassignadaaunaaltracomanda cadbl(palet), cadbl(bobina), cadbl(comanda), vresp
       End If
    End If
    inssql = "SELECT CDbl([comanda]) AS Expr1, Parcials.idpalet, Parcials.idbobina,parcials.orcomassignacio  From Parcials WHERE (((CDbl([orcomassignacio])<10000 and cdbl([orcomassignacio])>2000)) and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + ");"
    Set rst = dbstocks.OpenRecordset(inssql)

    If rst.EOF Then
      'he modificat la linia seguent per controlar un error quan el material era exacte, he deixa la coletilla al final
      ' em sembla que així es correcte pero no se si efectarà a quelcom mes
     If stockopacking = "E" Then  ' 29/08/23 aquestes 4 linies afegides per problema amb tipus de material diferent a Estoc
        vresp = comprovarsieselmateixmaterial(palet, bobina, cadbl(comanda), vgrup, vgrupmaterialcompatible)
        If InStr(1, vresp, "#materialerror") > 0 Then GoTo fi
     End If
     inssql = "select * from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and comanda=[orcomassignacio]" ''" + atrim(orcomassignacio) + "'"
     Set rst = dbstocks.OpenRecordset(inssql)
     
       Else:
          If stockopacking = "P" Then
             observacio = "S'ha agafat una bobina d'ESTOC en una comanda de PACKING-LIST."
             While vresp = ""
                  vresp = InputBox("Entra una explicació perquè utilitzes una bobina d'ESTOC en una comanda de PACKING-LIST.", "Comentari")
             Wend
          End If
          If stockopacking = "E" And cadbl(rst!orcomassignacio) <> cadbl(veuregrupsdestoc.tag) Then
              GoTo controlarcanvidematerialestoc
          End If
    End If
    
    If rst.EOF Then
controlarcanvidematerialestoc:
      If UCase(InputBox("El Palet: " + atrim(palet) + "/" + atrim(bobina) + " no està assignat per utilitzar-lo." + Chr(10) + Chr(13) + "Escriu ACCEPTO per agafar-la igualment", "Palet/Bobina equivocat")) = "ACCEPTO" Then
         If stockopacking = "E" Then
            observacio = "S'ha agafat una bobina que no era ESTOC per fer una comanda d'ESTOC."
            While vresp = ""
                vresp = InputBox("Entra una explicació perquè utilitzes una bobina d'ESTOC en una comanda de PACKING-LIST.", "Comentari")
            Wend
         End If
         passaravisdebobinaassignadaaunaaltracomanda cadbl(palet), cadbl(bobina), cadbl(comanda), vresp
         GoTo accepto
          Else: GoTo fi
      End If
      'MsgBox inssql
     Else
accepto:
       metresreals = bobinesdentrada.calcular_mtrsdispreals(palet, bobina)
       If metresreals < 500 Then
         If MsgBox("Aquesta bobina està donada per ACABADA." + Chr(10) + "VOLS AFEGIR-LA IGUALMENT?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then GoTo afegirbob
          Else:
afegirbob:
             afegir_labobinadentrada palet, bobina
             If palet <> 0 Then   'si palet es 0 es perque al afegir_labobinadentrada ha vist que estava repetida
             
               If Not hihaparcialassignat(palet, bobina, cadbl(comanda)) Then dbstocks.Execute "insert into parcials (idpalet,idbobina,metres,comanda,orcomassignacio) values (" + atrim(palet) + "," + atrim(bobina) + ",0,'" + atrim(cadbl(comanda)) + "'," + atrim(cadbl(comanda)) + ")"
               demanar_verificacio_espesoritractat palet, bobina
               imprimir_controlqualitatbobinaentrada palet, bobina, 0
             End If
       End If
    End If
  End If
  If observacio <> "" Or vresp <> "" Then mantenimentbobina.passaravis palet, bobina, observacio, comanda, observacio + IIf(vresp <> "", " Explicació OP: " + vresp, "[L'OPERARI NO HA POSSAT EXPLICACIÓ]")
fi:
  ratoli "normal"
  Set rst = Nothing
End Sub
Function hihaparcialassignat(p As Double, b As Double, vnumc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbstocks.OpenRecordset("select * from parcials where utilitzada=false and cdbl(orcomassignacio)=" + atrim(vnumc) + " and idpalet=" + atrim(p) + " and idbobina=" + atrim(b))
   If Not rst.EOF Then hihaparcialassignat = True
End Function
Sub passaravisdebobinaassignadaaunaaltracomanda(palet As Double, bobina As Double, comanda As Double, Optional vmotiu As String)
   Dim rst As Recordset
   Set rst = dbstocks.OpenRecordset("select * from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and comanda<>'" + atrim(comanda) + "' and utilitzada=false")
   If Not rst.EOF Then mantenimentbobina.passaravis palet, bobina, "!!!!! ATENCIÓ  BOBINA ASSIGNADA UTILITZADA PER UNA ALTRA COMANDA", atrim(comanda), "La bobina " + atrim(palet) + "/" + atrim(bobina) + " estava assignada a la comanda " + atrim(rst!comanda) + " i s'ha utilitzat per la " + atrim(comanda) + vbNewLine + IIf(vmotiu <> "", "Explicació operari: " + vmotiu, "")
   Set rst = Nothing
End Sub
Sub demanar_paletibobina(palet As Double, bobina As Double)
  Unload entradabobina
  Load entradabobina
  entradabobina.etdesb.visible = False
  entradabobina.desb.visible = False
  
  entradabobina.Show 1
  If cadbl(entradabobina.palet) > 0 Then
     palet = cadbl(entradabobina.palet)
     bobina = cadbl(entradabobina.bobina)
  End If
  Unload entradabobina
End Sub

Private Sub Command21_Click()
  Me.PopupMenu mmenucalculadores
  
End Sub

Private Sub Command22_Click()
   metresdolents.Show
   calcular_metresdolents
End Sub
Sub calcular_metresdolents()
   Dim rstdo As Recordset
   Set rstdo = dbtmpb.OpenRecordset("select sum(metres) as mtrs from impresores_mtrsdolents where idcomanda=" + atrim(cadbl(form1.comanda)))
   If Not rstdo.EOF Then
     form1.mtrsdolents = cadbl(rstdo!mtrs)
       Else: form1.mtrsdolents = 0
   End If
   Set rstdo = Nothing
   form1.guarda_totals
End Sub




Sub veureelimp()
  Dim rstc As Recordset
  Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(comanda)))
  obrir_imp_treball cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), cadbl(rstc!client), cadbl(rstc!direnvio)
End Sub
Sub obrir_imp_treball(treball As Double, modificacio As Double, codiclient As Double, direnvio As Double)
   Dim generarfitxer_imp As String
   If modificacio = 0 Then modificacio = 1
   generarfitxer_imp = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\IMP" + Format(treball, "00000") + "-" + Format(modificacio, "000") + "-" + Format(codiclient, "000000") + "_" + atrim(direnvio) + ".doc"
   If Not existeix(generarfitxer_imp) Then generarfitxer_imp = generarfitxer_imp + "x"
   If existeix(generarfitxer_imp) Then
     obrir_document generarfitxer_imp
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + generarfitxer_imp, vbCritical, "Error"
  End If
End Sub

Sub obrir_pdf_treball(treball As Double, modificacio As Double, Optional vnomfitxerpdf As String)
   Dim generarfitxer_pdf As String
   If modificacio = 0 Then modificacio = 1
   generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + "_SC.pdf"
   If Not existeix(generarfitxer_pdf) Then generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + ".pdf"
   If existeix(generarfitxer_pdf) Then
     If vnomfitxerpdf <> "noobrir" Then obrir_document generarfitxer_pdf
     vnomfitxerpdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + ".pdf"
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + generarfitxer_pdf + Chr(10) + " i tampoc el de separació de colors.", vbCritical, "Error"
  End If
End Sub

Sub veureelpdf(Optional vnomfitxerpdf As String)
  Dim rstc As Recordset
  Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(comanda)))
  obrir_pdf_treball cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), vnomfitxerpdf
  wait 1
  
End Sub
Sub veureelpdf2()
Dim nomfitxer As String
  Dim nomcarpeta As String
  Dim rstc As Recordset
  Dim rstc2 As Recordset
  Dim ruta As String
  Dim ruta_relativa_docs As String
  
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(comanda)))
  If rstc.EOF Then Exit Sub
  nomfitxer = atrim(rstc!arxiupdf)
  If nomfitxer = "" Then Exit Sub
  Set rstc2 = dbtmp.OpenRecordset("select * from carpeta_client where codiclient=" + atrim(cadbl(rstc!client)))
  If Not rstc2.EOF Then nomcarpeta = rstc2!nomcarpeta
  If cadbl(Mid(nomfitxer, 1, 6)) = 0 And nomfitxer <> "" Then nomfitxer = nomcarpeta + Mid(nomfitxer, InStr(1, nomfitxer, "\"))
  ruta = ruta_relativa_docs + "\" + nomfitxer
  If existeix(ruta) Then
     obrir_document ruta
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + ruta, vbCritical, "Error"
  End If

End Sub

Private Sub Command25_Click()
  Me.PopupMenu mllistat
 
End Sub

Private Sub Command26_Click()
  'avisosxrseccio.Show 1
  Dim datainicicomanda As Date
  datainicicomanda = Now
  If Not impresores.Recordset.EOF Then datainicicomanda = IIf(IsDate(impresores.Recordset!datainici), CVDate(impresores.Recordset!datainici), Now)
  'estemfentreprint Or
  If DateDiff("m", "01/10/19", datainicicomanda) > 0 And estemfentreprint Then
     Command9_Click
      botollaunesreprint_Click
     Exit Sub
  End If
     
  
 ' ensenyar_llaunespendentsderetorn cadbl(comanda)
  Unload formaniloxos
  Load formaniloxos
  formaniloxos.tag = Me.comanda
  formaniloxos.boto_nou(0).tag = Command26.tag
  If vestemfentfingerprint Then formaniloxos.checkeditar.tag = "fingerprint"
  formaniloxos.Show 1
  Command26.tag = ""
End Sub
Sub ensenyar_llaunespendentsderetorn(vnumc As Double)
   Dim vsql As String
   Dim rst As Recordset
   vsql = "SELECT impresores_aniloxos.comanda, impresores_retornllaunes.numllauna, impresores_retornllaunes.data "
   vsql = vsql + " FROM impresores_retornllaunes INNER JOIN impresores_aniloxos ON impresores_retornllaunes.idliniadetinta = impresores_aniloxos.id "
   vsql = vsql + " where comanda<>" + atrim(vnumc)
   
   Set rst = dbtmpb.OpenRecordset(vsql)
   If Not rst.EOF Then
      veurelesllaunespendentsderetorn rst
   End If
   Set rst = Nothing
End Sub
Sub veurelesllaunespendentsderetorn(rst As Recordset)
  Load formseleccio
  formseleccio.Data1.DatabaseName = cami
  Set formseleccio.Data1.Recordset = rst
  formseleccio.caption = "Llaunes pendents de retorn"
  formseleccio.refrescar
  formseleccio.Show 1
  Unload formseleccio
End Sub

Private Sub Command26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 And Shift = 0 Then Unload formcanvisanilox: formcanvisanilox.Show 1: Unload formcanvisanilox
  If Button = 2 And Shift = 2 Then Command9_Click
End Sub

Private Sub Command27_Click()
 If Not comprovamaq Then Exit Sub
  If comprovarsidescansorelleu Then Exit Sub
 If Not impresores.Recordset.EOF Then
  impresores.Recordset.MoveLast
  If impresores.Recordset!tipus = "V" Then
      numop = escullir_operari
      nomoperari = UCase(r)
      numop2 = escullir_operari("Escullir AJUDANT d'OPERARI")
      nomoperari2 = "Ajudant: " + UCase(r)
  End If
 End If
 crearseccio "V"
End Sub

Private Sub Command28_Click()
    Me.PopupMenu memail
End Sub
Sub enviar_email_oficines()
   Dim msg As String
   msg = InputBox("Escriu el missatge que vols enviar a oficines.", "E-mail a oficines")
   If atrim(msg) <> "" Then
        enviaremailgeneric "missatgesgenericsimpresores", "Missatge genèric d'impressores.  " + nommaq + " - " + nomoperari, treure_apostruf(msg)
        MsgBox "Missatge enviat.", vbInformation, "Eviar missatge genèric"
   End If
End Sub
Sub enviar_email_encarregat()
   Dim msg As String
   msg = InputBox("Escriu el missatge que vols enviar a l'Encarregat.", "E-mail ENCARREGAT")
   If atrim(msg) <> "" Then
        enviaremailgeneric "impresores@inplacsa.com", "Missatge a l'ENCARREGAT d'impressores.  " + nommaq + " - " + nomoperari, treure_apostruf(msg)
        MsgBox "Missatge enviat.", vbInformation, "Eviar missatge a l'ENCARREGAT"
   End If
End Sub

Function avisnohihatotselslots() As Boolean
    Dim vcont As Byte
    vcont = 5
    While vcont >= 1
       If MsgBox("No hi ha tots els lots entrats." + vbNewLine + "No pots continuar sense entrar-los", vbInformation + vbOKCancel, "Atenció (" + atrim(vcont) + ")") = vbOK Then avisnohihatotselslots = True: Exit Function
       vcont = vcont - 1
    Wend
    
End Function
Function nohihatotselslots() As Boolean
   Dim i  As Byte
   Dim vmsg As String
   If estemfentreprint Then
        If llistallaunesreprint.ListCount = 0 Or cadbl(ettotalkgreprint) = 0 Then nohihatotselslots = True
       Else
         For i = 0 To 7
            If pantone(i) <> "" And atrim(compantone(i)) = "" Then vmsg = "La tinta " + pantone(i) + " no te lots." + vbNewLine
         Next i
   End If
   If vmsg <> "" Then nohihatotselslots = True
End Function

Private Sub Command3_Click()
 Dim mtrsprova As String
 Dim mtrsparcials As Double
 Dim opantic As Byte
 Dim idbobina As Long
 Dim ensenyaraniloxalfinal As Boolean
 Dim vidnovaseccio As Double
 Dim vmetresarrancadavisual As String
 
 
 If Not comprovamaq Then Exit Sub
 comprovarsidescansorelleu
 'comprovo si tots els lots estan entrats i si no aviso i surto
 If nohihatotselslots Then
    If avisnohihatotselslots Then Exit Sub
 End If
 If Not impresores.Recordset.EOF Then
    impresores.Recordset.MoveLast
    If impresores.Recordset!tipus = "A" Then
        'mtrsprova = InputBox("Entra els Metres de prova.", "Atenció")
        impresores.Recordset.FindLast "tipus='A'"
        
        If Not impresores.Recordset.NoMatch Then
         'impresores.Recordset.Edit
         'impresores.Recordset!mtrsprova = cadbl(mtrsprova)
         
        ' impresores.Recordset.Update
           If Not sharevisatlasortidaalabobinazero(cadbl(comanda)) Then MsgBox "No pots fer funcionament sense haver revisat la sortida de la bobina zero." + vbNewLine + "Revisala a la segona pantalla.", vbCritical, "Error": Exit Sub
           If cadbl(impresores.Recordset!mtrsprova) = 0 Then
            MsgBox "No hi han les dades de les bobines de prova.", vbCritical, "Atenció"
            Exit Sub
           End If
           mirar_excesdemetresdeprova
        End If
'        veureelpdf
        MsgBox "REVISA QUE EL PDF COINCIDEIX AMB L'IMPRESIÓ QUE ESTAS FENT.", vbExclamation, "A T E N C I Ó"
        vmetresarrancadavisual = demanar_metres_arrancada
        ensenyaraniloxalfinal = True
    End If
    Else:
       If vestemfentfingerprint Then GoTo crearseccioF
       Exit Sub
 End If
 comprovar_reducciocilindre
 While panellavis.visible
   DoEvents
 Wend
 firmar_fulla
 If impresores.Recordset!tipus = "F" Then
 
    opantic = numop
    numop = escullir_operari
    nomoperari = UCase(r)
    numop2 = escullir_operari("Escullir AJUDANT d'OPERARI")
    nomoperari2 = "Ajudant: " + UCase(r)
     
 End If
 If Not bobines.Recordset.EOF Then
   bobines.Recordset.MoveLast
   If cadbl(bobines.Recordset!metres) = 0 Then
     mtrsprova = InputBox("Entra els metres parcials de bobina.", "Bobina no acabada")
     If cadbl(mtrsprova) <> 0 Then
        mtrsparcials = cadbl(mtrsprova)
        impresores.Recordset.Edit
        impresores.Recordset!metresparcial = mtrsparcials
        impresores.Recordset.Update
        bobines.Recordset.Edit
        bobines.Recordset!metresparcial = mtrsparcials
        bobines.Recordset!operari1 = numop
        bobines.Recordset!operari2 = opantic
        bobines.Recordset.Update
        idbobina = bobines.Recordset!id
     End If
   End If
 End If
 
crearseccioF:
 If Not sharevisatlasortidaalabobinazero(cadbl(comanda)) Then MsgBox "No pots fer funcionament sense haver revisat la sortida de la bobina zero." + vbNewLine + "Revisala a la segona pantalla.", vbCritical, "Error": Exit Sub
 crearseccio "F", vmetresarrancadavisual
 vidnovaseccio = cadbl(r)
   'And Not hihalabobina1creada(cadbl(comanda))    'en princi ho havia ficat al proxim IF per quan hi havia avaria no tornes a fer la VQ
   'però al fer funcionament i canvi de operari no ho demanava fent-ho així i no ho volen i ho tret per tot
 If Not vestemfentfingerprint Then   'si es fingerprint no ha de verificar ni imprimir etiqueta
    verificar_tubbase
    imprimir_controlqualitat cadbl(comanda), numop, 0
 End If
 dbtmpb.Execute "update bobinesimp set controlid=" + atrim(vidnovaseccio) + " where id=" + atrim(idbobina)  'paso la bobina del funcionament anterior a aquest funcionamnent nou
 mtrsparcials = 0
 impresores.Recordset.MoveLast
 While bobines.Recordset.RecordCount = 0 And mtrsparcials < 200
   DoEvents
   bobines.Refresh
   mtrsparcials = mtrsparcials + 1
 Wend
 mourealnouFsilabobina1siestasensemetres vidnovaseccio, cadbl(comanda)
 'If ensenyaraniloxalfinal Then formaniloxos.Show 1
End Sub
Function sharevisatlasortidaalabobinazero(vnumc As Double) As Boolean
   Dim rst As Recordset
   Dim dbbaixesannex As Database
   Set dbbaixesannex = OpenDatabase(rutadelfitxer(cami) + "baixes_annex.mdb")
   Set rst = dbbaixesannex.OpenRecordset("select * from RevisioCQ where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
      If rst!imp_verificat Then sharevisatlasortidaalabobinazero = True
   End If
   Set rst = Nothing
   Set dbbaixesannex = Nothing
End Function
Function demanar_metres_arrancada() As String
  Dim vpalet As Double
  Dim vbobina As Double
  Dim vmetres As Double
  
  demanar_paletibobina vpalet, vbobina
  If vpalet = 0 Or vbobina = 0 Then MsgBox "Aquest palet no es vàlid", vbCritical, "Error": Exit Function
  While vmetres < 100 Or vmetres > 999
     vmetres = InputBox("Entre els metres d'arrancada que has utilitzat." + Chr(13) + ">100m i <1000m", "Metres arrancada")
  Wend
  dbstocks.Execute "insert into parcials (idpalet,idbobina,metres,comanda,data,seccio,utilitzada,orcomassignacio,operari,observacions) values (" + atrim(vpalet) + "," + atrim(vbobina) + "," + atrim(vmetres) + ",'" + atrim(cadbl(form1.comanda)) + "',now,'" + lletraseccio + "',true,500," + atrim(numop) + ",'#arrancada')"
  demanar_metres_arrancada = "[A]" + atrim(vpalet) + "/" + atrim(vbobina) + "-" + atrim(vmetres) + "m"
  vmetresarrancada = vmetres
End Function
Function hihalabobina1creada(vnumc As Double) As Boolean
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("SELECT impressores.comanda, bobinesimp.numerodebobina, bobinesimp.metres, bobinesimp.controlid FROM bobinesimp INNER JOIN impressores ON bobinesimp.controlId = impressores.Id WHERE comanda=" + atrim(vnumc) + " and (((bobinesimp.numerodebobina)=1));")
  If Not rst.EOF Then hihalabobina1creada = True
  Set rst = Nothing
End Function
Sub mourealnouFsilabobina1siestasensemetres(vidnovaseccio As Double, vnumc As Double)
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("SELECT impressores.comanda, bobinesimp.numerodebobina, bobinesimp.metres, bobinesimp.controlid FROM bobinesimp INNER JOIN impressores ON bobinesimp.controlId = impressores.Id WHERE comanda=" + atrim(vnumc) + " and (((bobinesimp.numerodebobina)=1));")
   If Not rst.EOF Then
       If rst!metres = 0 Then  'si no hi ha metres es que la bobina està acabada i la mouré a la nova secció
          rst.Edit
          rst!controlid = vidnovaseccio
          rst.Update
       End If
   End If
  Set rst = Nothing
  bobines.Refresh
End Sub
Sub imprimir_controlqualitatVQ(numc As Double, Optional imprimirlo As Boolean, Optional vreimpres As Boolean)
   Dim rst As Recordset
   Dim rstm As Recordset
   Dim ultimalinia As String
   Dim vnumtreball As String
   Set rst = dbtmp.OpenRecordset("select impressio,refclient,numtreball,numordremodificacio from comandes where comanda=" + atrim(numc))
   If Not rst.EOF And Not imprimirlo Then
        If atrim(rst!impressio) <> "N" And atrim(rst!impressio) <> "M" Then Exit Sub
   End If
   If Not rst.EOF Then vnumtreball = atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio)
   ultimalinia = "Op: " + atrim(numop) + "    NºBob.Salida: 0   Fecha: " + Format(Now, "dd/mm/yy")
   For i = 0 To 100
     llistat.Formulas(i) = ""
   Next i
   Set rstm = dbtmp.OpenRecordset("select descripcio from maquines where maquina='I' and codi=" + atrim(nummaq))
   If rstm.EOF Then Exit Sub
   llistat.Formulas(0) = "lot=" + atrim(numc)
   llistat.Formulas(1) = "ultimalinia='" + atrim(ultimalinia) + "'"
   llistat.Formulas(2) = "data='" + Format(Now, "dd/mm/yy") + " " + IIf(vreimpres, "(R)'", "'")
   llistat.Formulas(3) = "referencia='" + atrim(texteimpresio) + "'"
   llistat.Formulas(4) = "client='" + atrim(client) + "'"
   llistat.Formulas(5) = "nommaquina='" + atrim(rstm!descripcio) + "'"
   llistat.Formulas(6) = "treball='" + atrim(vnumtreball) + "'"
   llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "verificacioqualitatVQimpresores.rpt"
   llistat.Destination = crptToPrinter
    llistat.CopiesToPrinter = 1
   llistat.DataFiles(0) = ""
   llistat.DiscardSavedData = True
   escullir_impresora_tickets
   DoEvents
   If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
   llistat.Action = 1
   
   Set rst = Nothing
   llistat.PrinterDriver = ""
   llistat.PrinterName = ""
   llistat.PrinterPort = ""
End Sub
Sub escullir_impresora_tickets()
  Dim X As Printer
  For Each X In Printers
     If InStr(1, UCase(X.DeviceName), "80 PRINTER") > 0 Then GoTo cont
  Next
  MsgBox "No he trobat la impresora de tickets [80 Printer]", vbCritical, "Error": GoTo fi
cont:
 ' If InStr(1, UCase(x.DeviceName), "80 PRINTER") > 0 Then MsgBox "No he trobat la impresora de tickets [80 Printer]", vbCritical, "Error": GoTo fi
  llistat.PrinterDriver = X.DriverName
  llistat.PrinterName = X.DeviceName
  llistat.PrinterPort = X.Port
  
fi:
End Sub
Sub firmar_fulla()
 If atrim(firmat) = "" Then
    Do
    firmat = InputBoxEx("Entra el codi d'operari o contrasenya que firma la fulla", "Atenció", , , , , , SPassword)
    If cadbl(firmat) = 1 Then MsgBox "Aquest operari ha d'apuntar la contrasenya."
    Loop Until cadbl(firmat) <> 1
    
    
    If cadbl(firmat) = 0 Then
       If LCase(firmat) = "jmok" Then
          firmat = "1"
         Else: firmat = ""
       End If
    End If
    
    Set rsttmp = dbtmp.OpenRecordset("select codi from operaris where actiu=1 and maquina='I' and codi=" + atrim(cadbl(firmat)))
    If rsttmp.EOF Then firmat = "": MsgBox "Aquest operari no existeix"
    guarda_totals
    passarcomandaacomençada
 End If
End Sub
Sub passarcomandaacomençada()
 dbtmp.Execute "update comandes set seccioactual='I' where comanda=" + atrim(comanda)
End Sub
Sub corretgirerrorbdestoc()
  dbstocks.Execute "update  parcials set orcomassignacio='0' where (orcomassignacio is null) or (orcomassignacio='')"
  'dbstocks.Execute "update table parcials set comanda='-1' where comanda=nul or comanda=''"
End Sub
Sub carregar_opcions_stock()
  Dim grupajust As Double
  veuregrupsdestoc.tag = ""
  vgrupmaterialcompatible = 0
  obrestocks
    ' CORRETGEIXO ERRORS DE LA BASE DE DADES D'ESTOC
  corretgirerrorbdestoc
  etmetresajust = textedajust(cadbl(comanda), grupajust)
  veuregrupsdestoc.tag = atrim(grupajust)
  If grupajust > 0 Then
    Set rsttmp = dbstocks.OpenRecordset("select codigrupmaterialscompatibles from grupsdepalets where numerogrup=" + atrim(grupajust))
    vgrupmaterialcompatible = cadbl(rsttmp!codigrupmaterialscompatibles)
  End If
  Set rsttmp = dbtmp.OpenRecordset("select codigrupmaterialcompatible,assignarstock from comandes_extres where comanda=" + atrim(cadbl(comanda)))
  etmetresajust.tag = ""
  If Not rsttmp.EOF Then
      If vgrupmaterialcompatible = 0 Then vgrupmaterialcompatible = cadbl(rsttmp!codigrupmaterialcompatible)
      If rsttmp!assignarstock Then
        etmetresajust = "-- STOCK --  " + etmetresajust
        etmetresajust.tag = "STOCK"
      End If
  End If
  Set dbstocks = Nothing
  Set rsttmp = Nothing
End Sub
Function textedajust(numc As Double, grupdestoc As Double) As String
  Dim rstopcions As Recordset
  Dim rstgrup As Recordset
  Dim t As String
  Dim sisaj As Byte
  
   Set rstopcions = dbstocks.OpenRecordset("select * from opcionsdajust where comanda=" + atrim(numc))
   If Not rstopcions.EOF Then
     sisaj = atrim(cadbl(rstopcions!sistemadajust))
     grupdestoc = cadbl(rstopcions!grupdestoc)
     If sisaj > 0 Then
      t = atrim(cadbl(rstopcions!mtrsajust)) + " Mtrs D'AJUST.  "
      If sisaj = 1 Then t = t + " S'HA D'UTILITZAR MATERIAL PER LLENÇAR."
      If sisaj = 2 Then
        If cadbl(rstopcions!grupdestoc) > 0 Then
           Set rstgrup = dbstocks.OpenRecordset("select numerogrup,nomdelgrup from grupsdepalets where numerogrup=" + atrim(cadbl(rstopcions!grupdestoc)))
           If Not rstgrup.EOF Then
            t = t + " S'HA D'UTILITZAR MATERIAL D'ESTOC DEL " + UCase(rstgrup!nomdelgrup)
            grupdestoc = cadbl(rstgrup!numerogrup)
           End If
        End If
      End If
      If sisaj = 3 And cadbl(rstopcions!paletajust) > 0 Then t = t + " S'HA D'UTILITZAR EL PALET " + atrim(rstopcions!paletajust) + " BOB: " + atrim(rstopcions!bobinaajust)
      
      
     End If
   End If
   textedajust = t
End Function


Function estamuntada(numc As Double) As Boolean
   Dim rst As Recordset
   estamuntada = False
   Set rst = dbtmpb.OpenRecordset("SELECT muntadoratot.comanda  FROM comandes INNER JOIN muntadoratot ON comandes.comanda = muntadoratot.comanda WHERE (((muntadoratot.acabada)=True) AND ((comandes.proximaseccio)='I') and muntadoratot.comanda=" + atrim(cadbl(numc)) + ");")
   If Not rst.EOF Then estamuntada = True
   If cadbl(llegir_ini("Baixes", "ultimacomanda", "comandes.ini")) = numc Then estamuntada = True
End Function


Sub carregaravisosmanteniment(tancarla As Boolean)
   Load avisosxrseccio
   avisosxrseccio.seccio.text = "Impresores"
   avisosxrseccio.nommaquina = nommaq
   avisosxrseccio.nommaquina.tag = nummaq
   'avisosxrseccio.datafi = "01/03/15"
   avisosxrseccio.buscaravisos
   'If Not avisosxrseccio.datamanteniments.Recordset.EOF Then
   '    botoavisos.BackColor = QBColor(12)
   '    botoavisos.tag = "avis"
   '   Else: botoavisos.BackColor = Command25.BackColor: botoavisos.tag = ""
   'End If
   If tancarla Then Unload avisosxrseccio
End Sub
Function posicioenlaruta(numc As Double) As String
  Dim rstp As Recordset
  Dim rstpr As Recordset
  Dim laruta As String
   
  'If InStr(1, "VPT", seccioactual) = 0 Then Exit Function
  Set rstp = dbtmpb.OpenRecordset("SELECT comandes.comanda,comandes.proximaseccio,comandes.producte, rebobinadorestot.acavada as acavadar, laminadorestot.acavada as acavadal, impressorestot.acavada as acavadai FROM ((comandes LEFT JOIN rebobinadorestot ON comandes.comanda = rebobinadorestot.comanda) LEFT JOIN laminadorestot ON comandes.comanda = laminadorestot.comanda) LEFT JOIN impressorestot ON comandes.comanda = impressorestot.comanda WHERE (((comandes.comanda)=" + atrim(numc) + "));")
  If Not rstp.EOF Then
     Set rstpr = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(rstp!producte) + "'")
     If rstpr.EOF Then Exit Function
     laruta = atrim(rstpr!ruta)
     If InStr(1, laruta, "R") > 0 And cadblnull_1(rstp!acavadar) = 0 Then posicioenlaruta = "R"
     If InStr(1, laruta, "L") > 0 And cadblnull_1(rstp!acavadal) = 0 Then posicioenlaruta = "L"
     If InStr(1, laruta, "I") > 0 And cadblnull_1(rstp!acavadai) = 0 Then posicioenlaruta = "I"
  End If
  If posicioenlaruta = "" Or atrim(rstp!proximaseccio) = "E" Then posicioenlaruta = rstp!proximaseccio
  
  Set rstp = Nothing
  Set rstpr = Nothing
End Function
Function cadblnull_1(acabada As Variant) As Double
   If IsNull(acabada) Then cadblnull_1 = -1: Exit Function
   cadblnull_1 = cadbl(acabada)
End Function
Function comandavalida(numc As Double, msg As String, Optional nocomprovarllista As Boolean) As Boolean
   Dim rst As Recordset
   msg = ""
   comandavalida = False
   If numc = 0 Then Exit Function
   If Not nocomprovarllista Then
     Set rst = dbbaixes.OpenRecordset("select * from muntadora_ordremuntatge where comanda=" + atrim(numc))
     If Not rst.EOF Then MsgBox "La comanda " + atrim(numc) + " ja està a la llista.": Exit Function
   End If
   Set rst = dbtmp.OpenRecordset("SELECT comandes.comanda, productes.ruta, comandes.proximaseccio,comandes.impressio FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(numc) + "));")
   If Not rst.EOF Then
       proximaseccio = posicioenlaruta(numc)
       If proximaseccio = "I" And InStr(1, rst!ruta, "I") > 0 Then
             comandavalida = True
               Else
                 If InStr(1, rst!ruta, "I") = 0 Then
                    'MsgBox "La comanda " + atrim(numc) + " no te seccio d'impresores"
                    msg = msg + "La comanda " + atrim(numc) + " no te seccio d'impresores" + Chr(10)
                 End If
                 If proximaseccio <> "I" Then
                   'MsgBox "La comanda " + atrim(numc) + " no està apunt per imprimir. La ruta no està a I."
                   msg = msg + "La comanda " + atrim(numc) + " no està apunt per imprimir. La seccio actual no es Impresores està amb [ " + atrim(proximaseccio) + "]" + Chr(10)
                 End If
                 
       End If
       If rst!impressio = "F" Then
           'MsgBox "A la comanda " + atrim(numc) + " li Falta Autoritzar.", vbCritical, "Atenció"
           msg = msg + "A la comanda " + atrim(numc) + " li Falta Autoritzar." + Chr(10)
           comandavalida = False
       End If
           
        Else:  msg = msg + "La comanda " + atrim(numc) + " no existeix." + Chr(10)
          'MsgBox "La comanda " + atrim(numc) + " no existeix."
   End If
   If comandavalida Then
      If Not tepackinglist(cadbl(numc)) Then
         'MsgBox "Aquesta comanda encara no te material assignat.", vbCritical, "Atenció"
         msg = msg + "Aquesta comanda encara no te material assignat." + Chr(10)
         comandavalida = False
      End If
   End If
   If comandavalida Then
        If Not clixesentratsafabrica(cadbl(numc)) Then
          comandavalida = False
          msg = msg + "La comanda " + atrim(numc) + " no te els CLIXES ENTRATS a disseny. No es poden utilitzar."
        End If
   End If
   'If comandavalida Then
    Set rst = dbtmp.OpenRecordset("select passaraimpresores,clientvindraarevisarimpresio from comandes_extres where comanda=" + atrim(numc))
    If Not rst.EOF Then
      If rst!clientvindraarevisarimpresio Then
       msg = msg + "ALERTA!!!" + Chr(10) + "AQUESTA COMANDA NECESSITA OK DEL CLIENT." + Chr(10)
       comandavalida = False
      End If
      If cadbl(rst!passaraimpresores) = 0 Then
       msg = msg + "ALERTA!!!" + Chr(10) + "AQUESTA COMANDA ESTÀ EN STANDBY, NO POTS UTILITZAR-LA." + Chr(10)
       comandavalida = False
      End If
    End If
   'End If
   
End Function
Function clixesentratsafabrica(numc As Double) As Boolean
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim rutaclixes As String
   Dim dbclixes As Database
   Dim ordrem As Integer
   clixesentratsafabrica = False
   rutaclixes = rutadelfitxer(cami) + "clixesnous.mdb"
   'rutaclixes = rutadelfitxer(cami) + "clixes.mdb"
   Set dbclixes = OpenDatabase(rutaclixes)
   Set rstc = dbtmp.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(numc))
   If rstc.EOF Then Exit Function
   ordrem = cadbl(rstc!numordremodificacio)
   If ordrem = 0 Then ordrem = 1
   Set rst = dbclixes.OpenRecordset("select id_estatclixe from clixes_modifi where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordremodificacio=" + atrim(ordrem) + " order by ordre DESC")
   If rst.EOF Then Exit Function
   If rst!id_estatclixe = 8 Then clixesentratsafabrica = True
   Set rst = Nothing
   Set rstc = Nothing
   Set dbclixes = Nothing
End Function
Function tepackinglist(numc As Double) As Boolean
   Dim rstt As Recordset
   obrestocks
   tepackinglist = False
   Set rstt = dbstocks.OpenRecordset("select * from parcials where  comanda='" + atrim(numc) + "'")
   If Not rstt.EOF Then tepackinglist = True
   Set rstt = dbtmp.OpenRecordset("select assignarstock from comandes_extres where comanda=" + atrim(numc))
   If Not rstt.EOF Then
      If rstt!assignarstock Then tepackinglist = True
   End If
   Set dbstocks = Nothing
End Function

Private Sub Command30_Click()
    Dim vtotal As Double
    Dim rst As Recordset
    vtotal = cadbl(InputBox("Entra el total de KG gastats.", "Total KG"))
    If vtotal > 0 Then
       dbtmpb.Execute "update impresores_llaunesgastades set kg=0 where tipus='R' and comanda=" + atrim(comanda)
       Set rst = dbtmpb.OpenRecordset("select * from impresores_llaunesgastades where tipus='R' and comanda=" + atrim(comanda))
       If Not rst.EOF Then
         rst.Edit
         rst!kg = vtotal
         rst.Update
         carregarllistadellaunesreprint
       End If
    End If
    Set rst = Nothing
           
End Sub

Private Sub Command31_Click()
  Dim vllauna As String
  Dim rst As Recordset
  Dim vkg As Double
  vllauna = "-"
  vkg = 0
  Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb", , True)
  While atrim(vllauna) <> ""
        vllauna = atrim(InputBox("Entra el numero de llauna que vols afegir al REPRINT.", "Llauna Reprint"))
        If vllauna <> "" Then
        '        vkg = cadbl(InputBox("Quants Kg has possat d'aquesta llauna?", "Kg gastats"))
        '       If vkg = 0 Then GoTo fi
            If lallaunaescorrecte(vllauna, etvernis.tag) Then
               Set rst = dbtmpb.OpenRecordset("select * from impresores_llaunesgastades where numllauna='" + atrim(vllauna) + "' and comanda=" + atrim(comanda))
               If rst.EOF Then
                   dbtmpb.Execute "insert into impresores_llaunesgastades (numllauna,comanda,tipus,kg) values ('" + atrim(vllauna) + "'," + atrim(comanda) + ",'R'," + atrim(vkg) + ")"
                   carregarllistadellaunesreprint
               End If
            End If
        End If
  Wend
fi:
  Set rst = Nothing
  'Set dbtintes = Nothing
  carregarllistadellaunesreprint
End Sub
Function lallaunaescorrecte(vllauna As String, vcodivernis As String) As Boolean
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from dadesllaunes where codi in (" + atrim(vcodivernis) + ") and numllauna='" + atrim(vllauna) + "'")
   If rst.EOF Then
      v = "-"
      While v <> ""
        v = InputBox("AQUESTA LLAUNA NO COINCIDEIX AMB EL VERNIS DE LA COMANDA." + Chr(10) + " NUMERO DE LLAUNA: " + vllauna, "ERROR DE VERNIS")
      Wend
      Exit Function
       Else: lallaunaescorrecte = True
   End If
   Set rst = Nothing
End Function
Sub carregarllistadellaunesreprint()
  Dim rst As Recordset
  Dim vtotalkg As Double
  llistallaunesreprint.Clear
  Set rst = dbtmpb.OpenRecordset("select * from impresores_llaunesgastades where tipus='R' and numllauna<>'' and comanda=" + atrim(comanda))
  While Not rst.EOF
    llistallaunesreprint.AddItem UCase(atrim(rst!numllauna)) + "-->" + atrim(rst!kg) + " Kg"
    vtotalkg = vtotalkg + cadbl(rst!kg)
    rst.MoveNext
  Wend
  ettotalkgreprint = atrim(vtotalkg)
  Set rst = Nothing
End Sub

Private Sub Command32_Click()
   Dim vnumllauna As String
    If llistallaunesreprint.ListIndex = -1 Then MsgBox "Primer has d'escullir una llauna de la llista.", vbInformation, "Atenció": Exit Sub
   If MsgBox("Segur que vols eliminar la llauna " + atrim(llistallaunesreprint) + "?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
       vnumllauna = Mid(atrim(llistallaunesreprint), 1, InStr(1, llistallaunesreprint, "-->") - 1)
       dbtmpb.Execute "delete * from impresores_llaunesgastades where comanda=" + atrim(comanda) + " and numllauna='" + vnumllauna + "'"
      ' wait 2
       carregarllistadellaunesreprint
   End If
End Sub

Private Sub Command33_Click()
  Dim vidtreball As Integer
  vidtreball = cadbl(InputBox("Entra el treball que vols consultar.", "Consultar observacions", atrim(ettreball.tag)))
  If vidtreball > 0 Then observacio_idtreball vidtreball
End Sub
Sub possarobservaciooperari(vnumtreball As Double)
   Dim rst As Recordset
   cobservacionsoperari.text = ""
   Set rst = dbtmpb.OpenRecordset("select * from idstreball where id=" + atrim(vnumtreball))
   If rst.EOF Then Exit Sub
   cobservacionsoperari.text = rst!obsidtreball
   Set rst = Nothing
End Sub

Private Sub Command34_Click()
  Dim vbob As String
   Dim rstb As Recordset
   Dim vpalet As String
   vbob = InputBox("Entra el Palet/Bobina que vols localitzar al magatzem (Situació):", "Localitzar la situació d'una bobina" + Chr(10) + "Ex: 99999/10")
   If InStr(1, vbob, "/") = 0 Then MsgBox "No hi ha el format correcte de bobina. Ex: 99999/11  Palet/Bobina", vbCritical, "Error": GoTo fi
   vpalet = cadbl(Mid(" " + vbob, 1, InStr(1, vbob + "  ", "/")))
   vbob = cadbl(Mid(vbob, InStr(1, vbob + "  ", "/") + 1))
   Set rstb = dbstocks.OpenRecordset("select sit from bobines where idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbob))
   If Not rstb.EOF Then
        MsgBox "La situació de la bobina " + vpalet + "/" + vbob + " es:  " + atrim(rstb!sit), vbInformation, "Palet localitzat"
         Else: MsgBox "No he localitzat aquesta bobina", vbCritical, "Error"
   End If
fi:
   Set rstb = Nothing
End Sub

Private Sub Command35_Click()
   Me.PopupMenu mmenuarrancariVQ
   
End Sub

Private Sub Command36_Click()
  formbobinesaimpresores.Show 1
End Sub

Private Sub Command37_Click()
   Load Formdesbobinadors
   Formdesbobinadors.nomoperari = form1.nomoperari
   Formdesbobinadors.Show 1
End Sub

Private Sub Command38_Click()
   carregar_llista_ordreimpressio
   actualitzar_comandaactiva_compartida
End Sub

Private Sub Command4_Click()
  Dim rst As Recordset
  Dim ruta As String
  Dim msg As String
  Dim resp As String
  Dim vmaterialexacte As Boolean
  Dim v As Double
  Dim vnomfitxerpdf As String
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  'Set ImagePDF.Picture = LoadPicture()
  If cadbl(comanda) < 100000 Then comanda = 0
  carregaravisosmanteniment True
  mantenimentbobina.comprovarnivellsdestoc
  estemfentreprint = False
  botollaunesreprint.visible = False
  etvernis.caption = ""
  etvernis.tag = ""
  firmat = ""
  codibarras = ""
  vvalidaciocodidebarres = ""
  'comprovo si la comanda està muntada
  If r <> "nopregunta" And Not vnopreguntar Then
    If Not estamuntada(cadbl(comanda)) Then
       If atrim(InputBox("Aquesta comanda no està entrada com a muntada o ja està IMPRESA." + Chr(10) + "Si es correcte escriu el Nº COMANDA si no fes CANCELAR", "NO MUNTAT")) <> comanda Then Exit Sub
       
    End If
    
    If Not comandavalida(cadbl(comanda), msg, True) Then
          '"Aquesta comanda ESTÀ PARADA O HI HA ALGUN MOTIU PER PARAR-LA."
        If InStr(1, UCase(msg), "STANDBY") > 0 Then MsgBox msg, vbCritical, "Atenció": comanda = "0": Exit Sub
        If InStr(1, UCase(msg), "FALTA AUTORITZAR") > 0 Then MsgBox msg, vbCritical, "Atenció": comanda = "0": Exit Sub
        If MsgBox(msg + Chr(10) + "VOLS CONTINUAR IGUALMENT?", vbCritical + vbYesNo + vbDefaultButton2, "ATENCIÓ") = vbNo Then Exit Sub
    End If
  End If
  'carrego l'annex
  formannex.carregarcomanda cadbl(comanda)
  'comprovo si existeix la comanda
  Set rsttmp = dbtmp.OpenRecordset("select producte,codibarras,impressora,espessor,comanda,refclient,comandaclient,texteimpressio,linkcomanda1,linkcomanda2,tubbaseimp from comandes where comanda=" + atrim(cadbl(comanda)))
  If Not rsttmp.EOF Then
  '//tret per ordre de lencarregat i en miralles
  '  If cadbl(rsttmp!linkcomanda1) <> 0 Or cadbl(rsttmp!linkcomanda2) <> 0 And r <> "nopregunta" And Not vnopreguntar Then
  '   While UCase(resp) <> "OK" And r <> "nopregunta" And Not vnopreguntar
  '      resp = InputBox("Aquesta comanda serà per LAMINAR tingues compte que les tintes siguin per LAMINACIÓ !!!" + Chr(10) + "Escriu OK per acceptar.")
  '   Wend
  '  End If
    vtubbase = cadbl(rsttmp!tubbaseimp)
  End If
  If rsttmp.EOF Or cadbl(comanda) = 0 Then
      MsgBox "No hi ha numero de comanda vàlida"
      Command1.Enabled = False:   Command2.Enabled = False:   Command3.Enabled = False: Exit Sub
        Else:
          ruta = ""
          Set rsttmp = dbtmp.OpenRecordset("select ruta from productes where codi='" + rsttmp!producte + "'")
          If Not rsttmp.EOF Then ruta = rsttmp!ruta
          
          If InStr(1, ruta, "I") = 0 Then
            MsgBox "Aquesta comanda no te secció d'IMPRESORES"
            Command1.Enabled = False:   Command2.Enabled = False:   Command3.Enabled = False: Exit Sub
          End If
  End If
  mirar_missatgeXroperaris cadbl(comanda) 'Si te un missatge assignat l'ensenyo
  escriure_ini "Baixes", "ultimacomanda", 0, "comandes.ini"
  ncomanda2 = cadbl(comanda): ncomanda = cadbl(comanda)
  comprovarsitepreuassignatosinoenviarunmail cadbl(comanda)
  ensenya_totals
  calcular_totals
  carregar_opcions_stock
  ncomanda = cadbl(comanda)
  bobines.RecordSource = "select * from bobinesimp where controlid=-1"
  bobines.Refresh
  
  'miro si es stock o packing
  stockopacking = "P"
  Set rst = dbtmp.OpenRecordset("SELECT comandes_extres.assignarstock as estoc, materialexacte frOM comandes_extres WHERE comanda=" + atrim(cadbl(comanda)) + ";")
  If Not rst.EOF Then
     If rst!estoc Then stockopacking = "E"
     If rst!materialexacte Then vmaterialexacte = True
  End If
  
  
  Set rsttmp = dbtmp.OpenRecordset("select marcailinia,materialex,cantitatex,mesuracantex,cmaquina,numtreball,numordremodificacio,codibarras,tubolam,espessor,mesuraesp,comanda,refclient,comandaclient,texteimpressio,direnvio,impressio from comandes where comanda=" + atrim(cadbl(comanda)))
  avisapantalla = ""
  direnvio = cadbl(rsttmp!direnvio)
  If atrim(rsttmp!cmaquina) <> "" And cadbl(rsttmp!cmaquina) > 0 Then avisapantalla = "Red.Cilindre: " + rsttmp!cmaquina
  calcularvalorsreducciocilindre cadbl(comanda), nummaq, 0, llistat
  mesuraespcomanda = ""
  If Not rsttmp.EOF Then
     Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rsttmp!mesuraesp)))
     If Not rsttmp2.EOF Then mesuraespcomanda = rsttmp2!descripcio
  End If
  vtipusimpresio = atrim(rsttmp!impressio)  'serveix per ensenyar o no el bloqueig de aniloxos
  '//tret per ordre de lencarregat i en miralles
  'If (atrim(rsttmp!impressio) = "N" Or atrim(rsttmp!impressio) = "M") And Not vnopreguntar Then
  '   MsgBox "IMPRESSIÓ NOVA O MODIFICADA " + Chr(10) + "PENSA A RECOLLIR MOSTRA PER MK", vbOKOnly, "Atenció"
  'End If
  
  
  id_treball = cadbl(rsttmp!numtreball)
  ordremodificacio = cadbl(rsttmp!numordremodificacio)
  metrescomanda = cadbl(rsttmp!cantitatex)
  If cadbl(rsttmp!mesuracantex) <> 1 Then metrescomanda = 999999
  ettreball = "Nº Treball: " + atrim(cadbl(rsttmp!numtreball))
  ettreball.tag = atrim(cadbl(rsttmp!numtreball))
  observacio_idtreball cadbl(ettreball.tag), True
  refclient = "": comandaclient = ""
  texteimpresio = ""
  refclient = atrim(rsttmp!refclient)
  comandaclient = atrim(rsttmp!comandaclient)
  possarobservaciooperari cadbl(rsttmp!numtreball)
  possartoleranciaample cadbl(rsttmp!numtreball), cadbl(rsttmp!numordremodificacio), rsttmp!materialex, atrim(rsttmp!cmaquina)
   'clixes.Enabled = True
  texteimpresio = IIf(atrim(rsttmp!marcailinia) = "", atrim(rsttmp!texteimpressio), atrim(rsttmp!marcailinia))
  micrescomanda = cadbl(rsttmp!espessor)
  If InStr(1, mesuraespcomanda, "GALGUES") > 0 Then micrescomanda = micrescomanda / IIf(rsttmp!tubolam = "L", 2, 4)
  codibarras = atrim(rsttmp!codibarras)
  
  Command1.Enabled = True: Command2.Enabled = True: Command3.Enabled = True
  
  
  'fins aqui comprovo comanda
  impresores.RecordSource = "select * from impressores where comanda=" + atrim(cadbl(comanda)) + " order by datainici,horainici"
  imppantones.RecordSource = "select * from impresorespantones where comanda=" + atrim(cadbl(comanda))
  impresores.Refresh
  imppantones.Refresh
  'comprova el material exacte
  
  If vmaterialexacte Then
     If impresores.Recordset.EOF And Not vnopreguntar Then MsgBox "Aquesta comanda es de MATERIAL ESPECÍFIC, s'ha de utilitzar sempre el MATERIAL EXACTE que demana el client.", vbCritical, "Atenció"
     etmaterialexacte.tag = rsttmp!materialex
     carregar_materialexacte
       Else: etmaterialexacte.tag = "": etmaterialexacte = ""
  End If
  
  '//tret per ordre de lencarregat i en miralles
  'If mirarsihihaCingularReal(cadbl(rsttmp!numtreball), cadbl(rsttmp!numordremodificacio)) And Not vnopreguntar Then MsgBox "Atenció aquest treball té un PDF de Cingular Real2", vbInformation, "Cingular Real2"
  
  '//tret per ordre de lencarregat i en miralles
  'If Not vnopreguntar Then mirarsihihaextensionsfetes cadbl(rsttmp!numtreball), cadbl(rsttmp!numordremodificacio)
  
  Set rsttmp = Nothing
  If imppantones.Recordset.EOF Then
     crear_pantones
     imppantones.RecordSource = "select * from impresorespantones where comanda=" + atrim(cadbl(comanda))
  End If
  carregar_client_ntintersialtres
  reixa.ReBind
  calcular_totals
  'busco el nom del fitxer pdf aprofinta la funció de obrirpdf passant parametre de no obrir
  vnomfitxerpdf = "noobrir"
  'If Not vestemfentfingerprint Then veureelpdf vnomfitxerpdf
  If impresores.Recordset.EOF And impresores.Recordset.BOF And Command1.Enabled Then
      If vestemfentfingerprint Then
           Command3_Click
          Else
           verificar_tubbase
           veureelpdf vnomfitxerpdf
           If vnopreguntar Then
              Command1_Click
               Else
                 If MsgBox("Vols començar ara la comanda?", vbInformation + vbYesNo + vbDefaultButton1, "Atenció") = vbYes Then
                   Command1_Click
                   
                 End If
           End If
      End If
  End If
 'dues linies per preparar el pdf a gif tretes per esperar a fer bé la orientació
  'preparaelPDF vnomfitxerpdf, 180
 ' If existeix("c:\temp\pdfimpresio.gif") Then Set ImagePDF.Picture = LoadPicture("c:\temp\pdfimpresio.gif")
  If lacomandatereprint(cadbl(comanda)) Then
      If hihabobines(cadbl(comanda)) Then
       If MsgBox("Aquesta comanda porta REPRINT, ara l'estàs fent?", vbInformation + vbYesNo, "Reprint") = vbYes Then
           estemfentreprint = True
           carregar_vernis
           'botollaunesreprint.visible = True
       End If
         Else: MsgBox "Atenció que aquesta comanda es de REPRINT." + vbNewLine + " RECORDA FER MES METRES DELS DEMANATS. +-1000 METRES MES", vbExclamation, "ATENCIÓ REPRINT"
      End If
  End If
  posar_color_material_reciclar cadbl(comanda)
  framebobines.Enabled = False: framepantones.visible = False
  carregarllistadellaunesreprint
  possarestadisticadeldia
'  If impresores.Recordset.EOF Then MsgBox "Baixa nova es començarà amb edició de Clixes.": Command4.Tag = "nou": crearseccio "C": Command4.Tag = ""
  Set rst = Nothing
  Set rsttmp = Nothing
'carrego la pantalla tactil si correspon
  If Checkescanerendollat.Value = 1 Then 'And Not existeix("c:\ordprog.ini") Then
     If Not isloaded("formrevisarCQ") Then Load formrevisarCQ
     formrevisarCQ.carregar_dades cadbl(comanda), 0
     If formrevisarCQ.visible = False Then
         formrevisarCQ.Show
         formrevisarCQ.Frame3.Enabled = False
         formrevisarCQ.Frame4.Enabled = False
         formrevisarCQ.Command3.Enabled = False
         obrirprogramalecturaCB
     End If
     ratoli "normal"
  End If
  escriure_ini "Baixes", "ultimacomanda", comanda, "comandes.ini"
  
End Sub
Sub obrirprogramalecturaCB()
  Dim vX As Integer
  Dim vY As Integer
  Dim vnomfitxer As String
  vX = formrevisarCQ.Left + formrevisarCQ.width
  vY = formrevisarCQ.Top
  vnomfitxer = "C:\Program Files\Axicon\Verifier\verifier.exe"
  If Not existeix(vnomfitxer) Then vnomfitxer = "C:\Program Files (x86)\Axicon\Verifier\verifier.exe"
  If Not existeix(vnomfitxer) Then Exit Sub
  r = Shell(vnomfitxer, vbNormalNoFocus)
  buscar_finestraIcolocarla "Axicon Linear Verifier", -700, 100, -1, 150
  buscar_finestraIcolocarla "Resumen", -700, 100, -1, 400
  
End Sub
Sub posar_color_material_reciclar(vnumc As Double)
  Dim vvalormesgran As Byte
    vcolorverd = &HFF00&
    vcolorblau = &HF3B378
    vcolorvermell = &HFF&
  vvalormesgran = NUMEROCOLORmaterialdelacomanda(vnumc)
  reciclarmaterial1.BackColor = IIf(vvalormesgran = 1, vcolorverd, IIf(vvalormesgran = 2, vcolorblau, vcolorvermell))
End Sub




Sub preparaelPDF(vnomfitxerpdf As String, vrotacio As Double, vMirall As String)
  If Not existeix(vnomfitxerpdf) Then Exit Sub
  vMirall = UCase(vMirall) 'vMirall si es V es vertical H es horitzotal
  If existeix("c:\temp\pdfimpresio.gif") Then Kill "c:\temp\pdfimpresio.gif"
  ConvertirFormats vnomfitxerpdf, "c:\temp\pdfimpresio.gif", 50
  If Not existeix("c:\temp\pdfimpresio.gif") Then GoTo fi
  If vMirall = "H" Then InvertirHVImatge "c:\temp\pdfimpresio.gif", "c:\temp\pdfimpresio.gif"
  If vMirall = "V" Then InvertirHVImatge "c:\temp\pdfimpresio.gif", "c:\temp\pdfimpresio.gif", True
  If vrotacio > 0 Then RotarImatge "c:\temp\pdfimpresio.gif", "c:\temp\pdfimpresio.gif", vrotacio
fi:
End Sub
Sub carregar_vernis()
   Dim rst As Recordset
   Dim rstalt As Recordset
   Dim vidtinter As Long
   Dim vcoditinta As String
   Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb", , True)
   Set rst = dbclixes.OpenRecordset("select * from tintes where not isnull(coditinta) and  id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio * -1))
   While Not rst.EOF
     vcoditinta = atrim(rst!coditinta)
     vidtinter = rst!id_tinter
     If cadbl(rst!tinterlinkambid_treball) > 0 Then
      Set rst = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(cadbl(rst!tinterlinkambid_treball)))
      If rst.EOF Then GoTo fi
      vcoditinta = atrim(rst!coditinta)
     End If
     If vcoditinta <> "" Then Set rstalt = dbclixes.OpenRecordset("select id_tinter,coditinta,color from tintes_alternatives where id_tinter=" + atrim(vidtinter))
     While Not rstalt.EOF
        etvernis.tag = etvernis.tag + IIf(etvernis.tag <> "", ",'", "'") + atrim(rstalt!coditinta) + "'"
        rstalt.MoveNext
     Wend
     If vcoditinta <> "" Then
      Set rst = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(vcoditinta) + "'")
      If Not rst.EOF Then
       etvernis.caption = atrim(rst!descripcio)
       etvernis.tag = etvernis.tag + IIf(etvernis.tag <> "", ",'", "'") + vcoditinta + "'"
      End If
      GoTo fi
     End If
     rst.MoveNext
   Wend
fi:
   Set rst = Nothing
   'Set dbtintes = Nothing
End Sub
Sub verificar_tubbase()
   Dim v As Double
   If vtubbase = 7.6 Then
         v = 0
         While v <> 760
           v = cadbl(InputBox("Aquesta comanda va amb canutu de 760mm" + Chr(10) + "Escriu 760 per continuar", "Canuto de 7,6cm"))
         Wend
   End If
End Sub
Sub mirarsihihaextensionsfetes(vtreball As Double, vordre As Double)
   Dim vrst As Recordset
   Dim vwhere As String
   Dim vmsg As String
   Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb", , True)
   vwhere = "Where numtreball=" + atrim(vtreball) + " and numordremodificacio=" + atrim(vordre)
'   Clipboard.Clear
'   Clipboard.SetText "SELECT extensions_treballsrelacionats.codiextensio, extensions_treballsrelacionats.numtreball, extensions_treballsrelacionats.numordremodificacio, tintes.codi, tintes.descripcio FROM extensions_treballsrelacionats LEFT JOIN tintes ON extensions_treballsrelacionats.coditinta = cdbl(tintes.codi) " + vwhere
   Set vrst = dbtintes.OpenRecordset("SELECT extensions_treballsrelacionats.codiextensio, extensions_treballsrelacionats.numtreball, extensions_treballsrelacionats.numordremodificacio, tintes.codi, tintes.descripcio FROM extensions_treballsrelacionats LEFT JOIN tintes ON trim(str(extensions_treballsrelacionats.coditinta)) = trim(tintes.codi) " + vwhere)
   'comprovar que les tintes resultants realment exixtreixen en el treball i si no existeixes eliminarles de extensions
   
   While Not vrst.EOF
     vmsg = vmsg + Chr(10) + atrim(vrst!codiextensio) + "-> " + atrim(vrst!codi) + "-" + atrim(vrst!descripcio)
     vrst.MoveNext
   Wend
   If vmsg <> "" Then MsgBox vmsg, vbInformation, "Extensions"
   Set vrst = Nothing
   'Set dbtintes = Nothing
End Sub

Function hihabobines(numc As Double) As Boolean
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("SELECT impressores.comanda, bobinesimp.numerodebobina FROM impressores INNER JOIN bobinesimp ON impressores.Id = bobinesimp.controlid WHERE (((impressores.comanda)=" + atrim(numc) + "))")
  If Not rst.EOF Then hihabobines = True
  Set rst = Nothing
End Function

Function mirarsihihaCingularReal(vnumtreball As Double, vordremodificacio As Double) As Boolean
   Dim vurl As String
   Dim generarfitxer_pdf As String
   generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(vnumtreball, "00000") + "\pdf" + Format(vnumtreball, "00000") + "-" + Format(vordremodificacio, "000") + "_CR.pdf"
   If existeix(generarfitxer_pdf) Then
      mirarsihihaCingularReal = True
   End If
   
   
End Function
Sub carregar_materialexacte()
    Dim rst As Recordset
    Set rst = dbtmp.OpenRecordset("select descripcio from materials where codi=" + atrim(cadbl(etmaterialexacte.tag)))
    If Not rst.EOF Then etmaterialexacte = atrim(rst!descripcio)
    Set rst = Nothing
End Sub
Sub possartoleranciaample(ntreball As Double, nmodificacio As Double, materialex As Double, vreduccio As String)
    Dim tanper100 As Byte
    Dim rst As Recordset
    Dim rstm As Recordset
    Dim msgobservacions As String
    tanper100 = 2
    If vreduccio <> "" Then ettoleranciaample.caption = "Tolerancia desarroll: Reducció cilindre": Exit Sub
    espe = False
    Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
    Set rst = dbclixes.OpenRecordset("select desarroll from modificacions where id_treball=" + atrim(ntreball) + " and ordre=" + atrim(nmodificacio))
    If rst.EOF Then Exit Sub
    Set rstm = dbtmp.OpenRecordset("SELECT comandes.comanda, familiesmaterials.descripcio FROM comandes INNER JOIN (familiesmaterials INNER JOIN materials ON familiesmaterials.codi = materials.familia) ON comandes.materialex = materials.codi WHERE (((comandes.comanda)=" + comanda + " ));")
    If Not rstm.EOF Then If Mid(atrim(rstm!descripcio), 1, 2) = "PE" And Mid(atrim(rstm!descripcio), 1, 3) <> "PET" Then tanper100 = 4
    ettoleranciaample.caption = "Tolerancia desarroll: " + atrim(cadbl(rst!desarroll) - tanper100) + " a " + atrim(cadbl(rst!desarroll) + tanper100) + " mm"
    Set rst = dbclixes.OpenRecordset("select * from tintes_observacions where id_treball=" + atrim(ntreball) + " and ordre=" + atrim(nmodificacio) + " order by id")
    While Not rst.EOF
       msgobservacions = msgobservacions + atrim(rst!observacio) + Chr(13)
       rst.MoveNext
    Wend
    If msgobservacions <> "" And Not vnopreguntar Then MsgBox msgobservacions, vbInformation, "Atenció"
    Set rst = Nothing
    Set rstm = Nothing
End Sub
Sub carregar_client_ntintersialtres()
  Dim rstnt As Recordset
  Dim codicli As Double
  Dim numcomanda As Double
  client.caption = "---"
  If cadbl(impresores.Recordset!comanda) = 0 Then
    numcomanda = cadbl(comanda.text)
    If numcomanda = 0 Then client.caption = "0": Exit Sub
   Else: numcomanda = impresores.Recordset!comanda
  End If
  Set rstnt = dbtmp.OpenRecordset("select client,proximaseccio,cilindres,numerotintes from comandes where comanda=" + atrim(cadbl(numcomanda)))
  DoEvents
  DoEvents
  DoEvents
  If Not rstnt.EOF Then
       ntintes = cadbl(rstnt!numerotintes)
       ncilindre = cadbl(rstnt!cilindres)
       framepantones.tag = atrim(rstnt!proximaseccio)
       codicli = cadbl(rstnt!client)
       Set rstnt = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(codicli))
       If Not rstnt.EOF Then client.caption = rstnt!nom
         Else: client.caption = "--- No trobat ---"
  End If
End Sub

Sub crear_pantones()
  r = " comanda "
  For i = 1 To 8
    r = r + ",tinta" + atrim(i) + "a "
  Next i
  Set rsttmp = dbtmp.OpenRecordset("select " + r + " from comandes where comanda=" + atrim(comanda))
  If Not rsttmp.EOF Then
   imppantones.Recordset.AddNew
   imppantones.Recordset!comanda = comanda
   For i = 1 To 8
      imppantones.Recordset.Fields("pantone" + atrim(i)) = rsttmp.Fields("tinta" + atrim(i) + "a")
   Next i
   imppantones.Recordset!comanda = comanda
   imppantones.Recordset!pantone9 = "ETOXI."
   imppantones.Recordset!comanda = comanda
   imppantones.Recordset!pantone10 = "R25."
   imppantones.Recordset.Update
  End If
  imppantones.Refresh
  imppantones.UpdateControls
End Sub
Function comprovarsilesbobinessoncorrelatives() As Boolean
  Dim rst As Recordset
  Dim nocorrelatiu As Boolean
  Dim cont As Long
  Dim noimpres As String

  Set rst = dbtmpb.OpenRecordset("SELECT impressores.comanda, bobinesimp.* FROM bobinesimp INNER JOIN impressores ON bobinesimp.controlid = impressores.Id WHERE (((impressores.comanda)=" + comanda + ")) order by numerodebobina;")

  'Set rst = dbtmpb.OpenRecordset("select * from bobinesimp where controlid in " + r + " order by numerodebobina")
  cont = 1
  If Not rst.EOF Then
     If cadbl(rst!numerodebobina) = 0 Then
        MsgBox "Hi ha una bobina creada amb numero zero... s'hauria de canviar.", vbCritical + vbOKOnly, "Atenció"
        nocorrelatiu = True
        GoTo fi
     End If
  End If
  While Not rst.EOF And Not nocorrelatiu
    If rst!numerodebobina <> cont Then nocorrelatiu = True
    If Not rst!fullbobinaimpres Then noimpres = noimpres + " " + atrim(rst!numerodebobina)
    cont = cont + 1
    rst.MoveNext
  Wend
  If nocorrelatiu Then MsgBox "Les bobines no son correlatives, hauries de canviar-ho.", vbCritical + vbOKOnly, "Atenció"
  
fi:
  Set rst = Nothing
  comprovarsilesbobinessoncorrelatives = nocorrelatiu
  If Not nocorrelatiu And atrim(noimpres) <> "" Then comprovarsilesbobinessoncorrelatives = True: MsgBox "Les bobina/es " + atrim(noimpres) + " no s'ha imprès el full de bobina," + Chr(10) + "aixó podria portar algun error, corretgeix-ho primer.", vbCritical, "Atenció"
End Function

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 2 Then obrir_directament
End Sub
Sub obrir_directament()
  r = InputBox("Escriu la comanda que vols modificar.", "Atenció")
  If cadbl(r) = 0 Then Exit Sub
  comanda = r
  vnopreguntar = True
  Command4_Click
  ensenyar_formanilox_i_tancar True
  vnopreguntar = False
End Sub
Private Sub Command5_Click()
'  If Not clixes.Enabled Then Exit Sub
Dim palet As Double
Dim bobina As Double
Dim paletant As Double
Dim bobinaant As Double
Dim utilitzada As Boolean
If Not comprovamaq Then Exit Sub
paletant = 0: bobinaant = 0
i = 0
While barraestat.caption = "Calculant els totals..."
  DoEvents
Wend
If comprovarsilesbobinessoncorrelatives Then Exit Sub
If bobinesent.Recordset.EOF And Not (bobines.Recordset.EOF And bobines.Recordset.BOF) Then If MsgBox("No hi ha bobines d'entrada vols crear una nova sense possar-les", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
  dblots.visible = False
  framepantones.visible = False
  frameempalmes.visible = False
  framebobentrada.visible = False

  bobines.UpdateRecord
 If impresores.Recordset!tipus = "F" Then
     If cadbl(reixabobines.Columns(4).text) = 0 And Not bobines.Recordset.EOF Then reixabobines.col = 4: reixabobines.SetFocus: MsgBox "Falten els metres a la bobina": Exit Sub
     If bobentrada.EditActive Then bobinesent.UpdateRecord
     If Not bobinesent.Recordset.EOF Then bobinesent.Recordset.MoveFirst
     While Not bobinesent.Recordset.EOF And Not bobines.Recordset.EOF
       If cadbl(bobinesent.Recordset!palet) > 0 And cadbl(bobinesent.Recordset!bobina) > 0 Then
         carregar_bobinesdentrada "mirarsiutilitzada", , bobinesent.Recordset!palet, bobinesent.Recordset!bobina, cadbl(comanda), utilitzada
          If Not utilitzada Then
             'carregar_bobinesdentrada "marcarutilitzadademanar", , bobinesent.Recordset!palet, bobinesent.Recordset!bobina, cadbl(comanda)
             If MsgBox("Canvi de bobina d'entrada?", vbExclamation + vbYesNo, "Canvi bobina?") = vbYes Then
                palet = cadbl(bobinesent.Recordset!palet)
                bobina = cadbl(bobinesent.Recordset!bobina)
                estatdelabobina palet, bobina, 0, ncomanda
               'wait 1
               'bobinesdentrada.imprimir_bobinaparcial palet, bobina, , 1
               Unload bobinesdentrada
                 Else: paletant = cadbl(bobinesent.Recordset!palet): bobinaant = cadbl(bobinesent.Recordset!bobina)
             End If
          End If
       End If
       bobinesent.Recordset.MoveNext
     Wend
     nova_bobina
     DoEvents
     If paletant > 0 Then afegir_labobinadentrada paletant, bobinaant
     Command13_Click
     bobentrada.SetFocus
     Else: MsgBox "Has d'escullir una linia de FUNCIONAMENT."
  End If
  ratoli "normal"
  ' calcular_totals
End Sub
Sub nova_bobina()
  Dim rstmp As Recordset
  Dim rsttmp2 As Recordset
  Dim col As Byte
  Dim elgran As Double
  reixabobines.tag = "afegint"
  If Not bobines.Recordset.EOF Then
   If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
   bobines.Recordset.Update
  End If
  Set rsttmp2 = dbtmpb.OpenRecordset("select id  from impressores where comanda=" + atrim(impresores.Recordset!comanda))
  elgran = 0
  While Not rsttmp2.EOF
   Set rstmp = dbtmpb.OpenRecordset("select max(numerodebobina) as elgran from bobinesimp where controlid=" + atrim(rsttmp2!id))
   If Not rstmp.EOF Then
      If cadbl(rstmp!elgran) > elgran Then elgran = cadbl(rstmp!elgran)
   End If
   rsttmp2.MoveNext
  Wend
  Set rstmp = dbtmpb.OpenRecordset("select * from bobinesimp where controlid=" + atrim(impresores.Recordset!id) + " and numerodebobina=" + atrim(elgran))
  bobines.Recordset.AddNew
  bobines.Recordset!numerodebobina = elgran + 1
  bobines.Recordset!controlid = atrim(impresores.Recordset!id)
  bobines.Recordset!numempalmes = 0
  bobines.Recordset!datafab = Date
  col = 0
  bobines.Recordset!operari1 = numop
 ' If Not rstmp.EOF Then
 '    bobines.Recordset!metres = rstmp!metres
 '    bobines.Recordset!kilos = rstmp!kilos
 '    col = 3 'escullo a la columne que es posa per defecte
 '  Else: col = 3
 ' End If
  bobines.Recordset.Update
  'bobines.Refresh
  reixabobines.Refresh
  bobines.Recordset.MoveLast
  'reixabobines.Refresh
  DoEvents
  reixabobines.col = col
  'reixabobines.SetFocus
  Set rstmp = Nothing
  Set rstmp2 = Nothing
If reixabobines.text = "0" Then reixabobines.SelLength = Len(reixabobines.text)
reixabobines.tag = ""
End Sub

Private Sub DBGrid1_DblClick()
r = "numeric"
Set campcontrol = ActiveControl
teclattactil.Show
End Sub

Private Sub Command6_Click()
 
  dblots.visible = False
  framepantones.visible = False
  frameempalmes.visible = False
  framebobentrada.visible = False
 
 If MsgBox("Segur que vols borrar aquesta bobina?", vbCritical + 4, "Atenció") = vbYes Then
     If Not bobines.Recordset.EOF Then
       dbtmpb.Execute "delete * from impempalmes where id=" + atrim(cadbl(bobines.Recordset!id))
       dbtmpb.Execute "delete * from bobentradaimp where idbobimp=" + atrim(cadbl(bobines.Recordset!id))
       bobines.Recordset.Delete
     End If
     On Error Resume Next
     bobines.Refresh
     reixabobines.Refresh
     bobines.Recordset.MoveLast
  End If
  calcular_totals
End Sub

Private Sub Command7_Click()
Dim numb As Integer
Dim mtrs As Double
Static cont As Byte
 If bobines.Recordset.EOF Then MsgBox "No hi ha cap bobina per imprimir": Exit Sub
 If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
 If cadbl(bobines.Recordset!metres) = 0 Then
   mtrs = 0
   While mtrs = 0
     mtrs = cadbl(InputBox("Entra els Metres de la bobina. ", "Atenció"))
   Wend
   bobines.Recordset!metres = cadbl(mtrs)
 End If
 If bobines.Recordset!fullbobinaimpres Then
    If UCase(InputBox("Aquest full de bobina ja s'ha imprès una vegada." + Chr(10) + "  Assegura't que no estiguis repetint bobina impresa." + Chr(10) + "ESCRIU [SEGUR] PER REIMPRIMIR-LA.", "Reimpresió de paper bobina")) <> "SEGUR" Then Exit Sub
    'desactivem l'enviament alicia es queixa massa reimpresions mantenimentbobina.passaravis 0, 0, "Full de bobina imprès mes d'una vegada", comanda, " El full de la bobina Nº: " + atrim(bobines.Recordset!numerodebobina) + " de la Comanda: " + atrim(comanda) + " s'ha imprès mes d'una vegada s´hauria de revisar que hi hagui algun error."
 End If
 If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
 bobines.Recordset.Update
  If cont = 3 Then cont = 0: form1.caption = "Baixes Comandes (Impresores)"
  If form1.caption = "Imprimint la bobina..." Then cont = cont + 1: Exit Sub
  form1.caption = "Imprimint la bobina..."
  bobines.UpdateRecord
  If Not bobines.Recordset.EOF Then numb = bobines.Recordset!numerodebobina
  calcular_totals
  wait (2)
  bobines.Recordset.FindFirst "numerodebobina=" + atrim(cadbl(numb))
   

  imprimir_bobina
  If Not bobines.Recordset.EOF Then
    If Not vestemfentfingerprint Then  'si no fa fingerprint fer etiqueta de qualitat
      imprimir_controlqualitat cadbl(comanda), numop, bobines.Recordset!numerodebobina
    End If
    If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
    bobines.Recordset!fullbobinaimpres = True
    bobines.Recordset.Update
    bobines.Recordset.FindFirst "numerodebobina=" + atrim(cadbl(numb))
      Else: MsgBox "Error no s'ha pogut modificar la bobina", vbCritical, "Error"
  End If
  form1.caption = "Baixes Comandes (Impresores)"
  controlstock.caption = ""
  ratoli "normal"
End Sub
Sub possarvalorsreduccioalllistat(numformula As Byte)
   
End Sub
Sub demanarvalorsdelta(vnumc As Double, vnumbob As Double, vcont As Byte)
    Dim rst As Recordset
    Dim v As Double
    Dim vresp As String
    Dim vi As Byte
    Dim vidtreball As Double
    Dim vordremodificacio As Double
    Dim vvalordelta As String
    Dim vdeltamaxim As Double
    Dim vsql As String
    Dim vdeltamaximTINTES As Double
    vi = 1
    vdeltamaxim = 2.5
    vdeltamaximTINTES = 2
    Set rst = dbtmp.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vnumc))
    If Not rst.EOF Then vidtreball = cadbl(rst!numtreball): vordremodificacio = cadbl(rst!numordremodificacio)
    Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
    Set rst = dbclixes.OpenRecordset("select valordeltamaxim from modificacions where id_treball=" + atrim(vidtreball) + " and ordre=" + atrim(vordremodificacio))
    If Not rst.EOF Then vdeltamaxim = cadbl(rst!valordeltamaxim)
    vsql = "select * from tintes where (id_treball=" + atrim(vidtreball) + " and ordremodificacio=" + atrim(vordremodificacio) + " and tinterlinkambid_treball<1) or id_tinter in (select tinterlinkambid_treball  from tintes where id_treball=" + atrim(vidtreball) + " and ordremodificacio=" + atrim(vordremodificacio) + " and tinterlinkambid_treball>0)"
    'Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(vidtreball) + " and ordremodificacio=" + atrim(vordremodificacio))
    Set rst = dbclixes.OpenRecordset(vsql)
    While Not rst.EOF
      If InStr(1, atrim(rst!color), "P-") > 0 And InStr(1, atrim(rst!color), "PRIMAR") = 0 Then
demanardelta:
          v = demanarvalordelta(atrim(rst!color), vdeltamaxim)
          If v = 9 Then GoTo sensedelta
          If v > vdeltamaxim Then
            vmsg = "Delta màxim superat  Màx:" + atrim(vdeltamaxim) + "  Llegit:" + atrim(v)
            If vmsg = "" Then GoTo demanardelta
          End If
          If v <= 0 Then vmsg = "Delta llegit valor zero."
          If vmsg <> "" Then vresp = InputBox("Escriu el motiu de: " + Chr(10) + vmsg + "SI T'HAS EQUIVOCAT POSSA [ERROR] SISPLAU")
          If vmsg <> "" Then enviaremailgeneric "missatgesgenericsimpresores", "ERROR Lectura Delta [" + nommaq + "] - " + nomoperari + "  Comanda:" + atrim(vnumc), "Color: " + atrim(rst!color) + Chr(10) + treure_apostruf("Error: " + vmsg + Chr(10) + "Resposta operari: " + vresp)
          If v > vdeltamaximTINTES Then enviaremailgeneric "tintes@inplacsa.com", "ERROR Lectura Delta>2 [" + nommaq + "] - " + nomoperari + "  Comanda:" + atrim(vnumc), "Color: " + atrim(rst!color) + Chr(10) + treure_apostruf("Error: " + vmsg + Chr(10) + "Resposta operari: " + vresp)
          guardarvalordelta vnumc, vnumbob, v, rst
sensedelta:
          '"missatgesgenericsimpresores"
          vmsg = ""
          vresp = ""
          vvalordelta = atrim(Redondejar(v, 2))
          If cadbl(vvalordelta) = 9 Then vvalordelta = "N/S"
          llistat.Formulas(vcont) = "delta" + atrim(vi) + "='VD: " + vvalordelta + " - " + atrim(rst!color) + "'"
          vcont = vcont + 1
          vi = vi + 1
      End If
      rst.MoveNext
    Wend
   Set dbclixes = Nothing
End Sub
Sub guardarvalordelta(vnumc As Double, vnumbob As Double, vvalor As Double, rst As Recordset)
  Dim vvalues As String
  vvalues = "(" + atrim(vnumbob) + "," + atrim(vnumc) + ",now," + atrim(cadbl(tmetres)) + "," + atrim(rst!coditinta) + ",'" + treure_apostruf(rst!color) + "'," + atrim(numop) + "," + passaradecimalpunt(atrim(vvalor)) + ")"
  dbtmpb.Execute "insert into impresores_valorsdelta (numbobina,comanda,hora,metres,coditinta,nomdelatinta,operari,valordelta) values " + vvalues
End Sub
Function demanarvalordelta(vdesc As String, vdeltamax As Double, Optional vX As Double, Optional vY As Double) As Double
  Dim v As String
  
  Unload formllegirdeltaibarres
  Load formllegirdeltaibarres
  formllegirdeltaibarres.framedelta.visible = True
  formllegirdeltaibarres.etdeltamaxicolor = "Màx: " + atrim(vdeltamax) + "  " + vdesc
  If vX = 0 Then vX = Screen.width / 2: vY = Screen.Height / 2
  formllegirdeltaibarres.Left = vX - formllegirdeltaibarres.width / 2: formllegirdeltaibarres.Top = vY - formllegirdeltaibarres.Height / 2

  formllegirdeltaibarres.Show 1
'  v = atrim(InputBox(UCase(vdesc) + Chr(10) + "Escriu el valor delta d'aquest tinter." + Chr(10) + "SI NO POTS LLEGIR EL DELTA ESCRIU 9", "Valor delta"))
 ' demanarvalordelta = cadbl(substituir(v, ".", ","))
 ' If cadbl(demanarvalordelta) = 0 Then GoTo demanardelta
 ' If cadbl(demanarvalordelta) > 9 Then GoTo demanardelta
 If formllegirdeltaibarres.tag = "No llegeix" Or formllegirdeltaibarres.tag = "" Then formllegirdeltaibarres.tag = "9"
 demanarvalordelta = cadbl(formllegirdeltaibarres.tag)
 Unload formllegirdeltaibarres
End Function
Function demanarvalorescaner(vnumc As Double, Optional vX As Long, Optional vY As Long) As Double
   Dim v As Double
   Dim vmsg As String
   Dim vresp As String
   Dim rst As Recordset
   Dim codibarras As String
   Set rst = dbtmp.OpenRecordset("SELECT comandes.comanda, Clixes.codidebarres FROM comandes LEFT JOIN Clixes ON comandes.numtreball = Clixes.id_treball where comanda=" + atrim(vnumc))
   If Not rst.EOF Then codibarras = atrim(rst!codidebarres)

   'v = 0
   'vmsg = "Valor de lectura d'escaner no òptim."
   'While v <= 1
   '  v = cadbl(InputBox("Entra el valor de l'Escaner:" + Chr(10) + "SI NO HI HA CODI DE BARRES ESCRIU 9", "Valor de l'escaner"))
   '  If v <= 1 Then
   '    If MsgBox("El valor de l'escaner hauria de ser mes gran de 1." + Chr(10) + "Estas correcte que vols possar " + atrim(v) + "?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
   '       vresp = InputBox("Escriu el motiu d'aquest valor. ", "Valor no òptim")
   '       GoTo cont
   '    End If
   '  End If
   'Wend
  If vvalidaciocodidebarres = "" Or atrim(vvalidaciocodidebarres) = "Er:" Then
       If vX = 0 Then
               vcodidebarres = InputBox("Escaneja el codi de barres de la mostra.", "Comprovar codi de barres")
                 Else: vcodidebarres = InputBox("Escaneja el codi de barres de la mostra.", "Comprovar codi de barres", , vX - 1500, vY - 1500)
       End If
       If atrim(vcodidebarres) <> atrim(codibarras) Then
          If Len(codibarras) = 15 Then
               If atrim(vcodidebarres) = Mid(codibarras, 1, 13) Then
                    MsgBox "El codi de barres es de 13+2 digits comprova els dos ultims que siguin correctes." + Chr(10) + "CODI: " + Mid(codibarras, 1, 13) + " +  " + Mid(codibarras, 14, 2)
                  Else: vcodidebarres = "Er:" + vcodidebarres: MsgBox "Error, el codi escanejat no coincideix amb el del producte." + Chr(10) + "REVISA QUE SIGUI TOT CORRECTE." + Chr(10) + "EL CODI CORRECTE SERIA: " + codibarras, vbCritical, "Error"
               End If
             Else:
                If Len(codibarras) = 13 Then
                      If atrim(vcodidebarres) = Mid(codibarras, 2, 12) Then
                        MsgBox "El codi de barres es de 1+12 digits comprova el primer digit que siguin correcta." + Chr(10) + "CODI: " + Mid(codibarras, 1, 1) + " +  " + Mid(codibarras, 2, 12)
                      Else: vcodidebarres = "Er:" + vcodidebarres: MsgBox "Error, el codi escanejat no coincideix amb el del producte." + Chr(10) + "REVISA QUE SIGUI TOT CORRECTE." + Chr(10) + "EL CODI CORRECTE SERIA: " + codibarras, vbCritical, "Error"
                  End If
                   Else: vcodidebarres = "Er:" + vcodidebarres: MsgBox "Error, el codi escanejat no coincideix amb el del producte." + Chr(10) + "REVISA QUE SIGUI TOT CORRECTE." + Chr(10) + "EL CODI CORRECTE SERIA: " + codibarras, vbCritical, "Error"
                End If
          End If
       End If
       vvalidaciocodidebarres = vcodidebarres
       dbtmpb.Execute "update impressorestot set validaciocodidebarres='" + treure_apostruf(vvalidaciocodidebarres) + "' where comanda=" + atrim(vnumc)
  End If
  If vhihadigimarc Then
      While UCase(InputBox("Aquesta comanda es DIGIMARC verifica que llegeix correctament." + Chr(10) + "Escriu [CORRECTE] quan ho hagis comprovat.", "Comprovar DIGIMARC")) <> "CORRECTE"
        DoEvents
      Wend
      vdigimarc = True
      dbtmpb.Execute "update impressorestot set validaciodigimarc=True where comanda=" + atrim(vnumc)
  End If
  
  Unload formllegirdeltaibarres
  Load formllegirdeltaibarres
  formllegirdeltaibarres.frameCB.visible = True
  If vX = 0 Then vX = Screen.width / 2: vY = Screen.Height / 2
  formllegirdeltaibarres.Left = vX - formllegirdeltaibarres.width / 2: formllegirdeltaibarres.Top = vY - formllegirdeltaibarres.Height / 2
  formllegirdeltaibarres.Show 1
  If formllegirdeltaibarres.tag = "No Llegeix" Then formllegirdeltaibarres.tag = "10"
  If formllegirdeltaibarres.tag = "No en té" Then formllegirdeltaibarres.tag = "9"
  v = cadbl(formllegirdeltaibarres.tag)
  If v = 10 Then
      If vX = 0 Then
           vresp = InputBox("Escriu el motiu d'aquest valor. ", "Valor no òptim")
            Else: vresp = InputBox("Escriu el motiu d'aquest valor. ", "Valor no òptim", , vX - 1500, vY - 1500)
      End If
  End If
  Unload formllegirdeltaibarres
cont:
  If v = 10 Then enviaremailgeneric "missatgesgenericsimpresores", "ERROR Lectura Escaner [" + nommaq + "] - " + nomoperari + "  Comanda:" + atrim(vnumc), treure_apostruf("Error: " + vmsg + Chr(10) + "Resposta operari: " + vresp)
  demanarvalorescaner = v
  Set rst = Nothing
End Function
Sub controldequalitatalapantallasecundaria(vnumc As Double, vnumbob As Double)
  'obrir el document word i colocarlo a la segona pantalla haig de buscar els valors de inici de la segona pantalla
  'If Not buscar_finestraIcolocarla("safata d'entrada", 1, 1, 400, 400) Then MsgBox "No he trobat la finestra"
  Dim vformjacarregat As Boolean
  vformjacarregat = isloaded("formrevisarCQ")
  framesegonapantalla.visible = True
  framesegonapantalla.ZOrder 0
  If Not vformjacarregat Then Load formrevisarCQ
  formrevisarCQ.carregar_dades vnumc, vnumbob
  'formrevisarCQ.colocarFormalasegonapantalla
  ratoli "normal"
  If Not vformjacarregat Then formrevisarCQ.Show
  While formrevisarCQ.Command3.Enabled
     DoEvents
  Wend
  ratoli "normal"
  SetCursorPos 200, 200
  framesegonapantalla.visible = False
End Sub
Sub imprimir_controlqualitat(numc As Double, op As Byte, numbob As Double)
   Dim ultimalinia As String
   Dim vvalorescaner As Double
   Dim vindexformules As Byte
   If Checkescanerendollat.Value = 1 Then vescanerendollat = True Else vescanerendollat = False
   If vescanerendollat Then controldequalitatalapantallasecundaria numc, numbob: GoTo fi
   ultimalinia = "Op: " + atrim(op) + "    NºBob.Salida: " + atrim(numbob) + "   Fecha: " + Format(Now, "dd/mm/yy")
   For i = 0 To 100
     llistat.Formulas(i) = ""
   Next i
   llistat.Formulas(0) = "lot=" + atrim(numc)
   llistat.Formulas(1) = "ultimalinia='" + atrim(ultimalinia) + "'"
   llistat.Formulas(2) = "nommaquina='" + atrim(nummaq) + "-" + atrim(nommaq) + "'"
   llistat.Formulas(3) = "nummaq='" + atrim(nummaq) + "'"
   
   vvalorescaner = demanarvalorescaner(numc)
   If vvalorescaner = 0 Then Exit Sub
   llistat.Formulas(4) = "valorescaner='" + atrim(vvalorescaner) + "'"
   If vvalidaciocodidebarres <> "-" Then
       llistat.Formulas(5) = "codidebarres='CB: " + atrim(vvalidaciocodidebarres) + "'"
        Else: llistat.Formulas(5) = ""
   End If
   If vdigimarc Then
        llistat.Formulas(6) = "digimarc='Digimarc OK'"
          Else: llistat.Formulas(6) = "digimarc=''"
   End If
   
   If (vavispeu <> "") And numbob = 0 Then
      If UCase(InputBox("Peu/Data: [" + vavispeu + "] VERIFICA'L." + Chr(10) + "ESCRIU [OK] PER CONTINUAR.", "PEU IMPRENTA")) <> "OK" Then Exit Sub
   End If
   llistat.Formulas(7) = "peuimprenta='" + atrim(vavispreu) + "'"
'   llistat.Formulas(7) = "peuimprenta='prova 123'"
   vindexformules = 8
   demanarvalorsdelta numc, numbob, vindexformules
   
   calcularvalorsreducciocilindre numc, nummaq, vindexformules, llistat
   llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "verificacioqualitatimpresores.rpt"
   llistat.Destination = crptToPrinter
    llistat.CopiesToPrinter = 1
   llistat.DataFiles(0) = ""
   llistat.DiscardSavedData = True
' llistat.PrinterName = llegir_ini("Impressores", "nomfulla", "baixesimpressora.ini")
' llistat.PrinterPort = llegir_ini("Impressores", "portfulla", "baixesimpressora.ini")
' llistat.PrinterDriver = llegir_ini("Impressores", "driverfulla", "baixesimpressora.ini")
   escullir_impresora_tickets
   DoEvents
   If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
   llistat.Action = 1
   llistat.PrinterDriver = ""
   llistat.PrinterName = ""
   llistat.PrinterPort = ""
   MsgBox "ATENCIÓ CONTROL DE VERIFICACIÓ DE QUALITAT." + Chr(10) + "VERIFICA LA IMPRESIÓ AMB L'ETIQUETA IMPRESA", vbInformation, "VERIFICACIÓ QUALITAT"
fi:
   
End Sub
Sub netejarreport(rpt As CrystalReport)
  Dim i As Byte
  rpt.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "reportblanc.rpt"
  rpt.Destination = crptToFile
  For i = 1 To 20
     rpt.Formulas(i) = ""
     rpt.DataFiles(i) = ""
  Next i
  rpt.Action = -1
End Sub
Function buscarproximaseccio(ncomanda As Double, ncomanda2 As Double) As String
   Dim rst As Recordset
   Dim rst2 As Recordset
   If Not estemfentreprint And lacomandatereprint(ncomanda) Then buscarproximaseccio = "rP": GoTo fi
   Set rst = dbtmp.OpenRecordset("SELECT comandes.comanda, comandes.linkcomanda2,productes.ruta, comandes.proximaseccio,comandes.impressio FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(ncomanda) + "));")
   If rst.EOF Then GoTo fi
   buscarproximaseccio = Mid(rst!ruta, InStr(1, rst!ruta, "I") + 1, 1)
fi:
   Set rst = Nothing
End Function
Private Sub imprimir_bobina()
 Dim vmaterialespecial As Boolean
If Command7.tag <> "imprimint" Then
   Command7.tag = "imprimint"
 Else: MsgBox "Ja està imprimint espera que acavi per tornar-hi": Exit Sub
End If
crear_taula_imp_empalmes

possar_valors_taula_imp_empalmes vmaterialespecial
llistatbob.DiscardSavedData = True
llistatbob.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "etempalmes2.rpt"
'llistat.Destination = crptToWindow
llistatbob.Destination = crptToPrinter
'dbtemp.Close
'Set dbtemp = Nothing
llistatbob.DataFiles(0) = nomfitxertemporalbobent 'nomfitxertemporal
DoEvents
'wait (4)

 If existeix("c:\ordprog.ini") Then llistatbob.Destination = crptToWindow
 llistatbob.Formulas(0) = ""
 llistatbob.Formulas(1) = ""
 llistatbob.Formulas(2) = ""
 llistatbob.Formulas(3) = ""
 
 If nummaq = 5 Then llistatbob.Formulas(0) = "nommaquina='" + "Màquina: FS    " + Format(Now, "hh:nn") + "'"
 If nummaq = 7 Then llistatbob.Formulas(0) = "nommaquina='" + "Màquina: FW 1508    " + Format(Now, "hh:nn") + "'"
 llistatbob.Formulas(1) = "proximaseccio='" + buscarproximaseccio(ncomanda, ncomanda2) + "'"
 llistatbob.Formulas(2) = "nohihamaterialLAM='" + mirarsihihamaterialLAM(ncomanda) + "'"
 llistatbob.CopiesToPrinter = 1
 'llistatbob.WindowControls = False
 
 'On Error GoTo noimprimeix
llistatbob.Action = 1
'Set dbtemp = Nothing
Command7.tag = ""
controlstock = ""
If vmaterialespecial Then
      Load fcalculant
      fcalculant.Command1.BackColor = QBColor(12)
      fcalculant.Command1.caption = "Aquesta bobina de sortida s'ha d'embolicar ràpidament." + vbNewLine + "MATERIAL ESPECIAL" + Chr(13) + "Fes Click"
      fcalculant.Command1.FontSize = 16
      fcalculant.Show 1
      Unload fcalculant
End If
Exit Sub
noimprimeix:
  MsgBox "No puc imprimir l'etiqueta."
  Resume Next
End Sub
Function mirarsihihamaterialLAM(vnumc As Double) As String
  Dim rst As Recordset
  Dim vnumc1 As Double
  Dim vnumc2 As Double
  Dim vmat1 As String
  Dim vmat2 As String
  
  obrestocks
  vmat1 = "N": vmat2 = "N"
  Set rst = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(vnumc))
  If rst.EOF Then Exit Function
  vnumc1 = rst!linkcomanda1
  vnumc2 = rst!linkcomanda2
  Set rst = dbstocks.OpenRecordset("SELECT * FROM percomandaoclient  WHERE (((percomandaoclient.numcomanda)=" + atrim(cadbl(vnumc1)) + "));")
  If Not rst.EOF Then vmat1 = "S"
  Set rst = dbstocks.OpenRecordset("select* from comandes_firmes where anulada=false and tipus='PK2' and comanda=" + atrim(vnumc1))
  If Not rst.EOF Then vmat1 = "S"
  Set rst = Nothing
  Set rst = dbstocks.OpenRecordset("SELECT * FROM percomandaoclient  WHERE (((percomandaoclient.numcomanda)=" + atrim(cadbl(vnumc2)) + "));")
  If Not rst.EOF Then vmat2 = "S"
  Set rst = dbstocks.OpenRecordset("select * from comandes_firmes where anulada=false and tipus='PK2' and comanda=" + atrim(vnumc2))
  If Not rst.EOF Then vmat2 = "S"
  Set rst = Nothing
  If vnumc1 > 0 And vmat1 = "N" Then mirarsihihamaterialLAM = "N"
  If vnumc2 > 0 And vmat2 = "N" Then mirarsihihamaterialLAM = "N"
End Function
Function nommaterialdelpalet(numpalet As Double) As String
  Dim rstmat As Recordset
  Dim rstpalet As Recordset
  Set rstpalet = dbstocks.OpenRecordset("select codimatprognou from palets where idpalet=" + atrim(numpalet))
  If rstpalet.EOF Then Exit Function
  Set rstmat = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstpalet!codimatprognou)))
  If rstmat.EOF Then Exit Function
  nommaterialdelpalet = IIf(rstmat!materialdelicat, "*", "") + descripciomaterial(rstmat)
  Set rstmat = Nothing
  Set rstpalet = Nothing

End Function
Function descripciomaterial(rstmat As Recordset) As String
  Dim desc As String
  Dim rstfam As Recordset
  Set rstfam = dbtmp.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(cadbl(rstmat!familia)))
  If Not rstfam.EOF Then desc = desc + atrim(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from subfamiliesmaterials where codi=" + atrim(cadbl(rstmat!subfamilia)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from familiescolorants where codi=" + atrim(cadbl(rstmat!familiacol)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from subfamiliescolorants where codi=" + atrim(cadbl(rstmat!subfamiliacol)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from familiesaditius where codi=" + atrim(cadbl(rstmat!familiaad)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from subfamiliesaditius where codi=" + atrim(cadbl(rstmat!subfamiliaad)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  descripciomaterial = desc
End Function
Function af(v As Variant) As String
  v = atrim(v)
  If Len(v) > 1 Then
     v = "+" + v
    Else: v = ""
  End If
  af = v
End Function

Sub possar_valors_taula_imp_empalmes(vmaterialespecial As Boolean)
 Dim rs As Recordset
 Dim rs2 As Recordset
 Dim bobe As String
 obrestocks
 If Not bobinesent.Recordset.EOF Then
   bobinesent.Recordset.MoveFirst
 End If
 Set rs = dbtemp.OpenRecordset("tmp_imp_empalmes")
 Set rststocks = dbstocks.OpenRecordset("select * from palets where idpalet=" + atrim(cadbl(bobinesent.Recordset!palet)))
 r = nommaterialdelpalet(cadbl(bobinesent.Recordset!palet))
 If Mid(r + " ", 1, 1) = "*" Then vmaterialespecial = True
 
 If bobines.Recordset.EOF Then Exit Sub
 rs.AddNew
 rs!numlot = comanda.text
 rs!numbobsort = cadbl(bobines.Recordset!numerodebobina)
 rs!numop = cadbl(bobines.Recordset!operari1)
 rs!numop2 = cadbl(bobines.Recordset!operari2)
 rs!datafab = Format(bobines.Recordset!datafab, "dd/mm/yy")
 rs!client = client.caption
 rs!texteimpressio = texteimpresio
 rs!refclient = refclient
 rs!observacio = bobines.Recordset!observacio
 rs!comandaclient = comandaclient
 rs!material = r
 For i = 1 To 4
   If Not bobinesent.Recordset.EOF Then
      rs.Fields("numbobent" + atrim(i)) = atrim(bobinesent.Recordset!palet) + "/" + atrim(bobinesent.Recordset!bobina)
      bobinesent.Recordset.MoveNext
   End If
 Next i
 rs!ample = cadbl(rststocks!ample)
 rs!plegat = cadbl(rststocks!plegat)
 rs!solapa = cadbl(rststocks!solapa)
 rs!metres = cadbl(bobines.Recordset!metres)
 rs!kilos = cadbl(bobines.Recordset!kilos)
 rs!espessor = micrescomanda
 'actualitzo els camps de bobina
    bobines.Recordset.Edit
    bobines.Recordset!ample = rs!ample
    bobines.Recordset!espessor = rs!espessor
    bobines.Recordset.Update
 ' fins aqui
 llistat.Formulas(0) = "mesuraesp='(" + mesuraespcomanda + ")'"
 rs!codibarres = codibarras
 empalmes.RecordSource = "select * from impempalmes where id=" + atrim(bobines.Recordset!id)
 empalmes.Refresh
 If Not empalmes.Recordset.EOF Then
  empalmes.Recordset.MoveFirst
  i = 0
  While Not empalmes.Recordset.EOF And i < 8
    rs.Fields("empalme" + atrim(i + 1)) = empalmes.Recordset!observacions
    rs.Fields("mtrs" + atrim(i + 1)) = empalmes.Recordset!metres
    i = i + 1
    empalmes.Recordset.MoveNext
  Wend
 End If
 rs.Update
 
 Set rs = Nothing
 Set rs2 = Nothing
 Set rststocks = Nothing
 Set dbstocks = Nothing
End Sub
Sub crear_taula_imp_empalmes()
  Dim camps As String
  'nomfitxertemporal = "c:\temp\" + Format(Now, "~biddmmhhnnss") + ".mdb"
  nomfitxertemporal = nomfitxertemporalbobent
  On Error Resume Next
   'MkDir "c:\temp"
   'Kill "c:\temp\~bi*.*"
   'DBEngine.CreateDatabase nomfitxertemporal, dbLangGeneral, dbVersion10
   
   'Set dbtemp = OpenDatabase(nomfitxertemporal)
   dbtemp.Execute "drop table tmp_imp_empalmes"
   'dbtemp.Execute "drop table tmp_imp_empalmes"
  On Error GoTo 0
  camps = "numlot double,numbobsort double, numop double,datafab string,numbobent1 string,numbobent2 string,client string,refclient string,comandaclient string,texteimpressio string,material string,ample double,plegat double,solapa double,espessor double,metres double,kilos double"
  camps = camps + " , empalme1 string,mtrs1 double,empalme2 string, mtrs2 double,empalme3 string, mtrs3 double,empalme4 string,mtrs4 double,empalme5 string,mtrs5 double,numbobent3 string,numbobent4 string, codibarres string, observacio string, numop2 double"
  camps = camps + " , empalme6 string,mtrs6 double,empalme7 string, mtrs7 double,empalme8 string, mtrs8 double"
  'ample double,plegat double,solapa double,espessor double,metres double,kilos double)"
  dbtemp.Execute ("create table tmp_imp_empalmes (" + camps) + ")"
End Sub






Sub imprimir_fulla()
  Dim mtrsparcialanteriors As Double
  Dim metrespackingfinal As Double
  Dim metrespacking As Double
  Dim metrescomanda As Double
   Dim nb As String
   Dim numc As Double
   Dim nmaq As Byte
   Dim np As Double
   Dim kgs As Double
   Dim mtrs As Double
   Dim linia As Double
   Dim formula As Byte
   Dim rsttmpbob As Recordset
   Dim rst As Recordset
   Dim rsttemp As Recordset
   Dim rsttmp2 As Recordset
   Dim rstopcions As Recordset
   Dim mtrsajust As Double
   Dim vmtrsreals As Double
   Dim vmtrsimpresos As Double
   Dim rstdr As Recordset
   Dim texteimp As String
   Dim vcarpetadesti As String
   
   numc = cadbl(comanda.text)
   form1.caption = "Imprimint..."
   nample = 0
   formula = 0
   For i = 0 To 20
    llistat.Formulas(i) = ""
   Next i
  crear_taula_impresio_baixa
  obrestocks
  Set rsttemp = dbtmpb.OpenRecordset("tmp_imp_baixa")
  imppantones.Refresh
  rsttemp.AddNew


  ' busco l'ample
   'ample_palet
  '-----------
  texteimp = ""
  Set rst = dbtmp.OpenRecordset("select marcailinia,texteimpressio,ampleesq,cantitatex from comandes where comanda=" + atrim(cadbl(numc)))
  If Not rst.EOF Then metrescomanda = cadbl(rst!cantitatex): nample = cadbl(rst!ampleesq): texteimp = IIf(atrim(rst!marcailinia) = "", atrim(rst!texteimpressio), atrim(rst!marcailinia)): Set rst = Nothing
  
  With rsttemp
  !comanda = atrim(comanda.text)
  '!client = atrim(client.Caption)
  !client = client.ToolTipText
  !firmat = atrim(firmat.caption)
  !nomfirmat = possarnomfirmat
  !tintersrentats = cadbl(trentats)
  !portaclixers = cadbl(pclixers)
  !canvienfilada = atrim(canvienfilada)
  !numtintes = cadbl(ntintes)
  !cilindre = cadbl(ncilindre)
  !espessor = micrescomanda
  !comandaacavada = IIf(comandaacavada.Value, 1, 0)
  !texteimp = texteimp
  'prep clixe
     'LA PREPaRACIÓ JA NO S'ENTRA A LA BAIXA
  
  Set rst = dbtmpb.OpenRecordset("select id,comanda,numeromaquina,operari,datainici,horainici,datafi,horafi,observacio from impressores where comanda=" + comanda.text + "  order by datainici,horainici")
  If Not rst.EOF Then
   nmaq = rst!numeromaquina
   i = 1
  End If
'afegir el descans relleu al llistat
    Set rstdr = dbtmpb.OpenRecordset("select * from controldescansrelleu where seccio='" + atrim(lletraseccio) + "' and comanda=" + atrim(ncomanda))
    While Not rstdr.EOF And i < 5
        .Fields("prepclixe_data" + Trim(i)) = Format(atrim(rstdr!datainici), "dd/mm/yy")
        .Fields("prepclixe_op" + Trim(i)) = cadbl(rstdr!operari)
        .Fields("prepclixe_de" + Trim(i)) = Format(atrim(rstdr!horainici), "hh:nn")
        .Fields("prepclixe_fins" + Trim(i)) = Format(atrim(rstdr!horafi), "hh:nn")
        .Fields("prepclixe_observacions" + Trim(i)) = atrim(cadbl(rstdr!hores)) + " Hores de " + atrim(rstdr!tipus)
         i = i + 1
        rstdr.MoveNext
    Wend
  
 'Avaria
  Set rst = dbtmpb.OpenRecordset("select id,numeromaquina,operari,datainici,horainici,datafi,horafi,observacio from impressores where comanda=" + comanda.text + " and tipus='V'  order by datainici,horainici")
  If Not rst.EOF Then
   nmaq = rst!numeromaquina
   rst.MoveLast
   If Not rst.BOF Then rst.MovePrevious:
   If rst.BOF Then
      rst.MoveNext
    Else: rst.MovePrevious: If rst.BOF Then rst.MoveNext
   End If
  End If
  i = 4
  'posso la i a 4 per aprofitar l'ultim espai de descans/relleu del llistat per possar una avaria i
    'posso el valor a la formula etavaria del llistat perque surti el nom al llistat
  If Not rst.EOF Then
    .Fields("prepclixe_data" + Trim(i)) = Format(atrim(rst!datainici), "dd/mm/yy")
    .Fields("prepclixe_op" + Trim(i)) = cadbl(rst!operari)
    .Fields("prepclixe_de" + Trim(i)) = Format(atrim(rst!horainici), "hh:nn")
    .Fields("prepclixe_fins" + Trim(i)) = Format(atrim(rst!horafi), "hh:nn")
    .Fields("prepclixe_observacions" + Trim(i)) = atrim(rst!observacio)
    llistat.Formulas(formula) = "etavaria='Avaria Màq.:'": formula = formula + 1
  End If
  If Not rst.EOF Then rst.MoveNext
  i = 3
  'posso la i a 3 per aprofitar l'ultim espai de descans/relleu del llistat per possar una avaria i
    'posso el valor a la formula etavaria del llistat perque surti el nom al llistat
  If Not rst.EOF Then
    .Fields("prepclixe_data" + Trim(i)) = Format(atrim(rst!datainici), "dd/mm/yy")
    .Fields("prepclixe_op" + Trim(i)) = cadbl(rst!operari)
    .Fields("prepclixe_de" + Trim(i)) = Format(atrim(rst!horainici), "hh:nn")
    .Fields("prepclixe_fins" + Trim(i)) = Format(atrim(rst!horafi), "hh:nn")
    .Fields("prepclixe_observacions" + Trim(i)) = atrim(rst!observacio)
    'llistat.Formulas(formula) = "etavaria='Avaria Màq.:'": formula = formula + 1
  End If
 
  
  'prep maquina
  Set rst = dbtmpb.OpenRecordset("select id,numeromaquina,operari,datainici,horainici,datafi,horafi,observacio from impressores where comanda=" + comanda.text + " and tipus='M'  order by datainici,horainici")
  If Not rst.EOF Then
   nmaq = rst!numeromaquina
   rst.MoveLast
   If Not rst.BOF Then rst.MovePrevious:
   If rst.BOF Then
      rst.MoveNext
    Else: rst.MovePrevious: If rst.BOF Then rst.MoveNext
   End If
  End If
  i = 1
  
  While Not rst.EOF
    .Fields("prepmaquina_data" + Trim(i)) = Format(atrim(rst!datainici), "dd/mm/yy")
    .Fields("prepmaquina_op" + Trim(i)) = cadbl(rst!operari)
    .Fields("prepmaquina_de" + Trim(i)) = Format(atrim(rst!horainici), "hh:nn")
    .Fields("prepmaquina_fins" + Trim(i)) = Format(atrim(rst!horafi), "hh:nn")
    .Fields("prepmaquina_observacions" + Trim(i)) = atrim(rst!observacio)
    'miro la submaquina  tintes i clixes etc...
     'Set rs = dbtmpb.OpenRecordset("select * from submaquina where comanda=" + atrim(cadbl(rst!id)))
     'r = ""
     '  If Not rs.EOF Then
     '     r = "NT: " + atrim(cadbl(rs!tinters))
     '     r = r + " NA: " + atrim(cadbl(rs!anilox))
     '     r = r + " NC: " + atrim(cadbl(rs!camises))
     '     r = r + " NR: " + atrim(cadbl(rs!rasquetes))
     '     !canvienfilada = atrim(rs!canvienfilad)
     '  End If
    '.Fields("prepmaquina_obscosesrentades" + Trim(i)) = atrim(r)
    rst.MoveNext
    i = i + 1
  Wend
  
  'If atrim(!canvienfilada) = "" Then !canvienfilada = "No"
  'ajust imp
  Set rst = dbtmpb.OpenRecordset("select id,numeromaquina,operari,datainici,horainici,datafi,horafi,observacio,mtrsprova,paletbobprova,paletprova2,bobinaprova2,metresprova2,paletprova,bobinaprova from impressores where comanda=" + comanda.text + " and tipus='A'  order by datainici,horainici")
  If Not rst.EOF Then
   nmaq = rst!numeromaquina
   rst.MoveLast
   'If Not rst.BOF Then rst.MovePrevious:
   'If rst.BOF Then
   '   rst.MoveNext
   ' Else: rst.MovePrevious: If rst.BOF Then rst.MoveNext
   'End If
   For i = 1 To 3
      If Not rst.BOF Then
         rst.MovePrevious
        Else: rst.MoveNext: i = 10
      End If
   Next i
  End If
  i = 1
  While Not rst.EOF
    .Fields("ajustimp_data" + Trim(i)) = Format(atrim(rst!datainici), "dd/mm/yy")
    .Fields("ajustimp_op" + Trim(i)) = cadbl(rst!operari)
    .Fields("ajustimp_de" + Trim(i)) = Format(atrim(rst!horainici), "hh:nn")
    .Fields("ajustimp_fins" + Trim(i)) = Format(atrim(rst!horafi), "hh:nn")
    .Fields("ajustimp_observacio" + Trim(i)) = atrim(rst!observacio)
    .Fields("ajustimp_prova" + Trim(i)) = cadbl(rst!mtrsprova)
    r = ""
    If cadbl(rst!paletprova2) > 0 Then r = "/" + atrim(rst!paletprova2) + "-" + atrim(rst!bobinaprova2) + "m" + atrim(rst!metresprova2)
    f = ""
    If cadbl(rst!paletprova) > 0 Then f = atrim(rst!paletprova) + "-" + atrim(rst!bobinaprova)
    llistat.Formulas(formula) = "paletbobprova" + atrim(i) + "='" + f + r + "'": formula = formula + 1

    
    i = i + 1
    rst.MoveNext
  Wend
  
  'temps funcionament
  Set rst = dbtmpb.OpenRecordset("select id,numeromaquina,operari,datainici,horainici,datafi,horafi,observacio,mtrsminut,mtrsprova,totalmetres, metresparcial from impressores where comanda=" + atrim(numc) + " and tipus='F'  order by datainici,horainici")
  If Not rst.EOF Then
    nmaq = rst!numeromaquina
    rst.MoveLast
   If Not rst.BOF Then rst.MovePrevious:
   If rst.BOF Then
      rst.MoveNext
    Else: rst.MovePrevious: If rst.BOF Then rst.MoveNext Else rst.MovePrevious: If rst.BOF Then rst.MoveNext
   End If
  End If
  i = 1
  While Not rst.EOF
    .Fields("temps_data" + Trim(i)) = Format(atrim(rst!datainici), "dd/mm/yy")
    .Fields("temps_op" + Trim(i)) = cadbl(rst!operari)
    .Fields("temps_de" + Trim(i)) = Format(atrim(rst!horainici), "hh:nn")
    .Fields("temps_fins" + Trim(i)) = Format(atrim(rst!horafi), "hh:nn")
    .Fields("temps_observacio" + Trim(i)) = atrim(rst!observacio)
    .Fields("temps_mtrsmin" + Trim(i)) = cadbl(rst!mtrsminut)
    .Fields("temps_mtrs" + Trim(i)) = cadbl(rst!totalmetres) - mtrsparcialanteriors + cadbl(rst!metresparcial)
    mtrsparcialanteriors = cadbl(rst!metresparcial)
    i = i + 1
    rst.MoveNext
  Wend
  
  'acavar comandes
  Set rst = dbtmpb.OpenRecordset("select * from impresorespantones where comanda=" + atrim(comanda))
  i = 1
  If Not rst.EOF Then
   For i = 1 To 10
    .Fields("pantone" + Trim(i)) = atrim(rst.Fields("pantone" + atrim(i)))
    .Fields("lot" + Trim(i)) = atrim(rst.Fields("lot" + atrim(i)))
    .Fields("kg" + Trim(i)) = cadbl(rst.Fields("kg" + atrim(i)))
   Next i
  End If
  'POSSO els camps extres de les tintes
'  Set rst = dbtmpb.OpenRecordset("select * from tintesclixes where comanda=" + comanda.Text + " order by tinta")
'  i = 1
'  While Not rst.EOF
'       Set rstbob = dbtmpb.OpenRecordset("select descripcio from adhesiusdoblecara where codi=" + atrim(rst!idadhesiu))
'        '-----------------
'        .Fields("cilindre" + Trim(i)) = rst!cilindre
'        If Not rstbob.EOF Then .Fields("adhesiu" + Trim(i)) = rstbob!descripcio
'        .Fields("espessorpol" + Trim(i)) = rst!espessorpol
'        .Fields("numpol" + Trim(i)) = rst!numpol
'        i = i + 1
'      rst.MoveNext
'  Wend
  'posso els camps de totals
  !hclixe = cadbl(hclixe): !hmaquina = cadbl(hmaquina): !hajusts = cadbl(hajusts): !hfunc = cadbl(hfunc): !tprova = cadbl(tprova): !tbob = cadbl(tbob): !tmtrs = cadbl(tmetres): !tkilos = cadbl(tkilos): !mtrsmin = cadbl(kiloshora): !tmetresdolents = cadbl(mtrsdolents)
  '!acavada = comandaacavada
  Set rstbob = Nothing
  Set rst = Nothing
  
  
    
  End With
  
  'passo les bobines a la taula del llistat
  Set rst = dbtmpb.OpenRecordset("select id,numeromaquina,operari,datainici,horainici,datafi,horafi,observacio,mtrsminut,mtrsprova from impressores where comanda=" + atrim(numc) + " and tipus='F'")
  If rst.EOF Then dbtmpb.Execute "insert into tmp_imp_baixa_bob (operari,operari2,palet,bobent,bobsort,kilos,metres) values (0,0,0,'0',0,0,0)"
  While Not rst.EOF
''     Set rsttmp = dbtmpb.OpenRecordset("Select * from bobinesimp where controlid=" + atrim(cadbl(rst!id)))
     
  ''   While Not rsttmp.EOF
         'la seguent linia es per anar a buscar els camps extres de la bobina palet i bobentrada
        'Set rsttmp2 = dbtmpb.OpenRecordset("select * from bobentradaimp where idbobimp=" + atrim(cadbl(rsttmp!id)))
        'If Not rsttmp2.EOF Then
        ' rsttmp2.MoveLast
        ' rsttmp2.MoveFirst
        'End If
        'While Not rsttmp2.EOF
        '  If rsttmp2.AbsolutePosition + 1 = rsttmp2.RecordCount Then
        '      dbtmpb.Execute "insert into tmp_imp_baixa_bob (operari,palet,bobent,bobsort,kilos,metres) values (" + atrim(cadbl(!operari1)) + "," + atrim(cadbl(rsttmp2!idpalet)) + "," + atrim(cadbl(rsttmp2!numbob)) + "," + atrim(cadbl(!numerodebobina)) + "," + atrim(cadbl(!kilos)) + "," + atrim(cadbl(!metres)) + ")"
        '    Else: dbtmpb.Execute "insert into tmp_imp_baixa_bob (operari,palet,bobent,bobsort,kilos,metres) values (" + atrim(cadbl(!operari1)) + "," + atrim(cadbl(rsttmp2!idpalet)) + "," + atrim(cadbl(rsttmp2!numbob)) + "," + atrim(cadbl(!numerodebobina)) + "," + atrim("0") + "," + atrim("0") + ")"
        '  End If
        '  rsttmp2.MoveNext
        'Wend
        Set rsttmp2 = dbtmpb.OpenRecordset("select * from bobinesimp where controlid=" + atrim(cadbl(rst!id)))
        
        With rsttmp2
        If Not rsttmp2.EOF Then
         rsttmp2.MoveLast
         rsttmp2.MoveFirst
          Else: dbtmpb.Execute "insert into tmp_imp_baixa_bob (operari,operari2,palet,bobent,bobsort,kilos,metres) values (0,0,0,'0',0,0,0)"
        End If
        While Not rsttmp2.EOF
          'If rsttmp2.AbsolutePosition + 1 = rsttmp2.RecordCount Then
              If Not rsttmp2.EOF Then Set rsttmpbob = dbtmpb.OpenRecordset("select * from bobinesentimp where id=" + atrim(cadbl(rsttmp2!id)))
              nb = 0
              np = 0
              If Not rsttmpbob.EOF Then
                 rsttmpbob.MoveLast
                 rsttmpbob.MoveFirst
                 'aprofito per buscar lamplada del palet
                 Set rststocks = dbstocks.OpenRecordset("select ample from palets where idpalet=" + atrim(np))
                 If Not rststocks.EOF Then nample = nample
                  Else
                    dbtmpb.Execute "insert into tmp_imp_baixa_bob (operari,operari2,palet,bobent,bobsort,kilos,metres) values (" + atrim(cadbl(!operari1)) + "," + atrim(cadbl(!operari2)) + "," + atrim(np) + ",'" + atrim(nb) + "'," + atrim(cadbl(!numerodebobina)) + "," + atrim(cadbl(!kilos)) + "," + atrim(cadbl(!metres)) + ")"
              End If
              mtrs = cadbl(!metres)
              kgs = cadbl(!kilos)
              While Not rsttmpbob.EOF
                 np = rsttmpbob!palet
                 nb = rsttmpbob!bobina
                 If rsttmpbob.RecordCount > 1 Then nb = "*" + nb
                 dbtmpb.Execute "insert into tmp_imp_baixa_bob (operari,operari2,palet,bobent,bobsort,kilos,metres) values (" + atrim(cadbl(!operari1)) + "," + atrim(cadbl(!operari2)) + "," + atrim(np) + ",'" + atrim(nb) + "'," + atrim(cadbl(!numerodebobina)) + "," + passaradecimalpunt(atrim(kgs)) + "," + passaradecimalpunt(atrim(mtrs)) + ")"
                 mtrs = 0
                 kbs = 0
                 rsttmpbob.MoveNext
              Wend
              
           ' Else: dbtmpb.Execute "insert into tmp_imp_baixa_bob (operari,palet,bobent,bobsort,kilos,metres) values (" + atrim(cadbl(!operari1)) + "," + "0" + "," + "0" + "," + atrim(cadbl(!numerodebobina)) + "," + atrim("0") + "," + atrim("0") + ")"`'          End If
          rsttmp2.MoveNext
'          rsttemp!ample = nample
        Wend
    ''    rsttmp.MoveNext
     ''Wend
     rst.MoveNext
     End With
  Wend
  rsttemp!ample = nample
  rsttemp.Update
  Set rsttmp2 = Nothing
  Set rsttmpbob = Nothing
  On Error Resume Next
  dbtmpb.Execute "drop table tmp_imp_totals"
  On Error GoTo 0
  dbstocks.Execute "select * into tmp_imp_totals IN '" + cami + "' from totals_full_packinglist where comanda=" + atrim(numc)
  Set rstopcions = dbstocks.OpenRecordset("select * from opcionsdajust where comanda=" + atrim(numc))
  If Not rstopcions.EOF Then
     If cadbl(rstopcions!sistemadajust) = 1 Then mtrsajust = cadbl(rstopcions!mtrsajust)
  End If
  
  metrespacking = calcularmetrespackinglist(cadbl(comanda), "historic_packinglist")
  metrespackingfinal = calcularmetrespackinglist(cadbl(comanda), "parcials")
  vtotalmetres = metrespackingfinal    'aquestvalor es per saber els metres impresos i asignarlos a lanilox
  'imprimir llistat
  If nmaq = 5 Then llistat.Formulas(formula) = "nommaquina='" + "Màquina: FS '": formula = formula + 1
  If nmaq = 7 Then llistat.Formulas(formula) = "nommaquina='" + "Màquina: FW 1508'": formula = formula + 1
  If nmaq = 9 Then llistat.Formulas(formula) = "nommaquina='" + "Màquina: COMEXI F2'": formula = formula + 1
  llistat.Formulas(formula) = "texteimpresio='" + texteimp + "'": formula = formula + 1
  llistat.Formulas(formula) = "metrespackinglist=" + atrim(Redondejar(metrespacking, 0)) + "": formula = formula + 1
  llistat.Formulas(formula) = "metrespackinglistfinal=" + atrim(Redondejar(metrespackingfinal, 0)) + "": formula = formula + 1
  llistat.Formulas(formula) = "mtscomanda=" + atrim(Redondejar(metrescomanda, 0)) + "": formula = formula + 1
  llistat.Formulas(formula) = "mtsajust=" + atrim(Redondejar(mtrsajust, 0)) + "": formula = formula + 1
  llistat.Formulas(formula) = "lotsreprint='" + crearliniareprint + "'": formula = formula + 1
     'reals {tmp_imp_totals.r_mtrs_assignat}+{tmp_imp_totals.r_mtrs_dolents}+{tmp_imp_totals.r_ajust_estoc}+{tmp_imp_totals.r_ajust_paletbob}+{tmp_imp_totals.r_ajust_llençar}
     'totalimpresos  {tmp_imp_totals.mtrs_impresos}+{tmp_imp_totals.mtrs_dolents}+{tmp_imp_totals.r_ajust_paletbob}+{tmp_imp_totals.r_ajust_estoc}+{tmp_imp_totals.r_ajust_llençar}
        
 
 'ATENCIÓ QUE FAIG SERVIR BAIXESIMPRESSORA.RPT PERÒ LA QUE S'IMPRIMEIX ES LA BAIXESIMPRESSORA_PDF perquè també fa el pdf
 '   i amb la versió que estava fet no es podia genera el PDF
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "baixesimpressora.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 2
 llistat.DataFiles(0) = cami
 llistat.DiscardSavedData = True
 
  DoEvents
' If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 'llistat.Action = 1
  escriure_ini "General", "exportantpdfs", "si", llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\organitzar.ini"
  crearlacarpetaperexportar cadbl(comanda.text), vcarpetadesti
  exportarllistatapdf llistat, llegir_ini("General", "rutallistats", "comandes.ini") + "baixesimpressora_PDF.rpt", cadbl(comanda.text), vcarpetadesti
  escriure_ini "General", "exportantpdfs", "no", llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\organitzar.ini"
   
  Set rsttmp = Nothing
  Set rst = Nothing
  Set dbstocks = Nothing
  
  Set rsttmpbob = Nothing
  Set rsttemp = Nothing
  Set rsttmp2 = Nothing

 form1.caption = "Baixes Comandes (Impresores)"
 
End Sub

Sub crearlacarpetaperexportar(numc As Double, carpetadesti As String)
   Dim carpetaprincipal As String
   Dim vcarpetatemporal As String
   Dim vubicaciocarpetadesti As String
   Dim vnomfitxer As String
   vcarpetatemporal = rutadelfitxer(llegir_ini("General", "cami", fitxerini))
   carpetadesti = llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini")
   'si no puc accedir a la carpeta ho guardo en una temporal en el servidor fins que es pugui descarregar
   If Not existeix(carpetadesti + "\cache_Fabricacio") Then carpetadesti = vcarpetatemporal 'MkDir carpetadesti + "\cache_Fabricacio"
   carpetadesti = carpetadesti + "\cache_Fabricacio"
   
   carpetaprincipal = "Les_" + atrim(atrim(Int(cadbl(numc) / 1000)) + "000")
   If Not existeix(carpetadesti) Then MkDir carpetadesti
   If Not existeix(carpetadesti + "\" + carpetaprincipal) Then MkDir carpetadesti + "\" + carpetaprincipal
   If Not existeix(carpetadesti + "\" + carpetaprincipal + "\" + atrim(numc)) Then MkDir carpetadesti + "\" + carpetaprincipal + "\" + atrim(numc)
   vubicaciocarpetadesti = carpetadesti
   carpetadesti = carpetadesti + "\" + carpetaprincipal + "\" + atrim(numc)
   
   'comprovo si hi ha quelcom a la carpeta temporal i ho copio a la definitiva
   If InStr(1, carpetadesti, vcarpetatemporal) = 0 Then
     vnomfitxer = Dir(vcarpetatemporal + "\cache_fabricacio\*.*", vbDirectory)
     While vnomfitxer <> ""
         If vnomfitxer <> "." And vnomfitxer <> ".." Then
          Copiar_Fitxer vcarpetatemporal + "cache_fabricacio\" + vnomfitxer + "\", vubicaciocarpetadesti + "\", 5
          borra_carpeta vcarpetatemporal + "cache_fabricacio\" + vnomfitxer
          vnomfitxer = Dir(vcarpetatemporal + "\cache_fabricacio\*.*", vbDirectory)
         End If
         vnomfitxer = Dir
     Wend
   End If
End Sub
Sub borra_carpeta(strRuta As String)
    'Elimina la carpeta sin necesidad de eliminar los ficheros en ella contenidos
   ' MsgBox "Eliminar " + strRuta
    Dim FSO As Object
    Dim i As Byte
    'quitamos la posible última barra \ de la ruta
    If Right(strRuta, 1) = "\" Then strRuta = Left(strRuta, Len(strRuta) - 1)
    'llamamos al script de FileSystem
    Set FSO = CreateObject("Scripting.FileSystemObject")
    'y acabamos borrando la carpeta y todo su contenido
    On Error GoTo fi
    If existeix(strRuta) Then
      FSO.DeleteFolder strRuta, True
      For i = 1 To 12
         If existeix(strRuta) Then wait 5
      Next i
    End If
fi:
    'MsgBox "Fi de Eliminar " + strRuta
    
End Sub

Function crearliniareprint() As String
  Dim rst As Recordset
  Dim vtot As Double
  If Not estemfentreprint Then GoTo fi
  Set rst = dbtmpb.OpenRecordset("select * from impresores_llaunesgastades where tipus='R' and comanda=" + atrim(comanda))
  If Not rst.EOF Then crearliniareprint = "REPRINT: " + etvernis + " Lots: "
  While Not rst.EOF
    vtot = vtot + cadbl(rst!kg)
    crearliniareprint = crearliniareprint + " " + atrim(rst!numllauna)
    rst.MoveNext
  Wend
  crearliniareprint = crearliniareprint + " TOTAL: " + atrim(vtot) + "Kg"
fi:
  Set rst = Nothing
End Function
Function possarnomfirmat() As String
  Dim rsttmp As Recordset
  Set rsttmp = dbtmp.OpenRecordset("select descripcio from operaris where maquina='I' and codi=" + atrim(cadbl(firmat)))
  If Not rsttmp.EOF Then
     possarnomfirmat = atrim(rsttmp!descripcio)
  End If
End Function
Function existeixlataula(vnomtaula As String) As Boolean
  Dim rstp As Recordset
  On Error GoTo errortaula
  existeixlataula = True
  Set rstp = dbtmpb.OpenRecordset("select * from " + vnomtaula)
  Set rstp = Nothing
  Exit Function
errortaula:
  existeixlataula = False
End Function
Sub crear_taula_impresio_baixa()
  Dim camps As String
  Dim camps2 As String
  Dim campspantone As String
  Dim campstotal As String
  If Not existeixlataula("tmp_imp_baixa") Or Not existeixlataula("tmp_imp_baixa_bob") Then
        campsextra = " nomfirmat text,firmat text,espessor double,"
        camps = " comanda double,client string,ample double ,comandaacavada byte,tintersrentats byte,portaclixers byte,canvienfilada string,numclixes byte,numtintes byte,cilindre double,"
        camps = camps + "prepclixe_data1 string,prepclixe_op1 byte, prepclixe_de1 string,prepclixe_fins1 string,prepclixe_observacions1 string,prepmaquina_obscosesrentades1 string,"
        camps = camps + "prepclixe_data2 string,prepclixe_op2 byte, prepclixe_de2 string,prepclixe_fins2 string,prepclixe_observacions2 string,prepmaquina_obscosesrentades2 string,"
        camps = camps + "prepclixe_data3 string,prepclixe_op3 byte, prepclixe_de3 string,prepclixe_fins3 string,prepclixe_observacions3 string,prepmaquina_obscosesrentades3 string,"
        camps = camps + "prepclixe_data4 string,prepclixe_op4 byte, prepclixe_de4 string,prepclixe_fins4 string,prepclixe_observacions4 string,prepmaquina_obscosesrentades4 string,"
        
        camps = camps + " prepmaquina_data1 string,prepmaquina_op1 byte,prepmaquina_de1 string,prepmaquina_fins1 string,prepmaquina_observacions1 string ,obscosesrentades1 string,"
        camps = camps + " prepmaquina_data2 string,prepmaquina_op2 byte,prepmaquina_de2 string,prepmaquina_fins2 string,prepmaquina_observacions2 string ,obscosesrentades2 string,"
        camps3 = " prepmaquina_data3 string,prepmaquina_op3 byte,prepmaquina_de3 string,prepmaquina_fins3 string,prepmaquina_observacions3 string ,obscosesrentades3 string,"
        camps3 = camps3 + " prepmaquina_data4 string,prepmaquina_op4 byte,prepmaquina_de4 string,prepmaquina_fins4 string,prepmaquina_observacions4 string ,obscosesrentades4 string,"
        
        camps3 = camps3 + "ajustimp_data1 string,ajustimp_op1 byte,ajustimp_de1 string, ajustimp_fins1 string,ajustimp_prova1 double,ajustimp_observacio1 string,"
        camps3 = camps3 + "ajustimp_data2 string,ajustimp_op2 byte,ajustimp_de2 string, ajustimp_fins2 string,ajustimp_prova2 double,ajustimp_observacio2 string,"
        camps3 = camps3 + "ajustimp_data3 string,ajustimp_op3 byte,ajustimp_de3 string, ajustimp_fins3 string,ajustimp_prova3 double,ajustimp_observacio3 string,"
        camps2 = "ajustimp_data4 string,ajustimp_op4 byte,ajustimp_de4 string, ajustimp_fins4 string,ajustimp_prova4 double,ajustimp_observacio4 string,"
        
        camps2 = camps2 + "temps_data1 string,temps_op1 byte, temps_de1 string,temps_fins1 string, temps_mtrsmin1 double,temps_mtrs1 double, temps_observacio1 string,"
        camps2 = camps2 + "temps_data2 string,temps_op2 byte, temps_de2 string,temps_fins2 string, temps_mtrsmin2 double,temps_mtrs2 double, temps_observacio2 string,"
        camps2 = camps2 + "temps_data3 string,temps_op3 byte, temps_de3 string,temps_fins3 string, temps_mtrsmin3 double,temps_mtrs3 double, temps_observacio3 string,"
        camps2 = camps2 + "temps_data4 string,temps_op4 byte, temps_de4 string,temps_fins4 string, temps_mtrsmin4 double,temps_mtrs4 double, temps_observacio4 string,"
        camps2 = camps2 + "temps_data5 string,temps_op5 byte, temps_de5 string,temps_fins5 string, temps_mtrsmin5 double,temps_mtrs5 double, temps_observacio5 string,"
        'creo els camps dels pantone
        For i = 1 To 10
          campspantone = campspantone + "pantone" + Trim(i) + " string, lot" + Trim(i) + " string,kg" + Trim(i) + " double, "
          campspantone2 = campspantone2 + "cilindre" + Trim(i) + " double, adhesiu" + Trim(i) + " string,espessorpol" + Trim(i) + " double,numpol" + Trim(i) + " byte, "
        Next i
        campspantone2 = campspantone2 + " fi string,  texteimp string,"
        
        'creo els camps de total
        campstotal = "hclixe double, hmaquina double, hajusts double, hfunc double, tprova double, tbob double,tmtrs double, tkilos double, mtrsmin double,tmetresdolents double "
        
        'ample double,plegat double,solapa double,espessor double,metres double,kilos double)"
        
        dbtmpb.Execute ("create table tmp_imp_baixa (" + campsextra + camps + camps3 + camps2 + campspantone + campspantone2 + campstotal) + ")"
        dbtmpb.Execute ("create table tmp_imp_baixa_bob (idbob integer,operari byte,operari2 byte,palet double,bobent string,bobsort integer,kilos double,metres double)")
           Else
              dbtmpb.Execute "delete * from tmp_imp_baixa"
              dbtmpb.Execute "delete * from tmp_imp_baixa_bob"
              
    End If
  
End Sub



Private Sub Command8_Click()
calcular_totals
If command15.tag = "Error" Then MsgBox "Hi ha un error a la comanda revisala": ratoli "normal": Exit Sub
wait 2
guardar_totals_packinglist cadbl(comanda)
If stockopacking = "E" Then imprimir_packinglist cadbl(comanda)
imprimir_fulla
End Sub

Private Sub Command9_Click()
If horaapretada <> 1 Then
    dblots.AllowAddNew = False
    dblots.AllowDelete = False
    dblots.AllowUpdate = False
    dblots.MarqueeStyle = 3
    dblots.visible = False
  framepantones.visible = Not framepantones.visible
  framepantones.ZOrder 0
  frameempalmes.visible = False
  framebobentrada.visible = False
  Framereprint.visible = False
End If
End Sub

Private Sub Command9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If horaapretada = 0 Then horaapretada = Now
  
End Sub

Private Sub Command9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
    horaapretada = 0
 
End Sub

Private Sub compantone_DblClick(Index As Integer)
  lots.Refresh
  dblots.visible = True
  dblots.tag = Index
End Sub

Private Sub compantone_LostFocus(Index As Integer)
  imppantones.Refresh
End Sub

Private Sub controlstock_Click()
  'obrir_llegirllaunes
End Sub

Private Sub cpostitcomanda_DblClick()
   canviartipificacioavaria
End Sub
Sub enviaremailsical_avaria(vtipificacio As String, vobservacio As String)
   Dim vdestinatari As String
   Dim vassumpte As String
   If InStr(1, UCase(vtipificacio), " TINTES") > 0 Or InStr(1, UCase(vtipificacio), " TINTA") > 0 Then
        vdestinatari = "grupimpressores@inplacsa.com"
       Else: vdestinatari = "jmiralles@inplacsa.com;impresores@inplacsa.com"
   End If
  vassumpte = "Comanda " + atrim(comanda) + " Avaria a " + nommaq + " - " + nomoperari + Chr(13) + Chr(10) + texteimpresio + vbNewLine + vbNewLine + "Observació de l'OPERARI: " + vobservacio
  enviaremailgeneric vdestinatari, vassumpte, treure_apostruf("Avaria... Motiu: " + UCase(vtipificacio) + Chr(13) + Chr(10) + vassumpte)
End Sub
Sub canviartipificacioavaria()
    Dim vtipusavaria As String
    vtipusavaria = escullir_avaria
    If vtipusavaria = "" Then vtipusavaria = InputBox("Escriu la descripció de l'avaria", "Error")
    If vtipusavaria <> "" Then
       If impresores.Recordset.EditMode = 0 Then impresores.Recordset.Edit
       impresores.Recordset!tipificacioavaria = Mid(vtipusavaria, 1, 50)
       impresores.Recordset.Update
       cpostitcomanda = vtipusavaria
       enviaremailsical_avaria vtipusavaria, ""
    End If
End Sub

Private Sub dblots_DblClick()
  Dim vbuscarlots As String
  Dim vnumlot As String
  Dim vllistalots(50) As String
  If dblots.MarqueeStyle = 6 Then
    dblots.visible = False
    dblots.AllowAddNew = False
    dblots.AllowDelete = False
    dblots.AllowUpdate = False
    dblots.MarqueeStyle = 3
    dblots.visible = False
    framepantones.visible = False
    Exit Sub
  End If
  If Not lots.Recordset.EOF Then
    'pantone(cadbl(dblots.tag)) = Mid(atrim(lots.Recordset!nomlot), 1, InStr(1, atrim(lots.Recordset!nomlot) + "#", "#") - 1)
    vbuscarlots = buscarlots(atrim(lots.Recordset!nomlot), "", vllistalots)
    vnumlot = IIf(atrim(lots.Recordset!codilot) = "" Or atrim(lots.Recordset!codilot) <> "0", atrim(lots.Recordset!codilot), "")
    compantone(cadbl(dblots.tag)) = vnumlot + IIf(vbuscarlots <> "" And vnumlot <> "", "+" + vbuscarlots, vbuscarlots)
  End If
  dblots.visible = False
End Sub
Function buscarlots(vnomlot As String, vcodilot As String, Optional vllistalots As Variant) As String
  Dim v As String
  Dim vnumerodelot As String
  Dim vcont As Byte
  vcont = 0
  If InStr(1, UCase(vcodilot), "I") > 0 Then
     vnumerodelot = agafarellotdelcomponent("#" + vcodilot)
     
     If vnumerodelot <> "" And vnumerodelot <> "0" Then
         vllistalots(vcont) = vnumerodelot
         vcont = vcont + 1
     End If
  End If
  
  If InStr(1, vnomlot, "#") > 0 Then
        v = "  " + vnomlot + "  "
        While InStr(1, v, "#") > 0 And vcont < 50
          v = Mid(v, InStr(2, v, "#"))
          vnumerodelot = agafarellotdelcomponent(atrim(Mid(v, 1, InStr(2, v, " ") - 1)))
          vllistalots(vcont) = vnumerodelot
          buscarlots = buscarlots + IIf(buscarlots <> "", "+", "") + vnumerodelot
          If InStr(2, v, "#") = 0 Then v = ""
          vcont = vcont + 1
        Wend
      Else: buscarlots = agafarellotdelcomponent("#" + vcodilot): vllistalots(0) = buscarlots
  End If
End Function
Function agafarellotdelcomponent(vdosificador As String, Optional vnomlot As String, Optional vcoditinta As String) As String
  Dim rst As Recordset
  vdosificador = UCase(vdosificador)
  If vdosificador = "#" Then GoTo fi
  If InStr(1, vdosificador, "#A") Then agafarellotdelcomponent = Mid(vdosificador, 2): GoTo fi
  vdosificador = substituir("  " + vdosificador, "#I", "")
  sql = "SELECT Componentsbase.nomcomponent AS nomlot,componentsbase.coditintarelacionada, detallnumeroslotsbase.numerodelot AS codilot FROM detallnumeroslotsbase INNER JOIN Componentsbase ON detallnumeroslotsbase.idcomponent = Componentsbase.idcomponent "
  sql = sql + " WHERE Componentsbase.numdosificador=" + atrim(vdosificador) + " order by data DESC"
  Set rst = lots.Database.OpenRecordset(sql)
  If Not rst.EOF Then
    vnomlot = atrim(rst!nomlot): agafarellotdelcomponent = rst!codilot
    If atrim(rst!coditintarelacionada) <> "" Then vcoditinta = atrim(rst!coditintarelacionada)
  End If
fi:
  Set rst = Nothing
End Function


Private Sub dblots_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then dblots.tag = "": dblots.visible = False
End Sub

Private Sub eliminarbobentrada_Click()
  If bobinesent.Recordset.EOF Then Exit Sub
  If MsgBox("Segur que vols eliminar la bobina d'entrada " + atrim(bobinesent.Recordset!palet) + "/" + atrim(bobinesent.Recordset!bobina), vbExclamation + vbYesNo, "Borrar bobina d'entrada") = vbYes Then
    'carregar_bobinesdentrada "marcarutilitzada", , bobinesent.Recordset!palet, bobinesent.Recordset!bobina, cadbl(comanda), False
    bobinesent.Recordset.Delete
    bobinesent.Refresh
    bobinesent.UpdateControls
  End If
End Sub

Private Sub ettreball_DblClick()
'  Dim vidtreball As Integer
'  vidtreball = cadbl(InputBox("Entra el treball que vols consultar.", "Consultar observacions", atrim(ettreball.tag)))
'  If vidtreball > 0 Then observacio_idtreball vidtreball
MsgBox "S'ha d'obrir pel botó de mes avall", vbInformation, "Atenció "
End Sub

Sub observacio_idtreball(numid As Integer, Optional comprovar As Boolean)
Dim rst As Recordset
  If numid = 0 Then Exit Sub
  Set rst = dbtmpb.OpenRecordset("select * from idstreball where id=" + atrim(numid))
  If rst.EOF Then rst.AddNew: rst!obsidtreball = " ": rst!id = numid: rst.Update
  Set rst = dbtmpb.OpenRecordset("select * from idstreball where id=" + atrim(numid))
  If comprovar Then
     GoTo coloretiqueta
  End If
  Load obsidtreball
  obsidtreball.obsid.text = rst!obsidtreball
 
  obsidtreball.Show 1
  If atrim(r) <> "" Then
     rst.Edit
     rst!obsidtreball = r
     rst.Update
       Else: rst.Delete
  End If
  
  'posar color etiqueta
coloretiqueta:
  Set rst = dbtmpb.OpenRecordset("select * from idstreball where id=" + atrim(ettreball.tag))
  ettreball.ForeColor = QBColor(1): ettreball.BorderStyle = 0
  If rst.EOF Then GoTo fi
  If atrim(rst!obsidtreball) <> "" Then
       ettreball.ForeColor = QBColor(12)
       ettreball.BorderStyle = 1
  End If
fi:
  Set rst = Nothing
  
End Sub
Private Sub firmat_DblClick()
If firmat <> "" Then
   firmat = ""
  Else: firmar_fulla
End If
End Sub

Sub veureelimp2()
 Dim nomfitxer As String
  Dim nomcarpeta As String
  Dim rstc As Recordset
  Dim rstc2 As Recordset
  Dim ruta As String
  Dim ruta_relativa_docs As String
  
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  
  Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(comanda)))
  If rstc.EOF Then Exit Sub
  nomfitxer = atrim(rstc!arxiuimpressora)
  Set rstc2 = dbtmp.OpenRecordset("select * from carpeta_client where codiclient=" + atrim(cadbl(rstc!client)))
  If Not rstc2.EOF Then nomcarpeta = rstc2!nomcarpeta
  If cadbl(Mid(nomfitxer, 1, 6)) = 0 Then nomfitxer = nomcarpeta + Mid(nomfitxer, InStr(1, nomfitxer, "\"))
  ruta = ruta_relativa_docs + "\" + nomfitxer
  If existeix(ruta) Then
     obrir_document ruta
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + ruta, vbCritical, "Error"
  End If

End Sub


Function isloaded(vnomform As String) As Boolean
  Dim f
  For Each f In Forms
   If f.Name = vnomform Then
         isloaded = True
   End If
  Next
End Function

Sub obrir_llegirllaunes()
'  If nummaq = 9 Then
   Command26.tag = "llegirllaunes"
   Command26_Click
 ' End If
End Sub
Sub ensenyar_formanilox_i_tancar(Optional notancar As Boolean)
 Unload formaniloxos
  Load formaniloxos
  formaniloxos.tag = Me.comanda
  formaniloxos.boto_nou(0).tag = Command26.tag
  If vestemfentfingerprint Then
    formaniloxos.checkeditar.tag = "fingerprint"
  End If
  formaniloxos.Show
  wait 2
  formaniloxos.ensenyar_els_lots
  If Not notancar Then Unload formaniloxos
End Sub

Sub feines_parar_engegar_maquina(vTipus As String, Optional vnummaq As Double)
  Dim vreiniciarchecks As Boolean
  Load formcosesengegariparar
  With formcosesengegariparar
   .desmarcarTOTS.tag = ""
   If vTipus = "ENGEGAR" Then .bparar.BackColor = &H8000000F: .bengegar.BackColor = &H17D062
   If vTipus = "PARAR" Then
        If MsgBox("Vols BORRAR totes les FEINES FETES?", vbCritical + vbDefaultButton2 + vbExclamation, "ATENCIÓ") = vbYes Then vreiniciarchecks = True
        vTipus = "ENGEGAR": .bengegar.BackColor = &H8000000F: .bparar.BackColor = &H17D062
        .desmarcarTOTS.tag = "1"
   End If
   .carregar_feines vnummaq, vTipus
   If vreiniciarchecks Then .passartotesaUNCHECK
   .Show 1
  End With
End Sub

Private Sub Form_Click()
  Dim vcarpetadesti As String
  Dim rst As Recordset
  Dim rst2 As Recordset
  
  
 ' borrar_bobines_aimpresoresquejanohison
   ' formbobinesaimpresores.ajustar_diametre_real "54098/1"
  'passar_comanda_a_acavada
  'MsgBox formbobinesaimpresores.comprovar_si_shautilitzat(53942, 1)
  'mirar_bobinesdentrada_noacavades
  'demanarvalorsdelta cadbl(comanda)
'imprimir_full_arrancar_rentar cadbl(comanda)

  Exit Sub
  Set rst = dbtmpb.OpenRecordset("select * from impressorestot where dataimpressio>#11/01/2022#")
  While Not rst.EOF
     Set rst2 = dbtmpb.OpenRecordset("select * from impressores where comanda=" + atrim(rst!comanda) + " and tipus='F' order by id desc")
     If Not rst2.EOF Then
        If rst!dataimpressio = rst2!datafi Then
            rst.Edit
            rst!dataimpressio = CVDate(atrim(rst2!datafi) + "  " + atrim(rst2!horafi))
            rst.Update
        End If
     End If
     rst.MoveNext
  Wend
  
  'If nohihatotselslots Then
  '  If avisnohihatotselslots Then Exit Sub
 'End If
' passar_comanda_a_acavada
  Exit Sub
  Set rst = dbtmpb.OpenRecordset("select distinct comanda from impressores where year(datainici)>2020")
  rst.MoveLast
  rst.MoveFirst
  While Not rst.EOF
    Me.caption = atrim(rst.AbsolutePosition) + "/" + atrim(rst.RecordCount)
    crearlacarpetaperexportar cadbl(rst!comanda), vcarpetadesti
    vcarpetadesti = "\\ord_copies\comandespdf\Les_" + Mid(atrim(rst!comanda), 1, 3) + "000" + "\" + atrim(rst!comanda)
    If Not existeix(vcarpetadesti + "\" + atrim(rst!comanda) + "_BaixaImpresores.pdf") Then MsgBox atrim(rst!comanda)
    rst.MoveNext
    DoEvents
  Wend
  Set rst = Nothing
'demanar_metres_arrancada
  ' guardar_metres_rasquetes
'mantenimentbobina.passarbobinaaacavada 11282, 1
' imprimir_etiquetallaunesdetintapercomanda
'imprimir_packinglistTICKET 201192
'guardar_aniloxositintesutilitzadescomadefinitives
  'passaravisdebobinaassignadaaunaaltracomanda 51592, 1, 199524
'  preparaelPDF "C:\Users\Usuari_Prog\Desktop\exemple.pdf", 0

 'MsgBox buscardadesbasiquescomanda(comanda)
' passarfitxertemperaturesf2 "", "", 198111
'  formconsumtintes.Show
 'Printer.Scale
 'formannex.PrintForm
imprimir_full_arrancar_rentar cadbl(comanda)
 Exit Sub
  
 ' Set rst = dbtmp.OpenRecordset("select * from comandes where materialex>=500 order by comanda desc")
 Set rst = dbtmp.OpenRecordset("SELECT comandes.comanda, InStr(1,[codi],'I') AS Expr1 FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((InStr(1,[codi],'I'))>0)) and comandes.materialex>499;")
  While Not rst.EOF
   guardar_fitxer_temperatures rst!comanda
   rst.MoveNext
  Wend
  Exit Sub
  
  Set rst = dbtmpb.OpenRecordset("select * from registrecomandes")
  dbtmpb.Execute "delete * from registrecomandes"
  r = Dir(nomordinadorcomexi + "\*.*")
  While r <> ""
    r = Dir
     If r <> "" Then
        rst.AddNew
        rst!comanda = r
        rst.Update
     End If
  Wend
MsgBox "Fet"
  'calcularvalorsreducciocilindre cadbl(comanda), nummaq
'obrestocks
'des_reservar 147000
 'mantenimentbobina.passaravis 0, 0, "Hi ha hagut massa metres gastats en aquesta comanda" + Chr(10) + Chr(13) + "La comanda es de " + atrim(metrescomanda) + " metres i s'han consumit " + atrim(totalmetres) + " metres"
'  MsgBox comanda.Container.Name
 'carregar_bobinesdentrada "marcarutilitzadademanar", , 18866, 1, 148779
  'mantenimentbobina.passaravis 14320, 1, "prova d'avis"
 ' netejarreport llistat
End Sub
Function guardar_fitxer_temperatures(numc As Double) As String
   Dim fitxerorigen As String
   Dim fitxerdesti As String
   Dim plantilla As String
   Dim rstb As Recordset
   Me.caption = atrim(numc)
   'fitxerdesti = llegir_ini("ruta", "ruta_documentacio_temperatures", rutadelfitxer(cami) + "valorsprograma.ini") + "\" + atrim(numc) + ".txt"
   fitxerdesti = "c:\temp\temperatures\" + atrim(numc) + ".txt"
   DoEvents
   
   Set rstb = dbtmpb.OpenRecordset("select * from impressores where comanda=" + atrim(numc))
   If rstb.EOF Then Exit Function
   If existeix(fitxerdesti) Then Kill fitxerdesti  ' Exit Function
   
   If existeix("c:\ordprog.ini") Then
    If rstb!numeromaquina = 7 Then nomordinadorcomexi = "\\comexi-service\M2833_FIC"
    If rstb!numeromaquina = 9 Then nomordinadorcomexi = "\\comexipc\files"
   End If
 fitxerorigen = nomordinadorcomexi + "\" + buscarcomandaacomexi(numc)
 If buscarcomandaacomexi(numc) = "" Or Not existeix(fitxerorigen) Then
     generarfitxernoudetemperatures numc: Exit Function
 End If
   
' guardar_fitxer_temperatures = passarfitxertemperaturesfw(fitxerorigen, fitxerdesti, numc)
 If rstb!numeromaquina = 7 Then guardar_fitxer_temperatures = passarfitxertemperaturesfw(fitxerorigen, fitxerdesti, numc)
 If rstb!numeromaquina = 9 Then guardar_fitxer_temperatures = passarfitxertemperaturesf2(fitxerorigen, fitxerdesti, numc)
   
   
   
End Function
Sub generarfitxernoudetemperatures(numc As Double)
   Dim fitxerdesti As String
   Dim fitxerorigen As String
   Dim rstc As Recordset
   Dim rstb As Recordset
   Dim rst As Recordset
   Dim rstm As Recordset
   Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
   Set rstb = dbtmpb.OpenRecordset("select * from impressores where comanda=" + atrim(numc) + " order by tipus")
   Set rstm = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(rstc!materialex))
   'fitxerdesti = llegir_ini("ruta", "ruta_documentacio_temperatures", rutadelfitxer(cami) + "valorsprograma.ini") + "\" + atrim(numc) + ".txt"
   fitxerdesti = "c:\temp\temperatures\" + atrim(numc) + ".txt"
   Set rst = dbtmp.OpenRecordset("SELECT comandes.comanda, comandes.materialex, materials.familia FROM materials INNER JOIN comandes ON materials.codi = comandes.materialex WHERE (((materials.familia)=" + atrim(rstm!familia) + "))" + " order by comanda desc")
   While Not rst.EOF And fitxerorigen = ""
      fitxerorigen = buscarcomandaacomexi(rst!comanda)
     rst.MoveNext
   Wend
   If rst.EOF Then Exit Sub 'MsgBox "No s'ha trobat cap comanda semblant a la " + atrim(numc): Exit Sub
   fitxerorigen = nomordinadorcomexi + "\" + fitxerorigen
   If rstb!numeromaquina = 9 Then inventarfitxertemperaturesf2 fitxerorigen, fitxerdesti, numc
   If rstb!numeromaquina <> 9 Then inventarfitxertemperaturesfw fitxerorigen, fitxerdesti, numc
   
End Sub
Sub inventarfitxertemperaturesf2(forigen As String, fdesti As String, numc As Double)
   Dim rstb As Recordset
   Dim rstc As Recordset
   Dim plantilla As String
   Dim vorigen As String
   Dim vplantilla As String
   Set rstb = dbtmpb.OpenRecordset("select * from impressores where  comanda=" + atrim(numc) + " order by tipus")
   Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
   If existeix(fdesti) Then Kill fdesti
   Open forigen For Input As 1
   Open fdesti For Output As 2
   Print #2, "########  #######  "
   Print #2, "##       ##     ## "
   Print #2, "##              ## "
   Print #2, "######    #######  "
   Print #2, "##       ##        "
   Print #2, "##       ##        "
   Print #2, "##       ######### "
   Print #2, "..."
   While Not EOF(1)
    vplantilla = ""
    vorigen = ""
    Input #1, vplantilla
    Input #1, vorigen
    If vplantilla = "ÿþFILE NAME **************************************************************************************" Then
       vorigen = atrim(numc)
    End If
    If vplantilla = "date creation" Then vorigen = Format(rstb!datainici, "mm/dd/yyyy")
    If vplantilla = "time creation" Then vorigen = Format(rstb!horainici, "hh:nn:ss")
    If vplantilla = "material thickness microns or units" Then vorigen = atrim(rstc!espessor)
    If vplantilla = "material width mm or units" Then vorigen = atrim(cadbl(rstc!ampleesq) * 10)
    If vplantilla = "SP temperature decks ºC  or units" Then vorigen = minimtinterotunel(rstc, "tinter") + CInt(Int((6 * Rnd()) + 2)): vplantilla = vplantilla + "     <------------------------- TINTERS"
    If vplantilla = "SP temperature tunel ºC  or units" Then vorigen = minimtinterotunel(rstc, "tunel") + CInt(Int((6 * Rnd()) + 2)): vplantilla = vplantilla + "     <------------------------- TUNEL"
    Print #2, vorigen + Chr(9) + vplantilla
   Wend
   Close #1
   Close #2
End Sub
Sub inventarfitxertemperaturesfw(forigen As String, fdesti As String, numc As Double)
   Dim rstb As Recordset
   Dim rstc As Recordset
   Dim plantilla As String
   Dim vorigen As String
   Dim vplantilla As String
   Set rstb = dbtmpb.OpenRecordset("select * from impressores where tipus='A' and comanda=" + atrim(numc))
   If rstb.EOF Then Set rstb = dbtmpb.OpenRecordset("select * from impressores where  comanda=" + atrim(numc))
   Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
   plantilla = llegir_ini("General", "rutallistats", fitxerini) + "plantilla_temperatures_fw.txt"
   If existeix(fdesti) Then Kill fdesti
   Open forigen For Input As 1
   Open fdesti For Output As 2
   Open plantilla For Input As 3
   Print #2, "######## ##      ## "
   Print #2, "##       ##  ##  ## "
   Print #2, "##       ##  ##  ## "
   Print #2, "######   ##  ##  ## "
   Print #2, "##       ##  ##  ## "
   Print #2, "##       ##  ##  ## "
   Print #2, "##        ###  ###  "
   Print #2, "..."
   While Not EOF(3) And Not EOF(1)
    vplantilla = ""
    vorigen = ""
    Input #3, vplantilla
    Input #1, vorigen
    If vplantilla = "File name" Then
       vorigen = atrim(numc) + ".fic"
    End If
    If vplantilla = "Date" Then vorigen = Format(rstb!datainici, "dd/mm/yyyy")
    If vplantilla = "Grammage" Then vorigen = atrim(rstc!espessor)
    If vplantilla = "Reel width" Then vorigen = atrim(cadbl(rstc!ampleesq) * 10)
    If vplantilla = "Decks temperature" Then vorigen = minimtinterotunel(rstc, "tinter") + CInt(Int((6 * Rnd()) + 2)):: vplantilla = vplantilla + "     <------------------------- TINTERS"
    If vplantilla = "Tunnel temperature" Then vorigen = minimtinterotunel(rstc, "tunel") + CInt(Int((6 * Rnd()) + 2)):: vplantilla = vplantilla + "     <------------------------- TUNEL"
    Print #2, vorigen + Chr(9) + vplantilla
   Wend
   Close #1
   Close #2
   Close #3
End Sub
Function micresmaterial(codimesuralineal As Byte, espesor As Double, tubolam As String) As String
  Dim rstmesural As Recordset
  Dim descripcio As String
 ' Dim r As String
  Set rstmesural = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(codimesuralineal))
  If rstmesural.EOF Then Exit Function
  descripcio = rstmesural!descripcio
  r = espesor
  If descripcio = "GALGUES" Then
            If tubolam = "T" Then
                 r = Format(espesor / 4, "#,##0.0")
                  Else: r = Format(espesor / 2, "#,##0.0")
            End If
  End If
  If InStr(1, descripcio, "GR/") > 0 Then
    micresmaterial = espesor * -1
  End If
  descripcio = IIf(descripcio = "MICRES", "Mic", descripcio)
  descripcio = IIf(descripcio = "GALGUES", "Mic", descripcio)
  If InStr(1, descripcio, "GR/") > 0 Then
     descripcio = "GR/MT2"
     r = cadbl(r) * -1
  End If
     
  micresmaterial = r
  r = descripcio
End Function
Function substituir(cadena As String, buscar As String, canviar As String) As String
   comença = InStr(1, cadena, buscar) - 1
   If comença < 1 Then substituir = cadena: Exit Function
   acaba = comença + Len(buscar) + 1
   cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   substituir = cadena
   'MsgBox linia
End Function

Function degramsamicres(codimat As Double) As Double
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("Select micresdelsgrm2 from materials where codi=" + atrim(codimat))
   If Not rst.EOF Then degramsamicres = IIf(IsNull(rst!micresdelsgrm2), 0, rst!micresdelsgrm2)
End Function
Function minimtinterotunel(rstc As Recordset, tinterotunel As String, Optional maxim As Boolean) As Double

   Dim rst As Recordset
   Dim micres As String
   minimtinterotunel = 0
   micres = micresmaterial(cadbl(rstc!mesuraesp), rstc!espessor, atrim(rstc!tubolam))
   If micres < 0 Then micres = micres * -1 ' degramsamicres(rstc!materialex)
   micres = substituir(micres, ",", ".")
   Set rst = dbtmp.OpenRecordset("SELECT  materials.codi, materialstoleranciestemp.toleranciatinterde, materialstoleranciestemp.toleranciatunelde,materialstoleranciestemp.toleranciatunela,materialstoleranciestemp.toleranciatintera FROM materials INNER JOIN materialstoleranciestemp ON materials.familia = materialstoleranciestemp.codifammaterial where materials.codi=" + atrim(rstc!materialex) + " and micresde<=" + micres + " and micresa>=" + micres)
   If Not rst.EOF Then
     If tinterotunel = "tinter" Then minimtinterotunel = cadbl(rst!toleranciatinterde)
     If tinterotunel = "tunel" Then minimtinterotunel = cadbl(rst!toleranciatunelde)
     If maxim Then
        If tinterotunel = "tinter" Then minimtinterotunel = cadbl(rst!toleranciatintera)
        If tinterotunel = "tunel" Then minimtinterotunel = cadbl(rst!toleranciatunela)
     End If
   End If
   Set rst = Nothing
End Function
Function convertirtemp(v As String) As String
    If InStr(1, v, ".") Then
           v = Mid(v, 1, InStr(1, v, ".") - 1) + Mid(v, InStr(1, v, ".") + 1)
           v = cadbl(v) / 10
     End If
     convertirtemp = v
End Function
Function passarfitxertemperaturesf2(forigen As String, fdesti As String, numc As Double) As String
   Dim plantilla As String
   Dim vorigen As String
   Dim vplantilla As String
   Dim vdetallcomanda As String
   Dim rstc As Recordset
   If existeix(fdesti) Then Kill fdesti
   vdetallcomanda = buscardadesbasiquescomanda(atrim(numc))
   Set rstc = dbtmp.OpenRecordset("SELECT comandes.*, clients.nom FROM comandes INNER JOIN clients ON comandes.client = clients.codi where comanda = " + atrim(cadbl(numc)))
   If rstc.EOF Then Exit Function
   Open forigen For Input As 1
   Open fdesti For Output As 2
   Print #2, "########  #######  "
   Print #2, "##       ##     ## "
   Print #2, "##              ## "
   Print #2, "######    #######  "
   Print #2, "##       ##        "
   Print #2, "##       ##        "
   Print #2, "##       ######### "
   Print #2, "  "
   While Not EOF(1)
    rectificar = False
    vplantilla = ""
    vorigen = ""
    Input #1, vplantilla
    Input #1, vorigen
    If atrim(vorigen) = "" Then vorigen = "0"
    If vplantilla = "SP temperature decks ºC  or units" Then
        vorigen = convertirtemp(vorigen)
        If CDbl(vorigen) < minimtinterotunel(rstc, "tinter") Or CDbl(vorigen) > minimtinterotunel(rstc, "tinter", True) Then rectificar = True
        If rectificar Then passarfitxertemperaturesf2 = passarfitxertemperaturesf2 + "Error de temperatura de tinters. Valor: " + vorigen + "  Minim: " + atrim(minimtinterotunel(rstc, "tinter", False)) + "    Màxim:" + atrim(minimtinterotunel(rstc, "tinter", True)) + Chr(10)
        'vplantilla = vplantilla + "     <------------------------- TINTERS"
    End If
    If vplantilla = "SP temperature tunel ºC  or units" Then
        vorigen = convertirtemp(vorigen)
        If CDbl(vorigen) < minimtinterotunel(rstc, "tunel") Or CDbl(vorigen) > minimtinterotunel(rstc, "tunel", True) Then rectificar = True
        If rectificar Then passarfitxertemperaturesf2 = passarfitxertemperaturesf2 + "Error de temperatura del Tunel. Valor: " + vorigen + "  Minim: " + atrim(minimtinterotunel(rstc, "tunel", False)) + "    Màxim:" + atrim(minimtinterotunel(rstc, "tunel", True)) + Chr(10)
        
    End If
    If vplantilla = "SP temperature decks ºC  or units" Then vorigen = minimtinterotunel(rstc, "tinter") + CInt(Int((6 * Rnd()) + 2))
    If vplantilla = "SP temperature tunel ºC  or units" Then vorigen = minimtinterotunel(rstc, "tunel") + CInt(Int((6 * Rnd()) + 2))
    vplantilla = vplantilla + "     <------------------------- "
    Print #2, vorigen + Chr(9) + vplantilla
   Wend
   Close #1
   Close #2
   If atrim(passarfitxertemperaturesf2) <> "" Then passarfitxertemperaturesf2 = "Comanda:  " + atrim(numc) + Chr(10) + passarfitxertemperaturesf2 + Chr(13) + Chr(10) + vdetallcomanda
End Function
Function passarfitxertemperaturesfw(forigen As String, fdesti As String, numc As Double) As String
   Dim plantilla As String
   Dim vorigen As String
   Dim vplantilla As String
   Dim rstc As Recordset
   Dim vdetallcomanda As String
   Dim rectificar As Boolean
   plantilla = llegir_ini("General", "rutallistats", fitxerini) + "plantilla_temperatures_fw.txt"
   Set rstc = dbtmp.OpenRecordset("SELECT comandes.*, clients.nom FROM comandes INNER JOIN clients ON comandes.client = clients.codi where comanda = " + atrim(numc))
   If rstc.EOF Then Exit Function
   vdetallcomanda = "Nom: " + atrim(rstc!client) + " - " + atrim(rstc!nom) + Chr(13) + Chr(10) + "Ref Client: " + atrim(rstc!refclient) + Chr(13) + Chr(10) + "Texte Imp.: " + atrim(rstc!marcailinia)
   If existeix(fdesti) Then Kill fdesti
   Open forigen For Input As 1
   Open fdesti For Output As 2
   Open plantilla For Input As 3
   Print #2, "######## ##      ## "
   Print #2, "##       ##  ##  ## "
   Print #2, "##       ##  ##  ## "
   Print #2, "######   ##  ##  ## "
   Print #2, "##       ##  ##  ## "
   Print #2, "##       ##  ##  ## "
   Print #2, "##        ###  ###  "
   
   While Not EOF(1)
    rectificar = False
    vplantilla = ""
    vorigen = ""
    If Not EOF(3) Then Input #3, vplantilla
    Input #1, vorigen
    If atrim(vorigen) = "" Then vorigen = "0"
    If vplantilla = "Decks temperature" Then
        vorigen = convertirtemp(vorigen)
        If CDbl(vorigen) < minimtinterotunel(rstc, "tinter") Or CDbl(vorigen) > minimtinterotunel(rstc, "tinter", True) Then rectificar = True
        If rectificar Then passarfitxertemperaturesfw = passarfitxertemperaturesfw + "Error de temperatura del tinter. Valor: " + vorigen + "  Minim: " + atrim(minimtinterotunel(rstc, "tinter", False)) + "    Màxim:" + atrim(minimtinterotunel(rstc, "tinter", True)) + Chr(10)
        vplantilla = vplantilla + "     <------------------------- TINTERS"
    End If
    If vplantilla = "Tunnel temperature" Then
        vorigen = convertirtemp(vorigen)
        If CDbl(vorigen) < minimtinterotunel(rstc, "tunel") Or CDbl(vorigen) > minimtinterotunel(rstc, "tunel", True) Then rectificar = True
        If rectificar Then passarfitxertemperaturesfw = passarfitxertemperaturesfw + "Error de temperatura del tunel. Valor: " + vorigen + "  Minim: " + atrim(minimtinterotunel(rstc, "tunel", False)) + "    Màxim:" + atrim(minimtinterotunel(rstc, "tunel", True)) + Chr(10)
         vplantilla = vplantilla + "     <------------------------- TUNEL"
    End If
    If atrim(Mid(vplantilla, 1, 20)) = "Decks temperature" And rectificar Then vorigen = minimtinterotunel(rstc, "tinter") + CInt(Int((6 * Rnd()) + 2))
    If atrim(Mid(vplantilla, 1, 20)) = "Tunnel temperature" And rectificar Then vorigen = minimtinterotunel(rstc, "tunel") + CInt(Int((6 * Rnd()) + 2))
    
    Print #2, vorigen + Chr(9) + vplantilla
   Wend
   Close #1
   Close #2
   Close #3
   If atrim(passarfitxertemperaturesfw) <> "" Then passarfitxertemperaturesfw = "Comanda:  " + atrim(numc) + Chr(10) + passarfitxertemperaturesfw + Chr(10) + Chr(13) + Chr(10) + Chr(13) + vdetallcomanda
End Function
Function buscarcomandaacomexi(numc As Double) As String
  Dim r As String
  If Not existeix(nomordinadorcomexi) Then GoTo fi
  r = Dir(nomordinadorcomexi + "\" + "*.*")
  If r = "" Then
      Shell "c:\windows\system32\net.exe use " + rutaordinador(nomordinadorcomexi) + " /user:inplacsa ipc123"
      r = Dir(nomordinadorcomexi + "\" + "*.*")
  End If
  While r <> ""
    r = Dir
    If InStr(1, r, atrim(numc)) > 0 Then buscarcomandaacomexi = r
  Wend
fi:
End Function
Sub passartemperaturestemporalsadirectorifinal()
   Dim fitxerdesti As String
   Dim d As String
   If Not existeix(llegir_ini("ruta", "ruta_documentacio_temperatures", rutadelfitxer(cami) + "valorsprograma.ini")) Then Exit Sub
   
   d = Dir("c:\temp\temperatures\*.txt")
   While d <> ""
    fitxerdesti = llegir_ini("ruta", "ruta_documentacio_temperatures", rutadelfitxer(cami) + "valorsprograma.ini") + "\" + d
    If existeix(fitxerdesti) Then Kill fitxerdesti
    FileCopy "c:\temp\temperatures\" + d, fitxerdesti
    Kill "c:\temp\temperatures\" + d
    d = Dir
   Wend
End Sub

Sub calcularvalorsreducciocilindre(numc As Double, nummaq As Byte, numformula As Byte, llistat As Control)
   Dim rstc As Recordset
   Dim rstclixes As Recordset
   Dim dbclixes As Database
   Dim rstmodifi As Recordset
   Dim desarrollteoric As Double
   Dim desarrollreal As Double
   Dim valorrealmostra As Double
   Dim motius As Double
   
   If nummaq < 7 Then Exit Sub
   Set rstc = dbtmp.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(numc))
   If rstc.EOF Then Exit Sub
   id_treball = cadbl(rstc!numtreball)
   ordremodificacio = cadbl(rstc!numordremodificacio)
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
   Set rstclixes = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(cadbl(rstc!numtreball)))
   If cadbl(rstclixes!reduccioxmetre) = 0 Then Exit Sub
   If rstclixes.EOF Then Exit Sub
   Set rstmodifi = dbclixes.OpenRecordset("select desarroll,digimarc from modificacions where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordre=" + atrim(cadbl(rstc!numordremodificacio)))
   If rstmodifi.EOF Then Exit Sub
   
   If cadbl(rstmodifi!desarroll) = 0 Then Exit Sub
   motius = Redondejar(1000 / cadbl(rstmodifi!desarroll), 0)
   desarrollteoric = motius * cadbl(rstmodifi!desarroll)
   desarrollreal = Redondejar((desarrollteoric * cadbl(rstclixes!reduccioxmetre)) / 1000, 1)
   valorrealmostra = Redondejar(desarrollteoric + desarrollreal, 1)
   
   avisapantalla = "Distorsió cilindre per metre: " + atrim(cadbl(rstclixes!reduccioxmetre)) + " mm" + Chr(10) + "Factor: " + atrim(cadbl(IIf(nummaq = 7, rstclixes!redcilindrefw, rstclixes!redcilindref2))) + " mm"
   ' posso els valors al report llistat
   If numformula = 0 Then GoTo fi
   llistat.Formulas(numformula) = "reducciopermetrelineal=" + passaradecimalpunt(atrim(rstclixes!reduccioxmetre))
   numformula = numformula + 1
   llistat.Formulas(numformula) = "parametrereduccio=" + passaradecimalpunt(atrim((IIf(nummaq = 7, rstclixes!redcilindrefw, rstclixes!redcilindref2))))
   numformula = numformula + 1
   llistat.Formulas(numformula) = "desarrollteoric=" + passaradecimalpunt(atrim(desarrollteoric))
   numformula = numformula + 1
   llistat.Formulas(numformula) = "motius=" + passaradecimalpunt(atrim(motius))
   numformula = numformula + 1
   llistat.Formulas(numformula) = "desarrollreal=" + passaradecimalpunt(atrim(desarrollreal))
   numformula = numformula + 1
   llistat.Formulas(numformula) = "valorrealmostra=" + passaradecimalpunt(atrim(valorrealmostra))
   numformula = numformula + 1
fi:
   Set dbclixes = Nothing
   Set rstclixes = Nothing
   Set rstmodifi = Nothing
End Sub
Sub imprimir_packinglist(numc As Double)
   If numc < 100000 Then Exit Sub
   Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "palets.exe comandes.ini " + atrim(numc), vbNormalFocus
End Sub
Sub crear_taulatemp_bobinesdentrada()
  'Dim camps As String
  If nomfitxertemporalbobent <> "" Then Exit Sub
  nomfitxertemporalbobent = "c:\temp\~bibe" + Format(Now, "ddmmhhnnss") + ".mdb"
  On Error Resume Next
   MkDir "c:\temp"
   Kill "c:\temp\~bibe*.*"
   On Error GoTo 0
   DBEngine.CreateDatabase nomfitxertemporalbobent, dbLangGeneral, dbVersion10
   Set dbtemp = OpenDatabase(nomfitxertemporalbobent)
   'dbtemp.Execute "drop table tmp_imp_empalmes"
  On Error GoTo 0
  camps = "sel bit,idpalet double,idbobina double,metres double, utilitzada bit,tipus string(1),taula string,idb double"
  dbtemp.Execute ("create table selecciobobentrada (" + camps) + ")"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 ' If KeyCode = 110 Then SendKeys "," 'KeyCode = 188
  tempseditant = Now
End Sub

Sub obrirportsseriesdecomunicacio()
   'Dim combocomretorn As Byte
   'Dim combocomconsumides As Byte
   'Dim combocombalança As Byte
   'combocomretorn = cadbl(substituir(" " + llegir_ini("Baixes", "EscanerRetornCom", "comandes.ini"), "Com", ""))
   'combocomconsumides = cadbl(substituir(" " + llegir_ini("Baixes", "EscanerConsumidesCom", "comandes.ini"), "Com", ""))
   'combocombalança = cadbl(substituir((" " + llegir_ini("Baixes", "EscanerBalançaCom", "comandes.ini")), "Com", ""))
   'If combascula > 0 Then obrirportseriebascula combascula, combocombalança
End Sub

Sub mirarsihihalafontTTFdecodidebarres()
   Dim objShell As Variant
   Dim objFolderItem As Variant
   If existeix("c:\windows\fonts\free3of9.ttf") Then Exit Sub
   Copiar_Fitxer llegir_ini("General", "rutallistats", fitxerini) + "\free3of9.ttf", "c:\windows\fonts"
   Set objShell = CreateObject("Shell.Application")
   Set objFolder = objShell.Namespace("C:\windows\Fonts")
   Set objFolderItem = objFolder.ParseName("free3of9.ttf")
   objFolderItem.InvokeVerb ("Install")
End Sub
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
Sub carregar_nummaq()
  Dim r As String
  r = cadbl(llegir_ini("Baixes", "nummaq", "comandes.ini"))
  If r = "{[}]" Then r = 100
  nummaq = cadbl(r)
  If existeix("c:\ordprog.ini") Then nummaq = 7 'temporal per fer proves
End Sub

Private Sub Form_Load()
   Dim camistocks As String
  camicomandes = llegir_ini("General", "cami", "comandes.ini")
  cami = llegir_ini("General", "camibaixes", "comandes.ini")
  ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
  fitxerini = "comandes.ini"
  arguments = ObtenerLíneaComando
  Set dbmissatges = OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
  v = ShellExecute(0, "runas", "c:\windows\system32\net.exe", " time \\serverprodu /set /y", "", 0)
  If Not existeix("c:\temp") Then MkDir "c:\temp"
  carregar_nummaq
  'If UCase(arguments(1)) = "DESBOBINADORS" Then
  '   Set dbtmpb = OpenDatabase(cami)
  '   Set dbtmp = OpenDatabase(camicomandes)
  '   rellotge.Enabled = False: Timer1.Enabled = False: Formdesbobinadors.Show 1: End
  'End If
  If UCase(arguments(1)) = "DESBOBINADORS" Then
     arguments(1) = "DESBOBINADORS"
     'If Not existeix("c:\ordprog.ini") Then v = ShellExecute(0, "runas", "\\SERVERPRODU\Dades\progcomandes\aplicacio\dccmd.exe", " -width=640 -height=480  -refresh=60", "", 0)
     If nummaq = 0 Then nummaq = 9
     If cadbl(arguments(2)) > 0 Then nummaq = cadbl(arguments(2))
     Set dbtmpb = OpenDatabase(cami)
     Set dbtmp = OpenDatabase(camicomandes)
     Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
     rellotge.Enabled = False: Timer1.Enabled = False: formbobinesaimpresores.Show
      formbobinesaimpresores.Top = 0
       formbobinesaimpresores.Left = -105
     While isloaded("formbobinesaimpresores")
       DoEvents
     Wend
     End
  End If
  escriure_ini "Impresores_Compartida", "imprimirPKGLST_maq_" + atrim(nummaq), "0", rutadelfitxer(cami) + "valorsprograma.ini"
  If UCase(arguments(1)) = "ORDREIMPRESSIO" Then
     arguments(1) = "ORDREIMPRESSIO"
     If nummaq = 0 Then nummaq = 9
     If cadbl(arguments(2)) > 0 Then nummaq = cadbl(arguments(2))
     Set dbtmpb = OpenDatabase(cami)
     Set dbtmp = OpenDatabase(camicomandes)
     Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
     rellotge.Enabled = False: Timer1.Enabled = False: formordreimpresio.Show 1: End
  End If
  
  If UCase(arguments(1)) = "CONTROLCLIXESENTRATS" Then
     If nummaq = 0 Then nummaq = 9
     If cadbl(arguments(2)) > 0 Then nummaq = cadbl(arguments(1))
     Set dbtmpb = OpenDatabase(cami)
     Set dbtmp = OpenDatabase(camicomandes)
     Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
     rellotge.Enabled = False: Timer1.Enabled = False: formcontrolclixesentrats.Show 1: End
  End If

  

  If Not existeix("c:\ordprog.ini") And Not existeix("\\ord_copies\temperaturesimpressores\fitxerdecontrolnotocar.txt") Then MsgBox "Error no puc conectar a \\ord_copies\temperaturesimpressores", vbCritical, "Atenció"
  lletraseccio = "I"
 ' obrirportsseriesdecomunicacio
  mirarsihihalafontTTFdecodidebarres
  sonar_sirena "tancar"
  obrestocks True
  

  If llegir_ini("Baixes", "programaamaquina", fitxerini) = "{[}]" Then escriure_ini "Baixes", "programaamaquina", "0", fitxerini
  If cami = "{[}]" Then
    escriure_ini "General", "camibaixes", InputBox("Entra la ruta de baixes", "Atenció", "y:\comandes\baixes.mdb"), "comandes.ini"
  End If
  crear_taulatemp_bobinesdentrada
  If Not existeix("c:\temp\temperatures") Then
    If Not existeix("c:\temp") Then MkDir "c:\temp"
    If Not existeix("c:\temp\temperatures") Then MkDir "c:\temp\temperatures"
  End If
  
  comanda = cadbl(llegir_ini("Baixes", "ultimacomanda", "comandes.ini"))
  
  If LCase(App.EXEName) = "baixesimpresoramaquina2" Then form1.BackColor = &HFF80FF
  
  

  If Not existeix("c:\ordprog.ini") And nummaq > 0 Then assignardecimalipunt
  
  Load formannex

  Me.Top = 1
  Me.Left = 1
  formannex.Top = 80
  formannex.Left = 0
  formannex.Left = Me.width
  formannex.Show
  ''cami = "\\SERVERprodu\dades\progcomandes\dades\baixesprova.mdb"
  Set dbtmpb = OpenDatabase(cami)
  Set dbtmp = OpenDatabase(camicomandes)
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
  Set rsttmp = dbtmp.OpenRecordset("select codi,descripcio,nomordinadorcomexi from maquines where maquina='I' and codi=" + atrim(nummaq))
  
  If Not rsttmp.EOF Then
     nommaq = atrim(rsttmp!descripcio)
     nomordinadorcomexi = atrim(rsttmp!nomordinadorcomexi)
  End If


  If llegir_ini("Baixes", "programaamaquina", fitxerini) = "1" Then
   Shell ("Runas /user:administrador net time \\serverprodu /set /y")
   'desactiva ctrlaltsupr
   Shell "Runas /user:administrador c:\windows\regedit.exe /s \\serverprodu\dades\progcomandes\aplicacio\desactivarctrl.reg"
   Shell "Runas /user:administrador c:\windows\system32\net.exe use " + rutaordinador(nomordinadorcomexi) + " /user:inplacsa ipc123"
  End If
  impresores.DatabaseName = cami
  imppantones.DatabaseName = cami
  bobines.DatabaseName = cami
  empalmes.DatabaseName = cami
  bobinesent.DatabaseName = cami
  
  lots.DatabaseName = rutadelfitxer(cami) + "tintes.mdb"
  lots.RecordSource = "SELECT Componentsbase.nomcomponent AS nomlot, detallnumeroslotsbase.numerodelot AS codilot FROM detallnumeroslotsbase INNER JOIN Componentsbase ON detallnumeroslotsbase.idcomponent = Componentsbase.idcomponent "
  lots.RecordSource = lots.RecordSource + " WHERE (((Componentsbase.numdosificador)>=100) AND ((detallnumeroslotsbase.data) In (SELECT  Max(detallnumeroslotsbase.data) AS ladata From detallnumeroslotsbase GROUP BY detallnumeroslotsbase.idcomponent;))) ORDER BY Componentsbase.nomcomponent;"
  lots.Refresh
  dblots.Left = 45
  dblots.Top = 180
  
  
  Set dbtmpb = OpenDatabase(impresores.DatabaseName)
  rellotge.Enabled = True
  rellotge.Interval = 900

  'nummaq = 9
  If nummaq > 90 Then
    i = nummaq
    nummaq = cadbl(InputBox("Entra el numero de maquina", "Atenció", "7"))
    If nummaq <> 7 And nummaq <> 5 And nummaq <> 2 Then MsgBox "Nomes hi ha la 2 la 5 i la 7": End
    If i = 0 Then escriure_ini "Baixes", "nummaq", atrim(nummaq), "comandes.ini"
  End If
  
  If nummaq = 0 Then
    maquina.visible = True
    imprimir.visible = True
   Else: maquina.visible = False: imprimir.visible = False
  End If
  impresores.RecordSource = "select * from impressores where comanda=-1"
  impresores.Refresh
  'r = ""
  'For i = 1 To 10
  '   r = r + ",pantone" + atrim(i) + " text,lot" + atrim(i) + " integer,kg" + atrim(i) + " double "
  'Next i
'  dbtmpb.Execute ("create table impresorespantones (comanda integer " + r + ")")
'For i = 1 To 10
'  dbtmpb.Execute ("alter table impresorespantones drop column lot" + atrim(i))
'Next i
'  wait 1
'  For i = 1 To 10
'  dbtmpb.Execute ("alter table impresorespantones add column lot" + atrim(i) + " text")
'Next i

  On Error Resume Next
  dbtmpb.Execute ("create table lots (nomlot string,codilot string)")
  On Error GoTo 0

  For Each objecte In Me
      If objecte.Name <> "reciclarmaterial1" And objecte.Name <> "nomoperari" And objecte.Name <> "Line1" And objecte.Name <> "rellotge" And objecte.Name <> "llistat" And objecte.Name <> "llistatbob" Then
        objecte.Enabled = False
      End If
     Next objecte

possarestadisticadeldia
borrar_verificacioescaner_nomesunmes

End Sub
Sub borrar_verificacioescaner_nomesunmes()
  Dim vinstruccio As String
  Dim vruta As String
  Dim vfitxer As String
  vruta = "\\serverprodu\Dades\progcomandes\dades\Lectures_Codisdebarres_Impresores\" + atrim(nummaq)
  If existeix(vruta) Then
      vfitxer = Dir(vruta + "\*.scn")
      While vfitxer <> ""
         vfitxer = vruta + "\" + vfitxer
         If DateDiff("d", FileDateTime(vfitxer), Now) > 30 Then Kill vfitxer
         vfitxer = Dir
      Wend
  End If
End Sub
Function rutaordinador(nomord As String) As String
     If atrim(nomord) <> "" Then
       rutaordinador = Mid(nomord, 1, InStr(3, nomord, "\") - 1)
     End If
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ' If Button = 2 And Shift = 1 Then imprimir_full_arrancar_rentar cadbl(comanda)
  If Shift = 1 Then imprimir_controlqualitat cadbl(comanda), numop, 0
  
  If Shift = 2 Then
    If InputBoxEx("Entra la contrasenya de configuració, la de sempre però de 4", "Programador", , , , , , SPassword) = "9909" Then
     If MsgBox("Prem si per activar el bloqueig del Ctrl+Alt+Supr i no per desactivar-lo", vbYesNo, "Atenció") = vbYes Then
       Shell "c:\windows\regedit.exe /s \\serverprodu\dades\progcomandes\aplicacio\desactivarctrl.reg"
        Else: Shell "c:\windows\regedit.exe /s \\serverprodu\dades\progcomandes\aplicacio\activarctrl.reg"
     End If
    End If
  End If
End Sub

Private Sub Form_Resize()
'  Me.caption = atrim(Me.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
 On Error Resume Next
' combascula.PortOpen = False
' comescanerbascula.PortOpen = False
' comescanerconsum.PortOpen = False
 dbtemp.Close
 dbtmp.Close
 dbtmpb.Close
 dbstocks.Close
 
 End
End Sub
Sub tancar_formularis()
Dim frm As Form
On Local Error Resume Next
For Each frm In Forms
   Unload frm
   frm.Hide
   Set frm = Nothing
Next
Set rsttmp = Nothing
    Set dbtmp = Nothing
    Set dbtmpb = Nothing
    Set bdllistat = Nothing
End Sub

Sub llistatpackinglist(numcomanda As Double)
   numcomanda = numcomanda * -1  'si el passo en negatiu imprimeix per pantalla
   Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "palets.exe comandes.ini " + atrim(numcomanda), vbNormalFocus
End Sub

Private Sub ImagePDF_DblClick()

End Sub

Private Sub impresores_Reposition()
 If Not impresores.Recordset.EOF Then
      ensenya_les_bobines
 End If


End Sub
Sub ensenya_les_bobines()
  Dim bk As String
  If Me.Name = "reixabobines" Then Exit Sub
  r = "-1"
  If impresores.Recordset!tipus = "F" Then r = atrim(cadbl(impresores.Recordset!id))
  If Not bobines.Recordset.EOF Then bk = bobines.Recordset!numerodebobina
  bobines.RecordSource = "select * from bobinesimp where controlid=" + r + " order by numerodebobina"
  bobines.Refresh
  
   bobines.Recordset.LockEdits = False
 bobinesent.Recordset.LockEdits = False
  
  On Error Resume Next
  If bk <> "" Then
     bobines.Recordset.FindFirst "numerodebobina=" + bk
   Else: bobines.Recordset.MoveLast
  End If
  'If Not IsEmpty(bk) Then bobines.Recordset.Bookmark = bk
  
End Sub

Private Sub imprimir_Click()
guardar_totals_packinglist cadbl(comanda)
'If stockopacking = "E" Then imprimir_packinglist cadbl(comanda)
imprimir_fulla
End Sub

Private Sub kbpantone_LostFocus(Index As Integer)
 Dim totaltinta As Double
 imppantones.Refresh
 
 totaltinta = 0
 For i = 0 To 11
   totaltinta = totaltinta + cadbl(kbpantone(i))
 Next i
 'impresores.Recordset.Edit3
 dbtmpb.Execute "update impressores set kgtinta=" + passaradecimalpunt(atrim(totaltinta)) + " where id=" + atrim(impresores.Recordset!id)
 'impresores.Recordset!kgtinta = totaltinta
 'impresores.Recordset.Update
 
 impresores.Recordset.Move 0
 
 impresores.UpdateControls
 
 If totaltinta > 0 And framepantones.tag = "E" Then dbtmp.Execute "update comandes set proximaseccio='I' where comanda=" + atrim(cadbl(comanda.text)): framepantones.tag = "I"
End Sub

Private Sub maquina_Click()
nummaq = cadbl(InputBox("Entra el numero de màquina [5,7 o 9]", "Atenció"))
   If nummaq <> 5 And nummaq <> 7 And nummaq <> 9 Then nummaq = 0
   maquina.caption = "Maq: " + atrim(nummaq)
   maquina.tag = nummaq
End Sub

Private Sub mm_Click()

End Sub

Private Sub nomoperari_Click()
 Dim numoptmp As Integer
 Dim numoptmp2 As Integer
 Dim nomoptmp As String
 Dim nomoptmp2 As String
 
 If barraestat.caption = "Calculant els totals..." Then Exit Sub
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "select codi,descripcio from operaris where maquina='I' and actiu<>0 order by codi "
  formseleccio.caption = "Selecció d'Operari"
  formseleccio.refrescar
  formseleccio.Height = form1.Height
   formseleccio.Top = 0
   formseleccio.DBGrid2.Font.Size = 16
   formseleccio.DBGrid2.RowHeight = 440
   formseleccio.DBGrid2.Height = form1.Height - formseleccio.DBGrid2.Top - 600
  formseleccio.Show 1
  If seleccioret = 1 Then
   numoptmp = cadbl(formseleccio.Data1.Recordset!codi)
   nomoptmp = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
  If numoptmp <> 0 Then
   Load formseleccio
   formseleccio.Data1.DatabaseName = camicomandes
   formseleccio.Data1.RecordSource = "select codi,descripcio from operaris where maquina='I' and actiu<>0 order by codi "
   formseleccio.caption = "Selecció AJUDANT d'OPERARI"
   formseleccio.BackColor = QBColor(12)
   formseleccio.refrescar
   formseleccio.Height = form1.Height
   formseleccio.Top = 0
   formseleccio.DBGrid2.Font.Size = 16
   formseleccio.DBGrid2.RowHeight = 440
   formseleccio.DBGrid2.Height = form1.Height - formseleccio.DBGrid2.Top - 600
   formseleccio.Show 1
   If seleccioret = 1 Then
    numoptmp2 = cadbl(formseleccio.Data1.Recordset!codi)
    nomoptmp2 = atrim(formseleccio.Data1.Recordset!descripcio)
   End If
   Unload formseleccio
  End If
  If numoptmp <> 0 And numoptmp2 <> 0 Then
     nomoperari = nomoptmp
     nomoperari2 = "Ajudant: " + nomoptmp2
     numop = numoptmp
     numop2 = numoptmp2
     For Each objecte In Me
      If objecte.Name <> "reciclarmaterial1" And objecte.Name <> "comanda" And objecte.Name <> "llistat" And objecte.Name <> "llistatbob" And objecte.Name <> "Line1" And objecte.Name <> "comandaacavada" Then
        objecte.Enabled = True
      End If
     Next objecte
      Else: If cadbl(numop) = 0 Then MsgBox "Has d'escullir un operari per treballar": Exit Sub
  End If
   If cadbl(comanda) > 0 Then
      Command4_Click
     Else: If comanda.Enabled Then comanda.SetFocus
   End If
 comprovarsihihamissatgesCHAT
 actualitzar_comandaactiva_compartida
End Sub

Private Sub pantone_LostFocus(Index As Integer)
 imppantones.Refresh
End Sub

Private Sub reixa_AfterUpdate()
  'calcular_totals
End Sub

Private Sub reixa_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
tempseditant = Now
End Sub

Private Sub reixa_BeforeDelete(Cancel As Integer)
  If Screen.ActiveControl.Name <> "Command14" Then
   If MsgBox("Segur que vols borrar aquesta linia i tot el seu contingut?", vbYesNo, "Atenció") = vbNo Then Cancel = 1
  End If
  If Cancel <> 1 Then
    'If impresores.Recordset!tipus = "F" Then
    r = atrim(cadbl(impresores.Recordset!id))
    dbtmpb.Execute "delete * from bobinesimp where controlid=" + r
  End If
End Sub

Private Sub reixa_Click()
ratoli "normal"
End Sub

Private Sub reixa_ColEdit(ByVal ColIndex As Integer)
tempseditant = Now
End Sub

Private Sub reixa_DblClick()
If reixa.col = 15 Then
   r = triar_observacio(impresores.Recordset!tipus)
   If Len(r) > 4 Then
     r = Mid(r, 4, Len(r))
     If r <> "" Then
       If reixa.text <> "" Then
           reixa.text = reixa.text + " <> " + r
           Else: reixa.text = r
       End If
     End If
   End If
   If r = "PARAR MÀQUINA" Then feines_parar_engegar_maquina "PARAR"
   If r = "ENGEGAR MÀQUINA" Then feines_parar_engegar_maquina "ENGEGAR"
End If

If reixa.col = 14 Then
  'If atrim(reixa.Text) = "" Then
  '   reixa.Text = demanar_bobinaprova(reixa.Text)
  '   Else
  '      If cadbl(impresores.Recordset!paletprova2) > 0 Then MsgBox "Nº del segon palet i metres:  " + atrim(impresores.Recordset!paletprova2) + "-" + atrim(impresores.Recordset!bobinaprova2) + "/" + atrim(impresores.Recordset!metresprova2) + " mtrs"
  '      If MsgBox("Vols introduir un segon palet? si es que no modificaras el primer.", vbYesNo, "Atenció") = vbYes Then
  '           If cadbl(impresores.Recordset!paletprova2) > 0 Then
  '              If MsgBox("Ja hi ha un segon palet entrat vols borrar-lo?", vbYesNo, "Atenció") = vbYes Then
  '                 impresores.Recordset.Edit
  '                 impresores.Recordset!paletprova2 = 0: impresores.Recordset!bobinaprova2 = 0: impresores.Recordset!metresprova2 = 0
  '                 impresores.Recordset.Update
  '                Else
  '                  demanar_bobinaprova , , True
  '              End If
  '               Else: demanar_bobinaprova , , True
  '            End If
  '          Else
  '            reixa.Text = demanar_bobinaprova(reixa.Text)
  '      End If
        
  'End If
  paletsajust.Show 1
  'fpaletsajust.Visible = True
  'fpaletsajust.Left = 6500
  'fpaletsajust.Top = 1500
  'fpaletsajust.Visible = True
  'impresores.Recordset.Edit
  'r = ""
  'If cadbl(impresores.Recordset!paletprova2) > 0 Then r = "*"
  'impresores.Recordset!paletbobprova = r + atrim(impresores.Recordset!paletprova) + "-" + atrim(impresores.Recordset!bobinaprova)
  'reixa.Text = impresores.Recordset!paletbobprova
  'impresores.Recordset.Update
End If

If reixa.col = 0 Then
  reixa.text = escullir_operari
  nomoperari = UCase(r)
  numop = reixa.text
  numop2 = escullir_operari("Escullir AJUDANT d'OPERARI")
  nomoperari2 = "Ajudant: " + UCase(r)
End If
End Sub
Function demanar_bobinaprova(Optional paletbobv As Variant, Optional nogravar As Boolean, Optional segon As Boolean) As String
  Dim palet As String
  Dim bob As String
  If atrim(paletbobv) = "" Then paletbobv = ""
  demanar_bobinaprova = ""
  Do
    palet = InputBox("Entra el numero de palet.", "Atenció")
  Loop Until cadbl(palet) <> 0 Or palet = ""
  If palet <> "" Then
     Do
      bob = InputBox("Entra el numero de bobina.", "Atenció")
     Loop Until cadbl(bob) <> 0 Or bob = ""
     If bob <> "" Then
       demanar_bobinaprova = atrim(palet) + "-" + atrim(bob)
       If Not nogravar Then impresores.Recordset.Edit
       If Not segon Then
          impresores.Recordset!paletprova = palet
          impresores.Recordset!bobinaprova = bob
           Else
             impresores.Recordset!paletprova2 = palet
             impresores.Recordset!bobinaprova2 = bob
             impresores.Recordset!metresprova2 = cadbl(InputBox("Entra els metres que has utilitzat de la bobina."))
       End If
       If Not nogravar Then impresores.Recordset.Update
     End If
  End If
  If demanar_bobinaprova = "" Then demanar_bobinaprova = paletbobv
End Function
Function triar_observacio(tipus As String) As String
  'Dim rsttriar As Recordset
  'Set rsttriar = dbtmp.OpenRecordset("select * from constantsobservacio where mid(observacio,1,1)='" + tipus + "'")
  'While Not rsttriar.EOF
  '  rsttriar.MoveNext
  'Wend
  
  Load formseleccio
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select * from constantsobservacio where mid(observacio,1,2)='I" + tipus + "'"
  formseleccio.caption = "Triar Observació"
  formseleccio.Height = 9100
  formseleccio.DBGrid2.Height = 8000
  formseleccio.refrescar
  formseleccio.Top = 2800
  formseleccio.Show 1
  If seleccioret = 1 Then
    triar_observacio = atrim(formseleccio.Data1.Recordset!observacio)
   'If InStr(1, nomoperari.Caption, "MARTINEZ") Then
   '    Command12.Visible = True
   '   Else: Command12.Visible = False
   'End If
  End If
  Unload formseleccio
  
End Function

Private Sub reixa_GotFocus()
   cpostitcomanda.visible = True
   cpostitcomanda.Top = reixa.RowTop(reixa.row) + reixa.Top
     cpostitcomanda.Left = reixa.Columns(4).Left + reixa.Left
End Sub

Private Sub reixa_KeyDown(KeyCode As Integer, Shift As Integer)
tempseditant = Now
End Sub

Private Sub reixa_KeyUp(KeyCode As Integer, Shift As Integer)
 If reixa.col = 4 And KeyCode > 46 Then
     If (Len(reixa.text)) >= 4 Then reixa.col = 5
  End If
  If reixa.col = 3 And KeyCode > 46 Then
     If (Len(reixa.text)) >= 6 Then reixa.col = 4
  End If
  If reixa.col = 2 And KeyCode > 46 Then
     If (Len(reixa.text)) >= 4 Then
       reixa.col = 3
     End If
  End If
  If reixa.col = 1 And KeyCode > 46 Then
     If (Len(reixa.text)) >= 6 Then reixa.col = 2
  End If
  If reixa.col = 15 And KeyCode > 46 Then
      If (Len(reixa.text)) > 99 Then reixa.text = Mid(reixa.text, 1, 99)
  End If
  If reixa.col = 14 And KeyCode = 46 Then
     If MsgBox("Borraras els palets apuntats en aquest ajust.", vbYesNo, "Atenció") = vbYes Then
         impresores.Recordset.Edit
         impresores.Recordset!paletprova = 0
         impresores.Recordset!bobinaprova = 0
         impresores.Recordset!paletprova2 = 0
         impresores.Recordset!bobinaprova2 = 0
         impresores.Recordset!metresprova2 = 0
         impresores.Recordset!paletbobprova = ""
         reixa.text = ""
         impresores.Recordset.Update
     End If
  End If
End Sub

Sub bloquejar_camps_innecesaris()
If impresores.Recordset.EOF Then Exit Sub
For i = 0 To 15
  reixa.Columns(i).Locked = False
Next i
reixa.Columns(5).Locked = True
reixa.Columns(6).Locked = True
reixa.Columns(7).Locked = True
reixa.Columns(8).Locked = True
reixa.Columns(9).Locked = True
reixa.Columns(10).Locked = True
reixa.Columns(11).Locked = True
reixa.Columns(12).Locked = True
reixa.Columns(13).Locked = True
reixa.Columns(14).Locked = True
'If impresores.Recordset!tipus = "A" Then reixa.Columns(7).Locked = False ': reixa.Columns(14).Locked = False
If impresores.Recordset!tipus = "F" Then reixa.Columns(11).Locked = False




End Sub
Function buscarnomoperari(vop As Long) As String
  Dim rst As Recordset
  
  If vop = 0 Then Exit Function
  Set rst = dbtmp.OpenRecordset("select * from operaris where codi=" + atrim(vop) + " and maquina='I'")
  If Not rst.EOF Then buscarnomoperari = atrim(rst!descripcio)
  Set rst = Nothing
End Function
Private Sub reixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 Dim valtmp As String
 Dim rst As Recordset
 cpostitcomanda.visible = False
 If reixa.row >= 0 Then
   If impresores.Recordset!tipus = "V" Then
    cpostitcomanda.text = ""
    cpostitcomanda.text = atrim(impresores.Recordset!tipificacioavaria)
    cpostitcomanda.visible = True
    cpostitcomanda.Top = reixa.RowTop(reixa.row) + reixa.Top + reixa.RowHeight
    cpostitcomanda.Left = reixa.Columns(4).Left + reixa.Left
   End If
 End If
 If reixa.col = 0 Then
   'If impresores.Recordset!tipus = "V" Then
    cpostitcomanda.text = ""
    cpostitcomanda.text = atrim(buscarnomoperari(cadbl(impresores.Recordset!operari2)))
    If atrim(cpostitcomanda.text) <> "" Then
        cpostitcomanda.text = "Ajudant: " + cpostitcomanda.text
        cpostitcomanda.visible = True
        cpostitcomanda.Top = reixa.RowTop(reixa.row) + reixa.Top + reixa.RowHeight
        cpostitcomanda.Left = reixa.Columns(0).Left + reixa.Left
   End If
   'End If
 End If
 
 
 bobsajust.visible = False
    If reixa.col = 0 Then reixa.EditActive = False
 '-------
 bloquejar_camps_innecesaris
 If Not impresores.Recordset.EOF Then
 'texteimpresio = atrim(impresores.Recordset!texteimpresio)
  If atrim(impresores.Recordset!tipus) = "F" Then
     framebobines.Enabled = True
       Else: framebobines.Enabled = False: framepantones.visible = False
  End If
 End If
 If Not impresores.Recordset.EOF Then
  If atrim(impresores.Recordset!tipus) = "A" Then
     bobsajust.visible = True
  End If
 End If
 
  If LastCol = 1 Or LastCol = 2 Then
  valtmp = reixa.Columns(LastCol).text
  
  If LastCol = 1 Then
      
      If InStr(1, valtmp, "/") = 0 Then valtmp = Mid(valtmp, 1, 2) + "/" + Mid(valtmp, 3, 2) + "/" + Mid(valtmp, 5, 2)
      If Not IsDate(valtmp) Then valtmp = ""
  End If
  
  If LastCol = 2 Then
    If InStr(1, valtmp, ":") = 0 Then valtmp = Mid(valtmp, 1, 2) + ":" + Mid(valtmp, 3, 2)
    If Not IsDate(Format(valtmp, "hh:nn")) Then valtmp = ""
  End If
  reixa.Columns(LastCol) = IIf(valtmp = "", Null, valtmp)
  End If
  
  If LastCol = 3 Or LastCol = 4 Then
  valtmp = reixa.Columns(LastCol).text
  If LastCol = 3 Then
      
      If InStr(1, valtmp, "/") = 0 Then valtmp = Mid(valtmp, 1, 2) + "/" + Mid(valtmp, 3, 2) + "/" + Mid(valtmp, 5, 2)
      If Not IsDate(valtmp) Then valtmp = ""
  End If
  If LastCol = 4 Then
    If InStr(1, valtmp, ":") = 0 Then valtmp = Mid(valtmp, 1, 2) + ":" + Mid(valtmp, 3, 2)
      If Not IsDate(Format(valtmp, "hh:nn")) Then valtmp = ""

  End If
  reixa.Columns(LastCol) = IIf(valtmp = "", Null, valtmp)
  End If
 'calcular_totals
 
 
 
End Sub

Private Sub reixabobines_AfterColUpdate(ByVal ColIndex As Integer)
If reixabobines.Columns(ColIndex) = "" And ColIndex < 8 Then reixabobines.Columns(ColIndex) = "0"
If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
 bobines.Recordset.Fields(reixabobines.Columns(ColIndex).DataField) = cadbl(reixabobines.Columns(ColIndex).text)
 reixabobines.EditActive = False
'bobines.Recordset.Update
End Sub

Private Sub reixabobines_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
tempseditant = Now
End Sub

Private Sub reixabobines_Click()
  ratoli "normal"
End Sub

Private Sub reixabobines_ColEdit(ByVal ColIndex As Integer)
  tempseditant = Now
End Sub

Private Sub reixabobines_DblClick()
If reixabobines.col = 8 Then
   r = triar_observacio("B")
   If r <> "" Then reixabobines.text = r
End If
If reixabobines.col = 1 Or reixabobines.col = 0 Then
  reixabobines.text = escullir_operari
  If reixabobines.col = 0 Then
   nomoperari = UCase(r)
   numop = cadbl(reixabobines.text)
   numop2 = escullir_operari("Escullir AJUDANT d'OPERARI")
   nomoperari2 = "Ajudant: " + UCase(r)
  End If
  
End If
End Sub
Function escullir_operari(Optional vtitolfinestre As String) As String
  Dim opvell As Byte
  If vtitolfinestre = "" Then vtitolfinestre = "Selecció d'Operari"
  opvell = numop
  r = nomoperari
 'While cadbl(escullir_operari) = 0
   Load formseleccio
   formseleccio.Data1.DatabaseName = camicomandes
   formseleccio.Data1.RecordSource = "select codi,descripcio from operaris where maquina='I' and actiu<>0"
   formseleccio.caption = vtitolfinestre
   If InStr(1, vtitolfinestre, "AJUDANT") > 0 Then formseleccio.BackColor = QBColor(12)
   formseleccio.refrescar
   formseleccio.Show 1
   If seleccioret = 1 Then
    escullir_operari = cadbl(formseleccio.Data1.Recordset!codi)
    r = formseleccio.Data1.Recordset!descripcio
   End If
   If cadbl(escullir_operari) = 0 Then MsgBox "Has d'escullir un operari per treballar"
 'Wend
 If cadbl(escullir_operari) = 0 Then escullir_operari = opvell
 Unload formseleccio
End Function

Private Sub reixabobines_Error(ByVal DataError As Integer, Response As Integer)
If reixabobines.Columns(3) = "" Then reixabobines.Columns(3) = "0"
If reixabobines.Columns(4) = "" Then reixabobines.Columns(4) = "0"
Response = 0
End Sub

Private Sub reixabobines_KeyDown(KeyCode As Integer, Shift As Integer)
tempseditant = Now
End Sub

Private Sub reixabobines_LostFocus()

If reixabobines.col > 1 And Screen.ActiveControl.Name <> "Command7" And Screen.ActiveControl.Name <> "DBGrid2" Then calcular_totals
End Sub

Private Sub reixabobines_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Static fila As Double
If IsNull(fila) Then fila = 0
If fila <> reixabobines.row Then
 'calcular_totals
End If
fila = reixabobines.row
End Sub

Private Sub reixaempalmes_AfterDelete()
If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
  bobines.Recordset!numempalmes = empalmes.Recordset.RecordCount
  bobines.Recordset.Update
End Sub

Private Sub reixaempalmes_AfterUpdate()
  If bobines.Recordset.EditMode = 0 And Not bobines.Recordset.EOF Then
    bobines.Recordset.Edit
    bobines.Recordset!numempalmes = empalmes.Recordset.RecordCount
    bobines.Recordset.Update
  End If
End Sub

Private Sub reixaempalmes_ColEdit(ByVal ColIndex As Integer)
  tempseditant = Now
End Sub

Private Sub reixaempalmes_DblClick()
If reixaempalmes.col = 1 Then
   r = triar_observacio("S")
   If r <> "" Then reixaempalmes.text = r
End If
End Sub

Private Sub reixaempalmes_OnAddNew()
 empalmes.Recordset!id = bobines.Recordset!id
 reixa.col = 0
 
End Sub

Sub mirarsiparar()
 Static contar
  If llegir_ini("General", "parar", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini") = "si" Then
    contar = contar + 1
     If contar = 1 Then MsgBox2 "El programa es pararà d'aqui a 1 minut. TANCA TOT I ESPERA CINC MINUTS.", 5, "Actualització", vbCritical
     If contar = 15 Then MsgBox2 "El programa es pararà d'aqui a 30 segons. TANCA TOT I ESPERA CINC MINUTS.", 5, "Actualització, vbCritical"
     If contar = 27 Then MsgBox2 "El programa es pararà d'aqui a 5 segons. TANCA TOT I ESPERA CINC MINUTS.", 3, "Actualització", vbCritical
     If contar > 30 Then End
   Else: contar = 0
  End If
  If llegir_ini("General", "parar", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini") = "ja" Then End
End Sub

Sub comprovareditantbobines()
If DateDiff("s", tempseditant, Now) > 6 And tempseditant > 0 Then
   On Error Resume Next
   If Not bobines.Recordset.EOF Then
    If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
    bobines.Recordset.Update
   End If
   bobines.UpdateControls
   If Not impresores.Recordset.EOF Then
    If impresores.Recordset.EditMode = 0 Then impresores.Recordset.Edit
    
    impresores.Recordset.Update
   End If
   impresores.UpdateControls
   
   If Not empalmes.Recordset.EOF Then
    If empalmes.Recordset.EditMode = 0 Then empalmes.Recordset.Edit
    empalmes.Recordset.Update
   End If
   empalmes.UpdateControls
   
   tempseditant = 0
 End If
End Sub
Sub pampalluguesbotoavisos()
  'If botoavisos.tag = "avis" Then
  '    botoavisos.BackColor = IIf(botoavisos.BackColor = QBColor(12), Command4.BackColor, QBColor(12))
  '      Else: botoavisos.BackColor = Command4.BackColor
  'End If
End Sub
Sub actualitzar_comandaactiva_compartida()
 Dim vnumc As Double
 Dim vnumpalet As Double
 Dim vbob As Double
 Dim vTipuspitu As String
 
 If Not existeix("c:\ordprog.ini") And nummaq > 0 Then
        Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
        escriure_ini "Impresores_Compartida", "dataihora_maq_" + atrim(nummaq), Now, rutadelfitxer(cami) + "valorsprograma.ini"
        escriure_ini "Impresores_Compartida", "comanda_maq_" + atrim(nummaq), atrim(comanda), rutadelfitxer(cami) + "valorsprograma.ini"
        escriure_ini "Impresores_Compartida", "numop1_maq_" + atrim(nummaq), atrim(numop), rutadelfitxer(cami) + "valorsprograma.ini"
        escriure_ini "Impresores_Compartida", "numop2_maq_" + atrim(nummaq), atrim(numop2), rutadelfitxer(cami) + "valorsprograma.ini"
        escriure_ini "Impresores_Compartida", "nomoperari_maq_" + atrim(nummaq), atrim(nomoperari), rutadelfitxer(cami) + "valorsprograma.ini"
        escriure_ini "Impresores_Compartida", "nomoperari2_maq_" + atrim(nummaq), atrim(nomoperari2), rutadelfitxer(cami) + "valorsprograma.ini"
        vnumpalet = cadbl(llegir_ini("Impresores_Compartida", "imprimirPalet_maq_" + atrim(nummaq), rutadelfitxer(cami) + "valorsprograma.ini"))
        vbob = cadbl(llegir_ini("Impresores_Compartida", "imprimirBobina_maq_" + atrim(nummaq), rutadelfitxer(cami) + "valorsprograma.ini"))
        vnumc = cadbl(llegir_ini("Impresores_Compartida", "imprimirPKGLST_maq_" + atrim(nummaq), rutadelfitxer(cami) + "valorsprograma.ini"))
        If vnumc > 0 Then
           escriure_ini "Impresores_Compartida", "imprimirPKGLST_maq_" + atrim(nummaq), 0, rutadelfitxer(cami) + "valorsprograma.ini"
           form1.imprimir_packinglistTICKET vnumc, False
        End If
        vTipuspitu = llegir_ini("Impresores_Compartida", "SonarSirena_maq_" + atrim(nummaq), rutadelfitxer(cami) + "valorsprograma.ini")
        If vTipuspitu <> "" And vTipuspitu <> "{[}]" Then
             escriure_ini "Impresores_Compartida", "SonarSirena_maq_" + atrim(nummaq), " ", rutadelfitxer(cami) + "valorsprograma.ini"
             sonar_sirena vTipuspitu
        End If
        If vnumpalet > 0 Then
              bobinesdentrada.imprimir_bobinaparcial cadbl(vnumpalet), cadbl(vbob), , 1
              escriure_ini "Impresores_Compartida", "imprimirBobina_maq_" + atrim(nummaq), "0", rutadelfitxer(cami) + "valorsprograma.ini"
              escriure_ini "Impresores_Compartida", "imprimirPalet_maq_" + atrim(nummaq), "0", rutadelfitxer(cami) + "valorsprograma.ini"
        End If
 End If
 
End Sub
Private Sub rellotge_Timer()
  Static tempsoperari As Byte
  Static controlminut As Date
'  Static ultimarow As Double
'  If ultimarow = 0 Then ultimarow = reixa.Row
'  If ultimarow <> reixa.Row Then
'     ultimarow = reixa.Row: calcular_totals
 ' End If
' Timer1.Enabled = True
 pampalluguesbotoavisos
 If DateDiff("n", controlminut, Now) > 0 Then
   controlminut = Now
   If Not existeix("c:\ordprog.ini") Then assignardecimalipunt
 End If
 mirarsiparar
 comprovareditantbobines
 
 On Error Resume Next
 If rststocks.EOF Then
    controlstock.caption = ""
      Else: controlstock.caption = "Stocks"
 End If
 On Error GoTo error_screen
 
 If Screen.ActiveControl.Name = "akjdfks" Then Me.caption = Me.caption
 On Error GoTo 0
 If client.caption = "" And (impresores.Recordset.BOF And impresores.Recordset.EOF) Then
   carregar_client_ntintersialtres
 End If
 
 If numop = 0 And Not formseleccio.visible And reixa.Enabled Then
   numop = escullir_operari
   nomoperari = UCase(r)
   numop2 = escullir_operari("Escullir AJUDANT d'OPERARI")
   nomoperari2 = "Ajudant: " + UCase(r)
 End If

 
 If reixa.col = 0 And Screen.ActiveControl.Name = "reixa" Then
   tempsoperari = cadbl(tempsoperari) + 1
   If tempsoperari > 2 Then reixa.col = 1: tempsoperari = 0
 End If
 If (reixabobines.col = 0 Or reixabobines.col = 1) And Screen.ActiveControl.Name = "reixabobines" Then
   tempsoperari = cadbl(tempsoperari) + 1
   If tempsoperari > 2 Then reixabobines.col = 2: tempsoperari = 0
 End If
 On Error GoTo cont
 If Screen.ActiveForm.Name = Me.Name Then
   Set campcontrol = ActiveControl
 End If
 
   'If Screen.ActiveControl.Tag = "888" And Not teclattactil.Visible Then
   '   r = "numeric"
   '   teclattactil.Show
   '     Else: Unload teclattactil
   'End If
cont:

 On Error GoTo 0
  If InStr(1, hora, ":") > 0 Then
     hora = Format(Now, "hh nn")
     Else: hora = Format(Now, "hh:nn")
  End If
  rellotge.tag = cadbl(rellotge.tag) + 1
  If rellotge.tag = "10" Then
    'calcular_totals
    rellotge.tag = "0"
  End If
  actualitzar_comandaactiva_compartida
  If Not impresores.Recordset.EOF Then
    Select Case atrim(impresores.Recordset!tipus)
       Case "A"
          Command1.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor: Command27.BackColor = Command4.BackColor
          Command2.BackColor = &HFF8080
      Case "V"
          Command1.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor: Command2.BackColor = Command4.BackColor
          Command27.BackColor = &HFF8080
          
       Case "M"
          Command2.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor: Command27.BackColor = Command4.BackColor
          Command1.BackColor = &HFF8080
       Case "F"
          Command1.BackColor = Command4.BackColor: Command2.BackColor = Command4.BackColor: Command27.BackColor = Command4.BackColor
          Command3.BackColor = &HFF8080
        Case Else
           Command1.BackColor = Command4.BackColor: Command2.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor: Command27.BackColor = Command4.BackColor
    End Select
  End If
  'Form1.Caption = DateDiff("s", horaapretada, Now)
  'miro si el boto de pantones ha estat apretat mes de 3 segons
  If horaapretada > 0 And DateDiff("s", horaapretada, Now) >= 1 Then
     modificataulapantonesstandard
     horaapretada = 1
  End If
   'AIXO DE LA COPIA S'HA DESABILITAT PER UN POSSIBLE PROBLEMA AL COPIA QUE ES QUEDA PENJAT ... S'HA DE REVISAR
  'copia la bd d'estoc del ser2 al serverprodu
 '' If Hour(Now) = 20 And Minute(Now) < 30 And (cadbl(llegir_ini("General", "diacopiafitxstoc", "comandes.ini")) <> Day(Now)) Then
 ''   Copiar_Fitxer camistocks, "\\serverprodu\dades\progcomandes\dades\copiaestocinplacsa.mdb"
 ''   escriure_ini "General", "diacopiafitxstoc", Day(Now), "comandes.ini"
 ''
 '' End If
  Exit Sub
error_screen:
'MsgBox "Error d'Screen en el Timer"
'End
End Sub
Sub modificataulapantonesstandard()
framepantones.visible = Not framepantones.visible
frameempalmes.visible = False
framebobentrada.visible = False
dblots.visible = True
dblots.AllowAddNew = True
dblots.AllowDelete = True
dblots.AllowUpdate = True
dblots.MarqueeStyle = 6
End Sub

Private Sub texteimpresio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    texteimpresio.width = 10000
    texteimpresio.ZOrder 0
End Sub

Private Sub texteimpresio_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   texteimpresio.width = 5280
End Sub

Private Sub Timer1_Timer()
  Static vcontadorminut As Byte
   'lectura_llaunes
   
   vcontadorminut = vcontadorminut + 1
   If vcontadorminut = 150 Then
        comprovarsihihamissatgesCHAT
        possarestadisticadeldia
        vcontadorminut = 0
   End If
   
End Sub
Sub comprovarsihihamissatgesCHAT()
   Dim rst As Recordset
   Set rst = dbmissatges.OpenRecordset("select * from converses_assumpte where datalectura=null and operariultimcanvi<>'T' and seccio='I' and operari=" + atrim(numop))
   If Not rst.EOF Then Command39.BackColor = &HFF00FF Else Command39.BackColor = &H17D062
   Set rst = dbmissatges.OpenRecordset("select * from converses_assumpte where  datalectura=null and operariultimcanvi<>'T' and seccio='I' and operari=" + atrim(numop2))
   If Not rst.EOF Then Command40.BackColor = &HFF80FF Else Command40.BackColor = &H6BEBB1
   Set rst = Nothing
End Sub
Function diaanteriortreballat(vdata As Date) As Date
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("select * from impressorestot where dataimpressio<#" + atrim(Format(vdata, "mm/dd/yy 00:00")) + "# order by dataimpressio desc")
   If Not rst.EOF Then
      diaanteriortreballat = rst!dataimpressio
       Else: diaanteriortreballat = vdata
   End If
   Set rst = Nothing
End Function
Sub possarestadisticadeldia()
   Dim rst As Recordset
   Dim vdatainici As String
   Dim vdatafi As String
   vdatainici = Format(diaanteriortreballat(Now), "mm/dd/yy " + Format(Now, "hh:nn"))
   vdatafi = Format(Now, "mm/dd/yy hh:mm")
   etestadistica = ""
   Set rst = dbtmpb.OpenRecordset("select sum(tmetres) as Totalmetres, count(*) as NComandes from impressorestot where impressora=7 and dataimpressio>=#" + vdatainici + "# and dataimpressio<=#" + vdatafi + "#")
   If Not rst.EOF Then etestadistica = "Estadistica ultimes 24H: Maq:FW -> " + atrim(rst!totalmetres) + " Metres i " + atrim(rst!NComandes) + " Comandes "
   Set rst = dbtmpb.OpenRecordset("select sum(tmetres) as Totalmetres, count(*) as NComandes from impressorestot where impressora=9 and dataimpressio>=#" + vdatainici + "# and dataimpressio<=#" + vdatafi + "#")
   If Not rst.EOF Then etestadistica = etestadistica + "  Maq:F2 -> " + atrim(rst!totalmetres) + " Metres i " + atrim(rst!NComandes) + " Comandes "
   Set rst = Nothing
End Sub

Sub lectura_retorn_llaunes()
  If combascula.tag = "Error" Then Exit Sub
End Sub
Sub lectura_llaunes()
   Dim v As String
   Dim rst As Recordset
   Dim rsttreball As Recordset
   Dim vpesnet As Double
   Dim rstc As Recordset
   Dim i As Byte
   If comescanerconsum.tag = "Error" Then Exit Sub
  ' v = llegirescaner(comescanerconsum)
   vpesnet = llegirpesbascula - 1.7
   If v <> "" Then
     If isloaded("formaniloxos") Then
      If guardar_llauna_consum(v) Then
          comescanerconsum.Output = "9" + Chr(7)
          formaniloxos.carregartintes cadbl(formaniloxos.tag)
         Else:
            errordellauna v
            comescanerconsum.Output = "3" + Chr(7)
      End If
        Else:
           If vpesnet > 0 Then
             ferelretorndetinta v, vpesnet, True
             comescanerconsum.Output = "9" + Chr(7)
             comprovar_retornsdellaunes
           End If
     End If
   End If
   
       'enviar pitu           comescanerconsum.Output = "9" + Chr(7)
       
End Sub
Sub errordellauna(v As String)
   formaniloxos.eterrorlectura.visible = True
   formaniloxos.eterrorlectura = formaniloxos.eterrorlectura + "   " + "Error llauna " + v + " no se a quin tinté pertany."
End Sub
Sub guardar_llauna(vnumllauna As String, id As Long)
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("select * from impresores_lotsdetinta where id=" + atrim(id) + " and numerodelot='" + atrim(vnumllauna) + "'")
   If rst.EOF Then
      dbtmpb.Execute "insert into impresores_lotsdetinta (id,numerodelot) values (" + atrim(id) + ",'" + atrim(vnumllauna) + "')"
   End If
   Set rst = Nothing
End Sub
Sub possar_llauna_pendentretorn(vnumllauna As String, id As Long)
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from impresores_retornllaunes where numllauna='" + atrim(vnumllauna) + "'")
   If rst.EOF Then
      dbtintes.Execute "insert into impresores_retornllaunes (numllauna,idliniadetinta) values ('" + atrim(vnumllauna) + "'," + atrim(id) + ")"
   End If
   Set rst = Nothing
End Sub
Sub comprovar_retornsdellaunes()
   Dim rst As Recordset
   Dim rstll As Recordset
   Set rst = dbtintes.OpenRecordset("select * from impresores_retornllaunes ")
   While Not rst.EOF
     vsql = "(((Llaunes.numllauna)='" + atrim(rst!numllauna) + "') AND ((historiallauna.tipusmoviment)='R') AND ((historiallauna.data)>#" + Format(rst!Data, "mm/dd/yy") + "#));"
     Set rstll = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, historiallauna.tipusmoviment, historiallauna.data FROM Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna WHERE " + vsql)
     If Not rstll.EOF Then dbtintes.Execute "delete from impresores_retornllaunes where numllauna='" + atrim(rst!numllauna) + "'"
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Function guardar_llauna_consum(vn As String) As Boolean
   Dim rstllauna As Recordset
   Dim rsttintatinter As Recordset
   If Mid(vn + " ", 1, 1) = "I" Then vn = formaniloxos.saber_lotactualdelcomponent(cadbl(Mid(vn + " ", 2)))
   Set rstllauna = dbtintes.OpenRecordset("select * from llaunes where numllauna=""" + vn + """")
   If Not rstllauna.EOF Then
      Set rstllauna = dbtintes.OpenRecordset("select * from tintes where idtinta=" + atrim(cadbl(rstllauna!idtinta)))
      If Not rstllauna.EOF Then
         For i = 0 To 7
          Set rsttintatinter = dbtintes.OpenRecordset("select * from tintes_tot where codi='" + atrim(cadbl(formaniloxos.tintacomanda(i).tag)) + "'")
          If Not rsttintatinter.EOF Then
            If Mid(rstllauna!referenciacolor + "   ", 1, 2) = "P-" Then
               If rsttintatinter!codi = rstllauna!codi Then
                  guardar_llauna vn, formaniloxos.compantone(i).tag
                  possar_llauna_pendentretorn vn, formaniloxos.compantone(i).tag
                  comprovar_retornsdellaunes
                  guardar_llauna_consum = True
               End If
                 Else
                   If rstllauna!idfamilia = rsttintatinter!idfamilia And rstllauna!idsubfamilia = rsttintatinter!idsubfamilia And rstllauna!idfamcolor = rsttintatinter!idfamcolor And rstllauna!idsubfamcolor = rsttintatinter!idsubfamcolor Then
                       guardar_llauna vn, formaniloxos.compantone(i).tag
                       guardar_llauna_consum = True
                   End If
            End If
          End If
         Next i
      End If
   End If
   Set rstllauna = Nothing
   Set rsttintatinter = Nothing
End Function

Private Sub veuregrupsdestoc_Click()
   Load escullirgrup
   escullirgrup.Show 1
End Sub

Sub obrirportseriebascula(vmscommx As Object, vport As Byte)
  On Error GoTo errordeport
    If Not vmscommx.PortOpen Then
      vmscommx.CommPort = vport
     ' 9600 baudios, sin paridad, 7 bits de datos y 1 bit de parada.
      vmscommx.Settings = "9600,n,8,1"
     ' If nummaq = 1 Then MSComm1.Settings = "2400,n,8,1"
     ' Indicar al control que lea todo el búfer al usar Input.
      vmscommx.InputLen = 0
     
      vmscommx.RTSEnable = True 'Por si necesitas habilitar el RTS
     
     'Abrir Puertos
     
      vmscommx.PortOpen = True
    End If
    Exit Sub
errordeport:
    MsgBox "No s'ha pogut connectar amb el port [" + atrim(vport) + "]", vbCritical, "Error"
    vmscommx.tag = "error"
End Sub
Function llegirescaner(vmscommx As Object) As String
Static buffer As String
Static nobascula As Boolean
 On Error GoTo nopossarpes
 i = 0
 
 buffer = buffer & vmscommx.Input
 If Len(buffer) > 1 Then
   'If InStr(1, buffer, "-") Then buffer = "0"
   If InStr(1, buffer, Chr$(13)) > 0 Then buffer = Mid(buffer, InStr(1, buffer, "+") + 1, InStr(1, buffer, Chr$(13)))
   'If InStr(1, buffer, ".") > 0 Then buffer = Mid(buffer, 1, InStr(1, buffer, ".") - 1) + "," + Mid(buffer, InStr(1, buffer, ".") + 1)
   llegirescaner = substituir(buffer, Chr(13), "")
'   MSComm1.Output = Chr(27) + "3,"
'   vmscommx.Output = "9" + Chr(7)
   buffer = ""
   'escriure_ini "Tintes", "pesbascula", atrim(pesbascula), fitxerini
 End If
 Exit Function
nopossarpes:
   llegirescaner = ""
End Function
Function llegirpesbascula() As Double
Static buffer As String
Static nobascula As Boolean
 On Error GoTo nopossarpes
 i = 0
 
 buffer = buffer & combascula.Input
 If Len(buffer) > 1 Then
   If InStr(1, buffer, Chr$(13)) > 0 Then buffer = Mid(buffer, InStr(1, buffer, "+") + 1, InStr(1, buffer, Chr$(13)))
   buffer = substituir(" " + buffer, "ST,GS,", "")
   buffer = substituir(buffer, ",kg", "")
   llegirpesbascula = cadbl(buffer)
   'combascula.Output = "9" + Chr(7)
   buffer = ""
 End If
 Exit Function
nopossarpes:
   llegirpesbascula = 0
End Function


Private Sub imprimir_etiquetallaunesdetintapercomanda()
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vllaunes As String
  Dim vnumc As Double
  Dim vprinter As Printer
  Dim vdatapreparada As String
  Dim vobservacions As String
  Set oapp = New CRAXDDRT.Application
  vnumc = cadbl(comanda)
  carregarobservacioidata vdatapreparada, vobservacions, vnumc
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "etiqueta_tintapreparada.rpt")
  oreport.DiscardSavedData
  carregarllaunes vllaunes, vnumc
  oreport.FormulaFields.GetItemByName("lot").text = "'" + Format(vnumc, "#,##0") + "'"
  oreport.FormulaFields.GetItemByName("llaunes").text = """" + atrim(vllaunes) + """"
  oreport.FormulaFields.GetItemByName("data").text = "'" + atrim(vdatapreparada) + "'"
  oreport.FormulaFields.GetItemByName("nummaquina").text = "'" + atrim(nummaq) + "'"
  oreport.FormulaFields.GetItemByName("observacions").text = "'" + atrim(vobservacions) + "'"
  Set vprinter = triarimpresoratickets
  'MsgBox vprinter.DeviceName
  If InStr(1, UCase(triarimpresoratickets.DeviceName), "TICKETS") > 0 Or InStr(1, UCase(triarimpresoratickets.DeviceName), "80 PRINTER") > 0 Then
     oreport.SelectPrinter vprinter.DriverName, vprinter.DeviceName, vprinter.Port
       Else: MsgBox "No s'ha trobat la impresora Tickets instal.lada al sistema", vbCritical, "Error": Exit Sub
  End If
  oreport.PaperOrientation = crDefaultPaperOrientation
  oreport.DisplayProgressDialog = False
   oreport.PrintOut False, 1

End Sub
Function triarimpresoratickets() As Printer
  
  For Each triarimpresoratickets In Printers
    If InStr(1, UCase(triarimpresoratickets.DeviceName), "TICKETS") > 0 Or InStr(1, UCase(triarimpresoratickets.DeviceName), "80 PRINTER") > 0 Then
       Exit For
    End If
  Next
  'Set triarimpresoratickets = Printer
End Function
Sub carregarllaunes(vllaunes As String, vnumc As Double)
  Dim rst As Recordset
  Dim rstdades As Recordset
  
  
  Set rstdades = dbtintes.OpenRecordset("select * from dadesllaunes", , ReadOnly)
  Set rst = dbtintes.OpenRecordset("SELECT * from assignaciollaunesacomandes where comanda=" + atrim(vnumc), , ReadOnly)
  While Not rst.EOF
     rstdades.FindFirst "numllauna='" + atrim(rst!numllauna) + "'"
     If Not rstdades.NoMatch Then
        vllaunes = vllaunes + atrim(rst!numllauna) + " " + atrim(rstdades!descripcio) + "¿"

     End If
     rst.MoveNext
  Wend
  
  Set rst = Nothing
  Set rstdades = Nothing
End Sub
Sub carregarobservacioidata(vdata As String, vobservacio As String, vnumc As Double)
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from comandesrevisadesatintes where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
      vobservacio = atrim(rst!observacio)
      vdata = atrim(rst!datacomandapreparada)
   End If
   Set rst = Nothing
End Sub

Sub actualitzarestatbobinesdesbobinadors()
   Dim cbob1 As String
   Dim cbob2 As String
   Dim cmtrs As Double
   Dim vp As Double
   Dim vb As Double
   Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
   
   cbob1 = llegir_ini("Bobines_Desbobinadors_" + atrim(nummaq), "Bobina1", rutadelfitxer(cami) + "valorsprograma.ini")
   cbob2 = llegir_ini("Bobines_Desbobinadors_" + atrim(nummaq), "Bobina2", rutadelfitxer(cami) + "valorsprograma.ini")
   convertirScanambPaletiBobina cbob1, vp, vb
   cmtrs = bobinesdentrada.calcular_mtrsdispreals(vp, vb)
   If cmtrs < 1 Then
      escriure_ini "Bobines_Desbobinadors_" + atrim(nummaq), "Bobina1", "", rutadelfitxer(cami) + "valorsprograma.ini"
      escriure_ini "Bobines_Desbobinadors_" + atrim(nummaq), "Horabob1", "", rutadelfitxer(cami) + "valorsprograma.ini"
   End If
   
   vp = 0
   vb = 0
   convertirScanambPaletiBobina cbob2, vp, vb
   cmtrs = bobinesdentrada.calcular_mtrsdispreals(vp, vb)
   If cmtrs < 1 Then
      escriure_ini "Bobines_Desbobinadors_" + atrim(nummaq), "Bobina2", "", rutadelfitxer(cami) + "valorsprograma.ini"
      escriure_ini "Bobines_Desbobinadors_" + atrim(nummaq), "Horabob2", "", rutadelfitxer(cami) + "valorsprograma.ini"
   End If
End Sub




    
 Function buscar_finestraIcolocarla(partialWindowName As String, X As Integer, Y As Integer, cx As Integer, cy As Integer) As Boolean
    Dim hwnd As Long
    Dim currentTitle As String
    Dim currentTitleLength As Long
    Dim maxTitleLength As Long

    ' Inicialitza el handle a la primera finestra.
    hwnd = FindWindowEx(0, 0, vbNullString, vbNullString)

    ' Itera a través de totes les finestres.
    Do While hwnd <> 0
        ' Obté la longitud màxima del títol de la finestra.
        maxTitleLength = 255 ' Un valor raonable.

        ' Inicialitza la variable per emmagatzemar el títol.
        currentTitle = String$(maxTitleLength, 0)

        ' Obté el títol de la finestra.
        currentTitleLength = GetWindowText(hwnd, currentTitle, maxTitleLength)

        ' Si s'ha obtingut el títol i conté el títol parcial, retorna el handle.
        If currentTitleLength > 0 Then
            currentTitle = Left$(currentTitle, currentTitleLength)
            If InStr(1, LCase$(currentTitle), LCase$(partialWindowName)) > 0 Then
                SetWindowPos hwnd, 0, X, Y, cx, cy, 0 ' &H1
                buscar_finestraIcolocarla = True
                Exit Function
            End If
        End If

        ' Obté el següent handle de finestra.
        hwnd = FindWindowEx(0, hwnd, vbNullString, vbNullString)
    Loop

    ' Si no troba la finestra, retorna 0.
    FindWindowByPartialTitle = 0
End Function

