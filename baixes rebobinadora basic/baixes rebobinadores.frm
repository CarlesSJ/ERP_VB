VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{3D20F47F-E818-4A03-AD52-45B708ACCF23}#1.0#0"; "FoxitReaderOCX.ocx"
Begin VB.Form Form1 
   Caption         =   "Baixes Comandes (Rebobinadores)"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   11895
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton botodescansrelleu 
      Height          =   390
      Left            =   11520
      Picture         =   "baixes rebobinadores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   138
      ToolTipText     =   "Control Descans i Relleu"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command28 
      Height          =   390
      Left            =   11115
      Picture         =   "baixes rebobinadores.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   140
      ToolTipText     =   "Calcul diametre"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command29 
      Height          =   390
      Left            =   10725
      Picture         =   "baixes rebobinadores.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   142
      ToolTipText     =   "Obrir Pdf Impresió"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command32 
      Height          =   390
      Left            =   10335
      Picture         =   "baixes rebobinadores.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   147
      ToolTipText     =   "Obrir word especificacions client."
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Err"
      Height          =   225
      Left            =   6660
      TabIndex        =   146
      Top             =   135
      Width           =   375
   End
   Begin FOXITREADEROCXLib.FoxitReaderOCX AcroPDF1 
      Height          =   1770
      Left            =   9240
      TabIndex        =   141
      Top             =   1470
      Visible         =   0   'False
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   3122
      _StockProps     =   0
      SRC             =   ""
   End
   Begin VB.Frame framebobentrada 
      Caption         =   "Bobines Entrada"
      Height          =   3315
      Left            =   6555
      TabIndex        =   67
      Top             =   4170
      Visible         =   0   'False
      Width           =   3435
      Begin VB.CommandButton Command30 
         Height          =   480
         Left            =   2460
         Picture         =   "baixes rebobinadores.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   145
         ToolTipText     =   "Afegir bobines d'una altra comanda."
         Top             =   2775
         Width           =   585
      End
      Begin VB.CommandButton Command24 
         Height          =   480
         Left            =   1875
         Picture         =   "baixes rebobinadores.frx":1795
         Style           =   1  'Graphical
         TabIndex        =   132
         ToolTipText     =   "Ensenyar bobines d'entrada si utilitzades."
         Top             =   2775
         Width           =   585
      End
      Begin VB.CommandButton eliminarbobentrada 
         Height          =   480
         Left            =   1260
         Picture         =   "baixes rebobinadores.frx":1D1F
         Style           =   1  'Graphical
         TabIndex        =   131
         ToolTipText     =   "Eliminar bobina d'entrada"
         Top             =   2775
         Width           =   585
      End
      Begin VB.CommandButton Command23 
         Height          =   480
         Left            =   645
         Picture         =   "baixes rebobinadores.frx":22A9
         Style           =   1  'Graphical
         TabIndex        =   130
         ToolTipText     =   "Afegir manualment el Palet/Bobina d'entrada"
         Top             =   2775
         Width           =   585
      End
      Begin VB.CommandButton botoensenyarpacking 
         Height          =   480
         Left            =   15
         Picture         =   "baixes rebobinadores.frx":2833
         Style           =   1  'Graphical
         TabIndex        =   129
         ToolTipText     =   "Sel.lecciona la bobina del Packinglist"
         Top             =   2790
         Width           =   585
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H0080FF80&
         Caption         =   "Marcar Acavada"
         Height          =   435
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   1710
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H0080FF80&
         Caption         =   "Bobines Gastades"
         Height          =   450
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   1725
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H0080FF80&
         Caption         =   "Eliminar Bobina Ent."
         Height          =   420
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   1740
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CheckBox ensenyartoteslesbobines 
         Caption         =   "Totes"
         Height          =   195
         Left            =   45
         TabIndex        =   77
         Top             =   2565
         Width           =   720
      End
      Begin MSDBGrid.DBGrid bobentrada 
         Bindings        =   "baixes rebobinadores.frx":2DBD
         Height          =   2445
         Left            =   45
         OleObjectBlob   =   "baixes rebobinadores.frx":2DD2
         TabIndex        =   109
         Top             =   90
         Width           =   3330
      End
   End
   Begin VB.CommandButton imprimir 
      Height          =   360
      Left            =   11355
      Picture         =   "baixes rebobinadores.frx":37C4
      Style           =   1  'Graphical
      TabIndex        =   139
      TabStop         =   0   'False
      ToolTipText     =   "Imprimir Etiqueta Mostra Client"
      Top             =   5205
      Width           =   375
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Pkg-Lst"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   11310
      Picture         =   "baixes rebobinadores.frx":3D4E
      Style           =   1  'Graphical
      TabIndex        =   136
      ToolTipText     =   "Imprimir Baixa sense acabar."
      Top             =   75
      Width           =   585
   End
   Begin Crystal.CrystalReport llistatpalet 
      Left            =   15
      Top             =   1245
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command22 
      Height          =   570
      Left            =   8325
      Picture         =   "baixes rebobinadores.frx":42D8
      Style           =   1  'Graphical
      TabIndex        =   128
      Top             =   150
      Width           =   465
   End
   Begin VB.CommandButton maquina 
      BackColor       =   &H00FF8080&
      Caption         =   "Maq: 0"
      Height          =   465
      Left            =   7185
      Style           =   1  'Graphical
      TabIndex        =   126
      Tag             =   "0"
      Top             =   195
      Width           =   1065
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Fulla"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   10725
      Picture         =   "baixes rebobinadores.frx":4AFA
      Style           =   1  'Graphical
      TabIndex        =   125
      ToolTipText     =   "Imprimir Baixa sense acabar."
      Top             =   75
      Width           =   585
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H008080FF&
      Caption         =   "No Acabada"
      Height          =   645
      Left            =   9915
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   75
      Width           =   810
   End
   Begin VB.TextBox linia 
      Height          =   360
      Left            =   8430
      MaxLength       =   65000
      ScrollBars      =   2  'Vertical
      TabIndex        =   123
      Text            =   $"baixes rebobinadores.frx":5084
      Top             =   -75
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Frame Frame3 
      Caption         =   "Kg"
      Height          =   570
      Left            =   5880
      TabIndex        =   117
      Top             =   180
      Width           =   1215
      Begin VB.Label etpesbascula 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Caption         =   "0,0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   45
         TabIndex        =   118
         Top             =   210
         Width           =   1110
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   -390
      Top             =   2895
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   5535
      TabIndex        =   95
      Top             =   7650
      Width           =   6315
      Begin VB.TextBox bobinesxpalet 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   5040
         TabIndex        =   120
         Top             =   465
         Width           =   465
      End
      Begin VB.TextBox tpescanutu 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   5670
         TabIndex        =   113
         Top             =   450
         Width           =   465
      End
      Begin VB.TextBox bandes 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1650
         TabIndex        =   102
         Top             =   510
         Width           =   435
      End
      Begin VB.TextBox amplemerma 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   4050
         TabIndex        =   101
         Top             =   495
         Width           =   840
      End
      Begin VB.TextBox ampleref 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2235
         TabIndex        =   100
         Top             =   510
         Width           =   840
      End
      Begin VB.TextBox bandesm 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3120
         TabIndex        =   99
         Top             =   495
         Width           =   840
      End
      Begin VB.TextBox espesor 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   885
         TabIndex        =   98
         Top             =   495
         Width           =   690
      End
      Begin VB.TextBox amplebob 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   90
         TabIndex        =   97
         Top             =   495
         Width           =   645
      End
      Begin VB.CheckBox comandaacavada 
         Caption         =   "Acavada"
         Enabled         =   0   'False
         Height          =   225
         Left            =   5250
         TabIndex        =   96
         Top             =   15
         Width           =   1005
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "BxPalet"
         Height          =   210
         Left            =   4995
         TabIndex        =   121
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "P.Canutu"
         Height          =   210
         Left            =   5595
         TabIndex        =   114
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Bandes M"
         Height          =   210
         Left            =   3165
         TabIndex        =   108
         Top             =   255
         Width           =   990
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Ample Merma"
         Height          =   210
         Left            =   3960
         TabIndex        =   107
         Top             =   255
         Width           =   990
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ample Ref"
         Height          =   210
         Left            =   2235
         TabIndex        =   106
         Top             =   255
         Width           =   990
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bandes"
         Height          =   210
         Left            =   1395
         TabIndex        =   105
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Espesor"
         Height          =   195
         Left            =   720
         TabIndex        =   104
         Top             =   255
         Width           =   990
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Ample Bob"
         Height          =   210
         Left            =   45
         TabIndex        =   103
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
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
      Picture         =   "baixes rebobinadores.frx":50A8
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "Ensenya Pantones utilitzats (Apretat x modificar)"
      Top             =   7020
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Data lots 
      Caption         =   "dblots"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10500
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "lotslam"
      Top             =   3150
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0080FF80&
      Caption         =   "Acabar Comanda"
      Height          =   645
      Left            =   8850
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   60
      Width           =   1080
   End
   Begin VB.Frame calculant 
      Height          =   2580
      Left            =   -90
      TabIndex        =   70
      Top             =   8385
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
         TabIndex        =   71
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
      Height          =   495
      Left            =   6030
      Picture         =   "baixes rebobinadores.frx":61CA
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   825
      Width           =   675
   End
   Begin VB.Data bobinesent 
      Caption         =   "bobinesentreb"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10995
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "bobinesentreb"
      Top             =   6795
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Data empalmes 
      Caption         =   "empalmes"
      Connect         =   "Access"
      DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\baixes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11070
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "lamempalmes"
      Top             =   6375
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Calcular Totals"
      Height          =   390
      Left            =   6855
      Picture         =   "baixes rebobinadores.frx":6644
      TabIndex        =   62
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data imppantones 
      Caption         =   "imppantones"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   1680
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   0
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data bobines 
      Caption         =   "bobines"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10770
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "bobinesreb"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Frame Frame2 
      Caption         =   "Totals"
      Height          =   765
      Left            =   120
      TabIndex        =   9
      Top             =   7665
      Width           =   5370
      Begin VB.TextBox hmaquina 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox hfunc 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox tkilos 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3615
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox tmetres 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   405
         Width           =   840
      End
      Begin VB.TextBox kiloshora 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   4455
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox tbob 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   405
         Width           =   840
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "H. Màquina"
         Height          =   210
         Left            =   90
         TabIndex        =   21
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Hores Func."
         Height          =   195
         Left            =   945
         TabIndex        =   20
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bob."
         Height          =   210
         Left            =   1830
         TabIndex        =   19
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Metres"
         Height          =   210
         Left            =   2700
         TabIndex        =   18
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Metres/Min"
         Height          =   210
         Left            =   4425
         TabIndex        =   17
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Kilos"
         Height          =   210
         Left            =   3630
         TabIndex        =   16
         Top             =   195
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
      Height          =   330
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Escull Operari"
      Top             =   1005
      Width           =   3675
   End
   Begin VB.Timer rellotge 
      Interval        =   255
      Left            =   345
      Top             =   420
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2025
      TabIndex        =   5
      Top             =   150
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Funcionament"
      Enabled         =   0   'False
      Height          =   525
      Left            =   4305
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   210
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Canvi Màquina"
      Enabled         =   0   'False
      Height          =   525
      Left            =   2865
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   210
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Capçalera"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7185
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   45
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Data Rebobinadores 
      Caption         =   "Rebobinadores"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Rebobinadores"
      Top             =   735
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "baixes rebobinadores.frx":731A
      Height          =   2235
      Left            =   345
      OleObjectBlob   =   "baixes rebobinadores.frx":7332
      TabIndex        =   6
      Top             =   1365
      Width           =   11610
   End
   Begin VB.TextBox comanda 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   555
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "888"
      Top             =   180
      Width           =   1215
   End
   Begin VB.Frame framebobines 
      Caption         =   "Bobines"
      Height          =   3600
      Left            =   120
      TabIndex        =   22
      Top             =   4065
      Width           =   11655
      Begin VB.TextBox etmetresbob 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2340
         Left            =   1710
         MultiLine       =   -1  'True
         TabIndex        =   143
         Top             =   990
         Visible         =   0   'False
         Width           =   3405
      End
      Begin VB.CommandButton Command26 
         Height          =   495
         Left            =   11205
         Picture         =   "baixes rebobinadores.frx":8DAB
         Style           =   1  'Graphical
         TabIndex        =   134
         TabStop         =   0   'False
         ToolTipText     =   "Bosses  i canutus utilitzats per embossar les bobines."
         Top             =   1515
         Width           =   405
      End
      Begin VB.CheckBox mostracli 
         Caption         =   "Mostra Cli."
         Height          =   195
         Left            =   10350
         TabIndex        =   127
         Top             =   885
         Width           =   1215
      End
      Begin VB.CommandButton agafarpesbascula 
         BackColor       =   &H00FFFFFF&
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
         Left            =   10230
         Picture         =   "baixes rebobinadores.frx":939D
         Style           =   1  'Graphical
         TabIndex        =   110
         ToolTipText     =   "Agafar el pes de la bàscula"
         Top             =   2490
         Width           =   945
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   10260
         Picture         =   "baixes rebobinadores.frx":B7F7
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1800
         Width           =   945
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFFFFF&
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
         Left            =   10290
         Picture         =   "baixes rebobinadores.frx":CEC9
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Ensenya Pantones utilitzats"
         Top             =   2265
         Visible         =   0   'False
         Width           =   945
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
         Height          =   660
         Left            =   10290
         TabIndex        =   24
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   10290
         Picture         =   "baixes rebobinadores.frx":DF13
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1110
         Width           =   930
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
         Left            =   11145
         TabIndex        =   25
         Top             =   255
         Width           =   375
      End
      Begin MSDBGrid.DBGrid reixabobines 
         Bindings        =   "baixes rebobinadores.frx":FB15
         Height          =   3225
         Left            =   180
         OleObjectBlob   =   "baixes rebobinadores.frx":FB27
         TabIndex        =   23
         Top             =   225
         Width           =   9990
      End
      Begin VB.Label barraestat 
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   330
         TabIndex        =   63
         Top             =   2820
         Width           =   6315
      End
   End
   Begin VB.Frame frameempalmes 
      Caption         =   "Senyals"
      Height          =   3795
      Left            =   5625
      TabIndex        =   64
      Top             =   4095
      Visible         =   0   'False
      Width           =   4725
      Begin MSDBGrid.DBGrid reixaempalmes 
         Bindings        =   "baixes rebobinadores.frx":115B3
         Height          =   3525
         Left            =   60
         OleObjectBlob   =   "baixes rebobinadores.frx":115C6
         TabIndex        =   65
         Top             =   195
         Width           =   4515
      End
   End
   Begin VB.Frame framepantones 
      Caption         =   "Adhesius"
      Height          =   3390
      Left            =   6885
      TabIndex        =   28
      Top             =   4455
      Visible         =   0   'False
      Width           =   3450
      Begin VB.TextBox Text1 
         DataField       =   "observacions"
         DataSource      =   "imppantones"
         Height          =   555
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   78
         Top             =   3330
         Width           =   3210
      End
      Begin MSDBGrid.DBGrid dblots 
         Bindings        =   "baixes rebobinadores.frx":12177
         Height          =   3705
         Left            =   30
         OleObjectBlob   =   "baixes rebobinadores.frx":12186
         TabIndex        =   73
         Top             =   180
         Visible         =   0   'False
         Width           =   3405
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg10"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   9
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   58
         Tag             =   "1"
         Top             =   2835
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot10"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   9
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   57
         Tag             =   "888"
         Top             =   2835
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone10"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   9
         Left            =   255
         MaxLength       =   40
         TabIndex        =   56
         Tag             =   "888"
         Top             =   2835
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg9"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   8
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   55
         Tag             =   "1"
         Top             =   2565
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot9"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   8
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   54
         Tag             =   "888"
         Top             =   2565
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone9"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   8
         Left            =   255
         MaxLength       =   40
         TabIndex        =   53
         Tag             =   "888"
         Top             =   2565
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg8"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   7
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   52
         Tag             =   "1"
         Top             =   2310
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot8"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   7
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   51
         Tag             =   "888"
         Top             =   2310
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone8"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   7
         Left            =   255
         MaxLength       =   40
         TabIndex        =   50
         Tag             =   "888"
         Top             =   2310
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg7"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   6
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   49
         Tag             =   "1"
         Top             =   2025
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot7"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   6
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   48
         Tag             =   "888"
         Top             =   2025
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone7"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   6
         Left            =   255
         MaxLength       =   40
         TabIndex        =   47
         Tag             =   "888"
         Top             =   2025
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg6"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   5
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   46
         Tag             =   "1"
         Top             =   1755
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot6"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   5
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   45
         Tag             =   "888"
         Top             =   1755
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone6"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   5
         Left            =   255
         MaxLength       =   40
         TabIndex        =   44
         Tag             =   "888"
         Top             =   1755
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg5"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   4
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   43
         Tag             =   "1"
         Top             =   1470
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot5"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   4
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   42
         Tag             =   "888"
         Top             =   1470
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone5"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   4
         Left            =   255
         MaxLength       =   40
         TabIndex        =   41
         Tag             =   "888"
         Top             =   1470
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg4"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   3
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   40
         Tag             =   "1"
         Top             =   1200
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot4"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   3
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   39
         Tag             =   "888"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone4"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   3
         Left            =   255
         MaxLength       =   40
         TabIndex        =   38
         Tag             =   "888"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg3"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   2
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   37
         Tag             =   "1"
         Top             =   930
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot3"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   2
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   36
         Tag             =   "888"
         Top             =   930
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone3"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   2
         Left            =   255
         MaxLength       =   40
         TabIndex        =   35
         Tag             =   "888"
         Top             =   930
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg2"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   1
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   34
         Tag             =   "1"
         Top             =   660
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot2"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   1
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   33
         Tag             =   "888"
         Top             =   660
         Width           =   1100
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone2"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   1
         Left            =   255
         MaxLength       =   40
         TabIndex        =   32
         Tag             =   "888"
         Text            =   "LIOFOL 6020"
         Top             =   660
         Width           =   1500
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg1"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   0
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   31
         Tag             =   "1"
         Top             =   375
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot1"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   0
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   30
         Tag             =   "888"
         Top             =   375
         Width           =   1100
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone1"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   0
         Left            =   255
         MaxLength       =   40
         TabIndex        =   29
         Tag             =   "888"
         Text            =   "LIOFOL 7724"
         Top             =   375
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "Observacions"
         Height          =   210
         Left            =   495
         TabIndex        =   79
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Re En"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2715
         Left            =   0
         TabIndex        =   60
         Top             =   420
         Width           =   330
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NOM            LOT               KG"
         Height          =   255
         Left            =   1065
         TabIndex        =   59
         Top             =   150
         Width           =   2295
      End
   End
   Begin VB.Frame framepalets 
      Height          =   540
      Left            =   120
      TabIndex        =   81
      Top             =   3540
      Width           =   11670
      Begin VB.CommandButton Command25 
         Height          =   360
         Left            =   9120
         Picture         =   "baixes rebobinadores.frx":12B64
         Style           =   1  'Graphical
         TabIndex        =   133
         ToolTipText     =   "Imprimir full de palet"
         Top             =   135
         Width           =   1005
      End
      Begin VB.TextBox pespalet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
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
         Height          =   360
         Left            =   495
         TabIndex        =   111
         ToolTipText     =   "Si vols pesar el palet posa't dins el camp i pitja el botó de pesar"
         Top             =   135
         Width           =   570
      End
      Begin VB.CommandButton Command18 
         Caption         =   "-30"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11145
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   120
         Width           =   480
      End
      Begin VB.CommandButton Command17 
         Caption         =   "-20"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10665
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   120
         Width           =   480
      End
      Begin VB.CommandButton Command16 
         Caption         =   "-10"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10185
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   120
         Width           =   480
      End
      Begin VB.CommandButton botopalets 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   135
         Width           =   795
      End
      Begin VB.CommandButton botopalets 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   7485
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   135
         Width           =   795
      End
      Begin VB.CommandButton botopalets 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   6690
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   135
         Width           =   795
      End
      Begin VB.CommandButton botopalets 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   5895
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   135
         Width           =   795
      End
      Begin VB.CommandButton botopalets 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   5100
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   135
         Width           =   795
      End
      Begin VB.CommandButton botopalets 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   4305
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   135
         Width           =   795
      End
      Begin VB.CommandButton botopalets 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   135
         Width           =   795
      End
      Begin VB.CommandButton botopalets 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2715
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   135
         Width           =   795
      End
      Begin VB.CommandButton botopalets 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   135
         Width           =   795
      End
      Begin VB.CommandButton botopalets 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   135
         Width           =   795
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Pes P."
         Height          =   270
         Left            =   30
         TabIndex        =   112
         Top             =   195
         Width           =   600
      End
   End
   Begin VB.Shape reciclarmaterial1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   1785
      Shape           =   3  'Circle
      Top             =   210
      Width           =   225
   End
   Begin VB.Label canutustallats 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   195
      Left            =   2595
      TabIndex        =   135
      Top             =   15
      Width           =   3195
   End
   Begin VB.Label etproblema 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   6720
      TabIndex        =   119
      Top             =   705
      Width           =   5130
   End
   Begin VB.Label hora 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   450
      TabIndex        =   7
      Top             =   870
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "Lot1:"
      Height          =   285
      Left            =   150
      TabIndex        =   76
      Top             =   225
      Width           =   525
   End
   Begin VB.Label firmat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7215
      TabIndex        =   74
      ToolTipText     =   "Codi operari que ha firmat"
      Top             =   45
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label texteimpresio 
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
      Left            =   0
      TabIndex        =   66
      Top             =   570
      Width           =   2865
   End
   Begin VB.Label client 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2280
      TabIndex        =   27
      Top             =   765
      Width           =   3675
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Nº de Comanda"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   1260
   End
   Begin VB.Label proces 
      Height          =   315
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Width           =   480
   End
   Begin VB.Label etbobinesimpost 
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
      Height          =   555
      Left            =   6720
      TabIndex        =   144
      Top             =   705
      Width           =   5130
   End
   Begin VB.Label ettoleranciaample 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   6810
      TabIndex        =   137
      Top             =   1140
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vpermetbobinesnocorrelatives As Boolean
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Sub calcular_totals(Optional obrint As Boolean)
  Dim total As Double
  Dim hores As Double
  Dim bkimp As Double
  Dim bkbob As Double
  Dim bkultim As Double
  barraestat.Caption = "Calculant els totals..."
  'calculant.Visible = True
  fcalculant.Show 0, Me
  calculant.Top = 2222
  DoEvents
  
  'On Error GoTo fi
  reixa.EditActive = False
  reixabobines.EditActive = False
  If Rebobinadores.Recordset.EOF Or cadbl(Rebobinadores.Recordset!id) = 0 Then GoTo fi
  
  '---- guardo la posicio de linies imp i de bobina x recuperarlames avall
  If Rebobinadores.Recordset!tipus = "F" Then bkimp = atrim(cadbl(Rebobinadores.Recordset!id))
  If Not bobines.Recordset.EOF Then bkbob = atrim(cadbl(bobines.Recordset!numerodebobina))
  '------
  
  On Error Resume Next
  Rebobinadores.Recordset.MoveLast
  bkultim = atrim(cadbl(Rebobinadores.Recordset!id))
  Rebobinadores.Recordset.MoveFirst
  While Not Rebobinadores.Recordset.EOF
   'On Error GoTo 0
   If Rebobinadores.Recordset!tipus = "F" Then
    If Rebobinadores.Recordset.EditMode = 0 Then Rebobinadores.Recordset.Edit
    Set rsttmp = dbtmpb.OpenRecordset("select count(*) as elgran from bobinesreb where controlid=" + atrim(Rebobinadores.Recordset!id))
    If Not rsttmp.EOF Then Rebobinadores.Recordset!totalbobines = rsttmp!elgran
  
    Set rsttmp = dbtmpb.OpenRecordset("select sum(kilos) as elgran from bobinesreb where controlid=" + atrim(Rebobinadores.Recordset!id))
    If Not rsttmp.EOF Then Rebobinadores.Recordset!totalkilos = rsttmp!elgran
  
    Set rsttmp = dbtmpb.OpenRecordset("select sum(metres) as elgran from bobinesreb where controlid=" + atrim(Rebobinadores.Recordset!id))
    If Not rsttmp.EOF Then Rebobinadores.Recordset!totalmetres = rsttmp!elgran
  
    Set rsttmp = dbtmpb.OpenRecordset("select id,metres from bobinesreb where (metres=0 or kilos=0) and controlid=" + atrim(Rebobinadores.Recordset!id))
    If Not rsttmp.EOF Then
     If rsttmp!id <> bobines.Recordset!id Then MsgBox "Hi ha bobines sense metres o kilos"
    End If
    Rebobinadores.Recordset.Update
   End If
  
   
   With Rebobinadores.Recordset
    total = 0
    'On Error Resume Next
     If Not IsDate(CVDate(atrim(!datainici))) Or Not IsDate(CVDate(atrim(!horainici))) Or Not IsDate(atrim(!horafi)) Or Not IsDate(CVDate(atrim(!datafi))) Then
      If Not obrint And Rebobinadores.Recordset!id <> bkimp And Rebobinadores.Recordset!id <> bkultim Then MsgBox "Error d'hora d'inici o final de funcionament. Corretgeix l'error per poder continuar correctament."
       Else
            total = DateDiff("n", CVDate(atrim(!datainici) + " " + atrim(!horainici)), CVDate(atrim(!datafi) + " " + atrim(!horafi)))
            total = Format(total / 60, "#,##0.00")
            
     End If
    If Rebobinadores.Recordset.EditMode = 0 Then Rebobinadores.Recordset.Edit
     Rebobinadores.Recordset!totalhores = total
     Rebobinadores.Recordset.Update
   End With
  Rebobinadores.Recordset.MoveNext
 Wend
  'If Not rsttmp.EOF Then
  'impresores.UpdateControls
  'impresores.UpdateRecord
  'reixa.Refresh
  
  On Error GoTo 0
  ensenyar_totalstotals
  possar_metres_min
  Set rstmp = Nothing
  barraestat.Caption = ""
  
  '---recupero la pocisio de linis imp i de bob
   If bkimp > 0 Then
     Rebobinadores.Recordset.FindFirst "id=" + atrim(bkimp)
     bobines.Recordset.FindFirst "numerodebobina=" + atrim(bkbob)
   Else: Rebobinadores.Recordset.MoveLast
  End If
  '---
fi:
'calculant.Visible = False
barraestat.Caption = ""
Unload fcalculant
Form1.SetFocus
End Sub

Sub possar_metres_min()
  Dim v As Double
  DoEvents
  v = cadbl(hfunc)
  f = (Int(v) * 60) + (((v - Int(v)) * 100) * 60 / 100)
  If f > 0 Then
     kiloshora = Format(cadbl(tmetres) / (f), "#.00")
    Else: kiloshora = "0"
  End If
End Sub

Sub ensenyar_totalstotals()
tbob = 0: hfunc = 0: hclixe = 0: hmaquina = 0: hajusts = 0: tkilos = 0: tmetres = 0: tprova = 0:
'total bobines
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalbobines) as elgran from Rebobinadores totalbobines where comanda=" + atrim(cadbl(comanda.Text)))
  If Not rsttmp.EOF Then tbob = cadbl(rsttmp!elgran)

  
'hores func
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from Rebobinadores totalbobines where comanda=" + atrim(cadbl(comanda.Text)) + " and tipus='F'")
  If Not rsttmp.EOF Then hfunc = cadbl(rsttmp!elgran)
  

'hores maquina
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from Rebobinadores totalbobines where comanda=" + atrim(cadbl(comanda.Text)) + " and tipus='C'")
  If Not rsttmp.EOF Then hmaquina = cadbl(rsttmp!elgran)

'total kilos
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalkilos) as elgran from Rebobinadores  where comanda=" + atrim(cadbl(comanda.Text)))
  If Not rsttmp.EOF Then tkilos = cadbl(rsttmp!elgran)
  
'total metres
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalmetres) as elgran from Rebobinadores totalbobines where comanda=" + atrim(cadbl(comanda.Text)))
  If Not rsttmp.EOF Then tmetres = cadbl(rsttmp!elgran)
  

  guarda_totals
  ensenya_totals
End Sub

Sub guarda_totals()
Set rsttmp = dbtmpb.OpenRecordset("select * from Rebobinadorestot where comanda=" + atrim(cadbl(comanda)))
  If rsttmp.EOF Then
      rsttmp.AddNew
    Else: rsttmp.Edit
  End If
  With rsttmp
    '!firmat = atrim(firmat.Caption)
    !comanda = cadbl(comanda)
    !hcanvi = cadbl(hmaquina)
    !hfuncio = cadbl(hfunc)
    !tbobines = cadbl(tbob)
    !tkilos = cadbl(tkilos)
    !tmetres = cadbl(tmetres)
    !mtrsmin = cadbl(kiloshora)
    !simulteneitat = cadbl(bandes)
    !amplebob = cadbl(amplebob)
    !espesor = cadbl(espesor)
    !ampleref = cadbl(ampleref)
    !bandesmerma = cadbl(bandesm)
    !amplemerma = cadbl(amplemerma)
    !acavada = cadbl(comandaacavada.Value)
    pescanutu = cadbl(tpescanutu.Text)
    !pescanutu = pescanutu
    !bobinesxpalet = cadbl(bobinesxpalet.Text)
    !mostraclient = IIf(mostracli.Value > 0, True, False)
    'If Not (bobines.Recordset.EOF Or bobines.Recordset.BOF) Then
    ' !kilostinta = cadbl(bobines.Recordset!kgtinta)
    ' If Not IsNull(bobines.Recordset!datafi) Then !dataimpressio = bobines.Recordset!datafi
     '!impressora = cadbl(impresores.Recordset!numeromaquina)
     '!operari = cadbl(bobines.Recordset!operari)
    'End If
   .Update
  End With
End Sub
Sub ensenya_totals()
Set rsttmp = dbtmpb.OpenRecordset("select * from Rebobinadorestot where comanda=" + atrim(cadbl(comanda)))

  With rsttmp
    'comanda = atrim(!comanda)
    'firmat = atrim(!firmat)
    hmaquina = atrim(!hcanvi)
    hfunc = atrim(!hfuncio)
    tbob = atrim(!tbobines)
    'tprova = atrim(!tprova)
    tkilos = atrim(!tkilos)
    tmetres = atrim(!tmetres)
    kiloshora = atrim(!mtrsmin)
    comandaacavada.Value = cadbl(!acavada)
    If pescanutu = 0 Then pescanutu = atrim(cadbl(!pescanutu))
    tpescanutu = pescanutu
    bobinesxpalet = cadbl(!bobinesxpalet)
    bandes = atrim(!simulteneitat)
    amplebob = atrim(!amplebob)
    espesor = atrim(!espesor)
    ampleref = atrim(!ampleref)
    bandesm = atrim(!bandesmerma)
    amplemerma = atrim(!amplemerma)
    mostracli.Value = IIf(cadbl(!mostraclient), 1, 0)
    'If Not (bobines.Recordset.EOF Or bobines.Recordset.BOF) Then
    ' !kilostinta = cadbl(bobines.Recordset!kgtinta)
    ' If Not IsNull(bobines.Recordset!datafi) Then !dataimpressio = bobines.Recordset!datafi
     '!impressora = cadbl(impresores.Recordset!numeromaquina)
     '!operari = cadbl(bobines.Recordset!operari)
    'End If
  
  End With
   missatge_exesdemtrskg
End Sub
Sub missatge_exesdemtrskg()
If cadbl(tmetres.Tag) > 0 Then
  If cadbl(tmetres.Tag) * cadbl(bandes) < cadbl(tmetres) Then
      etproblema.Caption = "Mes Metres que a la comanda. " + tmetres.Tag + " Mtrs"
       Else: etproblema.Caption = ""
  End If
End If
If cadbl(tkilos.Tag) > 0 Then
  If cadbl(tkilos) > cadbl(tkilos.Tag) Then
      etproblema.Caption = "Mes Kilos que a la comanda. " + tkilos.Tag + " Kilos"
       Else: etproblema.Caption = ""
  End If
End If

End Sub

Private Sub AcroPDF2_GotFocus()

End Sub

Sub tamany_visualitzadorpdf(vtamanygran As Boolean)
'  If vtamanygran Then
     AcroPDF1.Visible = Not AcroPDF1.Visible
     AcroPDF1.Width = 11000
     AcroPDF1.Height = 6500
     AcroPDF1.Left = 700
     AcroPDF1.ZOrder
     framebobentrada.Visible = Not AcroPDF1.Visible
 '      Else
 '       AcroPDF1.Width = 3000
 '       AcroPDF1.Height = 2000
 '       AcroPDF1.Left = 8500
 '
  'End If
End Sub

Private Sub AcroPDF1_LostFocus()
'  tamany_visualitzadorpdf False
  'AcroPDF1.Visible = False
End Sub

Private Sub agafarpesbascula_Click()
'primer miro si estic pesant el palet o la bobina
 If pespalet.Tag = "pesar" Then
       pespalet.Text = atrim(llegirpesbascula): gravar_pespalet: pespalet.Tag = "": pespalet.SetFocus: Exit Sub
 End If
 If tpescanutu.Tag = "pesarcanutu" Then
   If MsgBox("Estas segur que vols possar " + atrim(llegirpesbascula) + "Kg com a pes del canutu?" + Chr(10) + "Amb aixó les bobines pesaran pes net i pes brut.", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
      tpescanutu.Text = atrim(llegirpesbascula): guarda_totals: tpescanutu.Tag = "": pescanutu = cadbl(tpescanutu.Text): tpescanutu.SetFocus: Exit Sub
   End If
 End If
 
 
'si noes el palet doncs es la bobina

 If bobines.Recordset.EOF Then
       MsgBox "Has de sel.leccionar la bobina primer", vbInformation, "Atenció": Exit Sub
     Else
        If cadbl(bobines.Recordset!kilos) > 0 Then
           MsgBox "Aquesta bobina ja te un pes, si vols canviar-lo primer posal a zero", vbInformation, "Atenció"
           reixabobines.col = 5
           reixabobines.SetFocus
           Exit Sub
             Else:
               reixabobines.EditActive = True
               reixabobines.Columns("kilos") = atrim(llegirpesbascula)
               If pescanutu > 0 And tpescanutu.HelpContextID = 9999 Then reixabobines.Columns("pesnet") = cadbl(reixabobines.Columns("kilos")) - pescanutu
               reixabobines.col = 3
               guardar_reg_bobines
               MsgBox bobines.Recordset!kilos
               If bobines.Recordset.EditMode > 0 Then bobines.Recordset.Update
               reixabobines.EditActive = False
        End If
 End If
 reixabobines.col = 7
 reixabobines.SetFocus
 'calcular_totals
End Sub
Sub guardar_reg_bobines()
    Dim i As Byte
    Dim camp As String
    If reixabobines.row = -1 Then Exit Sub
    i = 0
    If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
    i = 0
    While i < bobines.Recordset.Fields.Count
     'reixabobines.col = i
     'camp = reixabobines.Columns(i + 1).DataField
     camp = bobines.Recordset.Fields(i).Name
     If existeixelcamp(camp) Then
       If bobines.Recordset.Fields(camp).Type <> 8 And bobines.Recordset.Fields(camp).Type <> 4 And bobines.Recordset.Fields(camp).Type <> 10 Then
         bobines.Recordset.Fields(camp) = cadbl(reixabobines.Columns(camp))
       End If
     End If
     i = i + 1
    Wend
    bobines.Recordset.Update
End Sub
Function existeixelcamp(camp As String) As Boolean
  For i = 0 To reixabobines.Columns.Count - 1
     If reixabobines.Columns(i).DataField = camp Then existeixelcamp = True
  Next i
End Function
Function llegirpesbascula() As Double
 llegirpesbascula = cadbl(etpesbascula)
End Function

Private Sub amplebob_LostFocus()
  guarda_totals

End Sub

Private Sub amplemerma_LostFocus()
  guarda_totals

End Sub

Private Sub ampleref_LostFocus()
  guarda_totals

End Sub

Private Sub bandes_LostFocus()
  guarda_totals

End Sub

Private Sub bandesm_LostFocus()
  guarda_totals

End Sub
Function micresdelmaterialcomanda(comanda As Double) As Double
  Dim rst As Recordset
  Dim valor As Double
  Set rst = dbtmp.OpenRecordset("SELECT tubolam, espessor, descripcio FROM comandes INNER JOIN mesureslineals ON comandes.mesuraesp = mesureslineals.codi where comanda=" + atrim(comanda) + ";")
  If rst.EOF Then Exit Function
  If atrim(rst!descripcio) = "GALGUES" Then
            If atrim(rst!tubolam) = "T" Then
                 valor = Format(cadbl(rst!espessor) / 4, "#,##0")
                  Else: valor = Format(cadbl(rst!espessor) / 2, "#,##0")
            End If
            Else: valor = cadbl(rst!espessor)
  End If
  Set rst = Nothing
  micresdelmaterialcomanda = valor
End Function

Function buscarmicrescomanda(comanda1 As Double) As Double
   Dim rstc1 As Recordset
   Dim rstc2 As Recordset
   Dim rstc3 As Recordset
   Dim comanda3 As Double
   Dim comanda2 As Double
   Dim espesor2 As Double
   Dim espesor1 As Double
   Dim espesor3 As Double
   
   Set rstc1 = dbtmp.OpenRecordset("select espessor,refilatd,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(comanda1))
   If rstc1.EOF Then Exit Function
   comanda2 = rstc1!linkcomanda1
   comanda3 = rstc1!linkcomanda2
   If comanda3 > 0 Then espesor3 = micresdelmaterialcomanda(comanda3)
   espesor1 = micresdelmaterialcomanda(comanda1)
   espesor2 = micresdelmaterialcomanda(comanda2)
   buscarmicrescomanda = espesor1 + espesor2 + espesor3
   Set rstc1 = Nothing
   Set rstc2 = Nothing
   Set rstc3 = Nothing
End Function

Private Sub bobentrada_DblClick()
  Dim numoptmp As Integer
  Dim nomoptmp As String
  Dim rsttmpbob As Recordset
  Dim rsttmpbobimp As Recordset
  Dim rsttmpimp As Recordset
  Dim ensenyar As String
  Dim carregataulatmp As Boolean
  Dim taulabob As String
  Exit Sub
  If r = "carregartaulatmp" Then carregartaulatmp = True
  If bobines.Recordset.EOF Then Exit Sub
  ratoli "esperar"
  On Error Resume Next
  Unload formseleccio
  On Error GoTo 0
  'If Not carregartaulatmp And cadbl(bobentrada.Columns(0).Text) = 0 Then
  '   If MsgBox("Desbobinador 1", vbYesNo, "Selecció de Desbobinador") = vbYes Then
  '       bobentrada.Columns(0).Text = "1"
  '        Else: bobentrada.Columns(0).Text = "2"
  '   End If
  'End If
  
  'If framebobentrada.Visible And Not fcalculant.Visible Then bobentrada.SetFocus
'  If bobinesent.Recordset.EOF Then
'     bobinesent.Recordset.AddNew: bobentrada_OnAddNew: bobinesent.Recordset.Update: bobentrada.Refresh
'     bobinesent.Recordset.MoveFirst
'  End If


  If ensenyartoteslesbobines <> 1 Then
     ensenyar = "not utilitzadaabaixa and"
   Else: ensenyar = ""
  End If
  
  If carregartaulatmp And sa <> "noutilitzades" Then
     ensenyar = ""
      Else:
        If sa <> "noutilitzades" And cadbl(bobentrada.Columns(0).Text) = 0 Then bobentrada.Columns(0).Text = "0": bobentrada.Columns(1).Text = "0"
  End If
  If sa = "utilitzadaabaixa and" Then ensenyar = sa
  If sa = "totes" Then ensenyar = ""
  ratoli "espera"
  obrestocks
  crear_taula_bobentrada
  Set rsttmpbob = dbtmpb.OpenRecordset("bobentradatmpreb" + atrim(nummaq))
  If proces.Tag = "E" Then
    r = "SELECT DISTINCTROW numcom, Idpalet, Idbobina FROM bobines where " + ensenyar + " (bobines.Numcom) = '" + atrim(cadbl(comanda)) + "' "
    Set rststocks = dbstocks.OpenRecordset(r)
    While Not rststocks.EOF
     rsttmpbob.AddNew
     rsttmpbob!idbobina = 0
     rsttmpbob!numlot = rststocks!numcom
     rsttmpbob!numpalet = rststocks!idpalet
     rsttmpbob!numbobent = rststocks!idbobina
     rsttmpbob!paletobob = "P"
     rsttmpbob.Update
     rststocks.MoveNext
    Wend
    Set rsttmpimp = dbtmpb.OpenRecordset("select * from impressores where tipus='F' and comanda=-1")
  End If
  
  i = 0
  
  If proces.Tag = "I" Then taulabob = "bobinesimp": Set rsttmpimp = dbtmpb.OpenRecordset("select * from impressores where tipus='F' and comanda=" + atrim(cadbl(comanda)))
  If proces.Tag = "L" Then
    If cadbl(vlink3) = 0 Then
      r = comanda
       Else: r = vlink3
    End If
    taulabob = "bobineslam": Set rsttmpimp = dbtmpb.OpenRecordset("select * from laminadores where tipus='F' and comanda=" + atrim(cadbl(r)))
  End If
  While Not rsttmpimp.EOF
    If proces.Tag = "I" Then Set rsttmpbobimp = dbtmpb.OpenRecordset("select * from bobinesimp where " + ensenyar + " controlid=" + atrim(cadbl(rsttmpimp!id)))
    If proces.Tag = "L" Then Set rsttmpbobimp = dbtmpb.OpenRecordset("select * from bobineslam where " + ensenyar + " controlid=" + atrim(cadbl(rsttmpimp!id)))
    While Not rsttmpbobimp.EOF
     rsttmpbob.AddNew
     rsttmpbob!idbobina = cadbl(rsttmpbobimp!id)
     rsttmpbob!numlot = cadbl(rsttmpimp!comanda)
     rsttmpbob!numpalet = cadbl(comanda) 'cadbl(rsttmpimp!comanda)
     rsttmpbob!numbobent = cadbl(rsttmpbobimp!numerodebobina)
     rsttmpbob!espessor = cadbl(rsttmpbobimp!espessor)
     rsttmpbob!paletobob = "B"
     rsttmpbob.Update
     rsttmpbobimp.MoveNext
    Wend
    rsttmpimp.MoveNext
  Wend
  
  Set rststocks = Nothing
  Set rsttmpbob = Nothing
  Set rsttmpbobimp = Nothing
  Set rsttmpimp = Nothing
  
  DoEvents
 ' MsgBox bobinesent.EditMode
  If carregartaulatmp Then ratoli "normal": Exit Sub
  dbtmpb.Close
  Set dbtmpb = OpenDatabase(Rebobinadores.DatabaseName)
  'MsgBox bobinesent.EditMode
  'wait (3)
 
  Set rsttmp = dbtmpb.OpenRecordset("bobentradatmpreb" + atrim(nummaq))
  If rsttmp.EOF Then
     MsgBox "No hi ha bobines d'entrada per escullir  " + Chr(13) + Chr(10) + " o estan totes utilitzades. Prova amb el botó de Totes.": dbstocks.Close: ratoli "normal"
     If bobentrada.Columns(1) = "" Then bobinesent.Recordset.CancelUpdate
     Exit Sub
   End If
   Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select * from bobentradatmpreb" + atrim(nummaq) + " order by numpalet,numbobent"
   formseleccio.Caption = "Selecció bobina d'entrada"
   formseleccio.refrescar
   'formseleccio.DBGrid2.Columns(4).Visible = False
   formseleccio.DBGrid2.Columns(0).Visible = False
   formseleccio.DBGrid2.Columns(1).Visible = False
   formseleccio.DBGrid2.Columns(2).Width = 2500
   formseleccio.DBGrid2.Columns(3).Width = 2500
   ratoli "normal"
   formseleccio.Show 1
  If sa = "utilitzadaabaixa and" Then Exit Sub
  If seleccioret = 1 Then
'   espessor = cadbl(formseleccio.Data1.Recordset!espessor)
   espessor = buscarmicrescomanda(cadbl(comanda))
   If espessor > 0 Then espesor = espessor: guarda_totals
   If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
     bobines.Recordset!espessor = cadbl(espesor)
   
   possar_camps_generals
   If cadbl(formseleccio.Data1.Recordset!idbobina) = 0 Then
       bobentrada.Columns(0) = cadbl(formseleccio.Data1.Recordset!numpalet)
       bobentrada.Columns(1) = cadbl(formseleccio.Data1.Recordset!numbobent)
       If bobinesent.Recordset.EditMode = 0 Then bobinesent.Recordset.Edit
       'si es final gravo amb majuscula si no amb minuscula per saber si estava acavada o no
       r = "b"
       If bobinesent.Recordset.RecordCount > 2 Then
        If MsgBox("Ès final de bobina?", vbYesNo, "Bobina") = vbYes Then
          r = "P": dbstocks.Execute "update  bobines set utilitzadaabaixa=True where idpalet=" + atrim(cadbl(bobentrada.Columns(0))) + " and idbobina=" + atrim(cadbl(bobentrada.Columns(1)))
            Else: r = "p": dbstocks.Execute "update  bobines set utilitzadaabaixa=False where idpalet=" + atrim(cadbl(bobentrada.Columns(0))) + " and idbobina=" + atrim(cadbl(bobentrada.Columns(1)))
        End If
       End If
       bobinesent.Recordset!paletobobina = r
       bobinesent.Recordset!idbobina = 0
       bobinesent.Recordset!id = bobines.Recordset!id
        Else
          
          bobentrada.Columns(0) = cadbl(formseleccio.Data1.Recordset!numpalet)
          bobentrada.Columns(1) = cadbl(formseleccio.Data1.Recordset!numbobent)
          If bobinesent.Recordset.EditMode = 0 Then bobinesent.Recordset.Edit
          'si es final gravo amb majuscula si no amb minuscula per saber si estava acavada o no
           r = "b"
          If bobinesent.Recordset.RecordCount > 2 Then
           If MsgBox("Ès final de bobina?", vbYesNo, "Bobina") = vbYes Then
           
            r = "B": dbtmpb.Execute "update  " + taulabob + " set utilitzadaabaixa=True where id=" + atrim(cadbl(formseleccio.Data1.Recordset!idbobina))
              Else: r = "b": dbtmpb.Execute "update  " + taulabob + " set utilitzadaabaixa=False where id=" + atrim(cadbl(formseleccio.Data1.Recordset!idbobina))
           End If
          End If
          bobinesent.Recordset!paletobobina = r
          bobinesent.Recordset!idbobina = cadbl(formseleccio.Data1.Recordset!idbobina)
          bobinesent.Recordset!id = bobines.Recordset!id
   End If
    Else: If bobinesent.Recordset.EditMode > 0 Then bobinesent.Recordset.CancelUpdate: bobentrada.Refresh
  End If
  If bobinesent.EditMode > 0 Then bobinesent.Recordset.Update
  If bobines.EditMode > 0 Then bobines.Recordset.Update
  Unload formseleccio
  If numoptmp <> 0 Then
     nomoperari = nomoptmp
     numop = numoptmp
     For Each objecte In Me
      If objecte.Name <> "llistat" And objecte.Name <> "Line1" Then
        objecte.Enabled = True
      End If
     Next objecte
      Else: If cadbl(numop) = 0 Then MsgBox "Has d'escullir un operari per treballar": Exit Sub
  End If
dbstocks.Close
possarnumbobent
End Sub
Sub possarnumbobent(Optional afegint As Boolean)
  Dim clon As Recordset
  Dim bk As String
  Dim r As String
  
  'If afegint Then GoTo cont
r = ""
 bk = atrim(bobines.Recordset!numerodebobina)
'bobinesent.UpdateRecord
bobinesent.Refresh
'If bobentrada.EditActive Then bobentrada.EditActive = False: bobinesent.UpdateRecord
Set clon = bobinesent.Recordset.Clone
 If clon.EOF Then GoTo cont
   
  clon.MoveFirst
  While Not clon.EOF
   If cadbl(clon!bobina) > 0 Then
    If r <> "" Then r = r + "/"
    r = r + atrim(clon!bobina)
      Else: clon.Delete
   End If
    clon.MoveNext
  Wend
cont:
 If Not bobines.Recordset.EOF Then
  dbtmpb.Execute "update bobinesreb set bobsent='" + atrim(r) + "' where id=" + atrim(bobines.Recordset!id)
  bobines.UpdateControls
 End If
  'bobines.Recordset.Edit
  'bobines.Recordset!bobsent = r
  'bobines.Recordset.Update
  
  
  
'  If bobinesent.Recordset.RecordCount > 1 Then
'    Set clon = bobinesent.Recordset.Clone
'    clon.MoveLast
'    clon.MovePrevious
'    marcarfidebobina cadbl(clon!palet), cadbl(clon!bobina)
'  End If
 'End If
 If bk <> "" Then
     bobines.Recordset.FindFirst "numerodebobina=" + bk
   Else: If Not bobines.Recordset.EOF Then bobines.Recordset.MoveLast
  End If
 
End Sub
Sub marcarfidebobina(nump As Double, numb As Double)
  r = "carregartaulatmp": bobentrada_DblClick: primer = False: r = ""
  bobinesent.Recordset.FindFirst "palet=" + atrim(nump) + " and bobina=" + atrim(numb)
   Set rsttmp2 = dbtmpb.OpenRecordset("select * from bobentradatmpreb" + atrim(nummaq) + " where " + "numpalet=" + atrim(nump) + " and numbobent=" + atrim(numb))
   If bobinesent.Recordset!paletobobina = "p" Or bobinesent.Recordset!paletobobina = "b" Then
    bobinesent.Recordset.Edit
    If MsgBox("Ès final de la bobina? " + atrim(rsttmp2!numpalet) + "/" + atrim(rsttmp2!numbobent), vbYesNo, "Bobina") = vbYes Then
      bobinesent.Recordset!paletobobina = UCase(bobinesent.Recordset!paletobobina)
      If UCase$(bobinesent.Recordset!paletobobina) = "P" Then
         dbstocks.Execute "update  bobines set utilitzadaabaixa=True where idpalet=" + bobentrada.Columns(0) + " and idbobina=" + bobentrada.Columns(1)
        Else:
           r = IIf(proces.Tag <> "L", "bobinesimp", "bobineslam")
           dbtmpb.Execute "update  " + r + " set utilitzadaabaixa=True where id=" + atrim(cadbl(rsttmp2!idbobina))
      End If
              
      Else
       bobinesent.Recordset!paletobobina = LCase(bobinesent.Recordset!paletobobina)
       If UCase$(bobinesent.Recordset!paletobobina) = "P" Then
          dbstocks.Execute "update  bobines set utilitzadaabaixa=False where idpalet=" + bobentrada.Columns(0) + " and idbobina=" + bobentrada.Columns(1)
        Else:
          r = IIf(proces.Tag <> "L", "bobinesimp", "bobineslam")
          dbtmpb.Execute "update  " + r + " set utilitzadaabaixa=False where id=" + atrim(cadbl(rsttmp2!idbobina))
       End If
       
    End If
    bobinesent.Recordset.Update
   End If
End Sub
Sub crear_taula_bobentrada()
  Dim camps As String
  Dim rst As Recordset
  On Error GoTo 0
  camps = "idbobina double,numlot double,numpalet double,numbobent double,espessor double,paletobob string"
  On Error GoTo borrar
  Set rst = dbtmpb.OpenRecordset("select * from bobentradatmpreb" + atrim(nummaq))
creartaula:
  'ample double,plegat double,solapa double,espessor double,metres double,kilos double)"
  dbtmpb.Execute ("delete * from bobentradatmpreb" + atrim(nummaq))
  Set rst = Nothing
  Exit Sub
borrar:
  'dbtmpb.Execute "drop table bobentradatmpreb" + atrim(nummaq)
  dbtmpb.Execute ("create table bobentradatmpreb" + atrim(nummaq) + " (" + camps) + ")"
  GoTo creartaula
  
End Sub
Sub possar_camps_generals()
  Dim rsttmp As Recordset
  If cadbl(amplebob.Text) = 0 Then
      Set rsttmp = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(comanda.Text)))
      If Not rsttmp.EOF Then
         With rsttmp
          If cadbl(bandes) = 0 Then
             bandes = atrim(!simulteneitatreb)
             If cadbl(bandes) = 0 Then bandes = "1"
          End If
          If cadbl(amplebob) = 0 Then amplebob = atrim(!amplereb)
          If cadbl(espesor) = 0 Then espesor = atrim(micrescomanda)
         End With
      End If
      Set rsttmp = Nothing
  End If
End Sub
Private Sub bobentrada_KeyUp(KeyCode As Integer, Shift As Integer)
If bobentrada.col = 1 And Len(bobentrada.Text) = 5 And KeyCode > 46 Then bobentrada.col = 2

End Sub

Private Sub bobentrada_LostFocus()
 ' SI FAIG UN LOSTFOCUS DONA ERROR AL COMPROVAR COSES AL ESCULLIR LES BOBINES D'ENTRADA
 
 'On Error Resume Next
 ' bobinesent.UpdateRecord
 ' si
 'On Error Resume Next
 If Not formseleccio.Visible Then bobinesent.UpdateRecord
 'If Not formseleccio.Visible And controlactiu <> "Command19" Then possarnumbobent
End Sub

Private Sub bobentrada_OnAddNew()
 bobinesent.Recordset!id = bobines.Recordset!id
 bobentrada.col = 0
End Sub

Private Sub bobines_Reposition()
On Error Resume Next
 empalmes.UpdateRecord
 If empalmes.Recordset.EditMode = 0 Then
   empalmes.RecordSource = "select * from lamempalmes where id=" + atrim(cadbl(bobines.Recordset!id))
   empalmes.Refresh
 End If
 bobinesent.UpdateRecord
 If bobinesent.Recordset.EditMode = 0 Then
   If cadbl(bobines.Recordset!id) = 0 Then
       bobinesent.RecordSource = "select * from bobinesentreb where id=99999999"
     Else
       bobinesent.RecordSource = "select * from bobinesentreb where id=" + atrim(cadbl(bobines.Recordset!id)) + " "
   End If
   bobinesent.Refresh
 End If
 
End Sub

Private Sub clixes_Click()
 
End Sub
Sub finalitza_seccio()
  On Error GoTo fi
  If Rebobinadores.Recordset.EOF Then Exit Sub
  On Error Resume Next
  Rebobinadores.Recordset.MoveLast
  If IsDate(Rebobinadores.Recordset!datafi) Then r = "no": Exit Sub
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
  Rebobinadores.Recordset.Edit
  Rebobinadores.Recordset!datafi = Date: Rebobinadores.Recordset!horafi = Time
  Select Case Rebobinadores.Recordset!tipus
   Case "C"
   Case "M"
   Case "A"
   Case "F"
  End Select
  Rebobinadores.Recordset.Update
calcular_totals
fi:
End Sub



Private Sub canvienfilada_DblClick()
If canvienfilada = "Si" Then
   canvienfilada = "No"
 Else: canvienfilada = "Si"
End If
End Sub

Private Sub bobinesxpalet_LostFocus()
guarda_totals
End Sub
Sub posacoloralsquehihaalgu()
  Dim rstcp As Recordset
  If bobines.Tag <> "" Then
    Set rstcp = dbtmpb.OpenRecordset("select distinct palet from bobinesreb where  controlid in(" + bobines.Tag + ")")
    While Not rstcp.EOF
     For i = 0 To 9
       If cadbl(botopalets(i).Caption) = cadbl(rstcp!palet) Then botopalets(i).BackColor = QBColor(9)
     Next i
     rstcp.MoveNext
    Wend
  End If
  Set rstcp = Nothing
End Sub

Private Sub botodescansrelleu_Click()
   Load formdescansirelleu
   If Not Rebobinadores.Recordset.EOF Then
        Rebobinadores.Recordset.MoveLast
        If Not Rebobinadores.Recordset.EOF Then
           If Not IsDate(Rebobinadores.Recordset!datafi) And Not IsDate(Rebobinadores.Recordset!horafi) Then
             formdescansirelleu.Command1.Enabled = False
             formdescansirelleu.etnonou = "Funcionament activat a baixes, primer acaba'l."
             If MsgBox("No has possat la hora de fi vols que la possi automàticament?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
                possarliniadefinalitzacio
                formdescansirelleu.Command1.Enabled = True
                formdescansirelleu.etnonou = ""
             End If
           End If
        End If
   End If
   formdescansirelleu.datacontroldescansirelleu.DatabaseName = cami
   formdescansirelleu.etnomoperari = nomoperari
   formdescansirelleu.etnomoperari.Tag = atrim(numop)
   formdescansirelleu.Show 1
End Sub
Sub possarliniadefinalitzacio()
    If Not Rebobinadores.Recordset.EOF Then
        Rebobinadores.Recordset.MoveLast
        If Not IsDate(atrim(Rebobinadores.Recordset!datafi)) Or Not IsDate(atrim(Rebobinadores.Recordset!horafi)) Then
           'If MsgBox("No hi ha la hora de fi de funcionament, Vols que el col.loqui automàticament?", vbInformation + vbYesNo, "Atenció") = vbYes Then
            Rebobinadores.Recordset.Edit
            Rebobinadores.Recordset!datafi = Date
            Rebobinadores.Recordset!horafi = Time
            Rebobinadores.Recordset.Update
           'End If
        End If
    End If
    Command4_Click
End Sub
Private Sub botoensenyarpacking_Click()

 Dim i As Byte
 Dim palet As Double
 Dim bobina As Double
 Dim utilitzades As String
 utilitzades = "noutilitzades"
 If ensenyartoteslesbobines.Value = 1 Then utilitzades = ""
 Form1.botoensenyarpacking.Tag = "afegidamanualment"
 carregar_bobinesdentrada "ensenyar" + utilitzades, 1, palet, bobina, ncomanda, , ncomanda2, IIf(proces.Tag = "invertit", True, False)
 If (bobines.Recordset.EOF And bobines.Recordset.BOF) Then MsgBox "No hiha bobina de sortida sel.leccionada": GoTo fi
 If palet > 0 And bobina > 0 Then
    'bobentrada.Columns("Palet") = atrim(palet): bobentrada.Columns("Bobina") = atrim(bobina)
'passo totes les altres a gastades
       'bobinesent.Refresh
       'While Not bobinesent.Recordset.EOF
         'carregar_bobinesdentrada "marcarutilitzadademanar", , bobinesent.Recordset!palet, bobinesent.Recordset!bobina, ncomanda, True, ncomanda2
         'bobinesent.Recordset.MoveNext
       'Wend
' fins aqui
    'afegir_labobinadentrada palet, bobina
    afegir_bobentradareb palet, bobina
    'imprimir_controlqualitatVQ cadbl(comanda)    HE PASSAT AIXÓ AL APRETAR CANVI MAQUINA DEMANAT PER PACO
    For i = 1 To cadbl(bandes)
      
      imprimiretiquetaverificacio cadbl(bobines.Recordset!numerodebobina) + (i - 1)
    Next i
    
 End If
fi:
 botoensenyarpacking.Tag = ""
 bobinesent.UpdateRecord
 possarnumbobent
 If espesor.Text <> espesorbobina Then
   espesor.Text = espesorbobina
   guarda_totals
 End If
 
End Sub
Sub imprimir_controlqualitatVQ(numc As Double)
    If cadbl(bobines.Recordset!numerodebobina) > 1 Then Exit Sub
    If preparar_etiqueta_verificaciovq(cadbl(comanda), numop, 0) Then
       imprimir_etiqueta_zebra True
   ' contadorverificacio = cadbl(tmetres) / cadbl(bandes)
       wait 2
    End If


  
End Sub
Sub imprimir_controlbobina0(numc As Double)
    
    If preparar_etiqueta_controlbobina0(cadbl(comanda), numop, 0) Then
       imprimir_etiqueta_zebra True
       wait 2
    End If


  
End Sub
Function preparar_etiqueta_verificaciovq(numc As Double, numop As Byte, numbob As Double) As Boolean
   Dim rst As Recordset
   Dim ultimalinia As String
   Dim rstproducte As Recordset
   Dim rstm As Recordset
   Dim rstc As Recordset
   preparar_etiqueta_verificaciovq = False
   Set rst = dbtmp.OpenRecordset("select client, producte,impressio,refclient,numordremodificacio,numtreball from comandes where comanda=" + atrim(numc))
   If Not rst.EOF Then
        If atrim(rst!impressio) <> "N" And atrim(rst!impressio) <> "M" Then Exit Function
   End If
   preparar_etiqueta_verificaciovq = True
   Set rstproducte = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(rst!producte) + "'")
   If rstproducte.EOF Then Exit Function
   Set rstc = dbtmp.OpenRecordset("select * from clients where codi=" + atrim(rst!client))
   If rstc.EOF Then Exit Function
   Set rstm = dbtmpb.OpenRecordset("SELECT comanda, numeromaquina FROM REBOBINADORES where comanda=" + atrim(numc))
   If rstm.EOF Then Exit Function
   Set rstm = dbtmp.OpenRecordset("select descripcio from maquines where maquina='R' and codi=" + atrim(rstm!numeromaquina))
   If rstm.EOF Then Exit Function
   ultimalinia = "Op: " + atrim(numop) + "    NºBob.Salida: 0   Fecha: " + Format(Now, "dd/mm/yy")
   
   
   Open llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetarqualitatVQrebobinadores.prn" For Input As #1
   linia.Text = Input(LOF(1), #1)
   Close #1
   With rsttmp
   substituir "#DATA#", Format(Now, "dd/mm/yy")
   substituir "#NOMMAQUINA#", atrim(rstm!descripcio)
   substituir "#TREBALL#", atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio)
   substituir "#LOT#", atrim(numc)
   substituir "#CLIENT#", Mid(atrim(rstc!nom), 1, 30)
   substituir "#REF1#", atrim(Mid(atrim(texteimpresio) + String(40, " "), 1, 30))
   substituir "#REF2#", atrim(Mid(atrim(texteimpresio) + String(40, " "), 31, 30))
   substituir "#linia#", "Op: " + atrim(numop) + "     NºBob: " + atrim(numbob) + "    Fecha: " + Format(Now, "dd/mm/yy")
   End With
   
  
End Function
Function preparar_etiqueta_controlbobina0(numc As Double, numop As Byte, numbob As Double) As Boolean
   Dim rst As Recordset
   Dim ultimalinia As String
   Dim rstproducte As Recordset
   Dim rstm As Recordset
   Dim rstc As Recordset
   
   
   preparar_etiqueta_controlbobina0 = False
   Set rst = dbtmp.OpenRecordset("select client, producte,microperforat,rebmacroperforat,impressio,refclient,numordremodificacio,numtreball from comandes where comanda=" + atrim(numc))
   
   preparar_etiqueta_controlbobina0 = True
   Set rstproducte = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(rst!producte) + "'")
   If rstproducte.EOF Then Exit Function
   Set rstc = dbtmp.OpenRecordset("select * from clients where codi=" + atrim(rst!client))
   If rstc.EOF Then Exit Function
   Set rstm = dbtmpb.OpenRecordset("SELECT comanda, numeromaquina FROM REBOBINADORES where comanda=" + atrim(numc))
   If rstm.EOF Then Exit Function
   Set rstm = dbtmp.OpenRecordset("select descripcio from maquines where maquina='R' and codi=" + atrim(rstm!numeromaquina))
   If rstm.EOF Then Exit Function
   ultimalinia = "Op: " + atrim(numop) + "    NºBob.Salida: 0   Fecha: " + Format(Now, "dd/mm/yy")
   
   Open llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetarqualitatbob0rebobinadores.prn" For Input As #1
   linia.Text = Input(LOF(1), #1)
   Close #1
   With rsttmp
   substituir "#DATA#", Format(Now, "dd/mm/yy")
   substituir "#NOMMAQUINA#", atrim(rstm!descripcio)
   substituir "#TREBALL#", atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio)
   substituir "#LOT#", atrim(numc)
   substituir "#CLIENT#", Mid(atrim(rstc!nom), 1, 30)
   substituir "#REF1#", atrim(Mid(atrim(texteimpresio) + String(40, " "), 1, 30))
   substituir "#REF2#", atrim(Mid(atrim(texteimpresio) + String(40, " "), 31, 30))
   substituir "#linia#", "Reb-" + atrim(nummaq) + " Op: " + atrim(numop) + " NºBob: " + atrim(numbob) + " Fecha: " + Format(Now, "dd/mm/yy")
   If Not vperforat Then substituir "Verificar perforat.", "": substituir "X11,463,8,41,490", ""
   End With
   
  
End Function


Sub afegir_bobentradareb(palet As Double, bobina As Double)
        marcaranteriorscomagastades IIf(proces.Tag = "invertit", True, False)
        bobinesent.Recordset.AddNew
        bobinesent.Recordset!id = bobines.Recordset!id
        bobinesent.Recordset!palet = palet
        bobinesent.Recordset!bobina = bobina
        bobinesent.Recordset.Update
        bobinesent.Refresh
End Sub
Private Sub botopalets_Click(Index As Integer)
Dim pesar As Boolean
 If Screen.ActiveControl.Name = "botopalets" And Index >= 1 Then MsgBox "Recordeu imprimir el full del palet.", vbInformation + vbOKOnly, "Recordatori"
 netejar_botons_palets
 If Index >= 0 Then numpalet = cadbl(botopalets(Index).Caption)
 botopalets(0).Tag = Trim(Index)
  
 'carrego els pesos dels palets
 Set rstpespalet = dbtmpb.OpenRecordset("select * from reb_pespalets where numpalet=" + atrim(numpalet) + " and comanda=" + atrim(cadbl(comanda.Text)))
 If Not rstpespalet.EOF Then
  If cadbl(rstpespalet!pespalet) > 0 Then
     pespalet.Text = rstpespalet!pespalet
  End If
 End If
 If Rebobinadores.Recordset.EOF Then Exit Sub
 If rstpespalet.EOF And Rebobinadores.Recordset!tipus = "F" And controlactiu = "botopalets" Then
      pespalet.Text = "0"
      While cadbl(pespalet.Text) = 0
        pespalet.Text = cadbl(InputBox("Has de possar el pes del palet o ACCEPTAR per llegir el de la bàscula.", "Possar pes del palet."))
        If cadbl(pespalet.Text) = 0 Then pespalet.Text = llegirpesbascula
        If cadbl(pespalet.Text) < 6 Or cadbl(pespalet.Text) > 30 Then
           MsgBox "Aquest pes de palet no pot ser correcte.", vbCritical + vbOKOnly, "Error de pes de palet"
           pespalet.Text = "0"
        End If
      Wend
      gravar_pespalet
 End If

'ensenyo les bobines
If Not Rebobinadores.Recordset.EOF Then
      ensenya_les_bobines
       Else: Exit Sub
 End If
 posacoloralsquehihaalgu
  If Index >= 0 Then botopalets(Index).BackColor = QBColor(14)
 


End Sub
Function controlactiu() As String
  On Error Resume Next
  controlactiu = Form1.ActiveControl.Name
End Function
Sub gravar_pespalet()
   If numpalet < 0 Or numpalet > 30 Then Exit Sub
   'On Error Resume Next
   dbtmpb.Execute "insert into reb_pespalets (comanda,numpalet,pespalet) values (" + atrim(cadbl(comanda.Text)) + "," + atrim(numpalet) + "," + passaradecimalpunt(cadbl(pespalet.Text)) + ")"
   'On Error GoTo 0
   dbtmpb.Execute "update  reb_pespalets set pespalet=" + passaradecimalpunt(cadbl(pespalet.Text)) + " where numpalet=" + atrim(numpalet) + " and comanda=" + atrim(cadbl(comanda.Text))
End Sub
Sub netejar_botons_palets()
 For i = 0 To 9
    botopalets(i).BackColor = Command4.BackColor
 Next i
End Sub

Private Sub comanda_GotFocus()
  Dim vnumc As String
  Dim vnumc_anterior As String
  vnumc = cadbl(InputBox("Entra la nova comanda", "Comanda"))
  If cadbl(vnumc) > 0 Then
     vnumc_anterior = cadbl(comanda)
     comanda = vnumc
     comanda.Tag = ""
     Command4_Click
     If comanda.Tag = "" Then comanda.Tag = atrim(vnumc_anterior): comanda = atrim(vnumc_anterior)
  End If
End Sub

Private Sub comanda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then reixa.SetFocus
End Sub

Private Sub comanda_LostFocus()
   'escriure_ini "Baixes", "ultimacomanda", comanda, "comandes.ini"
  ' Command4_Click
End Sub

Private Sub comandaacavada_Click()
  If Form1.ActiveControl.Name = "comandaacavada" Then guarda_totals
End Sub

Private Sub Command1_Click()
Load capcalera
capcalera.capcalera.DatabaseName = Rebobinadores.DatabaseName
capcalera.capcalera.RecordSource = "select * from Rebobinadorestot where comanda=" + atrim(cadbl(comanda))
capcalera.capcalera.Refresh
If capcalera.capcalera.Recordset.EOF Then
   capcalera.capcalera.Recordset.AddNew
   capcalera.capcalera.Recordset!comanda = cadbl(comanda)
   capcalera.capcalera.Recordset.Update
End If
capcalera.capcalera.Refresh
capcalera.capcalera.Recordset.Edit
capcalera.Show 1
If Form1.Rebobinadores.Recordset.EOF And Form1.Rebobinadores.Recordset.BOF Then Command2.SetFocus: Command2_Click
reixa.col = 5
reixa.SetFocus
End Sub

Private Sub Command10_Click()
Dim i As Double
Dim vbobinesErrorPesMetres As String
 If nummaq = 0 Then Exit Sub
If numbobinesnocorrelatiu(vbobinesErrorPesMetres) Then
        If UCase(InputBoxEx("Els numeros de bobines no son correlatius. Reviseu per continuar la bobina " + r + vbNewLine + "ESCRIU LA CONTRASENYA PER CONTINUAR SENSE BOBINES CORRELATIVES.", "BOBINES NO CORRELATIVES", , , , , , SPassword)) = "INPNOCORRELATIVAS" Then Exit Sub
        vpermetbobinesnocorrelatives = True
End If
If vbobinesErrorPesMetres <> "" Then MsgBox "Les bobines " + vbobinesErrorPesMetres + " no coincideix el pes amb els metres, revisa-ho abans de tancar la comanda.", vbCritical, "Error": Exit Sub

comandaacavada.Value = 0
Rebobinadores.Recordset.Move 0
If Not Rebobinadores.Recordset.EOF Then
    If Not IsDate(Rebobinadores.Recordset!datafi) Or Not IsDate(Rebobinadores.Recordset!horafi) Then
        Rebobinadores.Recordset.Edit
        Rebobinadores.Recordset!datafi = Date
        Rebobinadores.Recordset!horafi = Time
        Rebobinadores.Recordset.Update
    End If
End If

client.ToolTipText = client.Caption
crear_actualitzar_bobinesdentrada cadbl(comanda)
guarda_totals
wait 1
Command4_Click
wait 2
If MsgBox("Vols imprimir la comanda?", vbInformation + vbYesNo + vbDefaultButton1, "Atenció") = vbYes Then Command8_Click
i = cadbl(InputBox("Entra la nova comanda", "Canvi de comanda"))
If i > 0 Then comanda.Text = i: Command4_Click
End Sub

Private Sub Command11_Click()

calcular_totals
End Sub

Private Sub Command12_Click()
If bobines.Recordset.EOF Then
   MsgBox "No hi ha bobina creada"
  Else
    frameempalmes.Visible = Not frameempalmes.Visible
    framepantones.Visible = False
    framebobentrada.Visible = False
    If Not frameempalmes.Visible Then reixabobines.SetFocus
End If
End Sub

Private Sub Command13_Click()
 If bobines.Recordset.EOF Then
     MsgBox "No hi ha bobina creada"
  Else
    framebobentrada.Visible = Not framebobentrada.Visible
    framepantones.Visible = False
    frameempalmes.Visible = False
    If Not framebobentrada.Visible Then reixabobines.SetFocus
 End If
End Sub

Private Sub Command14_Click()
  Dim rstbobines As Recordset
     'reixa_BeforeDelete 0
'     If MsgBox("Segur que vols borrar aquesta linia i tot el seu contingut?", vbYesNo, "Atenció") = vbNo Then Cancel = 1
If nummaq = 0 Then Exit Sub
   If IsDate(Rebobinadores.Recordset!datafi) And IsDate(Rebobinadores.Recordset!horafi) Then
     If MsgBox("Aquesta linia ja te la hora de fi possada, SEGUR QUE VOLS ELIMINAR-LA?", vbCritical + vbYesNo + vbDefaultButton2, "ATENCIÓ") = vbNo Then Exit Sub
   End If
    r = 0
    If atrim(Rebobinadores.Recordset!tipus) = "C" Then
      If MsgBox("Segur que vols eliminar aquesta linia de CANVI?", vbCritical + vbYesNo, "Atenció ELIMINACIÓ") = vbNo Then Exit Sub
    End If
    If atrim(Rebobinadores.Recordset!tipus) = "F" Then
      Set rstbobines = dbtmpb.OpenRecordset("select * from bobinesreb where controlid=" + atrim(cadbl(Rebobinadores.Recordset!id)) + " order by numerodebobina")
      If Not rstbobines.EOF Then
       rstbobines.MoveLast
       If MsgBox("Eliminar aquesta linia pot suposar eliminar informació de " + IIf(rstbobines.RecordCount > 0, atrim(rstbobines.RecordCount), "") + " bobines.", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then MsgBox "No s'ha eliminar cap informació.": Exit Sub
       If rstbobines.RecordCount > 0 Then
        If InputBox("PER PODER ELIMINAR AQUESTES BOBINES HAS DE TECLEJAR " + Chr(13) + Chr(10) + "ELIMINAR " + atrim(rstbobines.RecordCount) + " BOBINES", "SEGURETAT PER ELIMINAR BOBINES") <> "ELIMINAR " + atrim(rstbobines.RecordCount) + " BOBINES" Then MsgBox "El texte no coincideix no s'eliminarà res.": Exit Sub
       End If
       rstbobines.MoveFirst
       On Error Resume Next
       While Not rstbobines.EOF
        If Not rstbobines.EOF Then
         'dbtmpb.Execute "delete * from lamempalmes where id=" + atrim(cadbl(rstbobines!id))
         dbtmpb.Execute "delete * from bobinesentreb where id=" + atrim(cadbl(rstbobines!id))
         rstbobines.Delete
        End If
        rstbobines.MoveNext
       Wend
       On Error GoTo 0
      End If

    End If
    'dbtmpb.Execute "delete * from bobinesreb where controlid=" + r
'    dbtmpb.Execute "delete * from impressores where id=" + atrim(r2)
    Rebobinadores.Recordset.Delete
    Rebobinadores.Recordset.MoveLast
    Rebobinadores.Refresh
    If Not Rebobinadores.Recordset.EOF Then Rebobinadores.Recordset.MoveLast
    bobines.Refresh
    Command4_Click
    
End Sub

Sub quedenbobinesentrada()
   Dim rsttmp2 As Recordset
   Dim taulabob
   If bobines.Recordset.EOF Then Exit Sub
   taulabob = IIf(proces.Tag = "I", "bobinesimp", "bobineslam")
   sa = "noutilitzades": r = "carregartaulatmp": bobentrada_DblClick
   Set rsttmp2 = dbtmpb.OpenRecordset("select * from bobentradatmpreb" + atrim(nummaq))
   r = ""
   While Not rsttmp2.EOF
     If MsgBox("He trobat la bobina " + atrim(rsttmp2!numpalet) + "/" + atrim(rsttmp2!numbobent) + " encara activa, vols donar-la per acavada?", vbCritical + vbYesNo, "Atenció") = vbYes Then
       If UCase(rsttmp2!paletobob) = "P" Then dbstocks.Execute "update  bobines set utilitzadaabaixa=True where idpalet=" + atrim(cadbl(rsttmp2!numpalet)) + " and idbobina=" + atrim(cadbl(rsttmp2!numbobent))
       If UCase(rsttmp2!paletobob) = "B" Then dbtmpb.Execute "update  " + taulabob + " set utilitzadaabaixa=True where id=" + atrim(cadbl(rsttmp2!idbobina))
     End If
     rsttmp2.MoveNext
   Wend
       
   'r = r + " " + atrim(rsttmp2!numpalet) + "/" + atrim(rsttmp2!numbobent)
   sa = ""
   Set rsttmp2 = Nothing
End Sub
Function marcarbobinacomacavada(nump As Double, numb As Double) As Boolean
Dim rsttmp2 As Recordset
   Dim taulabob
   marcarbobinacomacavada = False
   nump = cadbl(nump): numb = cadbl(numb)
   taulabob = IIf(proces.Tag = "I", "bobinesimp", "bobineslam")
   sa = "totes": r = "carregartaulatmp": bobentrada_DblClick
   Set rsttmp2 = dbtmpb.OpenRecordset("select * from bobentradatmpreb" + atrim(nummaq) + " where numpalet=" + atrim(nump) + " and numbobent=" + atrim(numb))
   If rsttmp2.EOF Then MsgBox "Aquesta bobina no la trobo asignada a aquesta comanda.": Exit Function
   If UCase(rsttmp2!paletobob) = "P" Then dbstocks.Execute "update  bobines set utilitzadaabaixa=True where idpalet=" + atrim(cadbl(nump)) + " and idbobina=" + atrim(cadbl(numb))
   If UCase(rsttmp2!paletobob) = "B" Then
     If Not rsttmp2.EOF Then dbtmpb.Execute "update  " + taulabob + " set utilitzadaabaixa=True where id=" + atrim(cadbl(rsttmp2!idbobina))
   End If
   marcarbobinacomacavada = True
   'r = r + " " + atrim(rsttmp2!numpalet) + "/" + atrim(rsttmp2!numbobent)
   sa = ""
   Set rsttmp2 = Nothing
End Function

'Sub mirar_bobinesdentrada_noacavades()
' Dim metres As Double
' Dim metresant As Double
' Dim palet As Double
' Dim bobina As Double
' Dim rstconsulta2 As Recordset
'   carregar_bobinesdentrada "carregarbobinesnoutilitzades", , , , ncomanda, , ncomanda2
'   If Not rstconsulta.EOF Or Not rstconsulta.BOF Then rstconsulta.MoveFirst
'   Set rstconsulta2 = rstconsulta.Clone
'   While Not rstconsulta2.EOF
'      palet = rstconsulta2!idpalet
'      bobina = rstconsulta2!idbobina
'      If palet > 0 And bobina > 0 And atrim(rstconsulta2!tipus) >= "O" Then
'         'es una bobina d'estock
'         metres = ncomanda
'         carregar_bobinesdentrada "metresbobinadisponible", , palet, bobina, metres, , ncomanda2
 '        metresant = metres
 '        metres = cadbl(InputBox("La bobina " + atrim(palet) + "/" + atrim(bobina) + " tenia " + atrim(metres) + " Mtrs." + Chr(10) + Chr(13) + " Quants metres has gastat?", "Bobina no acavada"))
 '         If (metresant - metres) < 500 Then
 '             If (metresant - metres) < 500 Then MsgBox "Bobines de menys de 500 metres es donen per gastades.", vbInformation, "Atenció"
 '             carregar_bobinesdentrada "metresbobinaassignar", metresant, palet, bobina, ncomanda, , ncomanda2
 '             carregar_bobinesdentrada "marcarutilitzada", , palet, bobina, ncomanda, True, ncomanda2
 '           Else:
 '              carregar_bobinesdentrada "metresbobinaassignar", metres, palet, bobina, ncomanda, , ncomanda2
 ''              carregar_bobinesdentrada "marcarutilitzada", , palet, bobina, ncomanda, True, ncomanda2
 '              If bobinesdentrada.calcular_mtrsdispreals(palet, bobina) Then carregar_bobinesdentrada "imprimirbobina", , palet, bobina
 '        End If
 '        Else
 '           'es una bobina feta a inplacsa
 '             If atrim(rstconsulta2!tipus) < "O" Then
 '                 carregar_bobinesdentrada "marcarutilitzadademanar", , palet, bobina, ncomanda, True, ncomanda2
 '             End If
 '     End If
 '     rstconsulta2.MoveNext
 '  Wend
'End Sub

Sub mirar_bobinesdentrada_noacavades()
 Dim metres As Double
 Dim metresant As Double
 Dim palet As Double
 Dim bobina As Double
 Dim rstconsulta2 As Recordset
 noespota0 = True
   carregar_bobinesdentrada "carregarbobinesnoutilitzades", , , , cadbl(comanda), , IIf(proces.Tag = "invertit", True, False)
   wait 1
   If Not rstconsulta.EOF Or Not rstconsulta.BOF Then rstconsulta.MoveFirst
   Set rstconsulta2 = rstconsulta.Clone
   'MsgBox rstconsulta!idpalet
   While Not rstconsulta2.EOF
      palet = rstconsulta2!idpalet
      bobina = rstconsulta2!idbobina
      PoB = IIf(rstconsulta2!taula = "parcials", "p", "b")
      If palet > 0 And bobina > 0 And UCase(PoB) = "P" Then 'atrim(rstconsulta2!tipus) >= "O"
           'demanar_final_palet_bobina_stock palet, bobina
           estatdelabobina palet, bobina, 0, ncomanda
           'bobinesdentrada.imprimir_bobinaparcial palet, bobina
         Else
            'es una bobina feta a inplacsa
              If UCase(PoB) = "B" Then
                  carregar_bobinesdentrada "marcarutilitzadademanar", , palet, bobina, cadbl(comanda), True, ncomanda2, IIf(proces.Tag = "invertit", True, False)
              End If
      End If
      rstconsulta2.MoveNext
   Wend
   comprovar_fi_bobsent cadbl(comanda)
   Set rstconsulta2 = Nothing
   Unload mantenimentbobina
   noespota0 = False
End Sub
Sub comprovar_fi_bobsent(numc As Double)
 Dim rstbobent As Recordset
 Dim rstpar As Recordset
 Dim palet As Double
 Dim bobina As Double
 'Set rstbobent = dbtmpb.OpenRecordset("SELECT bobinesentreb.palet, bobinesentreb.bobina, rebobinadores.comanda FROM (bobinesentreb INNER JOIN bobinesreb ON bobinesentreb.id = bobinesreb.Id) INNER JOIN rebobinadores ON bobinesreb.controlid = rebobinadores.Id WHERE (((rebobinadores.comanda)=" + atrim(numc) + "));")
 
 Set rstbobent = dbtmpb.OpenRecordset("SELECT distinct rebobinadores.comanda, bobinesentreb.palet, bobinesentreb.bobina FROM (bobinesentreb INNER JOIN bobinesreb ON bobinesentreb.id = bobinesreb.Id) INNER JOIN rebobinadores ON bobinesreb.controlid = rebobinadores.Id WHERE (((rebobinadores.comanda)=151013));")



 
 While Not rstbobent.EOF
    palet = rstbobent!palet
    bobina = rstbobent!bobina
    Set rstpar = dbstocks.OpenRecordset("select * from parcials where comanda='" + atrim(numc) + "' and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
    If Not rstpar.EOF Then
      estatdelabobina palet, bobina, 0, numc
      'bobinesdentrada.imprimir_bobinaparcial palet, bobina
    End If
    rstbobent.MoveNext
 Wend
Set rstpar = Nothing
Set rstbobent = Nothing

End Sub







Function metresfetsinferiorsacomanda(numc As Double) As Boolean
   Dim metresc As Double
   'miro els metres que no es facin mes d'un 80%
   If cadbl(tmetres) < (cadbl(tmetres.Tag) - ((cadbl(tmetres.Tag) / 10) * 2)) Then
          If UCase(InputBox("Aquesta comanda es de " + tmetres.Tag + " metres i tu has fet " + tmetres + " metres" + Chr(10) + "PASSARÉ LA COMANDA A NO ACABADA. ESCRIU ACABADA SI ESTÀ REALMENT ACABADA", "ATENCIÓ")) = "ACABADA" Then
              metresfetsinferiorsacomanda = False
               Else: metresfetsinferiorsacomanda = True
          End If
   End If
   'també vigilo els kilos fets mes d'un 80%
   If metresfetsinferiorsacomanda = False Then
        If cadbl(tkilos) < (cadbl(tkilos.Tag) - ((cadbl(tkilos.Tag) / 10) * 2)) Then
               If UCase(InputBox("Aquesta comanda es de " + tkilos.Tag + " Kilos i tu has fet " + tkilos + " Kilos" + Chr(10) + "PASSARÉ LA COMANDA A NO ACABADA. ESCRIU ACABADA SI ESTÀ REALMENT ACABADA", "ATENCIÓ")) = "ACABADA" Then
                   metresfetsinferiorsacomanda = False
                    Else: metresfetsinferiorsacomanda = True
               End If
        End If
   End If

End Function

Sub comprovar_calloffs(vcomanda As Double)
  Dim rst As Recordset
  Dim vlogcomandes As String
  Dim rutamdb As String
  Dim dbavisos As Database
  Dim rsta As Recordset
  Dim destinatari As String
  Dim cos As String
  Dim assumpte As String
  
  Set rst = dbtmp.OpenRecordset("SELECt *  FROM calloffs_detall where comanda=" + atrim(vcomanda))
  If Not rst.EOF Then
    rutamdb = rutadelfitxer(cami) + "avisosincidencies.mdb"
    Set dbavisos = DBEngine.OpenDatabase(rutamdb)
    Set rsta = dbavisos.OpenRecordset("select * from envios_mails where assumpte='" + atrim(assumpte) + "'")
    If rsta.EOF Then
       destinatari = "destinatari1"
       assumpte = treure_apostruf("Call-Offs que ja tenen producció. S'han d'assignar els palets.")
       cos = "Comanda: " + atrim(vcomanda) + "   Item: " + atrim(rst!Item) + "   Call-off: " + atrim(rst!numcalloff) + Chr(10) + Chr(13) + Chr(10) + Chr(13) + ""
       dbavisos.Execute "insert into envios_mails (data,destinatari,assumpte,cos) values (now,'" + destinatari + "','" + atrim(assumpte) + "','" + atrim(cos) + "')"
    End If
    dbavisos.Close
    Set dbavisos = Nothing
  End If
  Set rst = Nothing
  
End Sub
Private Sub Command15_Click()
Dim com As Double
Dim vbobinesErrorPesMetres As String
 If nummaq = 0 Then MsgBox "No hi ha numero de màquina assignat.": Exit Sub
 If Rebobinadores.Recordset.EOF Then Exit Sub
 If numbobinesnocorrelatiu(vbobinesErrorPesMetres) Then
     If MsgBox("Els numeros de bobines no son correlatius. Reviseu per continuar la bobina " + r + Chr(10) + "VOLS CONTINUAR IGUALMENT? O VOLS PARAR L'IMPRESIÓ I MODIFICAR-HO?", vbCritical + vbYesNo, "Atenció") <> vbYes Then Exit Sub
 End If
 If vbobinesErrorPesMetres <> "" Then
     If MsgBox("Les bobines " + vbobinesErrorPesMetres + " no coincideix el pes amb els metres, VOLS CONTINUAR IGUALMENT?", vbCritical + vbDefaultButton2 + vbYesNo, "Error") = vbNo Then Exit Sub
 End If
 quedenbobinesentrada
 client.ToolTipText = client.Caption
 If comprovarsifaltencamps Then Exit Sub
comandaacavada.Value = 1
Rebobinadores.Recordset.MoveLast
If Not Rebobinadores.Recordset.EOF Then
    If Not IsDate(Rebobinadores.Recordset!datafi) Or Not IsDate(Rebobinadores.Recordset!horafi) Then
        Rebobinadores.Recordset.Edit
        Rebobinadores.Recordset!datafi = Date
        Rebobinadores.Recordset!horafi = Time
        Rebobinadores.Recordset.Update
        Rebobinadores.Recordset.MoveFirst
        Rebobinadores.Recordset.MoveLast
        wait 1
        calcular_totals
    End If
End If
mirar_bobinesdentrada_noacavades
If metresfetsinferiorsacomanda(cadbl(comanda)) Then Command10_Click: Exit Sub
passar_comanda_a_acavada
comprovar_calloffs cadbl(comanda)
crear_actualitzar_bobinesdentrada cadbl(comanda)
calcular_totals
guarda_totals
verificacio_netejaidespeje
wait 1
Command4_Click
ratoli "espera"
wait 2
Command8_Click
If cadbl(kiloshora) = 0 Then Exit Sub
wait (3)
ratoli "normal"
com = cadbl(InputBox("Entra la nova comanda", "Fi de comanda"))
If com = 0 Then Exit Sub
comanda.Text = atrim(com)
ratoli "espera"
Command4_Click
ratoli "normal"
If cadbl(comanda.Text) = 0 Then Exit Sub
'trentats = InputBox("Quants tinters has rentat?", "Nova Comanda")
'pclixers = InputBox("Quants portaclixers?", "Nova Comanda")
'canvienfilada = InputBox("Has fet canvi d'enfilada?   S o N ", "Nova Comanda", "N")
'If Mid(canvienfilada, 1, 1) = "N" Then
'   canvienfilada = "No"
'    Else: canvienfilada = "Si"
'End If


End Sub
Sub verificacio_netejaidespeje()
  Dim v As String
  Dim vcont As Byte
  vcont = 9
  While UCase(v) <> "NETEJA" And vcont > 0
    v = InputBox("Verificació de Neteja i despeje de línia." + Chr(10) + "Escriu [neteja] per acceptar", "Neteja i despeje (" + atrim(vcont) + ")")
    vcont = vcont - 1
  Wend
End Sub

Sub crear_actualitzar_bobinesdentrada(vnumc As Double)
  Dim rsttmp As Recordset
  Dim vruta As String
  Set rsttmp = dbtmp.OpenRecordset("SELECT productes.ruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi where comandes.comanda=" + atrim(vnumc))
  If Not rsttmp.EOF Then vruta = rsttmp!ruta
  actualitzar_bobinesent vnumc, vruta
  Set rsttmp = Nothing
End Sub
Sub passar_comanda_a_acavada()
Dim estat As String
Dim ruta As String


Rebobinadores.Recordset.MoveLast
  'posso la data als totals de seccio
  If IsDate(Rebobinadores.Recordset!datafi) Then
   dbtmpb.Execute "update rebobinadorestot set datarebobinat=#" + Format(Rebobinadores.Recordset!datafi, "yy/mm/dd") + "# where comanda=" + atrim(cadbl(comanda))
   dbtmpb.Execute "update rebobinadorestot set operari=" + atrim(cadbl(Rebobinadores.Recordset!operari1)) + " where comanda=" + atrim(cadbl(comanda))
   dbtmpb.Execute "update rebobinadorestot set rebobinadora=" + atrim(cadbl(Rebobinadores.Recordset!numeromaquina)) + " where comanda=" + atrim(cadbl(comanda))
  End If

'si hi ha alguna bobina passo l'estat de la comanda a la proxima seccio
   'passo l'estat de comanda a la proxima
   Set rsttmp = dbtmp.OpenRecordset("select producte,proximaseccio from comandes where comanda=" + atrim(comanda))
   If Not rsttmp.EOF Then
     estat = atrim(rsttmp!proximaseccio)
     If estat = "" Then estat = "E"
   End If
   Set rsttmp = dbtmp.OpenRecordset("select ruta from productes where codi='" + rsttmp!producte + "'")
   If Not rsttmp.EOF Then ruta = rsttmp!ruta + "   "
   If estat = "R" Then
     seccio = Mid(ruta, InStr(1, ruta, "R") + 1, 1)
     If atrim(seccio) = "" Then seccio = "V"
     dbtmp.Execute "update comandes set proximaseccio='" + seccio + "' where comanda=" + atrim(comanda)
     dbtmp.Execute "update comandes set seccioactual='" + seccio + "' where comanda=" + atrim(comanda)
   End If
End Sub
Function comprovarsifaltencamps() As Boolean
  Dim faltenpatones As Boolean
  Dim faltenmtrs As Boolean
  Dim rstc As Recordset
  
  Rebobinadores.Recordset.FindLast "tipus='F'"
  If Not Rebobinadores.Recordset.NoMatch Then
      If cadbl(Rebobinadores.Recordset!metresminut) = 0 Then MsgBox "Falten els Metres per minut": comprovarsifaltencamps = True
  End If
  Set rstc = dbtmp.OpenRecordset("select rebkilos from comandes where comanda=" + atrim(comanda))
  If Not rstc.EOF Then
    If tkilos < (cadbl(rstc!rebkilos) - (cadbl(rstc!rebkilos) * 0.3)) Then
         If MsgBox("Els kilos que has fabricat son menys d'un 70% de la comanda." + Chr(10) + "Es segur que vols acabar comanda?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then
           comprovarsifaltencamps = True
         End If
    End If
  End If
  
End Function

Private Sub Command16_Click()
 framepalets.Tag = "0"
 possar_botons_palets
End Sub

Private Sub Command17_Click()
 framepalets.Tag = "10"
 possar_botons_palets
End Sub

Private Sub Command18_Click()
 framepalets.Tag = "20"
 possar_botons_palets
End Sub


Private Sub Command19_Click()
 If Not bobinesent.Recordset.EOF Then
  If MsgBox("Segur que vols borrar aquesta bobina d'entrada?", vbCritical + vbYesNo, "Atenció") = vbYes Then
        bobinesent.Recordset.Delete
        possarnumbobent
        bobentrada.SetFocus
  End If
 End If
End Sub

Private Sub Command2_Click()
 If nummaq = 0 Then Exit Sub
 If comprovarsidescansorelleu Then Exit Sub
 numpalet = 1
 If Not Rebobinadores.Recordset.EOF Then
  Rebobinadores.Recordset.MoveLast
  If Rebobinadores.Recordset!tipus = "C" Then
      numop = escullir_operari
      nomoperari = UCase(r)
  End If
 End If
 crearseccio "C"
 reixa.SetFocus
 ensenya_les_bobines
 colocarelsbotonsdelspalets
 'mostra 0 sempre al final d'aquest procediment
    imprimir_controlbobina0 cadbl(comanda)
    imprimir_controlqualitatVQ cadbl(comanda)
    MsgBox "Pensa a treure la mostra 0 enganxant l'etiqueta amb la direcció de sortida correcte.", vbInformation, "Mostra 0"
End Sub
Sub crearseccio(tipus As String)
 Dim com As Double
 Dim rsttmpcs As Recordset
 Dim canvicamisa As String
 canvicamisa = " "
  r = ""
  Set rsttmpcs = dbtmp.OpenRecordset("select comanda,texteimpressio from comandes where comanda=" + Trim(comanda))
  
   com = cadbl(comanda)
   If rsttmpcs.EOF Then MsgBox "No hi ha numero de comanda vàlida": com = 0
  If Not Rebobinadores.Recordset.EOF Then
      finalitza_seccio
      com = cadbl(Rebobinadores.Recordset!comanda)
  End If
  r = ""
  If com = 0 Then Exit Sub
  Rebobinadores.Recordset.AddNew
  Rebobinadores.Recordset!comanda = com
  Rebobinadores.Recordset!numeromaquina = nummaq
  Rebobinadores.Recordset!operari1 = numop
  Rebobinadores.Recordset!tipus = tipus
  Rebobinadores.Recordset!datainici = Date
  Rebobinadores.Recordset!horainici = Time
  
  'Rebobinadores.Recordset!texteimpresio = rsttmpcs!texteimpressio
  r = Rebobinadores.Recordset!id
  Rebobinadores.Recordset.Update
  Rebobinadores.Recordset.MoveLast
     Set rsttmpcs = Nothing
     
End Sub

Private Sub Command20_Click()
' r = "carregartaulatmp"
 r = ""
  sa = "utilitzadaabaixa and"
  bobentrada_DblClick
  sa = ""
  r = ""
  ratoli "normal"
End Sub

Private Sub Command21_Click()
   If marcarbobinacomacavada(cadbl(bobentrada.Columns(0)), cadbl(bobentrada.Columns(1))) Then
      MsgBox "Bobina " + atrim(cadbl(bobentrada.Columns(0))) + "/" + atrim(cadbl(bobentrada.Columns(1))) + " marcada com acavada."
   End If
   
End Sub

Private Sub Command22_Click()
   Static id As Double
 
 On Error GoTo cridar
 AppActivate id
 Exit Sub
cridar:
 id = Shell("C:\WINDOWS\SYSTEM32\CALC.EXE", vbNormalFocus)
End Sub

Private Sub Command23_Click()
 Dim desb As Byte
Dim palet As Double
  Dim bobina As Double
  Dim rst As Recordset
  Dim inssql As String
  Dim jaexisteix As Boolean
  Dim numc As Double
  Dim utili As Boolean
  Dim i As Byte
  Form1.botoensenyarpacking.Tag = "afegidamanualmentcaixa"
  demanar_paletibobina palet, bobina, desb
  'If hiha_diferencies_impostos_amblesaltres(palet, bobina) Then MsgBox "Diferencia IMPOST D'ENVASOS entre bobines d'entrada, no es pot barrejar en la mateixa bobina de sortida bobines amb IMPOST I SENSE", vbCritical, "ERROR"
  numc = ncomanda2
  If palet > 0 And bobina > 0 Then
    obrestocks
    inssql = "SELECT CDbl([comanda]) AS Expr1, Parcials.idpalet, Parcials.idbobina From Parcials WHERE (((CDbl([comanda]))<10000) and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + ");"
    Set rst = dbstocks.OpenRecordset(inssql)
    If rst.EOF Then
     inssql = "select * from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and comanda='" + atrim(numc) + "'"
     Set rst = dbstocks.OpenRecordset(inssql)
    End If
    If rst.EOF Then
      MsgBox "El Palet: " + atrim(palet) + "/" + atrim(bobina) + " no està assignat per utilitzar-lo.", vbCritical, "Palet/Bobina equivocat"
     Else
       carregar_bobinesdentrada "mirarsiutilitzada", , palet, bobina, ncomanda, utili, ncomanda2, IIf(proces.Tag = "invertit", True, False)
       If utili Then
          MsgBox "Aquesta bobina ja està marcada com utilitzada.", vbInformation + vbOKOnly, "bobina utilitzada"
           Else
            afegir_labobinadentrada palet, bobina, desb
            possarnumbobent
            For i = 1 To cadbl(bandes)
              imprimiretiquetaverificacio cadbl(bobines.Recordset!numerodebobina) + (i - 1)
            Next i
       End If
    End If
  End If
End Sub

Private Sub Command24_Click()
Dim palet As Double
Dim bobina As Double
  carregar_bobinesdentrada "ensenyarsiutilitzades", 1, palet, bobina, ncomanda, , ncomanda2
End Sub

Private Sub Command25_Click()
  Dim vnumbobsxrpalet As Double
  Dim vultimpalet As Boolean
  
  If InStr(1, Form1.Caption, "Imprimint la bobina") > 0 Then MsgBox "S'està imprimint la bobina espera a que acavi sisplau.", vbCritical, "Error": Exit Sub
  vnumbobsxrpalet = contarbobinesdelpalet(cadbl(comanda), cadbl(numpalet))
  If vnumbobsxrpalet > 0 Then
    If cadbl(InputBox("Quantes bobines hi ha en aquest palet?", "Verificació de bobines")) = vnumbobsxrpalet Then
        vultimpalet = False
        If MsgBox("A T E N C I Ó" + vbNewLine + "Es l'últim palet aquest?", vbExclamation + vbDefaultButton2 + vbYesNo, "U L T I M    P A L E T") = vbYes Then vultimpalet = True
        imprimirfullpalet cadbl(comanda), cadbl(numpalet), vultimpalet
        imprimirfullpaletdireccioentrega cadbl(comanda)
        imprimirfullPackinglistXrPalet cadbl(comanda), cadbl(numpalet)
         Else: MsgBox "No coincideix el numero de bobines amb el que has entrat," + Chr(10) + " hauria de ser " + atrim(vnumbobsxrpalet) + " Bobines, revisa que estigui tot bé.", vbExclamation, "Atenció"
    End If
  End If
End Sub
Sub imprimirfullPackinglistXrPalet(vnumc As Double, vnumpalet As Double)
    Dim rstp As Recordset
    Dim rstdire As Recordset
    Dim vnomclient As String
    Dim vdirclient As String
    Dim oapp As CRAXDDRT.Application
    Dim oreport As CRAXDDRT.Report
    Set oapp = New CRAXDDRT.Application
    Set rstp = dbtmp.OpenRecordset("SELECT comandes.comanda,comandes.direnvio as dire,clients.codi, clients.nom,comandes.refclient as refcli FROM comandes INNER JOIN clients ON comandes.client = clients.codi WHERE (((comandes.comanda)=" + atrim(vnumc) + "));")
    If Not rstp.EOF Then
        Set rstdire = dbtmp.OpenRecordset("select pfpackinglistXpalet from clients_envios where id=" + atrim(rstp!dire))
        If rstdire.EOF Then GoTo fi
        If Not cabool(rstdire!pfpackinglistXpalet) Then GoTo fi

        Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "PackinglistPerPalet.rpt", 1)
        oreport.RecordSelectionFormula = "{rebobinadores.comanda} = " + atrim(vnumc) + " and {bobinesreb.palet} = " + atrim(vnumpalet)
        oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "baixes.mdb"
        oreport.DiscardSavedData
        oreport.VerifyOnEveryPrint = True
        If existeix("c:\ordprog.ini") Then
            Load veurereport
            veurereport.CRViewer.ReportSource = oreport
            veurereport.CRViewer.DisplayGroupTree = False
            veurereport.CRViewer.ViewReport
            veurereport.Show 1, Me
              Else
               oreport.PrintOut False, 1
        End If
    End If
fi:
   Set rstp = Nothing
   Set rstdire = Nothing
End Sub

Sub imprimirfullpaletdireccioentrega(vnumc As Double)
    Dim rstp As Recordset
    Dim rstdire As Recordset
    Dim vnomclient As String
    Dim vdirclient As String
    Dim oapp As CRAXDDRT.Application
    Dim oreport As CRAXDDRT.Report
    Set oapp = New CRAXDDRT.Application
    Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "etiqueta_frontalpalet_direccioentrega_4X.rpt", 1)
    Set rstp = dbtmp.OpenRecordset("SELECT comandes.comanda,comandes.direnvio as dire,clients.codi, clients.nom,comandes.refclient as refcli FROM comandes INNER JOIN clients ON comandes.client = clients.codi WHERE (((comandes.comanda)=" + atrim(vnumc) + "));")
    If Not rstp.EOF Then
        Set rstdire = dbtmp.OpenRecordset("select nome,poblacioe from clients_envios where id=" + atrim(rstp!dire))
        vnomclient = atrim(rstp!nom)
        vdirclient = atrim(rstdire!poblacioe)
          'assigno les formules al report
        oreport.FormulaFields.GetItemByName("nomclient").Text = "'" + vnomclient + "'"
        oreport.FormulaFields.GetItemByName("direccioentrega").Text = "'" + vdirclient + "'"
        'imprimeix 4 vegades
        oreport.PrintOut False, 1
        wait 1
        oreport.PrintOut False, 1
        wait 1
        oreport.PrintOut False, 1
        wait 1
        oreport.PrintOut False, 1
    End If
   Set rstp = Nothing
   Set rstdire = Nothing
End Sub
Function contarbobinesdelpalet(vnumc As Double, nump As Double) As Double
    Dim rstp As Recordset
    contarbobinesdelpalet = 0
    Set rstp = dbtmpb.OpenRecordset("SELECT rebobinadores.comanda , bobinesreb.palet, Count(bobinesreb.numerodebobina) AS bobines, Sum(bobinesreb.kilos) AS skilos, Sum(bobinesreb.metres) AS smetres FROM rebobinadores INNER JOIN bobinesreb ON rebobinadores.Id = bobinesreb.controlid GROUP BY rebobinadores.comanda, bobinesreb.palet HAVING (((rebobinadores.comanda)=" + atrim(vnumc) + ") AND ((bobinesreb.palet)=" + atrim(nump) + "));")
    If rstp.EOF Then Exit Function
    contarbobinesdelpalet = cadbl(rstp!bobines)
    Set rstp = Nothing
End Function
Function buscarrefinplacsa(vnumc As Double) As String
    Dim rst As Recordset
    Set rst = dbtmpb.OpenRecordset("select refinplacsa from comandes_extres where comanda=" + atrim(vnumc), , ReadOnly)
    If Not rst.EOF Then buscarrefinplacsa = atrim(rst!refinplacsa)
    Set rst = Nothing
End Function
Sub imprimirfullpalet(numc As Double, nump As Double, vultimpalet As Boolean)
    Dim nomclient As String
    Dim rstp As Recordset
    Dim obsalb As String
    Dim numbobs As Double
    Dim kilos As Double
    Dim metres As Double
    Dim direnvio As String
    Dim direnvio2 As String
    Dim refclient As String
    Dim vimprimirrefinplacsa As String
    Dim dire As Double
    Set rstp = dbtmp.OpenRecordset("SELECT comandes.comanda,comandes.direnvio as dire,clients.codi, clients.nom,comandes.refclient as refcli FROM comandes INNER JOIN clients ON comandes.client = clients.codi WHERE (((comandes.comanda)=" + atrim(numc) + "));")
    If Not rstp.EOF Then
        nomclient = rstp!nom
        dire = rstp!dire
        refclient = atrim(rstp!refcli)
    End If
    Set rstp = dbtmpb.OpenRecordset("SELECT rebobinadores.comanda , bobinesreb.palet, Count(bobinesreb.numerodebobina) AS bobines, Sum(bobinesreb.kilos) AS skilos, Sum(bobinesreb.metres) AS smetres FROM rebobinadores INNER JOIN bobinesreb ON rebobinadores.Id = bobinesreb.controlid GROUP BY rebobinadores.comanda, bobinesreb.palet HAVING (((rebobinadores.comanda)=" + atrim(numc) + ") AND ((bobinesreb.palet)=" + atrim(nump) + "));")
    If rstp.EOF Then Exit Sub
    kilos = cadbl(rstp!skilos)
    metres = cadbl(rstp!smetres)
    numbobs = cadbl(rstp!bobines)
    direnvio = ""
    If dire > 0 Then
         Set rstp = dbtmp.OpenRecordset("select nome,poblacioe,observacionsalbara,paletreferenciainplacsa from clients_envios where id=" + atrim(dire))
     If Not rstp.EOF Then
        direnvio = atrim(rstp!nome)
        direnvio2 = atrim(rstp!poblacioe)
        obsalb = atrim(rstp!observacionsalbara)
        'If cabool(rstp!paletreferenciainplacsa) Then
        '   vimprimirrefinplacsa = buscarrefinplacsa(numc)
        'End If
     End If
    End If
 llistatpalet.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetapalet.rpt"

  llistatpalet.DiscardSavedData = True
 llistatpalet.Destination = crptToPrinter
 llistatpalet.CopiesToPrinter = 1
 llistatpalet.DataFiles(0) = cami
 llistatpalet.Formulas(1) = "numcomanda='" + passaradecimalpunt(Format(numc, "#,##0")) + "'"
 llistatpalet.Formulas(0) = "client='" + Mid(direnvio, 1, 20) + "'"
 llistatpalet.Formulas(6) = "client1='(" + Mid(direnvio2, 1, 15) + ")'"
 llistatpalet.Formulas(2) = "numpalet='" + atrim(nump) + "'"
 llistatpalet.Formulas(3) = "bobines='" + atrim(numbobs) + "'"
 llistatpalet.Formulas(4) = "kilos='" + passaradecimalpunt(Format(kilos, "#,##0")) + "'"
 llistatpalet.Formulas(5) = "metres='" + passaradecimalpunt(Format(metres, "#,##0")) + "'"
 llistatpalet.Formulas(7) = "envio='" + nomclient + "'"
 llistatpalet.Formulas(8) = "refclient='Ref.Client: " + refclient + "'"
 llistatpalet.Formulas(9) = "obsalbara='" + treure_apostruf(obsalb) + "'"
 llistatpalet.Formulas(10) = "ultimpalet='" + IIf(vultimpalet, "F", "") + "'"
 
 DoEvents
 If existeix("c:\ordprog.ini") Then llistatpalet.Destination = crptToWindow
 imprimir_etiquetapalets llistatpalet, llistatpalet.ReportFileName
 'llistatpalet.Action = 1
Set rstp = Nothing
    
End Sub

Private Sub Command26_Click()
  formbossesperembossar.Show 1
End Sub

Private Sub Command27_Click()
client.ToolTipText = client.Caption
calcular_totals
wait 2
imprimir_fulla "packinglistrebobinadora.rpt"
End Sub

Function comprovarsidescansorelleu() As Boolean
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("select * from controldescansrelleu where (hores=0 or hores=null) and nummaq=" + atrim(nummaq) + " and operari=" + atrim(numop) + " and seccio='" + atrim(lletraseccio) + "'")
  If rst.EOF Then Exit Function
  comprovarsidescansorelleu = True
  MsgBox UCase(nomoperari) + " en aquest moment està fent " + atrim(rst!tipus) + Chr(10) + "Primer dona per acabada la incidència.", vbExclamation, "Atenció"
End Function

Private Sub Command28_Click()
Load calculdiametre
  calculdiametre.micres = micrescomanda
  
  calculdiametre.Show 1
End Sub

Private Sub Command29_Click()
   'If AcroPDF1.Left < 5000 Then
   '   tamany_visualitzadorpdf False
      ' Else: tamany_visualitzadorpdf True
  ' End If
  tamany_visualitzadorpdf True
End Sub

Private Sub Command3_Click()
  Dim i As Byte
 Dim mtrsprova As String
 Dim mtrsparcials As Double
 Dim opantic As Byte
 Dim idbobina As Long
 If nummaq = 0 Then MsgBox "Escull primer numero de màquina": Exit Sub
 If comprovarsidescansorelleu Then Exit Sub
 
 If Not Rebobinadores.Recordset.EOF Then
    Rebobinadores.Recordset.MoveLast
    If Rebobinadores.Recordset!tipus = "A" Then
        mtrsprova = InputBox("Entra els Metres de prova.", "Atenció")
        Rebobinadores.Recordset.FindLast "tipus='A'"
        If Not Rebobinadores.Recordset.NoMatch Then
         Rebobinadores.Recordset.Edit
         Rebobinadores.Recordset!mtrsprova = cadbl(mtrsprova)
         Rebobinadores.Recordset.Update
        End If
    
    End If
    Else: Exit Sub
 End If
 'firmar_fulla
 If Rebobinadores.Recordset!tipus = "F" Then
 
    opantic = numop
    numop = escullir_operari
    nomoperari = UCase(r)
 End If
 
 crearseccio "F"
 
 
 If cadbl(bobinesxpalet) = 0 Then
   bobinesxpalet = InputBox("Entra les bobines per palet.", "Atenció")
   botopalets(0).SetFocus
   botopalets_Click 0
   If cadbl(client.Tag) = 6603 Then MsgBox "Aquest client demana un codi de barres extra per etiqueta, surtirà una etiqueta amb un codi de barres i s'ha d'enganxar amb l'ETIQUETA EXTERIOR.", vbInformation, "Atenció"
   reixa.SetFocus
 End If
 'dbtmpb.Execute "update bobinesreb set controlid=" + r + " where id=" + atrim(idbobina)
' mtrsparcials = 0
 Rebobinadores.Refresh
 Rebobinadores.Recordset.MoveLast
 While bobines.Recordset.RecordCount = 0 And mtrsparcials < 100
   DoEvents
   bobines.Refresh
   mtrsparcials = mtrsparcials + 1
 Wend
 'If bobines.Recordset.RecordCount = 0 Then Command5_Click
 colocarelsbotonsdelspalets
  tamany_visualitzadorpdf True
 ensenyar_DoscsRebiUlt cadbl(comanda)
 avisarquelacomandasestaacabant cadbl(comanda), "R"
End Sub
Function carpeta(ruta, client) As String
  If cadbl(Mid(ruta, 1, 6)) = 0 Then ruta = numcarpetaclient + " " + Trim(ruta)
  carpeta = treure_apostruf(ruta)
End Function
Sub ensenyar_DoscsRebiUlt(vnumc As Double)
  Dim rst As Recordset
  Dim rstenvio As Recordset
  Dim vnomfitxer As String
  Dim ruta_relativa_docs As String
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then
      Set rstenvio = dbtmp.OpenRecordset("select * from clients_envios where id=" + atrim(rst!direnvio))
      If Not rstenvio.EOF Then
           'If atrim(rstenvio!arxiuult) <> "" Then obrir_document "\\ord_copies\pautacli\" + atrim(rstenvio!arxiuult)
           vnomfitxer = ruta_relativa_docs + "\" + carpeta(atrim(rstenvio!arxiuult), cadbl(rst!client))
           If Not existeix(vnomfitxer) Then vnomfitxer = substituirtot(LCase(vnomfitxer), ".doc", ".docx")
           If existeix(vnomfitxer) Then obrir_document vnomfitxer
      End If
      If atrim(rst!arxiureb) <> "" Then
           vnomfitxer = ruta_relativa_docs + "\" + carpeta(rst!arxiureb, cadbl(rst!client))
           If Not existeix(vnomfitxer) Then vnomfitxer = substituirtot(LCase(vnomfitxer), ".doc", ".docx")
           obrir_document vnomfitxer
      End If
  End If
  Set rst = Nothing
End Sub
Sub avisarquelacomandasestaacabant(vnumc As Double, vseccioactual As String)
  Dim rst As Recordset
  Dim vruta As String
  Dim rstc As Recordset
  Dim vcos As String
  
  Set rstc = dbtmp.OpenRecordset("SELECT comandes.client, clients.nom, comandes.refclient, comandes.marcailinia FROM comandes INNER JOIN clients ON comandes.client = clients.codi where comanda=" + atrim(vnumc))
  Set rst = dbtmp.OpenRecordset("SELECT comandes.direnvio,comandes.comanda, comandes.producte, productes.ruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi where comanda=" + atrim(vnumc))
  If rst.EOF Then GoTo fi
  If rstc.EOF Then GoTo fi
  vruta = atrim(rst!ruta)
  If vseccioactual = Mid(vruta, Len(vruta), 1) Then
      Set rst = dbtmp.OpenRecordset("select * from clients_envios where id=" + atrim(cadbl(rst!direnvio)))
      If rst.EOF Then GoTo fi
         If atrim(rst!avisfiproduccio) <> "" Then
             vcos = atrim(rst!avisfiproduccio) + Chr(10) + Chr(10) + "Codi Client: " + atrim(rstc!client) + "-" + atrim(rstc!nom) + Chr(10) + "Ref.Client: " + atrim(rstc!refclient) + Chr(10) + "Texte imp: " + atrim(rstc!marcailinia)
             avisarfiproduccio "La comanda " + atrim(vnumc) + " està acabant la producció.", vcos
         End If
  End If
fi:
  Set rstc = Nothing
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
Sub imprimiretiquetaverificacio(numbob As Double)

    preparar_etiqueta_verificacio cadbl(comanda), numop, numbob
    imprimir_etiqueta_zebra True
    calcularvalorsreducciocilindre cadbl(comanda), numop, 1
   ' contadorverificacio = cadbl(tmetres) / cadbl(bandes)
    wait 2
End Sub
Sub firmar_fulla()
 If atrim(firmat) = "" Then
    Do
    firmat = InputBoxEx("Entra el codi d'operari o contrasenya que firma la fulla", "Atenció", , , , , , SPassword)
    If cadbl(firmat) = 1 Then MsgBox "Aquest operari ha d'apuntar la contrasenya."
    Loop Until cadbl(firmat) <> 1
    
    If cadbl(firmat) = 0 Then
       If LCase(firmat) = "picaso" Then
          firmat = "1"
         Else: firmat = ""
       End If
    End If
    guarda_totals
    passarcomandaacomençada
 End If
End Sub
Sub passarcomandaacomençada()
 dbtmp.Execute "update comandes set seccioactual='I' where comanda=" + atrim(comanda)
End Sub
Sub netejarcampsdeltotalcomanda()
    pescanutu = "0"
    tpescanutu = "0"
    bobinesxpalet = "0"
    bandes = "0"
    amplebob = "0"
    espesor = "0"
    ampleref = "0"
    bandesm = "0"
    amplemerma = "0"
    pescanutu = 0
    Command7.BackColor = &HFFFFFF
    bobines.RecordSource = "select * from bobinesreb where controlid=-1"
    bobines.Refresh
    Rebobinadores.RecordSource = "select * from Rebobinadores where comanda=-1"
    Rebobinadores.Refresh
    bobinesent.RecordSource = "select * from bobinesentreb where id=99999999"
    bobinesent.Refresh
End Sub
Function micresmaterial(descripcio As String, espesor As Double, tubolam As String) As Double
  r = espesor
  If descripcio = "GALGUES" Then
            If tubolam = "T" Then
                 r = Format(espesor / 4, "#,##0")
                  Else: r = Format(espesor / 2, "#,##0")
            End If
  End If
  If InStr(1, descripcio, "GR/") > 0 Then
    micresmaterial = espesor * -1
  End If
  micresmaterial = r
End Function



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
Function comandavalida(numc As Double, Optional nocomprovarllista As Boolean, Optional vpararcomanda As Boolean) As Boolean
   Dim rst As Recordset
   Dim proximaseccio As String
   comandavalida = False
   If numc = 0 Then Exit Function
   If Not nocomprovarllista Then
     Set rst = dbbaixes.OpenRecordset("select * from muntadora_ordremuntatge where comanda=" + atrim(numc))
     If Not rst.EOF Then MsgBox "La comanda " + atrim(numc) + " ja està a la llista.": Exit Function
   End If
   Set rst = dbtmp.OpenRecordset("SELECT comandes.comanda, productes.ruta, comandes.numordremodificacio,comandes.numtreball,comandes.proximaseccio,comandes.impressio FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(numc) + "));")
   If Not rst.EOF Then
  '     proximaseccio = posicioenlaruta(numc)
  '     If proximaseccio = "I" And InStr(1, rst!ruta, "I") > 0 Then
  '           comandavalida = True
  '             Else
  '               If InStr(1, rst!ruta, "I") = 0 Then MsgBox "La comanda " + atrim(numc) + " no te seccio d'impresores"
  '               If proximaseccio <> "I" Then
  '                 MsgBox "La comanda " + atrim(numc) + " no està apunt per imprimir. La ruta no està a I."
  '               End If
  '
  '     End If
       If rst!impressio = "F" Then
          MsgBox "A la comanda " + atrim(numc) + " li Falta Autoritzar.", vbCritical, "Atenció"
          vpararcomanda = True
            Else: comandavalida = True
       End If
        Else: MsgBox "La comanda " + atrim(numc) + " no existeix."
   End If
   If comandavalida Then
      If Not tepackinglist(cadbl(numc)) Then
         MsgBox "Aquesta comanda encara no te material assignat.", vbCritical, "Atenció"
         comandavalida = False
      End If
   End If
   carregar_pdf 0, 0
   If comandavalida And InStr(1, rst!ruta, "I") > 0 Then
     If Not clixesentratsafabrica(cadbl(numc)) Then
       comandavalida = False
       MsgBox "La comanda " + atrim(numc) + " no te els CLIXES ENTRATS a disseny. No es poden utilitzar.", vbCritical, "Atenció"
         Else
            carregar_pdf rst!numtreball, rst!numordremodificacio
     End If
   End If
End Function
Sub carregar_pdf(vnumtreball As Double, vordre As Double)
   Dim generarfitxer_pdf As String
   Dim ruta_documentacio_clixes As String
   ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
   generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(vnumtreball, "00000") + "\pdf" + Format(vnumtreball, "00000") + "-" + Format(vordre, "000") + ".pdf"
  
   If existeix(generarfitxer_pdf) Then
       AcroPDF1.OpenFile generarfitxer_pdf
       AcroPDF1.ZOrder 0
       
       'AcroPDF1.SetFocus
       'SendKeys "^H"
       'AcroPDF1.src = generarfitxer_pdf
        Else
          AcroPDF1.OpenFile rutadelfitxer(cami) + "pdfblanc.pdf"
       '   AcroPDF1.src = generarfitxer_pdf
   End If
   
   ' AcroPDF1.setShowToolbar False
   'AcroPDF1.setShowScrollbars False
'   AcroPDF1.setView ("Fit")
'   AcroPDF1.setViewScroll "Fit", 0
'   AcroPDF1.setLayoutMode "OneColumn"
'  ' AcroPDF1.setZoom 10
'   AcroPDF1.setPageMode "none"
  ' AcroPDF1.gotoFirstPage
   
   
End Sub
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
   If rst!id_estatclixe = 8 Or rst!id_estatclixe = 22 Then clixesentratsafabrica = True
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
  Dim vnumc As Double
  Dim vnumbob As Double
   escullir_comanda_semblant vnumc, vnumbob
End Sub
Sub escullir_comanda_semblant(vnumc As Double, vnumbob As Double)
   Dim vcodiclient As Double
   Dim vnumtreball As Double
   Dim vrefinplacsa As String
   Dim rstc As Recordset
   Set rstc = dbtmp.OpenRecordset("select refinplacsa from comandes_EXTRES where comanda=" + atrim(cadbl(comanda)))
   If rstc.EOF Then GoTo fi
   vrefinplacsa = "NC" + atrim(rstc!refinplacsa)
'   vcodiclient = cadbl(rstc!client)
'   vnumtreball = cadbl(rstc!numtreball)
   Load formseleccio
   formseleccio.Data1.DatabaseName = cami

   formseleccio.Data1.RecordSource = "SELECT rebobinadores.comanda, bobinesreb.numerodebobina,bobinesreb.id FROM bobinesreb LEFT JOIN rebobinadores ON bobinesreb.controlid = rebobinadores.Id WHERE rebobinadores.comanda In (select comanda from comandes_extres where refinplacsa='" + atrim(vrefinplacsa) + "') order by comanda,numerodebobina;"
   formseleccio.Caption = "Comanda amb bobines ja fetes"
   formseleccio.refrescar
   If formseleccio.Tag = "Error refrescar" Then MsgBox "No s'ha pogut carregar la consulta.", vbCritical, "Error": GoTo fi
   If formseleccio.Data1.Recordset.EOF Then MsgBox "No s'ha trobat cap comanda amb referencia " + vrefinplacsa + " per poder agafar les bobines.": GoTo fi
   formseleccio.DBGrid2.Columns(0).Width = 2200
   formseleccio.DBGrid2.Columns(1).Width = 2000
   formseleccio.DBGrid2.Columns(2).Visible = False
   formseleccio.Show 1
   If seleccioret = 1 Then
    vnumc = cadbl(formseleccio.Data1.Recordset!comanda)
    vnumbob = cadbl(formseleccio.Data1.Recordset!numerodebobina)
    vidbobreb = cadbl(formseleccio.Data1.Recordset!id)
   End If
   'afegeixo la bobina mirant primer si ja hi era
   While Not bobinesent.Recordset.EOF
         If cadbl(bobinesent.Recordset!palet) = vnumc And cadbl(bobinesent.Recordset!bobina) = vnumbob Then jaexisteix = True
         bobinesent.Recordset.MoveNext
   Wend
   If jaexisteix Then MsgBox "Aquesta bobina ja està entrada.", vbCritical, "Error": GoTo fi
   If Not jaexisteix Then
        bobinesent.Recordset.AddNew
        bobinesent.Recordset!id = vidbobreb
        bobinesent.Recordset!paletobobina = "C"
        bobinesent.Recordset!palet = vnumc
        bobinesent.Recordset!bobina = vnumbob
        bobinesent.Recordset.Update
        bobentrada.Refresh
   End If
fi:
   Set rstc = Nothing
   Unload formseleccio
End Sub

Private Sub Command31_Click()
   Dim vmsg As String
   Dim dbavisos As Database
   If MsgBox("La bàscula està pesant malament o donant error?", vbExclamation + vbDefaultButton2 + vbYesNo, "Error bàscula") = vbYes Then
       vmsg = "Etpesbascula.tag=" + etpesbascula.Tag + vbNewLine
       vmsg = vmsg + "Etpesbascula=" + etpesbascula + vbNewLine
       vmsg = vmsg + "PortOpen=" + atrim(MSComm1.PortOpen) + vbNewLine
      Set dbavisos = DBEngine.OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
      dbavisos.Execute "insert into envios_mails (data,destinatari,assumpte,cos) values (now,'miquel.inplacsa@gmail.com','Error bascula Reb:" + Trim(nummaq) + "','" + Trim(vmsg) + "')"
      MsgBox "Enviat... prova de tancar el programa i tornar a obrir.", vbCritical, "Atenció"
   End If
   Set dbavisos = Nothing
End Sub

Private Sub Command32_Click()
   ensenyar_DoscsRebiUlt cadbl(comanda)
End Sub

Private Sub Command4_Click()
  Dim rst As Recordset
  Dim rstenvio As Recordset
  Dim direnvio As Double
  Dim petit As Double
  Dim tubbase As Double
  Dim rsttmp As Recordset
  Dim nlinkcomanda2 As Double
  Dim vpararcomanda As Boolean
  Dim rstpes As Recordset
  Dim v As String
  
  vperforat = False
  If cadbl(bandes) > 0 Then
     contadorverificacio = cadbl(tmetres) / cadbl(bandes)
       Else: contadorverificacio = 1
  End If
  'comprovo si existeix la comanda
  netejarcampsdeltotalcomanda
  AcroPDF1.OpenFile "dsf"
'  AcroPDF1.src = ""
  
  Set rsttmp = dbtmp.OpenRecordset("select obsreb2,microperforat,rebmacroperforat,tubbase,refilatd,producte,client,direnvio,etrebvistiplau,rebmtrs,rebkilos,cantitatex,mesuracantex,amplereb,producte,linkcomanda1,linkcomanda2,lotmatdesb1,lotmatdesb2,rebobinadora,codibarras,espessor,comanda,refclient,comandaclient,texteimpressio,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(cadbl(comanda)))
  If rsttmp.EOF Or cadbl(comanda) = 0 Then
      MsgBox "No hi ha numero de comanda vàlida"
         Command1.Enabled = False:   Command2.Enabled = False:   Command3.Enabled = False: Exit Sub
  End If
  tubbase = IIf(Not IsNull(rsttmp!tubbase), rsttmp!tubbase, 0)
  'comprovo si hi ha seccio de rebobinadora
  If nohiharebobinadora(rsttmp!producte) Then
      MsgBox "No hi ha seccio de rebobinadora en aquesta comanda"
         Command1.Enabled = False:   Command2.Enabled = False:   Command3.Enabled = False: Exit Sub
  End If
  If Not existeix("c:\ordprog.ini") Then
    If nummaq > 0 And Not potfermicromacroperforat(atrim(rsttmp!microperforat), atrim(rsttmp!rebmacroperforat)) Then MsgBox "Aquesta comanda te Micro " + atrim(rsttmp!microperforat) + " o Macroperforat i aquesta Rebobinadora no pot fer-ho", vbCritical, "Error": Exit Sub
  End If
  If atrim(rsttmp!microperforat) <> "" And atrim(rsttmp!microperforat) <> "N" Then MsgBox "Aquesta comanda porta Microperforat en " + IIf(atrim(rsttmp!microperforat) = "C", "Calent", "Fred"), vbInformation, "ATENCIÓ": vperforat = True
  If atrim(rsttmp!rebmacroperforat) <> "N" And atrim(rsttmp!rebmacroperforat) <> "" Then MsgBox "Aquesta comanda porta MACROPERFORAT", vbInformation, "ATENCIÓ": vperforat = True
  
  If Not comandavalida(cadbl(comanda), True, vpararcomanda) Then
    If vpararcomanda Then comanda = "0": Exit Sub
    If MsgBox("Aquesta comanda ESTÀ PARADA O HI HA ALGUN MOTIU PER PARAR-LA." + Chr(10) + "Vols continuar igualment?", vbCritical + vbYesNo + vbDefaultButton2, "ATENCIÓ") = vbNo Then Exit Sub
  End If
  ncomanda = cadbl(comanda)
  ncomanda2 = IIf(cadbl(rsttmp!linkcomanda2) > 0, cadbl(rsttmp!linkcomanda2), cadbl(rsttmp!linkcomanda1))
  nlinkcomanda2 = cadbl(rsttmp!linkcomanda2)
  tpescanutu.HelpContextID = 0
  vpermetbobinesnocorrelatives = False
  numpalet = 0
  botopalets_Click 0
  bobines.Tag = ""
  proces.Tag = ""
  proces = rsttmp!producte
  If cadbl(rsttmp!refilatd) = 0 Then proces.Tag = "invertit"
  'fins aqui comprova rebobinadora
  If Not cabool(rsttmp!etrebvistiplau) Then
    avis_et_noverificada
  End If
  'miro si hi ha un missatge general per impresors
  If atrim(rsttmp!obsreb2) <> "" Then
    While v <> "LLEGIT"
      v = UCase(InputBox("Aquest client demana especificament:" + Chr(10) + Chr(10) + atrim(rsttmp!obsreb2) + Chr(10) + Chr(10) + "ESCRIU [LLEGIT] PER ACCEPTAR", "Missatge especific del client"))
    Wend
  End If
  'carrego el pes net de clients_envios
  Set rstenvio = dbtmp.OpenRecordset("select * from clients_envios where id=" + atrim(cadbl(rsttmp!direnvio)))
  etiquetesean13 = False
  If Not rstenvio.EOF Then
     If InStr(1, atrim(rstenvio!estilfrontal), "EAN13") > 0 Then
       MsgBox "INFORMACIÓ PEL MAQUINISTA..." + Chr(10) + Chr(13) + "Aquest client vol etiquetes de codidebarres per cada bobina, s'imprimirant a cada impresió d'etiqueta.", vbInformation, "Informació"
       etiquetesean13 = True
     End If
     If atrim(rstenvio!avisrebobinadora) <> "" Then
         While UCase(InputBox(atrim(rstenvio!avisrebobinadora) + Chr(10) + Chr(10) + "ESCRIU [OK] PER CONTINUAR.", "Missatge pel REBOBINADOR.")) <> "OK"
            DoEvents
         Wend
     End If
     If rstenvio!pesnetbrut Then
       direnvio = rstenvio!id
       If rstenvio!pesnetstd Then direnvio = -9999
       r = " mida<=" + atrim(cadbl(rsttmp!amplereb))
       Set rstpes = dbtmp.OpenRecordset("select * from taulapesnet where " + passaradecimalpunt(r) + " and idenvio=" + atrim(direnvio) + " order BY  mida DESC")
       If Not rstpes.EOF Then MsgBox "INFORMACIÓ PEL MAQUINISTA..." + Chr(10) + Chr(13) + "Aquesta comanda te un pes de canutu asignat.", vbInformation, "Informació": pescanutu = cadbl(rstpes!pes): tpescanutu.HelpContextID = 9999
       Set rstpes = Nothing
       
     End If
     
  End If
  'miro si la comanda te preu assignat
  comprovarsitepreuassignatosinoenviarunmail cadbl(comanda)
  'carrego els camps de l'etiqueta
  imprimir_bobina "sense imprimir"
  If Not rstopcionset.EOF Then mostracli.Visible = rstopcionset!etmostra
  
  'poso els botons de palets a punt
  framepalets.Tag = "0"
  tmetres.Tag = ""
  tkilos.Tag = ""
  'If cadbl(rsttmp!mesuracantex) = 1 Then
     tmetres.Tag = cadbl(rsttmp!rebmtrs)
     tkilos.Tag = cadbl(rsttmp!rebkilos)
  '  Else: tkilos.Tag = cadbl(rsttmp!cantitatex)
  'End If
  vlink3 = cadbl(rsttmp!linkcomanda2)
  
  amplereb = cadbl(rsttmp!amplereb)
  ettoleranciaample.Caption = "Tolerancia Ample Reb: " + atrim((amplereb * 10) - 2) + " a " + atrim((amplereb * 10) + 2) + " mm"
  
  ensenya_totals
  calcular_totals True
  bobines.RecordSource = "select * from bobinesreb where controlid=-1"
  bobines.Refresh
  
  Set rsttmp = dbtmp.OpenRecordset("select mtrslinbob,marcailinia,tubolam,codibarras,espessor,mesuraesp,comanda,refclient,comandaclient,texteimpressio from comandes where comanda=" + atrim(cadbl(comanda)))
  mesuraespcomanda = ""
  If Not rsttmp.EOF Then
     Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rsttmp!mesuraesp)))
     If Not rsttmp2.EOF Then mesuraespcomanda = rsttmp2!descripcio
  End If
  etmetresbob.Tag = atrim(cadbl(rsttmp!mtrslinbob))
  If cadbl(rsttmp!mtrslinbob) > 0 Then
    etmetresbob.Text = Chr(13) + Chr(10) + "ATENCIÓ" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "El Client vol " + etmetresbob.Tag + "Mts per bobina."
    etmetresbob.Visible = True
     Else: etmetresbob.Text = ""
  End If
  refclient = "": comandaclient = ""
  texteimpresio = ""
  refclient = atrim(rsttmp!refclient)
  comandaclient = atrim(rsttmp!comandaclient)
   'clixes.Enabled = True
  texteimpresio = IIf(atrim(rsttmp!marcailinia) = "", atrim(rsttmp!texteimpressio), atrim(rsttmp!marcailinia))
  'micrescomanda = micresmaterial(mesuraespcomanda, cadbl(rsttmp!espessor), rsttmp!tubolam)
  micrescomanda = buscarmicrescomanda(cadbl(comanda))
  codibarras = atrim(rsttmp!codibarras)
  Command1.Enabled = True: Command2.Enabled = True: Command3.Enabled = True
  
  Set rsttmp = Nothing
  'fins aqui comprovo comanda
  Rebobinadores.RecordSource = "select * from Rebobinadores where comanda=" + atrim(cadbl(comanda)) + " order by datainici,horainici"
  '* imppantones.RecordSource = "select * from Rebobinadoresadhesius where comanda=" + atrim(cadbl(comanda))
  Rebobinadores.Refresh
  'imppantones.Refresh
  
 '* If imppantones.Recordset.EOF Then
'*     crear_pantones
'* imppantones.RecordSource = "select * from Rebobinadoresadhesius where comanda=" + atrim(cadbl(comanda))
'*  End If
  possar_botons_palets
  carregar_client_ntintersialtres
  possar_camps_generals
  Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
  comprovar_sielpackinglist_hihabobinesdediferentsIMPOSTENVASOS cadbl(comanda)
  canutustallats = ""
  If comandaacavada.Value <> 1 Then
    comprovarcanutustallats tubbase
  End If
  reixa.ReBind
  calcular_totals True
  'If Rebobinadores.Recordset.EOF And Rebobinadores.Recordset.BOF And Command1.Enabled Then Command1_Click
  If Not rstenvio.EOF Then If Not rstenvio!pesnetbrut Then pescanutu = 0: tpescanutu = "0"
  framebobines.Enabled = False: framepantones.Visible = False
  'If Rebobinadores.Recordset.EOF Then Command1_Click
'  If impresores.Recordset.EOF Then MsgBox "Baixa nova es començarà amb edició de Clixes.": Command4.Tag = "nou": crearseccio "C": Command4.Tag = ""
  If bobines.Tag <> "" And bobines.Tag <> "-1" Then
   Set rsttmp = dbtmpb.OpenRecordset("select max(palet) as maxpalet from bobinesreb where  controlid in(" + bobines.Tag + ")")
   If Not rsttmp.EOF Then numpalet = cadbl(rsttmp!maxpalet)
    Else:
       numpalet = 1
  End If
  
  vcolor = comprovarsireciclarmaterial(cadbl(ncomanda))
  reciclarmaterial1.BackColor = vcolor
  If nlinkcomanda2 > 0 And vcolor <> 255 Then
    vcolor = comprovarsireciclarmaterial(cadbl(nlinkcomanda2))
    If vcolor <> 255 Then
      If Not (vcolor = &HFF00& And reciclarmaterial1.BackColor = &HFF00&) Then
           If reciclarmaterial1.BackColor <> &HFF00& Then vcolor = reciclarmaterial1.BackColor
      End If
    End If
  End If
  reciclarmaterial1.BackColor = vcolor
  colocarelsbotonsdelspalets
  If Rebobinadores.Recordset.EOF Then If proces.Tag = "invertit" Then passarbobinesentradanoutilitzades cadbl(comanda)
  comanda.Tag = comanda.Text
  tamany_visualitzadorpdf True
  ' ensenyardadaexpedicio ncomanda
  ratoli "normal"
End Sub
Sub comprovar_sielpackinglist_hihabobinesdediferentsIMPOSTENVASOS(vnumc As Double)
  Dim rst As Recordset
  Dim rstalbarans As Recordset
  Dim vpaletsimpost As Boolean
  Dim vpaletsSENSEimpost As Boolean
  etbobinesimpost = ""
  Set rst = dbstocks.OpenRecordset("SELECT palets.teimpost FROM Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet where parcials.comanda='" + atrim(vnumc) + "'")
  While Not rst.EOF
     If cabool(rst!teimpost) Then
             vpaletsimpost = True
           Else: vpaletsSENSEimpost = True
     End If
     rst.MoveNext
  Wend
  If vpaletsSENSEimpost And vpaletsimpost Then
      etbobinesimpost = "LES BOBINES AMB IMPOST DIFERENT NO ES PODEN AJUNTAR."
  End If
  Set rst = Nothing
End Sub
Sub ensenyardadaexpedicio(vnumc As Double)
  Dim rst As Recordset
  Dim dbplanificacio As Database
  Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificacio.mdb", , True)
  Set rst = dbplanificacio.OpenRecordset("select data from linies_expedicions where comanda=" + atrim(vnumc))
  If Not rst.EOF Then
      MsgBox "Aquesta comanda ja te una data d'EXPEDICIÓ.  " + atrim(rst!data), vbInformation, "A T E N C I Ó"
  End If
  Set dbplanificacio = Nothing
  Set rst = Nothing
End Sub
Function potfermicromacroperforat(vmicro As String, vmacro As String) As Boolean
   Dim rst As Recordset
   
   potfermicromacroperforat = True
   Set rst = dbtmp.OpenRecordset("select * from maquines where maquina='R' and codi=" + atrim(nummaq))
   If Not rst.EOF Then
      If atrim(rst!rebmicromacro) = "Tots" Then potfermicromacroperforat = True: GoTo fi
      If vmicro <> "N" And vmicro <> "" Then If InStr(1, atrim(rst!rebmicromacro), "Micro" + atrim(vmicro)) = 0 Then potfermicromacroperforat = False
      If vmacro = "S" Then If InStr(1, atrim(rst!rebmicromacro), "Macro") = 0 Then potfermicromacroperforat = False
        Else: potfermicromacroperforat = False
   End If
fi:
   Set rst = Nothing
End Function
Sub passarbobinesentradanoutilitzades(numc As Double)
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("SELECT laminadores.comanda, bobineslam.* FROM laminadores INNER JOIN bobineslam ON laminadores.Id = bobineslam.controlid WHERE (laminadores.comanda=" + atrim(numc) + ")")
  While Not rst.EOF
     rst.Edit
     rst!utilitzadaabaixa = False
     rst.Update
     rst.MoveNext
  Wend
  Set rst = Nothing
  
End Sub
Sub comprovarcanutustallats(tubbase As Double)
    Dim rst As Recordset
    Set rst = dbtmpb.OpenRecordset("select * from canutusestandard where ample_canutu=" + passaradecimalpunt2(cadbl(amplebob)) + " and mida_canutu=" + passaradecimalpunt2(tubbase))
    If Not rst.EOF Then canutustallats = "Canutus ESTANDARD " + "(" + atrim(tubbase) + " cm)": Exit Sub
    Set rst = dbtmpb.OpenRecordset("select * from canutusjatallats where comanda=" + atrim(comanda))
    If rst.EOF Then
        canutustallats = "Atenció canutus NO TALLATS" + " (" + atrim(tubbase) + " cm)"
        MsgBox "Atenció els canutus per aquesta comanda encara NO ESTAN TALLATS." + "(" + atrim(tubbase) + " cm)", vbCritical, "Atenció"
      Else
         If cabool(rst!agafarstd) Then
              canutustallats = "Canutus aprox. a Standard" + " (" + atrim(tubbase) + " cm)"
             Else
               canutustallats = "Els canutus estan TALLATS" + " (" + atrim(tubbase) + " cm)"
         End If
   End If
End Sub
Sub carregar_client_ntintersialtres()
  Dim rstnt As Recordset
  Dim codicli As Double
  client.Caption = ""
  client.Tag = ""
  Set rstnt = dbtmp.OpenRecordset("select client,proximaseccio,cilindres,numerotintes from comandes where comanda=" + atrim(cadbl(comanda)))
  If Not rstnt.EOF Then
       ntintes = cadbl(rstnt!numerotintes)
       ncilindre = cadbl(rstnt!cilindres)
       framepantones.Tag = atrim(rstnt!proximaseccio)
       codicli = cadbl(rstnt!client)
       Set rstnt = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(codicli))
       If Not rstnt.EOF Then client.Caption = rstnt!nom: client.Tag = atrim(codicli)
  End If
  If client.Tag = "7" Then
            reixabobines.Columns(6).Caption = "Ample"
            reixabobines.Columns(6).DataField = "ample"
             Else
               reixabobines.Columns(6).Caption = "Pes N"
               reixabobines.Columns(6).DataField = "pesnet"
  End If
End Sub
Function nohiharebobinadora(producte As String) As Boolean
  Dim rstreb As Recordset
  nohiharebobinadora = True
  Set rstreb = dbtmp.OpenRecordset("select ruta from productes where codi='" + producte + "'")
   If Not rstreb.EOF Then
        If InStr(1, rstreb!ruta, "R") > 0 Then nohiharebobinadora = False
   End If
End Function
Sub gravar_pantones()
On Error GoTo fi
 If Not imppantones.Recordset.EOF Then
  escriure_ini "Rebobinadores", "lot1", imppantones.Recordset!lot1, "comandes.ini"
  escriure_ini "Rebobinadores", "lot2", imppantones.Recordset!lot2, "comandes.ini"
 End If
fi:
End Sub
Sub crear_pantones()
  r = " comanda "
  For i = 1 To 2
    r = r + ",tinta" + atrim(i) + "a "
  Next i
  Set rsttmp = dbtmp.OpenRecordset("select " + r + " from comandes where comanda=" + atrim(comanda))
  If Not rsttmp.EOF Then
   imppantones.Recordset.AddNew
   imppantones.Recordset!comanda = comanda
   imppantones.Recordset!pantone1 = "LIOFOL 7724"
   imppantones.Recordset!pantone2 = "LIOFOL 6020"
   imppantones.Recordset!lot1 = llegir_ini("Rebobinadores", "lot1", "comandes.ini")
   imppantones.Recordset!lot2 = llegir_ini("Rebobinadores", "lot2", "comandes.ini")
   'For i = 1 To 8
   '   imppantones.Recordset.Fields("pantone" + atrim(i)) = rsttmp.Fields("tinta" + atrim(i) + "a")
   'Next i
   imppantones.Recordset!comanda = comanda
   'imppantones.Recordset!pantone9 = "METOXI."
   'imppantones.Recordset!comanda = comanda
   'imppantones.Recordset!pantone10 = "R25."
   imppantones.Recordset.Update
  End If
  imppantones.Refresh
  imppantones.UpdateControls
End Sub
Function numerodepaletmesgran()
   Dim rst As Recordset
   r = bobines.Tag
   If r = "" Then r = "-1"
   Set rst = dbtmpb.OpenRecordset("select max(palet) as elgran from bobinesreb where controlid in (" + atrim(r) + ")")
   If Not rst.EOF Then
      numerodepaletmesgran = cadbl(rst!elgran)
   End If
   If numerodepaletmesgran = 0 Then numerodepaletmesgran = 1
End Function
Private Sub Command5_Click()
'  If Not clixes.Enabled Then Exit Sub
If Command5.Tag = "" Then Command5.Tag = Now
If Command5.Tag <> "" Then If DateDiff("s", Command5.Tag, Now) > 5 Then Command5.Tag = "": Exit Sub


Dim elgran As Double
Dim numb As Double
If numpalet < 1 Then numpalet = 1
If numpalet <> numerodepaletmesgran Then
   If numpalet <> (numerodepaletmesgran + 1) Then
    If MsgBox("No estàs col.locat al palet mes gran" + Chr(10) + "VOLS CONTINUAR AFEGINT LA BOBINA A AQUEST PALET IGUALMENT?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   End If
End If
i = 0
If IsDate(Rebobinadores.Recordset!datafi) And IsDate(Rebobinadores.Recordset!horafi) Then MsgBox "La linia de Funcionament actual està finalitzada. Canvia a la linia de Funcionament.": Exit Sub
If numop <> Rebobinadores.Recordset!operari1 Then MsgBox "No pots afegir bobines a una linia d'un altre operari": Exit Sub
If bobines.Recordset.EditMode > 0 Then bobines.Recordset.Update
If cadbl(bobinesxpalet) > 0 And bobines.Recordset.RecordCount >= cadbl(bobinesxpalet) Then
   If MsgBox("Ja has posat " + atrim(cadbl(bobinesxpalet)) + " bobines en aquest palet, vols posar una altra?", vbYesNo, "Atenció") = vbNo Then Exit Sub
End If
If bobinesent.Recordset.EOF And Not bobines.Recordset.EOF Then MsgBox "No hi han bobines d'entrada a la ultima bobina feta": Exit Sub
While barraestat.Caption = "Calculant els totals..."
  DoEvents
Wend
  dblots.Visible = False
  framepantones.Visible = False
  frameempalmes.Visible = False
  framebobentrada.Visible = False

  bobines.UpdateRecord
 If Rebobinadores.Recordset!tipus = "F" Then
     If cadbl(reixabobines.Columns(7).Text) = 0 And Not bobines.Recordset.EOF Then reixabobines.col = 7: reixabobines.SetFocus: MsgBox "Falten els metres a la bobina": Exit Sub
     If cadbl(reixabobines.Columns(5).Text) = 0 And Not bobines.Recordset.EOF Then reixabobines.col = 5: reixabobines.SetFocus: MsgBox "Falten els kilos a la bobina": Exit Sub
    'caluclar totals
     demanarcomandadebossesicanutus
     sa = ""
     If bobinesent.Recordset.RecordCount > 1 And cadbl(bandes) > 1 Then
      If MsgBox("Vols copiar kilos i bobines d'entrada?", vbYesNo, "Atenció") = vbYes Then
        sa = "copia kilos"
       Else: sa = ""
      End If
     End If
       If numbobinesnocorrelatiu Then
            If UCase(InputBoxEx("Els numeros de bobines no son correlatius. Reviseu per continuar la bobina " + r + vbNewLine + "ESCRIU LA CONTRASENYA PER CONTINUAR SENSE BOBINES CORRELATIVES.", "BOBINES NO CORRELATIVES", , , , , , SPassword)) = "INPNOCORRELATIVAS" Then Exit Sub
            vpermetbobinesnocorrelatives = True
       End If
       nova_bobina elgran
       'copiarbobentanterior IIf(sa = "copia kilos", True, False)
       'copiarbobentanterior True
       copiarbobinaentanterior elgran
       'possarnumerodepalet
       'possarnumbobent True
       bobines.Refresh
       If Not bobines.Recordset.EOF Then bobines.Recordset.MoveLast
       avisar_pesbobinaTEORIC cadbl(reixabobines.Columns("Kilos")), cadbl(reixabobines.Columns("Metres"))
       sa = ""
       If numbobinesnocorrelatiu Then MsgBox "Els numeros de bobines no son correlatius. Reviseu per continuar la bobina " + r: Exit Sub
     Else: MsgBox "Has d'escullir una linia de FUNCIONAMENT."
  End If
  reixabobines.col = 7
  If Not bobines.Recordset.EOF Then
    If bobines.Recordset!palet = 1 Then numpalet = 1
  End If
  ensenya_les_bobines
  colocarelsbotonsdelspalets
  ' calcular_totals
   calcular_totals
     'While barraestat.Caption = "Calculant els totals..."
     '  DoEvents
     'Wend
  Command5.Tag = ""
  
  'gravo la ultima comanda
  escriure_ini "Baixes", "ultimacomanda", comanda, "comandes.ini"
  
  
End Sub
Function copiarbobinaentanterior(bobant As Double)
    Dim rstbobent As Recordset
    Dim utili As Boolean
    Dim palet As Double
    Dim bobina As Double
    Dim idbobina As Double
    Dim numbent As String
    Set rstbobent = dbtmpb.OpenRecordset("SELECT bobinesreb.id FROM bobinesreb WHERE (((bobinesreb.controlid) in (" + bobines.Tag + ")) AND ((bobinesreb.numerodebobina)=" + atrim(bobant + 1) + "));")
    If Not rstbobent.EOF Then
         idbobina = cadbl(rstbobent!id)
       Else: Exit Function
    End If
    Set rstbobent = dbtmpb.OpenRecordset("SELECT bobinesreb.controlid,bobinesreb.id, bobinesreb.numerodebobina, bobinesentreb.palet, bobinesentreb.bobina,bobinesentreb.paletobobina  FROM bobinesentreb INNER JOIN bobinesreb ON bobinesentreb.id = bobinesreb.Id WHERE (((bobinesreb.controlid) in (" + bobines.Tag + ")) AND ((bobinesreb.numerodebobina)=" + atrim(bobant) + "));")
    numbent = ""
    While Not rstbobent.EOF
      palet = rstbobent!palet
      bobina = rstbobent!bobina
      carregar_bobinesdentrada "mirarsiutilitzada", , palet, bobina, ncomanda, utili, ncomanda2, IIf(proces.Tag = "invertit", True, False)
      If sa = "copia kilos" Then utili = False
      If Not utili Then
        dbtmpb.Execute "Insert into bobinesentreb (id,palet,bobina) values (" + passaradecimalpunt(idbobina) + "," + passaradecimalpunt(rstbobent!palet) + "," + passaradecimalpunt(rstbobent!bobina) + ") "
        If cadbl(rstbobent!bobina) > 0 Then
         If numbent <> "" Then numbent = numbent + "/"
         numbent = numbent + atrim(rstbobent!bobina)
        End If
      End If
      rstbobent.MoveNext
    Wend
    If idbobina > 0 Then
       dbtmpb.Execute "update bobinesreb set bobsent='" + numbent + "' where id=" + atrim(idbobina)
    End If
    
    Set rstbobent = Nothing
    wait 1
    r = numbent
End Function

Function numbobinesnocorrelatiu(Optional vbobinesErrorPesMetres As String) As Boolean
  Dim rstcp As Recordset
  Dim i As Integer
  Dim vnoavisar As Boolean
  
  If Rebobinadores.Recordset.EOF Then Exit Function
  
  If vpermetbobinesnocorrelatives Then numbobinesnocorrelatiu = False: Exit Function
  Set rstcp = dbtmpb.OpenRecordset("select * from bobinesreb where  controlid in(" + bobines.Tag + ") order by numerodebobina")
  If Not rstcp.EOF Then i = rstcp!numerodebobina
  numbobinesnocorrelatiu = False
  While Not rstcp.EOF And Not numbobinesnocorrelatiu
    If i <> rstcp!numerodebobina Then numbobinesnocorrelatiu = True: r = atrim(i) + " <> " + atrim(rstcp!numerodebobina)
    vnoavisar = True
    avisar_pesbobinaTEORIC cadbl(rstcp!kilos), cadbl(rstcp!metres), vnoavisar
    If vnoavisar = True Then vbobinesErrorPesMetres = vbobinesErrorPesMetres + " " + atrim(rstcp!numerodebobina)
    i = i + 1
    rstcp.MoveNext
  Wend
  
End Function

Sub crearunempalmerestomalo()
  empalmes.Recordset.AddNew
  empalmes.Recordset!id = bobines.Recordset!id
  empalmes.Recordset!observacions = "RESTO MALO"
  empalmes.Recordset.Update
End Sub
Sub possarnumerodepalet()

End Sub
Sub copiarbobentanterior(Optional nopreguntarfibob As Boolean)
 Dim rsttmp1 As Recordset
 Dim primer As Boolean
 Dim rsttmp2 As Recordset
 If cadbl(bobinesent.Tag) = 0 Then Exit Sub
 Set rsttmp1 = dbtmpb.OpenRecordset("select * from bobinesentreb where id=" + atrim(cadbl(bobinesent.Tag))) ' + " and paletobobina='B'")
 obrestocks
 primer = True
 While Not rsttmp1.EOF
  If (rsttmp1!paletobobina <> "P" And rsttmp1!paletobobina <> "B") Or sa = "copia kilos" Then
   If primer Then r = "carregartaulatmp": bobentrada_DblClick: primer = False: r = ""
   bobinesent.Recordset.AddNew
   bobinesent.Recordset!id = bobines.Recordset!id
   'bobinesent.Recordset!desb = rsttmp1!desb
   bobinesent.Recordset!palet = rsttmp1!palet
   bobinesent.Recordset!bobina = rsttmp1!bobina
   bobinesent.Recordset!paletobobina = rsttmp1!paletobobina
   
   bobinesent.Recordset.Update
   bobinesent.Refresh
  If Not nopreguntarfibob Then
   bobinesent.Recordset.FindFirst "palet=" + atrim(rsttmp1!palet) + " and bobina=" + atrim(rsttmp1!bobina)
   Set rsttmp2 = dbtmpb.OpenRecordset("select * from bobentradatmpreb" + atrim(nummaq) + " where " + "numpalet=" + atrim(rsttmp1!palet) + " and numbobent=" + atrim(rsttmp1!bobina))
   If rsttmp1!paletobobina = "p" Or rsttmp1!paletobobina = "b" Then
    bobinesent.Recordset.Edit
    If MsgBox("Ès final de la bobina? " + atrim(rsttmp1!palet) + "/" + atrim(rsttmp1!bobina), vbYesNo, "Bobina") = vbYes Then
      bobinesent.Recordset!paletobobina = UCase(bobinesent.Recordset!paletobobina)
      If UCase$(bobinesent.Recordset!paletobobina) = "P" Then
         dbstocks.Execute "update  bobines set utilitzadaabaixa=True where idpalet=" + bobentrada.Columns(0) + " and idbobina=" + bobentrada.Columns(1)
        Else:
           r = IIf(proces.Tag <> "L", "bobinesimp", "bobineslam")
           dbtmpb.Execute "update  " + r + " set utilitzadaabaixa=True where id=" + atrim(cadbl(rsttmp2!idbobina))
      End If
              
      Else
       bobinesent.Recordset!paletobobina = LCase(bobinesent.Recordset!paletobobina)
       If UCase$(bobinesent.Recordset!paletobobina) = "P" Then
          dbstocks.Execute "update  bobines set utilitzadaabaixa=False where idpalet=" + bobentrada.Columns(0) + " and idbobina=" + bobentrada.Columns(1)
        Else:
          r = IIf(proces.Tag <> "L", "bobinesimp", "bobineslam")
          dbtmpb.Execute "update  " + r + " set utilitzadaabaixa=False where id=" + atrim(cadbl(rsttmp2!idbobina))
       End If
       
    End If
    bobinesent.Recordset.Update
   End If
  End If
      
   
  End If
  rsttmp1.MoveNext
 Wend
 bobinesent.Refresh
 Set rsttmp1 = Nothing
 Set rsttmp2 = Nothing
 dbstocks.Close
End Sub
Sub nova_bobina(elgran As Double)
  Dim rstmp As Recordset
  Dim rsttmp2 As Recordset
  Dim col As Byte
  
  Dim metresant As Double
  Dim kilosant As Double
  Dim kilosantnet As Double
  Dim bobsent As String
  metresant = 0
  kilosantnet = 0
  kilosant = 0
  reixabobines.Tag = "afegint"
  'If Not bobines.Recordset.EOF Then
  ' If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
  ' bobines.Recordset.Update
   'metresant = cadbl(bobines.Recordset!metres)
   'kilosant = cadbl(bobines.Recordset!kilos)
  'End If
  r = bobines.Tag
   If r = "" Then r = "-1"
  Set rsttmp2 = dbtmpb.OpenRecordset("select id  from Rebobinadores where comanda=" + atrim(Rebobinadores.Recordset!comanda))
   Set rstmp = dbtmpb.OpenRecordset("select kilos,metres,bobsent,pesnet from bobinesreb where controlid in (" + atrim(r) + ") order by numerodebobina")
   If Not rstmp.EOF Then rstmp.MoveLast: bobsent = atrim(rstmp!bobsent): metresant = cadbl(rstmp!metres): kilosantnet = cadbl(rstmp!pesnet):: kilosant = cadbl(rstmp!kilos)
   Set rsttmp = Nothing
  elgran = 0
  
  While Not rsttmp2.EOF
   r = bobines.Tag
   If r = "" Then r = "-1"
   Set rstmp = dbtmpb.OpenRecordset("select max(numerodebobina) as elgran from bobinesreb where controlid in (" + atrim(r) + ")")
   If Not rstmp.EOF Then
      If cadbl(rstmp!elgran) > elgran Then elgran = cadbl(rstmp!elgran)
      If r = "-1" Then elgran = 0
      'If cadbl(rstmp!paletgran) > numpalet Then numpalet = cadbl(rstmp!paletgran)
   End If
   rsttmp2.MoveNext
  Wend
  Set rstmp = dbtmpb.OpenRecordset("select * from bobinesreb where controlid=" + atrim(Rebobinadores.Recordset!id) + " and numerodebobina=" + atrim(elgran))
  'bobines.Recordset.AddNew
  If sa = "copia kilos" Then kilos = kilosant
  If llegirpesbascula > 0 Then kilos = llegirpesbascula
  If pescanutu > 0 And tpescanutu.HelpContextID = 9999 Then pesnet = cadbl(kilos) - cadbl(pescanutu)
  afegir_bobina elgran + 1, atrim(Rebobinadores.Recordset!id), 0, cadbl(numpalet), Date, cadbl(amplebob), cadbl(espesor), cadbl(metresant), atrim(bobsent), cadbl(kilos), cadbl(pesnet), cadbl(numop)
  col = 0
  'bobinesent.Tag = atrim(rstmp!id)
 ' If Not rstmp.EOF Then
 '    bobines.Recordset!metres = rstmp!metres
 '    bobines.Recordset!kilos = rstmp!kilos
 '    col = 3 'escullo a la columne que es posa per defecte
 '  Else: col = 3
 ' End If
'  bobines.Recordset.Update
  'bobines.Refresh
  'bobines.Refresh
  'bobines.Recordset.MoveLast
  'reixabobines.Refresh
  DoEvents
  reixabobines.col = col
  If reixabobines.Enabled Then reixabobines.SetFocus
  Set rstmp = Nothing
  Set rstmp2 = Nothing
If reixabobines.Text = "0" Then reixabobines.SelLength = Len(reixabobines.Text)
reixabobines.Tag = ""
End Sub
Sub afegir_bobina(numbobent As Integer, idreb As Double, numempalmes As Byte, numpalet As Double, data As Date, ampleb As Double, espesor As Double, metres As Double, bobsent As String, kilos As Double, pesnet As Double, numop As Byte)
   Dim camps As String
   Dim valors As String
   Dim nump As Double
   nump = numpalet
   If nump < 1 Then numpalet = 1: nump = 1
   camps = "(numerodebobina,controlid,numempalmes,palet,datafab,ample,espessor,metres,bobsent,kilos,pesnet,operari1)"
   valors = "(" + passaradecimalpunt(numbobent) + "," + passaradecimalpunt(idreb) + "," + passaradecimalpunt(numempalmes) + "," + passaradecimalpunt(nump) + ",#" + Format(data, "yy/mm/dd") + "#," + passaradecimalpunt(ampleb) + "," + passaradecimalpunt(espesor) + "," + atrim(metres) + ",'" + atrim(bobsent) + "'," + passaradecimalpunt(kilos) + "," + passaradecimalpunt(pesnet) + "," + atrim(numop) + ")"
   dbtmpb.Execute ("insert into bobinesreb " + camps + " values " + valors)
End Sub
Private Sub DBGrid1_DblClick()
r = "numeric"
Set campcontrol = ActiveControl
teclattactil.Show
End Sub

Private Sub Command6_Click()
 
  dblots.Visible = False
  framepantones.Visible = False
  frameempalmes.Visible = False
  framebobentrada.Visible = False
 
 If MsgBox("Segur que vols BORRAR LA BOBINA Nº: " + atrim(bobines.Recordset!numerodebobina) + " ?", vbCritical + 4 + vbDefaultButton2, "Atenció") = vbYes Then
     If Not bobines.Recordset.EOF Then
       dbtmpb.Execute "delete * from bobinesentreb where id=" + atrim(cadbl(bobines.Recordset!id))
       bobines.Recordset.Delete
       bobines.Recordset.MoveLast
     End If
     On Error Resume Next
     bobines.Refresh
     reixabobines.Refresh
     bobines.Recordset.MoveLast
     wait 2
     calcular_totals
 End If
End Sub
Sub possar_valors_taula_reb(numcom As String, idbobina As Double, situacioet As String, Optional mostra As Boolean)
   Dim rstbob As Recordset
   Dim rstcom As Recordset
   Dim rstenvio As Recordset
   Dim idio As String
   Dim vespesortotal As Double
   Dim rst2 As Recordset
   Dim ruta As String
   If idbobina = 0 Then idbobina = 112239 ' apanyu per poder imprimir l'etiqueta interiro canutu de la primera bobina
   taula_tmp = "tmp_reb_empalmes" + atrim(nummaq)
   Set rstcom = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(numcom)))
   Set rstbob = dbtmpb.OpenRecordset("select * from bobinesreb where id=" + atrim(cadbl(idbobina)))
   Set rstenvio = dbtmp.OpenRecordset("select * from clients_envios where id=" + atrim(cadbl(rstcom!direnvio)))
   Set rst2 = dbtmp.OpenRecordset("select * from productes where codi='" + atrim(rstcom!producte) + "'")
   If Not rstenvio.EOF Then Set rstopcionset = dbtmp.OpenRecordset("select * from clients_etbobina where id_envio=" + atrim(rstenvio!id))
   If Not rstcom.EOF And Not rstbob.EOF And Not rstenvio.EOF And Not rst2.EOF Then
      If rstopcionset.EOF Then rstopcionset.AddNew: rstopcionset!id_envio = rstenvio!id: rstopcionset.Update: rstopcionset.MoveFirst
      rsttmp.AddNew
      rsttmp!idiomaclient = atrim(rstenvio!idioma)
      If rsttmp!idiomaclient = "" Then rsttmp!idiomaclient = "ES"
      'rsttmp!idiomaclient = "EN"
      rsttmp!etmostra = rstopcionset!etmostra
      If Not mostra Then rsttmp!etmostra = False
      rsttmp!comandacli = atrim(rstcom!comandaclient)
      rsttmp!refclient = IIf(atrim(rstcom!refclientdeclient) <> "", atrim(rstcom!refclientdeclient), atrim(rstcom!refclient))
      rsttmp!numcomanda = atrim(rstcom!comanda)
      rsttmp!texteimpresio = IIf(InStr(1, rst2!ruta, "I") > 0, IIf(atrim(rstcom!marcailinia) = "", atrim(rstcom!texteimpressio), atrim(rstcom!marcailinia)), "")
      rsttmp!codiproducte = ""
      rsttmp!material = desc_mat(rstcom!comanda, 1, vespesortotal) + desc_mat(cadbl(rstcom!linkcomanda1), 2, vespesortotal) + desc_mat(cadbl(rstcom!linkcomanda2), 3, vespesortotal)
      If (vespesortotal > 0) Then rsttmp!material = rsttmp!material + " (" + atrim(vespesortotal) + ")"
      rsttmp!dataproduccio = rstbob!datafab
      rsttmp!midarebobinat = cadbl(rstbob!ample) * 10
      rsttmp!desarroll = IIf(InStr(1, rst2!ruta, "I") > 0, rstcom!dessarroll, 0)
      rsttmp!numbob = rstbob!numerodebobina
      rsttmp!metresbob = rstbob!metres
      rsttmp!pesbobina = cadbl(rstbob!kilos)
      If cadbl(rstbob!pesnet) > 0 Then rsttmp!pesbobina = rstbob!pesnet * -1
      rsttmp!pescanutu = pescanutu
      rsttmp!peces = 0
      If InStr(1, rst2!ruta, "I") > 0 And (atrim(rstcom!continu) <> "S" And rstcom!dessarroll > 0) Then rsttmp!peces = Redondejar((cadbl(rsttmp!metresbob * 1000) / cadbl(rstcom!dessarroll)), 0)
      rsttmp!codibarres = rstcom!codibarras
      rsttmp!obsetiqueta = IIf(atrim(rstopcionset!obsetiq) <> "", atrim(rstopcionset!obsetiq), atrim(rstcom!obsetiq))
      rsttmp!situacioet = situacioet
      If atrim(rstopcionset!campcodibarres) <> "" Then
        rsttmp!campcodibarres = rstcom.Fields(rstopcionset!campcodibarres) ' s'ha de agafar el que possi a client
        rsttmp!tipuscodibarres = rstopcionset!tipuscodibarres ' s'ha de agafar el qu epossi a client
      End If
      rsttmp!inplacsasino = IIf(cadbl(rstenvio!emb_anonim) = 0, "INPLACSA", "")
      If atrim(rstcom!obspedgen2) <> "" Then
         rsttmp!inplacsasino = atrim(rstcom!obspedgen2)
         rsttmp!nomclient = buscarnomclient(cadbl(rstenvio!codi))
          Else: rsttmp!nomclient = atrim(rstenvio!nome)
      End If
      If rstopcionset!nomclientfacturacio Then rsttmp!nomclient = buscarnomclientfacturacio(rstcom!comanda)
      rsttmp!operari = cadbl(IIf(IIf(Not Rebobinadores.Recordset.EOF, Rebobinadores.Recordset!id, rstbob!controlid) <> rstbob!controlid, Rebobinadores.Recordset!operari1, rstbob!operari1))
      idio = IIf(rsttmp!idiomaclient <> "ES", "EN", rsttmp!idiomaclient)
      rsttmp!descproducte = atrim(rst2.Fields("descpelclient_" + idio))
       siesinterirobobinasensepes
      rsttmp.Update
        Else: If idbobina > 0 Then MsgBox "Hi ha hagut un error de client d'envio o de comanda. NO ES POT IMPRIMIR LA ETIQUETA": idbobina = 0
   End If
   If idbobina > 0 Then rsttmp.MoveFirst
End Sub
Function buscarnomclientfacturacio(vnumc As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("SELECT comandes_extres.comanda, Clients_codiscomptables.nomclient FROM comandes_extres LEFT JOIN Clients_codiscomptables ON comandes_extres.codicomptable = Clients_codiscomptables.codicomptable where comandes_extres.comanda=" + atrim(vnumc))
   If Not rst.EOF Then buscarnomclientfacturacio = atrim(rst!nomclient)
   Set rst = Nothing
End Function
Function buscarnomclient(numclient As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select nom from clientS where codi=" + atrim(numclient))
   If Not rst.EOF Then buscarnomclient = atrim(rst!nom)
   Set rst = Nothing
End Function
Sub siesinterirobobinasensepes()
  Dim rstmp
  If cadbl(etpesbascula) = 0 Then
      r = bobines.Tag
      If r = "" Then r = "-1"
      Set rstmp = dbtmpb.OpenRecordset("select max(numerodebobina) as elgran from bobinesreb where controlid in (" + atrim(r) + ")")

      rsttmp!numbob = cadbl(rstmp!elgran) + cadbl(bandes.Tag)
      'If reixabobines.Row = -1 Then
       rsttmp!metresbob = 0
       rsttmp!pesbobina = 0
       rsttmp!peces = 0
       rsttmp!dataproduccio = Date
       rsttmp!midarebobinat = cadbl(amplebob.Text) * 10
      'End If
      Set rstmp = Nothing
  End If
  
End Sub
Function desc_mat(numlot As String, ordre As Byte, vespesortotal As Double)
  Dim esp As Double
  If numlot = 0 Then Exit Function
  Set rsttmp3 = dbtmp.OpenRecordset("select materialex,colorex,espessor,mesuraesp,tubolam from comandes where comanda=" + atrim(numlot))
  
  If Not rsttmp3.EOF Then
      Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rsttmp3!mesuraesp)))
      If Not rsttmp2.EOF Then esp = micresmaterial(rsttmp2!descripcio, rsttmp3!espessor, rsttmp3!tubolam)
      Set rsttmp2 = dbtmp.OpenRecordset("select familia from materials where codi=" + atrim(cadbl(rsttmp3!materialex)))
      If Not rsttmp2.EOF Then
        Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(cadbl(rsttmp2!familia)))
        If Not rsttmp2.EOF Then desc_mat = atrim(rsttmp2!descripcio)
      End If
  End If
  If desc_mat <> "" Then
     'desc_mat = desc_mat + "(" + atrim(esp) + ")"
     If Len(desc_mat) > 4 Then desc_mat = Mid(desc_mat, 1, InStr(4, desc_mat, " "))
     vespesortotal = vespesortotal + esp
  End If
  If ordre > 1 And desc_mat <> "" Then desc_mat = "+" + desc_mat
End Function
Sub borraretiquetestemporals()
  On Error Resume Next
  Kill "c:\temp\ettmp*.*"
End Sub
Sub avis_et_noverificada()
  MsgBox "ATENCIÓ NO HI HA VERIFICACIÓ D'ETIQUETA PER IMPRIMIR" + Chr(13) + Chr(10) + "CONTACTA AMB L'OFICINA PER ACTIVAR L'ETIQUETA", vbCritical + vbOKOnly, "ATENCIÓ"
  Command7.BackColor = QBColor(12)
  If MsgBox("Vols verificar-la tu?", vbCritical + vbYesNo, "Atenció") = vbYes Then
      If InputBox("Entra la paraula INPLACSA per verificar la etiqueta", "Verificació d'Etiqueta") = "INPLACSA" Then
          r = App.Path + "\etokoperaris.txt"
          If Not existeix(r) Then
              Open r For Output As 1
             Else: Open r For Append As 1
          End If
          Print #1, Trim(Now) + "   Comanda: " + comanda.Text + " Operari: " + Trim(numop) + "-" + nomoperari
          Close 1
          dbtmp.Execute "update  comandes set etrebvistiplau=True where comanda=" + atrim(cadbl(comanda))
          MsgBox "RECORDA A ASSEGURAR QUE EL QUE SURT A L'ETIQUETA ES CORRECTE", vbCritical
          Command7.BackColor = &HFFFFFF
      End If
  End If
End Sub
Sub demanarescriureokperoficina()
   Dim resp As String
   While resp <> "OFICINA"
        resp = UCase(InputBox("Atenció aquesta mostra es per la Oficina no per Expedicions." + Chr(10) + "Escriu OFICINA per continuar.", "MOSTRA PER OFICINA"))
   Wend
End Sub
Function etiquetadeclientdeclient(numc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("select obspedgen2 from comandes where comanda=" + atrim(numc))
   If Not rst.EOF Then If atrim(rst!obspedgen2) <> "" Then etiquetadeclientdeclient = True
End Function
Function comprovar_sipesimetresescorrecte(vkg As Double, vmetres As Double) As Boolean
  Dim rst As Recordset
  Dim vpesxrmetre As Double
  comprovar_sipesimetresescorrecte = True
  Set rst = dbtmpb.OpenRecordset("SELECT rebobinadores.comanda, bobinesreb.kilos, bobinesreb.metres FROM rebobinadores RIGHT JOIN bobinesreb ON rebobinadores.Id = bobinesreb.controlid Where comanda = " + atrim(cadbl(comanda)))
  If Not rst.EOF Then
    If cadbl(rst!metres) > 0 Then
      vpesxrmetre = cadbl(rst!kilos) / cadbl(rst!metres)
      vpesteoric = (vmetres * vpesxrmetre)
      If (vkg > (vpesteoric * 1.05)) Or (vkg < (vpesteoric / 1.05)) Then comprovar_sipesimetresescorrecte = False
    End If
  End If
End Function
Private Sub Command7_Click()
Dim numb As Integer
Dim mtrs As Double
Dim rstco As Recordset
Dim inte As String
Static cont As Byte
'comprovo que no estigui imprimint ja
If cont = 3 Then cont = 0: Form1.Caption = "Baixes Comandes (Rebobinadores)"
If InStr(1, Form1.Caption, "Imprimint la bobina") <> 0 Then cont = cont + 1: Exit Sub
Form1.Caption = "Imprimint la bobina."
If cadbl(bandes) = 0 Then MsgBox "Atenció el numero de bandes està a zero." + Chr(10) + "Aixó pot afectar a l'impresió d'etiquetes i creació de bobines noves", vbInformation, "Atenció"
guardar_reg_bobines
If Not bobines.Recordset.EOF Then
 If Not comprovar_sipesimetresescorrecte(bobines.Recordset!kilos, bobines.Recordset!metres) Then
   If MsgBox("Els metres i kilos entrats no sembla correspondres correctament." + Chr(10) + "Vols cancelar l'impressió i rectificar-ho?", vbCritical + vbYesNo, "Error") = vbYes Then
      cont = 3
      Form1.Caption = "Baixes Comandes (Rebobinadores)"
      Exit Sub
   End If
 End If
End If
If cadbl(etmetresbob.Tag) > 0 And cadbl(bobines.Recordset!numerodebobina) = 1 Then
  If cadbl(bobines.Recordset!metres) > (cadbl(etmetresbob.Tag) * 1.1) Or cadbl(bobines.Recordset!metres) < (cadbl(etmetresbob.Tag) / 1.1) Then
       If MsgBox("Els metres que has fet d'aquesta bobina son diferents que els que demana el client" + Chr(10) + "ASSEGURA'T QUE SIGUI CORRECTE." + Chr(10) + "Vols cancelar l'impresió?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then GoTo fi
  End If
End If
Form1.Caption = "Imprimint la bobina.."
Set rstco = dbtmp.OpenRecordset("select etrebvistiplau from comandes where comanda=" + atrim(cadbl(comanda)))
If Not rstco.EOF Then
   If Not cabool(rstco!etrebvistiplau) Then avis_et_noverificada: Set rstco = Nothing: Exit Sub
    
End If
Command7.BackColor = &HFFFFFF
Form1.Caption = "Imprimint la bobina..."
borraretiquetestemporals
'MsgBox "Encara no es pot imprimir la bobina"
'Exit Sub
Form1.Caption = "Imprimint la bobina...."
If cadbl(etpesbascula) <> 0 Then
 If cadbl(bobines.Recordset!metres) = 0 Then
   mtrs = cadbl(InputBox("Entra els Metres de la bobina", "Atenció"))
   If mtrs = 0 Then Exit Sub
   If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
   bobines.Recordset!metres = cadbl(mtrs)
   bobines.Recordset.Update
 End If
   If bobines.Recordset.EditMode > 0 Then bobines.Recordset.Update
   bobines.UpdateRecord
   If Not bobines.Recordset.EOF Then numb = bobines.Recordset!numerodebobina
   Form1.Caption = "Imprimint la bobina....."
 'comprova si ha de fer la etiqueta de mostra
   If mostracli.Visible Then
       If mostracli.Value = 0 Then
         If MsgBox("Encara no has imprès la etiqueta per fer la mostra pel client." + Chr(13) + Chr(10) + "Vols fer-ho ara?", vbInformation + vbYesNo, "Atenció") = vbYes Then
           imprimir_bobina "Muestra Cli", True: mostracli.Value = 1: guarda_totals
           If etiquetadeclientdeclient(cadbl(comanda)) Then demanarescriureokperoficina
         End If
       End If
   End If
End If
Form1.Caption = "Imprimint la bobina......"
'imprimeix la et.bob interiors
  If Not rstopcionset.EOF Then
   inte = atrim(rstopcionset!etinteriorbob): inte = inte + "     "
   inte = atrim(Mid(inte, 1, 3))
   bandes.Tag = "1"
   If inte = "Un" Then
      imprimir_bobina "Int.Bob."
      If cadbl(etpesbascula) > 0 Then
        imprimir_bobina "Ext.Bobina"
      End If
   End If
   If inte = "Dos" Then
      If cadbl(etpesbascula) = 0 Then wait 1
      If cadbl(etpesbascula) = 0 Then
        If cadbl(bandes) > 5 Then Exit Sub
        If cadbl(bandes) = 0 Then MsgBox "Ojo... que les bandes estan a zero"
        For x = 1 To cadbl(bandes)
         bandes.Tag = atrim(x)
         imprimir_bobina "Int.Bob.D"
         wait (1)
         imprimir_bobina "Int.Bob.I"
         wait (1)
        Next x
         'imprimeix la et.bob exterior
           Else: imprimir_bobina "Ext.Bobina"
      End If
   End If
   If inte = "" Then imprimir_bobina "Ext.Bobina"
  End If
  Form1.Caption = "Imprimint la bobina......."
fi:
  Form1.Caption = "Baixes Comandes (Rebobinadores)"

End Sub
Sub comprovarsitocaverificacio()
  Dim metres As Double
  If cadbl(bandes) < 1 Then Exit Sub
  metres = cadbl(cadbl(tmetres) / cadbl(bandes)) + (cadbl(rsttmp!metresbob) * cadbl(bandes))
  'metres = metres - (Int(metres / 7000) * 7000)
  If (metres - contadorverificacio) > 7000 Then contadorverificacio = metres * -1
End Sub
Function texte_impressio(vnumc As Double) As String
   Dim rst As Recordset
   Dim vnumtreball As Double
   Set rst = dbtmp.OpenRecordset("select producte,numtreball from comandes where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
        vnumtreball = rst!numtreball
        Set rst = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(rst!producte) + "'")
        If InStr(1, rst!ruta, "I") > 0 Then
            Set rst = dbtmp.OpenRecordset("select * from clixes where id_treball=" + atrim(vnumtreball))
            If Not rst.EOF Then texte_impressio = atrim(rst!marca) + " - " + atrim(rst!linia)
        End If
   End If
   Set rst = Nothing
End Function
Sub imprimir_etiqueta_inplacsa(vnumc As Double, vnumbobina As Long)
  Dim vliniaZPL As String
  Dim v1 As String
  Dim v2 As String
  Dim vtexteimpressio As String
  Dim vportZebra As String
  vtexteimpressio = texte_impressio(vnumc)
  vportZebra = llegir_ini("Baixes", "portetiquetareb", fitxerini)
  v1 = atrim(Format(vnumc, "#,##0") + "/" + atrim(vnumbobina))
  v2 = atrim(vnumc) + "/" + atrim(vnumbobina)
  vliniaZPL = "^XA"
  vliniaZPL = vliniaZPL + "^PW866^LL551^POI" + vbNewLine
  vliniaZPL = vliniaZPL + "^FO15,30^A0N,60,60^FDINPLACSA^FS" + vbNewLine
  vliniaZPL = vliniaZPL + "^FO0,120^A0N,90,90^FB866,1,0,C^FD#v1#^FS" + vbNewLine
  vliniaZPL = vliniaZPL + "^BY3,3,200^FO200,240^BCN,200,N,N,N^FD#v2#^FS" + vbNewLine
  vliniaZPL = vliniaZPL + "^FO0,450^A0N,40,40^FB866,1,0,C^FD" + vtexteimpressio + "^FS"
  vliniaZPL = vliniaZPL + "^XZ"
  vliniaZPL = substituirtot(vliniaZPL, "#v1#", v1)
  vliniaZPL = substituirtot(vliniaZPL, "#v2#", v2)
  Open vportZebra For Binary As 1
  Put 1, , vliniaZPL
  Close 1
End Sub
Function mirarproximaseccio(vnumc As Double) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select producte from comandes where comanda=" + atrim(vnumc))
  Set rst = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(rst!producte) + "'")
  mirarproximaseccio = Mid(rst!ruta + " ", InStr(1, rst!ruta, "R") + 1, 1)
  Set rst = Nothing
End Function

Sub imprimir_bobina(aon As String, Optional mostra As Boolean)
 Dim idbob As Double
 Dim rstbob As Recordset
 Static ultimabobinaimpresa As Double
 If mirarproximaseccio(cadbl(comanda.Text)) = "S" Then
      If aon = "Int.Bob." Then
           imprimir_etiqueta_inplacsa comanda.Text, bobines.Recordset!numerodebobina: Exit Sub
      End If
 End If
 Set rstbob = dbtmpb.OpenRecordset("select * from bobinesreb")
 taula_tmp = "tmp_reb_empalmes" + atrim(nummaq)
 Set rsttmp = Nothing: r = ""
 crear_taula_rev_empalmes
 Set rsttmp = dbtmpb.OpenRecordset(taula_tmp)
 idbob = cadbl(rstbob!id)
 Set rstbob = Nothing
 On Error Resume Next
 idbob = bobines.Recordset!id
 On Error GoTo 0
 possar_valors_taula_reb comanda.Text, idbob, aon, mostra
 If rsttmp.EOF Then Exit Sub

'etiqueta de verificació
'If ultimabobinaimpresa <> rsttmp!numbob Then
'    comprovarsitocaverificacio
'    If contadorverificacio < 1 Then
'      If cadbl(bandes) > 0 Then
'       For i = 1 To cadbl(bandes)
'        preparar_etiqueta_verificacio cadbl(rsttmp!numcomanda), cadbl(rsttmp!operari), cadbl(rsttmp!numbob)
'        imprimir_etiqueta_zebra True
'        wait 1
'       Next i
'      End If
'      contadorverificacio = contadorverificacio * -1
'    End If
'End If
 
preparar_etiqueta_zebra
If aon = "sense imprimir" Then Exit Sub
imprimir_etiqueta_zebra
If cadbl(etpesbascula) > 0 And etiquetesean13 And InStr(1, aon, "Int.Bob") = 0 Then
  preparar_etiqueta_ean13_zebra
  imprimir_etiqueta_zebra True
End If
'si faig etiqueta exterior comprovao si toca imprimir codidebarres extra
If aon = "Ext.Bobina" Then
    If cadbl(client.Tag) = 6603 Then  'si es videcart faig el codidebarres extra
        preparar_etiqueta_videcart_zebra
        imprimir_etiqueta_zebra True
    End If
End If
ultimabobinaimpresa = rsttmp!numbob
Set rsttmp = Nothing

End Sub
Sub preparar_etiqueta_videcart_zebra()
   Dim v As String
   Dim ref As String
   Dim numvidecart As String
   'If existeix("c:\temp\etiquetareb.prn") Then Kill "c:\temp\etiquetareb.prn"
   Open llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetarebean128.prn" For Input As #1
   numvidecart = generarnumvidecart(rsttmp)
   linia.Text = Input(LOF(1), #1)
   Close #1
   With rsttmp
   substituir "Linia1", numvidecart
   substituir "1111111111111111111111111111", numvidecart
   End With
End Sub

Function generarnumvidecart(rst As Recordset) As String
   Dim pesbobina As Double
   pesbobina = IIf(rst!pesbobina < 0, rst!pesbobina * -1, rst!pesbobina)
   generarnumvidecart = "0591" + codicomandavidecart(atrim(rst!refclient)) + Format(rst!numbob, "0000000000") + Format(Redondejar(pesbobina, 0), "0000")
End Function
Function codicomandavidecart(vcomandacli As String) As String
   Dim i As Byte
   i = 1
   codicomandavidecart = "0000000000"
   While IsNumeric(Mid(vcomandacli, i, 1))
     i = i + 1
     If i > Len(vcomandacli) Then GoTo cont
   Wend
cont:
   If i < 2 Then GoTo fi
   codicomandavidecart = Format(Mid(vcomandacli, 1, i - 1), "0000000000")
fi:
End Function
Function treurecaracters(refclient As String) As String
   Dim ref As String
   ref = refclient
   For i = 1 To Len(refclient)
     If Not IsNumeric(Mid(refclient, i, 1)) Then substituircaracter ref, Mid(refclient, i, 1), ""
   Next i
   treurecaracters = ref
End Function
Function emplena12zeros(codi As String) As String
   emplena12zeros = String(12 - Len(codi), "0") + codi
End Function
Sub preparar_etiqueta_verificacio(numc As Double, op As Byte, numbob As Double)
   Dim v As String
   Dim ref As String
   Dim vnumbobentrada As String
   'If existeix("c:\temp\etiquetareb.prn") Then Kill "c:\temp\etiquetareb.prn"
   Open llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetarebverificacio.prn" For Input As #1
   linia.Text = Input(LOF(1), #1)
   Close #1
   With rsttmp
   vnumbobentrada = 0
   If Not bobinesent.Recordset.EOF Then bobinesent.Recordset.MoveLast: vnumbobentrada = atrim(bobinesent.Recordset!bobina)
   substituir "#linia1.2#", "B.Ent: " + atrim(vnumbobentrada)
   substituir "#linia1#", "VERIFICACION CALIDAD: " + atrim(numc)
   substituir "#linia2#", "Reb-" + atrim(nummaq) + " Op: " + atrim(op) + " NºBob: " + atrim(numbob) + " Fecha: " + Format(Now, "dd/mm/yy")
   If Not vperforat Then substituir "Verificar perforado.", "": substituir "X11,343,8,41,371", ""
   End With
End Sub
Sub preparar_etiqueta_ean13_zebra()
   Dim v As String
   Dim ref As String
   'If existeix("c:\temp\etiquetareb.prn") Then Kill "c:\temp\etiquetareb.prn"
   Open llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetarebean13.prn" For Input As #1
   
   linia.Text = Input(LOF(1), #1)
   Close #1
   With rsttmp
   ref = treurecaracters(!refclient)
   substituir "Linia1", "PRODUCTION Nº: " + atrim(!numcomanda)
   substituir "Linia2", "REFERENCE: " + atrim(ref)
   substituir "111111111111", emplena12zeros(!numcomanda)
   substituir "111111111111", emplena12zeros(ref)
   End With
End Sub

Sub substituircaracter(cadena As String, buscar As String, canviar As String)
   comença = InStr(1, cadena, buscar)
   If comença < 1 Then Exit Sub
   comença = comença - 1
   acaba = comença + Len(buscar) + 1
   cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   'MsgBox linia
End Sub
Function buscar_refinp(vnumc As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("Select refinplacsa from comandes_extres where comanda=" + atrim(vnumc))
   If Not rst.EOF Then buscar_refinp = atrim(rst!refinplacsa)
   Set rst = Nothing
End Function
Sub preparar_etiqueta_zebra()
   Dim v As String
   'If existeix("c:\temp\etiquetareb.prn") Then Kill "c:\temp\etiquetareb.prn"
   Open llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetareb1.prn" For Input As #1
   
   linia.Text = Input(LOF(1), #1)
   Close #1
   With rsttmp
   idiomaclient = !idiomaclient
   possar_codidebarres
   substituir "P1", ""
   substituir "#linia1.1#", sitoca(!inplacsasino, "inplacsasino")
   substituir "#linia1.2#", sitoca(retallar(!nomclient, 22), "nomclient")
   substituir "#linia2.1#", sitoca(retallar(idioma("Producto: ") + atrim(!descproducte), 40), "descproducte")
   substituir "#linia3.1#", retallar(sitoca(idioma("RefC:") + !refclient, "refclient") + " " + sitoca(idioma("PedC:") + !comandacli, "comandacli"), 39)
   substituir "#linia17.1#", retallar("RefInp:" + atrim(buscar_refinp(!numcomanda)), 39)
   substituir "#linia4.1#", sitoca(retallar(!material, 40), "material")
   substituir "#linia5.1#", sitoca(retallar(!texteimpresio, 50), "texteimpresio")
   substituir "#linia6.1#", sitoca(idioma("Ancho:") + atrim(!midarebobinat) + " m/m ", "midarebobinat") + sitoca(idioma("Desar:") + atrim(!desarroll) + " m/m", "desarroll")
   
   substituir "#linia7.1#", sitoca(retallar(IIf(!obsetiqueta <> "", idioma("Obs.Et:") + !obsetiqueta, ""), 50), "obsetiqueta")
   substituir "#linia15.1#", sitoca(Format(!dataproduccio, "dd/mm/yy"), "dataproduccio") + " " + sitoca(idioma("Op:") + !operari, "operari") + " " + sitoca(idioma("Lote: ") + Format(!numcomanda, "#,##0"), "numcomanda") + " " + idioma(!situacioet)
   
   If Not !etmostra Then
      substituir "#linia14.1#", sitoca(IIf(!peces > 0, idioma("Unidades:") + atrim(Format(!peces, "#,##0")), ""), "peces")
      If !pesbobina >= 0 Then
         substituir "#linia9.1#", sitoca(idioma("Peso:") + atrim(Format(!pesbobina, "#,##0.0")) + " Kg", "pesbobina")
         substituir "#linia16.1#", ""
        Else:
          substituir "#linia9.1#", sitoca(idioma("Neto:") + atrim(Format(!pesbobina * -1, "#,##0.0")) + " Kg", "pesbobina")
          substituir "#linia16.1#", sitoca(idioma("Mandril:") + atrim(Format(!pescanutu, "#,##0.0")) + " Kg", "pescanutu")
      End If
      substituir "#linia10.1#", sitoca(idioma("Long:") + atrim(Format(!metresbob, "#,##0")) + " Mts", "metresbob")
      substituir "#linia8.1#", sitoca(idioma("NºBob:") + atrim(!numbob), "numbob")
      
     Else
       substituir "#linia8.1#", idioma("NºBob:") + atrim(1)
       linia = linia + "A10,300,0,5,1,1,N," + Chr$(34) + idioma("ETIQUETA") + Chr$(34) & vbCrLf
       linia = linia + "A10,355,0,5,1,1,N," + Chr$(34) + idioma("MUESTRA") + Chr$(34) & vbCrLf
   End If
   End With
   'tradueixo els textes de apte per consum
   If Not cabool(rstopcionset!noimprimirapteusalimentari) Then
      substituir "Apto para uso alimentario.", idioma("Apto para uso alimentario.")
       Else: substituir "Apto para uso alimentario.", idioma("          ")
   End If
   substituir "Proteger de altas y bajas temperaturas.", idioma("Proteger de altas y bajas temperaturas.")
   substituir "Proteger de la luz solar.", idioma("Proteger de la luz solar.")
   substituir "Recomendable utilizar antes de 9 meses.", idioma("Recomendable utilizar antes de 9 meses.")
   
   'TREC LES LINIES QUE NO FAI SERVIR
    
   substituir "#linia1.1#", ""
   substituir "#linia1.2#", ""
   substituir "#linia2.1#", ""
   substituir "#linia3.1#", ""
   substituir "#linia4.1#", ""
   substituir "#linia5.1#", ""
   substituir "#linia6.1#", ""
   substituir "#linia7.1#", ""
   substituir "#linia17.1#", ""
   substituir "#linia8.1#", ""
   substituir "#linia2.2#", ""
   substituir "#linia9.1#", ""
   substituir "#linia10.1#", ""
   substituir "#linia14.1#", ""
End Sub
Sub possar_codidebarres()
   Dim codib As String
   Dim numc As String
  If atrim(rsttmp!campcodibarres) <> "" Then
      If rsttmp!tipuscodibarres = "Ean-13" Then
          codib = "E30"
           Else
             If rsttmp!tipuscodibarres = "Ean-8" Then
                 codib = "E80"
                  Else
                    If rsttmp!tipuscodibarres = "Ean-128A" Then codib = "1A"
             End If
      End If
      numc = rsttmp!campcodibarres
      Else: codib = "": numc = ""
  End If
  substituir "#EAN#", codib
  substituir "1234567890128", numc
End Sub
Function idioma(txt As String) As String
 Dim v As String
 Dim fitxeridioma As String
 
 If idiomaclient = "" Then idiomaclient = "ES"
 fitxeridioma = llegir_ini("General", "rutallistats", "comandes.ini") + idiomaclient + "_etiquetareb.txt"
 f = llegir_ini("Idioma", txt, fitxeridioma)
 If f = "{[}]" Then escriure_ini "Idioma", txt, txt, fitxeridioma: f = txt
 idioma = f
End Function
Function sitoca(txt As String, camp As String) As String
  sitoca = ""
  If camp = "pescanutu" Then
    If rstopcionset.Fields("sivull_canutu") Then sitoca = txt
     GoTo fi
  End If
  If Not rstopcionset.Fields(camp) Then sitoca = txt
  If atrim(rsttmp.Fields(camp)) = "" Then sitoca = ""
  If rsttmp.Fields(camp).Type = 7 Then
     If cadbl(rsttmp.Fields(camp)) = 0 Then sitoca = ""
  End If
fi:
End Function
Function retallar(txt As String, tamany As Integer) As String
   retallar = Mid(txt, 1, tamany)
End Function

Sub substituir(buscar As String, canviar As String)
   comença = InStr(1, linia, buscar) - 1
   If comença < 1 Then Exit Sub
   acaba = comença + Len(buscar) + 1
   linia = Mid(linia, 1, comença) + canviar + Mid(linia, acaba)
   'MsgBox linia
End Sub

Sub imprimir_etiqueta_zebra(Optional sensegrafic As Boolean)
  Dim nomord As String * 255
  Dim ettmp As String
  Static contador As Byte
  Dim impresora As String
  If contador = 200 Then contador = 1
  ettmp = "ettmp" + atrim(contador) + ".prn"
  contador = contador + 1
  GetComputerName nomord, 255
  If existeix("c:\temp\etiquetareb.prn") Then Kill "c:\temp\etiquetareb.prn"
  Open "c:\temp\etiquetareb.prn" For Output As #2
  Print #2, linia.Text
  Close #2
  Copiar_Fitxer "c:\temp\etiquetareb.prn", "c:\temp\" + ettmp
  'linia = ""
  nomord = Mid(nomord, 1, InStr(1, nomord, Chr$(0)) - 1)
  impresora = "\\" + atrim(nomord) + "\zebra"
  r = llegir_ini("Baixes", "portetiquetareb", fitxerini)
  If r <> impresora And r <> "{[}]" Then
       impresora = r
      Else: escriure_ini "Baixes", "portetiquetareb", impresora, fitxerini
  End If
  ShellandWait "c:\windows\system32\cmd.exe /c type c:\temp\" + ettmp + ">" + impresora, 5
  If Not sensegrafic Then ShellandWait "c:\windows\system32\cmd.exe /c type " + llegir_ini("General", "rutallistats", "comandes.ini") + "graficetareb1.prn>" + impresora, 5
  
End Sub
Sub possar_valors_taula_lam_empalmes()
 Dim rs As Recordset
 Dim rs2 As Recordset
 Dim bobe As String
 Dim nample As Double
 
 obrestocks
 If Not bobinesent.Recordset.EOF Then
   bobinesent.Recordset.MoveFirst
 End If
 Set rs = dbtmpb.OpenRecordset("tmp_lam_empalmes")
 r = atrim(capcalera.capcalera.Recordset!matdesb1) + " + " + atrim(capcalera.capcalera.Recordset!matdesb2)
 
 rs.AddNew
 rs!numlot1 = comanda.Text
 rs!numlot2 = linkcomanda.Text
 rs!numbobsort = cadbl(bobines.Recordset!numerodebobina)
 rs!numop = cadbl(bobines.Recordset!operari1)
 rs!numop2 = cadbl(bobines.Recordset!operari2)
 rs!datafab = Format(bobines.Recordset!datafab, "dd/mm/yy")
 rs!client = client.Caption
 rs!texteimpressio = texteimpresio
 rs!refclient = refclient
 rs!observacio = bobines.Recordset!observacio
 rs!comandaclient = comandaclient
 rs!material = r
 For i = 1 To 4
   If Not bobinesent.Recordset.EOF Then
      rs.Fields("numbobent" + atrim(i)) = atrim(bobinesent.Recordset!paletobobina) + atrim(bobinesent.Recordset!palet) + "/" + atrim(bobinesent.Recordset!bobina)
      'aprofito per buscar lamplada del palet
      If UCase(bobinesent.Recordset!paletobobina) = "P" Then
          Set rststocks = dbstocks.OpenRecordset("select ample,plegat, solapa from palets where idpalet=" + atrim(cadbl(bobinesent.Recordset!palet)))
          If Not rststocks.EOF Then If cadbl(rststocks!ample) > nample Then nample = rststocks!ample
            Else:
              am = cadbl(buscarbobina(atrim(cadbl(bobinesent.Recordset!palet)), atrim(cadbl(bobinesent.Recordset!bobina)), "impressores", "ample"))
              If am > nample Then nample = am
      End If
      bobinesent.Recordset.MoveNext
   End If
 Next i
 
 rs!ample = nample
 On Error Resume Next
 rs!plegat = cadbl(rststocks!plegat)
 rs!solapa = cadbl(rststocks!solapa)
 On Error GoTo 0
 rs!metres = cadbl(bobines.Recordset!metres)
 rs!kilos = cadbl(bobines.Recordset!kilos)
 rs!espessor = micrescomanda
 'actualitzo les dades de la bobina
    bobines.Recordset.Edit
    bobines.Recordset!ample = rs!ample
    bobines.Recordset!espessor = rs!espessor
    bobines.Recordset.Update
 'fins aqui actualitzo
 'llistat.Formulas(0) = "mesuraesp='(" + mesuraespcomanda + ")'"
 llistat.Formulas(0) = "mesuraesp='(micres)'"
 rs!codibarres = codibarras
 empalmes.RecordSource = "select * from lamempalmes where id=" + atrim(bobines.Recordset!id)
 empalmes.Refresh
 If Not empalmes.Recordset.EOF Then
  empalmes.Recordset.MoveFirst
  i = 0
  While Not empalmes.Recordset.EOF And i < 5
    rs.Fields("empalme" + atrim(i + 1)) = empalmes.Recordset!observacions
    rs.Fields("mtrs" + atrim(i + 1)) = empalmes.Recordset!metres
    rs.Fields("dist" + atrim(i + 1)) = empalmes.Recordset!distancia
    i = i + 1
    empalmes.Recordset.MoveNext
  Wend
 End If
 rs.Update
 Set rs = Nothing
 Set rs2 = Nothing
 Set dbstocks = Nothing
End Sub
Function buscarbobina(comanda As String, bobina As Integer, seccio As String, Optional camp As String) As Double
  Dim rstbob As Recordset
  Dim rstsec As Recordset
  If atrim(camp) = "" Then camp = "id"
  secciobob = "bobines" + Mid(seccio, 1, 3)
  Set rstsec = dbtmpb.OpenRecordset("select id,tipus from " + seccio + " where comanda=" + comanda)
  While Not rstsec.EOF
    Set rstbob = dbtmpb.OpenRecordset("select id,numerodebobina,controlid " + "," + camp + " from " + secciobob + " where controlid=" + atrim(cadbl(rstsec!id)) + " and numerodebobina=" + atrim(cadbl(bobina)))
    If Not rstbob.EOF Then rstsec.MoveLast: buscarbobina = cadbl(rstbob.Fields(camp))
    rstsec.MoveNext
  Wend
End Function
Sub crear_taula_rev_empalmes()
  Dim taula_tmp As String
  Dim camps(100, 2) As String
   Dim td As TableDef, fld As Field
   Dim db As Database
  Dim l As Integer
  Dim k As Integer
  taula_tmp = "tmp_reb_empalmes" + atrim(nummaq)
  If Not existeixlataula(taula_tmp) Then
        i = 1
        camps(i, 1) = "comandacli": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "pesbobina": camps(i, 2) = "double": i = i + 1
        camps(i, 1) = "refclient": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "numcomanda": camps(i, 2) = "double": i = i + 1
        camps(i, 1) = "texteimpresio": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "codiproducte": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "dataproduccio": camps(i, 2) = "date": i = i + 1
        camps(i, 1) = "material": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "midarebobinat": camps(i, 2) = "double": i = i + 1
        camps(i, 1) = "desarroll": camps(i, 2) = "double": i = i + 1
        camps(i, 1) = "peces": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "numbob": camps(i, 2) = "integer": i = i + 1
        camps(i, 1) = "metresbob": camps(i, 2) = "double": i = i + 1
        camps(i, 1) = "codibarres": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "obsetiqueta": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "situacioet": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "inplacsasino": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "nomclient": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "descproducte": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "operari": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "campcodibarres": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "tipuscodibarres": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "etmostra": camps(i, 2) = "bit": i = i + 1
        camps(i, 1) = "idiomaclient": camps(i, 2) = "string": i = i + 1
        camps(i, 1) = "pescanutu": camps(i, 2) = "double": i = i + 1
        dbtmpb.Execute ("create table " + taula_tmp + "(p byte)")
        For i = 1 To 100
          If camps(i, 1) <> "" Then
             dbtmpb.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
              Else: i = 1000
          End If
        Next i
        'ample double,plegat double,solapa double,espessor double,metres double,kilos double)"
        'dbtmpb.Execute ("create table tmp_lam_empalmes (" + camps + camps2 + camps3 + camps4) + ")"
         Else
              dbtmpb.Execute "delete * from " + taula_tmp
              Exit Sub
   End If
        'passo tots els camps de texte a allowzerolength
On Error Resume Next
    Set db = dbtmpb
    For l = 0 To db.TableDefs.Count - 1
       Set td = db(l)
       If td.Name = taula_tmp Then
        For k = 0 To td.Fields.Count - 1
          Set fld = td(k)
          If (fld.Type = 10) And Not _
            fld.AllowZeroLength Then
             fld.AllowZeroLength = True
          End If
        Next k
       End If
    Next l
        
End Sub

Function cabool(valor As Variant) As Boolean
  If IsNull(valor) Then valor = False
  If valor Then
    cabool = True
   Else: cabool = False
  End If
End Function


Sub emplenar_capcalera_imp(rsttemp As Recordset)
 Dim rst As Recordset
 
 Set rst = dbtmpb.OpenRecordset("select * from Rebobinadorestot where comanda=" + comanda.Text)
 If Not rst.EOF Then
 
   rsttemp!amplebob = atrim(rst!amplebob)
   rsttemp!espesor = atrim(rst!espesor)
   rsttemp!bandesbones = atrim(rst!simulteneitat)
   rsttemp!ampleref = atrim(rst!ampleref)
   rsttemp!bandesmerma = atrim(rst!bandesmerma)
   rsttemp!amplemerma = atrim(rst!amplemerma)
   rsttemp!lotbosses = atrim(rst!comandabosses1) + IIf(atrim(rst!comandabosses2) <> "", " - " + atrim(rst!comandabosses2), "")
   rsttemp!lotcanutus = atrim(rst!comandacanutus1) + IIf(atrim(rst!comandacanutus1) <> "", " - " + atrim(comandacanutus2), "")
 End If
 Set rst = Nothing
End Sub



Sub imprimir_fulla(Optional nomllistat As String)
  Dim mtrsparcialanteriors As Double
  Dim rst As Recordset
  Dim rsttemp As Recordset
   Dim rsttmp2 As Recordset
   Dim nb As String
   Dim np As Double
   Dim linia As Double
   Dim rsttmpbob As Recordset
   Dim canvicam As String
   Dim vdesarroll As Double
   Dim vcarpetadesti As String
   Dim vnommaquina As String
   
   
   If nomllistat = "" Then nomllistat = "baixesrebobinadora.rpt"
   If cadbl(kiloshora) = 0 And cadbl(tmetres) > 0 Then
     If MsgBox("Error els metres minuts estan a zero." + Chr(10) + "Vols imprimir igualment aquesta Fulla?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   End If
   Form1.Caption = "Imprimint..."
   nample = 0
   vdesarroll = 0
   Set rst = dbtmp.OpenRecordset("SELECT comandes.dessarroll, productes.ruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi where comanda=" + atrim(comanda.Text))
   If Not rst.EOF Then If InStr(1, rst!ruta, "I") > 0 Then vdesarroll = cadbl(rst!dessarroll)
 'carregar_client_ntintersialtres
   'panelimprimir.Visible = True
'panelimprimir.Top = Frame3.Top
  crear_taula_laminadora_baixa
  obrestocks
  Set rsttemp = dbtemp.OpenRecordset("tmp_reb_baixa")
  imppantones.Refresh
  rsttemp.AddNew

  ' busco l'ample
   'ample_palet
  '-----------
  

  
  With rsttemp
  !comanda = atrim(comanda.Text)
  '!client = atrim(client.Caption)
  !client = client.ToolTipText
 ' !firmat = atrim(firmat.Caption)
  '!nomfirmat = possarnomfirmat
  '!tintersrentats = cadbl(trentats)
  '!portaclixers = cadbl(pclixers)
  '!canvienfilada = atrim(canvienfilada)
  '!numtintes = cadbl(ntintes)
  '!cilindre = cadbl(ncilindre)
  !comandaacavada = IIf(comandaacavada.Value, 1, 0)
  
  'relleus i descans
   i = 1
   Set rstdr = dbtmpb.OpenRecordset("select * from controldescansrelleu where seccio='" + atrim(lletraseccio) + "' and comanda=" + atrim(ncomanda) + " and comandafi=" + atrim(ncomanda))
   While Not rstdr.EOF And i < 4
        .Fields("prepdr_data" + Trim(i)) = Format(atrim(rstdr!datainici), "dd/mm/yy")
        .Fields("prepdr_op" + Trim(i)) = cadbl(rstdr!operari)
        .Fields("prepdr_de" + Trim(i)) = Format(atrim(rstdr!horainici), "hh:nn")
        .Fields("prepdr_fins" + Trim(i)) = Format(atrim(rstdr!horafi), "hh:nn")
        .Fields("prepdr_observacions" + Trim(i)) = atrim(cadbl(rstdr!hores)) + " Hores de " + atrim(rstdr!tipus)
         i = i + 1
        rstdr.MoveNext
   Wend
  
  
  'prep clixe
  emplenar_capcalera_imp rsttemp
  Set rst = dbtmpb.OpenRecordset("select id,operari1,datainici,horainici,datafi,horafi,observacio from Rebobinadores where comanda=" + comanda.Text + " and tipus='C' order by datainici,horainici")
  If Not rst.EOF Then rst.MoveLast
  If Not rst.EOF Then
   For i = 1 To 4
     rst.MovePrevious
     If rst.BOF Then rst.MoveNext: i = 10
   Next i
  End If
  i = 1
  If rst.EOF Then Exit Sub
  While Not rst.EOF
    .Fields("prepmaquina_data" + Trim(i)) = Format(atrim(rst!datainici), "dd/mm/yy")
    .Fields("prepmaquina_op" + Trim(i)) = cadbl(rst!operari1)
    .Fields("prepmaquina_de" + Trim(i)) = Format(atrim(rst!horainici), "hh:nn")
    .Fields("prepmaquina_fins" + Trim(i)) = Format(atrim(rst!horafi), "hh:nn")
    .Fields("prepmaquina_observacions" + Trim(i)) = atrim(rst!observacio)
    i = i + 1
    rst.MoveNext
    If i > 4 Then rst.MoveLast: rst.MoveNext
  Wend

  
  'temps funcionament
  Set rst = dbtmpb.OpenRecordset("select id,operari1,datainici,horainici,datafi,horafi,observacio,metresminut,totalmetres from Rebobinadores where comanda=" + comanda.Text + " and tipus='F' order by datainici,horainici")
  If Not rst.EOF Then rst.MoveLast
  For i = 1 To 8
    rst.MovePrevious
    If rst.BOF Then rst.MoveNext: i = 10
  Next i
  i = 1
  While Not rst.EOF
    .Fields("tempsreb_datai" + Trim(i)) = Format(atrim(rst!datainici), "dd/mm/yy")
    .Fields("tempsreb_op" + Trim(i)) = cadbl(rst!operari1)
    .Fields("tempsreb_horai" + Trim(i)) = Format(atrim(rst!horainici), "hh:nn")
    .Fields("tempsreb_horaf" + Trim(i)) = Format(atrim(rst!horafi), "hh:nn")
    .Fields("tempsreb_observacio" + Trim(i)) = atrim(rst!observacio)
    .Fields("tempsreb_mtrsmin" + Trim(i)) = cadbl(rst!metresminut)
    .Fields("tempsreb_metres" + Trim(i)) = cadbl(rst!totalmetres)
    i = i + 1
    If i > 8 Then rst.MoveLast
    rst.MoveNext
  Wend
  
  'acavar comandes

  'posso els camps de totals
    !pescanutu = pescanutu: !hmaquina = cadbl(hmaquina): !hfunc = cadbl(hfunc): !tbob = cadbl(tbob): !tmtrs = cadbl(tmetres): !tkilos = cadbl(tkilos): !mtrsmin = cadbl(kiloshora)
  '!acavada = comandaacavada
  Set rstbob = Nothing
  Set rst = Nothing
  
  
    
  End With
  
  'passo les bobines a la taula del llistat
  Set rst = dbtmpb.OpenRecordset("select id,operari1,datainici,horainici,datafi,horafi,observacio,metresminut from Rebobinadores where comanda=" + comanda.Text + " and tipus='F'")
  If rst.EOF Then dbtemp.Execute "insert into " + "tmp_reb_baixa_bob" + " (operari1,palet1,bobent1,paletsort,bobsort,kilos,metres) values (0,0,'0',0,0,0,0)"
  While Not rst.EOF
''     Set rsttmp =  dbtmpb.OpenRecordset("Select * from bobinesimp where controlid=" + atrim(cadbl(rst!id)))
     
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
        Set rsttmp2 = dbtmpb.OpenRecordset("select * from bobinesreb where controlid=" + atrim(cadbl(rst!id)))
        
        With rsttmp2
        If Not rsttmp2.EOF Then
         rsttmp2.MoveLast
         rsttmp2.MoveFirst
          'Else: dbtemp.Execute "insert into " + "tmp_reb_baixa_bob (operari,operari2,palet1,bobent1,bobsort,kilos,metres) values (0,0,0,'0',0,0,0)"
        End If
        While Not rsttmp2.EOF
          'If rsttmp2.AbsolutePosition + 1 = rsttmp2.RecordCount Then
              If Not rsttmp2.EOF Then Set rsttmpbob = dbtmpb.OpenRecordset("select * from bobinesentreb where id=" + atrim(cadbl(rsttmp2!id)) + " order by paletobobina ASC")
              nb = 0
              np = 0
              
              If Not rsttmpbob.EOF Then
                 rsttmpbob.MoveLast
                 rsttmpbob.MoveFirst
                 np = rsttmpbob!palet
                 nb = rsttmpbob!bobina
                 If rsttmpbob.RecordCount > 1 Then nb = "*" + nb
                 'aprofito per buscar lamplada del palet
                 Set rststocks = dbstocks.OpenRecordset("select ample from palets where idpalet=" + atrim(np))
                 If Not rststocks.EOF Then nample = rststocks!ample
              End If
              npalet = atrim(!palet)
              'If npalet = "0" Then npalet = "1"
              Set rstpesp = dbtmpb.OpenRecordset("select * from reb_pespalets where numpalet=" + npalet + " and comanda=" + atrim(rsttmp!comanda))
              If Not rstpesp.EOF Then pesp = atrim(cadbl(rstpesp!pespalet))
              pesp = cadcml(pesp)
              dbtemp.Execute "insert into " + "tmp_reb_baixa_bob (datapalet,paletsort,pespalet,operari,palet1,bobent1,bobsort,kilos,kilosnets,metres,senyals,observacions) values ('" + atrim(rst!datainici) + "'," + npalet + "," + pesp + "," + atrim(cadbl(!operari1)) + "," + atrim(np) + ",'" + atrim(nb) + "'," + atrim(cadbl(!numerodebobina)) + "," + atrim(cadcml(!kilos)) + "," + atrim(cadcml(!pesnet)) + "," + atrim(cadbl(!metres)) + "," + atrim(cadbl(!numempalmes)) + ",'" + atrim(!observacions) + "')"
              
           ' Else: dbtmpb.Execute "insert into tmp_imp_baixa_bob (operari,palet,bobent,bobsort,kilos,metres) values (" + atrim(cadbl(!operari1)) + "," + "0" + "," + "0" + "," + atrim(cadbl(!numerodebobina)) + "," + atrim("0") + "," + atrim("0") + ")"`'          End If
          rsttmp2.MoveNext
        '  rsttemp!ample = nample
        Wend
    ''    rsttmp.MoveNext
     ''Wend
     rst.MoveNext
     End With
  Wend
  
  rsttemp.Update
  dbtemp.Close
crear_taulatemp_bobinesdentrada
  
  Set rsttmp2 = Nothing
  Set rsttmpbob = Nothing
  'imprimir llistat
   
  'ATENCIÓ QUE FAIG SERVIR BAIXESREBOBINADORA.RPT PERÒ LA QUE S'IMPRIMEIX ES LA BAIXESREBOBINADORA_PDF perquè també fa el pdf
 '   i amb la versió que estava fet no es podia genera el PDF
 'vnommaquina = llegir_ini("baixes", "nommaquina", "comandes.ini")
 vnommaquina = buscarnommaquina(nummaq, "R")
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + nomllistat
' llistat.Destination = crptToWindow
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 2
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 llistat.Formulas(1) = "nommaquina='Reb. - " + atrim(nummaq) + " -" + vnommaquina + "'"
 llistat.Formulas(0) = "texteimpresio='" + treure_apostruf(texteimpresio) + "'"
 llistat.Formulas(2) = "pescanutu='" + atrim(pescanutu) + "'"
 llistat.Formulas(2) = "desarroll=" + atrim(vdesarroll)
' llistat.PrinterName = llegir_ini("Impressores", "nomfulla", "baixesimpressora.ini")
' llistat.PrinterPort = llegir_ini("Impressores", "portfulla", "baixesimpressora.ini")
' llistat.PrinterDriver = llegir_ini("Impressores", "driverfulla", "baixesimpressora.ini")
  DoEvents
  wait (2)
' If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
' llistat.PrintReport
' llistat.Action = 1
  escriure_ini "General", "exportantpdfs", "si", llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\organitzar.ini"
  crearlacarpetaperexportar cadbl(comanda.Text), vcarpetadesti
  exportarllistatapdf llistat, llegir_ini("General", "rutallistats", "comandes.ini") + "baixesrebobinadora_PDF.rpt", cadbl(comanda.Text), vcarpetadesti
  escriure_ini "General", "exportantpdfs", "no", llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\organitzar.ini"

  Set rsttmp = Nothing
  Set rst = Nothing
  Set dbstocks = Nothing
 'panelimprimir.Visible = False
 Form1.Caption = "Baixes Comandes (Rebobinadores)"
 
End Sub
Function buscarnommaquina(vm As Byte, vseccio As String) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from maquines where maquina='" + atrim(vseccio) + "' and codi=" + atrim(vm))
   If Not rst.EOF Then buscarnommaquina = atrim(rst!descripcio)
   Set rst = Nothing
End Function
Sub crearlacarpetaperexportar(numc As Double, carpetadesti As String)
   Dim carpetaprincipal As String
   Dim vcarpetatemporal As String
   Dim vubicaciocarpetadesti As String
   Dim vnomfitxer As String
   Dim vcont As Double
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
   vcont = 0
   If InStr(1, carpetadesti, vcarpetatemporal) = 0 Then
     vnomfitxer = Dir(vcarpetatemporal + "\cache_fabricacio\*.*", vbDirectory)
     While vnomfitxer <> "" And vcont < 100
         If vnomfitxer <> "." And vnomfitxer <> ".." Then
          Copiar_Fitxer vcarpetatemporal + "cache_fabricacio\" + vnomfitxer + "\", vubicaciocarpetadesti + "\", 5
          borra_carpeta vcarpetatemporal + "cache_fabricacio\" + vnomfitxer
          vnomfitxer = Dir(vcarpetatemporal + "\cache_fabricacio\*.*", vbDirectory)
         End If
         vcont = vcont + 1
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
Sub convertirformules(oreport As CRAXDDRT.Report, vllistat As CrystalReport)
  Dim i As Byte
  Dim vn As String
  Dim vv As String
  Dim v As String
  i = 0
  While vllistat.Formulas(i) <> ""
     v = vllistat.Formulas(i)
     vn = Mid(v, 1, InStr(1, v, "=") - 1)
     vv = Mid(v, InStr(1, v, "=") + 1)
     oreport.FormulaFields.GetItemByName(vn).Text = vv
     i = i + 1
  Wend
End Sub
Sub imprimir_etiquetapalets(vllistat As CrystalReport, vnomfitxerRPT As String)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim camp As TextObject
  Dim f  As OLEObject
  Dim vformula As String
  Dim i As Byte
  Dim vcopies As Byte
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(vnomfitxerRPT, 1)
  For i = 1 To oreport.Database.Tables.Count
    oreport.Database.Tables.Item(i).Location = vllistat.DataFiles(0)
  Next i
  'oreport.RecordSelectionFormula = "{Llaunes.numllauna}='" + UCase(atrim(numllauna)) + "'"
  'oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
  'oreport.DiscardSavedData
  convertirformules oreport, vllistat
'  oreport.DisplayProgressDialog = FalsE
  For i = 1 To vllistat.PrinterCopies
     oreport.PrintOut False
     wait 1
  Next i
End Sub
Sub exportarllistatapdf(vllistat As CrystalReport, vnomfitxerRPT As String, vnumc As Double, vcarpetadesti As String)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim camp As TextObject
  Dim f  As OLEObject
  Dim vformula As String
  Dim i As Byte
  Dim vcopies As Byte
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(vnomfitxerRPT, 1)
  For i = 1 To oreport.Database.Tables.Count
    oreport.Database.Tables.Item(i).Location = vllistat.DataFiles(0)
  Next i
  'oreport.RecordSelectionFormula = "{Llaunes.numllauna}='" + UCase(atrim(numllauna)) + "'"
  'oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
  'oreport.DiscardSavedData
  convertirformules oreport, vllistat
'  oreport.DisplayProgressDialog = FalsE
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat
  oreport.ExportOptions.DiskFileName = vcarpetadesti + "\" + atrim(vnumc) + "_BaixaRebobinadores.pdf"
  oreport.ExportOptions.PDFExportAllPages = True
  oreport.Export False
  For i = 1 To vllistat.PrinterCopies
     oreport.PrintOut False
     wait 1
  Next i
End Sub

Function cadcml(valor As Variant) As String
  valor = cadbl(valor)
  r = atrim(valor)
  cadcml = r
  If InStr(1, r, ",") <> 0 Then
     cadcml = Mid(r, 1, InStr(1, r, ",") - 1) + "." + Mid(r, InStr(1, r, ",") + 1)
  End If
  
End Function
Function possarnomfirmat() As String
  Dim rsttmp As Recordset
  Set rsttmp = dbtmp.OpenRecordset("select descripcio from operaris where maquina='R' and codi=" + atrim(cadbl(firmat)))
  If Not rsttmp.EOF Then
     possarnomfirmat = rsttmp!descripcio
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
Sub crear_taula_laminadora_baixa()
  Dim camps As String
  Dim campscapcalera As String
  Dim camps2 As String
  Dim camps3 As String
  Dim camps4 As String
  Dim campscapcalera2 As String
  Dim campspantone As String
  Dim campstotal As String
  nomfitxertemporal = "c:\temp\" + Format(Now, "~brddmmhhnnss") + ".mdb"
  On Error Resume Next
   MkDir "c:\temp"
   Kill "c:\temp\~br*.*"
   DBEngine.CreateDatabase nomfitxertemporal, dbLangGeneral, dbVersion10
   Set dbtemp = OpenDatabase(nomfitxertemporal)
   'dbtemp.Execute "drop table tmp_imp_empalmes"
   
'  On Error GoTo 0
'  On Error Resume Next
 '  dbtmpb.Execute "drop table tmp_reb_baixa"
 '  dbtmpb.Execute "drop table tmp_reb_baixa_bob"
  On Error GoTo 0
  campscapcalera = " comanda double, comanda2 double,comanda3 double, client string,comandaacavada byte,"
  campscapcalera = campscapcalera + "amplebob double,espesor double,bandesbones byte,ampleref double, bandesmerma byte,amplemerma double,lotcanutus string,lotbosses string, "
  camps = " prepmaquina_data1 string,prepmaquina_op1 byte,prepmaquina_de1 string,prepmaquina_fins1 string,prepmaquina_observacions1 string ,"
  camps = camps + " prepmaquina_data2 string,prepmaquina_op2 byte,prepmaquina_de2 string,prepmaquina_fins2 string,prepmaquina_observacions2 string ,"
  camps = camps + " prepmaquina_data3 string,prepmaquina_op3 byte,prepmaquina_de3 string,prepmaquina_fins3 string,prepmaquina_observacions3 string ,"
  camps = camps + " prepmaquina_data4 string,prepmaquina_op4 byte,prepmaquina_de4 string,prepmaquina_fins4 string,prepmaquina_observacions4 string ,"
  
  camps2 = "tempsreb_observacio1 string,tempsreb_op1 string,tempsreb_datai1 string,tempsreb_dataf1 string,tempsreb_horai1 string,tempsreb_horaf1 string,tempsreb_mtrsmin1 double,tempsreb_metres1 double,tempsreb_kilos1 double,"
  camps2 = camps2 + "tempsreb_observacio2 string,tempsreb_op2 string,tempsreb_datai2 string,tempsreb_dataf2 string,tempsreb_horai2 string,tempsreb_horaf2 string,tempsreb_mtrsmin2 double,tempsreb_metres2 double,tempsreb_kilos2 double,"
  camps2 = camps2 + "tempsreb_observacio3 string,tempsreb_op3 string,tempsreb_datai3 string,tempsreb_dataf3 string,tempsreb_horai3 string,tempsreb_horaf3 string,tempsreb_mtrsmin3 double,tempsreb_metres3 double,tempsreb_kilos3 double,"
  
  camps3 = "tempsreb_observacio4 string,tempsreb_op4 string,tempsreb_datai4 string,temspreb_dataf4 string,tempsreb_horai4 string,tempsreb_horaf4 string,tempsreb_mtrsmin4 double,tempsreb_metres4 double,tempsreb_kilos4 double,"
  camps3 = camps3 + "tempsreb_observacio5 string,tempsreb_op5 string,tempsreb_datai5 string,tempsreb_dataf5 string,tempsreb_horai5 string,tempsreb_horaf5 string,tempsreb_mtrsmin5 double,tempsreb_metres5 double,tempsreb_kilos5 double,"
  camps3 = camps3 + "tempsreb_observacio6 string,tempsreb_op6 string,tempsreb_datai6 string,tempsreb_dataf6 string,tempsreb_horai6 string,tempsreb_horaf6 string,tempsreb_mtrsmin6 double,tempsreb_metres6 double,tempsreb_kilos6 double,"
  
  camps4 = "tempsreb_observacio7 string,tempsreb_op7 string,tempsreb_datai7 string,temspreb_dataf7 string,tempsreb_horai7 string,tempsreb_horaf7 string,tempsreb_mtrsmin7 double,tempsreb_metres7 double,tempsreb_kilos7 double,"
  camps4 = camps4 + "tempsreb_observacio8 string,tempsreb_op8 string,tempsreb_datai8 string,tempsreb_dataf8 string,tempsreb_horai8 string,tempsreb_horaf8 string,tempsreb_mtrsmin8 double,tempsreb_metres8 double,tempsreb_kilos8 double,"
  camps4 = camps4 + " prepdr_data1 string,prepdr_op1 byte,prepdr_de1 string,prepdr_fins1 string,prepdr_observacions1 string ,"
  camps4 = camps4 + " prepdr_data2 string,prepdr_op2 byte,prepdr_de2 string,prepdr_fins2 string,prepdr_observacions2 string ,"
  camps4 = camps4 + " prepdr_data3 string,prepdr_op3 byte,prepdr_de3 string,prepdr_fins3 string,prepdr_observacions3 string ,"
  
  
    'creo els camps de total
  campstotal = " hmaquina double,  hfunc double,  tbob double,tmtrs double, tkilos double, mtrsmin double, pescanutu double "
  
  'ample double,plegat double,solapa double,espessor double,metres double,kilos double)"
  'escriure_ini "a", "b", campsextra + camps + camps3 + camps2 + campspantone + campspantone2 + campstotal, "prova.ini"
    dbtemp.Execute ("create table tmp_reb_baixa (" + campscapcalera + camps + camps2 + camps3 + camps4 + campstotal + ")")
    dbtemp.Execute ("create table tmp_reb_baixa_bob (idbob integer,operari byte,operari2 byte,palet1 double,bobent1 string,palet2 double,bobent2 string,paletsort integer,bobsort integer,kilos double,metres double,senyals byte,pespalet double,observacions string,kilosnets double,datapalet string)")
  
End Sub



Private Sub Command8_Click()
client.ToolTipText = client.Caption
calcular_totals
wait 2
imprimir_fulla
End Sub

Private Sub Command9_Click()
If horaapretada <> 1 Then
    dblots.AllowAddNew = False
    dblots.AllowDelete = False
    dblots.AllowUpdate = False
    dblots.MarqueeStyle = 3
    dblots.Visible = False
    DoEvents
  framepantones.Visible = Not framepantones.Visible
  frameempalmes.Visible = False
  framebobentrada.Visible = False
  If Not framepantones.Visible Then If reixabobines.Enabled Then reixabobines.SetFocus
End If
End Sub

Private Sub Command9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If horaapretada = 0 Then horaapretada = Now
  
End Sub

Private Sub Command9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 
    horaapretada = 0
 
End Sub

Private Sub compantone_DblClick(Index As Integer)
'dblots.Visible = True
'dblots.Tag = Index
End Sub

Private Sub compantone_LostFocus(Index As Integer)
imppantones.Refresh
End Sub

Private Sub dblots_DblClick()
'If dblots.MarqueeStyle = 6 Then
'    dblots.Visible = False
'    dblots.AllowAddNew = False
'    dblots.AllowDelete = False
'    dblots.AllowUpdate = False
''    dblots.MarqueeStyle = 3
'    dblots.Visible = False
''    framepantones.Visible = False
''    Exit Sub
'End If
'If Not lots.Recordset.EOF Then
'  compantone(cadbl(dblots.Tag)) = atrim(lots.Recordset!codilot)
'End If
'dblots.Visible = False
End Sub

Private Sub dblots_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then dblots.Tag = "": dblots.Visible = False
End Sub

Private Sub eliminarbobentrada_Click()
  If bobinesent.Recordset.EOF Then Exit Sub
  If MsgBox("Segur que vols eliminar la bobina d'entrada " + atrim(bobinesent.Recordset!palet) + "/" + atrim(bobinesent.Recordset!bobina), vbExclamation + vbYesNo, "Borrar bobina d'entrada") = vbYes Then
    'carregar_bobinesdentrada "marcarutilitzada", , bobinesent.Recordset!palet, bobinesent.Recordset!bobina, ncomanda, False
    bobinesent.Recordset.Delete
    bobinesent.Refresh
    possarnumbobent
  End If
  
End Sub

Private Sub espesor_LostFocus()
  guarda_totals

End Sub

Private Sub etpesbascula_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Shift = 2 Then
     etpesbascula = cadbl(InputBox("Entra el pes"))
     If cadbl(etpesbascula) > 0 Then
        etpesbascula.Tag = "manual"
        etpesbascula.BackColor = QBColor(13)
          Else: etpesbascula.Tag = "": etpesbascula.BackColor = &HC0C0FF
     End If
  End If
End Sub

Private Sub firmat_DblClick()
Exit Sub
If firmat <> "" Then
   firmat = ""
  Else: firmar_fulla
End If
End Sub

Sub mirarsihihalafontTTFdecodidebarres()
   Dim objshell As Variant
   Dim objFolderItem As Variant
   On Error Resume Next
   If existeix("c:\windows\fonts\free3of9.ttf") Then Exit Sub
   Copiar_Fitxer llegir_ini("General", "rutallistats", fitxerini) + "\free3of9.ttf", "c:\windows\fonts"
  ' Set objshell = CreateObject("Shell.Application")
  ' Set objFolder = objshell.Namespace("C:\windows\Fonts")
  ' Set objFolderItem = objFolder.ParseName("free3of9.ttf")
  ' objFolderItem.InvokeVerb ("Install")
  
   Shell "reg add """ + "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" + """ /v """ + "Free 3 of 9 Regular (TrueType)" + """ /t REG_SZ /d free3of9.ttf /f"
End Sub
Private Sub Form_Activate()
 If cadbl(numop) = 0 And Form1.Tag <> "carregant" Then nomoperari_Click
 
End Sub
Sub demanarcomandadebossesicanutus()
    Dim rst As Recordset
    Dim rstc As Recordset
    Set rstc = dbtmp.OpenRecordset("select tubbase from comandes where comanda=" + atrim(comanda))
    Set rst = dbtmpb.OpenRecordset("select comandabosses1,comandabosses2,comandacanutus1,comandacanutus2 from rebobinadorestot where comanda=" + atrim(cadbl(Form1.comanda)))
    If Not rst.EOF Then
        While Len(atrim(rst!comandabosses1)) < 3 Or Len(atrim(rst!comandacanutus1)) < 3
         Load formbossesperembossar
         formbossesperembossar.Show
         formbossesperembossar.escullirisortir cadbl(rstc!tubbase)
         Set rst = dbtmpb.OpenRecordset("select comandabosses1,comandabosses2,comandacanutus1,comandacanutus2 from rebobinadorestot where comanda=" + atrim(cadbl(Form1.comanda)))
         If Len(atrim(rst!comandabosses1)) < 3 Or Len(atrim(rst!comandacanutus1)) < 3 Then
           MsgBox "Hi ha d'haver el lot de bosses i el de canutus entrat per poder continuar", vbCritical + vbOKOnly, "Lots"
           formbossesperembossar.Show 1
         End If
        Wend
    End If
    Set rst = Nothing
    Set rstc = Nothing
    
    canutustallats = ""
End Sub
Private Sub Form_Click()
 ' imprimirfullPackinglistXrPalet 227507, 1
'MSComm1_OnComm
  'avisarquelacomandasestaacabant cadbl(comanda), "R"
'comprovar_calloffs 170598
' imprimir_controlbobina0 cadbl(comanda)
 'imprimiretiquetaverificacio 1
'imprimir_controlbobina0 cadbl(comanda)

'  demanarcomandadebosses
'imprimir_bobina "Ext.Bobina"
   'imprimir_controlqualitatVQ cadbl(comanda)

  'dbtmpb.Execute "UPDATE bobinesreb INNER JOIN rebobinadores ON bobinesreb.controlid = rebobinadores.Id SET bobinesreb.metres = 495 WHERE (((rebobinadores.comanda)=148992) AND ((bobinesreb.numerodebobina)=39));"
  
'imprimir_bobina
'preparar_etiqueta_zebra
'imprimir_etiqueta_zebra
' appac = Shell("C:\Archivos de programa\swetiq.exe c:\prova.swe")
' wait (5)
' AppActivate appac
' wait (1)
' SendKeys ("%d")
' SendKeys ("{RIGHT}")
' SendKeys ("{ENTER}")
' SendKeys ("{ENTER}")
' SendKeys ("{ENTER}")
'MsgBox llegirpesbascula
'If numbobinesnocorrelatiu Then MsgBox "Els numeros de bobines no son correlatius. Reviseu per continuar la bobina " + r
End Sub

Sub possar_botons_palets()
Dim grup As Byte
Dim i As Byte
netejar_botons_palets
grup = cadbl(framepalets.Tag)
For i = 0 To 9
  botopalets(i).Caption = atrim((i + 1) + grup)
Next i
If Command16.Tag <> "E" Then botopalets_Click 0
End Sub
'Sub obrestocks(Optional noobrirbd As Boolean)
'camistocks = llegir_ini("General", "ruta_stocksmdb", "comandes.ini")
'If camistocks = "{[}]" Then camistocks = "\\Ser2\documentos\Stock Reclamaciones\Estoc inplacsa.mdb"
'If Not existeix(camistocks) Then camistocks = "\\serverprodu\dades\progcomandes\dades\copiaestocinplacsa.mdb"
'If Not noobrirbd Then Set dbstocks = OpenDatabase(camistocks)
'
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If Chr$(KeyAscii) = "'" Then KeyAscii = Asc("´")
End Sub

Sub crear_taulatemp_bobinesdentrada()
  'Dim camps As String
  If nomfitxertemporalbobent <> "" Then Exit Sub
  nomfitxertemporalbobent = "c:\temp\~bibe" + Format(Now, "ddmmhhnnss") + ".mdb"
  On Error Resume Next
   MkDir "c:\temp"
   Kill "c:\temp\~bibe*.*"
   DBEngine.CreateDatabase nomfitxertemporalbobent, dbLangGeneral, dbVersion10
   Set dbtemp = OpenDatabase(nomfitxertemporalbobent)
   'dbtemp.Execute "drop table tmp_imp_empalmes"
  On Error GoTo 0
  camps = "sel bit,idpalet double,idbobina double,metres double, utilitzada bit,tipus string(1),taula string,idb double"
  dbtemp.Execute ("create table selecciobobentrada (" + camps) + ")"

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

Private Sub Form_Load()
  Dim camistocks As String
  Dim MyWorkspace As Workspace
  Dim arguments As Variant
  On Error Resume Next
  
  Form1.Tag = "carregant"
  '
  'If EstaCorriendo("baixesrebobinadora.exe") Then MsgBox "El programa ja està funcionant.": End
  If App.PrevInstance Then MsgBox "El programa ja està funcionant.": End
  arguments = ObtenerLíneaComando
  Form1.Tag = ""
  fitxerini = "comandes.ini"
  Shell "c:\windows\regedit.exe /s \\serverprodu\dades\progcomandes\aplicacio\desactivarctrl.reg"
  Shell "c:\windows\regedit.exe /s \\serverprodu\dades\progcomandes\aplicacio\activarctrl.reg"
  Shell ("net time \\serverprodu /set /y")
  camicomandes = llegir_ini("General", "cami", "comandes.ini")
  cami = llegir_ini("General", "camibaixes", "comandes.ini")
  
  If LCase(App.EXEName) <> "baixesrebobinadora" And LCase(App.EXEName) <> "baixes rebobinadores" Then Form1.BackColor = &HFF80FF
   
  
  obrestocks True
  If cami = "{[}]" Then
    escriure_ini "General", "camibaixes", InputBox("Entra la ruta de baixes", "Atenció", "y:\comandes\baixes.mdb"), "comandes.ini"
  End If
  
  If Not existeix("c:\ordprog.ini") Then comanda = cadbl(llegir_ini("Baixes", "ultimacomanda", "comandes.ini"))
  r = cadbl(llegir_ini("Baixes", "nummaq", "comandes.ini"))
  nummaq = cadbl(r)
  assignardecimalipunt
  If nummaq = 0 Then
    maquina.Visible = True
   Else: maquina.Visible = False
  End If
  lletraseccio = "R"
  
  centerscreen Me
  'cami = "\\SERVERprodu\dades\progcomandes\dades\baixesprova.mdb"
  'Set MyWorkspace = DBEngine.CreateWorkspace("New", "Rebobinadora" + atrim(nummaq), "")
  Set dbtmpb = OpenDatabase(cami)
  Set dbtmp = OpenDatabase(camicomandes)
  crear_taulatemp_bobinesdentrada
  crear_taula_bobentrada
  On Error Resume Next
  dbtmpb.Execute ("create table lotslam (nomlot string,codilot string)")
  'dbtmpb.Execute "drop table bobentradatmpreb" + atrim(nummaq)
  
  If UCase(arguments(2)) = "LOTSBOSSES" Then
    lletraseccio = "LOTSBOSSES"
    rellotge.Enabled = False
    formbossesperembossar.Show 1
    End
  End If
  On Error GoTo 0
  
  Rebobinadores.DatabaseName = cami
  imppantones.DatabaseName = cami
  bobines.DatabaseName = cami
  
  empalmes.DatabaseName = cami
  bobinesent.DatabaseName = cami
  
  lots.DatabaseName = cami
  lots.Refresh
  Set dbtmpb = OpenDatabase(Rebobinadores.DatabaseName)
  rellotge.Enabled = True
  rellotge.Interval = 900
  
 'If cadbl(nummaq) = 0 Then MsgBox "No hi ha el numero de màquina posat": Exit Sub
  
  
  Rebobinadores.RecordSource = "select * from Rebobinadores where comanda=-1"
  Rebobinadores.Refresh
  bobinesent.RecordSource = "select * from bobinesentreb where id=99999999"
  bobinesent.Refresh
  
  mirarsihihalafontTTFdecodidebarres
  
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
  
  
  For Each objecte In Me
      If objecte.Name <> "reciclarmaterial1" And objecte.Name <> "AcroPDF1" And objecte.Name <> "MSComm1" And objecte.Name <> "llistatpalet" And objecte.Name <> "nomoperari" And objecte.Name <> "Line1" And objecte.Name <> "rellotge" And objecte.Name <> "llistat" Then
        objecte.Enabled = False
      End If
     Next objecte
     
     
  frameempalmes.ZOrder 0
    framepantones.ZOrder 0
    framebobentrada.ZOrder 0
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  'If Shift = 2 Then MsgBox Trim(App.Major) + "." + Trim(App.Minor) + "." + Trim(App.Revision)
  If Shift = 2 Then
    If InputBoxEx("Entra la contrasenya de configuració, la de sempre però de 4", "Programador", , , , , , SPassword) = "9909" Then
     If MsgBox("Prem si per activar el bloqueig del Ctrl+Alt+Supr i no per desactivar-lo", vbYesNo, "Atenció") = vbYes Then
       Shell "c:\windows\regedit.exe /s \\serverprodu\dades\progcomandes\aplicacio\desactivarctrl.reg"
        Else: Shell "c:\windows\regedit.exe /s \\serverprodu\dades\progcomandes\aplicacio\activarctrl.reg"
     End If
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'finalitza_seccio
  Cancel = 0
 If MSComm1.PortOpen Then MSComm1.PortOpen = False
 Form1.Tag = "tancant"
 Unload capcalera
 Cancel = 0
 End
End Sub

Private Sub impresores_Reposition()
 


End Sub
Sub ensenya_les_bobines()
  Dim bk As String
  Dim rstcp As Recordset
  
  If Me.Name = "reixabobines" Then Exit Sub
  r = "-1"
  If Rebobinadores.Recordset!tipus = "F" Then
   Set rstcp = Rebobinadores.Recordset.Clone
   r = ""
   rstcp.MoveFirst
   While Not rstcp.EOF
    If rstcp!tipus = "F" Then
       r = r + IIf(r <> "", ",", "") + atrim(cadbl(rstcp!id))
    End If
    rstcp.MoveNext
   Wend
  End If
  If Not bobines.Recordset.EOF And Not bobines.Recordset.BOF Then
    On Error Resume Next
    bk = bobines.Recordset!numerodebobina
    On Error GoTo 0
  End If
  bobines.Tag = r
  bobines.RecordSource = "select * from bobinesreb where palet=" + atrim(numpalet) + " and controlid in(" + r + ") order by numerodebobina"
  If numpalet = 1 Then bobines.RecordSource = "select * from bobinesreb where (palet=" + atrim(numpalet) + " or palet=0) and controlid in (" + r + ") order by numerodebobina"
  bobines.Refresh
  If bobines.Recordset.EOF And r <> "-1" Then
    
           Set rstcp = dbtmpb.OpenRecordset("select max(palet) as maxpalet from bobinesreb where  controlid in(" + r + ")")
           If Not rstcp.EOF Then
             If cadbl(rstcp!maxpalet) > 30 Then
                  MsgBox "Hi ha un numero de palet mes gran de 30. " + atrim(rstcp!maxpalet): numpalet = 1
               Else: numpalet = cadbl(rstcp!maxpalet) + 1
             End If
            Else: numpalet = 1
           End If
           
  End If
  bobines.Recordset.LockEdits = False
 bobinesent.Recordset.LockEdits = False
  'If bobines.Recordset.EOF Then
  '  bobines.RecordSource = "select * from bobinesreb where  controlid=" + r + " order by numerodebobina"
  '  bobines.Refresh
  'End If
  On Error Resume Next
  If bk <> "" Then
     bobines.Recordset.FindFirst "numerodebobina=" + bk
   Else: bobines.Recordset.MoveLast
  End If
  'If Not IsEmpty(bk) Then bobines.Recordset.Bookmark = bk
  
End Sub
Sub colocarelsbotonsdelspalets()
  If numpalet < 1 Then Exit Sub
  i = Fix((numpalet - 1) / 10)
  Command16.Tag = "E"
  Select Case i
      Case 0: Command16_Click
      Case 1: Command17_Click
      Case 2: Command18_Click
  End Select
  Command16.Tag = ""
 botopalets_Click (numpalet - (Fix((numpalet - 1) / 10) * 10)) - 1
 If Not bobines.Recordset.EOF Then bobines.Recordset.MoveLast
End Sub

Private Sub imprimir_Click()
    imprimir_bobina "Muestra Cli", True
End Sub

Private Sub kbpantone_LostFocus(Index As Integer)
' Dim totaltinta As Double
'imppantones.Refresh
'totaltinta = 0
'For i = 0 To 9
'   totaltinta = totaltinta + cadbl(kbpantone(i))
'Next i
'impresores.Recordset.Edit
'impresores.Recordset!kgtinta = totaltinta
'impresores.Recordset.Update
'If totaltinta > 0 And framepantones.Tag = "E" Then dbtmp.Execute "update comandes set proximaseccio='I' where comanda=" + atrim(cadbl(comanda.Text)): framepantones.Tag = "I"
End Sub

Private Sub maquina_Click()
   nummaq = cadbl(InputBox("Entra el numero de màquina [1,2 o 3]", "Atenció"))
   If nummaq > 0 And numaq < 4 Then
      'framebobines.Enabled = True
    Else: nummaq = 0 ': framebobines.Enabled = False
   End If
   maquina.Caption = "Maq: " + atrim(nummaq)
   maquina.Tag = nummaq
End Sub

Private Sub mostracli_Click()
guarda_totals
End Sub

Private Sub PDFX1_OnError(lErr As Long, sErr As String)

End Sub

Private Sub MSComm1_OnComm()

Dim sData As String

    Select Case MSComm1.CommEvent
        Case comEvReceive ' Dades rebudes
            ' Assegura't que hi hagi dades per llegir
            If MSComm1.InBufferCount > 0 Then
                sData = MSComm1.Input ' Llegir les dades del buffer d'entrada
                etpesbascula.ToolTipText = sData ' Afegir les dades al final del TextBox
            End If

        Case comEvCTS, comEvDSR, comEvCD, comEvRing, comEvEOF
            ' Altres esdeveniments, pots afegir lògica si cal
            ' Per exemple, monitoritzar senyals de control
            etpesbascula.ToolTipText = "Esdeveniment: " & MSComm1.CommEvent

        Case comEvBreak, comEvCDTO, comEvDSRTO, comEvFrame, comEvOverrun, comEvRxOver, comEvRxParity, comEvTxFull
            ' Errors de comunicació
            etpesbascula.ToolTipText = "Error de comunicació: " & MSComm1.CommEvent
            ' Podries mostrar el tipus d'error o tancar el port
    End Select
End Sub

Private Sub pespalet_Change()
  If Screen.ActiveControl.Name = "pespalet" Then
   If cadbl(pespalet) > 0 Then gravar_pespalet
  End If
End Sub

Private Sub pespalet_DblClick()
pespalet.Tag = "pesar"
agafarpesbascula_Click
End Sub

Private Sub pespalet_LostFocus()
 
  If controlactiu = "agafarpesbascula" Then
     pespalet.Tag = "pesar"
    Else: pespalet.Tag = ""
  End If
End Sub

Private Sub proces_Change()
 Dim rsttmpp As Recordset
 
 Set rsttmpp = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(proces) + "'")
 If InStr(1, rsttmpp!ruta, "R") = 0 Then proces.Tag = "": Exit Sub
 If Not rsttmpp.EOF Then proces.Tag = Mid(rsttmpp!ruta, InStr(1, rsttmpp!ruta, "R") - 1, 1)
End Sub

Private Sub Rebobinadores_Reposition()
  
    If Not Rebobinadores.Recordset.EOF Then
          If atrim(Rebobinadores.Recordset!tipus) = "F" Then ensenya_les_bobines
       If barraestat.Caption <> "Calculant els totals..." Then colocarelsbotonsdelspalets
       
       'framebobines.Enabled = False
     End If
     missatge_exesdemtrskg
End Sub

Private Sub nomoperari_Click()
 Dim numoptmp As Integer
 Dim nomoptmp As String
 If barraestat.Caption = "Calculant els totals..." Then Exit Sub
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "select codi,descripcio from operaris where maquina='R' and actiu<>0 order by codi asc"
  formseleccio.Caption = "Selecció d'Operari"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   numoptmp = cadbl(formseleccio.Data1.Recordset!codi)
   nomoptmp = atrim(formseleccio.Data1.Recordset!descripcio)
   'If InStr(1, nomoperari.Caption, "MARTINEZ") Then
   '    Command12.Visible = True
   '   Else: Command12.Visible = False
   'End If
  End If
  Unload formseleccio
  If numoptmp <> 0 Then
     nomoperari = nomoptmp
     numop = numoptmp
     For Each objecte In Me
      If objecte.Name <> "reciclarmaterial1" And objecte.Name <> "AcroPDF1" And objecte.Name <> "MSComm1" And objecte.Name <> "llistatpalet" And objecte.Name <> "llistat" And objecte.Name <> "Line1" And objecte.Name <> "comandaacavada" Then
        objecte.Enabled = True
      End If
     Next objecte
      Else: If cadbl(numop) = 0 Then MsgBox "Has d'escullir un operari per treballar": Exit Sub
  End If
   Command4_Click
End Sub

Private Sub pantone_LostFocus(Index As Integer)
imppantones.Refresh
End Sub

Private Sub reixa_AfterUpdate()
  'calcular_totals
End Sub

Private Sub reixa_BeforeDelete(Cancel As Integer)
  If controlactiu <> "Command14" Then
   If MsgBox("Segur que vols borrar aquesta linia i tot el seu contingut?", vbYesNo, "Atenció") = vbNo Then Cancel = 1
  End If
  If Cancel <> 1 Then
    If Rebobinadores.Recordset!tipus = "F" Then r = atrim(cadbl(Rebobinadores.Recordset!id))
    dbtmpb.Execute "delete * from bobinesreb where controlid=" + r
  End If
End Sub

Private Sub reixa_DblClick()
If reixa.col = 14 Then
   r = triar_observacio(Rebobinadores.Recordset!tipus)
   If Len(r) > 4 Then
     r = Mid(r, 4, Len(r))
     If r <> "" Then
       If reixa.Text <> "" Then
           reixa.Text = reixa.Text + " <> " + r
           Else: reixa.Text = r
       End If
     End If
   End If
End If

If reixa.col = 0 Then
  reixa.Text = escullir_operari
  nomoperari = UCase(r)
  numop = reixa.Text
End If
End Sub
Function triar_observacio(tipus As String) As String
  'Dim rsttriar As Recordset
  'Set rsttriar = dbtmp.OpenRecordset("select * from constantsobservacio where mid(observacio,1,1)='" + tipus + "'")
  'While Not rsttriar.EOF
  '  rsttriar.MoveNext
  'Wend
  
  Load formseleccio
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select * from constantsobservacio where mid(observacio,1,2)='L" + tipus + "'"
  formseleccio.Caption = "Triar Observació"
  formseleccio.refrescar
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
  AcroPDF1.Visible = False
End Sub

Private Sub reixa_KeyUp(KeyCode As Integer, Shift As Integer)
 If reixa.col = 4 And KeyCode > 46 Then
     If (Len(reixa.Text)) >= 4 Then reixa.col = 5
  End If
  If reixa.col = 3 And KeyCode > 46 Then
     If (Len(reixa.Text)) >= 6 Then reixa.col = 4
  End If
  If reixa.col = 2 And KeyCode > 46 Then
     If (Len(reixa.Text)) >= 4 Then
       reixa.col = 3
     End If
  End If
  If reixa.col = 1 And KeyCode > 46 Then
     If (Len(reixa.Text)) >= 6 Then reixa.col = 2
  End If
  If reixa.col = 12 And KeyCode > 46 And reixa.Columns(5) = "C" Then
     If UCase(Chr$(KeyCode)) = "S" Then
        reixa = "Sí"
       Else: reixa = "No"
     End If
     KeyCode = 0
  End If
  If reixa.col = 14 And KeyCode > 46 Then
      If (Len(reixa.Text)) > 99 Then reixa.Text = Mid(reixa.Text, 1, 99)
  End If
End Sub

Sub bloquejar_camps_innecesaris()
If Rebobinadores.Recordset.EOF Then Exit Sub
For i = 0 To 11
  reixa.Columns(i).Locked = False
Next i
reixa.Columns(5).Locked = True
reixa.Columns(6).Locked = True
reixa.Columns(7).Locked = True
reixa.Columns(8).Locked = True
reixa.Columns(9).Locked = True
reixa.Columns(10).Locked = True
reixa.Columns(11).Locked = False
'If Rebobinadores.Recordset!tipus = "C" Then reixa.Columns(12).Locked = False: reixa.Columns(14).Locked = False  ': reixa.Columns(11).Locked = False:reixa.Columns(7).Locked = False
If Rebobinadores.Recordset!tipus = "F" Then reixa.Columns(9).Locked = False




End Sub

Private Sub reixa_LostFocus()
'   AcroPDF1.Visible = True
End Sub

Private Sub reixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 Dim valtmp As String
 
 
 If reixa.col = 0 Then reixa.EditActive = False
 '-------
 bloquejar_camps_innecesaris
 If Not Rebobinadores.Recordset.EOF Then
 'texteimpresio = atrim(impresores.Recordset!texteimpresio)
  If atrim(Rebobinadores.Recordset!tipus) = "F" Then
     framebobines.Enabled = True
       Else: framebobines.Enabled = False: framepantones.Visible = False
  End If
 End If
 'If LastCol >= 0 Then valtmp = atrim(reixa.Columns(LastCol).CellValue(LastRow))
 If LastCol = 1 Or LastCol = 2 Then
   valtmp = reixa.Columns(LastCol).Text
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
  valtmp = reixa.Columns(LastCol).Text
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
 If LastCol <> 0 Then comprovardiferenciesdehoresinicifi LastRow
 'calcular_totals
 
 frameempalmes.Visible = False
 framepantones.Visible = False
 
End Sub
Sub comprovardiferenciesdehoresinicifi(LastRow As Variant)
   Dim vhinici As String
   Dim vhfi As String
   Dim vdinici As String
   Dim vdfi As String
   Dim vdatainici As Date
   Dim vdatafi As Date
   If atrim(LastRow) = "" Then Exit Sub
   On Error Resume Next
   vdinici = atrim(reixa.Columns(1).CellValue(LastRow))
   vhinici = atrim(reixa.Columns(2).CellValue(LastRow))
   vdfi = atrim(reixa.Columns(3).CellValue(LastRow))
   vhfi = atrim(reixa.Columns(4).CellValue(LastRow))
   If vdinici = "" Or vhinici = "" Or vdfi = "" Or vhfi = "" Then Exit Sub
   vdatainici = vdinici + " " + vhinici
   vdatafi = vdfi + " " + vhfi
   If IsDate(vdatainici) And IsDate(vdatafi) Then
       If DateDiff("n", vdatainici, vdatafi) > 490 Then MsgBox "Hi ha mes de 8 hores de diferencia entre inici i final de la baixa, comprova si es correcte", vbCritical, "Atenció"
   End If
End Sub
Private Sub reixabobines_AfterColUpdate(ByVal ColIndex As Integer)
 If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
 On Error Resume Next
 bobines.Recordset.Fields(reixabobines.Columns(ColIndex).DataField) = reixabobines.Columns(ColIndex).Text
 reixabobines.EditActive = False
 If reixabobines.Columns(ColIndex).DataField = "kilos" Or reixabobines.Columns(ColIndex).DataField = "metres" Then avisar_pesbobinaTEORIC cadbl(reixabobines.Columns("Kilos")), cadbl(reixabobines.Columns("Metres"))
'bobines.Recordset.Update
End Sub
Sub avisar_pesbobinaTEORIC(vpesentrat As Double, vmetresentrats As Double, Optional vnoavisar As Boolean)
   Dim rst As Recordset
   Dim vpesTeoricbobina As Double
   Dim vbandes As Double
   Dim vNOmsg As Boolean
   vNOmsg = vnoavisar
   vnoavisar = False
   If vpesentrat = 0 Or vmetresentrats = 0 Then Exit Sub
   Set rst = dbtmp.OpenRecordset("select rebkilos,rebmtrs from comandes where comanda=" + atrim(comanda))
   If rst.EOF Then GoTo fi
   vbandes = IIf(cadbl(bandes) = 0, 1, cadbl(bandes))
   If cadbl(rst!rebmtrs) = 0 Then GoTo fi
   vpesTeoricbobina = Redondejar((cadbl(rst!rebkilos) / rst!rebmtrs) * vmetresentrats, 2)
   vpesTeoricbobina = vpesTeoricbobina + (cadbl(amplebob) / 50)  'pes canutu 1kg cada 50 cms
   If (vpesTeoricbobina / 1.2) > vpesentrat Or (vpesTeoricbobina * 1.2) < vpesentrat Then      '+- 20%
       If Not vNOmsg Then MsgBox "Aquesta bobina hauria de pesar aproximadament " + atrim(vpesTeoricbobina) + "Kg revisa que el pes sigui correcte pels metres que has possat.", vbCritical, "Error"
       vnoavisar = True
   End If
fi:
   Set rst = Nothing
End Sub

Private Sub reixabobines_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
  tempseditant = Now
End Sub

Private Sub reixabobines_ColEdit(ByVal ColIndex As Integer)
tempseditant = Now
End Sub

Private Sub reixabobines_DblClick()
Dim nop As Double
If reixabobines.col = 11 Then
   r = triar_observacio("B")
   If r <> "" Then reixabobines.Text = r
End If
If reixabobines.col = 0 Then
  nop = cadbl(escullir_operari)
  If nop > 0 Then
   'nomoperari = UCase(r)
   'numop = nop
   reixabobines.Columns("operari1") = atrim(nop)
  End If
End If
End Sub
Function escullir_operari() As String
  Dim opvell As Byte
  opvell = numop
  r = nomoperari
 'While cadbl(escullir_operari) = 0
   Load formseleccio
   formseleccio.Data1.DatabaseName = camicomandes
   formseleccio.Data1.RecordSource = "select codi,descripcio from operaris where maquina='R' and actiu<>0 order by codi asc"
   formseleccio.Caption = "Selecció d'Operari"
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

Private Sub reixabobines_GotFocus()
etmetresbob.Visible = False
 frameempalmes.Visible = False
 framepantones.Visible = False
 If reixabobines.col <> 7 Then
     framebobentrada.Visible = True
   Else: framebobentrada.Visible = False
 End If
End Sub

Private Sub reixabobines_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim c As String
 tempseditant = Now
 If reixabobines.col = 3 And KeyCode > 48 And KeyCode < 58 Then
   c = Chr$(KeyCode)
   If cadbl(c) = 0 Then c = ""
   nump = InputBox("Entra el nou numero de palet.", "Nou Palet", c)
   If cadbl(nump) > 0 And cadbl(nump) < 31 Then
     If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
     bobines.Recordset!palet = nump
     bobines.Recordset.Update
   End If
 End If
End Sub

Private Sub reixabobines_LostFocus()
Dim camps As String
camps = "bobentradaagafarpesbasculabotopaletsCommand7Command9Command12Command13Command3Command5Command6"

If reixabobines.col > 1 And InStr(1, camps, controlactiu) = 0 And Screen.ActiveForm.Name = "Form1" Then
  calcular_totals
End If
End Sub

Private Sub reixabobines_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Static fila As Double
Dim pesnetanterior As Double
Dim vreixatmp  As DBGrid
'If Not Screen.ActiveControl.Name = "reixabobines" Then Exit Sub
guardar_reg_bobines
If IsNull(fila) Then fila = 0
If fila <> reixabobines.row Then
 'calcular_totals
End If
fila = reixabobines.row
If LastCol = 5 And LastRow = reixabobines.Bookmark And cadbl(reixabobines.Columns(6)) = 0 Then
   If pescanutu > 0 Then
      reixabobines.Columns(6) = cadbl(reixabobines.Columns(5)) - pescanutu
       Else: reixabobines.Columns(6) = "0"
   End If
End If
If reixabobines.col <> 11 And reixabobines.Columns(reixabobines.col).DataField <> "metres" And reixabobines.Columns(reixabobines.col).DataField <> "kilos" Then
     framebobentrada.Visible = True
   Else: framebobentrada.Visible = False
 End If
 If pescanutu > 0 Then
   pesnetanterior = cadbl(reixabobines.Columns("pesnet"))
   If cadbl(reixabobines.Columns("kilos")) - cadbl(pescanutu) > 0 Then
     reixabobines.Columns("pesnet") = cadbl(reixabobines.Columns("kilos")) - cadbl(pescanutu)
       Else: reixabobines.Columns("pesnet") = 0
   End If
   If pesnetanterior <> cadbl(reixabobines.Columns("pesnet")) Then bobines.Recordset.Edit: bobines.Recordset.Update
 End If
 If client.Tag <> "7" Then
    If cadbl(reixabobines.Columns("pesnet")) > 0 And cadbl(pescanutu) = 0 Then
         MsgBox "Hi ha un pes net entrat sense pes de canutu, es correcte?", vbCritical, "Atenció"
    End If
 End If
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

Private Sub reixaempalmes_DblClick()
If reixaempalmes.col = 1 Then
   r = triar_observacio("S")
   If r <> "" Then reixaempalmes.Text = r
End If
End Sub

Private Sub reixaempalmes_OnAddNew()
 empalmes.Recordset!id = bobines.Recordset!id
 'reixa.col = 0
 
End Sub

Sub posarpesbascula()
Static buffer As String
Static nobascula As Boolean
Dim vnumport As Double

If etpesbascula.Tag = "manual" Then Exit Sub
If nobascula Then Exit Sub
If Not MSComm1.PortOpen Then
  vnumport = cadbl(llegir_ini("Baixes", "numportbascula", "comandes.ini"))
  If vnumport = 0 Then
     vnumport = 1
     escriure_ini "Baixes", "numportbascula", "1", "comandes.ini"
  End If
  MSComm1.CommPort = vnumport
 ' 9600 baudios, sin paridad, 7 bits de datos y 1 bit de parada.
  MSComm1.Settings = "9600,n,8,1"
 ' If nummaq = 1 Then MSComm1.Settings = "2400,n,8,1"
 ' Indicar al control que lea todo el búfer al usar Input.
  MSComm1.InputLen = 0
 
  MSComm1.RTSEnable = True 'Por si necesitas habilitar el RTS
 
 'Abrir Puertos
 On Error GoTo nopossarpes
  MSComm1.PortOpen = True
  On Error GoTo 0
End If
 buffer = buffer & MSComm1.Input
 i = 0
 If Len(buffer) > 20 Then
   If InStr(1, buffer, "-") Then buffer = "0"
   'escriure_ini "InfoBufferBascula", "Bascula" + atrim(nummaq), buffer, rutadelfitxer(cami) + "valorsprograma.ini"
   If InStr(1, buffer, Chr$(13)) > 0 Then buffer = Mid(buffer, InStr(1, buffer, "+") + 1, InStr(1, buffer, Chr$(13)))
   'If InStr(1, buffer, ".") > 0 Then buffer = Mid(buffer, 1, InStr(1, buffer, ".") - 1) + "," + Mid(buffer, InStr(1, buffer, ".") + 1)
   If atrim(buffer) <> "" Then
        etpesbascula = buffer
   End If
   buffer = ""
 End If
 Exit Sub
nopossarpes:
  nobascula = True
  etpesbascula.BackColor = QBColor(11)
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

Sub no_editar_bobines()
 If bobines.Recordset.EditMode > 0 Then bobines.Recordset.Update
   bobines.UpdateControls
   tempseditant = 0
 
End Sub

Private Sub rellotge_Timer()
  Static tempsoperari As Byte
'  Static ultimarow As Double
'  If ultimarow = 0 Then ultimarow = reixa.Row
'  If ultimarow <> reixa.Row Then
'     ultimarow = reixa.Row: calcular_totals
 ' End If
 'si estic a canvi maquina faig pampalluga al canutu
 If Not Rebobinadores.Recordset.EOF Then
     If Rebobinadores.Recordset!tipus = "C" Then
        canutustallats.Visible = Not canutustallats.Visible
       Else: canutustallats.Visible = True
     End If
       Else: canutustallats.Visible = True
 End If
 If DateDiff("s", tempseditant, Now) > 3 And tempseditant > 0 Then
   no_editar_bobines
 End If
 etproblema.Visible = Not etproblema.Visible
 etbobinesimpost.Visible = Not etbobinesimpost.Visible
 mirarsiparar
 posarpesbascula
 On Error GoTo error_screen
 If controlactiu = "akjdfks" Then Me.Caption = Me.Caption
 On Error GoTo 0
 If client.Caption = "" And (Rebobinadores.Recordset.BOF And Rebobinadores.Recordset.EOF) Then
   carregar_client_ntintersialtres
 End If
 
 If numop = 0 And Not formseleccio.Visible And reixa.Enabled Then
   numop = escullir_operari
   nomoperari = UCase(r)
 End If

 
 If reixa.col = 0 And controlactiu = "reixa" Then
   tempsoperari = cadbl(tempsoperari) + 1
   If tempsoperari > 2 Then reixa.col = 1: tempsoperari = 0
 End If
 If (reixabobines.col = 0 Or reixabobines.col = 1) And controlactiu = "reixabobines" Then
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
  rellotge.Tag = cadbl(rellotge.Tag) + 1
  If rellotge.Tag = "100" Then
    'calcular_totals
    If Not existeix("c:\ordprog.ini") And InStr(1, UCase(Environ("computername")), "EXPEDICIONS") = 0 Then assignardecimalipunt
    rellotge.Tag = "0"
  End If
  
  If Not Rebobinadores.Recordset.EOF Then
    Select Case atrim(Rebobinadores.Recordset!tipus)
       
       Case "C"
          Command1.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor
          Command2.BackColor = &HFF8080
       Case "F"
          Command1.BackColor = Command4.BackColor: Command2.BackColor = Command4.BackColor
          Command3.BackColor = &HFF8080
        Case Else
           Command1.BackColor = Command4.BackColor: Command2.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor
    End Select
    'If Screen.ActiveForm.Name = "capcalera" Then
    '  Command2.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor
    '      Command1.BackColor = &HFF8080
    'End If
  End If
  'Form1.Caption = DateDiff("s", horaapretada, Now)
  'miro si el boto de pantones ha estat apretat mes de 3 segons
  If horaapretada > 0 And DateDiff("s", horaapretada, Now) >= 1 Then
     modificataulapantonesstandard
     horaapretada = 1
  End If
  'copia la bd d'estoc del ser2 al serverprodu
  'If Hour(Now) = 20 And Minute(Now) < 30 And (cadbl(llegir_ini("General", "diacopiafitxstoc", "comandes.ini")) <> Day(Now)) Then
  '  Copiar_Fitxer camistocks, "\\serverprodu\dades\progcomandes\dades\copiaestocinplacsa.mdb"
  '  escriure_ini "General", "diacopiafitxstoc", Day(Now), "comandes.ini"
    
  'End If
  Exit Sub
error_screen:
'MsgBox "Error d'Screen en el Timer"
'End
End Sub
Sub modificataulapantonesstandard()
framepantones.Visible = Not framepantones.Visible
frameempalmes.Visible = False
framebobentrada.Visible = False
dblots.Visible = True
dblots.AllowAddNew = True
dblots.AllowDelete = True
dblots.AllowUpdate = True
dblots.MarqueeStyle = 6
End Sub

Private Sub Text2_Change()

End Sub

Private Sub tpescanutu_LostFocus()
  If controlactiu = "agafarpesbascula" Then
     tpescanutu.Tag = "pesarcanutu"
    Else: tpescanutu.Tag = ""
  End If
  guarda_totals

End Sub
Sub calcularvalorsreducciocilindre(numc As Double, ByVal numerodemaquina As Byte, numformula As Byte)
   Dim rstc As Recordset
   Dim rstclixes As Recordset
   Dim dbclixes As Database
   Dim rstmodifi As Recordset
   Dim desarrollteoric As Double
   Dim desarrollreal As Double
   Dim valorrealmostra As Double
   Dim motius As Double
   Dim a1 As String
   Dim a2 As String
   Dim a3 As String
   Dim a4 As String
   Dim a5 As String
   Dim a6 As String
   
   
   
   numerodemaquina = maquinaquehaimpres(numc)
   If numerodemaquina < 7 Then Exit Sub
   Set rstc = dbtmp.OpenRecordset("select numtreball,microperforat,rebmacroperforat,numordremodificacio,microperforat,rebmacroperforat from comandes where comanda=" + atrim(numc))
   
   If rstc.EOF Then Exit Sub
   id_treball = rstc!numtreball
   ordremodificacio = rstc!numordremodificacio
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
   Set rstclixes = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(cadbl(rstc!numtreball)))
   If cadbl(rstclixes!reduccioxmetre) = 0 Then Exit Sub
   If rstclixes.EOF Then Exit Sub
   Set rstmodifi = dbclixes.OpenRecordset("select desarroll from modificacions where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordre=" + atrim(cadbl(rstc!numordremodificacio)))
   If rstmodifi.EOF Then Exit Sub
   If cadbl(rstmodifi!desarroll) = 0 Then Exit Sub
   motius = Redondejar(1000 / cadbl(rstmodifi!desarroll), 0)
   desarrollteoric = motius * cadbl(rstmodifi!desarroll)
   desarrollreal = Redondejar((desarrollteoric * cadbl(rstclixes!reduccioxmetre)) / 1000, 1)
   valorrealmostra = Redondejar(desarrollteoric + desarrollreal, 1)
   
   avisapantalla = "Distorsió cilindre per metre: " + atrim(cadbl(rstclixes!reduccioxmetre)) + " mm" + Chr(10) + "Factor: " + atrim(cadbl(IIf(numerodemaquina = 7, rstclixes!redcilindrefw, rstclixes!redcilindref2))) + " mm"
   ' posso els valors al report llistat
   a1 = passaradecimalpunt(atrim(rstclixes!reduccioxmetre))
   a2 = passaradecimalpunt(atrim((IIf(numerodemaquina = 7, rstclixes!redcilindrefw, rstclixes!redcilindref2))))
   a3 = passaradecimalpunt(atrim(desarrollteoric))
   a4 = passaradecimalpunt(atrim(motius))
   a5 = passaradecimalpunt(atrim(desarrollreal))
   a6 = passaradecimalpunt(atrim(valorrealmostra))
   If Not vperforat Then substituir "Verificar perforado.", "": substituir "X11,463,8,41,490", ""
   
   preparar_etiqueta_verificacioreducciocilindre numc, numop, a1, a2, a3, a4, a5, a6
   imprimir_etiqueta_zebra True
   'llistat.Formulas(numformula) = "reducciopermetrelineal=" + passaradecimalpunt(atrim(rstclixes!reduccioxmetre))
   'numformula = numformula + 1
   'llistat.Formulas(numformula) = "parametrereduccio=" + passaradecimalpunt(atrim((IIf(numerodemaquina = 7, rstclixes!redcilindrefw, rstclixes!redcilindref2))))
   'numformula = numformula + 1
   'llistat.Formulas(numformula) = "desarrollteoric=" + passaradecimalpunt(atrim(desarrollteoric))
   'numformula = numformula + 1
   'llistat.Formulas(numformula) = "motius=" + passaradecimalpunt(atrim(motius))
   'numformula = numformula + 1
   'llistat.Formulas(numformula) = "desarrollreal=" + passaradecimalpunt(atrim(desarrollreal))
   'numformula = numformula + 1
   'llistat.Formulas(numformula) = "valorrealmostra=" + passaradecimalpunt(atrim(valorrealmostra))
   'numformula = numformula + 1
fi:
   Set dbclixes = Nothing
   Set rstclixes = Nothing
   Set rstmodifi = Nothing
End Sub

Function maquinaquehaimpres(numc As Double) As Byte
   Dim rst As Recordset
   maquinaquehaimpres = 0
   Set rst = dbbaixes.OpenRecordset("select * from impressores where comanda=" + atrim(numc))
   If Not rst.EOF Then maquinaquehaimpres = cadbl(rst!numeromaquina)
   
End Function
Sub preparar_etiqueta_verificacioreducciocilindre(numc As Double, numop As Byte, reducciopermetre As String, parametrereduccio As String, desarrollteric As String, motius As String, desarrollreal As String, valorrealmostra As String)
   Dim rst As Recordset
   Dim ultimalinia As String
   Dim rstproducte As Recordset
   Dim rstm As Recordset
   Dim rstc As Recordset
   Set rst = dbtmp.OpenRecordset("select client, producte,impressio,refclient,numordremodificacio,numtreball from comandes where comanda=" + atrim(numc))
   Set rstproducte = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(rst!producte) + "'")
   If rstproducte.EOF Then Exit Sub
   Set rstc = dbtmp.OpenRecordset("select * from clients where codi=" + atrim(rst!client))
   If rstc.EOF Then Exit Sub
   Set rstm = dbtmpb.OpenRecordset("SELECT comanda, numeromaquina FROM REBOBINADORES where comanda=" + atrim(numc))
   If rstm.EOF Then Exit Sub
   Set rstm = dbtmp.OpenRecordset("select descripcio from maquines where maquina='R' and codi=" + atrim(rstm!numeromaquina))
   If rstm.EOF Then Exit Sub
   
   Open llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetarqualitatreducciocilindrerebobinadores.prn" For Input As #1
   linia.Text = Input(LOF(1), #1)
   Close #1
   With rsttmp
   substituir "#DATA#", Format(Now, "dd/mm/yy")
   substituir "#NOMMAQUINA#", atrim(rstm!descripcio)
   'substituir "#TREBALL#", atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio)
   substituir "#LOT#", atrim(numc)
   substituir "#CLIENT#", Mid(atrim(rstc!nom), 1, 30)
   substituir "#METRELINEAL#", atrim(reducciopermetre)
   substituir "#PARAMETREREDUCCIO#", atrim(parametrereduccio)
   substituir "#DESARROLLTEORIC#", atrim(desarrollteric)
   substituir "#MOTIUS#", atrim(motius)
   substituir "#DESARROLLREAL#", atrim(desarrollreal)
   substituir "#VALORREALMOSTRA#", atrim(valorrealmostra)
   substituir "#LINIA#", "Operari: " + atrim(numop)
   
   End With
      
   
End Sub

Sub actualitzarestatbobinesdesbobinadors()

End Sub

