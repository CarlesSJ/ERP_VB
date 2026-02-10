VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   Caption         =   "Baixes Soldadores"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   11895
   ClipControls    =   0   'False
   Icon            =   "baixes soldadores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command32 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Capçalera"
      Enabled         =   0   'False
      Height          =   525
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   144
      Top             =   210
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   1155
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Avaria"
      Enabled         =   0   'False
      Height          =   525
      Left            =   3675
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   210
      Width           =   795
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Parada"
      Enabled         =   0   'False
      Height          =   525
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   210
      Width           =   750
   End
   Begin VB.Frame framebobentrada 
      Caption         =   "Bobines Entrada"
      Height          =   3315
      Left            =   6555
      TabIndex        =   67
      Top             =   4170
      Visible         =   0   'False
      Width           =   3435
      Begin VB.CheckBox carrastrar2bobs 
         Caption         =   "2 Bobs"
         Height          =   195
         Left            =   2520
         TabIndex        =   145
         ToolTipText     =   "Arrastrar 2 bobines d'entrada"
         Top             =   3045
         Width           =   840
      End
      Begin VB.CommandButton Command24 
         Height          =   480
         Left            =   1875
         Picture         =   "baixes soldadores.frx":048A
         Style           =   1  'Graphical
         TabIndex        =   127
         ToolTipText     =   "Ensenyar bobines d'entrada si utilitzades."
         Top             =   2775
         Width           =   585
      End
      Begin VB.CommandButton eliminarbobentrada 
         Height          =   480
         Left            =   1260
         Picture         =   "baixes soldadores.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "Eliminar bobina d'entrada"
         Top             =   2775
         Width           =   585
      End
      Begin VB.CommandButton Command23 
         Height          =   480
         Left            =   645
         Picture         =   "baixes soldadores.frx":0F9E
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "Afegir manualment el Palet/Bobina d'entrada"
         Top             =   2775
         Width           =   585
      End
      Begin VB.CommandButton botoensenyarpacking 
         Height          =   480
         Left            =   15
         Picture         =   "baixes soldadores.frx":1528
         Style           =   1  'Graphical
         TabIndex        =   124
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
         TabIndex        =   117
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
         TabIndex        =   111
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
         TabIndex        =   110
         Top             =   1740
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CheckBox ensenyartoteslesbobines 
         Caption         =   "Totes"
         Height          =   195
         Left            =   2520
         TabIndex        =   76
         Top             =   2805
         Width           =   720
      End
      Begin MSDBGrid.DBGrid bobentrada 
         Bindings        =   "baixes soldadores.frx":1AB2
         Height          =   2520
         Left            =   45
         OleObjectBlob   =   "baixes soldadores.frx":1AC7
         TabIndex        =   104
         Top             =   195
         Width           =   3330
      End
   End
   Begin VB.CommandButton Command29 
      Height          =   390
      Left            =   10725
      Picture         =   "baixes soldadores.frx":24B9
      Style           =   1  'Graphical
      TabIndex        =   136
      ToolTipText     =   "Document descriptiu de com ha de ser la sortida"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command28 
      Height          =   390
      Left            =   11115
      Picture         =   "baixes soldadores.frx":2A43
      Style           =   1  'Graphical
      TabIndex        =   135
      ToolTipText     =   "Calcul diametre"
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton imprimir 
      Height          =   360
      Left            =   11355
      Picture         =   "baixes soldadores.frx":2FCD
      Style           =   1  'Graphical
      TabIndex        =   134
      TabStop         =   0   'False
      ToolTipText     =   "Imprimir Etiqueta Mostra Client"
      Top             =   5205
      Width           =   375
   End
   Begin VB.CommandButton botodescansrelleu 
      Height          =   390
      Left            =   11520
      Picture         =   "baixes soldadores.frx":3557
      Style           =   1  'Graphical
      TabIndex        =   133
      ToolTipText     =   "Control Descans i Relleu"
      Top             =   960
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
      Picture         =   "baixes soldadores.frx":3AE1
      Style           =   1  'Graphical
      TabIndex        =   131
      ToolTipText     =   "Imprimeix el Packing-List"
      Top             =   75
      Visible         =   0   'False
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
      Picture         =   "baixes soldadores.frx":406B
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   150
      Width           =   465
   End
   Begin VB.CommandButton maquina 
      BackColor       =   &H00FF8080&
      Caption         =   "Maq: 0"
      Height          =   465
      Left            =   7185
      Style           =   1  'Graphical
      TabIndex        =   121
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
      Picture         =   "baixes soldadores.frx":488D
      Style           =   1  'Graphical
      TabIndex        =   120
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
      TabIndex        =   119
      Top             =   75
      Width           =   810
   End
   Begin VB.TextBox linia 
      Height          =   360
      Left            =   8430
      MaxLength       =   65000
      ScrollBars      =   2  'Vertical
      TabIndex        =   118
      Text            =   $"baixes soldadores.frx":4E17
      Top             =   -75
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Frame Frame3 
      Caption         =   "Kg"
      Height          =   570
      Left            =   11670
      TabIndex        =   112
      Top             =   1365
      Visible         =   0   'False
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
         TabIndex        =   113
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
      Left            =   6765
      TabIndex        =   94
      Top             =   7665
      Width           =   5115
      Begin VB.TextBox unitatsxfunda 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   3690
         TabIndex        =   115
         Top             =   405
         Width           =   465
      End
      Begin VB.TextBox tpescanutu 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   4320
         TabIndex        =   108
         Top             =   390
         Width           =   465
      End
      Begin VB.TextBox bandes 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   300
         TabIndex        =   99
         Top             =   450
         Width           =   435
      End
      Begin VB.TextBox amplemerma 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   2700
         TabIndex        =   98
         Top             =   435
         Width           =   840
      End
      Begin VB.TextBox ampleref 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   885
         TabIndex        =   97
         Top             =   450
         Width           =   840
      End
      Begin VB.TextBox bandesm 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1770
         TabIndex        =   96
         Top             =   435
         Width           =   840
      End
      Begin VB.CheckBox comandaacavada 
         Caption         =   "Acavada"
         Enabled         =   0   'False
         Height          =   225
         Left            =   3930
         TabIndex        =   95
         Top             =   -30
         Width           =   1005
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "UxFunda"
         Height          =   210
         Left            =   3600
         TabIndex        =   116
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "P.Canutu"
         Height          =   210
         Left            =   4245
         TabIndex        =   109
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Bandes M"
         Height          =   210
         Left            =   1815
         TabIndex        =   103
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Ample Merma"
         Height          =   210
         Left            =   2610
         TabIndex        =   102
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ample Ref"
         Height          =   210
         Left            =   885
         TabIndex        =   101
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Simult."
         Height          =   210
         Left            =   45
         TabIndex        =   100
         Top             =   180
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
      Picture         =   "baixes soldadores.frx":4E3B
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
      Left            =   8835
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   75
      Width           =   1080
   End
   Begin VB.Frame calculant 
      Height          =   2580
      Left            =   -15
      TabIndex        =   70
      Top             =   8415
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
      Picture         =   "baixes soldadores.frx":5F5D
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
      Picture         =   "baixes soldadores.frx":63D7
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
      DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\baixes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10770
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "bobinessol"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Frame Frame2 
      Caption         =   "Totals"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   7665
      Width           =   6585
      Begin VB.TextBox hparada 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1350
         TabIndex        =   141
         Top             =   390
         Width           =   645
      End
      Begin VB.TextBox havaria 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2055
         TabIndex        =   140
         Top             =   390
         Width           =   690
      End
      Begin VB.TextBox hmaquina 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   390
         Width           =   540
      End
      Begin VB.TextBox hfunc 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   390
         Width           =   570
      End
      Begin VB.TextBox tkilos 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4725
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox tunitats 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3810
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox tunitatshora 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   5610
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox tsacs 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2955
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   390
         Width           =   840
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "H.Parada"
         Height          =   210
         Left            =   1350
         TabIndex        =   143
         Top             =   150
         Width           =   990
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "H.Avaria"
         Height          =   195
         Left            =   1920
         TabIndex        =   142
         Top             =   135
         Width           =   990
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "H. Màq."
         Height          =   210
         Left            =   75
         TabIndex        =   21
         Top             =   165
         Width           =   675
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "H. Func."
         Height          =   195
         Left            =   690
         TabIndex        =   20
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sacs/Caixes"
         Height          =   210
         Left            =   2940
         TabIndex        =   19
         Top             =   150
         Width           =   990
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "T.Unitats"
         Height          =   210
         Left            =   3930
         TabIndex        =   18
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Unitats/Hora"
         Height          =   210
         Left            =   5550
         TabIndex        =   17
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Kilos"
         Height          =   210
         Left            =   4740
         TabIndex        =   16
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
      Height          =   330
      Left            =   2265
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
      Left            =   5985
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   210
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Canvi"
      Enabled         =   0   'False
      Height          =   525
      Left            =   5250
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   210
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Capçalera"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7005
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Data soldadores 
      Caption         =   "soldadores"
      Connect         =   "Access"
      DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\baixes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "soldadores"
      Top             =   735
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "baixes soldadores.frx":70AD
      Height          =   2235
      Left            =   345
      OleObjectBlob   =   "baixes soldadores.frx":70C2
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
      Text            =   "135716"
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
         TabIndex        =   137
         Top             =   990
         Visible         =   0   'False
         Width           =   3405
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Lots"
         Height          =   750
         Left            =   10260
         Picture         =   "baixes soldadores.frx":8CF3
         Style           =   1  'Graphical
         TabIndex        =   129
         TabStop         =   0   'False
         ToolTipText     =   "Bosses  i Caixes utilitzades per embossar les bosses."
         Top             =   2505
         Width           =   960
      End
      Begin VB.CheckBox mostracli 
         Caption         =   "Mostra Cli."
         Height          =   195
         Left            =   10350
         TabIndex        =   122
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
         Picture         =   "baixes soldadores.frx":A18A
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Agafar el pes de la bàscula"
         Top             =   2490
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   10260
         Picture         =   "baixes soldadores.frx":C5E4
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
         Picture         =   "baixes soldadores.frx":DCB6
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
         Picture         =   "baixes soldadores.frx":ED00
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
         Bindings        =   "baixes soldadores.frx":10902
         Height          =   3225
         Left            =   180
         OleObjectBlob   =   "baixes soldadores.frx":10914
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
         Bindings        =   "baixes soldadores.frx":117F0
         Height          =   3525
         Left            =   60
         OleObjectBlob   =   "baixes soldadores.frx":11803
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
         TabIndex        =   77
         Top             =   3330
         Width           =   3210
      End
      Begin MSDBGrid.DBGrid dblots 
         Bindings        =   "baixes soldadores.frx":123B4
         Height          =   3705
         Left            =   30
         OleObjectBlob   =   "baixes soldadores.frx":123C3
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
         TabIndex        =   78
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
      TabIndex        =   80
      Top             =   3540
      Width           =   11670
      Begin VB.CommandButton Command25 
         Height          =   360
         Left            =   9120
         Picture         =   "baixes soldadores.frx":12DA1
         Style           =   1  'Graphical
         TabIndex        =   128
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
         TabIndex        =   106
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
         TabIndex        =   93
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
         TabIndex        =   92
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
         TabIndex        =   91
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
         Index           =   8
         Left            =   7485
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
         Index           =   7
         Left            =   6690
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
         Index           =   6
         Left            =   5895
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
         Index           =   5
         Left            =   5100
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
         Index           =   4
         Left            =   4305
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
         Index           =   3
         Left            =   3510
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
         Index           =   2
         Left            =   2715
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
         Index           =   1
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   82
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
         TabIndex        =   81
         Top             =   135
         Width           =   795
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Pes P."
         Height          =   270
         Left            =   30
         TabIndex        =   107
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
      TabIndex        =   132
      Top             =   1140
      Width           =   3735
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
      TabIndex        =   130
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
      TabIndex        =   114
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
      Left            =   30
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
      Left            =   615
      TabIndex        =   1
      Top             =   0
      Width           =   1260
   End
   Begin VB.Label proces 
      Height          =   315
      Left            =   0
      TabIndex        =   79
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Sub calcular_totals(Optional obrint As Boolean)
  Dim total As Double
  Dim hores As Double
  Dim bkimp As Double
  Dim bkbob As Double
  Dim bkultim As Double
  barraestat.caption = "Calculant els totals..."
  'calculant.Visible = True
  fcalculant.Show 0, Me
  calculant.Top = 2222
  DoEvents
  
  'On Error GoTo fi
  reixa.EditActive = False
  reixabobines.EditActive = False
  If soldadores.Recordset.EOF Or cadbl(soldadores.Recordset!id) = 0 Then GoTo fi
  
  '---- guardo la posicio de linies imp i de bobina x recuperarlames avall
  If soldadores.Recordset!tipus = "F" Then bkimp = atrim(cadbl(soldadores.Recordset!id))
  If Not bobines.Recordset.EOF Then bkbob = atrim(cadbl(bobines.Recordset!numerodesac))
  '------
  
  'On Error Resume Next
  soldadores.Recordset.MoveLast
  bkultim = atrim(cadbl(soldadores.Recordset!id))
  soldadores.Recordset.MoveFirst
  While Not soldadores.Recordset.EOF
   'On Error GoTo 0
   If soldadores.Recordset!tipus = "F" Then
    If soldadores.Recordset.EditMode = 0 Then soldadores.Recordset.Edit
    Set rsttmp = dbtmpb.OpenRecordset("select count(*) as elgran from bobinessol where controlid=" + atrim(soldadores.Recordset!id))
    If Not rsttmp.EOF Then soldadores.Recordset!totalsacs = rsttmp!elgran
  
'    Set rsttmp = dbtmpb.OpenRecordset("select sum(kilos) as elgran from bobinessol where controlid=" + atrim(soldadores.Recordset!id))
'    If Not rsttmp.EOF Then soldadores.Recordset!totalkilos = rsttmp!elgran
  
    Set rsttmp = dbtmpb.OpenRecordset("select sum(unitatsxsac) as elgran from bobinessol where controlid=" + atrim(soldadores.Recordset!id))
    If Not rsttmp.EOF Then soldadores.Recordset!totalunitats = rsttmp!elgran
  
    Set rsttmp = dbtmpb.OpenRecordset("select id,unitatsxsac from bobinessol where unitatsxsac=0 and controlid=" + atrim(soldadores.Recordset!id))
    If Not rsttmp.EOF Then
     If rsttmp!id <> bobines.Recordset!id Then MsgBox "Hi ha algun sac sense unitats posades"
    End If
    soldadores.Recordset.Update
   End If
  
   
   With soldadores.Recordset
    total = 0
    'On Error Resume Next
     If Not IsDate(CVDate(atrim(!datainici))) Or Not IsDate(CVDate(atrim(!horainici))) Or Not IsDate(atrim(!horafi)) Or Not IsDate(atrim(!datafi)) Then
      If Not obrint And soldadores.Recordset!id <> bkimp And soldadores.Recordset!id <> bkultim Then MsgBox "Error d'hora d'inici o final de funcionament. Corretgeix l'error per poder continuar correctament."
       Else
            total = DateDiff("n", CVDate(atrim(!datainici) + " " + atrim(!horainici)), CVDate(atrim(!datafi) + " " + atrim(!horafi)))
            total = Format(total / 60, "#,##0.00")
            
     End If
    If soldadores.Recordset.EditMode = 0 Then soldadores.Recordset.Edit
     soldadores.Recordset!totalhores = total
     soldadores.Recordset.Update
   End With
  soldadores.Recordset.MoveNext
 Wend
  'If Not rsttmp.EOF Then
  'impresores.UpdateControls
  'impresores.UpdateRecord
  'reixa.Refresh
  
  On Error GoTo 0
  ensenyar_totalstotals
  possar_metres_min
  Set rstmp = Nothing
  barraestat.caption = ""
  
  '---recupero la pocisio de linis imp i de bob
   If bkimp > 0 Then
     soldadores.Recordset.FindFirst "id=" + atrim(bkimp)
     bobines.Recordset.FindFirst "numerodesac=" + atrim(bkbob)
   Else: soldadores.Recordset.MoveLast
  End If
  '---
fi:
'calculant.Visible = False
barraestat.caption = ""
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
tbob = 0: hfunc = 0: hparada = 0: havaria = 0: hclixe = 0: hmaquina = 0: hajusts = 0: tkilos = 0: tmetres = 0: tprova = 0:
'total bobines
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalsacs) as elgran from soldadores where comanda=" + atrim(cadbl(comanda.Text)))
  If Not rsttmp.EOF Then tbob = cadbl(rsttmp!elgran)

  
'hores func
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from soldadores  where comanda=" + atrim(cadbl(comanda.Text)) + " and tipus='F'")
  If Not rsttmp.EOF Then hfunc = cadbl(rsttmp!elgran)
'hores maquina
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from soldadores  where comanda=" + atrim(cadbl(comanda.Text)) + " and tipus='C'")
  If Not rsttmp.EOF Then hmaquina = cadbl(rsttmp!elgran)
'hores parada
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from soldadores  where comanda=" + atrim(cadbl(comanda.Text)) + " and tipus='P'")
  If Not rsttmp.EOF Then hparada = cadbl(rsttmp!elgran)
'hores avaria
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from soldadores  where comanda=" + atrim(cadbl(comanda.Text)) + " and tipus='A'")
  If Not rsttmp.EOF Then havaria = cadbl(rsttmp!elgran)


  
'total metres
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalunitats) as elgran from soldadores  where comanda=" + atrim(cadbl(comanda.Text)))
  If Not rsttmp.EOF Then tunitats = cadbl(rsttmp!elgran)
  

  guarda_totals
  ensenya_totals
End Sub

Sub guarda_totals()
Set rsttmp = dbtmpb.OpenRecordset("select * from soldadorestot where comanda=" + atrim(cadbl(comanda)))
  If rsttmp.EOF Then
      rsttmp.AddNew
    Else: rsttmp.Edit
  End If
  With rsttmp
    '!firmat = atrim(firmat.Caption)
    !comanda = cadbl(comanda)
    !hcanvi = cadbl(hmaquina)
    !hfuncio = cadbl(hfunc)
    !hparada = cadbl(hparada)
    !havaria = cadbl(havaria)
    '!sacsocaixes = cadbl(tunitats)
   ' !tkilos = cadbl(tkilos)
    !unitatshora = cadbl(tunitatshora)
    '!kiloshora = cadbl(kiloshora)
    !simultaneitat = cadbl(bandes)
   ' !amplebob = cadbl(amplebob)
   ' !espesor = cadbl(espesor)
   ' !ampleref = cadbl(ampleref)
   ' !bandesmerma = cadbl(bandesm)
   ' !amplemerma = cadbl(amplemerma)
    !acavada = cadbl(comandaacavada.Value)
   ' pescanutu = cadbl(tpescanutu.Text)
   ' !pescanutu = pescanutu
    !unitatsperfunda = cadbl(unitatsxfunda)
    '!mostraclient = IIf(mostracli.Value > 0, True, False)
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
Set rsttmp = dbtmpb.OpenRecordset("select * from soldadorestot where comanda=" + atrim(cadbl(comanda)))

  With rsttmp
    'comanda = atrim(!comanda)
    'firmat = atrim(!firmat)
    hmaquina = atrim(!hcanvi)
    hfunc = atrim(!hfuncio)
    tsacs = atrim(!tsacs)
    'tprova = atrim(!tprova)
    'tkilos = atrim(!tkilos)
    tunitats = atrim(!tunitats)
    tunitatshora = atrim(!unitatshora)
    comandaacavada.Value = cadbl(!acavada)
 '   If pescanutu = 0 Then pescanutu = atrim(cadbl(!pescanutu))
 '   tpescanutu = pescanutu
    unitatsxfunda = cadbl(!unitatsperfunda)
    bandes = atrim(!simultaneitat)
   ' amplebob = atrim(!amplebob)
   ' espesor = atrim(!espesor)
   ' ampleref = atrim(!ampleref)
   ' bandesm = atrim(!simultaneitat)
   ' amplemerma = atrim(!amplemerma)
   ' mostracli.Value = IIf(cadbl(!mostraclient), 1, 0)
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
If cadbl(tunitats.tag) > 0 Then
  If cadbl(tunitats.tag) * cadbl(bandes) < cadbl(tunitats) Then
      etproblema.caption = "Mes Unitats que a la comanda. " + tunitats.tag + " Unitats"
       Else: etproblema.caption = ""
  End If
End If
'If cadbl(tkilos.Tag) > 0 Then
'  If cadbl(tkilos) > cadbl(tkilos.Tag) Then
'      etproblema.Caption = "Mes Kilos que a la comanda. " + tkilos.Tag + " Kilos"
'       Else: etproblema.Caption = ""
'  End If
'End If

End Sub

Private Sub AcroPDF2_GotFocus()

End Sub

Sub tamany_visualitzadorpdf(vtamanygran As Boolean)
'  If vtamanygran Then
     AcroPDF1.visible = Not AcroPDF1.visible
     AcroPDF1.width = 11000
     AcroPDF1.Height = 6500
     AcroPDF1.Left = 700
     AcroPDF1.ZOrder
     framebobentrada.visible = Not AcroPDF1.visible
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
 If pespalet.tag = "pesar" Then
       pespalet.Text = atrim(llegirpesbascula): gravar_pespalet: pespalet.tag = "": pespalet.SetFocus: Exit Sub
 End If
 If tpescanutu.tag = "pesarcanutu" Then
   If MsgBox("Estas segur que vols possar " + atrim(llegirpesbascula) + "Kg com a pes del canutu?" + Chr(10) + "Amb aixó les bobines pesaran pes net i pes brut.", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
      tpescanutu.Text = atrim(llegirpesbascula): guarda_totals: tpescanutu.tag = "": pescanutu = cadbl(tpescanutu.Text): tpescanutu.SetFocus: Exit Sub
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
       If bobines.Recordset.Fields(camp).Type <> 8 Then
         bobines.Recordset.Fields(camp) = reixabobines.Columns(camp)
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
  If proces.tag = "E" Then
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
  
  If proces.tag = "I" Then taulabob = "bobinesimp": Set rsttmpimp = dbtmpb.OpenRecordset("select * from impressores where tipus='F' and comanda=" + atrim(cadbl(comanda)))
  If proces.tag = "L" Then
    If cadbl(vlink3) = 0 Then
      r = comanda
       Else: r = vlink3
    End If
    taulabob = "bobineslam": Set rsttmpimp = dbtmpb.OpenRecordset("select * from laminadores where tipus='F' and comanda=" + atrim(cadbl(r)))
  End If
  While Not rsttmpimp.EOF
    If proces.tag = "I" Then Set rsttmpbobimp = dbtmpb.OpenRecordset("select * from bobinesimp where " + ensenyar + " controlid=" + atrim(cadbl(rsttmpimp!id)))
    If proces.tag = "L" Then Set rsttmpbobimp = dbtmpb.OpenRecordset("select * from bobineslam where " + ensenyar + " controlid=" + atrim(cadbl(rsttmpimp!id)))
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
  Set dbtmpb = OpenDatabase(soldadores.DatabaseName)
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
   formseleccio.caption = "Selecció bobina d'entrada"
   formseleccio.refrescar
   'formseleccio.DBGrid2.Columns(4).Visible = False
   formseleccio.DBGrid2.Columns(0).visible = False
   formseleccio.DBGrid2.Columns(1).visible = False
   formseleccio.DBGrid2.Columns(2).width = 2500
   formseleccio.DBGrid2.Columns(3).width = 2500
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
  'If afegint Then GoTo cont
r = ""
 bk = atrim(bobines.Recordset!numerodesac)
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
 'If Not bobines.Recordset.EOF Then
 ' dbtmpb.Execute "update bobinessol set bobsent='" + atrim(r) + "' where id=" + atrim(bobines.Recordset!id)
 ' bobines.Refresh
 'End If
  
  
  
'  If bobinesent.Recordset.RecordCount > 1 Then
'    Set clon = bobinesent.Recordset.Clone
'    clon.MoveLast
'    clon.MovePrevious
'    marcarfidebobina cadbl(clon!palet), cadbl(clon!bobina)
'  End If
 'End If
 If bk <> "" Then
     bobines.Recordset.FindFirst "numerodesac=" + bk
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
           r = IIf(proces.tag <> "L", "bobinesimp", "bobineslam")
           dbtmpb.Execute "update  " + r + " set utilitzadaabaixa=True where id=" + atrim(cadbl(rsttmp2!idbobina))
      End If
              
      Else
       bobinesent.Recordset!paletobobina = LCase(bobinesent.Recordset!paletobobina)
       If UCase$(bobinesent.Recordset!paletobobina) = "P" Then
          dbstocks.Execute "update  bobines set utilitzadaabaixa=False where idpalet=" + bobentrada.Columns(0) + " and idbobina=" + bobentrada.Columns(1)
        Else:
          r = IIf(proces.tag <> "L", "bobinesimp", "bobineslam")
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
  Set rst = dbtmpb.OpenRecordset("select * from bobentradatmpsol" + atrim(nummaq))
creartaula:
  'ample double,plegat double,solapa double,espessor double,metres double,kilos double)"
  dbtmpb.Execute ("delete * from bobentradatmpsol" + atrim(nummaq))
  Set rst = Nothing
  Exit Sub
borrar:
  'dbtmpb.Execute "drop table bobentradatmpreb" + atrim(nummaq)
  dbtmpb.Execute ("create table bobentradatmpsol" + atrim(nummaq) + " (" + camps) + ")"
  GoTo creartaula
  
End Sub
Sub possar_camps_generals()
  Dim rsttmp As Recordset
  If cadbl(bandes) = 0 Then
      Set rsttmp = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(comanda.Text)))
      If Not rsttmp.EOF Then
         With rsttmp
          If cadbl(bandes) = 0 Then
             bandes = atrim(!simulteneitatsol)
             If cadbl(bandes) = 0 Then bandes = "1"
          End If
          If cadbl(amplebob) = 0 Then amplebob = atrim(!amplesol)
          If cadbl(espesor) = 0 Then espesor = atrim(espessorsol)
          
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
 If Not formseleccio.visible Then bobinesent.UpdateRecord
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
       bobinesent.RecordSource = "select * from bobinesentsol where id=99999999"
     Else
       bobinesent.RecordSource = "select * from bobinesentsol where id=" + atrim(cadbl(bobines.Recordset!id)) + " order by id_entrada"
   End If
   bobinesent.Refresh
 End If
 
End Sub

Private Sub clixes_Click()
 
End Sub
Sub finalitza_seccio()
  On Error GoTo fi
  If soldadores.Recordset.EOF Then Exit Sub
  On Error Resume Next
  soldadores.Recordset.MoveLast
  If IsDate(soldadores.Recordset!datafi) Then r = "no": Exit Sub
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
  soldadores.Recordset.Edit
  soldadores.Recordset!datafi = Date: soldadores.Recordset!horafi = Time
  Select Case soldadores.Recordset!tipus
   Case "C"
   Case "M"
   Case "A"
   Case "F"
  End Select
  soldadores.Recordset.Update
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
  If bobines.tag <> "" Then
    Set rstcp = dbtmpb.OpenRecordset("select distinct palet from bobinesreb where  controlid in(" + bobines.tag + ")")
    While Not rstcp.EOF
     For i = 0 To 9
       If cadbl(botopalets(i).caption) = cadbl(rstcp!palet) Then botopalets(i).BackColor = QBColor(9)
     Next i
     rstcp.MoveNext
    Wend
  End If
  Set rstcp = Nothing
End Sub

Private Sub botodescansrelleu_Click()
   Load formdescansirelleu
   If Not soldadores.Recordset.EOF Then
        soldadores.Recordset.MoveLast
        If Not soldadores.Recordset.EOF Then
           If Not IsDate(soldadores.Recordset!datafi) And Not IsDate(soldadores.Recordset!horafi) Then
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
   formdescansirelleu.etnomoperari.tag = atrim(numop)
   formdescansirelleu.Show 1
End Sub
Sub possarliniadefinalitzacio()
    If Not soldadores.Recordset.EOF Then
        soldadores.Recordset.MoveLast
        If Not IsDate(atrim(soldadores.Recordset!datafi)) Or Not IsDate(atrim(soldadores.Recordset!horafi)) Then
           'If MsgBox("No hi ha la hora de fi de funcionament, Vols que el col.loqui automàticament?", vbInformation + vbYesNo, "Atenció") = vbYes Then
            soldadores.Recordset.Edit
            soldadores.Recordset!datafi = Date
            soldadores.Recordset!horafi = Time
            soldadores.Recordset.Update
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
 carregar_bobinesdentrada "ensenyar" + utilitzades, 1, palet, bobina, ncomanda, , ncomanda2, IIf(proces.tag = "invertit", True, False)
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
    afegir_bobentradasol palet, bobina
    'imprimir_controlqualitatVQ cadbl(comanda)    HE PASSAT AIXÓ AL APRETAR CANVI MAQUINA DEMANAT PER PACO
   
   
   'imprimiretiquetaverificacio cadbl(bobines.Recordset!numerodebobina) + (i - 1)
   
    
 End If
fi:
 botoensenyarpacking.tag = ""
 bobinesent.UpdateRecord
 possarnumbobent
' If espesor.Text <> espesorbobina Then
'   espesor.Text = espesorbobina
'   guarda_totals
' End If
 
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
   Set rstm = dbtmpb.OpenRecordset("SELECT comanda, numeromaquina FROM soldadores where comanda=" + atrim(numc))
   If rstm.EOF Then Exit Function
   Set rstm = dbtmp.OpenRecordset("select descripcio from maquines where maquina='R' and codi=" + atrim(rstm!numeromaquina))
   If rstm.EOF Then Exit Function
   ultimalinia = "Op: " + atrim(numop) + "    NºBob.Salida: 0   Fecha: " + Format(Now, "dd/mm/yy")
   
   
   Open llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetarqualitatVQsoldadores.prn" For Input As #1
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
   Set rstm = dbtmpb.OpenRecordset("SELECT comanda, numeromaquina FROM soldadores where comanda=" + atrim(numc))
   If rstm.EOF Then Exit Function
   Set rstm = dbtmp.OpenRecordset("select descripcio from maquines where maquina='R' and codi=" + atrim(rstm!numeromaquina))
   If rstm.EOF Then Exit Function
   ultimalinia = "Op: " + atrim(numop) + "    NºBob.Salida: 0   Fecha: " + Format(Now, "dd/mm/yy")
   
   Open llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetarqualitatbob0soldadores.prn" For Input As #1
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


Sub afegir_bobentradasol(palet As Double, bobina As Double)
        marcaranteriorscomagastades
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
 If Index >= 0 Then numpalet = cadbl(botopalets(Index).caption)
 botopalets(0).tag = Trim(Index)
 pespalet.Text = ""
 'carrego els pesos dels palets
 Set rstpespalet = dbtmpb.OpenRecordset("select * from sol_pespalets where numpalet=" + atrim(numpalet) + " and comanda=" + atrim(cadbl(comanda.Text)))
 If Not rstpespalet.EOF Then
  If cadbl(rstpespalet!pespalet) > 0 Then
     pespalet.Text = rstpespalet!pespalet
  End If
 End If
 If soldadores.Recordset.EOF Then Exit Sub
 If rstpespalet.EOF And soldadores.Recordset!tipus = "F" And controlactiu = "botopalets" Then
      pespalet.Text = "0"
      demanar_el_pespalet
 End If

'ensenyo les bobines
If Not soldadores.Recordset.EOF Then
      ensenya_les_bobines
       Else: Exit Sub
 End If
 posacoloralsquehihaalgu
  If Index >= 0 Then botopalets(Index).BackColor = QBColor(14)

End Sub
Sub demanar_el_pespalet()
      While cadbl(pespalet.Text) = 0
        pespalet.Text = cadbl(InputBox("Has de possar el pes del palet.", "Possar pes del palet.", "23"))
        If cadbl(pespalet.Text) < 6 Or cadbl(pespalet.Text) > 30 Then
           MsgBox "Aquest pes de palet no pot ser correcte.", vbCritical + vbOKOnly, "Error de pes de palet"
           pespalet.Text = "0"
        End If
      Wend
      gravar_pespalet
End Sub
Function controlactiu() As String
  On Error Resume Next
  controlactiu = Form1.ActiveControl.Name
End Function
Sub gravar_pespalet()
   If numpalet < 0 Or numpalet > 30 Then Exit Sub
   'On Error Resume Next
   dbtmpb.Execute "insert into sol_pespalets (comanda,numpalet,pespalet) values (" + atrim(cadbl(comanda.Text)) + "," + atrim(numpalet) + "," + passaradecimalpunt(cadbl(pespalet.Text)) + ")"
   'On Error GoTo 0
   dbtmpb.Execute "update  sol_pespalets set pespalet=" + passaradecimalpunt(cadbl(pespalet.Text)) + " where numpalet=" + atrim(numpalet) + " and comanda=" + atrim(cadbl(comanda.Text))
End Sub
Sub netejar_botons_palets()
 For i = 0 To 9
    botopalets(i).BackColor = Command4.BackColor
 Next i
End Sub

Private Sub comanda_GotFocus()
  Dim vnumc As String
  Dim vnumc_anterior As String
  If nummaq = 0 Then MsgBox "Escull una màquina primer.", vbCritical, "Atenció": Exit Sub
  vnumc = cadbl(InputBox("Entra la nova comanda", "Comanda"))
  If cadbl(vnumc) > 0 Then
     vnumc_anterior = cadbl(comanda)
     comanda = vnumc
     comanda.tag = ""
     If Command4.Enabled Then Command4.SetFocus
     Command4_Click
     If comanda.tag = "" Then comanda.tag = atrim(vnumc_anterior): comanda = atrim(vnumc_anterior)
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
capcalera.capcalera.DatabaseName = soldadores.DatabaseName
capcalera.capcalera.RecordSource = "select * from soldadorestot where comanda=" + atrim(cadbl(comanda))
capcalera.capcalera.Refresh
If capcalera.capcalera.Recordset.EOF Then
   capcalera.capcalera.Recordset.AddNew
   capcalera.capcalera.Recordset!comanda = cadbl(comanda)
   capcalera.capcalera.Recordset.Update
End If
capcalera.capcalera.Refresh
capcalera.capcalera.Recordset.Edit
capcalera.Show 1
If Form1.soldadores.Recordset.EOF And Form1.soldadores.Recordset.BOF Then Command2.SetFocus: Command2_Click
reixa.col = 5
reixa.SetFocus
End Sub

Private Sub Command10_Click()
Dim i As Double
 If nummaq = 0 Then MsgBox "Primer has d'escullir una màquina": Exit Sub
If numbobinesnocorrelatiu Then MsgBox "Els numeros de bobines no son correlatius. Reviseu per continuar la bobina " + r: Exit Sub

comandaacavada.Value = 0
soldadores.Recordset.Move 0
If Not soldadores.Recordset.EOF Then
    If Not IsDate(soldadores.Recordset!datafi) Or Not IsDate(soldadores.Recordset!horafi) Then
        soldadores.Recordset.Edit
        soldadores.Recordset!datafi = Date
        soldadores.Recordset!horafi = Time
        soldadores.Recordset.Update
    End If
End If

client.ToolTipText = client.caption
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
    frameempalmes.visible = Not frameempalmes.visible
    framepantones.visible = False
    framebobentrada.visible = False
    If Not frameempalmes.visible Then reixabobines.SetFocus
End If
End Sub

Private Sub Command13_Click()
 If bobines.Recordset.EOF Then
     MsgBox "No hi ha bobina creada"
  Else
    framebobentrada.visible = Not framebobentrada.visible
    framepantones.visible = False
    frameempalmes.visible = False
    If Not framebobentrada.visible Then reixabobines.SetFocus
 End If
End Sub

Private Sub Command14_Click()
  Dim rstbobines As Recordset
     'reixa_BeforeDelete 0
'     If MsgBox("Segur que vols borrar aquesta linia i tot el seu contingut?", vbYesNo, "Atenció") = vbNo Then Cancel = 1
If nummaq = 0 Then Exit Sub
   If IsDate(soldadores.Recordset!datafi) And IsDate(soldadores.Recordset!horafi) Then
     If MsgBox("Aquesta linia ja te la hora de fi possada, SEGUR QUE VOLS ELIMINAR-LA?", vbCritical + vbYesNo + vbDefaultButton2, "ATENCIÓ") = vbNo Then Exit Sub
   End If
    r = 0
    If atrim(soldadores.Recordset!tipus) = "C" Then
      If MsgBox("Segur que vols eliminar aquesta linia de CANVI?", vbCritical + vbYesNo, "Atenció ELIMINACIÓ") = vbNo Then Exit Sub
    End If
    If atrim(soldadores.Recordset!tipus) = "F" Then
      Set rstbobines = dbtmpb.OpenRecordset("select * from bobinessol where controlid=" + atrim(cadbl(soldadores.Recordset!id)) + " order by numerodesac")
      If Not rstbobines.EOF Then
       rstbobines.MoveLast
       If MsgBox("Eliminar aquesta linia pot suposar eliminar informació de " + IIf(rstbobines.RecordCount > 0, atrim(rstbobines.RecordCount), "") + " bobines.", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then MsgBox "No s'ha eliminar cap informació.": Exit Sub
       If rstbobines.RecordCount > 0 Then
        If InputBox("PER PODER ELIMINAR AQUESTES CAIXESS HAS DE TECLEJAR " + Chr(13) + Chr(10) + "ELIMINAR " + atrim(rstbobines.RecordCount) + " BOBINES", "SEGURETAT PER ELIMINAR BOBINES") <> "ELIMINAR " + atrim(rstbobines.RecordCount) + " CAIXES" Then MsgBox "El texte no coincideix no s'eliminarà res.": Exit Sub
       End If
       rstbobines.MoveFirst
       On Error Resume Next
       While Not rstbobines.EOF
        If Not rstbobines.EOF Then
         'dbtmpb.Execute "delete * from lamempalmes where id=" + atrim(cadbl(rstbobines!id))
         dbtmpb.Execute "delete * from bobinesentsol where id=" + atrim(cadbl(rstbobines!id))
         rstbobines.Delete
        End If
        rstbobines.MoveNext
       Wend
       On Error GoTo 0
      End If

    End If
    'dbtmpb.Execute "delete * from bobinesreb where controlid=" + r
'    dbtmpb.Execute "delete * from impressores where id=" + atrim(r2)
    soldadores.Recordset.Delete
    soldadores.Recordset.MoveLast
    soldadores.Refresh
    If Not soldadores.Recordset.EOF Then soldadores.Recordset.MoveLast
    bobines.Refresh
    Command4_Click
    
End Sub

Sub quedenbobinesentrada()
   Dim rsttmp2 As Recordset
   Dim taulabob
   If bobines.Recordset.EOF Then Exit Sub
   taulabob = IIf(proces.tag = "I", "bobinesimp", "bobineslam")
   sa = "noutilitzades": r = "carregartaulatmp": bobentrada_DblClick
   Set rsttmp2 = dbtmpb.OpenRecordset("select * from bobentradatmpsol" + atrim(nummaq))
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
   taulabob = IIf(proces.tag = "I", "bobinesimp", "bobineslam")
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
   carregar_bobinesdentrada "carregarbobinesnoutilitzades", , , , cadbl(comanda), , IIf(proces.tag = "invertit", True, False)
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
                  carregar_bobinesdentrada "marcarutilitzadademanar", , palet, bobina, cadbl(comanda), True, ncomanda2, IIf(proces.tag = "invertit", True, False)
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
 Set rstbobent = dbtmpb.OpenRecordset("SELECT bobinesentreb.palet, bobinesentreb.bobina, soldadores.comanda FROM (bobinesentreb INNER JOIN bobinesreb ON bobinesentreb.id = bobinesreb.Id) INNER JOIN soldadores ON bobinesreb.controlid = soldadores.Id WHERE (((soldadores.comanda)=" + atrim(numc) + "));")
 
 'Set rstbobent = dbtmpb.OpenRecordset("SELECT distinct soldadores.comanda, bobinesentreb.palet, bobinesentreb.bobina FROM (bobinesentreb INNER JOIN bobinesreb ON bobinesentreb.id = bobinesreb.Id) INNER JOIN soldadores ON bobinesreb.controlid = soldadores.Id WHERE (((soldadores.comanda)=151013));")



 
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
   If cadbl(tmetres) < (cadbl(tmetres.tag) - ((cadbl(tmetres.tag) / 10) * 2)) Then
          If UCase(InputBox("Aquesta comanda es de " + tmetres.tag + " metres i tu has fet " + tmetres + " metres" + Chr(10) + "PASSARÉ LA COMANDA A NO ACABADA. ESCRIU ACABADA SI ESTÀ REALMENT ACABADA", "ATENCIÓ")) = "ACABADA" Then
              metresfetsinferiorsacomanda = False
               Else: metresfetsinferiorsacomanda = True
          End If
   End If
End Function

Private Sub Command15_Click()
 Dim com As Double
 If nummaq = 0 Then MsgBox "No hi ha numero de màquina assignat.": Exit Sub
 If soldadores.Recordset.EOF Then Exit Sub
 If numbobinesnocorrelatiu Then
     If MsgBox("Els numeros de sacs/caixes no son correlatius, reviseu-ho per continuar. " + r + Chr(10) + "VOLS CONTINUAR IGUALMENT? O VOLS PARAR L'IMPRESIÓ I MODIFICAR-HO?", vbCritical + vbYesNo, "Atenció") <> vbYes Then Exit Sub
 End If
 quedenbobinesentrada
 client.ToolTipText = client.caption
 If comprovarsifaltencamps Then Exit Sub
comandaacavada.Value = 1
soldadores.Recordset.MoveLast
If Not soldadores.Recordset.EOF Then
    If Not IsDate(soldadores.Recordset!datafi) Or Not IsDate(soldadores.Recordset!horafi) Then
        soldadores.Recordset.Edit
        soldadores.Recordset!datafi = Date
        soldadores.Recordset!horafi = Time
        soldadores.Recordset.Update
        wait 1
    End If
End If
mirar_bobinesdentrada_noacavades
'If metresfetsinferiorsacomanda(cadbl(comanda)) Then Command10_Click: Exit Sub
passar_comanda_a_acavada
crear_actualitzar_bobinesdentrada cadbl(comanda)
calcular_totals
guarda_totals
verificacio_netejaidespeje
wait 1
Command4_Click
ratoli "espera"
wait 2
Command8_Click
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


soldadores.Recordset.MoveLast
  'posso la data als totals de seccio
  If IsDate(soldadores.Recordset!datafi) Then
   dbtmpb.Execute "update soldadorestot set datasoldadora=#" + Format(soldadores.Recordset!datafi, "yy/mm/dd") + "# where comanda=" + atrim(cadbl(comanda))
   dbtmpb.Execute "update soldadorestot set operari=" + atrim(cadbl(soldadores.Recordset!operari1)) + " where comanda=" + atrim(cadbl(comanda))
   dbtmpb.Execute "update soldadorestot set soldadora=" + atrim(cadbl(soldadores.Recordset!numeromaquina)) + " where comanda=" + atrim(cadbl(comanda))
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
   If estat = "S" Then
     'seccio = Mid(ruta, InStr(1, ruta, "R") + 1, 1)
     'If atrim(seccio) = "" Then
     seccio = "V"
     dbtmp.Execute "update comandes set proximaseccio='" + seccio + "' where comanda=" + atrim(comanda)
     dbtmp.Execute "update comandes set seccioactual='" + seccio + "' where comanda=" + atrim(comanda)
   End If
End Sub
Function comprovarsifaltencamps() As Boolean
  Dim faltenpatones As Boolean
  Dim faltenmtrs As Boolean
  Dim rstc As Recordset
  
  soldadores.Recordset.FindLast "tipus='F'"
  'If Not soldadores.Recordset.NoMatch Then
  '    If cadbl(soldadores.Recordset!metresminut) = 0 Then MsgBox "Falten els Metres per minut": comprovarsifaltencamps = True
  'End If
 ' Set rstc = dbtmp.OpenRecordset("select rebkilos from comandes where comanda=" + atrim(comanda))
 ' If Not rstc.EOF Then
 '   If tkilos < (cadbl(rstc!rebkilos) - (cadbl(rstc!rebkilos) * 0.3)) Then
 '        If MsgBox("Els kilos que has fabricat son menys d'un 70% de la comanda." + Chr(10) + "Es segur que vols acabar comanda?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then
 '          comprovarsifaltencamps = True
 '        End If
 '   End If
 ' End If
  
End Function

Private Sub Command16_Click()
 framepalets.tag = "0"
 possar_botons_palets
End Sub

Private Sub Command17_Click()
 framepalets.tag = "10"
 possar_botons_palets
End Sub

Private Sub Command18_Click()
 framepalets.tag = "20"
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
 If Not soldadores.Recordset.EOF Then
  soldadores.Recordset.MoveLast
  If soldadores.Recordset!tipus = "C" Then
      numop = escullir_operari
      nomoperari = UCase(r)
  End If
 End If
 crearseccio "C"
 reixa.SetFocus
 ensenya_les_bobines
 colocarelsbotonsdelspalets
 'mostra 0 sempre al final d'aquest procediment
'    imprimir_controlbobina0 cadbl(comanda)
'    imprimir_controlqualitatVQ cadbl(comanda)
'    MsgBox "Pensa a treure la mostra 0 enganxant l'etiqueta amb la direcció de sortida correcte.", vbInformation, "Mostra 0"
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
  If Not soldadores.Recordset.EOF Then
      finalitza_seccio
      com = cadbl(soldadores.Recordset!comanda)
  End If
  r = ""
  If com = 0 Then Exit Sub
  soldadores.Recordset.AddNew
  soldadores.Recordset!comanda = com
  soldadores.Recordset!numeromaquina = nummaq
  soldadores.Recordset!operari1 = numop
  soldadores.Recordset!tipus = tipus
  soldadores.Recordset!datainici = Date
  soldadores.Recordset!horainici = Time
  
  'soldadores.Recordset!texteimpresio = rsttmpcs!texteimpressio
  r = soldadores.Recordset!id
  soldadores.Recordset.Update
  soldadores.Recordset.MoveLast
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
  demanar_paletibobina palet, bobina, desb
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
       carregar_bobinesdentrada "mirarsiutilitzada", , palet, bobina, ncomanda, utili, ncomanda2, IIf(proces.tag = "invertit", True, False)
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
  If InStr(1, Form1.caption, "Imprimint la bobina") > 0 Then MsgBox "S'està imprimint la bobina espera a que acavi sisplau.", vbCritical, "Error": Exit Sub
  vnumbobsxrpalet = contarbobinesdelpalet(cadbl(comanda), cadbl(numpalet))
  If vnumbobsxrpalet > 0 Then
    If cadbl(InputBox("Quantes Caixes/Sacs hi ha en aquest palet?", "Verificació de Caixes")) = vnumbobsxrpalet Then
        imprimirfullpalet cadbl(comanda), cadbl(numpalet)
         Else: MsgBox "No coincideix el numero de Caixes/Sacs amb el que has entrat," + Chr(10) + " hauria de ser " + atrim(vnumbobsxrpalet) + " Caixes/Sacs, revisa que estigui tot bé.", vbExclamation, "Atenció"
    End If
  End If
End Sub
Function contarbobinesdelpalet(vnumc As Double, nump As Double) As Double
    Dim rstp As Recordset
    contarbobinesdelpalet = 0
    Set rstp = dbtmpb.OpenRecordset("SELECT soldadores.comanda , bobinessol.palet, Count(bobinessol.numerodesac) AS bobines FROM soldadores INNER JOIN bobinessol ON soldadores.Id = bobinessol.controlid GROUP BY soldadores.comanda, bobinessol.palet HAVING (((soldadores.comanda)=" + atrim(vnumc) + ") AND ((bobinessol.palet)=" + atrim(nump) + "));")
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
Sub imprimirfullpalet(numc As Double, nump As Double)
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
    Set rstp = dbtmpb.OpenRecordset("SELECT soldadores.comanda , bobinessol.palet, Count(bobinessol.numerodesac) AS bobines FROM soldadores INNER JOIN bobinessol ON soldadores.Id = bobinessol.controlid GROUP BY soldadores.comanda, bobinessol.palet HAVING (((soldadores.comanda)=" + atrim(numc) + ") AND ((bobinesSOL.palet)=" + atrim(nump) + "));")
    If rstp.EOF Then Exit Sub
'    kilos = cadbl(rstp!skilos)
'    metres = cadbl(rstp!smetres)
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
' llistat.Destination = crptToWindow
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
 llistatpalet.Formulas(10) = "seccio='S'"
 'llistatpalet.Formulas(10) = "refinplacsa='" + IIf(vimprimirrefinplacsa <> "", "Ref.Inplacsa: ", "") + vimprimirrefinplacsa + "'"
' llistat.PrinterName = llegir_ini("Impressores", "nomfulla", "baixesimpressora.ini")
' llistat.PrinterPort = llegir_ini("Impressores", "portfulla", "baixesimpressora.ini")
' llistat.PrinterDriver = llegir_ini("Impressores", "driverfulla", "baixesimpressora.ini")
  DoEvents
 If existeix("c:\ordprog.ini") Then llistatpalet.Destination = crptToWindow
 llistatpalet.Action = 1
Set rstp = Nothing
    
End Sub

Private Sub Command26_Click()
  formbossesperembossar.Show 1
End Sub

Private Sub Command27_Click()
client.ToolTipText = client.caption
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
 obrir_DOC_ArxiuSOL
End Sub
Sub obrir_DOC_ArxiuSOL()
 Dim rst As Recordset
  Dim vnomfitxer As String
  
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(comanda))
  If Not rst.EOF Then
    vnomfitxer = ruta_relativa_docs + "\" + atrim(rst!arxiusol)
    If Not existeix(vnomfitxer) And InStr(1, UCase(vnomfitxer), ".DOC") > 0 Then vnomfitxer = vnomfitxer + "x"
    If existeix(vnomfitxer) And InStr(1, UCase(vnomfitxer), ".DOC") > 0 Then obrir_document vnomfitxer
  End If
  Set rst = Nothing
End Sub

Private Sub Command3_Click()
  Dim i As Byte
 Dim mtrsprova As String
 Dim mtrsparcials As Double
 Dim opantic As Byte
 Dim idbobina As Long
 If nummaq = 0 Then MsgBox "Escull primer numero de màquina": Exit Sub
 If comprovarsidescansorelleu Then Exit Sub
 
 If Not soldadores.Recordset.EOF Then
    soldadores.Recordset.MoveLast
    If soldadores.Recordset!tipus = "A" Then
        mtrsprova = InputBox("Entra els Metres de prova.", "Atenció")
        soldadores.Recordset.FindLast "tipus='A'"
        If Not soldadores.Recordset.NoMatch Then
         soldadores.Recordset.Edit
         soldadores.Recordset!mtrsprova = cadbl(mtrsprova)
         soldadores.Recordset.Update
        End If
    
    End If
    Else: Exit Sub
 End If
 'firmar_fulla
 If soldadores.Recordset!tipus = "F" Then
 
    opantic = numop
    numop = escullir_operari
    nomoperari = UCase(r)
 End If
 
 crearseccio "F"
 
 If cadbl(pespalet) = 0 Then demanar_el_pespalet
' If cadbl(bobinesxpalet) = 0 Then
'   bobinesxpalet = InputBox("Entra les bobines per palet.", "Atenció")
'   botopalets(0).SetFocus
   
'   If cadbl(client.tag) = 6603 Then MsgBox "Aquest client demana un codi de barres extra per etiqueta, surtirà una etiqueta amb un codi de barres i s'ha d'enganxar amb l'ETIQUETA EXTERIOR.", vbInformation, "Atenció"
'   reixa.SetFocus
' End If
 
 soldadores.Refresh
 soldadores.Recordset.MoveLast
 While bobines.Recordset.RecordCount = 0 And mtrsparcials < 100
   DoEvents
   bobines.Refresh
   mtrsparcials = mtrsparcials + 1
 Wend
 'If bobines.Recordset.RecordCount = 0 Then Command5_Click
 colocarelsbotonsdelspalets
'  tamany_visualitzadorpdf True
 avisarquelacomandasestaacabant cadbl(comanda), "R"
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
  Set rstp = dbtmpb.OpenRecordset("SELECT comandes.comanda,comandes.proximaseccio,comandes.producte, soldadorestot.acavada as acavadar, laminadorestot.acavada as acavadal, impressorestot.acavada as acavadai FROM ((comandes LEFT JOIN soldadorestot ON comandes.comanda = soldadorestot.comanda) LEFT JOIN laminadorestot ON comandes.comanda = laminadorestot.comanda) LEFT JOIN impressorestot ON comandes.comanda = impressorestot.comanda WHERE (((comandes.comanda)=" + atrim(numc) + "));")
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
'   carregar_pdf 0, 0
   If comandavalida And InStr(1, rst!ruta, "I") > 0 Then
     If Not clixesentratsafabrica(cadbl(numc)) Then
       comandavalida = False
       MsgBox "La comanda " + atrim(numc) + " no te els CLIXES ENTRATS a disseny. No es poden utilitzar.", vbCritical, "Atenció"
         Else
            carregar_DocX_Soldadora
            'carregar_pdf rst!numtreball, rst!numordremodificacio
     End If
   End If
End Function
Sub carregar_DocX_Soldadora()

End Sub
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
If nummaq = 0 Then Exit Sub
 If comprovarsidescansorelleu Then Exit Sub
 numpalet = 1
 If Not soldadores.Recordset.EOF Then
  soldadores.Recordset.MoveLast
  If soldadores.Recordset!tipus = "P" Then
      numop = escullir_operari
      nomoperari = UCase(r)
  End If
 End If
 crearseccio "P"
 reixa.SetFocus
 ensenya_les_bobines
 colocarelsbotonsdelspalets
End Sub

Private Sub Command31_Click()
If nummaq = 0 Then Exit Sub
 If comprovarsidescansorelleu Then Exit Sub
 numpalet = 1
 If Not soldadores.Recordset.EOF Then
  soldadores.Recordset.MoveLast
  If soldadores.Recordset!tipus = "A" Then
      numop = escullir_operari
      nomoperari = UCase(r)
  End If
 End If
 crearseccio "A"
 reixa.SetFocus
 ensenya_les_bobines
 colocarelsbotonsdelspalets
End Sub

Private Sub Command32_Click()
    carregar_capcalera
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
  If nummaq = 0 Then MsgBox "Escull una màquina primer.", vbCritical, "Atenció": Exit Sub
  vperforat = False
  If cadbl(bandes) > 0 Then
     contadorverificacio = cadbl(tmetres) / cadbl(bandes)
       Else: contadorverificacio = 1
  End If
  'comprovo si existeix la comanda
  netejarcampsdeltotalcomanda
  'AcroPDF1.OpenFile "dsf"
'  AcroPDF1.src = ""
  Set rsttmp = dbtmp.OpenRecordset("select cantitatsol,microperforat,rebmacroperforat,tubbase,refilatd,producte,client,direnvio,etrebvistiplau,rebmtrs,rebkilos,cantitatex,mesuracantex,amplereb,producte,linkcomanda1,linkcomanda2,lotmatdesb1,lotmatdesb2,rebobinadora,codibarras,espessor,comanda,refclient,comandaclient,texteimpressio,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(cadbl(comanda)))
  If rsttmp.EOF Or cadbl(comanda) = 0 Then
      MsgBox "No hi ha numero de comanda vàlida"
         Command1.Enabled = False:   Command2.Enabled = False:   Command3.Enabled = False: Exit Sub
  End If
  tubbase = IIf(Not IsNull(rsttmp!tubbase), rsttmp!tubbase, 0)
  'comprovo si hi ha seccio de rebobinadora
  If nohihasoldadora(rsttmp!producte) Then
      MsgBox "No hi ha seccio de soldadora en aquesta comanda"
         Command1.Enabled = False:   Command2.Enabled = False:   Command3.Enabled = False: Exit Sub
  End If
  
  If Not comandavalida(cadbl(comanda), True, vpararcomanda) Then
    If vpararcomanda Then comanda = "0": Exit Sub
    If MsgBox("Aquesta comanda ESTÀ PARADA O HI HA ALGUN MOTIU PER PARAR-LA." + Chr(10) + "Vols continuar igualment?", vbCritical + vbYesNo + vbDefaultButton2, "ATENCIÓ") = vbNo Then Exit Sub
  End If
  ncomanda = cadbl(comanda)
  ncomanda2 = IIf(cadbl(rsttmp!linkcomanda2) > 0, cadbl(rsttmp!linkcomanda2), cadbl(rsttmp!linkcomanda1))
  nlinkcomanda2 = cadbl(rsttmp!linkcomanda2)
  tpescanutu.HelpContextID = 0
  numpalet = 0
  botopalets_Click 0
  bobines.tag = ""
  proces.tag = ""
  
  'fins aqui comprova rebobinadora
  'carrego el pes net de clients_envios
  'miro si la comanda te preu assignat
  comprovarsitepreuassignatosinoenviarunmail cadbl(comanda)
  'carrego els camps de l'etiqueta
  imprimir_bobina "sense imprimir"
  
  
  'poso els botons de palets a punt
  framepalets.tag = "0"
  tunitats.tag = ""
  tkilos.tag = ""
  'If cadbl(rsttmp!mesuracantex) = 1 Then
     tunitats.tag = cadbl(rsttmp!cantitatsol)
     'tkilos.Tag = cadbl(rsttmp!rebkilos)
  '  Else: tkilos.Tag = cadbl(rsttmp!cantitatex)
  'End If
  vlink3 = cadbl(rsttmp!linkcomanda2)
  
  amplereb = cadbl(rsttmp!amplereb)
  ettoleranciaample.caption = "Tolerancia Ample Reb: " + atrim((amplereb * 10) - 2) + " a " + atrim((amplereb * 10) + 2) + " mm"
  
  ensenya_totals
  calcular_totals True
  bobines.RecordSource = "select * from bobinessol where controlid=-1"
  bobines.Refresh
  
  Set rsttmp = dbtmp.OpenRecordset("select mtrslinbob,marcailinia,tubolam,codibarras,espessor,mesuraesp,comanda,refclient,comandaclient,texteimpressio from comandes where comanda=" + atrim(cadbl(comanda)))
  mesuraespcomanda = ""
  If Not rsttmp.EOF Then
     Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rsttmp!mesuraesp)))
     If Not rsttmp2.EOF Then mesuraespcomanda = rsttmp2!descripcio
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
  soldadores.RecordSource = "select * from soldadores where comanda=" + atrim(cadbl(comanda)) + " order by datainici,horainici"
  '* imppantones.RecordSource = "select * from soldadoresadhesius where comanda=" + atrim(cadbl(comanda))
  soldadores.Refresh
  'imppantones.Refresh
  
 '* If imppantones.Recordset.EOF Then
'*     crear_pantones
'* imppantones.RecordSource = "select * from soldadoresadhesius where comanda=" + atrim(cadbl(comanda))
'*  End If
  possar_botons_palets
  carregar_client_ntintersialtres
  possar_camps_generals
  canutustallats = ""
  reixa.ReBind
  calcular_totals True
  'If soldadores.Recordset.EOF And soldadores.Recordset.BOF And Command1.Enabled Then Command1_Click
  framebobines.Enabled = False: framepantones.visible = False
  'If soldadores.Recordset.EOF Then Command1_Click
'  If impresores.Recordset.EOF Then MsgBox "Baixa nova es començarà amb edició de Clixes.": Command4.Tag = "nou": crearseccio "C": Command4.Tag = ""
  If bobines.tag <> "" And bobines.tag <> "-1" Then
   Set rsttmp = dbtmpb.OpenRecordset("select max(palet) as maxpalet from bobinessol where  controlid in(" + bobines.tag + ")")
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
  If soldadores.Recordset.EOF Then If proces.tag = "invertit" Then passarbobinesentradanoutilitzades cadbl(comanda)
  comanda.tag = comanda.Text
'  tamany_visualitzadorpdf True
ratoli "normal"
  If UCase(Screen.ActiveControl.Name) = "COMMAND4" Then
    obrir_DOC_ArxiuSOL
    carregar_capcalera
  End If
End Sub
Sub carregar_capcalera()
 Load FormResumComanda
  FormResumComanda.carregar_dadescomanda cadbl(comanda)
  FormResumComanda.Show 1
End Sub
Function potfermicromacroperforat(vmicro As String, vmacro As String) As Boolean
   Dim rst As Recordset
   
   potfermicromacroperforat = True
   Set rst = dbtmp.OpenRecordset("select * from maquines where maquina='R' and codi=" + atrim(nummaq))
   If Not rst.EOF Then
      If vmicro <> "N" And vmicro <> "" Then If InStr(1, atrim(rst!rebmicromacro), "Micro" + atrim(vmicro)) = 0 Then potfermicromacroperforat = False
      If vmacro = "S" Then If InStr(1, atrim(rst!rebmicromacro), "Macro") = 0 Then potfermicromacroperforat = False
        Else: potfermicromacroperforat = False
   End If
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
  client.caption = ""
  Set rstnt = dbtmp.OpenRecordset("select client,proximaseccio,cilindres,numerotintes from comandes where comanda=" + atrim(cadbl(comanda)))
  If Not rstnt.EOF Then
       ntintes = cadbl(rstnt!numerotintes)
       ncilindre = cadbl(rstnt!cilindres)
       framepantones.tag = atrim(rstnt!proximaseccio)
       codicli = cadbl(rstnt!client)
       Set rstnt = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(codicli))
       If Not rstnt.EOF Then client.caption = rstnt!nom: client.tag = atrim(codicli)
  End If
End Sub
Function nohihasoldadora(producte As String) As Boolean
  Dim rstreb As Recordset
  nohihasoldadora = True
  Set rstreb = dbtmp.OpenRecordset("select ruta from productes where codi='" + producte + "'")
   If Not rstreb.EOF Then
        If InStr(1, rstreb!ruta, "S") > 0 Then nohihasoldadora = False
   End If
End Function
Sub gravar_pantones()
On Error GoTo fi
 If Not imppantones.Recordset.EOF Then
  escriure_ini "soldadores", "lot1", imppantones.Recordset!lot1, "comandes.ini"
  escriure_ini "soldadores", "lot2", imppantones.Recordset!lot2, "comandes.ini"
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
   imppantones.Recordset!lot1 = llegir_ini("soldadores", "lot1", "comandes.ini")
   imppantones.Recordset!lot2 = llegir_ini("soldadores", "lot2", "comandes.ini")
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
   r = bobines.tag
   If r = "" Then r = "-1"
   Set rst = dbtmpb.OpenRecordset("select max(palet) as elgran from bobinessol where controlid in (" + atrim(r) + ")")
   If Not rst.EOF Then
      numerodepaletmesgran = cadbl(rst!elgran)
   End If
   If numerodepaletmesgran = 0 Then numerodepaletmesgran = 1
End Function
Private Sub Command5_Click()
 Dim vunitatsxfunda As String
'  If Not clixes.Enabled Then Exit Sub
If Command5.tag = "" Then Command5.tag = Now
If Command5.tag <> "" Then If DateDiff("s", Command5.tag, Now) > 5 Then Command5.tag = "": Exit Sub


Dim elgran As Double
Dim numb As Double
If numpalet < 1 Then numpalet = 1
If numpalet <> numerodepaletmesgran Then
   If numpalet <> (numerodepaletmesgran + 1) Then
    If MsgBox("No estàs col.locat al palet mes gran" + Chr(10) + "VOLS CONTINUAR AFEGINT LA CAIXA/SAC A AQUEST PALET IGUALMENT?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   End If
End If
i = 0
If IsDate(soldadores.Recordset!datafi) And IsDate(soldadores.Recordset!horafi) Then MsgBox "La linia de Funcionament actual està finalitzada. Canvia a la linia de Funcionament.": Exit Sub
If numop <> soldadores.Recordset!operari1 Then MsgBox "No pots afegir caixes a una linia d'un altre operari": Exit Sub
If bobines.Recordset.EditMode > 0 Then bobines.Recordset.Update
If cadbl(bobinesxpalet) > 0 And bobines.Recordset.RecordCount >= cadbl(bobinesxpalet) Then
   If MsgBox("Ja has posat " + atrim(cadbl(bobinesxpalet)) + " caixes/sacs en aquest palet, vols posar una altra?", vbYesNo, "Atenció") = vbNo Then Exit Sub
End If
If bobinesent.Recordset.EOF And Not bobines.Recordset.EOF Then MsgBox "No hi han bobines d'entrada a la ultima bobina feta": Exit Sub
While barraestat.caption = "Calculant els totals..."
  DoEvents
Wend
  dblots.visible = False
  framepantones.visible = False
  frameempalmes.visible = False
  framebobentrada.visible = False

  bobines.UpdateRecord
 If soldadores.Recordset!tipus = "F" Then
     'If cadbl(reixabobines.Columns(7).Text) = 0 And Not bobines.Recordset.EOF Then reixabobines.col = 7: reixabobines.SetFocus: MsgBox "Falten els metres a la bobina": Exit Sub
     'If cadbl(reixabobines.Columns(5).Text) = 0 And Not bobines.Recordset.EOF Then reixabobines.col = 5: reixabobines.SetFocus: MsgBox "Falten els kilos a la bobina": Exit Sub
    'caluclar totals
     'demanarcomandadebossesicanutus
     sa = ""
    ' If bobinesent.Recordset.RecordCount > 1 Then
    '  If MsgBox("Vols copiar kilos i bobines d'entrada?", vbYesNo, "Atenció") = vbYes Then
    '    sa = "copia kilos"
    '   Else: sa = ""
    '  End If
    ' End If
      If numbobinesnocorrelatiu Then MsgBox "Els numeros de sacs no son correlatius. Reviseu per continuar la " + r: Exit Sub
       nova_bobina elgran
       'copiarbobentanterior IIf(sa = "copia kilos", True, False)
       copiarbobentanterior True
     ''  copiarbobinaentanterior elgran
       'possarnumerodepalet
       'possarnumbobent True
       bobines.Refresh
       If Not bobines.Recordset.EOF Then
           bobines.Recordset.MoveLast
           If bobines.Recordset!numerodesac = 1 And cadbl(unitatsxfunda) = 0 Then
               vunitatsxfunda = InputBox("Vols possar les unitats en fundes?" + vbNewLine + "Escriu quantes en vols a cada funda.", "Utilitzar fundes")
               If cadbl(vunitatsxfunda) > 0 Then
                   unitatsxfunda = vunitatsxfunda
                   guarda_totals
               End If
           End If
       End If
       
       sa = ""
       If numbobinesnocorrelatiu Then MsgBox "Els numeros de caixes/sacs no son correlatius. Reviseu per continuar la " + r: Exit Sub
     Else: MsgBox "Has d'escullir una linia de FUNCIONAMENT."
  End If
'  reixabobines.col = 7
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
  Command5.tag = ""
  
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
    Set rstbobent = dbtmpb.OpenRecordset("SELECT bobinesreb.id FROM bobinessol WHERE (((bobinesreb.controlid) in (" + bobines.tag + ")) AND ((bobinesreb.numerodebobina)=" + atrim(bobant + 1) + "));")
    If Not rstbobent.EOF Then
         idbobina = cadbl(rstbobent!id)
       Else: Exit Function
    End If
    Set rstbobent = dbtmpb.OpenRecordset("SELECT bobinessol.controlid,bobinesreb.id, bobinessol.numerodesac, bobinesentsol.palet, bobinesentreb.bobina,bobinesentsol.paletobobina  FROM bobinesentsol INNER JOIN bobinessol ON bobinesentsol.id = bobinessol.Id WHERE (((bobinessol.controlid) in (" + bobines.tag + ")) AND ((bobinessol.numerodesac)=" + atrim(bobant) + "));")
    numbent = ""
    While Not rstbobent.EOF
      palet = rstbobent!palet
      bobina = rstbobent!bobina
      carregar_bobinesdentrada "mirarsiutilitzada", , palet, bobina, ncomanda, utili, ncomanda2, IIf(proces.tag = "invertit", True, False)
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

Function numbobinesnocorrelatiu() As Boolean
  Dim rstcp As Recordset
  Dim i As Integer
  If soldadores.Recordset.EOF Then Exit Function
  Set rstcp = dbtmpb.OpenRecordset("select * from bobinessol where  controlid in(" + bobines.tag + ") order by numerodesac")
  If Not rstcp.EOF Then i = rstcp!numerodesac
  numbobinesnocorrelatiu = False
  While Not rstcp.EOF And Not numbobinesnocorrelatiu
    If i <> rstcp!numerodesac Then numbobinesnocorrelatiu = True: r = atrim(i) + " <> " + atrim(rstcp!numerodesac)
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
 Dim vnumc As Double
 Dim rsttmp2 As Recordset
 vnumc = cadbl(comanda.Text)
 
 'Set rsttmp2 = dbtmpb.OpenRecordset("SELECT bobinesentsol.id_entrada FROM (soldadores LEFT JOIN bobinessol ON soldadores.Id = bobinessol.controlid) LEFT JOIN bobinesentsol ON bobinessol.Id = bobinesentsol.id Where (((soldadores.comanda) = " + atrim(vnumc) + ")) ORDER BY bobinesentsol.id_entrada DESC;")
 Set rsttmp2 = dbtmpb.OpenRecordset("SELECT bobinessol.id FROM (soldadores LEFT JOIN bobinessol ON soldadores.Id = bobinessol.controlid) LEFT JOIN bobinesentsol ON bobinessol.Id = bobinesentsol.id Where (((soldadores.comanda) = " + atrim(vnumc) + ")) ORDER BY bobinesentsol.id_entrada DESC;")
 If rsttmp2.EOF Then Exit Sub
 'If cadbl(bobinesent.tag) = 0 Then Exit Sub
 Set rsttmp1 = dbtmpb.OpenRecordset("select * from bobinesentsol where id=" + atrim(rsttmp2!id) + " order by id_entrada desc") ' + " and paletobobina='B'")
 While Not rsttmp1.EOF
   bobinesent.Recordset.AddNew
   bobinesent.Recordset!id = bobines.Recordset!id
   'bobinesent.Recordset!desb = rsttmp1!desb
   bobinesent.Recordset!palet = rsttmp1!palet
   bobinesent.Recordset!bobina = rsttmp1!bobina
   bobinesent.Recordset!paletobobina = rsttmp1!paletobobina
   bobinesent.Recordset.Update
   bobinesent.Refresh
   If carrastrar2bobs.Value = 0 Then GoTo cont
   rsttmp1.MoveNext
 Wend
cont:
 bobinesent.Refresh
 Set rsttmp1 = Nothing
End Sub

Sub copiarbobentanterior_novalid(Optional nopreguntarfibob As Boolean)
 Dim rsttmp1 As Recordset
 Dim primer As Boolean
 Dim rsttmp2 As Recordset
 If cadbl(bobinesent.tag) = 0 Then Exit Sub
 Set rsttmp1 = dbtmpb.OpenRecordset("select * from bobinesentsol where id=" + atrim(cadbl(bobinesent.tag))) ' + " and paletobobina='B'")
 obrestocks
 primer = True
 While Not rsttmp1.EOF
  If (atrim(rsttmp1!paletobobina) <> "P" And atrim(rsttmp1!paletobobina) <> "B") Or sa = "copia kilos" Then
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
           r = IIf(proces.tag <> "L", "bobinesimp", "bobineslam")
           dbtmpb.Execute "update  " + r + " set utilitzadaabaixa=True where id=" + atrim(cadbl(rsttmp2!idbobina))
      End If
              
      Else
       bobinesent.Recordset!paletobobina = LCase(bobinesent.Recordset!paletobobina)
       If UCase$(bobinesent.Recordset!paletobobina) = "P" Then
          dbstocks.Execute "update  bobines set utilitzadaabaixa=False where idpalet=" + bobentrada.Columns(0) + " and idbobina=" + bobentrada.Columns(1)
        Else:
          r = IIf(proces.tag <> "L", "bobinesimp", "bobineslam")
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
  Dim vunitatsxsac As Double
  
  Dim metresant As Double
  Dim kilosant As Double
  Dim kilosantnet As Double
  Dim bobsent As String
  metresant = 0
  kilosantnet = 0
  kilosant = 0
  reixabobines.tag = "afegint"
  'If Not bobines.Recordset.EOF Then
  ' If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
  ' bobines.Recordset.Update
   'metresant = cadbl(bobines.Recordset!metres)
   'kilosant = cadbl(bobines.Recordset!kilos)
  'End If
  r = bobines.tag
   If r = "" Then r = "-1"
  Set rsttmp2 = dbtmpb.OpenRecordset("select id  from soldadores where comanda=" + atrim(soldadores.Recordset!comanda))
   Set rstmp = dbtmpb.OpenRecordset("select * from bobinessol where controlid in (" + atrim(r) + ") order by numerodesac")
   If Not rstmp.EOF Then rstmp.MoveLast: vunitatsxsac = atrim(rstmp!unitatsxsac) ': metresant = cadbl(rstmp!metres): kilosantnet = cadbl(rstmp!pesnet):: kilosant = cadbl(rstmp!kilos)
   If cadbl(vunitatsxsac) = 0 Then vunitatsxsac = cadbl(InputBox("Entra les unitats per Sac.", "Unitats per sac."))
   Set rsttmp = Nothing
  elgran = 0
  
  While Not rsttmp2.EOF
   r = bobines.tag
   If r = "" Then r = "-1"
   Set rstmp = dbtmpb.OpenRecordset("select max(numerodesac) as elgran from bobinessol where controlid in (" + atrim(r) + ")")
   If Not rstmp.EOF Then
      If cadbl(rstmp!elgran) > elgran Then elgran = cadbl(rstmp!elgran)
      If r = "-1" Then elgran = 0
      'If cadbl(rstmp!paletgran) > numpalet Then numpalet = cadbl(rstmp!paletgran)
   End If
   rsttmp2.MoveNext
  Wend
  Set rstmp = dbtmpb.OpenRecordset("select * from bobinessol where controlid=" + atrim(soldadores.Recordset!id) + " and numerodesac=" + atrim(elgran))
  'bobines.Recordset.AddNew
'  If sa = "copia kilos" Then kilos = kilosant
 ' If llegirpesbascula > 0 Then kilos = llegirpesbascula
  'If pescanutu > 0 Then pesnet = cadbl(kilos) - cadbl(pescanutu)
  afegir_bobina elgran + 1, atrim(soldadores.Recordset!id), vunitatsxsac, cadbl(numpalet), Date, atrim(bobsent), cadbl(numop)
  col = 0
  bobinesent.tag = atrim(rstmp!id)
 ' If Not rstmp.EOF Then
 '    bobines.Recordset!metres = rstmp!metres
 '    bobines.Recordset!kilos = rstmp!kilos
 '    col = 3 'escullo a la columne que es posa per defecte
 '  Else: col = 3
 ' End If
  'bobines.Recordset.Update
  bobines.Refresh
  'bobines.Refresh
  bobines.Recordset.MoveLast
  'reixabobines.Refresh
  DoEvents
  reixabobines.col = col
  If reixabobines.Enabled Then reixabobines.SetFocus
  Set rstmp = Nothing
  Set rstmp2 = Nothing
If reixabobines.Text = "0" Then reixabobines.SelLength = Len(reixabobines.Text)
reixabobines.tag = ""
End Sub
Sub afegir_bobina(numbobent As Integer, idreb As Double, vunitatxsac As Double, numpalet As Double, data As Date, bobsent As String, numop As Byte)
   Dim camps As String
   Dim valors As String
   Dim nump As Double
   Dim vlotzipper As String
   Dim vlotcinta As String
   vlotcinta = llegir_ini("Baixes", "LotCinta", "comandes.ini")
   vlotzipper = llegir_ini("Baixes", "LotZipper", "comandes.ini")
   nump = numpalet
   If nump < 1 Then numpalet = 1: nump = 1
   camps = "(numerodesac,controlid,palet,datafab,operari1,unitatsxsac,lotcintaadhesiva,lotzipper)"
   valors = "(" + passaradecimalpunt(numbobent) + "," + passaradecimalpunt(idreb) + "," + passaradecimalpunt(nump) + ",#" + Format(data, "yy/mm/dd") + "#," + atrim(numop) + "," + atrim(vunitatxsac) + ",'" + vlotcinta + "','" + vlotzipper + "')"
   dbtmpb.Execute ("insert into bobinessol " + camps + " values " + valors)
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
 
 If MsgBox("Segur que vols BORRAR LA CAIXA/SAC Nº: " + atrim(bobines.Recordset!numerodesac) + " ?", vbCritical + 4 + vbDefaultButton2, "Atenció") = vbYes Then
     If Not bobines.Recordset.EOF Then
       dbtmpb.Execute "delete * from bobinesentsol where id=" + atrim(cadbl(bobines.Recordset!id))
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
   Dim vLongSol As Double
   If idbobina = 0 Then idbobina = 112239 ' apanyu per poder imprimir l'etiqueta interiro canutu de la primera bobina
   taula_tmp = "tmp_sol_empalmes" + atrim(nummaq)
   Set rstcom = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(numcom)))
   Set rstbob = dbtmpb.OpenRecordset("select * from bobinessol where id=" + atrim(cadbl(idbobina)))
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
     ' rsttmp!midarebobinat = cadbl(rstbob!ample) * 10
      rsttmp!desarroll = IIf(InStr(1, rst2!ruta, "I") > 0, rstcom!dessarroll, 0)
      rsttmp!numbob = rstbob!numerodesac
      rsttmp!metresbob = rstbob!unitatsxsac
      'rsttmp!pesbobina = cadbl(rstbob!kilos)
'      If cadbl(rstbob!pesnet) > 0 Then rsttmp!pesbobina = rstbob!pesnet * -1
      rsttmp!pescanutu = pescanutu
      rsttmp!peces = cadbl(rstbob!unitatsxsac)
      'If InStr(1, rst2!ruta, "I") > 0 And (atrim(rstcom!continu) <> "S" And rstcom!dessarroll > 0) Then rsttmp!peces = Redondejar((cadbl(rsttmp!metresbob * 1000) / cadbl(rstcom!dessarroll)), 0)
      vLongSol = cadbl(rstcom!longitudsol) - cadbl(rstcom!fuellebasesol) - cadbl(rstcom!fuellebocasol)
      rsttmp!liniamides = justificar(atrim(cadbl(rstcom!amplesol)), 5, "D") + "/" + justificar(atrim(rstcom!ampleplegsol), 5, "E") + " X " + justificar(atrim(vLongSol), 5, "E") + " (" + justificar(atrim(rstcom!fuellebasesol), 4, "D") + "/" + justificar(atrim(rstcom!fuellebocasol), 4, "E") + ")"
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
      rsttmp!operari = cadbl(rstbob!operari1)
      idio = IIf(rsttmp!idiomaclient <> "ES", "EN", rsttmp!idiomaclient)
      rsttmp!descproducte = atrim(rst2.Fields("descpelclient_" + idio))
     ' siesinterirobobinasensepes
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
Function justificar(v As String, longitut As Integer, Optional DoE As String) As String
    v = Mid(v, 1, longitut)
    If DoE <> "D" Then
       v = v + Space(longitut - Len(v))
      Else: v = Space(longitut - Len(v)) + v
    End If
    justificar = v
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
      r = bobines.tag
      If r = "" Then r = "-1"
      Set rstmp = dbtmpb.OpenRecordset("select max(numerodesac) as elgran from bobinessol where controlid in (" + atrim(r) + ")")

      rsttmp!numbob = cadbl(rstmp!elgran) + cadbl(bandes.tag)
      'If reixabobines.Row = -1 Then
       rsttmp!metresbob = 0
       rsttmp!pesbobina = 0
       rsttmp!peces = 0
       rsttmp!dataproduccio = Date
       'rsttmp!midarebobinat = cadbl(amplebob.Text) * 10
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
  Set rst = dbtmpb.OpenRecordset("SELECT soldadores.comanda, bobinesreb.kilos, bobinesreb.metres FROM soldadores RIGHT JOIN bobinesreb ON soldadores.Id = bobinesreb.controlid Where comanda = " + atrim(cadbl(comanda)))
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
If cont = 3 Then cont = 0: Form1.caption = "Baixes Comandes (soldadores)"
If InStr(1, Form1.caption, "Imprimint la bobina") <> 0 Then cont = cont + 1: Exit Sub
Form1.caption = "Imprimint la bobina."
If cadbl(bandes) = 0 Then MsgBox "Atenció el numero de bandes està a zero." + Chr(10) + "Aixó pot afectar a l'impresió d'etiquetes i creació de bobines noves", vbInformation, "Atenció"
guardar_reg_bobines
'If Not bobines.Recordset.EOF Then
' If Not comprovar_sipesimetresescorrecte(bobines.Recordset!kilos, bobines.Recordset!metres) Then
'   If MsgBox("Els metres i kilos entrats no sembla correspondres correctament." + Chr(10) + "Vols cancelar l'impressió i rectificar-ho?", vbCritical + vbYesNo, "Error") = vbYes Then
'      cont = 3
'      Form1.Caption = "Baixes Comandes (soldadores)"
'      Exit Sub
'   End If
' End If
'End If
If cadbl(etmetresbob.tag) > 0 And cadbl(bobines.Recordset!numerodesac) = 1 Then
  If cadbl(bobines.Recordset!metres) > (cadbl(etmetresbob.tag) * 1.1) Or cadbl(bobines.Recordset!metres) < (cadbl(etmetresbob.tag) / 1.1) Then
       If MsgBox("Els metres que has fet d'aquesta bobina son diferents que els que demana el client" + Chr(10) + "ASSEGURA'T QUE SIGUI CORRECTE." + Chr(10) + "Vols cancelar l'impresió?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then GoTo fi
  End If
End If
Form1.caption = "Imprimint etiqueta Sac.."
'Set rstco = dbtmp.OpenRecordset("select etrebvistiplau from comandes where comanda=" + atrim(cadbl(comanda)))
'If Not rstco.EOF Then
'   If Not cabool(rstco!etrebvistiplau) Then avis_et_noverificada: Set rstco = Nothing: Exit Sub
    
'End If
Command7.BackColor = &HFFFFFF
Form1.caption = "Imprimint etiqueta Sac..."
borraretiquetestemporals
'MsgBox "Encara no es pot imprimir la bobina"
'Exit Sub
Form1.caption = "Imprimint etiqueta Sac...."
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
   Form1.caption = "Imprimint la bobina....."
 'comprova si ha de fer la etiqueta de mostra
   If mostracli.visible Then
       If mostracli.Value = 0 Then
         If MsgBox("Encara no has imprès la etiqueta per fer la mostra pel client." + Chr(13) + Chr(10) + "Vols fer-ho ara?", vbInformation + vbYesNo, "Atenció") = vbYes Then
           imprimir_bobina "Muestra Cli", True: mostracli.Value = 1: guarda_totals
           If etiquetadeclientdeclient(cadbl(comanda)) Then demanarescriureokperoficina
         End If
       End If
   End If
End If
Form1.caption = "Imprimint etiqueta Sac......"
imprimir_bobina "Soldadores"
Form1.caption = "Imprimint etiqueta Sac......."
fi:
  Form1.caption = "Baixes Comandes (soldadores)"

End Sub
Sub comprovarsitocaverificacio()
  Dim metres As Double
  If cadbl(bandes) < 1 Then Exit Sub
  metres = cadbl(cadbl(tmetres) / cadbl(bandes)) + (cadbl(rsttmp!metresbob) * cadbl(bandes))
  'metres = metres - (Int(metres / 7000) * 7000)
  If (metres - contadorverificacio) > 7000 Then contadorverificacio = metres * -1
End Sub
Sub imprimir_bobina(aon As String, Optional mostra As Boolean)
 Dim idbob As Double
 Dim rstbob As Recordset
 Static ultimabobinaimpresa As Double
 
 Set rstbob = dbtmpb.OpenRecordset("select * from bobinessol")
 taula_tmp = "tmp_sol_empalmes" + atrim(nummaq)
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
    If cadbl(client.tag) = 6603 Then  'si es videcart faig el codidebarres extra
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
   'If existeix("c:\temp\etiquetareb.prn") Then Kill "c:\temp\etiquetareb.prn"
   Open llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetarebverificacio.prn" For Input As #1
   linia.Text = Input(LOF(1), #1)
   Close #1
   With rsttmp
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
   substituir "#linia4.1#", sitoca(retallar(!material, 40), "material")
   substituir "#linia5.1#", sitoca(retallar(!texteimpresio, 50), "texteimpresio")
   'substituir "#linia6.1#", sitoca(idioma("Ancho:") + atrim(!midarebobinat) + " m/m ", "midarebobinat") + sitoca(idioma("Desar:") + atrim(!desarroll) + " m/m", "desarroll")
   substituir "#linia6.1#", idioma("Ancho/Pleg. X Long. (F.Ba/F.Bo)")
   substituir "#linia7.1#", !liniamides
   substituir "A0,235,0,3,1,1,N,", "A0,235,0,1,2,2,N,"
   substituir "A0,204,0,1,2,2,N,", "A0,204,0,1,2,2,R,"
   substituir "#linia9.1#", idioma("Unidades:") + atrim(!peces)
   'substituir "#linia7.1#", sitoca(retallar(IIf(!obsetiqueta <> "", idioma("Obs.Et:") + !obsetiqueta, ""), 50), "obsetiqueta")
   substituir "#linia15.1#", sitoca(Format(!dataproduccio, "dd/mm/yy"), "dataproduccio") + " " + sitoca(idioma("Op:") + !operari, "operari") + " " + sitoca(idioma("Lote: ") + Format(!numcomanda, "#,##0"), "numcomanda") + " " + idioma(!situacioet)
   substituir "#linia8.1#", sitoca(idioma("NºCaixa/Sac:") + atrim(!numbob), "numbob")
   If Not !etmostra Then
      'substituir "#linia14.1#", sitoca(IIf(!peces > 0, idioma("Unidades:") + atrim(Format(!peces, "#,##0")), ""), "peces")
      If !pesbobina >= 0 Then
         substituir "#linia9.1#", sitoca(idioma("Peso:") + atrim(Format(!pesbobina, "#,##0.0")) + " Kg", "pesbobina")
         substituir "#linia16.1#", ""
        Else:
          substituir "#linia9.1#", sitoca(idioma("Neto:") + atrim(Format(!pesbobina * -1, "#,##0.0")) + " Kg", "pesbobina")
          substituir "#linia16.1#", sitoca(idioma("Mandril:") + atrim(Format(!pescanutu, "#,##0.0")) + " Kg", "pescanutu")
      End If
      'substituir "#linia10.1#", sitoca(idioma("Long:") + atrim(Format(!metresbob, "#,##0")) + " Mts", "metresbob")
      'substituir "#linia8.1#", sitoca(idioma("NºCaixa/Sac:") + atrim(!numbob), "numbob")
      
     Else
       substituir "#linia8.1#", idioma("Unidades:") + atrim(1)
       linia = linia + "A10,300,0,5,1,1,N," + Chr$(34) + idioma("ETIQUETA") + Chr$(34) & vbCrLf
       linia = linia + "A10,355,0,5,1,1,N," + Chr$(34) + idioma("MUESTRA") + Chr$(34) & vbCrLf
   End If
   substituir "#linia17.1#", retallar("RefInp:" + atrim(buscar_refinp(!numcomanda)), 39)
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
   substituir "#linia2.2#", ""
   substituir "#linia10.1#", ""
   substituir "#linia14.1#", ""
   
End Sub
Function buscar_refinp(vnumc As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("Select refinplacsa from comandes_extres where comanda=" + atrim(vnumc))
   If Not rst.EOF Then buscar_refinp = atrim(rst!refinplacsa)
   Set rst = Nothing
End Function
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
 rs!client = client.caption
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
  taula_tmp = "tmp_sol_empalmes" + atrim(nummaq)
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
        camps(i, 1) = "liniamides": camps(i, 2) = "string": i = i + 1
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
  Dim rstc As Recordset
  Set rstc = dbtmpb.OpenRecordset("select * from comandes where comanda=" + comanda.Text)
 
 Set rst = dbtmpb.OpenRecordset("select * from soldadorestot where comanda=" + comanda.Text)
 If Not rst.EOF Then
 
   rsttemp!amplebob = atrim(rstc!amplesol)
   rsttemp!espesor = atrim(rstc!espessorsol)
   rsttemp!ampleref = atrim(rstc!longitudsol)
'   rsttemp!bandesbones = atrim(rst!simulteneitat)
'   rsttemp!ampleref = atrim(rst!ampleref)
'   rsttemp!bandesmerma = atrim(rst!bandesmerma)
'   rsttemp!amplemerma = atrim(rst!amplemerma)
   rsttemp!lotbosses = atrim(rst!comandabosses1) + IIf(atrim(rst!comandabosses2) <> "", " - " + atrim(rst!comandabosses2), "")
   rsttemp!lotcanutus = atrim(rst!comandacaixes1) + IIf(atrim(rst!comandacaixes1) <> "", " - " + atrim(comandacaixes2), "")
 End If
 Set rst = Nothing
 Set rstc = Nothing
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
   Dim v As String
   
   
   If nomllistat = "" Then nomllistat = "baixessoldadora.rpt"
   Form1.caption = "Imprimint..."
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
  Set rst = dbtmpb.OpenRecordset("select id,operari1,datainici,horainici,datafi,tipus,horafi,observacio from soldadores where comanda=" + comanda.Text + " and tipus<>'F' order by datainici,horainici")
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
    .Fields("prepmaquina_observacions" + Trim(i)) = atrim(rst!tipus) + " " + atrim(rst!observacio)
    i = i + 1
    rst.MoveNext
    If i > 4 Then rst.MoveLast: rst.MoveNext
  Wend

  
  'temps funcionament
  Set rst = dbtmpb.OpenRecordset("select * from soldadores where comanda=" + comanda.Text + " and tipus='F' order by datainici,horainici")
  If Not rst.EOF Then rst.MoveLast
  If Not rst.EOF Then
        For i = 1 To 8
          rst.MovePrevious
          If rst.BOF Then rst.MoveNext: i = 10
        Next i
  End If
  i = 1
  While Not rst.EOF
    .Fields("tempsreb_datai" + Trim(i)) = Format(atrim(rst!datainici), "dd/mm/yy")
    .Fields("tempsreb_op" + Trim(i)) = cadbl(rst!operari1)
    .Fields("tempsreb_horai" + Trim(i)) = Format(atrim(rst!horainici), "hh:nn")
    .Fields("tempsreb_horaf" + Trim(i)) = Format(atrim(rst!horafi), "hh:nn")
    .Fields("tempsreb_observacio" + Trim(i)) = atrim(rst!observacio)
    If cadbl(rst!totalhores) > 0 Then .Fields("tempsreb_mtrsmin" + Trim(i)) = Redondejar(cadbl(rst!totalunitats) / ((60 * cadbl(rst!totalhores))), 0) 'cadbl(rst!metresminut)
    .Fields("tempsreb_metres" + Trim(i)) = 0 ' cadbl(rst!totalmetres)
    i = i + 1
    If i > 8 Then rst.MoveLast
    rst.MoveNext
  Wend
  
  'acavar comandes

  'posso els camps de totals
    !pescanutu = pescanutu: !hparada = cadbl(hparada): !havaria = havaria: !hmaquina = cadbl(hmaquina): !hfunc = cadbl(hfunc): !tbob = cadbl(tbob): !tmtrs = cadbl(tmetres): !tkilos = cadbl(tkilos): !mtrsmin = cadbl(kiloshora)
  '!acavada = comandaacavada
  Set rstbob = Nothing
  Set rst = Nothing
  
  
    
  End With
  
  'passo les bobines a la taula del llistat
  Set rst = dbtmpb.OpenRecordset("select * from soldadores where comanda=" + comanda.Text + " and tipus='F'")
  If rst.EOF Then dbtemp.Execute "insert into " + "tmp_reb_baixa_bob" + " (operari,palet1,bobent1,paletsort,bobsort,kilos,metres) values (0,0,'0',0,0,0,0)"
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
        Set rsttmp2 = dbtmpb.OpenRecordset("select * from bobinessol where controlid=" + atrim(cadbl(rst!id)))
        
        With rsttmp2
        If Not rsttmp2.EOF Then
         rsttmp2.MoveLast
         rsttmp2.MoveFirst
          'Else: dbtemp.Execute "insert into " + "tmp_reb_baixa_bob (operari,operari2,palet1,bobent1,bobsort,kilos,metres) values (0,0,0,'0',0,0,0)"
        End If
        While Not rsttmp2.EOF
          'If rsttmp2.AbsolutePosition + 1 = rsttmp2.RecordCount Then
              If Not rsttmp2.EOF Then Set rsttmpbob = dbtmpb.OpenRecordset("select * from bobinesentsol where id=" + atrim(cadbl(rsttmp2!id)) + " order by paletobobina ASC")
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
              
              Set rstpesp = dbtmpb.OpenRecordset("select * from sol_pespalets where numpalet=" + npalet + " and comanda=" + atrim(rsttmp!comanda))
              If Not rstpesp.EOF Then pesp = atrim(cadbl(rstpesp!pespalet))
              pesp = cadcml(pesp)
              
              dbtemp.Execute "insert into " + "tmp_reb_baixa_bob (datapalet,paletsort,pespalet,operari,palet1,bobent1,bobsort,observacions,metres) values ('" + atrim(rst!datainici) + "'," + npalet + "," + atrim(pesp) + "," + atrim(cadbl(!operari1)) + "," + atrim(np) + ",'" + atrim(nb) + "'," + atrim(cadbl(!numerodesac)) + ",'" + treure_apostruf(atrim(!observacions)) + "'," + passaradecimalpunt(atrim(cadbl(!unitatsxsac))) + ")"
              '(idbob integer,operari byte,operari2 byte,palet1 double,bobent1 string,palet2 double,bobent2 string,paletsort integer,bobsort integer,kilos double,metres double,senyals byte,pespalet double,observacions string,kilosnets double,datapalet string)
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
  
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + nomllistat
' llistat.Destination = crptToWindow
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 2
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 llistat.Formulas(1) = "nommaquina='Soladora - " + atrim(nummaq) + "-" + nom_maquina(cadbl(nummaq)) + "'"
 llistat.Formulas(0) = "texteimpresio='" + treure_apostruf(texteimpresio) + "'"
 llistat.Formulas(2) = "pescanutu='" + atrim(pescanutu) + "'"
 llistat.Formulas(3) = "desarroll=" + atrim(vdesarroll)
 v = atrim(buscar_altres_lots(comanda.Text))
 llistat.Formulas(4) = "altresLots='" + IIf(v <> "", "Altres Lots: " + v, "") + "'"

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
  exportarllistatapdf llistat, llegir_ini("General", "rutallistats", "comandes.ini") + "baixessoldadora_PDF.rpt", cadbl(comanda.Text), vcarpetadesti
  escriure_ini "General", "exportantpdfs", "no", llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\organitzar.ini"

  Set rsttmp = Nothing
  Set rst = Nothing
  Set dbstocks = Nothing
 'panelimprimir.Visible = False
 Form1.caption = "Baixes Comandes (soldadores)"
 
End Sub
Function buscar_altres_lots(vnumc As Double) As String
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("select * from soldadores_accessorisutilitzats where comanda=" + atrim(vnumc) + " and lottraçabilitat<>''")
  While Not rst.EOF
    buscar_altres_lots = buscar_altres_lots + treure_apostruf(rst!nomaccessori) + " Lot:" + treure_apostruf(atrim(rst!lottraçabilitat)) + " "
    rst.MoveNext
  Wend
  Set rst = Nothing
End Function
Function nom_maquina(vnummaq As Long) As String
  Dim rstm As Recordset
  Set rstm = dbtmp.OpenRecordset("select descripcio from maquines where maquina='S' and codi=" + atrim(vnummaq))
  If rstm.EOF Then Exit Function
  nom_maquina = atrim(rstm!descripcio)
  Set rstm = Nothing
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
'Load veurereport
'   veurereport.CRViewer.ReportSource = oreport
'   veurereport.CRViewer.DisplayGroupTree = False
'   veurereport.CRViewer.ViewReport
'   veurereport.Show 1, Me
'Exit Sub
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat
  oreport.ExportOptions.DiskFileName = vcarpetadesti + "\" + atrim(vnumc) + "_Baixesoldadores.pdf"
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
  campstotal = " hmaquina double,  hfunc double, hparada double,havaria double,  tbob double,tmtrs double, tkilos double, mtrsmin double, pescanutu double "
  
  'ample double,plegat double,solapa double,espessor double,metres double,kilos double)"
  'escriure_ini "a", "b", campsextra + camps + camps3 + camps2 + campspantone + campspantone2 + campstotal, "prova.ini"
    dbtemp.Execute ("create table tmp_reb_baixa (" + campscapcalera + camps + camps2 + camps3 + camps4 + campstotal + ")")
    dbtemp.Execute ("create table tmp_reb_baixa_bob (idbob integer,operari byte,operari2 byte,palet1 double,bobent1 string,palet2 double,bobent2 string,paletsort integer,bobsort integer,kilos double,metres double,senyals byte,pespalet double,observacions string,kilosnets double,datapalet string)")
  
End Sub



Private Sub Command8_Click()
client.ToolTipText = client.caption
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
    dblots.visible = False
    DoEvents
  framepantones.visible = Not framepantones.visible
  frameempalmes.visible = False
  framebobentrada.visible = False
  If Not framepantones.visible Then If reixabobines.Enabled Then reixabobines.SetFocus
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
  If KeyCode = 27 Then dblots.tag = "": dblots.visible = False
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
        etpesbascula.tag = "manual"
          Else: etpesbascula.tag = ""
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

Sub actualitzarestatbobinesdesbobinadors()

End Sub
Private Sub Form_Activate()
 If cadbl(numop) = 0 And Form1.tag <> "carregant" Then nomoperari_Click
 
 
End Sub
Sub demanarcomandadebossesicanutus()
    Dim rst As Recordset
    Dim rstc As Recordset
    Set rstc = dbtmp.OpenRecordset("select tubbase from comandes where comanda=" + atrim(comanda))
    Set rst = dbtmpb.OpenRecordset("select comandabosses1,comandabosses2,comandacanutus1,comandacanutus2 from soldadorestot where comanda=" + atrim(cadbl(Form1.comanda)))
    If Not rst.EOF Then
        While Len(atrim(rst!comandabosses1)) < 3 Or Len(atrim(rst!comandacanutus1)) < 3
         Load formbossesperembossar
         formbossesperembossar.Show
         formbossesperembossar.escullirisortir cadbl(rstc!tubbase)
         Set rst = dbtmpb.OpenRecordset("select comandabosses1,comandabosses2,comandacanutus1,comandacanutus2 from soldadorestot where comanda=" + atrim(cadbl(Form1.comanda)))
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
 
'avisarquelacomandasestaacabant cadbl(comanda), "R"

' imprimir_controlbobina0 cadbl(comanda)
' imprimiretiquetaverificacio 1
'imprimir_controlbobina0 cadbl(comanda)

'  demanarcomandadebosses
'imprimir_bobina "Ext.Bobina"
   'imprimir_controlqualitatVQ cadbl(comanda)

  'dbtmpb.Execute "UPDATE bobinesreb INNER JOIN soldadores ON bobinesreb.controlid = soldadores.Id SET bobinesreb.metres = 495 WHERE (((soldadores.comanda)=148992) AND ((bobinesreb.numerodebobina)=39));"
  
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
grup = cadbl(framepalets.tag)
For i = 0 To 9
  botopalets(i).caption = atrim((i + 1) + grup)
Next i
If Command16.tag <> "E" Then botopalets_Click 0
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


Private Sub Form_Load()
  Dim camistocks As String
  Dim MyWorkspace As Workspace
  On Error Resume Next
 
  Form1.tag = "carregant"
  '
  'If EstaCorriendo("baixesrebobinadora.exe") Then MsgBox "El programa ja està funcionant.": End
  If App.PrevInstance Then MsgBox "El programa ja està funcionant.": End
  Form1.tag = ""
  fitxerini = "comandes.ini"
  Shell "c:\windows\regedit.exe /s \\serverprodu\dades\progcomandes\aplicacio\desactivarctrl.reg"
  Shell "c:\windows\regedit.exe /s \\serverprodu\dades\progcomandes\aplicacio\activarctrl.reg"
  Shell ("net time \\serverprodu /set /y")
  camicomandes = llegir_ini("General", "cami", "comandes.ini")
  cami = llegir_ini("General", "camibaixes", "comandes.ini")
  
   If LCase(App.EXEName) <> "baixessoldadores" And LCase(App.EXEName) <> "baixes soldadores" Then Form1.BackColor = &HFF80FF
   
  
  obrestocks True
  If cami = "{[}]" Then
    escriure_ini "General", "camibaixes", InputBox("Entra la ruta de baixes", "Atenció", "y:\comandes\baixes.mdb"), "comandes.ini"
  End If
  
  comanda = cadbl(llegir_ini("Baixes", "ultimacomanda", "comandes.ini"))
  r = cadbl(llegir_ini("Baixes", "nummaq", "comandes.ini"))
  nummaq = cadbl(r)

  If Not existeix("c:\ordprog.ini") And nummaq > 0 Then assignardecimalipunt
  If nummaq = 0 Then
    maquina.visible = True
   Else: maquina.visible = False
  End If
  lletraseccio = "S"
  
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
  
  On Error GoTo 0
  
  soldadores.DatabaseName = cami
  imppantones.DatabaseName = cami
  bobines.DatabaseName = cami
  
  empalmes.DatabaseName = cami
  bobinesent.DatabaseName = cami
  
  lots.DatabaseName = cami
  lots.Refresh
  Set dbtmpb = OpenDatabase(soldadores.DatabaseName)
  rellotge.Enabled = True
  rellotge.Interval = 900
  
 'If cadbl(nummaq) = 0 Then MsgBox "No hi ha el numero de màquina posat": Exit Sub
  
  
  soldadores.RecordSource = "select * from soldadores where comanda=-1"
  soldadores.Refresh
  bobinesent.RecordSource = "select * from bobinesentsol where id=99999999"
  bobinesent.Refresh
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
      If objecte.Name <> "MSComm2" And objecte.Name <> "reciclarmaterial1" And objecte.Name <> "AcroPDF1" And objecte.Name <> "MSComm1" And objecte.Name <> "llistatpalet" And objecte.Name <> "nomoperari" And objecte.Name <> "Line1" And objecte.Name <> "rellotge" And objecte.Name <> "llistat" Then
        objecte.Enabled = False
      End If
     Next objecte
     
     
  frameempalmes.ZOrder 0
    framepantones.ZOrder 0
    framebobentrada.ZOrder 0
    
    
  'netejo els numeros de lots si cal
  If llegir_ini("Baixes", "LotCinta", "comandes.ini") = "{[}]" Then escriure_ini "Baixes", "LotCinta", "", "comandes.ini"
  If llegir_ini("Baixes", "LotZipper", "comandes.ini") = "{[}]" Then escriure_ini "Baixes", "LotZipper", "", "comandes.ini"
    
    
  If existeix("c:\ordprog.ini") Then formEstacions.Show 1: End
    
    
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
 Form1.tag = "tancant"
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
  If soldadores.Recordset!tipus = "F" Then
   Set rstcp = soldadores.Recordset.Clone
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
    bk = bobines.Recordset!numerodesac
    On Error GoTo 0
  End If
  bobines.tag = r
  bobines.RecordSource = "select * from bobinessol where palet=" + atrim(numpalet) + " and controlid in(" + r + ") order by numerodesac"
  If numpalet = 1 Then bobines.RecordSource = "select * from bobinessol where (palet=" + atrim(numpalet) + " or palet=0) and controlid in (" + r + ") order by numerodesac"

  bobines.Refresh
  If bobines.Recordset.EOF And r <> "-1" Then
    
           Set rstcp = dbtmpb.OpenRecordset("select max(palet) as maxpalet from bobinessol where  controlid in(" + r + ")")
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
  Command16.tag = "E"
  Select Case i
      Case 0: Command16_Click
      Case 1: Command17_Click
      Case 2: Command18_Click
  End Select
  Command16.tag = ""
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
   maquina.caption = "Maq: " + atrim(nummaq)
   maquina.tag = nummaq
End Sub

Private Sub mostracli_Click()
guarda_totals
End Sub

Private Sub PDFX1_OnError(lErr As Long, sErr As String)

End Sub

Private Sub pespalet_Change()
  If Screen.ActiveControl.Name = "pespalet" Then
   If cadbl(pespalet) > 0 Then gravar_pespalet
  End If
End Sub

Private Sub pespalet_DblClick()
pespalet.tag = "pesar"
agafarpesbascula_Click
End Sub

Private Sub pespalet_LostFocus()
 
  If controlactiu = "agafarpesbascula" Then
     pespalet.tag = "pesar"
    Else: pespalet.tag = ""
  End If
End Sub

Private Sub proces_Change()
 Dim rsttmpp As Recordset
 
 Set rsttmpp = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(proces) + "'")
 If InStr(1, rsttmpp!ruta, "R") = 0 Then proces.tag = "": Exit Sub
 If Not rsttmpp.EOF Then proces.tag = Mid(rsttmpp!ruta, InStr(1, rsttmpp!ruta, "R") - 1, 1)
End Sub

Private Sub soldadores_Reposition()
If Not soldadores.Recordset.EOF Then
      If atrim(soldadores.Recordset!tipus) = "F" Then ensenya_les_bobines
   If barraestat.caption <> "Calculant els totals..." Then colocarelsbotonsdelspalets
   'framebobines.Enabled = False
 End If
 missatge_exesdemtrskg
End Sub

Private Sub nomoperari_Click()
 Dim numoptmp As Integer
 Dim nomoptmp As String
 If barraestat.caption = "Calculant els totals..." Then Exit Sub
  Load formseleccio2
  formseleccio2.Data1.DatabaseName = camicomandes
  formseleccio2.Data1.RecordSource = "select codi,descripcio from operaris where maquina='S' and actiu<>0 order by codi asc"
  formseleccio2.caption = "Selecció d'Operari"
  formseleccio2.refrescar
  formseleccio2.Show 1
  If seleccioret = 1 Then
   numoptmp = cadbl(formseleccio2.Data1.Recordset!codi)
   nomoptmp = atrim(formseleccio2.Data1.Recordset!descripcio)
   'If InStr(1, nomoperari.Caption, "MARTINEZ") Then
   '    Command12.Visible = True
   '   Else: Command12.Visible = False
   'End If
  End If
  Unload formseleccio2
  If numoptmp <> 0 Then
     nomoperari = nomoptmp
     numop = numoptmp
     For Each objecte In Me
      If objecte.Name <> "MSComm2" And objecte.Name <> "reciclarmaterial1" And objecte.Name <> "AcroPDF1" And objecte.Name <> "MSComm1" And objecte.Name <> "llistatpalet" And objecte.Name <> "llistat" And objecte.Name <> "Line1" And objecte.Name <> "comandaacavada" Then
        objecte.Enabled = True
      End If
     Next objecte
      Else: If cadbl(numop) = 0 Then MsgBox "Has d'escullir un operari per treballar": Exit Sub
  End If
  If nummaq = 0 Then maquina_Click
   Command4.SetFocus
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
    If soldadores.Recordset!tipus = "F" Then r = atrim(cadbl(soldadores.Recordset!id))
    dbtmpb.Execute "delete * from bobinessol where controlid=" + r
  End If
End Sub

Private Sub reixa_DblClick()
   If reixa.col = 12 Then
'   r = triar_observacio(soldadores.Recordset!tipus)
'   If Len(r) > 4 Then
'     r = Mid(r, 4, Len(r))
'     If r <> "" Then
'       If reixa.Text <> "" Then
'           reixa.Text = reixa.Text + " <> " + r
'           Else: reixa.Text = r
'       End If
'     End If
'   End If
   r = InputBox("Escriu la observació:", "Observacio", reixa.Text)
   reixa.Text = Mid(r, 1, 50)
  End If
  If reixa.Columns(reixa.col).caption = "Sim." Then
   r = InputBox("Escriu la simulteneitat:", "Simulteneitat", reixa.Text)
   reixa.Text = atrim(cadbl(r))
  End If
  If reixa.Columns(reixa.col).caption = "Taco" Then
   r = InputBox("Escriu el valor del Tacometre:", "Tacometre", reixa.Text)
   reixa.Text = atrim(cadbl(r))
  End If
  If reixa.Columns(reixa.col).caption = "PreSel." Then
   r = InputBox("Escriu la Preselecció:", "Preselecció", reixa.Text)
   reixa.Text = atrim(cadbl(r))
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
  formseleccio.Data1.RecordSource = "select * from constantsobservacio where mid(observacio,1,2)='S" + tipus + "'"
  formseleccio.caption = "Triar Observació"
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
  'AcroPDF1.visible = False
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
If soldadores.Recordset.EOF Then Exit Sub
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
'If soldadores.Recordset!tipus = "C" Then reixa.Columns(12).Locked = False: reixa.Columns(14).Locked = False  ': reixa.Columns(11).Locked = False:reixa.Columns(7).Locked = False
If soldadores.Recordset!tipus = "F" Then reixa.Columns(9).Locked = False




End Sub

Private Sub reixa_LostFocus()
'   AcroPDF1.Visible = True
End Sub

Private Sub reixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 Dim valtmp As String
 If reixa.col = 0 Then reixa.EditActive = False
 '-------
 bloquejar_camps_innecesaris
 If Not soldadores.Recordset.EOF Then
 'texteimpresio = atrim(impresores.Recordset!texteimpresio)
  If atrim(soldadores.Recordset!tipus) = "F" Then
     framebobines.Enabled = True
       Else: framebobines.Enabled = False: framepantones.visible = False
  End If
 End If
 
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
 'calcular_totals
 
 frameempalmes.visible = False
 framepantones.visible = False
 
End Sub

Private Sub reixabobines_AfterColUpdate(ByVal ColIndex As Integer)
If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
On Error Resume Next
 bobines.Recordset.Fields(reixabobines.Columns(ColIndex).DataField) = reixabobines.Columns(ColIndex).Text
 reixabobines.EditActive = False
'bobines.Recordset.Update
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
   Load formseleccio2
   formseleccio2.Data1.DatabaseName = camicomandes
   formseleccio2.Data1.RecordSource = "select codi,descripcio from operaris where maquina='S' and actiu<>0 order by codi asc"
   formseleccio2.caption = "Selecció d'Operari"
   formseleccio2.refrescar
   formseleccio2.Show 1
   If seleccioret = 1 Then
    escullir_operari = cadbl(formseleccio2.Data1.Recordset!codi)
    r = formseleccio2.Data1.Recordset!descripcio
   End If
   If cadbl(escullir_operari) = 0 Then MsgBox "Has d'escullir un operari per treballar"
 'Wend
 If cadbl(escullir_operari) = 0 Then escullir_operari = opvell
 Unload formseleccio2
End Function

Private Sub reixabobines_Error(ByVal DataError As Integer, Response As Integer)
If reixabobines.Columns(3) = "" Then reixabobines.Columns(3) = "0"
If reixabobines.Columns(4) = "" Then reixabobines.Columns(4) = "0"
Response = 0
End Sub

Private Sub reixabobines_GotFocus()
 etmetresbob.visible = False
 frameempalmes.visible = False
 framepantones.visible = False
 If reixabobines.col <> 7 Then
     framebobentrada.visible = True
   Else: framebobentrada.visible = False
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
guardar_reg_bobines
If IsNull(fila) Then fila = 0
If fila <> reixabobines.row Then
 'calcular_totals
End If
fila = reixabobines.row
If reixabobines.col <> 11 Then
     framebobentrada.visible = True
   Else: framebobentrada.visible = False
 End If
 If pescanutu > 0 Then
   pesnetanterior = cadbl(reixabobines.Columns("pesnet"))
   If cadbl(reixabobines.Columns("kilos")) - cadbl(pescanutu) > 0 Then
     reixabobines.Columns("pesnet") = cadbl(reixabobines.Columns("kilos")) - cadbl(pescanutu)
       Else: reixabobines.Columns("pesnet") = 0
   End If
   If pesnetanterior <> cadbl(reixabobines.Columns("pesnet")) Then bobines.Recordset.Edit: bobines.Recordset.Update
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

If etpesbascula.tag = "manual" Then Exit Sub
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
End If
 i = 0
 buffer = buffer & MSComm1.Input
 If Len(buffer) > 20 Then
   If InStr(1, buffer, "-") Then buffer = "0"
   If InStr(1, buffer, Chr$(13)) > 0 Then buffer = Mid(buffer, InStr(1, buffer, "+") + 1, InStr(1, buffer, Chr$(13)))
   'If InStr(1, buffer, ".") > 0 Then buffer = Mid(buffer, 1, InStr(1, buffer, ".") - 1) + "," + Mid(buffer, InStr(1, buffer, ".") + 1)
   etpesbascula = buffer
   buffer = ""
 End If
 Exit Sub
nopossarpes:
   nobascula = True
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
 If Not soldadores.Recordset.EOF Then
     If soldadores.Recordset!tipus = "C" Then
        canutustallats.visible = Not canutustallats.visible
       Else: canutustallats.visible = True
     End If
       Else: canutustallats.visible = True
 End If
 If DateDiff("s", tempseditant, Now) > 3 And tempseditant > 0 Then
   no_editar_bobines
 End If
 etproblema.visible = Not etproblema.visible
 mirarsiparar
 posarpesbascula
 On Error GoTo error_screen
 If controlactiu = "akjdfks" Then Me.caption = Me.caption
 On Error GoTo 0
 If client.caption = "" And (soldadores.Recordset.BOF And soldadores.Recordset.EOF) Then
   carregar_client_ntintersialtres
 End If
 
 If numop = 0 And Not formseleccio2.visible And reixa.Enabled Then
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
  rellotge.tag = cadbl(rellotge.tag) + 1
  If rellotge.tag = "100" Then
    'calcular_totals
    If Not existeix("c:\ordprog.ini") Then assignardecimalipunt
    rellotge.tag = "0"
  End If
  
  If Not soldadores.Recordset.EOF Then
    Select Case atrim(soldadores.Recordset!tipus)
       
       Case "C"
          Command1.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor: Command2.BackColor = Command4.BackColor
          Command30.BackColor = Command4.BackColor: Command31.BackColor = Command4.BackColor
          Command2.BackColor = &HFF8080
       Case "F"
          Command1.BackColor = Command4.BackColor: Command2.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor
          Command30.BackColor = Command4.BackColor: Command31.BackColor = Command4.BackColor
          Command3.BackColor = &HFF8080
       Case "A"
          Command1.BackColor = Command4.BackColor: Command2.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor
          Command30.BackColor = Command4.BackColor: Command31.BackColor = Command4.BackColor
          Command31.BackColor = &HFF8080
       Case "P"
          Command1.BackColor = Command4.BackColor: Command2.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor
          Command30.BackColor = Command4.BackColor: Command31.BackColor = Command4.BackColor
          Command30.BackColor = &HFF8080
        Case Else
          Command1.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor: Command2.BackColor = Command4.BackColor
          Command30.BackColor = Command4.BackColor: Command31.BackColor = Command4.BackColor
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
framepantones.visible = Not framepantones.visible
frameempalmes.visible = False
framebobentrada.visible = False
dblots.visible = True
dblots.AllowAddNew = True
dblots.AllowDelete = True
dblots.AllowUpdate = True
dblots.MarqueeStyle = 6
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub tpescanutu_LostFocus()
  If controlactiu = "agafarpesbascula" Then
     tpescanutu.tag = "pesarcanutu"
    Else: tpescanutu.tag = ""
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
   Set rstm = dbtmpb.OpenRecordset("SELECT comanda, numeromaquina FROM soldadores where comanda=" + atrim(numc))
   If rstm.EOF Then Exit Sub
   Set rstm = dbtmp.OpenRecordset("select descripcio from maquines where maquina='R' and codi=" + atrim(rstm!numeromaquina))
   If rstm.EOF Then Exit Sub
   
   Open llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetarqualitatreducciocilindresoldadores.prn" For Input As #1
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

Private Sub unitatsxfunda_DblClick()
  Dim vunitatsxfunda As String
  vunitatsxfunda = InputBox("Escriu quantes unitats vols a cada funda." + vbNewLine + "Escriu [CAP] si no vols utilitzar fundes.", "Utilitzar fundes")
  If UCase(vunitatsxfunda) = "CAP" Then unitatsxfunda = 0: guarda_totals
  If cadbl(vunitatsxfunda) > 0 Then
            unitatsxfunda = vunitatsxfunda
            guarda_totals
  End If
End Sub
