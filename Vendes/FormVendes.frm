VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formvendes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Albarà d'expedicions"
   ClientHeight    =   9705
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   15960
   Icon            =   "FormVendes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   15960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   345
      TabIndex        =   131
      Top             =   7350
      Width           =   10275
      Begin VB.Label etimpostenvasos 
         BackColor       =   &H006BEBB1&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   0
         TabIndex        =   132
         Top             =   0
         Width           =   10290
      End
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Càlcul Impost Env."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   10920
      TabIndex        =   128
      Top             =   8205
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton bpendents 
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   2790
      Picture         =   "FormVendes.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   123
      ToolTipText     =   "Llista pendents"
      Top             =   75
      Width           =   765
   End
   Begin VB.CommandButton Command10 
      Height          =   360
      Left            =   2220
      Picture         =   "FormVendes.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   97
      ToolTipText     =   "Actualitzar/Grabar Registres"
      Top             =   75
      Width           =   525
   End
   Begin VB.CommandButton benviat 
      BackColor       =   &H0000FF00&
      Caption         =   "Marcar enviat"
      Height          =   285
      Left            =   9660
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   1140
      Width           =   1380
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   9585
      TabIndex        =   121
      Top             =   1140
      Width           =   1695
      Begin VB.Label etdataenviament 
         Caption         =   "Data Enviament"
         Height          =   270
         Left            =   90
         TabIndex        =   122
         Top             =   75
         Width           =   1620
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imprimir Frontals"
      Height          =   615
      Left            =   8655
      TabIndex        =   114
      Top             =   -30
      Width           =   2850
      Begin VB.CheckBox Checkpapersfrontalsimpresos 
         Enabled         =   0   'False
         Height          =   195
         Left            =   2670
         TabIndex        =   135
         ToolTipText     =   "Papers frontals impresos."
         Top             =   45
         Width           =   225
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FDDECE&
         Caption         =   "Et.Bases"
         Height          =   375
         Left            =   915
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "Imprimir Albarà d'expedicions"
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Paper Std"
         Height          =   375
         Left            =   1785
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Imprimir Albarà d'expedicions"
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00F3B378&
         Caption         =   "EAN128"
         Height          =   360
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   195
         Width           =   855
      End
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   14040
      Picture         =   "FormVendes.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   90
      Width           =   600
   End
   Begin VB.CommandButton Command11 
      Height          =   375
      Left            =   11550
      Picture         =   "FormVendes.frx":1968
      Style           =   1  'Graphical
      TabIndex        =   111
      ToolTipText     =   "Organitzar PackingList"
      Top             =   90
      Width           =   810
   End
   Begin VB.CommandButton bdesbloquejarsap 
      Height          =   375
      Left            =   13200
      Picture         =   "FormVendes.frx":1EF2
      Style           =   1  'Graphical
      TabIndex        =   96
      ToolTipText     =   "Desbloquejar SAP"
      Top             =   90
      Width           =   810
   End
   Begin VB.CommandButton Command8 
      Height          =   375
      Left            =   12390
      Picture         =   "FormVendes.frx":247C
      Style           =   1  'Graphical
      TabIndex        =   92
      ToolTipText     =   "Imprimir Albarà d'expedicions"
      Top             =   90
      Width           =   810
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      DisabledPicture =   "FormVendes.frx":2A06
      Height          =   375
      Left            =   14670
      Picture         =   "FormVendes.frx":5220
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Enviar albarà al SAP (No es pot retrocedir)"
      Top             =   90
      Width           =   1200
   End
   Begin VB.Frame Framecapcalera 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Capçalera de l'Albarà de Venda"
      Height          =   2250
      Left            =   285
      TabIndex        =   0
      Top             =   435
      Width           =   15540
      Begin VB.Frame Frame7 
         Caption         =   "Transportista"
         Height          =   1485
         Left            =   4155
         TabIndex        =   22
         Top             =   150
         Width           =   6930
         Begin VB.CommandButton bassignartransport 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Assignar transport"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6075
            Style           =   1  'Graphical
            TabIndex        =   138
            Top             =   165
            Width           =   780
         End
         Begin VB.TextBox cidtransport 
            DataField       =   "id_transport"
            DataSource      =   "datacapcalera"
            Height          =   285
            Left            =   1035
            TabIndex        =   88
            Top             =   225
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox cobservacionstransport 
            DataField       =   "observacionsports"
            DataSource      =   "datacapcalera"
            Height          =   330
            Left            =   480
            MaxLength       =   50
            TabIndex        =   30
            ToolTipText     =   "Observacions pel transportista"
            Top             =   1035
            Width           =   6195
         End
         Begin VB.ComboBox combotipusdeports 
            DataField       =   "tipusports"
            DataSource      =   "datacapcalera"
            Height          =   315
            ItemData        =   "FormVendes.frx":7A3A
            Left            =   1215
            List            =   "FormVendes.frx":7A47
            TabIndex        =   24
            Top             =   645
            Width           =   2325
         End
         Begin VB.ComboBox combotransportista 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1260
            TabIndex        =   23
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label etmetrescubicscalculats 
            BackStyle       =   0  'Transparent
            Caption         =   "-----------------"
            DataSource      =   "datacapcalera"
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
            Height          =   300
            Left            =   3570
            TabIndex        =   136
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label etmetrescubics 
            BackStyle       =   0  'Transparent
            DataSource      =   "datacapcalera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3690
            TabIndex        =   133
            Top             =   600
            Width           =   990
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Nom Transport:"
            Height          =   195
            Left            =   60
            TabIndex        =   58
            Top             =   285
            Width           =   1260
         End
         Begin VB.Label Label6 
            Caption         =   "Obs:"
            Height          =   195
            Left            =   60
            TabIndex        =   31
            Top             =   1095
            Width           =   1290
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipus de Ports:"
            Height          =   195
            Left            =   45
            TabIndex        =   25
            Top             =   660
            Width           =   1305
         End
      End
      Begin VB.TextBox cobservacions 
         DataField       =   "observacions"
         DataSource      =   "datacapcalera"
         Height          =   330
         Left            =   4455
         MaxLength       =   100
         TabIndex        =   19
         Top             =   1650
         Width           =   11010
      End
      Begin VB.TextBox cdataalbara 
         DataField       =   "dataalbara"
         DataSource      =   "datacapcalera"
         Height          =   285
         Left            =   1125
         TabIndex        =   17
         Top             =   1590
         Width           =   1365
      End
      Begin VB.TextBox cnumalbara 
         DataField       =   "numalbara"
         DataSource      =   "datacapcalera"
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1275
         Width           =   1830
      End
      Begin VB.ComboBox comboquifactura 
         DataField       =   "empresa"
         DataSource      =   "datacapcalera"
         Height          =   315
         ItemData        =   "FormVendes.frx":7A66
         Left            =   1110
         List            =   "FormVendes.frx":7A70
         TabIndex        =   5
         Top             =   240
         Width           =   2475
      End
      Begin VB.Frame frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dades del Client                                                                  "
         Height          =   1515
         Left            =   11115
         TabIndex        =   3
         Top             =   135
         Width           =   4350
         Begin VB.CommandButton bcanvienvio 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Canvi Direcció"
            Height          =   225
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   1230
            Width           =   1290
         End
         Begin VB.Label etgrupclient 
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
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   1515
            TabIndex        =   101
            Top             =   0
            Width           =   2820
         End
         Begin VB.Label etinfodelclient 
            BackStyle       =   0  'Transparent
            Caption         =   "Informació del client"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1275
            Left            =   120
            TabIndex        =   32
            Top             =   210
            Width           =   4065
         End
      End
      Begin VB.CheckBox Checkalbaravalorat 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Albarà valorat"
         DataField       =   "albaravalorat"
         DataSource      =   "datacapcalera"
         Height          =   450
         Left            =   3195
         TabIndex        =   98
         Top             =   1245
         Width           =   945
      End
      Begin VB.Label etfiproduccio 
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
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1935
         TabIndex        =   137
         Top             =   1965
         Width           =   13575
      End
      Begin VB.Image logoinplacsa 
         Height          =   720
         Left            =   675
         Picture         =   "FormVendes.frx":7A86
         Top             =   555
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Observacions:"
         Height          =   225
         Left            =   3300
         TabIndex        =   20
         Top             =   1740
         Width           =   1050
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Alb:"
         Height          =   225
         Left            =   180
         TabIndex        =   18
         Top             =   1620
         Width           =   825
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Albarà:"
         Height          =   225
         Left            =   195
         TabIndex        =   16
         Top             =   1305
         Width           =   825
      End
      Begin VB.Image logoplasel 
         Height          =   555
         Left            =   1230
         Picture         =   "FormVendes.frx":F2C8
         Top             =   555
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qui factura?"
         Height          =   345
         Left            =   135
         TabIndex        =   6
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.CheckBox checknogenerarfitxersap 
      BackColor       =   &H006BEBB1&
      Caption         =   "Check1"
      Height          =   195
      Left            =   15555
      TabIndex        =   112
      ToolTipText     =   "No generar fitxer per importar a SAP"
      Top             =   135
      Width           =   210
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   2355
   End
   Begin VB.CommandButton Command9 
      Height          =   360
      Left            =   2595
      Picture         =   "FormVendes.frx":128CE
      Style           =   1  'Graphical
      TabIndex        =   93
      ToolTipText     =   "Actualitzar bobines per entregar"
      Top             =   2670
      Width           =   420
   End
   Begin VB.CommandButton Command6 
      Height          =   360
      Left            =   1650
      Picture         =   "FormVendes.frx":12E58
      Style           =   1  'Graphical
      TabIndex        =   89
      ToolTipText     =   "Actualitzar/Grabar Registres"
      Top             =   75
      Width           =   570
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF8080&
      Height          =   390
      Left            =   12090
      Picture         =   "FormVendes.frx":133E2
      Style           =   1  'Graphical
      TabIndex        =   85
      ToolTipText     =   "Canvia el modo de visualització de les dades"
      Top             =   4965
      Width           =   450
   End
   Begin VB.Frame frameliniesalpaper 
      BackColor       =   &H0080FF80&
      Caption         =   "Visualització sobre el paper"
      Height          =   3390
      Left            =   2190
      TabIndex        =   86
      Top             =   8445
      Visible         =   0   'False
      Width           =   12315
      Begin VB.ListBox llistasobrepaper 
         Height          =   2595
         Left            =   915
         TabIndex        =   90
         Top             =   510
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label etsobrepaper 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3000
         Left            =   90
         TabIndex        =   87
         Top             =   240
         Width           =   12180
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   360
      Left            =   1080
      Picture         =   "FormVendes.frx":1396C
      Style           =   1  'Graphical
      TabIndex        =   84
      ToolTipText     =   "Modificar capçalera d'albarà"
      Top             =   75
      Width           =   570
   End
   Begin VB.Frame Framebobines 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Bobines seleccionades"
      Height          =   3645
      Left            =   12615
      TabIndex        =   26
      ToolTipText     =   "Verd si Kg comanda +-10% seleccionats sino vermell."
      Top             =   4815
      Width           =   3180
      Begin VB.CommandButton seltot 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2655
         TabIndex        =   29
         ToolTipText     =   "Des/Seleccionar-ho tot"
         Top             =   165
         Width           =   465
      End
      Begin VB.ListBox llistabobinessel 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2940
         ItemData        =   "FormVendes.frx":13EF6
         Left            =   75
         List            =   "FormVendes.frx":13EF8
         Style           =   1  'Checkbox
         TabIndex        =   27
         Top             =   195
         Width           =   2550
      End
      Begin VB.Label etentregaparcial 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   142
         Top             =   3420
         Width           =   2880
      End
      Begin VB.Label etdemanats 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   141
         Top             =   3285
         Width           =   2880
      End
      Begin VB.Label ettotals 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   28
         Top             =   3120
         Width           =   2880
      End
   End
   Begin VB.CommandButton Command3 
      Height          =   360
      Left            =   315
      Picture         =   "FormVendes.frx":13EFA
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Crear nou albarà de venda"
      Top             =   75
      Width           =   765
   End
   Begin VB.Data datacapcalera 
      Caption         =   "Albarans Venda"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   3555
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from capcaleraalbara order by numalbara desc"
      Top             =   60
      Width           =   2280
   End
   Begin VB.Data datalinies 
      Caption         =   "datalinies"
      Connect         =   "Access"
      DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\vendes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2820
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "liniesalbara"
      Top             =   2670
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.Frame Framepeualbara 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Peu de l'albarà"
      Height          =   1230
      Left            =   285
      TabIndex        =   12
      Top             =   8490
      Width           =   15480
      Begin VB.Data dataliniespeu 
         Caption         =   "dataliniespeu"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   7740
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   375
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSDBGrid.DBGrid reixapeu 
         Bindings        =   "FormVendes.frx":14484
         Height          =   945
         Left            =   60
         OleObjectBlob   =   "FormVendes.frx":1449C
         TabIndex        =   13
         Top             =   225
         Width           =   15150
      End
   End
   Begin VB.Frame Framecontrolslinia 
      Height          =   525
      Left            =   285
      TabIndex        =   7
      Top             =   2625
      Width           =   1875
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   30
         Picture         =   "FormVendes.frx":14E90
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Afegir entrega de Comanda"
         Top             =   120
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   945
         Picture         =   "FormVendes.frx":1541A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Eliminacio del registre"
         Top             =   120
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   480
         Picture         =   "FormVendes.frx":159A4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Modificar Registres"
         Top             =   120
         Width           =   420
      End
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   1395
         Picture         =   "FormVendes.frx":15F2E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Actualitzar/Grabar Registres"
         Top             =   120
         Width           =   420
      End
   End
   Begin VB.Frame Framelinies 
      BackColor       =   &H00C0FFC0&
      Caption         =   "                                        Linies de l'albarà"
      Height          =   1800
      Left            =   240
      TabIndex        =   1
      Top             =   3045
      Width           =   15525
      Begin MSDBGrid.DBGrid reixalinies 
         Bindings        =   "FormVendes.frx":164B8
         Height          =   1500
         Left            =   105
         OleObjectBlob   =   "FormVendes.frx":164CD
         TabIndex        =   4
         Top             =   225
         Width           =   15180
      End
      Begin VB.Label etresumbobinesipalets 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "                              "
         Height          =   195
         Left            =   3315
         TabIndex        =   118
         Top             =   0
         Width           =   1350
      End
   End
   Begin VB.Frame framedadeslinia 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Detall de la linia d'albarà"
      Height          =   3660
      Left            =   270
      TabIndex        =   2
      Top             =   4830
      Width           =   12315
      Begin VB.TextBox Text32 
         DataField       =   "kgimpostenvasos"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   11340
         TabIndex        =   126
         Text            =   " "
         Top             =   2580
         Width           =   840
      End
      Begin VB.TextBox Text12 
         DataField       =   "quantitat"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   11445
         TabIndex        =   124
         Text            =   " "
         Top             =   1680
         Width           =   750
      End
      Begin VB.TextBox Text11 
         DataField       =   "kgtotalsnets"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   3795
         TabIndex        =   109
         Text            =   " "
         Top             =   1950
         Width           =   705
      End
      Begin VB.TextBox Text10 
         DataField       =   "macroperforat"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   8775
         MaxLength       =   1
         TabIndex        =   107
         Text            =   " "
         Top             =   1680
         Width           =   300
      End
      Begin VB.TextBox Text3 
         DataField       =   "microperforat"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   7815
         MaxLength       =   1
         TabIndex        =   105
         Text            =   " "
         Top             =   1665
         Width           =   300
      End
      Begin VB.ComboBox combotipusentrega 
         BackColor       =   &H00FFC0FF&
         DataField       =   "tipusdeentrega"
         DataSource      =   "datalinies"
         Height          =   315
         ItemData        =   "FormVendes.frx":17234
         Left            =   11430
         List            =   "FormVendes.frx":1723E
         TabIndex        =   99
         Top             =   1050
         Width           =   615
      End
      Begin VB.TextBox Text2 
         DataField       =   "pespalets"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   5400
         TabIndex        =   94
         Text            =   " "
         Top             =   1935
         Width           =   480
      End
      Begin VB.TextBox Text30 
         DataField       =   "tipusproducte"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   10380
         TabIndex        =   57
         Text            =   " "
         Top             =   1980
         Width           =   1830
      End
      Begin VB.TextBox Text29 
         DataField       =   "lotinplacsa"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   9705
         TabIndex        =   56
         Text            =   " "
         Top             =   780
         Width           =   1380
      End
      Begin VB.TextBox Text28 
         DataField       =   "unitats"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   8520
         TabIndex        =   55
         Text            =   " "
         Top             =   1950
         Width           =   825
      End
      Begin VB.TextBox Text27 
         DataField       =   "metreslineals"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   6960
         TabIndex        =   54
         Text            =   " "
         Top             =   1935
         Width           =   825
      End
      Begin VB.TextBox Text26 
         DataField       =   "kgtotalsbruts"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   2355
         TabIndex        =   53
         Text            =   " "
         Top             =   1935
         Width           =   705
      End
      Begin VB.TextBox Text25 
         DataField       =   "numcalloff"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   4110
         TabIndex        =   52
         Text            =   " "
         Top             =   1650
         Width           =   2775
      End
      Begin VB.TextBox Text24 
         DataField       =   "codibarres"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   10365
         TabIndex        =   51
         Text            =   " "
         Top             =   1365
         Width           =   1845
      End
      Begin VB.TextBox Text23 
         DataField       =   "datafabricacio"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   7980
         TabIndex        =   50
         Text            =   "  "
         Top             =   1365
         Width           =   1395
      End
      Begin VB.TextBox Text22 
         DataField       =   "refclientdeclient"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   4785
         TabIndex        =   49
         Text            =   " "
         Top             =   1365
         Width           =   2100
      End
      Begin VB.TextBox Text21 
         DataField       =   "refclient"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   4110
         TabIndex        =   48
         Text            =   " "
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox Text20 
         DataField       =   "descripciomides"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   7350
         TabIndex        =   47
         Text            =   " "
         Top             =   510
         Width           =   4860
      End
      Begin VB.TextBox Text19 
         DataField       =   "mesuraespesor"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   4845
         TabIndex        =   46
         Text            =   " "
         Top             =   510
         Width           =   1065
      End
      Begin VB.TextBox Text18 
         DataField       =   "espesor"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   2835
         TabIndex        =   45
         Text            =   " "
         Top             =   510
         Width           =   690
      End
      Begin VB.TextBox Text17 
         DataField       =   "preuvenda"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   8670
         TabIndex        =   44
         Text            =   " "
         Top             =   225
         Width           =   1005
      End
      Begin VB.TextBox Text16 
         DataField       =   "quantitat"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   7350
         TabIndex        =   43
         Text            =   " "
         Top             =   225
         Width           =   750
      End
      Begin VB.TextBox Text15 
         DataField       =   "descripcioproducte"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   3825
         TabIndex        =   42
         Text            =   " "
         Top             =   225
         Width           =   2820
      End
      Begin VB.TextBox Text14 
         DataField       =   "unitatmesura"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   10290
         TabIndex        =   41
         Text            =   " "
         Top             =   225
         Width           =   1065
      End
      Begin VB.TextBox Text13 
         DataField       =   "observacionslinia"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   1125
         TabIndex        =   40
         Text            =   " "
         Top             =   2280
         Width           =   11085
      End
      Begin VB.TextBox Text9 
         DataField       =   "numbobs"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   1140
         TabIndex        =   39
         Text            =   " "
         Top             =   1935
         Width           =   495
      End
      Begin VB.TextBox Text8 
         DataField       =   "numcontracte"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   1140
         TabIndex        =   38
         Text            =   " "
         Top             =   1650
         Width           =   2130
      End
      Begin VB.TextBox Text7 
         DataField       =   "numcomandaclideclient"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   1140
         TabIndex        =   37
         Text            =   " "
         Top             =   1365
         Width           =   2145
      End
      Begin VB.TextBox Text6 
         DataField       =   "numcomandacli"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   1140
         TabIndex        =   36
         Text            =   " "
         Top             =   1080
         Width           =   2145
      End
      Begin VB.TextBox Text5 
         DataField       =   "marcailinia"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   1140
         TabIndex        =   35
         Text            =   " "
         Top             =   795
         Width           =   7440
      End
      Begin VB.TextBox Text4 
         DataField       =   "ampladamaterial"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   1140
         TabIndex        =   34
         Text            =   " "
         Top             =   510
         Width           =   945
      End
      Begin VB.TextBox Text1 
         DataField       =   "codiproducte"
         DataSource      =   "datalinies"
         Height          =   285
         Left            =   1140
         TabIndex        =   33
         Text            =   " "
         Top             =   225
         Width           =   1740
      End
      Begin VB.TextBox cKgImpost100per100 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         DataField       =   "KgImpost100per100"
         DataSource      =   "datalinies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   11580
         TabIndex        =   130
         Text            =   " 1234"
         Top             =   2925
         Width           =   600
      End
      Begin VB.Label ettanx100mermalot 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   10425
         TabIndex        =   140
         Top             =   3150
         Width           =   1800
      End
      Begin VB.Label ettanper100merma 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   10350
         TabIndex        =   139
         Top             =   3330
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "-1(No impost)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   10440
         TabIndex        =   134
         Top             =   2760
         Width           =   1470
      End
      Begin VB.Label etKgImpost100per100 
         BackStyle       =   0  'Transparent
         Caption         =   "Impost (100%):"
         Height          =   270
         Left            =   10410
         TabIndex        =   129
         Top             =   2970
         Width           =   1170
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg Impost: "
         Height          =   270
         Left            =   10425
         TabIndex        =   127
         Top             =   2565
         Width           =   1110
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantitat:"
         Height          =   270
         Left            =   10755
         TabIndex        =   125
         Top             =   1710
         Width           =   870
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg Nets:"
         Height          =   270
         Left            =   3150
         TabIndex        =   110
         Top             =   1995
         Width           =   675
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "MacroP:"
         Height          =   270
         Left            =   8160
         TabIndex        =   108
         Top             =   1695
         Width           =   630
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "MicroP:"
         Height          =   270
         Left            =   7230
         TabIndex        =   106
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipus d'entrega"
         Height          =   270
         Left            =   11145
         TabIndex        =   100
         Top             =   810
         Width           =   1170
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Pes Palets:"
         Height          =   270
         Left            =   4530
         TabIndex        =   95
         Top             =   1965
         Width           =   870
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipus Prod:"
         Height          =   270
         Left            =   9450
         TabIndex        =   83
         Top             =   2025
         Width           =   900
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Inplacsa:"
         Height          =   270
         Left            =   8700
         TabIndex        =   82
         Top             =   825
         Width           =   990
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Unitats:"
         Height          =   270
         Left            =   7935
         TabIndex        =   81
         Top             =   1965
         Width           =   705
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Mtrs Lineals:"
         Height          =   270
         Left            =   5985
         TabIndex        =   80
         Top             =   1995
         Width           =   990
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg Bruts:"
         Height          =   270
         Left            =   1695
         TabIndex        =   79
         Top             =   1965
         Width           =   990
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Calloff:"
         Height          =   270
         Left            =   3330
         TabIndex        =   78
         Top             =   1665
         Width           =   1155
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Codi Barres:"
         Height          =   270
         Left            =   9465
         TabIndex        =   77
         Top             =   1395
         Width           =   990
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Fàb:"
         Height          =   270
         Left            =   6960
         TabIndex        =   76
         Top             =   1395
         Width           =   765
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. Client de Client:"
         Height          =   270
         Left            =   3315
         TabIndex        =   75
         Top             =   1395
         Width           =   1665
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. Client:"
         Height          =   270
         Left            =   3300
         TabIndex        =   74
         Top             =   1110
         Width           =   975
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripció Mides:"
         Height          =   270
         Left            =   6000
         TabIndex        =   73
         Top             =   525
         Width           =   1425
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Mesura espesor:"
         Height          =   270
         Left            =   3600
         TabIndex        =   72
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Espesor:"
         Height          =   270
         Left            =   2130
         TabIndex        =   71
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Preu:"
         Height          =   270
         Left            =   8220
         TabIndex        =   70
         Top             =   255
         Width           =   480
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantitat:"
         Height          =   270
         Left            =   6660
         TabIndex        =   69
         Top             =   255
         Width           =   870
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripció:"
         Height          =   270
         Left            =   2970
         TabIndex        =   68
         Top             =   255
         Width           =   870
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Unitat:"
         Height          =   270
         Left            =   9810
         TabIndex        =   67
         Top             =   255
         Width           =   600
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Observacions:"
         Height          =   270
         Left            =   30
         TabIndex        =   66
         Top             =   2295
         Width           =   1155
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Bobines:"
         Height          =   270
         Left            =   45
         TabIndex        =   65
         Top             =   1965
         Width           =   1155
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Codi Inplacsa:"
         Height          =   270
         Left            =   45
         TabIndex        =   64
         Top             =   255
         Width           =   1155
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Contracte:"
         Height          =   270
         Left            =   45
         TabIndex        =   63
         Top             =   1665
         Width           =   1155
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "NºCom.Cli/Cli:"
         Height          =   270
         Left            =   45
         TabIndex        =   62
         Top             =   1395
         Width           =   1155
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "NºCom. Client:"
         Height          =   270
         Left            =   45
         TabIndex        =   61
         Top             =   1095
         Width           =   1155
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Marca i Linia:"
         Height          =   270
         Left            =   45
         TabIndex        =   60
         Top             =   810
         Width           =   1155
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Ample Mat:"
         Height          =   270
         Left            =   45
         TabIndex        =   59
         Top             =   525
         Width           =   1155
      End
   End
   Begin VB.Label etean128 
      BackColor       =   &H0000FFFF&
      Caption         =   "Aquest client vol Frontal EAN128"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   12915
      TabIndex        =   117
      Top             =   2670
      Width           =   2820
   End
   Begin VB.Label etfiltreactivat 
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
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   5850
      MouseIcon       =   "FormVendes.frx":17248
      TabIndex        =   103
      Top             =   195
      Width           =   4470
   End
   Begin VB.Label etmissatge 
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
      Height          =   315
      Left            =   4215
      TabIndex        =   102
      Top             =   2685
      Width           =   8685
   End
   Begin VB.Label ettraspasasap 
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
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   5835
      TabIndex        =   91
      Top             =   15
      Width           =   2910
   End
   Begin VB.Menu mfiltre 
      Caption         =   "Filtre"
      Begin VB.Menu mpendentsdesap 
         Caption         =   "Pendents de passar a SAP"
      End
      Begin VB.Menu mpendentsenviats 
         Caption         =   "Pendents marcar ENVIATS"
      End
      Begin VB.Menu menviatsasap 
         Caption         =   "Enviats a SAP"
      End
      Begin VB.Menu mtotsdeunclient 
         Caption         =   "Tots de un client"
      End
   End
   Begin VB.Menu menurma 
      Caption         =   "Recepció RMA"
   End
   Begin VB.Menu malbcontenidors 
      Caption         =   "Albarans de contenidors"
   End
   Begin VB.Menu mtransportistes 
      Caption         =   "Transportistes"
   End
   Begin VB.Menu mescanejar 
      Caption         =   "Escanejar"
      Begin VB.Menu mescanejaralbaransproveidor 
         Caption         =   "Albarans del proveïdor"
      End
      Begin VB.Menu mcertificaciolotsproveidor 
         Caption         =   "Certificació de Lots de proveïdor"
      End
      Begin VB.Menu malbaransSAPsegellats 
         Caption         =   "Albarans del SAP segellats i CMR"
      End
   End
   Begin VB.Menu mllistatg 
      Caption         =   "Llistats"
      Begin VB.Menu llistatentreguestransport 
         Caption         =   "Llistat entregues transportistes per dia"
      End
      Begin VB.Menu mvveure 
         Caption         =   "Veure CMR o Albarà SAP Signat"
         Begin VB.Menu malbSAP 
            Caption         =   "Albarà SAP Signat"
         End
         Begin VB.Menu mCMR 
            Caption         =   "CMR Signat"
         End
      End
   End
   Begin VB.Menu mutil 
      Caption         =   "Utilitats"
      Begin VB.Menu mveurebases 
         Caption         =   "Veure palets embolicats de l'albarà"
      End
      Begin VB.Menu mfacturarclixes 
         Caption         =   "Facturar només Clixes"
      End
      Begin VB.Menu mpemail 
         Caption         =   "Parametres Email"
         Begin VB.Menu musrsmtp 
            Caption         =   "Usuari SMTP"
         End
         Begin VB.Menu mpwdsmpt 
            Caption         =   "Password SMTP"
         End
      End
      Begin VB.Menu mpersonalitzarbases 
         Caption         =   "Personalitzar bases"
      End
      Begin VB.Menu mprogramadembolicarpalets 
         Caption         =   "Programa d' embolicar Palets"
      End
      Begin VB.Menu mCMRsNoEscanejats 
         Caption         =   "CMR's no escanejats"
         Begin VB.Menu mtotsCmrNoescanejats 
            Caption         =   "Tots"
         End
         Begin VB.Menu CMRmarcatssensepdf 
            Caption         =   "Marcats com escanejats pero sense PDF"
         End
      End
   End
End
Attribute VB_Name = "formvendes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vTipusImpostLinia As String
Dim vObrirAlbaràSAP As Double
Dim valbaraSAPportaimpost As Boolean
Dim vTkilosseleccionats As Double
Dim vTkilosseleccionatsmesparcials As Double
Dim dbvendes As Database
Dim dbimpost As Database
'Dim arguments As Variant
Dim vidliniaclixe As Double
Dim vdesccosmailclixes As String
Dim vcomandesambproforma As String
Dim vimpostinclosalPVP As Boolean
Dim vcancelarSAP As Boolean
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


Private Sub Combo1_Change()

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   
End Sub

Private Sub a_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub alta_Click()
   escullir_comandaxralbaranar False
   
End Sub
Function jahihaaquestacomanda(numc As Double, numalbara As Double) As Boolean
  Dim rst As Recordset
  Set rst = datacapcalera.Database.OpenRecordset("select * from liniesalbara where lotinplacsa=" + atrim(numc) + " and numalbara=" + atrim(numalbara))
  If Not rst.EOF Then jahihaaquestacomanda = True
  Set rst = Nothing
End Function
Function comandaentregadaambT(numc As Double) As Boolean
  Dim rst As Recordset
  Set rst = datacapcalera.Database.OpenRecordset("select * from liniesalbara where tipusdeentrega='T' and lotinplacsa=" + atrim(numc))
  If Not rst.EOF Then comandaentregadaambT = True
  Set rst = Nothing
End Function

Function buscarestatcomanda(numc As Double) As String
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("select proximaseccio from comandes where comanda=" + atrim(numc))
  buscarestatcomanda = "T"
  If rst.EOF Then Exit Function
  buscarestatcomanda = atrim(rst!proximaseccio)
  Set rst = Nothing
End Function
Function buscarmonedadelclient(vcodiclient As Double, vcodicomptable As Double) As String
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select * from clients_codiscomptables where codifabricacio=" + atrim(vcodiclient) + " and codicomptable=" + atrim(vcodicomptable))
   If Not rst.EOF Then
     buscarmonedadelclient = atrim(rst!moneda)
     If buscarmonedadelclient = "" Then buscarmonedadelclient = "Euros"
   End If
   Set rst = Nothing
End Function
Function mirarsiaquestalbaraesdecalloffisiente(vnumc As Double) As Byte
    Dim rst As Recordset
    Set rst = dbcomandes.OpenRecordset("select client from comandes where comanda=" + atrim(vnumc))
    If rst.EOF Then Exit Function
    If cadbl(rst!client) <> 6841 Then mirarsiaquestalbaraesdecalloffisiente = 0: Exit Function 'no es client calloff
    mirarsiaquestalbaraesdecalloffisiente = 1  'es client calloff pero no te calloff assignats
    Set rst = dbbaixes.OpenRecordset("select numcalloff from bobinesent where (numalbara=null or numalbara=0) and comanda=" + atrim(vnumc) + " and (numcalloff<>'' and numcalloff<>null)")
    If Not rst.EOF Then mirarsiaquestalbaraesdecalloffisiente = 2  'es client calloff i te calloffs assignats
    Set rst = Nothing
End Function
Function comandaambnumerodeproforma(numc As Double, crearalbaranou As Boolean, vensenyarmissatge As Boolean) As Boolean
    Dim rst As Recordset
    Dim v As String
    Dim vnumc As Double
    Dim vcomandeslotafegit As String
    Dim vcomandeslotalbara As String
    
    vnumc = numc
    comandaambnumerodeproforma = True
    'es nou
    If crearalbaranou Then
        v = comandesproforma(numc)
        If v <> "" And vensenyarmissatge Then
             MsgBox "PROFORMA. Aquestes comandes han d'anar juntament amb aquestes altres" + Chr(10) + v, vbInformation + vbOKOnly, "Proformes"
        End If
        vcomandesambproforma = v
        numc = vnumc
        Exit Function
    End If
    'no es nou
    vcomandeslotafegit = comandesproforma(numc)
    If Not datalinies.Recordset.EOF Then vcomandeslotalbara = comandesproforma(datalinies.Recordset!lotinplacsa)
    'mirar si afegit està a albarà
    If vcomandeslotalbara <> "" Then
        If InStr(1, " " + vcomandeslotalbara, atrim(numc)) = 0 Then
           comandaambnumerodeproforma = False
           MsgBox "La comanda " + atrim(numc) + " no està a la llista de comandes de la PROFORMA.", vbCritical, "ERROR"
        End If
          Else
            If vcomandeslotafegit <> "" Then
                If InStr(1, " " + vcomandeslotafegit, atrim(datalinies.Recordset!lotinplacsa)) = 0 Then
                   comandaambnumerodeproforma = False
                   MsgBox "La comanda " + atrim(numc) + " està agrupada amb PROFORMA i no coincideix amb les d'aquest albarà.", vbCritical, "ERROR"
                End If
            End If
    End If
    numc = vnumc
End Function
Function comandesproforma(numc As Double) As String
    Dim rst As Recordset
    Dim v As String
    Set rst = dbcomandes.OpenRecordset("select obsext2,client from comandes where comanda=" + atrim(numc))
    If rst.EOF Then Exit Function
    If atrim(rst!obsext2) <> "" Then
      Set rst = dbcomandes.OpenRecordset("select comanda from comandes where client=" + atrim(rst!client) + " and obsext2='" + atrim(rst!obsext2) + "'")
      While Not rst.EOF
        v = v + " " + atrim(rst!comanda)
        rst.MoveNext
      Wend
    End If
    comandesproforma = v
End Function
Sub escullir_comandaxralbaranar(crearalbaranou As Boolean)
   Dim numc As Double
   If datalinies.Recordset.EditMode > 0 Or datacapcalera.Recordset.EditMode > 0 Then MsgBox "Abans d'afegir comanda primer guarda el registre actual", vbCritical, "Error": Exit Sub
   If Not datacapcalera.Recordset.EOF Then numc = datacapcalera.Recordset!numalbara
   
   datacapcalera.RecordSource = "select * from capcaleraalbara order by numalbara desc"
   datacapcalera.Refresh
   datacapcalera.Recordset.FindFirst "numalbara=" + atrim(numc)
   numc = 0
   etfiltreactivat = ""
   numc = cadbl(InputBox("Entra el numero de comanda que vols afegir.", "Comanda"))
   If numc = 0 Then Exit Sub
   If comandaentregadaambT(numc) Then MsgBox "Aquesta comanda ja està entregada en TOTAL no pots tornar a albaranar-lo.", vbCritical, "Error": Exit Sub
  ' If buscarestatcomanda(numc) = "T" Then MsgBox "Aquesta comanda ja està entregada no es pot albaranar", vbCritical, "Error": Exit Sub
   If Not crearalbaranou Then If jahihaaquestacomanda(numc, cadbl(cnumalbara)) Then MsgBox "Aquesta comanda ja està afegida a aquest albarà", vbCritical, "Error": Exit Sub
   If Not comandaestaapuntperexpedir(numc) Then Exit Sub
   If comandashadalbaranarsola(numc, crearalbaranou) Then Exit Sub
   If Not comandaambnumerodeproforma(numc, crearalbaranou, True) Then Exit Sub
   If mirarsiaquestalbaraesdecalloffisiente(numc) = 1 Then
       If MsgBox("Aquest Lot necessita Call-off per entregar-lo i no hi ha cap palet amb call-off assignat." + Chr(10) + "Vols continuar igualment?", vbCritical + vbYesNo + vbDefaultButton2, "Error") = vbNo Then Exit Sub
   End If
   crea_capcalera_i_linies numc, crearalbaranou
   If numc > 0 Then
    carregar_bobinesentrega numc, cnumalbara
    sumarkilosimetresseleccionats
   End If
   modificar_Click
   Command1_Click
End Sub
Function comandashadalbaranarsola(vnumc As Double, Optional crearalbaranou As Boolean) As Boolean
   Dim rst As Recordset
   Dim rst2 As Recordset
   comandashadalbaranarsola = True
   Set rst = dbcomandes.OpenRecordset("select * from comandes_extres where comanda=" + atrim(vnumc))
   If rst.EOF Then
       Exit Function
         Else
           'mirar si s'ha d'albaranar sola
           If Not crearalbaranou Then
             If cadbl(rst!pararaexpedicions) > 1 And datalinies.Recordset.RecordCount > 0 Then MsgBox "Aquesta comanda s'ha d'albaranar sola.", vbCritical, "Error": Exit Function
            'comprovo que si hi ha un albarà entrat no s'hagi de albaranar sol
                If datalinies.Recordset.RecordCount = 1 Then
                Set rst2 = dbcomandes.OpenRecordset("select * from comandes_extres where comanda=" + atrim(datalinies.Recordset!lotinplacsa))
                If Not rst2.EOF Then If cadbl(rst2!pararaexpedicions) > 1 Then MsgBox "La comanda ja entrada en aquest albarà s'ha d'albaranar sola no pots afegir cap mes albarà.", vbCritical, "Error": Exit Function
             End If
           End If
   End If
   Set rst = Nothing
   Set rst2 = Nothing
   comandashadalbaranarsola = False
End Function
Function comandaestaapuntperexpedir(vnumc As Double) As Boolean
   Dim rst As Recordset
   
   Set rst = dbcomandes.OpenRecordset("select * from comandes_extres where comanda=" + atrim(vnumc))
   If rst.EOF Then
       Exit Function
         Else
           'mirar les comprovacions que calgui
           'si està parada la expedició
           If cadbl(rst!pararaexpedicions) = 1 Then MsgBox "Aquesta comanda no es pot enviar, l'han parat desde oficina.", vbCritical, "No enviar": Exit Function
           If cadbl(rst!pararaexpedicions) = 2 Then MsgBox "Aquesta comanda s'ha d'albaranar sola.", vbCritical, "No agrupar albarans"
           If atrim(rst!refinplacsa) = "" Then MsgBox "ATENCIÓ AQUESTA COMANDA NO TÉ REFERENCIA D'INPLACSA NO ES POT ALBARANAR PERQUÈ DONARÀ ERROR DE SAP." + vbNewLine + "REVISEU LA COMANDA QUE TINGUI LA REFERENCIA D'INPLACSA.", vbCritical, "ERROR"
   End If
   Set rst = Nothing
   comandaestaapuntperexpedir = True
End Function
Sub crea_capcalera_i_linies(numc As Double, crearalbaranou As Boolean)
   Dim rstc As Recordset
   Dim rstdirenvio As Recordset
   Dim nounumalb As Double
   Dim vmoneda As String
   If datacapcalera.Recordset.EOF And Not crearalbaranou Then MsgBox "No pots afegir una comanda si no hi ha albarà creat, afegeig primer un albarà amb l'altra botó.", vbExclamation, "Atenció": numc = 0: Exit Sub
   'Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
   Set rstc = dbcomandes.OpenRecordset("SELECT comandes.*, comandes_extres.observacionsalbara, comandes_extres.codicomptable,comandes_extres.solpesgrmcm2,comandes_extres.transportista_albara FROM comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda where comandes.comanda=" + atrim(numc))
   
   If rstc.EOF Then Exit Sub
   Set rstdirenvio = dbcomandes.OpenRecordset("select * from clients_envios where id=" + atrim(rstc!direnvio))
   If rstdirenvio.EOF Then MsgBox "Error al carregar la direcció d'enviament d'aquest client.", vbCritical, "Atenció": GoTo fi
   idiomaclient = atrim(rstdirenvio!Idioma)
   If idiomaclient <> "CA" And idiomaclient <> "ES" And idiomaclient <> "EN" And idiomaclient <> "FR" Then
      idiomaclient = "EN"
      etmissatge = "Aquesta direcció d'enviament no te idioma assignat."
   End If
   If crearalbaranou Then
      vmoneda = buscarmonedadelclient(cadbl(rstc!client), cadbl(rstc!codicomptable))
      If vmoneda = "" Then MsgBox "No s'ha trobat el codicomptable d'aquest client reviseu les dades del client", vbCritical, "Error": GoTo fi
      nounumalb = proximnumalbara
      datacapcalera.Recordset.AddNew
      cnumalbara = nounumalb
      cdataalbara = Format(Now, "dd/mm/yy")
      datacapcalera.Recordset!id_direnvio = cadbl(rstc!direnvio)
      datacapcalera.Recordset!codiclient = cadbl(rstc!codicomptable)
      datacapcalera.Recordset!albaravalorat = cabool(rstdirenvio!albaravalorat)
      datacapcalera.Recordset!moneda = vmoneda
      datacapcalera.Recordset!tipusports = "Pagats"
      datacapcalera.Recordset!observacions = atrim(rstdirenvio!observacionsalbara)
      If cadbl(rstc!transportista_albara) > 0 Then datacapcalera.Recordset!id_transport = rstc!transportista_albara
      comboquifactura = atrim(rstdirenvio!empresa)
      datacapcalera.Recordset.Update
      datacapcalera.Recordset.FindFirst "numalbara=" + atrim(nounumalb)
      If datacapcalera.Recordset.NoMatch Then MsgBox "Error no s'ha trobat l'albarà nou Nº: " + atrim(nounumalb), vbCritical, "Error": GoTo fi
      possarliniesdepeualbarapredeterminadesdelclient
   End If
   
   If Not crearalbaranou Then
      If comprovarrepeticionscomandesdecrops Then MsgBox "Aquesta comandaclient es diferent que la resta de comandes entrades en aquest albarà de CROP'S", vbCritical, "Atenció": GoTo fi
      If cadbl(rstc!direnvio) <> cadbl(datacapcalera.Recordset!id_direnvio) And cadbl(rstc!codicomptable) <> cadbl(datacapcalera.Recordset!codiclient) Then
       MsgBox "Aquesta comanda no te el mateix CLIENT DE FACTURACIÓ ni la mateixa DIRECCIÓ d'enviament no es pot adjuntar al mateix albarà.", vbCritical, "Error": GoTo fi
         Else
           If cadbl(rstc!codicomptable) = cadbl(datacapcalera.Recordset!codiclient) Then
                If cadbl(rstc!direnvio) <> cadbl(datacapcalera.Recordset!id_direnvio) Then If MsgBox("Aquestes comandes tenen el mateix client de facturació però les direccions d'enviament tenen codis diferents. Vols albaranar-ho junt?", vbInformation + vbYesNo + vbDefaultButton2, "Direccions diferents") = vbNo Then GoTo fi
              Else: MsgBox "Aquesta comanda no te el mateix CLIENT DE FACTURACIÓ ni la mateixa DIRECCIÓ d'enviament no es pot adjuntar al mateix albarà.", vbCritical, "Error": GoTo fi
           End If
      End If
   End If
   If cadbl(rstc!transportista_albara) > 0 And datacapcalera.Recordset!id_transport > 0 Then
        If datacapcalera.Recordset!id_transport <> rstc!transportista_albara Then MsgBox "Aquesta comanda té un tranportista diferent assignat." + vbNewLine + "REVISA QUE ESTIGUI TOT CORRECTE SISPLAU.", vbCritical, "ERROR"
   End If
   datalinies.Recordset.AddNew
   datalinies.Recordset!numalbara = datacapcalera.Recordset!numalbara
   datalinies.Recordset!ordre = proximnumordre(datacapcalera.Recordset!numalbara)
   possar_els_campsdelalinia rstc, rstdirenvio
   mirarsihihabobinesambelcalloffdemanat datalinies.Recordset
   datalinies.Recordset.Update
   
   datalinies.Recordset.FindFirst "lotinplacsa=" + atrim(numc)
fi:
   Set rstc = Nothing
End Sub
Sub mirarsihihabobinesambelcalloffdemanat(rstlinia As Recordset)
   dbbaixes.Execute "update bobinesent set numalbara=" + atrim(rstlinia!numalbara) + " where (numalbara=null or numalbara=0) and comanda=" + atrim(cadbl(rstlinia!lotinplacsa)) + " and numcalloff='" + atrim(cadbl(rstlinia!numcalloff)) + "'"
End Sub
Function buscarmesura(idmesurapvp As Long, ByVal Idioma As String) As String
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select * from mesures where codi=" + atrim(idmesurapvp))
   If rst.EOF Then GoTo fi
   If atrim(Idioma) = "" Then Idioma = "ES"
   Idioma = "descripcio_" + Idioma
   buscarmesura = atrim(rst.Fields(Idioma))
fi:
   Set rst = Nothing
End Function
Function buscardescripcioproducte(codiproducte As String, ByVal Idioma As String) As String
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select * from productes where codi='" + atrim(codiproducte) + "'")
   If rst.EOF Then GoTo fi
   If atrim(Idioma) = "" Then Idioma = "ES"
   Idioma = "descpelclient_" + Idioma
   buscardescripcioproducte = atrim(rst.Fields(Idioma))
fi:
   Set rst = Nothing
End Function
Function triarelvalordepenguentdelaunitat() As Double
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda, mesures.unitatinterna, Clients_envios.packinglistalbara, Clients_envios.pesnetbrut,clients_envios.albaraarrodonirkg as arrodonirkg FROM (comandes INNER JOIN Clients_envios ON comandes.direnvio = Clients_envios.id) INNER JOIN mesures ON comandes.mesurapvp = mesures.codi WHERE (((comandes.comanda)=" + atrim(cadbl(datalinies.Recordset!lotinplacsa)) + "));")
   If rst.EOF Then Exit Function
   With datalinies.Recordset
   triarelvalordepenguentdelaunitat = 0
   Select Case rst!unitatinterna
     Case "€/1000U"
       triarelvalordepenguentdelaunitat = Redondejar(!unitats / 1000, 3)
     Case "€/U"
       triarelvalordepenguentdelaunitat = cadbl(!unitats)
     Case "€/B"
       triarelvalordepenguentdelaunitat = !numbobs
     Case "€/K"
       If Not rst!pesnetbrut Then
            triarelvalordepenguentdelaunitat = Redondejar(!kgtotalsbruts, 1)
             Else: triarelvalordepenguentdelaunitat = Redondejar(!kgtotalsnets, 1)
       End If
     Case "€/M"
       triarelvalordepenguentdelaunitat = !metreslineals
     Case "€/KM"
       triarelvalordepenguentdelaunitat = Redondejar(!metreslineals / 1000, 2)
     Case "€/FIX", "€/PROVA"
       triarelvalordepenguentdelaunitat = 1
     Case "€/M2"
       triarelvalordepenguentdelaunitat = Redondejar(!metreslineals * (!ampladamaterial / 1000), 2)
   End Select
   End With
   If rst!unitatinterna = "€/K" And rst!arrodonirkg Then triarelvalordepenguentdelaunitat = Redondejar(triarelvalordepenguentdelaunitat, 0)
End Function
Function mirarsimicroperforat(rstc As Recordset) As String
   mirarsimicroperforat = "N"
   vmicropsol = IIf(atrim(rstc!microperforatsol) = "", "N", atrim(rstc!microperforatsol))
   vmicrop = IIf(atrim(rstc!microperforat) = "", "N", atrim(rstc!microperforat))
   vmicropex = IIf(atrim(rstc!micropex) = "", "N", atrim(rstc!micropex))
   If vmicrop <> "N" Then mirarsimicroperforat = "S"
   If vmicropex <> "N" Then mirarsimicroperforat = "S"
   If vmicropsol <> "N" Then mirarsimicroperforat = "S"
End Function
Function buscarcolormaterial(rstc As Recordset, vruta As String) As String
   Dim vcolor1 As String
   Dim vcolor2 As String
   Dim vcolor3 As String
   Dim rst As Recordset
   Dim vsql As String
   vsql = "SELECT comandes.comanda, familiescolorants.descripcio as color1, familiescolorants_1.descripcio as color2, familiescolorants_2.descripcio as color3 "
   vsql = vsql + " FROM (((((((comandes LEFT JOIN materials ON comandes.materialex = materials.codi) LEFT JOIN comandes AS comandes_1 ON comandes.linkcomanda1 = comandes_1.comanda) LEFT JOIN comandes AS comandes_2 ON comandes.linkcomanda2 = comandes_2.comanda) LEFT JOIN materials AS materials_1 ON comandes_1.materialex = materials_1.codi) LEFT JOIN materials AS materials_2 ON comandes_2.materialex = materials_2.codi) LEFT JOIN familiescolorants ON materials.familiacol = familiescolorants.codi) LEFT JOIN familiescolorants AS familiescolorants_1 ON materials_1.familiacol = familiescolorants_1.codi) LEFT JOIN familiescolorants AS familiescolorants_2 ON materials_2.familiacol = familiescolorants_2.codi "
   vsql = vsql + " WHERE (((comandes.comanda)=" + atrim(rstc!comanda) + "));"
   Set rst = dbcomandes.OpenRecordset(vsql)
   If Not rst.EOF Then vcolor1 = atrim(rst!color1): vcolor2 = atrim(rst!color2): vcolor3 = atrim(rst!color3)
   If InStr(1, vruta, "I") = 0 Then
      If InStr(1, vcolor1, "TRANSP") = 0 And vcolor1 <> "" Then buscarcolormaterial = vcolor1
      If InStr(1, vcolor2, "TRANSP") = 0 And vcolor2 <> "" Then buscarcolormaterial = vcolor2
      If InStr(1, vcolor3, "TRANSP") = 0 And vcolor3 <> "" Then buscarcolormaterial = vcolor3
 
      If InStr(1, vcolor1 + vcolor2 + vcolor3, "METAL") > 0 Then buscarcolormaterial = "METALIZADO"
      If buscarcolormaterial = "" Then buscarcolormaterial = "TRANSPARENTE"
   End If
   Set rst = Nothing
End Function
Sub possar_els_campsdelalinia(rstc As Recordset, rstdirenvio As Recordset)
   Dim rstcextra As Recordset
   Dim vsemionosemi As String
   Dim vruta As String
   vruta = rutadelproducte(rstc!comanda)
   If vruta = "E" Or vruta = "EI" Then
        If atrim(rstc!oberturaex) = "1" Then vsemionosemi = "SEMI "
        If atrim(rstc!oberturaex) = "2" Then vsemionosemi = "2s SEMI "
   End If
   Set rstcextra = dbcomandes.OpenRecordset("select * from comandes_extres where comanda=" + atrim(rstc!comanda))
   With datalinies.Recordset
   !codiproducte = atrim(rstcextra!refinplacsa)
   !unitatmesura = buscarmesura(cadbl(rstc!mesurapvp), idiomaclient)
   !unitatpvp = rstc!mesurapvp
   !descripcioproducte = vsemionosemi + buscardescripcioproducte(rstc!producte, idiomaclient)
   !colormaterial = Mid(buscarcolormaterial(rstc, vruta), 1, 15)
   !preuvenda = IIf(datacapcalera.Recordset!moneda = "Euros", cadbl(rstc!pvp), cadbl(rstc!pvpdolar))
   If !preuvenda = -1 Then !preuvenda = 0
   !microperforat = mirarsimicroperforat(rstc)
   !macroperforat = atrim(rstc!rebmacroperforat)
   possar_descripciomaterial_liniaitintes rstc, rstdirenvio
   possar_referenciesdelclient rstc
   possar_datafab_contractesicalloff rstc
   possar_kilos_metres_unitats_etc rstc
   !quantitat = triarelvalordepenguentdelaunitat  'aixó al final que ja tindré tots els camps
   !observacionslinia = Mid(atrim(rstc!observacionsalbara), 1, 100)
   !kgimpostenvasos = calcularkgimpost(rstc)
   !eurokg_impost = cadbl(llegir_ini("General", "PreuImpostEnvasos", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"))
fi:
   End With
   Set rstcextra = Nothing
End Sub
Function calcularkgimpost(rstc As Recordset) As Double

End Function
Function rutadelproducte(numc As Double) As String
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda, productes.ruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(numc) + "));")
   If rst.EOF Then GoTo fi
   rutadelproducte = atrim(rst!ruta)
fi:
   Set rst = Nothing
End Function
Sub possar_kilos_metres_unitats_etc(rstc As Recordset)
   Dim rst As Recordset
   Dim vruta As String
   Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda, mesures.unitatinterna, Clients_envios.packinglistalbara, Clients_envios.pesnetbrut FROM (comandes INNER JOIN Clients_envios ON comandes.direnvio = Clients_envios.id) INNER JOIN mesures ON comandes.mesurapvp = mesures.codi WHERE (((comandes.comanda)=" + atrim(cadbl(rstc!comanda)) + "));")
   If Not rst.EOF Then datalinies.Recordset!espesnet = cabool(rst!pesnetbrut)
   
   vruta = rutadelproducte(cadbl(rstc!comanda))
   If InStr(1, vruta, "S") = 0 Then
        possar_quantitats_bobines rstc
         Else: possar_quantitats_formats rstc
   End If
   Set rst = Nothing
End Sub
Sub possar_quantitats_formats(rstc As Recordset)
   Dim rstmesura As Recordset
   Dim rst As Recordset
   Dim vmesura As String
   Dim vmetres As Double
   Dim vkilosb As Double
   Dim vkilosn As Double
   Dim vbobs As Double
   Dim vdesarroll As Double
   Dim vunitats As Double
   Dim rstsol As Recordset
   Dim rstbobinesreb As Recordset
   Dim vtipusembalatge As String
   Dim vpalets As Integer
   'busco les bobines seleccionades
   Set rstsol = dbbaixes.OpenRecordset("select sacsocaixes from soldadorestot where comanda=" + atrim(rstc!comanda))
   If rstsol.EOF Then MsgBox "No trobo la secció de baixes de soldadores", vbCritical, "Atenció": Exit Sub
   If atrim(rstsol!sacsocaixes) = "" Then
       MsgBox "No està especificat a la baixa de soladores si son sacs o caixes, possaré caixes.", vbCritical, "Error": vtipusembalatge = "Caixes"
          Else: vtipusembalatge = atrim(rstsol!sacsocaixes)
   End If
   Set rstsol = Nothing
  ' Set rstbobinesreb = dbbaixes.OpenRecordset("SELECT bobinessol.numerodesac,soldadores.comanda, bobinessol.unitatsxsac FROM soldadores RIGHT JOIN bobinessol ON soldadores.Id = bobinessol.controlid where comanda=" + atrim(rstc!comanda))
   Set rst = dbbaixes.OpenRecordset("select * from bobinesent where comanda=" + atrim(rstc!comanda) + " and (numalbara=" + cnumalbara + " ) order by numbob")  'or numalbara='' or numalbara=null
   While Not rst.EOF
     ' rstbobinesreb.FindFirst "numerodesac=" + atrim(rst!numbob)
     ' If Not rstbobinesreb.NoMatch Then
      vunitats = vunitats + cadbl(rst!metresisacs)
     ' End If
      rst.MoveNext
   Wend
   'compto els palets que hi ha
   vpalets = 0
   Set rst = dbbaixes.OpenRecordset("select * from bobinesent where comanda=" + atrim(rstc!comanda) + " and (numalbara=" + cnumalbara + " ) ")
   If Not rst.EOF Then
     rst.MoveLast: rst.MoveFirst
     vbobs = rst.RecordCount
   End If
   
   Set rst = dbbaixes.OpenRecordset("select distinct numpalet from bobinesent where comanda=" + atrim(rstc!comanda) + " and (numalbara=" + cnumalbara + " ) ")
   vpalets = 0
   If Not rst.EOF Then
     rst.MoveLast: rst.MoveFirst
     vpalets = cadbl(rst.RecordCount)
   End If
     
   With datalinies.Recordset
   !numbobs = vbobs
   !kgtotalsbruts = Redondejar(calcularpesxrpeça(rstc) * vunitats, 1)
   !kgtotalsnets = 0
   !pespalets = buscarpesdelspalets(rstc!comanda, cabool(!espesnet))
   !metreslineals = IIf(InStr(1, !unitatmesura, "€/MT.") > 0, vunitats, 0) 'En ramon em va dir que si facturen per metres possi les unitats com a metres
   !unitats = IIf(InStr(1, !unitatmesura, "€/MT.") = 0, vunitats, 0)
   !tipusproducte = IIf(InStr(1, !unitatmesura, "€/MT.") > 0, "AvBobines", vtipusembalatge)
   !numerodepalets = vpalets
   !lotinplacsa = rstc!comanda
   End With
   Set rst = Nothing
   Set rstmesura = Nothing
End Sub

Function buscarpesdelspalets(numc As Double, espesnet As Boolean) As Double
   Dim rst As Recordset
   Set rst = dbbaixes.OpenRecordset("select sum(pespalet) as tpespalets from reb_pespalets where comanda=" + atrim(numc) + " group by comanda")
   If Not rst.EOF Then buscarpesdelspalets = cadbl(rst!tpespalets)
   If buscarpesdelspalets < 1 And espesnet Then
       buscarpesdelspalets = cadbl(InputBox("No he trobat el pes dels palets per aquesta comanda i es le client vol PES NET." + Chr(10) + "SI VOLS POSSAR ELS KILOS DELS PALETS I ES SUMARÀ AL TOTAL DE KG BRUTS", "Falta pes de palets"))
   End If
End Function
Sub possar_quantitats_bobines(rstc As Recordset)
   Dim rstmesura As Recordset
   Dim rst As Recordset
   Dim vmesura As String
   Dim vmetres As Double
   Dim vkilosb As Double
   Dim vkilosn As Double
   Dim vcontinuu As Boolean
   Dim vbobs As Double
   Dim vdesarroll As Double
   Dim vunitats
   Dim vpalets
   Dim rstbobinesreb As Recordset
   Set rstmesura = dbcomandes.OpenRecordset("select unitatinterna from mesures where codi=" + atrim(cadbl(rstc!mesurapvp)))
   If rstmesura.EOF Then MsgBox "Aquesta comanda no te mesura de PVP assignada.", vbCritical, "Atenció": Exit Sub
   vmesura = atrim(rstmesura!unitatinterna): Set rstmesura = Nothing
   Set rstbobinesreb = dbbaixes.OpenRecordset("SELECT bobinesreb.numerodebobina,rebobinadores.comanda, bobinesreb.kilos, bobinesreb.pesnet FROM rebobinadores RIGHT JOIN bobinesreb ON rebobinadores.Id = bobinesreb.controlid where comanda=" + atrim(rstc!comanda))
   'busco les bobines seleccionades
   Set rst = dbbaixes.OpenRecordset("select * from bobinesent where comanda=" + atrim(rstc!comanda) + " and (numalbara=" + cnumalbara + ") order by numbob")
'   vpalets = 0
   While Not rst.EOF
      rstbobinesreb.FindFirst "numerodebobina=" + atrim(rst!numbob)
      vmetres = vmetres + cadbl(rst!metresisacs)
      If Not rstbobinesreb.NoMatch Then
        vkilosb = vkilosb + Redondejar(cadbl(rstbobinesreb!kilos), 1)
        vkilosn = vkilosn + cadbl(rstbobinesreb!pesnet)
          Else: vkilosb = vkilosb + Redondejar(cadbl(rst!kilosiunitats), 1)
      End If
 '     If cadbl(rst!numpalet) > vpalets Then vpalets = cadbl(rst!numpalet)
      rst.MoveNext
   Wend
   vpalets = 0
   Set rst = dbbaixes.OpenRecordset("select distinct numpalet from bobinesent where comanda=" + atrim(rstc!comanda) + " and (numalbara=" + cnumalbara + " ) ")
   If Not rst.EOF Then
     rst.MoveLast: rst.MoveFirst
     vpalets = cadbl(rst.RecordCount)
   End If
   Set rst = dbbaixes.OpenRecordset("select numpalet from bobinesent where comanda=" + atrim(rstc!comanda) + " and (numalbara=" + cnumalbara + " ) ")
   If Not rst.EOF Then
     rst.MoveLast: rst.MoveFirst
     vbobs = cadbl(rst.RecordCount)
   End If
  ' vbobs = rst.RecordCount
   vdesarroll = mirardesarrolldeltreball(cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), vcontinuu)
   If vdesarroll <> 0 Then vunitats = Redondejar(vmetres / vdesarroll, 0)
   If vcontinuu Then vunitats = 0
   With datalinies.Recordset
   !numbobs = vbobs
   !pespalets = buscarpesdelspalets(rstc!comanda, cabool(!espesnet))
   !numerodepalets = vpalets
   !kgtotalsbruts = Redondejar(vkilosb, 1)
   !kgtotalsnets = Redondejar(vkilosn, 1)
   !metreslineals = vmetres
   !unitats = vunitats
   !lotinplacsa = rstc!comanda
   End With
End Sub
Function mirardesarrolldeltreball(ntreball As Double, nordre As Double, vcontinuu As Boolean) As Double
   Dim rstclixes As Recordset
   Dim rsttintes As Recordset
   Dim vdesarrollcalculat As Double
   Set rstclixes = dbclixes.OpenRecordset("select * from modificacions where id_Treball=" + atrim(ntreball) + " and ordre=" + atrim(nordre))
   If rstclixes.EOF Then Exit Function
     '16/12/21 modificada aquesta linia per un error en el calcul del desarroll en el cas que un color hagi canviat
       'aixó feia que l'albarà estigui equivocat calculant les peces malament
   '''Set rsttintes = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(rstclixes!id_treball) + " and ordremodificacio=" + atrim(rstclixes!ordre) + "and cilindre<>0 order by continuu DESC")
   Set rsttintes = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(rstclixes!id_treball) + " and ordremodificacio=" + atrim(rstclixes!ordre) + "and color<>'' and cilindre<>0 order by continuu DESC")
   If Not rsttintes.EOF Then
     If cadbl(rstclixes!desarroll) > 0 Then
        vdesarrollcalculat = cadbl(rsttintes!cilindre) / Redondejar(cadbl(rsttintes!cilindre) / cadbl(rstclixes!desarroll), 0)
     End If
   End If
   If vdesarrollcalculat = 0 Then vdesarrollcalculat = cadbl(rstclixes!desarroll)
   mirardesarrolldeltreball = cadbl(vdesarrollcalculat) / 1000
   If Not rstclixes.EOF Then vcontinuu = cabool(rsttintes!continuu)
   Set rstclixes = Nothing
   Set rsttintes = Nothing
End Function
Sub possar_datafab_contractesicalloff(rstc As Recordset)
   Dim rst As Recordset
   Dim vcontracte As String
   Dim vcalloff As String
   Set rst = dbbaixes.OpenRecordset("select * from " + taulaultimaseccio(rstc!producte) + " where comanda=" + atrim(rstc!comanda))
   If Not rst.EOF Then
     If IsDate(rst!datainici) Then
       datalinies.Recordset!datafabricacio = Format(rst!datainici, "dd/mm/yy")
         Else: datalinies.Recordset!datafabricacio = Date
     End If
   End If
   While Len(vcontracte) > 30
    vcontracte = InputBox("Entra el numero de Contracte si el tens.", "Nº contracte")
   Wend
   vcalloff = buscarsihihacalloff(rstc!comanda)
   vcalloff = InputBox("Entra el número de CallOff si el tens." + vbNewLine + "Màxim 30 digits", "Nº CallOff", vcalloff)
   While Len(vcalloff) > 30
        vcalloff = InputBox("Entra el número de CallOff si el tens.", "Nº CallOff", vcalloff)
   Wend
   datalinies.Recordset!numcalloff = atrim(vcalloff)
   datalinies.Recordset!numcontracte = atrim(vcontracte)
   Set rst = Nothing
End Sub
Function buscarsihihacalloff(vnumc As Double) As String
  Dim rst As Recordset
  Set rst = dbbaixes.OpenRecordset("select numcalloff from bobinesent where (numalbara=null or numalbara=0) and comanda=" + atrim(vnumc) + " and (numcalloff<>'' and numcalloff<>null)")
  If Not rst.EOF Then
    buscarsihihacalloff = atrim(rst!numcalloff)
    dbcomandes.Execute "delete * from calloffs_detall where comanda=" + atrim(vnumc)
      Else
        Set rst = dbcomandes.OpenRecordset("select * from calloffs_detall where comanda=" + atrim(vnumc))
        If Not rst.EOF Then buscarsihihacalloff = atrim(rst!numcalloff)
  End If
  If buscarsihihacalloff = "" Then
     Set rst = dbcomandes.OpenRecordset("select numcalloff from comandes_extres where comanda=" + atrim(vnumc))
     If Not rst.EOF Then buscarsihihacalloff = atrim(rst!numcalloff)
  End If
  Set rst = Nothing
End Function
Function taulaultimaseccio(vproducte As String) As String
   Dim rst As Recordset
   Dim vruta As String
   Set rst = dbcomandes.OpenRecordset("select ruta from productes where codi='" + atrim(vproducte) + "'")
   taulaultimaseccio = "impressores"
   If Not rst.EOF Then
      vruta = atrim(rst!ruta) + " "
      Select Case Mid(vruta, Len(vruta) - 1, 1)
         Case "E"
           taulaultimaseccio = "extrussores"
         Case "I"
           taulaultimaseccio = "impressores"
         Case "L"
           taulaultimaseccio = "laminadores"
         Case "R"
           taulaultimaseccio = "rebobinadores"
         Case "S"
           taulaultimaseccio = "soldadores"
      End Select
   End If
   Set rst = Nothing
End Function

Sub possar_referenciesdelclient(rstc As Recordset)
   Dim vref1 As String
   Dim vcom1 As String
   Dim vref2 As String
   Dim vcom2 As String
   Dim vresposta
   With rstc
    If (atrim(!comandaclient) <> "" Or atrim(!refclient) <> "") And (atrim(!obspedgen2) <> "" Or atrim(!refclientdeclient) <> "") Then
       vresposta = UCase(InputBox("Hi ha valors a comanda de client i a Comanda client de client quina vols possar?" + Chr(10) + " Si vols client (C), client de client (CdC), o totes (T)", "Escullir", "C"))
       If vresposta = "CdC" Then
           vcom2 = atrim(!obspengen2): vref2 = atrim(refclientdeclient)
               Else
                  If vresposta = "C" Then
                     vcom1 = atrim(!comandaclient): vref1 = atrim(!refclient)
                      Else: GoTo tots
                  End If
       End If
        Else:
tots:
           vcom1 = atrim(!comandaclient): vref1 = atrim(!refclient)
           vcom2 = atrim(!obspedgen2): vref2 = atrim(refclientdeclient)
    End If
   End With
   datalinies.Recordset!numcomandacli = vcom1
   datalinies.Recordset!refclient = Mid(vref1, 1, 30)
   vcom2 = " " + vcom2
   datalinies.Recordset!numcomandaclideclient = Trim(Mid(vcom2, 1, 30))
   datalinies.Recordset!refclientdeclient = vref2
End Sub
Sub possarlespesordelmaterial(rstc As Recordset)
  Dim rst As Recordset
  Dim vespesor As Double
  Dim vmesura As String
  Set rst = dbcomandes.OpenRecordset("select comanda,mesuraesp,espessor,tubolam from comandes where comanda=" + atrim(rstc!comanda) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2))
  While Not rst.EOF
    If rst!comanda > 0 Then
      vespesor = vespesor + micresmaterial(cadbl(rst!mesuraesp), rst!espessor, rst!tubolam)
      vmesura = r
      If InStr(1, r, "GALGUES") Then vmesura = "MICRES"
      If vmesura = "GR/MT2" Then vmesura = "GRMS/M2"
    End If
    rst.MoveNext
  Wend
  If vespesor < 0 Then vmesura = "GRMS/M2"
  datalinies.Recordset!espesor = vespesor
  datalinies.Recordset!mesuraespesor = vmesura
End Sub
Function micresmaterial(codimesuralineal As Byte, espesor As Double, tubolam As String) As Double
  Dim rstmesural As Recordset
  Set rstmesural = dbcomandes.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(codimesuralineal))
  r = ""
  If rstmesural.EOF Then Exit Function
  r = espesor
  If rstmesural!descripcio = "GALGUES" Then
            If tubolam = "T" Then
                 r = formatar(espesor / 4, "#,##0")
                  Else: r = formatar(espesor / 2, "#,##0")
            End If
  End If
  If InStr(1, rstmesural!descripcio, "GR/") > 0 Then
    r = espesor * -1
  End If
  micresmaterial = r
  r = rstmesural!descripcio
End Function
Function possarliniaitintes(numtreball As Double, numordre As Double) As Byte
  Dim rstclixes As Recordset
  If numordre = 0 Then numordre = 1
  Set rstclixes = dbclixes.OpenRecordset("SELECT Clixes.id_treball,clixes.codidebarres, Modificacions.ordre, Clixes.marca, Clixes.linia, Modificacions.tinters FROM Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball where modificacions.id_treball=" + atrim(numtreball) + " and ordre=" + atrim(numordre))
  If rstclixes.EOF Then GoTo fi
   datalinies.Recordset!marcailinia = atrim(rstclixes!marca) + " - " + atrim(rstclixes!linia)
  datalinies.Recordset!codibarres = atrim(rstclixes!codidebarres)
  possarliniaitintes = cadbl(rstclixes!tinters)
fi:
  Set rstclixes = Nothing
End Function
Function generaelcoditipusINP(rstc As Recordset) As String
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select materialex from comandes where comanda=" + atrim(rstc!comanda) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2))
   While Not rst.EOF
     If cadbl(rst!materialex) > 0 Then generaelcoditipusINP = generaelcoditipusINP + atrim(rst!materialex)
     rst.MoveNext
   Wend
   If generaelcoditipusINP <> "" Then generaelcoditipusINP = "INP" + generaelcoditipusINP
   Set rst = Nothing
End Function
Function possar_descripciomaterial_liniaitintes(rstc As Recordset, rstdirenvio As Recordset) As String
   Dim rstp As Recordset
   Dim esmaterialimpres As Boolean
   Dim coditipusmaterial As String
   Dim vtintes As Byte
   Dim vruta As String
   Set rstp = dbcomandes.OpenRecordset("select * from productes where codi='" + atrim(rstc!producte) + "'")
   If rstp.EOF Then GoTo fi
   esmaterialimpres = IIf(InStr(1, rstp!ruta, "I"), True, False)
   With datalinies.Recordset
   coditipusmaterial = generaelcoditipusINP(rstc) 'falta fer aixó encara
   possarlespesordelmaterial rstc
   If esmaterialimpres Then vtintes = possarliniaitintes(rstc!numtreball, rstc!numordremodificacio)
   vruta = atrim(rstp!ruta) + " "
   If InStr(1, vruta, "S") > 0 Then
        'formats
         !ampladamaterial = cadbl(rstc!amplesol) * 10
         !descripciomides = atrim(cadbl(!ampladamaterial)) + IIf(cadbl(rstc!ampleplegsol) > 0, "/" + atrim(cadbl(rstc!ampleplegsol) * 10), "") + "X"
         !descripciomides = !descripciomides + atrim(cadbl(rstc!longitudsol) * 10) + IIf(cadbl(rstc!solapasol) > 0, "+" + atrim(cadbl(rstc!solapasol) * 10), "") + " mm"
         If cadbl(rstc!fuellebasesol) > 0 Or cadbl(rstc!fuellebocasol) > 0 Then !descripciomides = !descripciomides + " F(" + atrim(cadbl(rstc!fuellebasesol) * 10) + "+" + atrim(cadbl(rstc!fuellebocasol) * 10) + ") "
        Else
           'Bobines
           If Mid(vruta, Len(vruta) - 1, 1) = "S" Then !ampladamaterial = cadbl(rstc!amplesol) * 10
           If Mid(vruta, Len(vruta) - 1, 1) = "R" Then !ampladamaterial = cadbl(rstc!amplereb) * 10
           If Mid(vruta, Len(vruta) - 1, 1) = "L" Then !ampladamaterial = cadbl(rstc!ampleutil) * 10
           If Mid(vruta, Len(vruta) - 1, 1) = "I" Then !ampladamaterial = cadbl(rstc!ampleesq) * 10
           If Mid(vruta, Len(vruta) - 1, 1) = "E" Then !ampladamaterial = cadbl(rstc!ampleesq) * 10
           !descripciomides = atrim(cadbl(!ampladamaterial)) + " mm "
           
   End If
   
   !descripciomides = !descripciomides + " " + atrim(IIf(!espesor < 0, "", atrim(!espesor) + "µ "))
   If rstc!linkcomanda1 > 0 Or rstc!linkcomanda2 > 0 Then
       !descripciomides = !descripciomides + " Cod:" + coditipusmaterial
         Else: !descripciomides = !descripciomides + buscarprimeraparaulanommaterial(rstc!materialex)
   End If
   If esmaterialimpres Then !descripciomides = !descripciomides + "  " + traducciodeabreviatures("AvTintersTitol", idiomaclient) + "  " + atrim(vtintes) + traducciodeabreviatures("AvTinters", idiomaclient)

fi:
  End With
   Set rstp = Nothing
End Function
Function traducciodeabreviatures(av As String, Idioma As String)

   If Idioma = "CA" Then
     Select Case av
      Case "AvClixes"
       traducciodeabreviatures = "FOTOCOMPOSICIÓ I FOTOPOLÍMERS"
      Case "Avmacrop"
       traducciodeabreviatures = "Macro-P"
      Case "Avmicrop"
       traducciodeabreviatures = "Micro-P"
      Case "Avpesbrut"
       traducciodeabreviatures = "Brut"
      Case "COM:"
       traducciodeabreviatures = "COM."
      Case "Ports"
       traducciodeabreviatures = "Ports"
      Case "Pagats"
       traducciodeabreviatures = "Pagats"
      Case "Deguts"
       traducciodeabreviatures = "Deguts"
      Case "Facturats"
       traducciodeabreviatures = "En Factura"
      Case "AvTintersTitol"
       traducciodeabreviatures = "Imp."
      Case "AvTinters"
       traducciodeabreviatures = "C"
      Case "AvDataproduccio"
       traducciodeabreviatures = "Data/prod."
      Case "AvLot"
       traducciodeabreviatures = "Lot"
      Case "AvPeces"
       traducciodeabreviatures = "Pcs"
      Case "AVBOBINES"
       traducciodeabreviatures = "Bobines"
      Case "AvBobines"
       traducciodeabreviatures = "Bobines"
      Case "SACS"
       traducciodeabreviatures = "Sacs"
      Case "CAIXES"
       traducciodeabreviatures = "Caixes"
     End Select
   End If

   If Idioma = "ES" Then
    Select Case av
      Case "AvClixes"
       traducciodeabreviatures = "FOTOCOMPOSICION Y FOTOPOLIMEROS"
      Case "Avmacrop"
       traducciodeabreviatures = "Macro-P"
      Case "Avmicrop"
       traducciodeabreviatures = "Micro-P"
      Case "Avpesbrut"
       traducciodeabreviatures = "Bruto"
      Case "COM:"
       traducciodeabreviatures = "PED:"
      Case "Ports"
       traducciodeabreviatures = "Portes:"
      Case "Pagats"
       traducciodeabreviatures = "Pagados"
      Case "Deguts"
       traducciodeabreviatures = "Debidos"
      Case "Facturats"
       traducciodeabreviatures = "En Factura"
      Case "AvTintersTitol"
       traducciodeabreviatures = "Imp."
      Case "AvTinters"
       traducciodeabreviatures = "C"
      Case "AvDataproduccio"
       traducciodeabreviatures = "Fecha/prod."
      Case "AvLot"
       traducciodeabreviatures = "Lote"
      Case "AvPeces"
       traducciodeabreviatures = "Pzs"
      Case "AVBOBINES"
       traducciodeabreviatures = "Bobinas"
      Case "AvBobines"
       traducciodeabreviatures = "Bobinas"
      Case "SACS"
       traducciodeabreviatures = "Sacos"
      Case "CAIXES"
       traducciodeabreviatures = "Cajas"
      
   End Select
   End If
   If Idioma = "EN" Then
    Select Case av
      Case "AvClixes"
       traducciodeabreviatures = "TYPESETTING & PHOTOPOLYMERS"
      Case "Avmacrop"
       traducciodeabreviatures = "Macro-P"
      Case "Avmicrop"
       traducciodeabreviatures = "Micro-P"
      Case "Avpesbrut"
       traducciodeabreviatures = "Gross"
      Case "COM:"
       traducciodeabreviatures = "Order:"
      Case "Ports"
       traducciodeabreviatures = "Carriage:"
      Case "Pagats"
       traducciodeabreviatures = "paid"
      Case "Deguts"
       traducciodeabreviatures = "forward"
      Case "Facturats"
       traducciodeabreviatures = "Extra cost"
      Case "AvTintersTitol"
       traducciodeabreviatures = "Print."
      Case "AvTinters"
       traducciodeabreviatures = "C"
      Case "AvDataproduccio"
       traducciodeabreviatures = "Prod.date"
      Case "AvLot"
       traducciodeabreviatures = "Batch"
      Case "AvPeces"
       traducciodeabreviatures = "Pcs"
       Case "AVBOBINES"
       traducciodeabreviatures = "Reels"
      Case "AvBobines"
       traducciodeabreviatures = "Reels"
      Case "SACS"
       traducciodeabreviatures = "Bags"
      Case "CAIXES"
       traducciodeabreviatures = "Boxes"
    End Select
   End If
   If Idioma = "FR" Then
    Select Case av
      Case "AvClixes"
       traducciodeabreviatures = "PHOTOCOMPOSITION ET PHOTOPOLYMÈRES"
      Case "Avmacrop"
       traducciodeabreviatures = "Macro-P"
      Case "Avmicrop"
       traducciodeabreviatures = "Micro-P"
      Case "Avpesbrut"
       traducciodeabreviatures = "Brut"
      Case "COM:"
       traducciodeabreviatures = "Comm:"
      Case "Ports"
       traducciodeabreviatures = "Port:"
      Case "Pagats"
       traducciodeabreviatures = "payé"
      Case "Deguts"
       traducciodeabreviatures = "dus"
      Case "Facturats"
       traducciodeabreviatures = "Dépenses supplémentaires"
      Case "AvTintersTitol"
       traducciodeabreviatures = "Print."
      Case "AvTinters"
       traducciodeabreviatures = "C"
      Case "AvDataproduccio"
       traducciodeabreviatures = "Date/prod."
      Case "AvLot"
       traducciodeabreviatures = "Lot"
      Case "AvPeces"
       traducciodeabreviatures = "Pcs"
       Case "AVBOBINES"
       traducciodeabreviatures = "Rolls"
      Case "AvBobines"
       traducciodeabreviatures = "Rolls"
      Case "SACS"
       traducciodeabreviatures = "Sacs"
      Case "CAIXES"
       traducciodeabreviatures = "Boîtes"
    End Select
   End If
End Function
Function buscarprimeraparaulanommaterial(codimaterial As Double)
    Dim rstm As Recordset
    Dim vdesc As String
    Set rstm = dbcomandes.OpenRecordset("select descripcio from materials where codi=" + atrim(codimaterial))
    If rstm.EOF Then GoTo fi
    vdesc = atrim(rstm!descripcio) + "   "
    buscarprimeraparaulanommaterial = atrim(Mid(vdesc, 1, InStr(1, vdesc, " ")))
fi:
    Set rstm = Nothing
End Function
Function proximnumordre(numc As Double) As Double
  Dim rst As Recordset
  Set rst = datacapcalera.Database.OpenRecordset("select ordre from liniesalbara where numalbara=" + atrim(numc) + " order by ordre desc")
  If rst.EOF Then
     proximnumordre = 10
      Else: proximnumordre = rst!ordre + 10
  End If
  Set rst = Nothing
End Function
Function proximnumalbara() As Double
   Dim rst As Recordset
   Set rst = datacapcalera.Database.OpenRecordset("select numalbara from capcaleraalbara order by numalbara desc")
   If rst.EOF Then
         proximnumalbara = 1
       Else: proximnumalbara = cadbl(rst!numalbara) + 1
   End If
   Set rst = Nothing
End Function
Sub carregar_bobinesentrega(numc As Double, vnumalbara As Double)
   Dim rst As Recordset
   Dim ruta As String
   Dim ventregat As String
   Static jahisoc As Boolean
   If jahisoc Then Exit Sub
   seltot.tag = "1"
   ruta = rutadelproducte(numc)
   If ruta = "" Then MsgBox "No trobo la ruta d'aquesta comanda o la comanda.", vbCritical, "Error": Exit Sub
   jahisoc = True
   Set rst = dbbaixes.OpenRecordset("select * from bobinesent where comanda=" + atrim(numc) + " and (numalbara=" + atrim(cadbl(vnumalbara)) + " or numalbara=null or numalbara=0) order by numbob")
   llistabobinessel.Clear
   If Not rst.EOF Then If Not IsNull(rst!dataentrega) And atrim(rst!dataentrega) <> "0:00:00" Then ventregat = Format(rst!dataentrega, "dd/mm/yy")
   While Not rst.EOF
     If Mid(ruta, Len(ruta)) <> "S" Then
        llistabobinessel.AddItem justificar(rst!numbob, 3, "D") + "B" + justificar(rst!metresisacs, 6, "D") + "M" + justificar(Redondejar(cadbl(rst!kilosiunitats), 1), 6, "D") + "K"
          Else: llistabobinessel.AddItem justificar(rst!numbob, 3, "D") + "S" + justificar(rst!metresisacs, 6, "D") + " Unitats" '+ justificar(cadbl(rst!kilosiunitats), 6, "D") + "K"
     End If
     llistabobinessel.ItemData(llistabobinessel.NewIndex) = cadbl(rst!numbob)
     If cadbl(rst!numalbara) = vnumalbara Then llistabobinessel.Selected(llistabobinessel.NewIndex) = True
     rst.MoveNext
   Wend
   llistabobinessel.tag = atrim(numc)
   jahisoc = False
   seltot.tag = ""
   If ventregat <> "" Then
        benviat.visible = False
        etdataenviament = "Entregat: " + ventregat
        datacapcalera.Database.Execute "update linies_expedicions set enviat=true where albara=" + atrim(cadbl(cnumalbara))
           Else:
             etdataenviament = "": benviat.visible = True
                 If cadbl(cnumalbara) > 0 Then datacapcalera.Database.Execute "update linies_expedicions set enviat=false where albara=" + atrim(cadbl(cnumalbara))
   End If
        
End Sub
Function justificar(v As String, longitut As Integer, DoE As String) As String
    v = Mid(v, 1, longitut)
    If DoE = "E" Then
       v = v + Space(longitut - Len(v))
      Else: v = Space(longitut - Len(v)) + v
    End If
    justificar = v
End Function

Private Sub bassignartransport_Click()
  MsgBox "Per assignar un transport s'ha de fer desde el programa d'assignació de transports al programa de comandes.", vbCritical, "Atenció"
End Sub

Private Sub bcanvienvio_Click()
   triar_client_direnvio
End Sub
Function codidelclientdelaidenvio() As Double
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select * from clients_envios where id=" + atrim(datacapcalera.Recordset!id_direnvio))
   If Not rst.EOF Then codidelclientdelaidenvio = rst!codi
   Set rst = Nothing
End Function
Sub triar_client_direnvio()
  
   Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select id ,nome,domicilie,poblacioe,provinciae from clients_envios where codi=" + atrim(codidelclientdelaidenvio)
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(2).width = 900
  formseleccio.width = 9000
  formseleccio.Left = formseleccio.Left - 3000
  If formseleccio.Data1.Recordset.EOF Then MsgBox "Aquest client no te cap DIRECCIO D'ENVIO ASSIGNADA.": Exit Sub
  formseleccio.Data1.Recordset.MoveLast
  formseleccio.Data1.Recordset.MoveFirst
  If formseleccio.Data1.Recordset.RecordCount > 1 Then
     formseleccio.Show 1
    Else: MsgBox "Nomes hi ha una direcció d'enviament no es pot canviar", vbCritical, "Error"
  End If
  If seleccioret = 1 Then
        
        If Not formseleccio.Data1.Recordset.EOF Then
           datacapcalera.Recordset.id_direnvio = cadbl(formseleccio.DBGrid2.Columns("id"))
        End If
   End If
   Unload formseleccio
   SendKeys "{TAB}"
   Command6_Click
   'codimuntadora.SetFocus
End Sub

Function triar_direnvio_client_busqueda()
  
   Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select id ,nome,domicilie,poblacioe,provinciae from clients_envios order by nome"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(2).width = 2000
  formseleccio.width = 15000
  formseleccio.Left = formseleccio.Left - 3000
  'If formseleccio.Data1.Recordset.EOF Then MsgBox "Aquest client no te cap DIRECCIO D'ENVIO ASSIGNADA.": Exit Function
  'formseleccio.Data1.Recordset.MoveLast
  'formseleccio.Data1.Recordset.MoveFirst
  'If formseleccio.Data1.Recordset.RecordCount > 1 Then
     formseleccio.Show 1
  '  Else: MsgBox "Nomes hi ha una direcció d'enviament no es pot canviar", vbCritical, "Error"
  'End If
  If seleccioret = 1 Then
        
        If Not formseleccio.Data1.Recordset.EOF Then
           triar_direnvio_client_busqueda = cadbl(formseleccio.DBGrid2.Columns("id"))
        End If
   End If
   Unload formseleccio
   'SendKeys "{TAB}"
   'Command6_Click
   'codimuntadora.SetFocus
End Function


Private Sub bdesbloquejarsap_Click()
   If datacapcalera.Recordset.EOF Then Exit Sub
   datalinies.Refresh
   If UCase(InputBox("Entra la contrasenya per desbloquejar l'albarà." + Chr(10) + "PENSA QUE ABANS DE TRASPASSAR-LO HAS D'ELIMINAR-LO DEL SAP", "CONTRASENYA")) <> "INPLACSA" Then MsgBox "Contrasenya no vàlida", vbCritical, "Error": Exit Sub
   datacapcalera.Database.Execute "update capcaleraalbara set dataenvioasap=null,numalbarasap=0,numfacturasap=0 where numalbara=" + atrim(cadbl(datacapcalera.Recordset!numalbara))
   While Not datalinies.Recordset.EOF
      passarbobinesaentregades False, cadbl(datalinies.Recordset!lotinplacsa), cadbl(datalinies.Recordset!numalbara), Now, cadbl(datacapcalera.Recordset!id_transport)
      datacapcalera.Database.Execute "update linies_expedicions set enviat=false where albara=" + atrim(cadbl(datacapcalera.Recordset!numalbara))
      datalinies.Recordset.MoveNext
   Wend
   datacapcalera.Recordset.Move 0
End Sub
Function revisarsishaREVISATelspaletsabansdenviar(vnumalb As Double) As Boolean
    Dim rst As Recordset
    Set rst = dbbaixes.OpenRecordset("select * from bobinesent where numalbara=" + atrim(vnumalb) + " and (revisatTORERU<>'S' or revisatTORERU=null)")
    If rst.EOF Then
        revisarsishaREVISATelspaletsabansdenviar = True
    End If
    Set rst = Nothing
End Function
Private Sub benviat_Click()
    Dim vdataenviament As String
    refrescarnumerosdalbara
    vdataenviament = InputBox("Entra la data que ho va recullir el transportista.", "Atenció", Format(Now, "dd/mm/yy"))
    If Not IsDate(vdataenviament) Then MsgBox "Aquesta data no es vàlida.", vbCritical, "Error": Exit Sub
    dbbaixes.Execute "update bobinesent set dataentrega=#" + atrim(Format(vdataenviament, "mm/dd/yy")) + "# where numalbara=" + atrim(cadbl(cnumalbara))
    datacapcalera.Database.Execute "update linies_expedicions set enviat=true where albara=" + atrim(cadbl(cnumalbara))
    datacapcalera.Recordset.Move 0
    
End Sub
Sub refrescarnumerosdalbara()
  Dim rstc As Recordset
  Dim vcont As Byte
  Set rstc = datacapcalera.Database.OpenRecordset("SELECT expedicions.enviat as enviatexp, linies_expedicions.enviat as enviatlinies,linies_expedicions.comanda as [Comanda], linies_expedicions.data as [Data_Ex],linies_expedicions.albara as [Albarà], linies_expedicions.observacio as [Obs_Exp], Expedicions.observaciogeneral as [Obs_General] FROM Expedicions RIGHT JOIN linies_expedicions ON Expedicions.data = linies_expedicions.data   order by linies_expedicions.data desc;")
  While Not rstc.EOF And vcont < 150
     Set rst = datacapcalera.Database.OpenRecordset("SELECT capcaleraalbara.numalbara, capcaleraalbara.dataalbara, liniesalbara.lotinplacsa FROM capcaleraalbara INNER JOIN liniesalbara ON capcaleraalbara.numalbara = liniesalbara.numalbara where liniesalbara.lotinplacsa=" + atrim(rstc!comanda))
     If Not rst.EOF Then
        If rst!dataalbara = rstc!data_Ex Then
          If cadbl(rstc![Albarà]) <> cadbl(rst!numalbara) Then
            rstc.Edit
            rstc![Albarà] = rst!numalbara
            rstc.Update
          End If
        End If
     End If
     If (cadbl(rstc![Albarà]) = 0 And rstc!enviatlinies) Or rst.EOF Then
         rstc.Edit
         rstc!enviatlinies = False
         rstc.Update
         datacapcalera.Database.Execute "update expedicions set enviat=true where data=#" + Format(rstc!data_Ex, "mm/dd/yy") + "#"
     End If
     rstc.MoveNext
     vcont = vcont + 1
  Wend
  Set rstc = Nothing
End Sub


Private Sub bpendents_Click()
  Dim rst As Recordset
  refrescarnumerosdalbara
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  Set formseleccio.Data1.Recordset = datacapcalera.Database.OpenRecordset("SELECT linies_expedicions.comanda as [Comanda], linies_expedicions.data as [Data_Ex],linies_expedicions.albara as [Albarà],linies_expedicions.nomclient as [Client], linies_expedicions.observacio as [Obs_Exp], Expedicions.observaciogeneral as [Obs_General] FROM Expedicions RIGHT JOIN linies_expedicions ON Expedicions.data = linies_expedicions.data where not linies_expedicions.enviat order by linies_expedicions.data,linies_expedicions.nomclient,linies_expedicions.comanda;")
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 800
  formseleccio.DBGrid2.Columns(1).width = 1200
  formseleccio.DBGrid2.Columns(2).width = 800
  formseleccio.DBGrid2.Columns(3).width = 2000
  formseleccio.DBGrid2.Columns(4).width = 5000
  formseleccio.DBGrid2.Columns(5).width = 5000
  formseleccio.width = 16000
  formseleccio.caption = "Comandes pendents d'enviar."
  formseleccio.Show 1
  If seleccioret = 1 Then
   'escullir_albara = cadbl(formseleccio.DBGrid2.Columns(0))
  End If
  Unload formseleccio
End Sub

Private Sub CMRmarcatssensepdf_Click()
  CMRsNoEscanejats "Marcats"
End Sub

Private Sub comboquifactura_Change()
possar_logocorrecte
End Sub

Sub possar_logocorrecte()
   If comboquifactura = "Inplacsa" Then logoinplacsa.visible = True: logoplasel.visible = False
   If comboquifactura = "Plasel" Then logoinplacsa.visible = False: logoplasel.visible = True
End Sub

Private Sub comboquifactura_Click()
possar_logocorrecte
End Sub

Private Sub comboquifactura_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub combotransportista_Click()
  cidtransport = combotransportista.ItemData(combotransportista.ListIndex)
  If Framecapcalera.Enabled = True Then
     demanar_metrescubicstransport
     possarobservacio_transport cidtransport
  End If
End Sub
Sub demanar_metrescubicstransport()
   datacapcalera.Recordset!metrescubicstransport = 0
   If InStr(1, combotransportista, "DACHSER") > 0 Then demanar_metrecubics
End Sub
Sub demanar_metrecubics()
  Dim v As String
  v = InputBox("Aquest transportista demana els metres cubics a l'albarà del client, escriu-los sisplau.", "METRES CUBICS")
  If cadbl(v) > 0 Then datacapcalera.Recordset!metrescubicstransport = cadbl(v)
End Sub
Sub possarobservacio_transport(coditransport As Long)
   Dim rst As Recordset
   If cadbl(coditransport) > 0 Then
      'cobservacionstransport = ""
      Set rst = dbcomandes.OpenRecordset("Select * from transportistes where codi=" + atrim(coditransport) + " order by descripcio")
      If Not rst.EOF Then
         If atrim(rst!observaciopredeterminada) <> "" Then
             cobservacionstransport = atrim(rst!observaciopredeterminada)
         End If
      End If
   End If
   Set rst = Nothing
End Sub

Sub escullirtransportista(Optional coditransport As Long)
   Dim rst As Recordset
   If cadbl(coditransport) > 0 Then
      combotransportista = ""
      Set rst = dbcomandes.OpenRecordset("Select * from transportistes where visible=1 and codi=" + atrim(coditransport) + " order by descripcio")
      If Not rst.EOF Then
         combotransportista = atrim(rst!descripcio)
      End If
      Set rst = Nothing
      Exit Sub
        Else
          Set rst = dbcomandes.OpenRecordset("Select codi,descripcio from transportistes where visible=1 order by descripcio")
   End If
   combotransportista.Clear
   While Not rst.EOF
      combotransportista.AddItem atrim(rst!descripcio)
      combotransportista.ItemData(combotransportista.NewIndex) = cadbl(rst!codi)
      rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Private Sub combotransportista_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub combotransportista_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Command1_Click()
   Dim vlotinplacsa As Double
   actualitzartotalsdelinia
   activar_frames False
   If Not datalinies.Recordset.EOF Then
      vlotinplacsa = cadbl(datalinies.Recordset!lotinplacsa)
      If combotipusentrega <> "T" And combotipusentrega <> "P" Then
        preguntartipusdentrega
        If combotipusentrega = "T" Then seltot.tag = "totes": seltot_Click
      End If
   End If
   If datalinies.Recordset.EditMode > 0 Then datalinies.Recordset.Update
   ratoli "espera"
   'calculo l'impost denvasos, si s'ha entrat manualment ho calculo diferent
    If Text32.tag <> "1" Then
         calcular_impost_envasos_liniaalbara
          Else: calcular_impost_envasos_liniaalbara cadbl(datalinies.Recordset!kgimpostenvasos)
    End If
   refrescarnumerosdalbara
   datalinies.Refresh
   datalinies.Recordset.FindFirst "lotinplacsa=" + atrim(vlotinplacsa)
   possar_etiqueta_tanx100mermes datacapcalera.Recordset!numalbara
   If comprovar_quantitatsVSdemanats(datalinies.Recordset, cadbl(datacapcalera.Recordset!id_direnvio)) Then
        MsgBox "Aquest client ha demanat que la quantitat entregada coincideixi amb la demanada." + vbNewLine + "ARREGLA L'ENTREGA PERQUÈ COINCIDEIXI EXACTAMENT.", vbCritical, "ATENCIÓ"
   End If
   ratoli "normal"
   Text32.tag = ""
End Sub
Function comprovar_quantitatsVSdemanats(rstlinies As Recordset, vdirenvio As Double) As Boolean
    Dim rst As Recordset
    Dim rstc As Recordset
    Set rst = dbcomandes.OpenRecordset("select * from clients_envios where id=" + atrim(vdirenvio))
    If Not rst.EOF Then
         If rst!entregaigualademanada Then
             Set rstc = dbcomandes.OpenRecordset("select tubbaseext from comandes where comanda=" + atrim(rstlinies!lotinplacsa))
             If Not rstc.EOF Then
                 If cadbl(rstc!tubbaseext) <> cadbl(rstlinies!quantitat) Then comprovar_quantitatsVSdemanats = True
             End If
         End If
    End If
    Set rst = Nothing
    Set rstc = Nothing
End Function
Sub preguntartipusdentrega()
   Dim te As String
   Dim rst As Recordset
   Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificacio.mdb")
   Set rst = dbplanificacio.OpenRecordset("select entregaparcial from planificaciototes where comanda=" + atrim(datalinies.Recordset!lotinplacsa))
   If Not rst.EOF Then
       If rst!entregaparcial Then
            MsgBox "A T E N C I Ó     A Q U E S T A    E N T R E G A    S E R À   P A R C I A L...", vbExclamation, "A T E N C I Ó"
            te = "P"
       End If
   End If
   While te <> "T" And te <> "P"
     te = UCase(InputBox("Aquesta entrega serà (T)Total o (P)Parcial?", "Tipus d'entrega"))
   Wend
   combotipusentrega = te
End Sub
Private Sub Command10_Click()
   Dim numlot As Double
   Dim rst As Recordset
   Dim numalb As Double
   numlot = cadbl(InputBox("Entra el numero de lot que busques.", "Buscar comanda"))
   If numlot = 0 Then Exit Sub
   If numlot < 150000 Then numalb = numlot: GoTo fi
   Set rst = datacapcalera.Database.OpenRecordset("select * from liniesalbara where lotinplacsa=" + atrim(numlot))
   If rst.EOF Then MsgBox "No s'ha trobat aquesta comanda", vbCritical, "Buscar comanda": Exit Sub
   rst.MoveLast
   If rst.RecordCount = 1 Then numalb = rst!numalbara: GoTo fi
   numalb = escullir_albara(numlot)
fi:
   If numalb > 0 Then
         datacapcalera.Recordset.FindFirst "numalbara=" + atrim(numalb)
         DoEvents
         datalinies.Recordset.FindFirst "lotinplacsa=" + atrim(numlot)
   End If
   Set rst = Nothing
End Sub
Function escullir_albara(numlot As Double) As Double
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  Set formseleccio.Data1.Recordset = datacapcalera.Database.OpenRecordset("SELECT capcaleraalbara.numalbara as NumAlbara, dataalbara, descripcioproducte FROM capcaleraalbara INNER JOIN liniesalbara ON capcaleraalbara.numalbara = liniesalbara.numalbara where liniesalbara.lotinplacsa=" + atrim(numlot))
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 800
  formseleccio.DBGrid2.Columns(1).width = 1000
  formseleccio.DBGrid2.Columns(2).width = 4500
  formseleccio.width = 8000
  formseleccio.Show 1
  If seleccioret = 1 Then
   escullir_albara = cadbl(formseleccio.DBGrid2.Columns(0))
  End If
  Unload formseleccio
End Function

Private Sub Command11_Click()
  formpackinglist.Show 1
End Sub

Private Sub Command12_Click()
 Set dbtmpb = dbcomandes
 Set dbtmp = datacapcalera.Database
 Load paperfrontal
 paperfrontal.comanda = atrim(datalinies.Recordset!lotinplacsa)
 paperfrontal.carregar_elspalets
 paperfrontal.Show 1
    
End Sub

Private Sub Command13_Click()
  Dim vhihainplacsa As Boolean
  Dim vhihaplasel As Boolean
  
  assignardecimalipunt  ' afegeixo aixó perquè en algun cas el sap ha fet el raru amb el decimal i he provat aixó
  
  comprovarsihihafitxerspendentsdimportarasap vhihainplacsa, vhihaplasel, True, True
  If vhihaplasel Then
    MsgBox "Importació de Plasel", vbInformation, "Atenció"
    ShellAndWait "\\servidorsap\seidor_COMUNICADOR\PROGRAMA\Vendes\Plasel\SEI_Importacions.exe"
    mirar_resultat_importacio "\\servidorsap\seidor_COMUNICADOR\LOG\Plasel"
    actualitzar_numerosalbaraSAP_a_produccio
  End If
  If vhihainplacsa Then
  '  Shell "\\servidorsap\seidor_COMUNICADOR\PROGRAMA\Vendes\Inplacsa\SEI_Importacions.exe"
    MsgBox "Importació de Inplacsa", vbInformation, "Atenció"
   ' Unload formimportaciosap
    'formimportaciosap.Show 1
    ShellAndWait "\\servidorsap\seidor_COMUNICADOR\PROGRAMA\Vendes\Inplacsa\SEI_Importacions.exe"
    mirar_resultat_importacio "\\servidorsap\seidor_COMUNICADOR\LOG\Inplacsa"
  End If
  ratoli "espera"
  wait 2
  ' sobretot primer actualitzar factures i despres albarans
  actualitzar_numerosfacturaSAP_a_produccio
  actualitzar_numerosalbaraSAP_a_produccio
  ratoli "normal"
End Sub

Sub actualitzar_numerosfacturaSAP_a_produccio()
    Dim rst As Recordset
    Dim rstlin As Recordset
    Dim rstSAP_inp As Recordset
    Dim rstSAP_pla As Recordset
    Dim rstSAP_albinp As Recordset
    Dim rstSAP_albpla As Recordset
    
    Dim dbsap As Database
    On Error GoTo err
    Set dbsap = OpenDatabase(rutadelfitxer(cami) + "\connexiosap.mdb")
    Set rstSAP_inp = dbsap.OpenRecordset("select * from liniesfacturessap_inplacsa")
    Set rstSAP_pla = dbsap.OpenRecordset("select * from liniesfacturessap_plasel")
    Set rstSAP_albinp = dbsap.OpenRecordset("select * from albaransilinies_inplacsa where [CANCELED]<>'Y' order by Docnum Desc")
    Set rstSAP_albpla = dbsap.OpenRecordset("select * from albaransilinies_plasel  where canceled<>'Y' order by Docnum Desc")
    On Error GoTo 0
    Set rst = datacapcalera.Database.OpenRecordset("select * from capcaleraalbara where numfacturaSAP=0 or numfacturaSAP=null")
    While Not rst.EOF
      If UCase(rst!empresa) = "INPLACSA" Then
'       If rst!numalbara = 11253 Then Stop
   '  Me.caption = atrim(rst.RecordCount) + " / " + atrim(rst.AbsolutePosition)
       rstSAP_inp.FindFirst "trim(albara_produccio) ='" + atrim(rst!numalbara) + "'"
       If rstSAP_inp.NoMatch Then rstSAP_inp.FindFirst "trim(albara_produccio) like '" + atrim(rst!numalbara) + ";*'"
       If rstSAP_inp.NoMatch Then rstSAP_inp.FindFirst "trim(albara_produccio) like ' " + atrim(rst!numalbara) + ";*'"
       If rstSAP_inp.NoMatch Then rstSAP_inp.FindFirst "trim(albara_produccio) like ' " + atrim(rst!numalbara) + ".'"
       If Not rstSAP_inp.NoMatch Then
            rst.Edit
            rst!numfacturasap = cadbl(rstSAP_inp!NumFact)
            rstSAP_albinp.FindFirst "DocEntry=" + atrim(cadbl(rstSAP_inp!BaseEntry))
            If Not rstSAP_albinp.NoMatch Then rst!numalbaraSAP = cadbl(rstSAP_albinp!DocNum)
            rst.Update
             Else
               rst.Edit
               rst!numfacturasap = 0
               rst.Update
       End If
      End If
      If UCase(rst!empresa) = "PLASEL" Then
       If rstSAP_inp.NoMatch Then
        rstSAP_pla.FindFirst "albara_produccio like '*" + atrim(rst!numalbara) + "*'"
        If Not rstSAP_pla.NoMatch Then
             rst.Edit
             rst!numfacturasap = cadbl(rstSAP_pla!NumFact)
             rstSAP_albpla.FindFirst "DocEntry=" + atrim(cadbl(rstSAP_pla!BaseEntry))
             If Not rstSAP_albpla.NoMatch Then rst!numalbaraSAP = cadbl(rstSAP_albpla!DocNum)
             rst.Update
              Else
               rst.Edit
               rst!numalbaraSAP = 0
               rst.Update
        End If
       End If
      End If
      DoEvents
      rst.MoveNext
    Wend
    Exit Sub
err:
    MsgBox "No he pogut accedir al servidor de SAP per mirar els albarans generats.", vbCritical, "Error"
End Sub

Sub actualitzar_numerosalbaraSAP_a_produccio()
    Dim rst As Recordset
    Dim rstlin As Recordset
    Dim rstSAP_inp As Recordset
    Dim rstSAP_pla As Recordset
    Dim dbsap As Database
    On Error GoTo err
    Set dbsap = OpenDatabase(rutadelfitxer(cami) + "\connexiosap.mdb")
    Set rstSAP_inp = dbsap.OpenRecordset("select * from albaransilinies_inplacsa where [CANCELED]<>'Y' order by Docnum Desc")
    Set rstSAP_pla = dbsap.OpenRecordset("select * from albaransilinies_plasel  where canceled<>'Y' order by Docnum Desc")
    On Error GoTo 0
    Set rst = datacapcalera.Database.OpenRecordset("select * from capcaleraalbara where dataenvioasap<>null and numalbaraSAP=0 or numalbaraSAP=null ")
    'Set rstlin = datacapcalera.Database.OpenRecordset("select numalbara,lotinplacsa from liniesalbara")
    While Not rst.EOF
            'busca el lot a SAP empresa INPLACSA
          If UCase(rst!empresa) = "INPLACSA" Then
            rstSAP_inp.FindFirst "AlbaraProd='" + atrim(rst!numalbara) + "' and (linestatus='O' or Targettype<>-1) "
            If Not rstSAP_inp.NoMatch Then
                rst.Edit
                rst!numalbaraSAP = cadbl(rstSAP_inp!DocNum)
                rst.Update
                 Else
                  If DateDiff("d", rst!dataalbara, Now) > 7 Then
                      rst.Edit
                      rst!numalbaraSAP = -1
                      rst!numfacturasap = -1
                      rst.Update
                  End If
            End If
          End If
          If UCase(rst!empresa) = "PLASEL" Then
            If rstSAP_inp.NoMatch Then
                rstSAP_pla.FindFirst "AlbaraProd='" + atrim(rst!numalbara) + "' and (linestatus='O' or targettype<>-1) "
                If Not rstSAP_pla.NoMatch Then
                    rst.Edit
                    rst!numalbaraSAP = cadbl(rstSAP_pla!DocNum)
                    rst.Update
                     Else
                        If DateDiff("d", rst!dataalbara, Now) > 7 Then
                         rst.Edit
                         rst!numalbaraSAP = -1
                         rst!numfacturasap = -1
                         rst.Update
                        End If
                End If
                
            End If
          End If
       rst.MoveNext
    Wend
    Exit Sub
err:
    
End Sub
Sub mirar_resultat_importacio(vdir As String)
  Dim v As String
  Dim vlinia As String
  Dim vmesactual As String
  v = Dir(vdir + "\Log_" + Format(Now, "yyyymmdd") + "*.txt")
  While v <> ""
    If cadbl(Mid(v, Len(v) - 9, 6)) > vgran Then
      vgran = cadbl(Mid(v, Len(v) - 9, 6))
      vmesactual = v
    End If
    'v = Dir(vdir + "\Log_" + Format(Now, "yyyymmdd") + "*.txt")
    v = Dir
  Wend
  If vmesactual <> "" Then
    Open vdir + "\" + vmesactual For Input As #1
    While Not EOF(1)
       Input #1, vlinia
       If InStr(1, UCase(vlinia), "ERROR") > 0 Then obrir_document vdir + "\" + vmesactual:  GoTo fi
    Wend
fi:
    Close #1
  End If

End Sub

Private Sub Command14_Click()
  Dim v As String
  Dim vbases As Integer
  v = cadbl(InputBox("Quantes BASES te aquest albarà?", "IMPRESIO DE PAPERS PER LES BASES", atrim(datacapcalera.Recordset!numbases)))
  If StrPtr(v) = 0 Then Exit Sub
  vbases = v
  If vbases > 0 Then
    imprimirbases vbases
    If MsgBox("Ara imprimiré les etiquetes d'enviament, sisplau possa " + atrim(vbases) + " etiquetes a la bandeja manual de la impresora i fes Acceptar.", vbExclamation + vbDefaultButton2 + vbOKCancel, "IMPRIMIR ETIQUETES ENVIO") = vbOK Then
        imprimiretiquetesenvio vbases
    End If
      Else: datacapcalera.Recordset.Edit: datacapcalera.Recordset!numbases = 0: datacapcalera.Recordset.Update
  End If
End Sub
Sub imprimiretiquetesenvio(vbases As Integer)
  Dim oapp As CRAXDDRT.Application
  Dim vnumcopies As Byte
  Dim i As Byte
  Dim rstcli As Recordset
  Dim rstdire As Recordset
  Dim vnumpalet As Integer
  Dim vpalets As Integer
  Dim oreport As CRAXDDRT.Report
  Dim vdata As Date
  Dim vmesosany As Variant
  vmesosany = Array("Gener", "Febrer", "Març", "Abril", "Maig", "Juny", "Juliol", "Agost", "Setembre", "Octubre", "Novembre", "Desembre")
  If vbases = 0 Then Exit Sub
  vdata = CVDate(datacapcalera.Recordset!dataalbara)
  Set rstdire = dbcomandes.OpenRecordset("select * from clients_envios where id=" + atrim(cadbl(datacapcalera.Recordset!id_direnvio)))
  If rstdire.EOF Then Exit Sub
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "etiqueta_envio_bases.rpt", 1)
  oreport.FormulaFields.GetItemByName("dir1").Text = "'" + atrim(treure_apostruf(atrim(rstdire!nome))) + "'"
  oreport.FormulaFields.GetItemByName("dir2").Text = "'" + atrim(treure_apostruf(atrim(rstdire!domicilie))) + "'"
  oreport.FormulaFields.GetItemByName("dir3").Text = "'" + atrim(treure_apostruf(atrim(rstdire!codipostale))) + " " + atrim(treure_apostruf(atrim(rstdire!poblacioe))) + "'"
  oreport.FormulaFields.GetItemByName("dir4").Text = "'" + atrim(treure_apostruf(atrim(rstdire!provinciae))) + "'"
  oreport.FormulaFields.GetItemByName("bultos").Text = "'" + atrim(vbases) + "'"
  oreport.FormulaFields.GetItemByName("ANY").Text = "'" + atrim(Format(vdata, "yy")) + "'"
  oreport.FormulaFields.GetItemByName("mes").Text = "'" + atrim(vmesosany(Month(vdata) - 1)) + "'"
  oreport.FormulaFields.GetItemByName("dia").Text = "'" + atrim(Day(vdata)) + "'"
  vnumcopies = 1
  If existeix("c:\ordprog.ini") Then vistaprevia = True
  For vnumpalet = 1 To vbases
        If vistaprevia Then
          Load veurereport
          veurereport.CRViewer.ReportSource = oreport
          veurereport.CRViewer.DisplayGroupTree = False
          veurereport.CRViewer.ViewReport
          veurereport.Show 1, Me
          Else
            For i = 1 To vnumcopies
              oreport.PrintOut False, 1
            Next i
        End If
        
  Next vnumpalet
fi:

End Sub

Sub imprimirbases(vbases As Integer, Optional vpersonalitzat As Boolean)
  Dim oapp As CRAXDDRT.Application
  Dim vnumcopies As Byte
  Dim i As Byte
  Dim rstcli As Recordset
  Dim rstdire As Recordset
  Dim vnumpalet As Integer
  Dim vpalets As Integer
  Dim oreport As CRAXDDRT.Report
  Dim vnomclientpersonalitzat As String
  Dim vtamanyfontpersonalitzat As Double
  Dim rst As Recordset
  
  Set rstdire = dbcomandes.OpenRecordset("select * from clients_envios where id=" + atrim(cadbl(datacapcalera.Recordset!id_direnvio)))
  vtamanyfontpersonalitzat = 50
  vnomclientpersonalitzat = atrim(treure_apostruf(atrim(rstdire!nome)))
  Set rst = dbvendes.OpenRecordset("select * from impressiobases_personalitzacio WHERE id_direnvio=" + atrim(cadbl(datacapcalera.Recordset!id_direnvio)))
       If Not rst.EOF Then
           vtamanyfontpersonalitzat = cadbl(rst!tamanyfont)
           vnomclientpersonalitzat = atrim(rst!nome)
       End If
  If vpersonalitzat Then
       vnomclientpersonalitzat = treure_apostruf(InputBox("Escriu com vols el texte del nom del client.", "Nom client", atrim(vnomclientpersonalitzat)))
       vtamanyfontpersonalitzat = cadbl(InputBox("Escriu el tamany de la lletra que vols utilitzar.", "Tamany lletra", atrim(vtamanyfontpersonalitzat)))
       vistaprevia = True
       If vtamanyfontpersonalitzat > 0 Then
         If rst.EOF Then rst.AddNew Else rst.Edit
         rst!tamanyfont = vtamanyfontpersonalitzat
         rst!nome = vnomclientpersonalitzat
         rst!id_direnvio = cadbl(datacapcalera.Recordset!id_direnvio)
         rst.Update
            Else: GoTo fi
       End If
  End If
         
  If vbases = 0 Then Exit Sub
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetnomclientalpalet.rpt", 1)
  'oreport.FormulaFields.GetItemByName("nomdirenvio").Text = "'" + treure_apostruf(etinfodelclient.tag) + "'"
  Set rstcli = dbcomandes.OpenRecordset("select * from clients_codiscomptables where codicomptable=" + atrim(cadbl(datacapcalera.Recordset!codiclient)))
  If rstcli.EOF Or rstdire.EOF Then Exit Sub
  Set rstcli = dbcomandes.OpenRecordset("select * from clients where codi=" + atrim(rstcli!codifabricacio))
  If rstcli.EOF Then Exit Sub
  oreport.FormulaFields.GetItemByName("nomclient").Text = "'" + atrim(treure_apostruf(atrim(rstdire!nome))) + "'"
  oreport.FormulaFields.GetItemByName("poblacio").Text = "'" + atrim(treure_apostruf(atrim(rstdire!poblacioe))) + "'"
  oreport.FormulaFields.GetItemByName("transportista").Text = "'" + atrim(treure_apostruf("Transportista: " + atrim(combotransportista))) + "'"
  oreport.FormulaFields.GetItemByName("nomclient").Text = "'" + atrim(vnomclientpersonalitzat) + "'"
  oreport.Sections("PH").ReportObjects("nomclient1").Font.Size = vtamanyfontpersonalitzat
  oreport.Sections("PH").ReportObjects("nomclient3").Font.Size = vtamanyfontpersonalitzat
  datacapcalera.Recordset.Edit: datacapcalera.Recordset!numbases = vbases: datacapcalera.Recordset.Update
  vnumcopies = 1
  If existeix("c:\ordprog.ini") Then vistaprevia = True
  If Not vistaprevia Then If MsgBox("Vols imprimir-los ara?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then GoTo fi
  For vnumpalet = 1 To vbases
        oreport.FormulaFields.GetItemByName("numdepalets").Text = "'" + atrim(vnumpalet) + "/" + atrim(vbases) + "'"
        If vistaprevia Then
          Load veurereport
          veurereport.CRViewer.ReportSource = oreport
          veurereport.CRViewer.DisplayGroupTree = False
          veurereport.CRViewer.ViewReport
          veurereport.Show 1, Me
          GoTo fi
          Else
            For i = 1 To vnumcopies
              oreport.PrintOut False, 1
            Next i
        End If
        
  Next vnumpalet
fi:
Set rst = Nothing
End Sub
Function comprovarsiesTotalihihabobinessenseseleccionar() As Boolean
   Dim rst As Recordset
   Dim vcomandes As String
   Dim rst2 As Recordset
   vcancelarSAP = False
   comprovarsiesTotalihihabobinessenseseleccionar = False
   Set rst = datacapcalera.Database.OpenRecordset("SELECT quantitat,lotinplacsa as comanda,tipusdeentrega,lotinplacsa  FROM liniesalbara Where numalbara = " + atrim(cadbl(cnumalbara)) + ";")
   While Not rst.EOF
      If rst!tipusdeentrega = "T" Then
       Set rst2 = datacapcalera.Database.OpenRecordset("SELECT bobinesent.comanda, bobinesent.numalbara From bobinesent where comanda=" + atrim(rst!comanda) + " ORDER BY bobinesent.numalbara;")
       If Not rst.EOF Then
         If cadbl(rst2!numalbara) = 0 Then vcomandes = vcomandes + " " + atrim(rst2!comanda)
       End If
      End If
      If comprovar_quantitatsVSdemanats(rst, cadbl(datacapcalera.Recordset!id_direnvio)) Then
           comprovarsiesTotalihihabobinessenseseleccionar = True
           MsgBox "Hi ha lots en aquest albarà que no tenen les quantitas demanades pel client ABANS DE CONTINUAR HAURIES D'ARRECLAR-HO", vbCritical, "Error"
           GoTo fi
      End If
      rst.MoveNext
   Wend
   
   If vcomandes <> "" Then
     If MsgBox("Les comandes " + atrim(vcomandes) + " tenen Bobines/Caixes pendents d'entrega i estas marcant l'albarà com a entrega Total." + Chr(10) + "Vols continuar sense revisar els errors?", vbCritical + vbYesNo + vbDefaultButton2, "A t e n c i ó") = vbYes Then
       comprovarsiesTotalihihabobinessenseseleccionar = False
         Else: comprovarsiesTotalihihabobinessenseseleccionar = True
     End If
   End If
fi:
   Set rst = Nothing
   Set rst2 = Nothing
End Function

Private Sub Command15_Click()
   Dim vvaloranterior As Double
   If ettraspasasap <> "" Then If UCase(InputBox("No es pot calcular un albarà passat a SAP." + vbNewLine + "ESCRIU LA CONTRASENYA PER CALCULAR-HO IGUALMENT." + vbNewLine, "RECALCUL IMPOST")) <> "INPLACSA" Then Exit Sub
   calcular_impost_envasos_liniaalbara
   datalinies.Recordset.Move 0
  Exit Sub
  Dim rstimpost As Recordset
  Dim vmsg As String
  datacapcalera.Recordset.FindFirst "month(dataalbara)=3"
  While Year(datacapcalera.Recordset!dataalbara) = 2024 'And Month(datacapcalera.Recordset!Dataalbara) = 3
   Set rstimpost = datacapcalera.Database.OpenRecordset("select sum(ad_intracom_kgimpost) as Tintracom from impostenvasos where numalbara=" + atrim(datacapcalera.Recordset!numalbara) + " and comanda=" + atrim(datalinies.Recordset!lotinplacsa))
   If rstimpost!Tintracom = 0 Then
    vvaloranterior = cadbl(datalinies.Recordset!kgimpostenvasos)
    calcular_impost_envasos_liniaalbara
    datalinies.Recordset.Move 0
    
    If vvaloranterior <> cadbl(datalinies.Recordset!kgimpostenvasos) Then
         vmsg = vmsg + "Comanda: " + atrim(datalinies.Recordset!lotinplacsa) + vbNewLine + atrim(vvaloranterior) + "Kg  --->  " + atrim(datalinies.Recordset!kgimpostenvasos) + " Kg" + vbNewLine
         'MsgBox "Comanda: " + atrim(datalinies.Recordset!lotinplacsa) + vbNewLine + atrim(vvaloranterior) + "Kg  --->  " + atrim(datalinies.Recordset!kgimpostenvasos) + " Kg"
    End If
   End If
  ' MsgBox "PROXIM"
   datalinies.Recordset.MoveNext
   If datalinies.Recordset.EOF Then datacapcalera.Recordset.MoveNext
   DoEvents
  Wend
  Clipboard.Clear
  Clipboard.SetText vmsg
  MsgBox "Resultat copiat al portapapers." + vbNewLine + vmsg
End Sub
Sub calcular_impost_envasos_liniaalbara(Optional vKgimpostmanual As Double)
 Dim rstalb As Recordset
   Dim rstimpost As Recordset
   Set rstalb = datacapcalera.Database.OpenRecordset("select * from liniesalbara where id=" + atrim(datalinies.Recordset!ID))
   If rstalb.EOF Then GoTo fi
   Set rstimpost = datacapcalera.Database.OpenRecordset("select * from impostenvasos")
   calcular_impost_comanda rstalb, rstimpost, cadbl(vKgimpostmanual)
   If cadbl(datalinies.Recordset!kgimpostenvasos) >= cadbl(datalinies.Recordset!kgtotalsbruts) Then
       If Not existeix("c:\ordprog.ini") Then MsgBox "Els Kg d'IMPOST son superiors als Kg de l'albarà." + vbNewLine + "PARLAR AMB EN MIRALLES.", vbCritical, "ATENCIÓ"
   End If
fi:
   Set rstc = Nothing
   Set rstalb = Nothing
   Set rstimpost = Nothing
End Sub
Function sumatotalpackinglist(vnumc As Double) As Double
   Dim rst As Recordset
   Dim vnumcomandes As String
   Dim rstc As Recordset
   Set rstc = dbstocks.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(vnumc))
   If Not rstc.EOF Then vnumcomandes = "'" + atrim(vnumc) + "'" + IIf(rstc!linkcomanda1 > 0, ",'" + atrim(rstc!linkcomanda1) + "'", "") + IIf(rstc!linkcomanda2 > 0, ",'" + atrim(rstc!linkcomanda2) + "'", "")
   Set rst = dbstocks.OpenRecordset("SELECT Sum(Bobines.pesdelproveidor) AS SumaDepesdelproveidor, Sum(([bobines].[pesdelproveidor]/[bobines].[mts])*[parcials].[metres]) AS pesTbobines FROM Bobines LEFT JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet) WHERE Parcials.comanda in (" + vnumcomandes + ") AND Parcials.orcomassignacio<>'500';")
   If Not rst.EOF Then sumatotalpackinglist = rst!pesTbobines
   Set rst = Nothing
End Function
Sub calcular_totals_productorextern(vnumc As Double, rstimpost As Recordset)
   Dim rstc As Recordset
   Dim rstcompra As Recordset
   Dim rstprov As Recordset
   Dim vmicres As Double
   Dim vfuelles As Double
   Dim vkgcompra As Double
   Dim vkgimpost As Double
   Dim vquantitat As Double
   
   Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
   
   Set rstc = dbcompres.OpenRecordset("SELECT albaransbip.Id,comandesxlinia.kgcompra,comandesxlinia.kgcomanda,liniescompra.quantitatkg,albaransbip.quantitat as QuantitatAlbaransBip FROM (liniescompra RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra) LEFT JOIN albaransbip ON liniescompra.idliniacompra = albaransbip.idliniacompra WHERE (((comandesxlinia.numcomanda)=" + atrim(vnumc) + "));")
   If rstc.EOF Then GoTo fi
   Set rstprov = dbcompres.OpenRecordset("SELECT proveidors.tipusproveidorIMPOST,materials.mesuarespcompra as MesuraCompraMaterial, materials.grmcm3,albaransbip.KgImpostEnvasos FROM (((albaransbip LEFT JOIN liniescompra ON albaransbip.idliniacompra = liniescompra.idliniacompra) LEFT JOIN capcalera ON liniescompra.idcompra = capcalera.id) LEFT JOIN proveidors ON capcalera.codiproveidor = proveidors.codi) LEFT JOIN materials ON liniescompra.codimaterial = materials.codi WHERE albaransbip.id=" + atrim(rstc!ID))
   'set rstprov=dbcompres.OpenRecordset("SELECT proveidors.tipusproveidorIMPOST, materials.grmcm3, albaransbip.KgImpostEnvasos, comandesxlinia.numcomanda, comandesxlinia.kgcompra FROM comandesxlinia RIGHT JOIN ((((albaransbip LEFT JOIN liniescompra ON albaransbip.idliniacompra = liniescompra.idliniacompra) LEFT JOIN capcalera ON liniescompra.idcompra = capcalera.id) LEFT JOIN proveidors ON capcalera.codiproveidor = proveidors.codi) LEFT JOIN materials ON liniescompra.codimaterial = materials.codi) ON comandesxlinia.idliniacompra = liniescompra.idliniacompra WHERE (((albaransbip.Id)=20884) AND ((comandesxlinia.numcomanda)="+atrim(vnumc)+"));"
   vkgcompra = rstc![kgcompra]
   
   If Not rstprov.EOF Then
       vkgimpost = cadbl(rstprov!kgimpostenvasos)
       If rstprov!MesuraCompraMaterial = "Un" Or rstprov!MesuraCompraMaterial = "Mts" Then
             vkgimpost = (cadbl(rstprov!kgimpostenvasos))
               Else
                vkgimpost = (vkgcompra * vkgimpost) / rstc!quantitatkg
       End If
       Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
       If rstc.EOF Then Exit Sub
       
        rstimpost!Espanya_KgIMPOST = 0
        rstimpost!eSPANYA_KgNOIMPOST = 0
        rstimpost!Imp_mes_Esp_KgIMPOST = 0
        rstimpost!Imp_mes_Esp_KgNOIMPOST = 0
        rstimpost!Ad_Intracom_KgIMPOST = 0
        rstimpost!aD_iNTRACOM_KgNOIMPOST = 0
       
       If rstprov!tipusproveidorIMPOST = "Espanyol" Then
        rstimpost!Espanya_KgIMPOST = cadbl(vkgimpost)
        rstimpost!eSPANYA_KgNOIMPOST = 0
       End If
       If rstprov!tipusproveidorIMPOST = "Importació" Then
        rstimpost!Imp_mes_Esp_KgIMPOST = cadbl(vkgimpost)
        rstimpost!Imp_mes_Esp_KgNOIMPOST = 0
       End If
       If rstprov!tipusproveidorIMPOST = "Intracomunitari" Then
        rstimpost!Ad_Intracom_KgIMPOST = cadbl(vkgimpost)
        rstimpost!aD_iNTRACOM_KgNOIMPOST = 0
       End If
       With rstc
       vmicres = micresmaterial(cadbl(!mesuraesp), cadbl(!espessor), atrim(!tubolam))
       'vfuelles = ((cadbl(!fuellebasesol) + cadbl(!fuellebocasol)) / 100)
       rstimpost!m2perpeça = ((((!longitudsol + (!solapasol / 2)) / 100) * (!amplesol / 100)) * IIf(!migelaboratsol <> "L", 2, 1))
'       rstimpost!kgm2 = (rstimpost!m2perpeça * vmicres * rstprov!grmcm3) / 1000
       End With
       rstimpost!kgm2 = (vmicres / 1000) * cadbl(rstprov!grmcm3)
       rstimpost!impostFabricacioExterna = True
       If vkgimpost = 0 Then MsgBox "Es calcularà un valor teoric d'IMPOST ENVASOS per aquesta comanda." + vbNewLine + "S'ha de demanar a compres quin IMPOST s'ha d'aplicar i canviar-lo del teòric si correspon.", vbExclamation, "ATENCIÓ"
   End If
fi:
   
   Set rstc = Nothing
   Set rstprov = Nothing
   Set rstcompra = Nothing
End Sub
Sub calcular_impost_comanda(rstalb As Recordset, rstimpost As Recordset, Optional vKgimpostmanual As Double)
   Dim vsql As String
   Dim vsqlwhere As String
   Dim rstc As Recordset
   Dim vnumc As Double
   Dim vtotalkgm2 As Double
   Dim vkgimpost As Double
   Dim vkgimpostTotal As Double
   Dim vkgtotalpackinglist As Double
   Dim vproporciokgimpost As Double
   Dim vKGventaPROPORCIOPERLACAPA As Double
   Dim vkgzipper As Double
   Dim vRegimFiscal As String
   Dim vsenseimpost As Boolean
   
   If rstalb.EOF Then Exit Sub
   Set rstc = datacapcalera.Database.OpenRecordset("select * from comandes where comanda=" + atrim(rstalb!lotinplacsa))
   If rstc.EOF Then Exit Sub
   vnumc = cadbl(rstc!comanda)
   vtotalkgm2 = 0
   If cadbl(vKgimpostmanual) = -1 Then vsenseimpost = True
   'vkgtotalpackinglist = sumatotalpackinglist(rstalb!lotinplacsa)
calcularvalorspackinglist:
   If cadbl(vnumc) = 0 Then GoTo ficalculvalorpackinglist
      'CALCULAR EL TAN PERCENT DE MATERIAL SOBRE EL PACKINGLIST
   
   Set rstimpost = datacapcalera.Database.OpenRecordset("select * from impostenvasos where numalbara=" + atrim(rstalb!numalbara) + " and comanda=" + atrim(vnumc))
   If rstimpost.EOF Then
        rstimpost.AddNew
        rstimpost!numalbara = rstalb!numalbara
        rstimpost!comanda = vnumc
         Else: rstimpost.Edit
   End If
   If cadbl(rstalb!ID) = 0 Then MsgBox "HI HA UN ERROR EN L'IDLINIAALBARA DE CALCUL DE L'IMPOST AVISEU A L'INFORMÀTIC SOBRE AQUEST ERROR EN AQUEST ALBARÀ.", vbCritical, "GRACIES"
   rstimpost!idliniAaLbara = rstalb!ID
   vRegimFiscal = ImpostEnv_regimfiscalREFINPLACSA(rstalb!codiproducte)
   rstimpost!regimfiscal = vRegimFiscal
   rstimpost!idliniAaLbara = rstalb!ID
   rstimpost!paisventa = paisdedirecciodenviament(rstalb!numalbara)
   rstimpost!kgm2 = 0
   If nohihapackinglist(vnumc) Then
      calcular_totals_productorextern vnumc, rstimpost: GoTo continuarcalculs
   End If
   
   If vsenseimpost Then
         rstimpost!Ad_Intracom_TKg = 0:   rstimpost!Imp_mes_Esp_TKg = 0:   rstimpost!Espanya_TKg = 0
         rstimpost!Imp_mes_Esp_KgIMPOST = 0: rstimpost!Imp_mes_Esp_KgNOIMPOST = 0: rstimpost!Espanya_KgIMPOST = 0: rstimpost!eSPANYA_KgNOIMPOST = 0: rstimpost!Ad_Intracom_KgIMPOST = 0: rstimpost!aD_iNTRACOM_KgNOIMPOST = 0
         rstimpost![Ad_Intracom_%NOIMPOST] = 0: rstimpost![Espanya_%NOIMPOST] = 0: rstimpost![Espanya_%impost] = 0: rstimpost![Ad_intracom_%impost] = 0: rstimpost![Imp_mes_Esp_%impost] = 0: rstimpost![Imp_mes_Esp_%NOIMPOST] = 0
         rstimpost!Ad_intracom_Ttanper100 = 0:   rstimpost!Imp_mes_Esp_Ttanper100 = 0:   rstimpost!Espanya_Ttanper100 = 0
         rstimpost!kgmermaad_intracom = 0: rstimpost!kgmermaimp_mes_esp = 0: rstimpost!kgmermaespanya = 0: rstimpost!kgmermaad_intracom = 0: rstimpost!kgmermaimp_mes_esp = 0: rstimpost!kgmermaespanya = 0:
        ' rstimpost.Update
        ' rstimpost.Edit
         vKgimpostmanual = 0
         GoTo continuarcalculs
   End If
   
   dbimpost.Execute "delete * from ParcialsBobinesCalculats where idtaulaimpost=" + atrim(rstimpost!ID)
   
   'faig un calcul per treure els kgm2*** AQUEST NO COMPTA PER DADES
   vTipusImpostLinia = ""
   vsqlwhere = " WHERE Parcials_DBL.orcomassignacio<>'500' and Parcials_DBL.comanda='" + atrim(vnumc) + "'"
   rstimpost!Imp_mes_Esp_KgIMPOST = sumar_impost_linies(vsqlwhere, rstimpost)
   
   'importadors amb impost
   vTipusImpostLinia = "IMf"
   vsqlwhere = " WHERE Parcials_DBL.orcomassignacio<>'500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and tipusproveidorIMPOST='Importació' and (palets.teimpost=true)"
   rstimpost!Imp_mes_Esp_KgIMPOST = sumar_impost_linies(vsqlwhere, rstimpost)
   'importadors sense impost
   vTipusImpostLinia = "IMno"
   vsqlwhere = " WHERE Parcials_DBL.orcomassignacio<>'500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and tipusproveidorIMPOST='Importació' and (palets.teimpost=false)"
   rstimpost!Imp_mes_Esp_KgNOIMPOST = sumar_impost_linies(vsqlwhere, rstimpost)
   
   'ESPANYOL amb impost
   vTipusImpostLinia = "ESf"
   vsqlwhere = " WHERE Parcials_DBL.orcomassignacio<>'500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and tipusproveidorIMPOST='Espanyol' and (palets.teimpost=true)"
   rstimpost!Espanya_KgIMPOST = sumar_impost_linies(vsqlwhere, rstimpost)
   'ESPANYOL sense impost
   vTipusImpostLinia = "ESno"
   vsqlwhere = " WHERE Parcials_DBL.orcomassignacio<>'500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and tipusproveidorIMPOST='Espanyol' and (palets.teimpost=false)"
   rstimpost!eSPANYA_KgNOIMPOST = sumar_impost_linies(vsqlwhere, rstimpost)
   
   'iNTRACOMUNITARI amb impost
   vTipusImpostLinia = "ICf"
   vsqlwhere = " WHERE Parcials_DBL.orcomassignacio<>'500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and (tipusproveidorIMPOST='Intracomunitari' or tipusproveidorIMPOST is null) and (palets.teimpost=true)"
   rstimpost!Ad_Intracom_KgIMPOST = sumar_impost_linies(vsqlwhere, rstimpost)
   'INTRACOMUNITARI sense impost
   vTipusImpostLinia = "ICno"
   vsqlwhere = " WHERE Parcials_DBL.orcomassignacio<>'500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and (tipusproveidorIMPOST='Intracomunitari' or tipusproveidorIMPOST is null) and (palets.teimpost=false)"
   rstimpost!aD_iNTRACOM_KgNOIMPOST = sumar_impost_linies(vsqlwhere, rstimpost)
   'posso el tan % que correspont a cada import
continuarcalculs:
 '% importacio
   rstimpost![Imp_mes_Esp_%NOIMPOST] = 0
   If (rstimpost!Imp_mes_Esp_KgIMPOST + rstimpost!Imp_mes_Esp_KgNOIMPOST) > 0 Then rstimpost![Imp_mes_Esp_%NOIMPOST] = (rstimpost!Imp_mes_Esp_KgNOIMPOST * 100) / (rstimpost!Imp_mes_Esp_KgIMPOST + rstimpost!Imp_mes_Esp_KgNOIMPOST)
   rstimpost![Imp_mes_Esp_%impost] = 0
   If (rstimpost!Imp_mes_Esp_KgIMPOST + rstimpost!Imp_mes_Esp_KgNOIMPOST) > 0 Then rstimpost![Imp_mes_Esp_%impost] = (rstimpost!Imp_mes_Esp_KgIMPOST * 100) / (rstimpost!Imp_mes_Esp_KgIMPOST + rstimpost!Imp_mes_Esp_KgNOIMPOST)
 '% espanya
   rstimpost![Espanya_%NOIMPOST] = 0
   If (rstimpost!Espanya_KgIMPOST + rstimpost!eSPANYA_KgNOIMPOST) > 0 Then rstimpost![Espanya_%NOIMPOST] = (rstimpost!eSPANYA_KgNOIMPOST * 100) / (rstimpost!Espanya_KgIMPOST + rstimpost!eSPANYA_KgNOIMPOST)
   rstimpost![Espanya_%impost] = 0
   If (rstimpost!Espanya_KgIMPOST + rstimpost!eSPANYA_KgNOIMPOST) > 0 Then rstimpost![Espanya_%impost] = (rstimpost!Espanya_KgIMPOST * 100) / (rstimpost!Espanya_KgIMPOST + rstimpost!eSPANYA_KgNOIMPOST)
 '% Intracomunitari
   rstimpost![Ad_Intracom_%NOIMPOST] = 0
   If (rstimpost!Ad_Intracom_KgIMPOST + rstimpost!aD_iNTRACOM_KgNOIMPOST) > 0 Then rstimpost![Ad_Intracom_%NOIMPOST] = (rstimpost!aD_iNTRACOM_KgNOIMPOST * 100) / (rstimpost!Ad_Intracom_KgIMPOST + rstimpost!aD_iNTRACOM_KgNOIMPOST)
   rstimpost![Ad_intracom_%impost] = 0
   If (rstimpost!Ad_Intracom_KgIMPOST + rstimpost!aD_iNTRACOM_KgNOIMPOST) > 0 Then rstimpost![Ad_intracom_%impost] = (rstimpost!Ad_Intracom_KgIMPOST * 100) / (rstimpost!Ad_Intracom_KgIMPOST + rstimpost!aD_iNTRACOM_KgNOIMPOST)
   
   rstimpost!Ad_Intracom_TKg = rstimpost![Ad_Intracom_KgIMPOST] + rstimpost![aD_iNTRACOM_KgNOIMPOST]
   rstimpost!Imp_mes_Esp_TKg = rstimpost![Imp_mes_Esp_KgIMPOST] + rstimpost![Imp_mes_Esp_KgNOIMPOST]
   rstimpost!Espanya_TKg = rstimpost![Espanya_KgIMPOST] + rstimpost![eSPANYA_KgNOIMPOST]
   
   rstimpost!Ad_intracom_Ttanper100 = 0
   rstimpost!Imp_mes_Esp_Ttanper100 = 0
   rstimpost!Espanya_Ttanper100 = 0
   
   If (cadbl(rstimpost!Ad_Intracom_KgIMPOST) + cadbl(rstimpost!Imp_mes_Esp_KgIMPOST) + cadbl(rstimpost!Espanya_KgIMPOST)) > 0 Then
          'rstimpost!Ad_intracom_Ttanper100 = (rstimpost!Ad_Intracom_KgIMPOST * 100) / (rstimpost!Ad_Intracom_KgIMPOST + rstimpost!Imp_mes_Esp_KgIMPOST + rstimpost!Espanya_KgIMPOST)
          'rstimpost!Imp_mes_Esp_Ttanper100 = (rstimpost!Imp_mes_Esp_KgIMPOST * 100) / (rstimpost!Ad_Intracom_KgIMPOST + rstimpost!Imp_mes_Esp_KgIMPOST + rstimpost!Espanya_KgIMPOST)
          'rstimpost!Espanya_Ttanper100 = (rstimpost!Espanya_KgIMPOST * 100) / (rstimpost!Ad_Intracom_KgIMPOST + rstimpost!Imp_mes_Esp_KgIMPOST + rstimpost!Espanya_KgIMPOST)
          rstimpost!Ad_intracom_Ttanper100 = (cadbl(rstimpost!Ad_Intracom_KgIMPOST) * 100) / (cadbl(rstimpost!Ad_Intracom_KgIMPOST) + cadbl(rstimpost!Imp_mes_Esp_KgIMPOST) + cadbl(rstimpost!Espanya_KgIMPOST) + cadbl(rstimpost!aD_iNTRACOM_KgNOIMPOST) + cadbl(rstimpost!Imp_mes_Esp_KgNOIMPOST) + cadbl(rstimpost!eSPANYA_KgNOIMPOST))
          rstimpost!Imp_mes_Esp_Ttanper100 = (cadbl(rstimpost!Imp_mes_Esp_KgIMPOST) * 100) / (cadbl(rstimpost!Ad_Intracom_KgIMPOST) + cadbl(rstimpost!Imp_mes_Esp_KgIMPOST) + cadbl(rstimpost!Espanya_KgIMPOST) + cadbl(rstimpost!aD_iNTRACOM_KgNOIMPOST) + cadbl(rstimpost!Imp_mes_Esp_KgNOIMPOST) + cadbl(rstimpost!eSPANYA_KgNOIMPOST))
          rstimpost!Espanya_Ttanper100 = (cadbl(rstimpost!Espanya_KgIMPOST * 100)) / (rstimpost!Ad_Intracom_KgIMPOST + cadbl(rstimpost!Imp_mes_Esp_KgIMPOST) + cadbl(rstimpost!Espanya_KgIMPOST) + cadbl(rstimpost!aD_iNTRACOM_KgNOIMPOST) + cadbl(rstimpost!Imp_mes_Esp_KgNOIMPOST) + cadbl(rstimpost!eSPANYA_KgNOIMPOST))
   End If
   vkgtotalpackinglist = vkgtotalpackinglist + (cadbl(rstimpost!Espanya_KgIMPOST) + cadbl(rstimpost!eSPANYA_KgNOIMPOST) + rstimpost!Imp_mes_Esp_KgIMPOST + rstimpost!Imp_mes_Esp_KgNOIMPOST + rstimpost!Ad_Intracom_KgIMPOST + rstimpost!aD_iNTRACOM_KgNOIMPOST)
   vtotalkgm2 = vtotalkgm2 + Redondejar(cadbl(rstimpost!kgm2), 6)
   If atrim(rstalb!tipusdeentrega) = "T" And Not vsenseimpost Then posarlamerma rstimpost
   rstimpost.Update
   If vnumc = cadbl(rstc!comanda) Then vnumc = cadbl(rstc!linkcomanda1): GoTo calcularvalorspackinglist
   If vnumc = cadbl(rstc!linkcomanda1) Then vnumc = cadbl(rstc!linkcomanda2): GoTo calcularvalorspackinglist
ficalculvalorpackinglist:
    'CALCULAR ELS KG DE MATERIAL REAL ENTREGAT AL CLIENT
    Set rstimpost = datacapcalera.Database.OpenRecordset("select * from impostenvasos where numalbara=" + atrim(rstalb!numalbara) + " and comanda=" + atrim(rstalb!lotinplacsa))
    If cadbl(rstimpost!m2perpeça) > 0 Then
          vkgimpost = cadbl(rstalb!unitats) * (cadbl(rstimpost!m2perpeça) * Redondejar(vtotalkgm2, 6))
            Else:
              vkgimpost = (vtotalkgm2 * migelaboratultimaseccio(rstalb!lotinplacsa)) * ((cadbl(rstalb!ampladamaterial) / 1000) * cadbl(rstalb!metreslineals))
              If datalinies.Recordset!espesnet And cadbl(rstc!linkcomanda1) = 0 Then
                  If vkgimpost > datalinies.Recordset!kgtotalsnets Then vkgimpost = datalinies.Recordset!kgtotalsnets
              End If
    End If
    'si s'ha possat els kg manualment faig el calcul de mermes i altres a partir d'aquest valor
    If cadbl(vKgimpostmanual) > 0 Then
       vkgimpost = vKgimpostmanual
    End If
    If rstimpost!impostFabricacioExterna Then
       vkgimpost = vkgtotalpackinglist
    End If
    Set rstimpost = datacapcalera.Database.OpenRecordset("select * from impostenvasos where (comanda=" + atrim(rstc!comanda) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2) + ") and numalbara=" + atrim(datalinies.Recordset!numalbara))
    While Not rstimpost.EOF
       rstimpost.Edit
       If vkgtotalpackinglist > 0 Then   'busco el tan100 d'aquesta capa sobra el total de la venta
          rstimpost!Tanper100KgVendaVsTotal = ((cadbl(rstimpost!Espanya_KgIMPOST) + cadbl(rstimpost!eSPANYA_KgNOIMPOST) + cadbl(rstimpost!Imp_mes_Esp_KgIMPOST) + cadbl(rstimpost!Imp_mes_Esp_KgNOIMPOST) + cadbl(rstimpost!Ad_Intracom_KgIMPOST) + cadbl(rstimpost!aD_iNTRACOM_KgNOIMPOST)) * 100) / vkgtotalpackinglist
       End If
      
    'proporcio de la venta sobre aquesta capa
       vproporciokgimpost = (vkgimpost * cadbl(rstimpost!Tanper100KgVendaVsTotal)) / 100
       rstimpost!kgventaEspanya = ((vproporciokgimpost * cadbl(rstimpost!Espanya_Ttanper100) / 100)) ' * rstimpost![Espanya_%impost]) / 100
       rstimpost!kgventaImp_mes_esp = ((vproporciokgimpost * cadbl(cadbl(rstimpost!Imp_mes_Esp_Ttanper100)) / 100)) ' * rstimpost![Imp_mes_Esp_%impost]) / 100
       rstimpost!kgventaad_intracom = ((vproporciokgimpost * cadbl(cadbl(rstimpost!Ad_intracom_Ttanper100)) / 100)) ' * cadbl(rstimpost![Ad_intracom_%impost])) / 100
       vkgimpostTotal = vkgimpostTotal + (rstimpost!kgventaad_intracom + rstimpost!kgventaImp_mes_esp + rstimpost!kgventaEspanya)
       rstimpost!impostposatmanualment = IIf(cadbl(vKgimpostmanual) > 0 Or vsenseimpost, True, False)
       rstimpost.Update
       rstimpost.MoveNext
    Wend
    vkgzipper = 0
    vkgzipper = calcularKgzipper(vkgimpostTotal)
    
    '  calcular la merma
    
    Set rstimpost = datacapcalera.Database.OpenRecordset("select * from impostenvasos where (comanda=" + atrim(rstc!comanda) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2) + ") and numalbara=" + atrim(datalinies.Recordset!numalbara))
    With rstimpost
    While Not rstimpost.EOF
        .Edit
        !kgmermaad_intracom = 0: !kgmermaimp_mes_esp = 0: !kgmermaespanya = 0: !kgmermaad_intracom = 0: !kgmermaimp_mes_esp = 0: !kgmermaespanya = 0
        If atrim(rstalb!tipusdeentrega) = "T" And Not vsenseimpost Then
            !kgmermaad_intracom = !Ad_Intracom_KgIMPOST - sumaKgVendaImpost(rstimpost!comanda, "AI", rstc) 'rstimpost!KgVentaAd_Intracom
            !kgmermaimp_mes_esp = !Imp_mes_Esp_KgIMPOST - sumaKgVendaImpost(rstimpost!comanda, "IE", rstc) 'rstimpost!KgVentaImp_mes_Esp
            !kgmermaespanya = !Espanya_KgIMPOST - sumaKgVendaImpost(rstimpost!comanda, "ES", rstc) 'rstimpost!KgVentaespanya
            !kgmermaad_intracom = !kgmermaad_intracom + !KgMermaIMPOST_AD_capa
            !kgmermaimp_mes_esp = !kgmermaimp_mes_esp + !KgMermaIMPOST_IE_capa
            !kgmermaespanya = !kgmermaespanya + !KgMermaIMPOST_ES_capa
             Else
               !kgmermaad_intracom = 0: !kgmermaimp_mes_esp = 0: !kgmermaespanya = 0
        End If
        !regimfiscal = vRegimFiscal
        .Update
        .MoveNext
    Wend
    End With
    
    
'GUARDO ELS VALORS DE VENTA IMPOST A LA LINIA D'ALBARA
    rstalb.Edit
    rstalb!kgimpost100per100 = Redondejar(vkgimpost, 1)  'AQUI POSSO EL 100% PER PODER COMPROVAR-LO SI ESTÀ CORRECTE

    rstalb!KgImpost_primari = Redondejar(vkgimpostTotal, 3)
    rstalb!KgImpost_secondari = calcularKGbosses(rstalb!KgImpost_primari)
    rstalb!KgImpost_terciari = calcularKgfilmestirable(rstalb!KgImpost_primari)
    rstalb!kgimpost_zipper = vkgzipper
    
    If vKgimpostmanual = 0 Then
           rstalb!kgimpostenvasos = Redondejar(rstalb!KgImpost_primari + rstalb!KgImpost_secondari + rstalb!KgImpost_terciari + rstalb!kgimpost_zipper, 2) 'AQUI HAURÀ DE SER vkgimpostTotal QUAN COMPROVEM SI ES CORRECTE
            Else: rstalb!kgimpostenvasos = vKgimpostmanual
    End If
    
    rstalb!eurokg_impost = cadbl(llegir_ini("General", "PreuImpostEnvasos", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"))
    If noescobraalclient(rstalb!lotinplacsa) Then rstalb!eurokg_impost = 0
    rstalb.Update
    
End Sub
Function noescobraalclient(vnumc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select pvp,pvpdolar from comandes where comanda=" + atrim(vnumc), , dbReadOnly)
   If Not rst.EOF Then
      If cadbl(rst!pvp) = -1 Or cadbl(rst!pvpdolar) = -1 Then
            noescobraalclient = True
      End If
   End If
   Set rst = Nothing
End Function
Function sumaKgVendaImpost(vnumc As Double, vTipus As String, rstc As Recordset) As Double
  Dim rst As Recordset
  Set rst = datacapcalera.Database.OpenRecordset("select * from impostenvasos where comanda=" + atrim(vnumc)) ' + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2))
'  MsgBox rst!comanda
  While Not rst.EOF
    sumaKgVendaImpost = sumaKgVendaImpost + IIf(vTipus = "AI", rst!kgventaad_intracom, IIf(vTipus = "IE", rst!kgventaImp_mes_esp, IIf(vTipus = "ES", cadbl(rst!kgventaEspanya), 0)))
    rst.MoveNext
  Wend
  Set rst = Nothing
End Function
Function nohihapackinglist(vnumc As Double) As Boolean
  Dim rst As Recordset
  Set rst = dbstocks.OpenRecordset("select * from parcials where comanda='" + atrim(vnumc) + "'")
  If rst.EOF Then nohihapackinglist = True
  Set rst = Nothing
End Function
Function calcularKgzipper(vimpostprimari As Double) As Double
  Dim rst As Recordset
  Dim rstaccessoris As Recordset
  'si no hi ha primari no compto zipper
  If vimpostprimari = 0 Then Exit Function
  Set rst = dbcomandes.OpenRecordset("SELECT comandes.cinta, productes.ruta FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi where comanda=" + atrim(datalinies.Recordset!lotinplacsa))
  If rst.EOF Then GoTo fi
  Set rstaccessoris = dbcomandes.OpenRecordset("Select * from accessoris where codi=" + atrim(cadbl(rst!cinta)))
  If rstaccessoris.EOF Then GoTo fi
  If InStr(1, rst!ruta, "S") Then
      calcularKgzipper = ((cadbl(rstaccessoris!Kg_Imp_Env) / 1000) * (datalinies.Recordset!ampladamaterial / 1000)) * datalinies.Recordset!unitats
  End If
fi:
  Set rst = Nothing
End Function
Function calcularKgfilmestirable(vimpostprimari As Double) As Double
   Dim vpespalet As Double
   vpespalet = 1.5
   If vimpostprimari = 0 Then
         'si no hi ha impost no comptem film
         calcularKgfilmestirable = 0
          Else
                'palets de la comanda multiplicat pel pes de film estimat per palet  2,2Kg
            calcularKgfilmestirable = cadbl(datalinies.Recordset!numerodepalets) * vpespalet
    End If
End Function
Function calcularKGbosses(vimpostprimari As Double) As Double
   Dim rst As Recordset
   Dim vpessacrebobinadora As Double
   Dim vpessacsoldadora As Double
   vpessacsoldadora = 0.14
   vpessacrebobinadora = 0.07
   If vimpostprimari = 0 Then
            'si no hi ha impost no comptem bosses
         calcularKGbosses = 0
          Else
            Set rst = dbcomandes.OpenRecordset("SELECT comandes.unitatsxsac, productes.ruta FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi where comanda=" + atrim(datalinies.Recordset!lotinplacsa))
            If rst.EOF Then Exit Function
            'si es SOLDADORA i son sacs i es compten els sacs si fossin caixes no
            vpessac = vpessacrebobinadora
            If Mid(rst!ruta, Len(rst!ruta), 1) = "S" Then vpessac = 0: If cadbl(rst!unitatsxsac) > 0 Then vpessac = vpessacsoldadora
            calcularKGbosses = cadbl(datalinies.Recordset!numbobs) * vpessac
   End If
   Set rst = Nothing
End Function
Sub posarlamerma(rstimpost As Recordset)
   Dim vsqlwhere As String
   Dim vnumc As Double
   
   vnumc = rstimpost!comanda
   With rstimpost
   'Espanya
   'Verd
Espanya:
   vTipusImpostLinia = "ESmv"
   vsqlwhere = " WHERE materials.colorreciclatge='Verd' and Parcials_DBL.orcomassignacio='500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and tipusproveidorIMPOST='Espanyol' and (palets.teimpost=true)"
   !KgMermaIMPOST_ES_capa_Verd = sumar_impost_linies(vsqlwhere, rstimpost)
   'Blau
   vTipusImpostLinia = "ESmb"
   vsqlwhere = " WHERE materials.colorreciclatge='Blau' and Parcials_DBL.orcomassignacio='500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and tipusproveidorIMPOST='Espanyol' and (palets.teimpost=true)"
   !KgMermaIMPOST_ES_capa_Blau = sumar_impost_linies(vsqlwhere, rstimpost)
   'Vermell
   vTipusImpostLinia = "ESmm"
   vsqlwhere = " WHERE materials.colorreciclatge='Vermell' and Parcials_DBL.orcomassignacio='500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and tipusproveidorIMPOST='Espanyol' and (palets.teimpost=true)"
   !KgMermaIMPOST_ES_capa_Vermell = sumar_impost_linies(vsqlwhere, rstimpost)
   
   !KgMermaIMPOST_ES_capa = !KgMermaIMPOST_ES_capa_Verd + !KgMermaIMPOST_ES_capa_Blau + !KgMermaIMPOST_ES_capa_Vermell
   
   
   'Importador
   'Verd
Importador:
   vTipusImpostLinia = "IMmv"
   vsqlwhere = " WHERE materials.colorreciclatge='Verd' and Parcials_DBL.orcomassignacio='500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and tipusproveidorIMPOST='Importació' and (palets.teimpost=true)"
   !KgMermaIMPOST_IE_capa_Verd = sumar_impost_linies(vsqlwhere, rstimpost)
   'Blau
   vTipusImpostLinia = "IMmb"
   vsqlwhere = " WHERE materials.colorreciclatge='Blau' and Parcials_DBL.orcomassignacio='500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and tipusproveidorIMPOST='Importació' and (palets.teimpost=true)"
   !KgMermaIMPOST_IE_capa_Blau = sumar_impost_linies(vsqlwhere, rstimpost)
   'Vermell
   vTipusImpostLinia = "IMmm"
   vsqlwhere = " WHERE materials.colorreciclatge='Vermell' and Parcials_DBL.orcomassignacio='500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and tipusproveidorIMPOST='Importació' and (palets.teimpost=true)"
   !KgMermaIMPOST_IE_capa_Vermell = sumar_impost_linies(vsqlwhere, rstimpost)
   
   !KgMermaIMPOST_IE_capa = !KgMermaIMPOST_IE_capa_Verd + !KgMermaIMPOST_IE_capa_Blau + !KgMermaIMPOST_IE_capa_Vermell
   
   
   'adquirent amb impost
   'Verd
Intracomunitari:
   vTipusImpostLinia = "ICmv"
   vsqlwhere = " WHERE materials.colorreciclatge='Verd' and Parcials_DBL.orcomassignacio='500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and (tipusproveidorIMPOST='Intracomunitari' or tipusproveidorIMPOST is null) and (palets.teimpost=true)"
   !KgMermaIMPOST_AD_capa_verd = sumar_impost_linies(vsqlwhere, rstimpost)
   'Blau
   vTipusImpostLinia = "ICmb"
   vsqlwhere = " WHERE materials.colorreciclatge='Blau' and Parcials_DBL.orcomassignacio='500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and (tipusproveidorIMPOST='Intracomunitari' or tipusproveidorIMPOST is null) and (palets.teimpost=true)"
   !KgMermaIMPOST_AD_capa_blau = sumar_impost_linies(vsqlwhere, rstimpost)
   'Vermell
   vTipusImpostLinia = "ICmm"
   vsqlwhere = " WHERE materials.colorreciclatge='Vermell' and Parcials_DBL.orcomassignacio='500' and Parcials_DBL.comanda='" + atrim(vnumc) + "' and (tipusproveidorIMPOST='Intracomunitari' or tipusproveidorIMPOST is null) and (palets.teimpost=true)"
   !KgMermaIMPOST_AD_capa_vermell = sumar_impost_linies(vsqlwhere, rstimpost)
   
   !KgMermaIMPOST_AD_capa = !KgMermaIMPOST_AD_capa_verd + !KgMermaIMPOST_AD_capa_blau + !KgMermaIMPOST_AD_capa_vermell
   End With
End Sub
Function migelaboratultimaseccio(vnumc As Double) As Byte
   Dim rst As Recordset
   Dim vultimaseccio As String
   migelaboratultimaseccio = 1
   Set rst = dbcomandes.OpenRecordset("SELECT *,productes.ruta FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi where comanda=" + atrim(vnumc) + ";")
   If rst.EOF Then Exit Function
   vultimaseccio = Mid(rst!ruta, Len(rst!ruta), 1)
   If vultimaseccio = "S" Then GoTo fi
   If vultimaseccio = "R" Then
           migelaboratultimaseccio = IIf(atrim(rst!migelaborat) <> "L", 2, 1)
        Else: migelaboratultimaseccio = IIf(atrim(rst!tubolam) <> "L", 2, 1)
   End If
fi:
   Set rst = Nothing
End Function
Function sumar_impost_linies(vsqlwhere As String, rstimpost As Recordset) As Double
   Dim vsuma As Double
   Dim vsumaKgm2 As Double
   Dim rst As Recordset
   Dim vsql As String
   Dim qdf As QueryDef
   Dim vsubconsulta As String
   Dim vnomsubconsulta As String
   
   
   
   'vsql = "SELECT Parcials_DBL.comanda, Palets.Idpalet,Palets.teimpost, Bobines.Idbobina, Parcials_DBL.orcomassignacio, Palets.micres, Palets.semielaborat, Palets.Ample, Parcials_DBL.metres, Bobines.pesdelproveidor, Bobines.Mts, comandes.amplesol, comandes.migelaboratsol, comandes.longitudsol, comandes.solapasol, proveidors.tipusproveidorIMPOST, proveidors.nom, albaransbip.kgbaseimposableimpostenvasos, albaransbip.KgImpostEnvasos, materials.tanpercentimpostenvasos, materials.grmm2, materials.grmcm3, (((((([bobines].[pesdelproveidor]/[bobines].[mts])*[parcials_DBL].[metres])/[SUBCONSULTA_CALCULIMPOST].[pesTbobines])*([materials].[grmcm3])*[palets].[micres])/1000)*[materials].[tanpercentimpostenvasos])/100 AS Kgm2, "
   'vsql = vsql + " ((([materials].[grmcm3]*[palets].[micres])/1000)*[parcials_DBL].[metres]*(([palets].[ample]+([palets].[solapa]/2))/100)*IIf([palets].[semielaborat]<>'L',2,1)) AS KgTmaterial , IIf(Mid([ruta],Len([ruta]))<>'S',0,((([comandes].[longitudsol]+([comandes].[solapasol]/2))/100)*([comandes].[amplesol]/100))*IIf([migelaboratsol]<>'L',2,1)) AS m2perpeça, (((([bobines].[pesdelproveidor]/[bobines].[mts])*[parcials_DBL].[metres])/[SUBCONSULTA_CALCULIMPOST].[pesTbobines]))*100 AS tanper100_mitjana_Grm3 "
   'vsql = vsql + " FROM (((((((Palets LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi) LEFT JOIN albaransbip ON Palets.Idpalet = albaransbip.numpalet) LEFT JOIN Parcials_DBL ON Palets.Idpalet = Parcials_DBL.idpalet) LEFT JOIN Bobines ON (Parcials_DBL.idpalet = Bobines.Idpalet) AND (Parcials_DBL.idbobina = Bobines.Idbobina)) LEFT JOIN comandes ON Parcials_DBL.comandaDBL = comandes.comanda) LEFT JOIN productes ON comandes.producte = productes.codi) LEFT JOIN SUBCONSULTA_CALCULIMPOST ON Parcials_DBL.comanda = SUBCONSULTA_CALCULIMPOST.comanda "
   vsql = ""
   vsql = "SELECT Parcials_DBL.comanda,parcials_dbl.id as iddelparcial, Palets.Idpalet, Palets.teimpost, Bobines.Idbobina, Parcials_DBL.orcomassignacio, Palets.micres, Palets.semielaborat, Palets.Ample, Parcials_DBL.metres, Bobines.pesdelproveidor, Bobines.Mts, comandes.amplesol, comandes.migelaboratsol, comandes.longitudsol, comandes.solapasol, proveidors.tipusproveidorIMPOST, proveidors.nom, materials.tanpercentimpostenvasos, materials.grmm2, materials.grmcm3, (((((([bobines].[pesdelproveidor]/[bobines].[mts])*[parcials_DBL].[metres])/[SUBCONSULTA_CALCULIMPOST].[pesTbobines])*([materials].[grmcm3])*[palets].[micres])/1000)*[materials].[tanpercentimpostenvasos])/100 AS Kgm2, ((([materials].[grmcm3]*[palets].[micres])/1000)*[parcials_DBL].[metres]*(([palets].[ample]+([palets].[solapa]/2))/100)*([materials].[tanpercentimpostenvasos]/100)*IIf([palets].[semielaborat]<>'L',2,1)) AS KgTmaterial,"
   vsql = vsql + " IIf(Mid([ruta],Len([ruta]))<>'S',0,((([comandes].[longitudsol]+([comandes].[solapasol]/2))/100)*([comandes].[amplesol]/100))*IIf([migelaboratsol]<>'L',2,1)) AS m2perpeça,"
   vsql = vsql + "(((([bobines].[pesdelproveidor]/[bobines].[mts])*[parcials_DBL].[metres])/[SUBCONSULTA_CALCULIMPOST].[pesTbobines]))*100 AS tanper100_mitjana_Grm3 "
   vsql = vsql + " FROM ((((((Palets LEFT JOIN materials ON Palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi) LEFT JOIN Parcials_DBL ON Palets.Idpalet = Parcials_DBL.idpalet) LEFT JOIN Bobines ON (Parcials_DBL.idpalet = Bobines.Idpalet) AND (Parcials_DBL.idbobina = Bobines.Idbobina)) LEFT JOIN comandes ON Parcials_DBL.comandaDBL = comandes.comanda) LEFT JOIN productes ON comandes.producte = productes.codi) LEFT JOIN SUBCONSULTA_CALCULIMPOST ON Parcials_DBL.comanda = SUBCONSULTA_CALCULIMPOST.comanda"
   
   vsubconsulta = "SELECT Sum(Bobines.pesdelproveidor) AS SumaDepesdelproveidor, Parcials.comanda, Sum(([bobines].[pesdelproveidor]/[bobines].[mts])*[parcials].[metres]) AS pesTbobines"
   vsubconsulta = vsubconsulta + " FROM Bobines LEFT JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet) "
   vsubconsulta = vsubconsulta + " Where parcials.comanda='" + atrim(rstimpost!comanda) + "' and Parcials.orcomassignacio <> '500' GROUP BY Parcials.comanda;"
   vnomsubconsulta = "SUBCONSULTA__CALCULIMPOST_" + nomordinador
   eliminar_querydef_subconsulta vnomsubconsulta 'elimino primer la subconsulta de la base de dades per no liarla
   vsql = substituir(vsql, "SUBCONSULTA_CALCULIMPOST", vnomsubconsulta)
   Set qdf = dbstocks.CreateQueryDef(vnomsubconsulta, vsubconsulta)
  ' Clipboard.Clear
  ' Clipboard.SetText vsql + vsqlwhere
   Set rst = dbstocks.OpenRecordset(vsql + vsqlwhere + " ORDER BY Parcials_DBL.orcomassignacio DESC;")
   If Not rst.EOF Then
       rstimpost!m2perpeça = cadbl(rst!m2perpeça)
   End If
   vsumaKgm2 = 0
   vsuma = 0
   While Not rst.EOF
      vsuma = vsuma + cadbl(rst!KgTmaterial)
      vsumaKgm2 = vsumaKgm2 + Redondejar(cadbl(rst!kgm2), 6)
      If vTipusImpostLinia <> "" Then dbimpost.Execute "insert into ParcialsBobinesCalculats (idtaulaimpost,idparcial,comanda,kg,tipusimpost) values (" + atrim(rstimpost!ID) + "," + atrim(rst!iddelparcial) + "," + atrim(rst!comanda) + "," + passaradecimalpunt(cadbl(rst!KgTmaterial)) + ",'" + vTipusImpostLinia + "')"
      rst.MoveNext
   Wend
'   If vsumaKgm2 > 0 Then rstimpost!kgm2 = vsumaKgm2
   If rstimpost!kgm2 = 0 Then rstimpost!kgm2 = vsumaKgm2
   sumar_impost_linies = vsuma
   Set rst = Nothing
   eliminar_querydef_subconsulta vnomsubconsulta
        
End Function
Sub eliminar_querydef_subconsulta(vnom)
   Dim i As Long
   For i = 0 To dbstocks.QueryDefs.Count - 1
      If dbstocks.QueryDefs(i).Name = vnom Then dbstocks.QueryDefs.Delete vnom: GoTo fi
   Next i
fi:
End Sub
Private Sub Command2_Click()
   Dim numalb As Double
   Dim vcomandesambclixes As String
   Dim rstc As Recordset
   Dim vhihaliniessenseimpost As Boolean
   
   valbaraSAPportaimpost = False
   If datacapcalera.Recordset.EditMode > 0 Or datalinies.Recordset.EditMode > 0 Then
      MsgBox "S'està editant la capçalera o les linies d'albarà primer accepta els canvis abans d'enviar-ho al SAP.", vbCritical, "A T E N C I Ó"
      Exit Sub
   End If
   If IsDate(datacapcalera.Recordset!dataenvioasap) Then MsgBox "Ja vas enviar aquest albarà a SAP no es pot tornar a fer.", vbCritical, "Error": Exit Sub
   If datacapcalera.Recordset.EOF Then Exit Sub
   
   If etmetrescubicscalculats.ForeColor = QBColor(12) Then MsgBox "Hi ha palets que no tenen base assignada a baixes d'embalar, revisa-ho abans de pujar l'albarà.", vbCritical, "Error": Exit Sub
   If Not datacapcalera.Recordset!papersfrontalsimpresos Then
       If MsgBox("AQUESTA COMANDA NO S'HAN IMPRÈS ELS PAPERS FRONTALS." + vbNewLine + "ES CORRECTE?", vbDefaultButton2 + vbCritical + vbYesNo, "ATENCIÓ") = vbNo Then Exit Sub
   End If
   If Not revisarsishaREVISATelspaletsabansdenviar(cadbl(cnumalbara)) Then If UCase(InputBox("No s'ha verificat els palets abans de fer l'enviament." + vbNewLine + "VOLS CONTINUAR IGUALMENT? ESCRIU [VERIFICAT] PER CONTINUAR.", "ATENCIÓ NO VERIFICAT")) <> "VERIFICAT" Then Exit Sub
      
   If CDbl(datacapcalera.Recordset!id_transport) = 0 Or atrim(datacapcalera.Recordset!id_transport) = "" Then MsgBox "No pots enviar a SAP sense transportista i tipus de ports.", vbCritical, "Error": Exit Sub
   numalb = cadbl(datacapcalera.Recordset!numalbara)
   If Not comprovarcodiclientexisteixasap(datacapcalera.Recordset!codiclient, logoinplacsa.visible) Then MsgBox "Aquest codi de client no existeix a SAP hi ha algun error.", vbCritical, "Error GREU": Exit Sub
   If Not toteslesbobinesdonadesdebaixaocanviunitatpvp(vhihaliniessenseimpost) Then Exit Sub
   If vhihaliniessenseimpost Then If MsgBox("Veig que hi ha alguna comanda sense Impost d'envasos i es un client ESPANYOL, vols comprovar que estigui correcte? " + vbnewlinw + " Potser fent editar sobre la linia i acceptar, actualitzarà l'IMPOST.", vbCritical + vbDefaultButton1 + vbYesNo, "ATENCIÓ") = vbYes Then GoTo fi
   
   If comprovarsiesTotalihihabobinessenseseleccionar Then Exit Sub
   If vcancelarSAP Then GoTo fi
   If MsgBox("Segur que vols enviar aquest albarà al SAP el procés es irreversible.", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
      If checknogenerarfitxersap.Value = 0 Then
         generar_fitxer_sap datacapcalera.Recordset!numalbara
           Else:
             checknogenerarfitxersap.Value = 0
             datacapcalera.Recordset.Edit
             datacapcalera.Recordset!dataenvioasap = Now
             datacapcalera.Recordset.Update
      End If
      datalinies.Refresh
      wait 1
      While Not datalinies.Recordset.EOF
        Set rstc = dbcomandes.OpenRecordset("select linkcomanda1,linkcomanda2,comanda from comandes where comanda=" + atrim(cadbl(datalinies.Recordset!lotinplacsa)))
        passarbobinesaentregades True, cadbl(datalinies.Recordset!lotinplacsa), cadbl(datalinies.Recordset!numalbara), atrim(datacapcalera.Recordset!dataalbara), cadbl(datacapcalera.Recordset!id_transport)
        canviarlestatdelacomanda cadbl(datalinies.Recordset!lotinplacsa), atrim(datalinies.Recordset!tipusdeentrega), datacapcalera.Recordset!dataalbara
        comprovarsihihaclixesperenviar cadbl(datalinies.Recordset!lotinplacsa), vcomandesambclixes, True
        If Not rstc.EOF Then
           comprovarsihihaclixesperenviar cadbl(rstc!linkcomanda1), vcomandesambclixes, True
           comprovarsihihaclixesperenviar cadbl(rstc!linkcomanda2), vcomandesambclixes, True
        End If
        datalinies.Recordset.MoveNext
      Wend
      If ettanper100merma.tag <> "" Then enviar_mermesmassagrans ettanper100merma.tag
   End If
   
   If vcomandesambclixes <> "" Then MsgBox "La comanda " + vcomandesambclixes + " te clixes pendents de facturar i s'han albaranat al SAP " + Chr(10) + "PENSA A FER LA IMPORTACIÓ CORRESPONENT TAN A L'EMPRESA INPLACSA COM PLASEL", vbInformation, "CLIXES ALBARANATS"
   datacapcalera.Refresh
   datacapcalera.Recordset.FindFirst "numalbara=" + atrim(numalb)
   comprovarsicalpackinglist
   comprovar_avisos_del_client datacapcalera.Recordset!id_direnvio, numalb
   comprovar_extracost numalb
   If esalbaradARTA(datacapcalera.Recordset!numalbara) Then generar_CSV_dArta datacapcalera.Recordset!numalbara
fi:
   Set rstc = Nothing
   ratoli "normal"
   
End Sub
Function esalbaradARTA(vnumalbara As String) As Boolean
   Dim rstcli As Recordset
   
   Set rstcli = dbvendes.OpenRecordset("SELECT clients.nom,comandes.client FROM liniesalbara LEFT JOIN (clients RIGHT JOIN comandes ON clients.codi = comandes.client) ON liniesalbara.lotinplacsa = comandes.comanda WHERE (((liniesalbara.numalbara)=" + vnumalbara + "));")
   If Not rstcli.EOF Then If rstcli!client = 6393 Or (InStr(1, rstcli!nom, "DARTA,") > 0) Or (InStr(1, rstcli!nom, "DARTA ") > 0) Then esalbaradARTA = True
   
   Set rstcli = Nothing
End Function
Sub comprovar_extracost(vnumalb As Double)
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim vmsg As String
  Set rst = dbvendes.OpenRecordset("select lotinplacsa from liniesalbara where numalbara=" + atrim(vnumalb))
  While Not rst.EOF
    Set rst2 = dbcomandes.OpenRecordset("select * from comandes_observaciopvp where extracost>0 and comanda=" + atrim(rst!lotinplacsa))
    If rst2.EOF Then Set rst2 = dbcomandes.OpenRecordset("select * from comandes_observaciopvp where extracost>0 and  comandesafectades like '*" + atrim(rst!lotinplacsa) + "*'")
    If Not rst2.EOF Then
      If InStr(1, vmsg, atrim(rst2!comanda)) = 0 Then
       vmsg = "La comanda " + atrim(rst2!comanda) + " te un Extra Cost de " + atrim(rst2!extracost) + "€"
       vmsg = vmsg + vbNewLine + " La descripció de la linia es: " + atrim(rst2!liniaalbara)
       vmsg = vmsg + vbNewLine + " Aquest extracost es reparteix entre les seguents comandes: " + vbNewLine + "      " + atrim(rst2!comandesafectades) + vbNewLine + vbNewLine
      End If
    End If
    rst.MoveNext
  Wend
  If vmsg <> "" Then MsgBox vmsg
  Set rst2 = Nothing
  Set rst = Nothing
End Sub
Function buscar_pressupostos_del_albara(vnumalb As Double, vnumlots As String) As String
   Dim rst As Recordset
   Dim vnumpressupostos As String
   
   Set rst = dbcomandes.OpenRecordset("select lotinplacsa from liniesalbara where numalbara=" + atrim(vnumalb))
   While Not rst.EOF
     vnumlots = vnumlots + IIf(vnumlots <> "", ",", "") + atrim(rst!lotinplacsa)
     rst.MoveNext
   Wend
   If vnumlots = "" Then GoTo fi
   Set rst = dbcomandes.OpenRecordset("select distinct numpressupost from comandes where comanda in(" + atrim(vnumlots) + ")")
   While Not rst.EOF
      vnumpressupostos = vnumpressupostos + IIf(vnumpressupostos <> "", ",", "") + atrim(rst!numpressupost)
      rst.MoveNext
   Wend
   If vnumpressupostos <> "" Then buscar_pressupostos_del_albara = vnumpressupostos
fi:
   Set rst = Nothing
End Function
Sub comprovar_avisos_del_client(viddirenvio As Double, vnumalb As Double)
  Dim rst As Recordset
  Dim vnumpressupostos As String
  Dim vnumlots As String
  Set rst = dbcomandes.OpenRecordset("select avisfiproduccio,avisalbaragenerat,codi,nome from clients_envios where id=" + atrim(viddirenvio))
  If Not rst.EOF Then
    If atrim(rst!avisfiproduccio) <> "" Then MsgBox "Hi ha un missatge de fi de producció." + vbNewLine + vbNewLine + atrim(rst!avisfiproduccio), vbExclamation, "ATENCIÓ"
    If atrim(rst!avisalbaragenerat) <> "" Then
        If InStr(1, rst!avisalbaragenerat, "@") > 0 And InStr(1, rst!avisalbaragenerat, ".") > 0 Then
             vnumpressupostos = buscar_pressupostos_del_albara(vnumalb, vnumlots)
             enviaremailgeneric rst!avisalbaragenerat, "S'ha generat un albarà al SAP del client " + atrim(rst!codi) + "-" + atrim(rst!nome) + " " + atrim(Now), " S'ha generat albarà dels següents pressupostos: " + vnumpressupostos + Chr(10) + Chr(13) + " Les comandes relacionades son: " + vnumlots
             'enviaremailgeneric "miquel.inplacsa@gmail.com", "S'ha generat un albarà al SAP del client " + atrim(rst!codi) + "-" + atrim(rst!nome), " S'ha generat albarà dels següents pressupostos: " + vnumpressupostos + vbNewLine + " Les comandes relacionades son: " + vnumlots
        End If
    End If
  End If
  Set rst = Nothing
End Sub
Sub comprovarsicalpackinglist()
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("select packinglistalbara from clients_envios where id=" + atrim(formvendes.datacapcalera.Recordset!id_direnvio))
  If Not rst.EOF Then
     If atrim(rst!packinglistalbara) = "Cap" Then etavis = ""
     If atrim(rst!packinglistalbara) = "Detal Bobina per Bobina" Then etavis = "" ': cdetallbobines.Value = 1
     If atrim(rst!packinglistalbara) = "Totalitzat" Then etavis = "" ': cdetallbobines.Value = 0
  End If
  If etavis <> "" Then MsgBox etavis, vbInformation, "Informació de Packing-List"
 Set rst = Nothing
End Sub
Sub comprovarsihihaclixesperenviar(numc As Double, vcomandesambclixes As String, Optional valbaranar As Boolean, Optional vtreball As Double, Optional vordre As Byte)
   Dim rst As Recordset
   'Dim vordre As Byte
   'Dim vtreball As Double
   Dim rstpressupost As Recordset
   Dim rstm As Recordset
   Dim vnomempresafacturadora As String
   If cadbl(vtreball) > 0 Then GoTo sensecomanda
   If InStr(1, rutadelproducte(numc), "I") = 0 Then Exit Sub
   Set rst = dbcomandes.OpenRecordset("SELECT numtreball,numordremodificacio FROM comandes   WHERE comanda=" + atrim(numc) + ";")
   If rst.EOF Then GoTo fi
   vtreball = cadbl(rst!numtreball)
   vordre = cadbl(rst!numordremodificacio)
sensecomanda:
   Set rstm = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(vtreball) + " and ordre=" + atrim(vordre))
   If Not rstm.EOF Then vnomempresafacturadora = IIf(atrim(rstm!empresafacturadora) = "P", "(Plasel)", IIf(atrim(rstm!empresafacturadora) = "I", "(Inplacsa)", ""))
   Set rst = dbclixes.OpenRecordset("select * from clixes_albarans where id_treball=" + atrim(vtreball) + " and ordremodificacio=" + atrim(vordre) + " and not facturat order by ordre")
   If Not rst.EOF Then
    Set rstpressupost = dbclixes.OpenRecordset("select * from pressupostos where (lotambelqueshafacturat=null or lotambelqueshafacturat=0) and preu>0 and  id_treball=" + atrim(vtreball) + " and ordremodificacio=" + atrim(vordre))
      If Not rstpressupost.EOF Then
          idiomaclientclixes = convertir_idiomaclixes(rstpressupost!Idioma)
          'If rstpressupost!enviat Then 'hauria de ser verificat però a vegades no hi ha verificació
              If valbaranar Then albaranar_albaranspendentsdeltreball numc, vtreball, vordre, rst, rstpressupost
              vcomandesambclixes = vcomandesambclixes + " " + atrim(numc) + " " + vnomempresafacturadora
          'End If
    End If
   End If
fi:
   Set rst = Nothing
End Sub
Function convertir_idiomaclixes(vidioma As String) As String
   Select Case UCase(vidioma)
      Case "ESP"
         convertir_idiomaclixes = "ES"
      Case "CAT"
         convertir_idiomaclixes = "ES"
      Case "ANG"
         convertir_idiomaclixes = "EN"
      Case "FRA"
         convertir_idiomaclixes = "FR"
   End Select
End Function
Function noexisteixelclientasap(vclient As Double, vempresa As String) As Boolean
   Dim rst As Recordset
   vempresa = UCase(vempresa)
   If vempresa = "I" Then
      vempresa = ""
        Else: vempresa = "PLASEL"
   End If
   Set rst = dbcomandes.OpenRecordset("select * from clients_codisSAP" + vempresa + " where codiSAP=" + atrim(vclient))
   If rst.EOF Then noexisteixelclientasap = True
   Set rst = Nothing
End Function
Sub albaranar_albaranspendentsdeltreball(numc As Double, vtreball As Double, vordre As Byte, rst As Recordset, rstpressupost As Recordset)
   Dim rstm As Recordset
   Set rstm = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(vtreball) + " and ordre=" + atrim(vordre))
   If rstm.EOF Then Exit Sub
   If cadbl(rstm!codiclientfactclixes) = 0 Or atrim(rstm!empresafacturadora) = "" Then
      MsgBox "No hi ha el client a qui facturar-li els clixes assignat al treball  " + atrim(vtreball) + " de producció." + Chr(10) + " S'enviarà un e-mail a Disseny per donar-lo d'alta, prova-ho en un altra moment."
      enviaremailgeneric "mkinplacsa@inplacsa.com", "Albarans clixes. Client per facturar-li no assignat al treball " + atrim(vtreball), "No hi ha el client a qui facturar-li els clixes assignat al treball  " + atrim(vtreball) + " de producció."
      Exit Sub
   End If
   If noexisteixelclientasap(cadbl(rstm!codiclientfactclixes), atrim(rstm!empresafacturadora)) Then
      MsgBox "No s'ha trobat el client " + atrim(cadbl(rstm!codiclientfactclixes)) + " a a l'empresa " + IIf(UCase(atrim(rstm!empresafacturadora)) = "P", "PLASEL", "INPLACSA") + " del SAP." + Chr(10) + " S'enviarà un e-mail a Comptablitat per donar-lo d'alta, prova-ho en un altra moment."
      enviaremailgeneric "incidenciesillistatsSAPcomptabilitat", "Albarans clixes. Client no existent a " + IIf(UCase(atrim(rstm!empresafacturadora)) = "P", "PLASEL", "INPLACSA"), "El client " + atrim(rstm!codiclientfactclixes) + " no existeix a l'empresa " + IIf(UCase(atrim(rstm!empresafacturadora)) = "P", "PLASEL", "INPLACSA") + " s'hauria de donar d'alta el mes ràpid possible."
      Exit Sub
   End If
   If cadbl(rstm!codiclientfactclixes) = 0 Then
      MsgBox "El treball " + atrim(vtreball) + " no te assignat el client de facturació de Clixes, no es pot albaranar." + Chr(10) + "ES PASSA UN MAIL A DISSENY PER REVISAR-HO" + Chr(10) + "Es pot facturar igualment el material però els clixes no.", vbCritical, "Error"
      enviaremailsensecodiclientaclixes numc, vtreball, vordre
      Exit Sub
   End If
 'fer la exportacio
   generar_fitxer_sap_clixes rstm, numc, rstpressupost
   If numc <> -1 Then
         passar_linies_albara_clixe_a_facturades vtreball, vordre, numc 'el numc -1 controla si s'ha generat el fitxer SAP
         enviaremaildeclixefacturat
   End If
   Set rstm = Nothing
    
End Sub
Sub enviaremailsensecodiclientaclixes(numc As Double, vtreball As Double, vordre As Byte)
   Dim dbenvio As Database
   Set dbenvio = OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
   dbenvio.Execute "insert into envios_mails (data,destinatari,assumpte,cos) values (now,'incidenciesalbaranarclixes','Error albarà clixes " + atrim(vtreball) + "/" + atrim(vordre) + "  NºLot: " + atrim(numc) + "','No hi ha client vàlid assignat a albarans del clixé, s´ha de revisar per poder albaranar aquests clixes al client.')"
   Set dbenvio = Nothing
End Sub
Sub enviaremailgeneric(destinatari As String, assumpte As String, cos As String)
   Dim dbenvio As Database
   If atrim(cos) = "" Then Exit Sub
   Set dbenvio = OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
   dbenvio.Execute "insert into envios_mails (data,destinatari,assumpte,cos) values (now,'" + destinatari + "','" + treuresimbols(assumpte) + "','" + treuresimbols(cos) + "')"
   Set dbenvio = Nothing
End Sub
Sub enviaremaildeclixefacturat()
   Dim dbenvio As Database
   Dim vcos1 As String
   Dim vcos2 As String
   If atrim(vdesccosmailclixes) = "" Then Exit Sub
   Set dbenvio = OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
   vdesccosmailclixes = atrim(vdesccosmailclixes)
   vcos1 = Mid(vdesccosmailclixes, 1, 255)
   If Len(vdescosmailclixes > 255) Then vcos2 = Mid(vdesccosmailclixes, 256, 510)
   dbenvio.Execute "insert into envios_mails (data,destinatari,assumpte,cos,cos2) values (now,'incidenciesalbaranarclixes','Clixe albaranat desde Expedicions','" + treure_apostruf(vcos1) + "','" + treure_apostruf(vcos2) + "')"
   Set dbenvio = Nothing
End Sub
Sub passar_linies_albara_clixe_a_facturades(vtreball As Double, vordre As Byte, numc As Double)
    dbclixes.Execute "update clixes_albarans set facturat=true,lotambelqueshafacturat=" + atrim(numc) + " where id_treball=" + atrim(vtreball) + " and ordremodificacio=" + atrim(vordre) + " and not facturat"
    dbclixes.Execute "update pressupostos set datafacturacio=now,lotambelqueshafacturat=" + atrim(numc) + " where id_treball=" + atrim(vtreball) + " and ordremodificacio=" + atrim(vordre)
End Sub
Sub canviarlestatdelacomanda(numc As Double, estat As String, vdataentrega As Date)
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(numc))
   If rst.EOF Then Exit Sub
   If estat = "" Then MsgBox "La comanda " + atrim(numc) + " no tenia possat l'estat final de la comanda, per defecte es passarà a (T).", vbInformation, "Atenció"
   If estat = "T" Or estat = "P" Then
      dbcomandes.Execute "update comandes set proximaseccio='" + estat + "' where comanda=" + atrim(numc) + IIf(estat = "T", " or comanda=" + atrim(cadbl(rst!linkcomanda1)) + " or comanda=" + atrim(cadbl(rst!linkcomanda2)), "")
      actualitzar_totals_a_comandes numc
      dbcomandes.Execute "update comandes_extres set dataentrega=#" + atrim(Format(vdataentrega, "mm/dd/yy")) + "# where comanda=" + atrim(numc)
      If estat = "T" Then
          dbbaixes.Execute "update impressorestot set acavada='1' where comanda=" + atrim(numc)
          dbbaixes.Execute "update laminadorestot set acavada='1' where comanda=" + atrim(numc)
          dbbaixes.Execute "update rebobinadorestot set acavada='1' where comanda=" + atrim(numc)
          dbbaixes.Execute "update soldadorestot set acavada='1' where comanda=" + atrim(numc)
      End If
   End If
End Sub
Sub enviar_mermesmassagrans(vmsg As String)
    enviaremailgeneric "jmiralles@inplacsa.com", "[%] de mermes mes gran de 17% igual a zero i negatives. " + atrim(Now), "Els seguents lots han superat la merma esperada, s'han de revisar: " + vbNewLine + vmsg
End Sub
Sub actualitzar_totals_a_comandes(numc As Double)
   Dim vsql As String
   Dim rst As Recordset
   Set rst = datacapcalera.Database.OpenRecordset("select sum(quantitat*preuvenda) as tpvp,sum(kgtotalsbruts) as tkilos,sum(metreslineals) as tmetres,sum(unitats) as tunitats from liniesalbara where lotinplacsa=" + atrim(numc))
   vsql = "update comandes_extres set metresentregats=" + atrim(cadbl(rst!tmetres))
   vsql = vsql + ", numpecesentregades=" + atrim(cadbl(rst!tunitats))
   vsql = vsql + ", kilosentregats=" + passaradecimalpunt(cadbl(rst!tkilos))
   vsql = vsql + ",pvptotal=" + passaradecimalpunt(Redondejar(cadbl(rst!tpvp), 2))
   vsql = vsql + " where comanda=" + atrim(numc)
   dbcomandes.Execute vsql
   Set rst = Nothing
End Sub
Function toteslesbobinesdonadesdebaixaocanviunitatpvp(vhihaliniessenseimpost As Boolean) As Boolean
  Dim bobassignades As Boolean
  Dim canviunitatpvp As String
  Dim rstc As Recordset
  Dim vpvp As Double
  Dim vEsEspanyol As Boolean
  
  datalinies.Refresh
  toteslesbobinesdonadesdebaixaocanviunitatpvp = True
  vEsEspanyol = elclientesESPANYOL(datacapcalera.Recordset!numalbara)
  canviunitatpvp = ""
  While Not datalinies.Recordset.EOF
    bobassignades = sihihanbobinesassignades(cadbl(datalinies.Recordset!lotinplacsa), atrim(datalinies.Recordset!tipusdeentrega))
    If bobassignades Then toteslesbobinesdonadesdebaixaocanviunitatpvp = False
    Set rstc = dbcomandes.OpenRecordset("SELECT pvp,pvpdolar, mesurapvp FROM comandes Where comanda = " + atrim(datalinies.Recordset!lotinplacsa))
    If Not rstc.EOF Then
        vpvp = IIf(datacapcalera.Recordset!moneda = "Euros", cadbl(rstc!pvp), cadbl(rstc!pvpdolar))
        If cadbl(vpvp) <> cadbl(datalinies.Recordset!preuvenda) Then
          datalinies.Database.Execute "update liniesalbara set preuvenda=" + passaradecimalpunt(atrim(vpvp)) + " where id=" + atrim(datalinies.Recordset!ID)
        End If
        If atrim(rstc!mesurapvp) <> atrim(datalinies.Recordset!unitatpvp) And cadbl(datalinies.Recordset!unitatpvp) <> 0 Then canviunitatpvp = canviunitatpvp + " " + atrim(datalinies.Recordset!lotinplacsa)
    End If
    If vEsEspanyol Then
        If cadbl(datalinies.Recordset!kgimpostenvasos) = 0 Then
          If ImpostEnv_regimfiscalREFINPLACSA(datalinies.Recordset!codiproducte) = "" Then vhihaliniessenseimpost = True
        End If
    End If
    datalinies.Recordset.MoveNext
  Wend
  If atrim(canviunitatpvp) <> "" Then MsgBox "Hi ha canvis de mesura de PVP a la comanda " + canviunitatpvp + " corretgeix-ho.", vbCritical, "Error": toteslesbobinesdonadesdebaixaocanviunitatpvp = False
End Function
Sub passarbobinesaentregades(ventregades As Boolean, numc As Double, vnumalbara As Double, vdata As String, vnumtransport As Long)
  Dim rst As Recordset
  Set rst = dbbaixes.OpenRecordset("select * from bobinesent where comanda=" + atrim(numc) + " and (numalbara=" + atrim(cadbl(vnumalbara)) + " or numalbara=null or numalbara=0) order by numbob")
  If Not rst.EOF Then rst.MoveLast: rst.MoveFirst
  While Not rst.EOF
   If cadbl(rst!numalbara) = vnumalbara Then
    If ventregades Then
        'If atrim(rst!Data) = "" Then
          rst.Edit
          rst!Data = CVDate(vdata)
          rst!entregat = "S"
          rst!Transportista = vnumtransport
          rst!numalbara = atrim(vnumalbara)
          rst.Update
        'End If
         Else
           dbbaixes.Execute "update bobinesent set  dataentrega=null,data=null,transportista=0,entregat='N' where comanda=" + atrim(numc) + " and numalbara=" + atrim(cadbl(vnumalbara))
           dbcomandes.Execute "update comandes set proximaseccio='V' where comanda=" + atrim(numc)
           GoTo fi
     End If
   End If
   rst.MoveNext
  Wend
fi:
  Set rst = Nothing
  
End Sub
Sub generar_fitxer_sap_clixes(rstm As Recordset, numc As Double, rstpressupost As Recordset)
   Dim rstlinalb As Recordset
   Dim rstalb As Recordset
   Dim nomfitxer As String
   Dim numid As Double
   Dim vempresa As String
   Dim vnomfitxerseidor As String
   Dim r As String
   vempresa = IIf(rstm!empresafacturadora = "P", "Plasel", "Inplacsa")
   If llegir_ini("Vendes", "rutasap_" + UCase(vempresa), "comandes.ini") = "{[}]" Then MsgBox "No hi ha la ruta d'importació del SAP " + "rutasap_" + UCase(atrim(vempresa)): Exit Sub
   nomfitxer = "V-Clixes_" + atrim(rstm!codiclientfactclixes) + "_" + atrim(numc) + ".csv"
   vnomfitxerseidor = nomfitxer
   r = llegir_ini("Vendes", "rutasap_" + UCase(atrim(vempresa)), "comandes.ini") + "\" + atrim(nomfitxer)
   vnomfitxerseidor = llegir_ini("Vendes", "rutaSapSeidor_" + UCase(atrim(vempresa)), "comandes.ini") + "\" + atrim(vnomfitxerseidor)
   If existeix(r) Then
         r = substituir(r, ".csv", "_" + Format(Now, "hhnnss")) + ".csv"
         vnomfitxerseidor = substituir(vnomfitxerseidor, ".csv", "_" + Format(Now, "hhnnss")) + ".csv"
   End If
   If Not existeix(r) Then
              Open r For Output As 1
              generar_capcalera_clixes_fitxer_sap rstm, vempresa
          Else: MsgBox "No es pot crear el fitxer " + vbNewLine + r + " Potser està pendent de pujar a SAP.", vbCritical, "Error": numc = -1: Exit Sub
   End If
   vidliniaclixe = buscar_idlinia + 1
   generar_linia_fitxer_sap_clixes rstm, numc, rstpressupost
   Close 1
   If numc = -1 Then
      If existeix(r) Then Kill r
     Else: copiar_fitxersap_seidor r, vnomfitxerseidor
   End If
End Sub
Function buscar_idlinia() As Double
   Dim linia As String
   buscar_idlinia = 0
   While Not EOF(1)
      Input #1, linia
   Wend
   If Mid(linia, 1, 1) = "#" Then
      buscar_idlinia = cadbl(Mid(linia, 2, InStr(1, linia, ".") - 2))
   End If
End Function
Sub generar_linia_fitxer_sap_clixes(rstm As Recordset, numc As Double, rstpressupost As Recordset)
   Dim r As String
   Dim rstc As Recordset
   Dim rstclixe As Recordset
   Dim linia As String
   Dim vl As String
   Dim vpreu As Double
   Dim vdescpressupost As String
   Dim vcomandaireferencia As String
   If numc = 0 Then numc = InputBox("Al facturar els clixes independentment no hi ha cap comanda relacionada, si vols possar una comanda així quedarà reflexada la Referencia.", "Comanda relacionada si vols.")
   Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
   Set rstclixe = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(rstm!id_treball))
   If rstc.EOF Or rstclixe.EOF Then Exit Sub
   If numc > 0 Then If rstc!numtreball <> rstm!id_treball Then MsgBox "La comanda entrada no té el mateix treball que el clixé que estàs facturant.", vbCritical, "Error": numc = -1: GoTo fi
   If Not rstpressupost.EOF Then
     vpreu = cadbl(rstpressupost!preu)
     vdescpressupost = treuresimbols(atrim(rstpressupost!descripcio))
   End If
   With rstm
    linia = "#" + atrim(vidliniaclixe) + ".1;" + treuresimbols("PLATES") + ";€/U;FOTOCOMPOSICION Y FOTOPOLIMEROS;1;" + treurecomaperpunt(atrim(vpreu))
    Print #1, linia
    linia = "#" + atrim(vidliniaclixe) + ".2;;;;"
    Print #1, linia
    linia = "#" + atrim(vidliniaclixe) + ".3;" + treuresimbols(atrim(rstclixe!marca) + " - " + atrim(rstclixe!linia))
    Print #1, linia
    linia = "#" + atrim(vidliniaclixe) + ".4;" + treuresimbols(atrim(rstc!comandaclient)) + ";" + treuresimbols(atrim(rstc!refclient))
    Print #1, linia
    linia = "#" + atrim(vidliniaclixe) + ".5;" + treuresimbols(atrim(rstc!obspedgen2)) + ";" + treuresimbols(atrim(rstc!refclientdeclient)) + ";;" + treuresimbols(atrim(rstclixe!codidebarres))
    Print #1, linia
    linia = "#" + atrim(vidliniaclixe) + ".6;;"
    Print #1, linia
    linia = "#" + atrim(vidliniaclixe) + ".7;0;0;0;0;" + treuresimbols(atrim(numc)) + ";CLIXES"
    Print #1, linia
    linia = "#" + atrim(vidliniaclixe) + ".8;"
    Print #1, linia
    
    'POSO LES VISUALS
    vcomandaireferencia = ""
    If (atrim(rstc!comandaclient) + atrim(rstc!refclient)) <> "" Then
         vcomandaireferencia = IIf(atrim(rstc!comandaclient) <> "", traducciodeabreviatures("COM: ", idiomaclientclixes) + treuresimbols(atrim(rstc!comandaclient)), "") + "  " + IIf(atrim(rstc!refclient) <> "", "   REF: " + treuresimbols(atrim(rstc!refclient)), "")
    End If
    If atrim(rstm!codifacturacioclixes) <> "" And atrim(rstm!codifacturacioclixes) <> "0" Then vcomandaireferencia = "Com.Clixé: " + atrim(rstm!codifacturacioclixes) + "  " + vcomandaireferencia
    vl = traducciodeabreviatures("AvClixes", idiomaclientclixes)
    Print #1, "LV" + atrim(vidliniaclixe) + ".1;" + vl
    vdesccosmailclixes = vdesccosmailclixes + vl + Chr(10) + Chr(13)
    vl = treuresimbols(atrim(rstclixe!marca) + " - " + atrim(rstclixe!linia))
    Print #1, "LV" + atrim(vidliniaclixe) + ".2;" + vl
    vdesccosmailclixes = vdesccosmailclixes + vl + Chr(10) + Chr(13)
    vl = treuresimbols(vcomandaireferencia) + "  " + traducciodeabreviatures("AvLot", idiomaclientclixes) + ": " + atrim(numc) + "   T: " + atrim(rstm!id_treball) + "/" + atrim(rstm!ordre)
    Print #1, "LV" + atrim(vidliniaclixe) + ".3;" + vl
    vdesccosmailclixes = vdesccosmailclixes + vl + Chr(10) + Chr(13)
    If atrim(rstm!observacionsfacturaclixes) <> "" Then
      vl = treuresimbols(rstm!observacionsfacturaclixes)
      Print #1, "LV" + atrim(vidliniaclixe) + ".4;" + vl
      vdesccosmailclixes = vdesccosmailclixes + vl + Chr(10) + Chr(13)
    End If
    If cadbl(rstm!codiclientfactclixes) = 43000007419# Then   'si es de ARDO FOODS BELGICA POSSAR COLETILLA
        vl = "This cost is related to ARDO FOODS SLU C2050."
        Print #1, "LV" + atrim(vidliniaclixe) + ".5;" + vl
        vdesccosmailclixes = vdesccosmailclixes + vl + Chr(10) + Chr(13)
        vl = "The management of the design: pdf validation, plate order, etc. has been done"
        Print #1, "LV" + atrim(vidliniaclixe) + ".6;" + vl
        vdesccosmailclixes = vdesccosmailclixes + vl + Chr(10) + Chr(13)
        vl = "directly with Mr. Bryan Germán in Ardo Benimodo (Bryan.German@ardo.com)."
        Print #1, "LV" + atrim(vidliniaclixe) + ".7;" + vl
        vdesccosmailclixes = vdesccosmailclixes + vl + Chr(10) + Chr(13)
    End If
    
    vdesccosmailclixes = vdesccosmailclixes + Chr(10) + Chr(13) + "PRESSUPOST" + Chr(10) + Chr(13)
    vdesccosmailclixes = vdesccosmailclixes + atrim(vdescpressupost) + "  ---->   Preu: " + atrim(vpreu) + " Euros"
    
    End With
fi:
    Set rstclixe = Nothing
    Set rstc = Nothing
    Exit Sub
errorgravar:
      MsgBox err.Description & Chr(10) & "No s'ha gravat el fitxer: " + Chr(10) + r
End Sub


Sub generar_fitxer_sap(numalbp As Double)
   Dim rstlinalb As Recordset
   Dim rstalb As Recordset
   Dim nomfitxer As String
   Dim numid As Double
   Dim idlinia As Long
   Dim vnomfitxerseidor As String
   Dim r As String
    Set rstalb = datacapcalera.Database.OpenRecordset("select * from capcaleraalbara where numalbara=" + atrim(numalbp))
   If Not rstalb.EOF Then
        If llegir_ini("Vendes", "rutasap_" + UCase(atrim(rstalb!empresa)), "comandes.ini") = "{[}]" Then MsgBox "No hi ha la ruta d'importació del SAP " + "rutasap_" + UCase(atrim(rstalb!empresa)): Exit Sub
         nomfitxer = "V-" + Format(cadbl(rstalb!numalbara), "0000000") + "-" + atrim(rstalb!codiclient) + ".csv"
         r = llegir_ini("Vendes", "rutasap_" + UCase(atrim(rstalb!empresa)), "comandes.ini")
         r = r + "\" + atrim(nomfitxer)
         vnomfitxerseidor = llegir_ini("Vendes", "rutaSapSeidor_" + UCase(atrim(rstalb!empresa)), "comandes.ini")
         vnomfitxerseidor = vnomfitxerseidor + "\" + atrim(nomfitxer)
         If existeix(r) Then Kill r
         If existeix(vnomfitxerseidor) Then Kill vnomfitxerseidor
          Else: Exit Sub
   End If
   Set rstlinalb = datacapcalera.Database.OpenRecordset("select * from liniesalbara where numalbara=" + atrim(numalbp))
   idlinia = 1
   While Not rstlinalb.EOF
      vimpostinclosalPVP = False
      datalinies.Recordset.FindFirst "id=" + atrim(rstlinalb!ID)
      generar_linia_fitxer_sap rstlinalb, rstalb, r, idlinia
      If elclientesESPANYOL(rstalb!numalbara) And Not vimpostinclosalPVP Then
          'si es espanyol te que surtir encara que sigui zero
          'If cadbl(rstlinalb!kgimpostenvasos) > 0 Then
            idlinia = idlinia + 1
            generarLINIAimpost rstlinalb, r, idlinia, ImpostEnv_regimfiscalREFINPLACSA(rstlinalb!codiproducte)
          'End If
          Else: If cadbl(rstlinalb!kgimpostenvasos) > 0 Then valbaraSAPportaimpost = True
      End If
      rstlinalb.MoveNext
      idlinia = idlinia + 1
   Wend
   possar_peusalbara numalbp, r
   rstalb.Edit
   rstalb!dataenvioasap = Now
   rstalb.Update
   copiar_fitxersap_seidor r, vnomfitxerseidor
   'If valbaraSAPportaimpost Then MsgBox "Atenció aquest albarà porta impost, has de fer copia per en Miralles i posar-hi el valor de l'impost a cada comanda.", vbExclamation, "Atenció"
End Sub
Sub generarLINIAimpost(rstlinalb As Recordset, vnomfitxer As String, idlinia As Long, vRegimFiscal As String)
    Dim vkgimpost As Double
    Dim v€kgimpost As Double
    Dim vdescripcioimpost1 As String
    Dim vdescripcioimpost2 As String
    Dim vimpostinclosPVPlocal As Boolean
    If Not existeix(vnomfitxer) Then Exit Sub
    With rstlinalb
    vkgimpost = Redondejar(cadbl(!kgimpostenvasos), 0)
    v€kgimpost = cadbl(!eurokg_impost)
    If vRegimFiscal <> "" Then vkgimpost = 0
    vimpostinclosPVPlocal = carregar_camp_PVPinclos(rstlinalb!lotinplacsa)
    Open vnomfitxer For Append As 1
    
    
    'Si idioma es català ho tradueixo --- qualsevol altra idioma es fa en ESPAÑOL tot i que ara nomes es a espanya
    If idiomaclient = "CA" Then
        ' vdescripcioimpost1 = atrim(Format(vkgimpost, "#,##0.000")) + " Kg Base imposable (" + atrim(!eurokg_impost) + "€/Kg)"
         vdescripcioimpost1 = "Kg Base imposable (" + atrim(!eurokg_impost) + "€/Kg)" + IIf(vimpostinclosPVPlocal, "Preu INCLÒS EN EL PVP", "")
         vdescripcioimpost2 = "   segons Llei 7/2022 d'envasos no reutilitzables."
         If vRegimFiscal <> "" Then vdescripcioimpost2 = "Llei 7/2022 exempt per règim fiscal (Lletra " + vRegimFiscal + ")"
        Else:
           'vdescripcioimpost1 = atrim(Format(vkgimpost, "#,##0.000")) + " Kg Base imponible (" + atrim(!eurokg_impost) + "€/Kg)"
           vdescripcioimpost1 = "Kg Base imponible (" + atrim(!eurokg_impost) + "€/Kg) " + IIf(vimpostinclosPVPlocal, "Precio INCLUIDO EN PVP", "")
           vdescripcioimpost2 = "   según Ley 7/2022 de envases no reutilizables."
           If vRegimFiscal <> "" Then vdescripcioimpost2 = "Ley 7/2022 exento por régimen fiscal (Letra " + vRegimFiscal + ")"
    End If
    If vimpostinclosPVPlocal Then v€kgimpost = 0
    linia = "#" + atrim(idlinia) + ".1;IMP_ENV;€/KG;" + vdescripcioimpost1 + ";" + treurecomaperpunt(atrim(vkgimpost)) + ";" + treurecomaperpunt(atrim(v€kgimpost))
    Print #1, linia
    linia = "#" + atrim(idlinia) + ".7;" + ";" + ";" + ";" + ";" + treuresimbols(atrim(!lotinplacsa)) + ";"
    Print #1, linia
    Print #1, "LV" + atrim(idlinia) + ".2;" + vdescripcioimpost2
    Print #1, "LV" + atrim(idlinia) + ".3; "
    End With
    Close 1
    Exit Sub
errorgravar:
      MsgBox err.Description & Chr(10) & "No s'ha gravat el fitxer: " + Chr(10) + r
End Sub
Function elclientesESPANYOL(vnumalb As Double) As Boolean
  If paisdedirecciodenviament(vnumalb) = "ES" Then elclientesESPANYOL = True
End Function
Function paisdedirecciodenviament(vnumalb As Double) As String
  Dim rst As Recordset
  Set rst = datacapcalera.Database.OpenRecordset("SELECT capcaleraalbara.numalbara, Clients_envios.pais FROM capcaleraalbara LEFT JOIN Clients_envios ON capcaleraalbara.id_direnvio = Clients_envios.id where numalbara=" + atrim(vnumalb))
  If Not rst.EOF Then paisdedirecciodenviament = atrim(rst!pais)
  
  If paisdedirecciodenviament = "" Then MsgBox "Aquest client no té el país posat a la direcció d'enviament.", vbCritical, "Error"
  Set rst = Nothing
End Function
Sub copiar_fitxersap_seidor(vnomfitxer As String, vnomfitxerseidor As String)
   If existeix(vnomfitxerseidor) Then Kill vnomfitxerseidor
   Copiar_Fitxer vnomfitxer, vnomfitxerseidor
End Sub
Private Sub Command3_Click()
  escullir_comandaxralbaranar True
End Sub
Function substituir(cadena As String, buscar As String, canviar As String) As String
   If buscar = canviar Then GoTo fi
   While InStr(1, cadena, buscar) > 0
    comença = InStr(1, cadena, buscar) - 1
    If comença < 1 Then substituir = cadena: Exit Function
    acaba = comença + Len(buscar) + 1
    cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   Wend
fi:
   substituir = cadena
   'MsgBox linia
End Function
Function treurecomaperpunt(desc As String) As String
   desc = substituir(desc, ",", ".")
   treurecomaperpunt = desc
End Function
Function treuresimbols(desc As String) As String
   desc = substituir(desc, ":", "_")
   desc = substituir(desc, "'", "´")
   desc = substituir(desc, "|", "_")
   desc = substituir(desc, ";", "_")
   desc = substituir(desc, Chr(10), "")
   desc = substituir(desc, Chr(13), "")
   treuresimbols = desc
End Function
Sub generar_capcalera_fitxer_sap(rstalb As Recordset)
    Dim rst As Recordset
    With rstalb
    Set rst = dbcomandes.OpenRecordset("select * from clients_envios where id=" + atrim(!id_direnvio))
    linia = atrim(!empresa) + ";" + atrim(!numalbara) + ";" + Format(!dataalbara, "dd/mm/yy") + ";" + atrim(!codiclient) + ";" + atrim(!id_direnvio) + ";" + atrim(!id_transport) + ";" + Mid(atrim(!tipusports), 1, 1) + ";" + treuresimbols(atrim(!observacionsports)) + ";" + treuresimbols(atrim(!observacions))
    If Not rst.EOF Then
       linia = linia + ";" + treuresimbols(atrim(rst!nome)) + ";" + treuresimbols(atrim(rst!domicilie)) + ";" + treuresimbols(atrim(rst!codipostale)) + ";" + treuresimbols(atrim(rst!poblacioe)) + ";" + treuresimbols(atrim(rst!provinciae))
       linia = linia + ";" + IIf(cabool(!albaravalorat), "Y", "N") + ";" + IIf(atrim(rst!pais) = "", "ES", atrim(rst!pais)) + ";" + IIf(atrim(rst!Idioma) = "", "ES", atrim(rst!Idioma))
    End If
    Print #1, linia
    End With
End Sub
Sub generar_capcalera_clixes_fitxer_sap(rstm As Recordset, vempresa As String)
    Dim rst As Recordset
    Dim rstclientsap As Recordset
    Set rstclientsap = dbcomandes.OpenRecordset("select * from clients_codissap where codisap=" + atrim(cadbl(rstm!codiclientfactclixes)))
    With rstm
    vdesccosmailclixes = ""
    If Not rstclientsap.EOF Then vdesccosmailclixes = "Facturats per " + atrim(vempresa) + " al client " + atrim(rstm!codiclientfactclixes) + "-" + atrim(rstclientsap!nomclient) + "  Treball: " + atrim(!id_treball) + "/" + atrim(!ordre) + " Idioma: " + atrim(idiomaclientclixes) + Chr(13)
    linia = atrim(vempresa) + ";" + atrim(!id_treball) + atrim(Format(!ordre, "000")) + ";" + Format(Now, "dd/mm/yy") + ";" + atrim(!codiclientfactclixes) + ";1;;;;;-"
    linia = linia + ";;;;"
    linia = linia + ";Y;;" + idiomaclientclixes
    Print #1, linia
    End With
    Set rst = Nothing
    Set rstclientsap = Nothing
End Sub
Function cabool(valor As Variant) As Boolean
'On Error Resume Next
   If atrim(valor) = "S" Then valor = True
   If atrim(valor) = "N" Then valor = False
   If IsNull(valor) Or atrim(valor) = "" Then
      cabool = False
        Else: cabool = valor
   End If
End Function
Sub generar_linia_fitxer_sap(rstlinalb As Recordset, rstalb As Recordset, nomfitxer As String, idlinia As Long)
   Dim r As String
   Dim linia As String
   Dim ruta As String
   Dim rstidcompra As Recordset
   Dim espesor As Double
   Dim mesuraespesor As String
   Dim numlotfabricacio As String
   Dim comandesrelacionades As String
   Dim numalbaraprov As String
   Dim vunitatmesura As String
   On Error GoTo errorgravar
    r = nomfitxer
    If Not existeix(r) Then
              Open r For Output As 1
              generar_capcalera_fitxer_sap rstalb
          Else: Open r For Append As 1
    End If
    With rstlinalb
    vunitatmesura = atrim(!unitatmesura)
    If InStr(1, vunitatmesura, "/FIX") > 0 Then vunitatmesura = " "
    linia = "#" + atrim(idlinia) + ".1;" + treuresimbols(atrim(!codiproducte)) + ";" + treuresimbols(atrim(vunitatmesura)) + ";" + treuresimbols(atrim(!descripcioproducte)) + ";" + treurecomaperpunt(atrim(!quantitat)) + ";" + treurecomaperpunt(atrim(!preuvenda))
    Print #1, linia
    linia = "#" + atrim(idlinia) + ".2;" + treurecomaperpunt(atrim(!ampladamaterial)) + ";" + treurecomaperpunt(atrim(!espesor)) + ";" + treuresimbols(atrim(!mesuraespesor)) + ";" + treuresimbols(atrim(!descripciomides))
    Print #1, linia
    linia = "#" + atrim(idlinia) + ".3;" + treuresimbols(atrim(!marcailinia))
    Print #1, linia
    linia = "#" + atrim(idlinia) + ".4;" + treuresimbols(atrim(!numcomandacli)) + ";" + treuresimbols(atrim(!refclient))
    Print #1, linia
    linia = "#" + atrim(idlinia) + ".5;" + treuresimbols(atrim(!numcomandaclideclient)) + ";" + treuresimbols(atrim(!refclientdeclient)) + ";" + treuresimbols(atrim(!datafabricacio)) + ";" + treuresimbols(Mid(atrim(!codibarres) + " ", 1, 13))
    Print #1, linia
    linia = "#" + atrim(idlinia) + ".6;" + treuresimbols(atrim(!numcontracte)) + ";" + treuresimbols(Mid(atrim(!numcalloff), 1, 15))
    Print #1, linia
    linia = "#" + atrim(idlinia) + ".7;" + treuresimbols(atrim(!numbobs)) + ";" + treurecomaperpunt(atrim(!kgtotalsbruts)) + ";" + treurecomaperpunt(atrim(!metreslineals)) + ";" + treurecomaperpunt(atrim(!unitats)) + ";" + treuresimbols(atrim(!lotinplacsa)) + ";" + treuresimbols(atrim(!tipusproducte))
    Print #1, linia
    linia = "#" + atrim(idlinia) + ".8;" + treuresimbols(atrim(!observacionslinia))
    Print #1, linia
    possar_liniesvisuals idlinia
    End With
    Close 1
    Exit Sub
errorgravar:
      MsgBox err.Description & Chr(10) & "No s'ha gravat el fitxer: " + Chr(10) + r
End Sub
Sub possar_liniesvisuals(idlinia As Long)
   Dim i As Byte
   If llistasobrepaper.ListCount = 0 Then Exit Sub
   For i = 0 To llistasobrepaper.ListCount - 1
      If InStr(1, llistasobrepaper.List(i), "Ley 7/2022") = 0 And InStr(1, llistasobrepaper.List(i), "LLei 7/2022") = 0 Then  'No passo a l'albarà la linia de detall de l'IMPOST SOBRE ENVASOS
           Print #1, "LV" + atrim(idlinia) + "." + atrim(i + 1) + ";" + substituir(llistasobrepaper.List(i), ";", "_")
         Else
            If vimpostinclosalPVP = True Then
              Print #1, "LV" + atrim(idlinia) + "." + atrim(i + 1) + ";" + substituir(llistasobrepaper.List(i), ";", "_")
            End If
      End If
   Next i
   
End Sub
Sub possar_peusalbara(numalb As Double, nomfitxer As String)
   Dim rst As Recordset
   Dim idpa As Byte
   Dim linia As String
   If Not existeix(nomfitxer) Then
         Exit Sub
          Else: Open nomfitxer For Append As 1
    End If
   Set rst = datacapcalera.Database.OpenRecordset("select descripcio from liniespeu where numalbara=" + atrim(numalb) + " order by ordre asc")
   idpa = 1
   While Not rst.EOF
     linia = "PA" + atrim(idpa) + ";" + Mid(treuresimbols(rst!descripcio), 1, 79)
     Print #1, linia
     rst.MoveNext
     idpa = idpa + 1
   Wend
   Close 1
End Sub

Private Sub Command4_Click()
  Framecapcalera.Enabled = True
  If datacapcalera.Recordset.EditMode = 0 Then datacapcalera.Recordset.Edit
End Sub

Private Sub Command5_Click()
Frame3.visible = Not Frame3.visible
frameliniesalpaper.visible = Not frameliniesalpaper.visible
frameliniesalpaper.Left = framedadeslinia.Left
frameliniesalpaper.Top = framedadeslinia.Top
'frameliniesalpaper.ZOrder 1
End Sub

Private Sub Command6_Click()
   Framecapcalera.Enabled = False
   If datacapcalera.Recordset.EditMode > 0 Then
      datacapcalera.Recordset.Update
      actualitzarpeudepagina
      comprovarsicalpackinglist
   End If
End Sub
Sub possarliniesdepeualbarapredeterminadesdelclient()
   Dim rst As Recordset
   Dim vordre As Byte
   dataliniespeu.RecordSource = "select * from liniespeu where numalbara=" + atrim(datacapcalera.Recordset!numalbara) + " order by ordre"
   dataliniespeu.Refresh
   vordre = 5
   If dataliniespeu.Recordset.EOF Then
      Set rst = dbcomandes.OpenRecordset("select * from clients_notespeu where id_direnvio=" + atrim(datacapcalera.Recordset!id_direnvio) + " order by ordre")
      While Not rst.EOF And vordre < 21
         dataliniespeu.Recordset.AddNew
         dataliniespeu.Recordset!numalbara = datacapcalera.Recordset!numalbara
         dataliniespeu.Recordset!ordre = vordre
         dataliniespeu.Recordset!descripcio = atrim(rst!descripcio)
         dataliniespeu.Recordset.Update
         vordre = vordre + 5
         rst.MoveNext
      Wend
   End If
   Set rst = dbcomandes.OpenRecordset("select albaracaducitatmaterial from clients_envios where id=" + atrim(datacapcalera.Recordset!id_direnvio))
   If Not rst.EOF Then
      If cabool(rst!albaracaducitatmaterial) = True And vordre < 21 Then
         dataliniespeu.Recordset.AddNew
         dataliniespeu.Recordset!numalbara = datacapcalera.Recordset!numalbara
         dataliniespeu.Recordset!ordre = vordre
         dataliniespeu.Recordset!descripcio = possartextecaducitat
         dataliniespeu.Recordset.Update
         vordre = vordre + 5
      End If
   End If
   Set rst = dbcomandes.OpenRecordset("Select pais from clients_envios where id=" + atrim(datacapcalera.Recordset!id_direnvio))
   If atrim(rst!pais) = "GB" Then
         dataliniespeu.Recordset.AddNew
         dataliniespeu.Recordset!numalbara = datacapcalera.Recordset!numalbara
         dataliniespeu.Recordset!ordre = vordre
         dataliniespeu.Recordset!descripcio = "Cod.Mercancia: 39202029"
         dataliniespeu.Recordset.Update
         vordre = vordre + 5
   End If
   If atrim(rst!pais) = "US" Then
         dataliniespeu.Recordset.AddNew
         dataliniespeu.Recordset!numalbara = datacapcalera.Recordset!numalbara
         dataliniespeu.Recordset!ordre = vordre
         dataliniespeu.Recordset!descripcio = "Cod.Mercancia: 392021025"
         dataliniespeu.Recordset.Update
         vordre = vordre + 5
   End If
   Set rst = Nothing
End Sub
Function possartextecaducitat() As String
   Dim txt As String
   txt = "Recomendable utilizar antes de 9 meses."
   If idiomaclient = "" Then idiomaclient = "ES"
   fitxeridioma = llegir_ini("General", "rutallistats", "comandes.ini") + idiomaclient + "_etiquetareb.txt"
   f = llegir_ini("Idioma", txt, fitxeridioma)
   possartextecaducitat = f
End Function
Function buscar_INCOTERMS(vnumalbara As Double) As String
  Dim rst As Recordset
  Dim vnumc As Double
  Set rst = dbvendes.OpenRecordset("select lotinplacsa from liniesalbara where numalbara=" + atrim(vnumalbara))
  If Not rst.EOF Then
     vnumc = rst!lotinplacsa
     Set rst = dbvendes.OpenRecordset("select incoterm_envio from comandes_extres where comanda=" + atrim(vnumc))
     If Not rst.EOF Then
         buscar_INCOTERMS = atrim(rst!incoterm_envio)
     End If
  End If
  Set rst = Nothing

End Function
Sub actualitzarpeudepagina()
   Dim rsttrans As Recordset
   Dim vtransportista As String
   dataliniespeu.Database.Execute "delete * from liniespeu where numalbara=" + atrim(datacapcalera.Recordset!numalbara) + " and ordre=-1"
   If atrim(datacapcalera.Recordset!observacionsports) <> "" Then
        'dataliniespeu.Recordset.FindFirst "ordre=-1"
        'If dataliniespeu.Recordset.NoMatch Then
        dataliniespeu.Recordset.AddNew
        dataliniespeu.Recordset!numalbara = datacapcalera.Recordset!numalbara
        dataliniespeu.Recordset!ordre = -1
         '     Else
        '        dataliniespeu.Recordset.Edit
        'End If
        dataliniespeu.Recordset!descripcio = atrim(datacapcalera.Recordset!observacionsports) + IIf(cadbl(datacapcalera.Recordset!metrescubicstransport) > 0, " [" + atrim(datacapcalera.Recordset!metrescubicstransport) + " m3" + "]", "")
        dataliniespeu.Recordset.Update
   End If
   dataliniespeu.Database.Execute "delete * from liniespeu where numalbara=" + atrim(datacapcalera.Recordset!numalbara) + " and ordre=-2"
   If cadbl(datacapcalera.Recordset!id_transport) > 0 Or atrim(datacapcalera.Recordset!tipusports) <> "" Then
        If cadbl(datacapcalera.Recordset!id_transport) > 0 Then
             Set rsttrans = datacapcalera.Database.OpenRecordset("Select * from transportistes where  codi=" + atrim(datacapcalera.Recordset!id_transport))
             If Not rsttrans.EOF Then vtransportista = atrim(rsttrans!descripcio)
        End If
        'dataliniespeu.Recordset.FindFirst "ordre=-2"
        'If dataliniespeu.Recordset.NoMatch Then
        dataliniespeu.Recordset.AddNew
        dataliniespeu.Recordset!numalbara = datacapcalera.Recordset!numalbara
        dataliniespeu.Recordset!ordre = -2
        '      Else
        '        dataliniespeu.Recordset.Edit
        'End If
        txtports = " - " + traducciodeabreviatures("Ports", idiomaclient) + " " + traducciodeabreviatures(atrim(datacapcalera.Recordset!tipusports), idiomaclient) + " [" + buscar_INCOTERMS(datacapcalera.Recordset!numalbara) + "]"
        dataliniespeu.Recordset!descripcio = "Transport: " + atrim(vtransportista) + IIf(atrim(datacapcalera.Recordset!tipusports) <> "", txtports, "")
        dataliniespeu.Recordset.Update
   End If
   dataliniespeu.Refresh
End Sub

Sub imprimir_albara()
 
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistatalbaransexpedicions.rpt", 1)
 ' oreport.SQLQueryString = ""
  oreport.RecordSelectionFormula = "{capcaleraalbara.numalbara}=" + atrim(datacapcalera.Recordset!numalbara)
  oreport.FormulaFields.GetItemByName("nomdirenvio").Text = "'" + treure_apostruf(etinfodelclient.tag) + "'"
  oreport.SQLQueryString = ""
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "vendes.mdb"
  oreport.Database.Tables.Item(2).Location = rutadelfitxer(cami) + "vendes.mdb"
  oreport.DiscardSavedData
  oreport.VerifyOnEveryPrint = False
  
  
  
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
    Else
      oreport.PrintOut False, 1
  End If
  

End Sub

Private Sub Command7_Click()
  Dim rst As Recordset
  Dim vistaprevia As Boolean
  If etean128.visible Then MsgBox "Atenció aquest client vol frontals amb EAN 128.", vbCritical, "Error"
  Set rst = dbbaixes.OpenRecordset("select distinct numpalet from bobinesent where numalbara=" + atrim(datacapcalera.Recordset!numalbara) + " and comanda=" + atrim(datalinies.Recordset!lotinplacsa))
  If Not rst.EOF Then
     rst.MoveLast: rst.MoveFirst
     If MsgBox("Vols vista previa?", vbInformation + vbYesNo, "Vista previa") = vbYes Then vistaprevia = True
     datacapcalera.Database.Execute "update capcaleraalbara set papersfrontalsimpresos=True where numalbara=" + atrim(datacapcalera.Recordset!numalbara)
  End If
  While Not rst.EOF
     imprimir_frontals cadbl(rst!numpalet), cadbl(rst.AbsolutePosition) + 1, cadbl(rst.RecordCount), vistaprevia
     rst.MoveNext
  Wend
possarcampscapçalera
  
End Sub
Function buscaretiqueta(vetiqueta As String, idiomaclient As String)
  Dim i As Byte
  Dim colidioma As Byte
  Dim vector(100, 5)
  Dim vectorsoldadores(100, 5)
  inizialitzaretiquetes vector, vectorsoldadores
  colidioma = 2
  If idiomaclient = "CA" Then colidioma = 5
  If idiomaclient = "ES" Then colidioma = 4
  If idiomaclient = "FR" Then colidioma = 3
  i = 1
  While vector(i, 1) <> "--"
     If vector(i, 1) = UCase(vetiqueta) Then buscaretiqueta = vector(i, colidioma)
     i = i + 1
  Wend
End Function
Sub possar_etiquetesidioma(oreport As CRAXDDRT.Report, vsecant As String, idiomaclient As String)
  Dim i As Byte
  Dim colidioma As Byte
  Dim vector(100, 5)
  Dim vectorsoldadores(100, 5)
  inizialitzaretiquetes vector, vectorsoldadores
  colidioma = 2
  If idiomaclient = "CA" Then colidioma = 5
  If idiomaclient = "ES" Then colidioma = 4
  If idiomaclient = "FR" Then colidioma = 3
  i = 1
  While vector(i, 1) <> "-"
     oreport.FormulaFields.GetItemByName("e" + LCase(vector(i, 1))).Text = "'" + treure_apostruf(vector(i, colidioma)) + "'"
     i = i + 1
  Wend
  'si es soladores canviho les etiquetes corresponents
  If vsecant = "S" Then
    i = 1
    While vectorsoldadores(i, 1) <> "-"
       oreport.FormulaFields.GetItemByName("e" + LCase(vectorsoldadores(i, 1))).Text = "'" + treure_apostruf(vectorsoldadores(i, colidioma)) + "'"
       i = i + 1
    Wend
  End If
  
End Sub
Sub inizialitzaretiquetes(vector, vectorsoldadores)
  Dim i As Byte
  
  'cas que sigui material rebobinat
  i = 1
  vector(i, 1) = "SUPP": vector(i, 2) = "Supp": vector(i, 3) = "Fournisseur": vector(i, 4) = "Proveedor": vector(i, 5) = "Proveïdor": i = i + 1
  vector(i, 1) = "BATCH": vector(i, 2) = "Batch": vector(i, 3) = "Lot": vector(i, 4) = "Lote": vector(i, 5) = "Lot": i = i + 1
  vector(i, 1) = "PALET": vector(i, 2) = "Palet": vector(i, 3) = "Palet": vector(i, 4) = "Palet": vector(i, 5) = "Palet": i = i + 1
  vector(i, 1) = "CUSTOMER": vector(i, 2) = "Customer": vector(i, 3) = "Client": vector(i, 4) = "Cliente": vector(i, 5) = "Client": i = i + 1
  vector(i, 1) = "DELIVERY": vector(i, 2) = "Delivery Addres": vector(i, 3) = "Adresse de livraison": vector(i, 4) = "Dirección de entrega": vector(i, 5) = "Direcció d´entrega": i = i + 1
  vector(i, 1) = "CONTRACT": vector(i, 2) = "Contract": vector(i, 3) = "Contrat": vector(i, 4) = "Contrato": vector(i, 5) = "Contracte": i = i + 1
  vector(i, 1) = "ORDER": vector(i, 2) = "Order No.": vector(i, 3) = "Commande Nº": vector(i, 4) = "Pedido Nº": vector(i, 5) = "Comanda Nº": i = i + 1
  vector(i, 1) = "DELIVERYORDER": vector(i, 2) = "Delivery Order No.": vector(i, 3) = "Bon de livraison Nº": vector(i, 4) = "Orden de entrega Nº": vector(i, 5) = "Ordre d´entrega Nº": i = i + 1
  vector(i, 1) = "REF": vector(i, 2) = "Ref.": vector(i, 3) = "Réf.": vector(i, 4) = "Ref": vector(i, 5) = "Ref": i = i + 1
  vector(i, 1) = "MATERIAL": vector(i, 2) = "Material": vector(i, 3) = "Matériau": vector(i, 4) = "Material": vector(i, 5) = "Material": i = i + 1
  vector(i, 1) = "WIDTH": vector(i, 2) = "Width": vector(i, 3) = "Largeur": vector(i, 4) = "Ancho": vector(i, 5) = "Ample": i = i + 1
  vector(i, 1) = "REPEAT": vector(i, 2) = "Repeat": vector(i, 3) = "Développement": vector(i, 4) = "Desarroll": vector(i, 5) = "": i = i + 1
  vector(i, 1) = "PRINTING": vector(i, 2) = "Printing": vector(i, 3) = "Texte d'impression": vector(i, 4) = "Texto de impresión": vector(i, 5) = "Texte d´impresió": i = i + 1
  vector(i, 1) = "BARCODE": vector(i, 2) = "Barcode": vector(i, 3) = "Code-barres": vector(i, 4) = "Código de barras": vector(i, 5) = "Còdi de barres": i = i + 1
  vector(i, 1) = "DATE": vector(i, 2) = "Date": vector(i, 3) = "Date": vector(i, 4) = "Fecha": vector(i, 5) = "Data": i = i + 1
  vector(i, 1) = "BEST": vector(i, 2) = "Best": vector(i, 3) = "À consommer avant": vector(i, 4) = "Fecha consumo preferente": vector(i, 5) = "Data consum preferent": i = i + 1
  vector(i, 1) = "PCS": vector(i, 2) = "Pcs": vector(i, 3) = "Pcs": vector(i, 4) = "Pzs": vector(i, 5) = "Pcs": i = i + 1
  vector(i, 1) = "REELS": vector(i, 2) = "Reels": vector(i, 3) = "Bobines": vector(i, 4) = "Bobinas": vector(i, 5) = "Bobines": i = i + 1
  vector(i, 1) = "METERS": vector(i, 2) = "Meters": vector(i, 3) = "Mètres": vector(i, 4) = "Metros": vector(i, 5) = "Metres": i = i + 1
  vector(i, 1) = "NETW": vector(i, 2) = "Net Weight": vector(i, 3) = "Poids net": vector(i, 4) = "Peso neto": vector(i, 5) = "Pes net": i = i + 1
  vector(i, 1) = "GROSSW": vector(i, 2) = "Gross Weight": vector(i, 3) = "Poids brut": vector(i, 4) = "Peso bruto": vector(i, 5) = "Pes brut": i = i + 1
  vector(i, 1) = "BOBSXMTS": vector(i, 2) = "Reels X Mts.": vector(i, 3) = "bobines X Mts.": vector(i, 4) = "Bobinas X Mts.": vector(i, 5) = "Bobines X Mts.": i = i + 1
  vector(i, 1) = "-": i = i + 1
  'AQUI POSO ALTRES ETIQUETES QUE PUC NECESSITAR
  vector(i, 1) = "PESO": vector(i, 2) = "Weight": vector(i, 3) = "Poids": vector(i, 4) = "Peso": vector(i, 5) = "Pes": i = i + 1
  vector(i, 1) = "--": i = i + 1
  
  'cas que sigui material soldadores
  i = 1
  vectorsoldadores(i, 1) = "REELS": vectorsoldadores(i, 2) = "Boxes": vectorsoldadores(i, 3) = "Caisses": vectorsoldadores(i, 4) = "Cajas": vectorsoldadores(i, 5) = "Caixes": i = i + 1
  vectorsoldadores(i, 1) = "METERS": vectorsoldadores(i, 2) = "Bags": vectorsoldadores(i, 3) = "Sacs": vectorsoldadores(i, 4) = "Bolsas": vectorsoldadores(i, 5) = "Bosses": i = i + 1
  vectorsoldadores(i, 1) = "BOBSXMTS": vectorsoldadores(i, 2) = "Bags X Box": vectorsoldadores(i, 3) = "Sacs X Caisse": vectorsoldadores(i, 4) = "Bolsas X Caja.": vectorsoldadores(i, 5) = "Bosses X caixa.": i = i + 1
  vectorsoldadores(i, 1) = "-"
  
End Sub
Function triartipusmaterial(vdesc As String, vmetres As Double) As String
  Dim vinici As Byte
  Dim vfi As Byte
  vinici = InStr(1, vdesc, "µ")
  If vinici > 0 Then
     vfi = InStr(vinici + 3, vdesc, " ")
     If vinici > vfi Then
        vinici = 1
        vfi = InStr(vinici + 3, vdesc, " ")
          Else: vinici = vinici - 3
     End If
     
     If vmetres = 0 Then vinici = 1 'si els metres son 0 es que son bosses i agafo tota la linia
     If vfi > vinici Then
        triartipusmaterial = Mid(vdesc, vinici, vfi - vinici)
     End If
     If vmetres <> 0 And InStr(vdesc, "Cod:") = 0 Then
        vinici = InStr(1, vdesc, "µ") - 3
        vfi = InStr(1, vdesc, "Imp.")
        If vfi = 0 Then
           vfi = Len(vdesc)
            Else: vfi = vfi - vinici
        End If
        triartipusmaterial = Mid(vdesc, vinici, vfi)
     End If
  End If
End Function
Function buscardesarrollcomanda(vnumc As Double) As Double
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
  If rst.EOF Then Exit Function
  Set rst = dbclixes.OpenRecordset("select desarroll from modificacions where id_treball=" + atrim(cadbl(rst!numtreball)) + " and ordre=" + atrim(cadbl(rst!numordremodificacio)))
  If Not rst.EOF Then buscardesarrollcomanda = cadbl(rst!desarroll)
End Function
Sub netejarformules(oreport As CRAXDDRT.Report)
  Dim i As Byte
  For i = 1 To oreport.FormulaFields.Count
    If oreport.FormulaFields.Item(i).Name <> "{@codidebarrescomandamespalet}" Then oreport.FormulaFields.Item(i).Text = ""
  Next i
End Sub
Sub possar_dadesinformefrontals(oreport As CRAXDDRT.Report, vnumpalet As Byte, vnumpaletdeX As Byte, vtotalpalets As Byte, vnumcopies As Byte)
  Dim rsttotals As Recordset
  Dim rstc As Recordset
  Dim rstcli As Recordset
  Dim vnumc As Double
  Dim vpespalet As Double
  Dim vdesarroll As Double
  Dim vpestotal As Double
  Dim rstcextres As Recordset
  vnumc = cadbl(datalinies.Recordset!lotinplacsa)
  Set rsttotals = dbbaixes.OpenRecordset("Select count(*) as tbobines,sum(metresisacs) as tmetres ,sum(kilosiunitats) as tkilos,sum(kilosnets) as tkilosnets,first(seccio) as seccioanterior from bobinesent where numalbara=" + atrim(datacapcalera.Recordset!numalbara) + " and comanda=" + atrim(datalinies.Recordset!lotinplacsa) + " and numpalet=" + atrim(vnumpalet) + " group by numpalet")
  Set rstc = dbcomandes.OpenRecordset("SELECT comandes.comanda, clients.nom, clients.numproveidor, Clients_envios.nome, Clients_envios.domicilie, Clients_envios.codipostale,clients_envios.poblacioe, Clients_envios.provinciae, Clients_envios.pfpaperfrontal, Clients_envios.pfdatafab,clients_envios.albarasensetexteimpresio FROM (comandes LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN Clients_envios ON comandes.direnvio = Clients_envios.id WHERE (((comandes.comanda)=" + atrim(vnumc) + "));")
  Set rstcextres = dbcomandes.OpenRecordset("select * from comandes_extres where comanda=" + atrim(vnumc), , ReadOnly)
  Set rstcli = dbcomandes.OpenRecordset("select * from clients_envios where id=" + atrim(cadbl(datacapcalera.Recordset!id_direnvio)))
  If rstc.EOF Or rsttotals.EOF Or rstcli.EOF Then Exit Sub
  vnumcopies = cadbl(rstcli!copiespaperfrontal)
  If vnumcopies < 1 Then vnumcopies = 1
  If vnumcopies > 4 Then vnumcopies = 4
  possar_etiquetesidioma oreport, atrim(rsttotals!seccioanterior), idiomaclient
  With datalinies.Recordset
  oreport.FormulaFields.GetItemByName("dnumpaletreal").Text = "'" + atrim(vnumpalet) + "'"
  oreport.FormulaFields.GetItemByName("dsupp").Text = "'" + atrim(rstc!numproveidor) + "'"
  oreport.FormulaFields.GetItemByName("dbatch").Text = atrim(rstc!comanda)
  oreport.FormulaFields.GetItemByName("dpalet").Text = "'" + atrim(vnumpaletdeX) + "/" + atrim(vtotalpalets) + "'"
  oreport.FormulaFields.GetItemByName("dcustomer").Text = "'" + treure_apostruf(atrim(rstcli!nome)) + "'"
  oreport.FormulaFields.GetItemByName("ddelivery1").Text = "'" + treure_apostruf(atrim(rstcli!nome)) + "'"
  oreport.FormulaFields.GetItemByName("ddelivery2").Text = "'" + atrim(rstcli!codipostale) + IIf(atrim(rstcli!codipostale) <> "", " - ", "") + atrim(rstcli!poblacioe) + "'"
  oreport.FormulaFields.GetItemByName("ddelivery3").Text = "'" + atrim(rstcli!provinciae) + "'"
  oreport.FormulaFields.GetItemByName("dcontract").Text = "'" + atrim(!numcontracte) + "'"
  oreport.FormulaFields.GetItemByName("dorder").Text = "'" + atrim(!numcomandacli) + "'"
  oreport.FormulaFields.GetItemByName("dmaterial1").Text = "'" + atrim(!descripcioproducte) + IIf(atrim(!colormaterial) <> "", " (" + atrim(!colormaterial) + ")", "") + "'"
  oreport.FormulaFields.GetItemByName("dmaterial2").Text = "'" + triartipusmaterial(!descripciomides, !metreslineals) + "'"
  oreport.FormulaFields.GetItemByName("ddeliveryorder").Text = "'" + atrim(!numcalloff) + "'"
  oreport.FormulaFields.GetItemByName("dref").Text = "'" + atrim(!refclient) + "'"
  oreport.FormulaFields.GetItemByName("refinplacsa").Text = "'" + IIf(cabool(rstcli!paletreferenciainplacsa), "Inplacsa Ref: " + atrim(rstcextres!refinplacsa), "") + "'"
  oreport.FormulaFields.GetItemByName("dwidth").Text = atrim(!ampladamaterial)
  oreport.FormulaFields.GetItemByName("dprinting").Text = "'" + IIf(cabool(rstc!albarasensetexteimpresio), "", atrim(!marcailinia)) + "'"
  oreport.FormulaFields.GetItemByName("dbarcode").Text = "'" + atrim(!codibarres) + "'"
  If InStr(1, atrim(rstcli!estilfrontal), "EAN-13") > 0 Then
       oreport.FormulaFields.GetItemByName("valorcodidebarres").Text = "'*" + atrim(!codibarres) + "*'"
         Else: oreport.FormulaFields.GetItemByName("valorcodidebarres").Text = "''"
  End If
  oreport.FormulaFields.GetItemByName("drepeat").Text = "0"
  vdesarroll = buscardesarrollcomanda(vnumc)
  If InStr(1, !codiproducte, "I") Then oreport.FormulaFields.GetItemByName("drepeat").Text = passaradecimalpunt(atrim(vdesarroll))
  If rstc!pfdatafab Then
     oreport.FormulaFields.GetItemByName("ddate").Text = "'" + atrim(!datafabricacio) + "'"
     oreport.FormulaFields.GetItemByName("dbest").Text = "'" + atrim(DateAdd("m", 9, !datafabricacio)) + "'"
       Else:
          oreport.FormulaFields.GetItemByName("ddate").Text = "'" + atrim(datacapcalera.Recordset!dataalbara) + "'"
          oreport.FormulaFields.GetItemByName("dbest").Text = "''"
          oreport.FormulaFields.GetItemByName("ebest").Text = "''"
  End If
  oreport.FormulaFields.GetItemByName("dreels").Text = atrim(rsttotals!tbobines)
  oreport.FormulaFields.GetItemByName("dmeters").Text = atrim(rsttotals!tmetres)
  If vdesarroll > 0 Then oreport.FormulaFields.GetItemByName("dpcs").Text = Redondejar(cadbl(atrim(rsttotals!tmetres) / (vdesarroll / 1000)), 0)
  oreport.FormulaFields.GetItemByName("dnetw").Text = passaradecimalpunt(atrim(rsttotals!tkilosnets))
  's'ha de buscar el pes del palet si es de rebobinadora a bobinesent hi ha la secció
  'si no hi ha pes o es Soldadora sumarem 22 kg al pes brut
  vpespalet = buscarpespalet(vnumc, vnumpalet)
  If cadbl(rsttotals!tkilosnets) < 1 Then
     vpespalet = 0
     oreport.FormulaFields.GetItemByName("enetw").Text = "''"
     oreport.FormulaFields.GetItemByName("dnetw").Text = "''"
     oreport.FormulaFields.GetItemByName("egrossw").Text = "'" + buscaretiqueta("PESO", idiomaclient) + "'"
  End If
  vpestotal = cadbl(rsttotals!tkilos)
  oreport.FormulaFields.GetItemByName("dgrossw").Text = atrim(Redondejar(vpestotal + vpespalet, 0))
  oreport.FormulaFields.GetItemByName("dbobsxmts").Text = "'" + generarliniadepackinglist(vnumpalet, datacapcalera.Recordset!numalbara, datalinies.Recordset!lotinplacsa) + "'"
  If !metreslineals = 0 Then sisonbossesnetejarcamps oreport
  End With
  Set rstc = Nothing
  Set rsttotals = Nothing
  Set rstcli = Nothing
End Sub
Sub sisonbossesnetejarcamps(oreport As CRAXDDRT.Report)
    oreport.FormulaFields.GetItemByName("dpcs").Text = ""
    oreport.FormulaFields.GetItemByName("epcs").Text = ""
    oreport.FormulaFields.GetItemByName("dwidth").Text = ""
    oreport.FormulaFields.GetItemByName("ewidth").Text = ""
    oreport.FormulaFields.GetItemByName("drepeat").Text = ""
    oreport.FormulaFields.GetItemByName("erepeat").Text = ""
End Sub
Function generarliniadepackinglist(vnumpalet As Byte, vnumalb As Double, vnumc As Double, Optional vambunitats As Boolean) As String
  Dim rst As Recordset
  Dim rstc As Recordset
  Dim vlinia As String
  Dim vliniaunitats As String
  Dim vdesarroll As Double
  Dim vcontinuu As Boolean
  Dim vruta As String
  Set rstc = dbcomandes.OpenRecordset("SELECT productes.ruta, comandes.comanda FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(vnumc) + "));")
  If Not rstc.EOF Then vruta = atrim(rstc!ruta)
  If vambunitats Then
     Set rstc = dbcomandes.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vnumc))
     If Not rstc.EOF Then
         vdesarroll = mirardesarrolldeltreball(cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), vcontinuu)
         If vcontinuu Then vdesarroll = 0
     End If
  End If
  Set rst = dbbaixes.OpenRecordset("select count(*) as tbobs,first(metresisacs) as tmetresisacs from bobinesent where comanda=" + atrim(vnumc) + IIf(vnumpalet > 0, " and numpalet=" + atrim(vnumpalet), "") + " and numalbara=" + atrim(vnumalb) + " group by metresisacs")
  While Not rst.EOF
    vliniaunitats = ""
    If vdesarroll > 0 Then vliniaunitats = atrim(Redondejar(cadbl(rst!tmetresisacs) / vdesarroll, 0)) + "p"
    vlinia = vlinia + IIf(vlinia <> "", " + ", "")
    If InStr(1, vruta, "S") = 0 Then
         vlinia = vlinia + atrim(cadbl(rst!tbobs)) + "X" + IIf(vliniaunitats <> "", "(" + vliniaunitats + "/", "") + atrim(cadbl(rst!tmetresisacs)) + "m" + IIf(vliniaunitats > "", ")", "")
          Else
            vlinia = vlinia + IIf(vliniaunitats <> "", "(" + vliniaunitats + "/", "") + atrim(cadbl(rst!tmetresisacs)) + "u" + IIf(vliniaunitats > "", ")", "") + "X" + atrim(cadbl(rst!tbobs))
    End If
    rst.MoveNext
  Wend
  generarliniadepackinglist = vlinia
  Set rst = Nothing
  Set rstc = Nothing
End Function
Function buscarpespalet(vnumc As Double, vnumpalet As Byte) As Double
   Dim rst As Recordset
   Set rst = dbbaixes.OpenRecordset("select pespalet from reb_pespalets where numpalet=" + atrim(vnumpalet) + " and comanda=" + atrim(vnumc))
   If Not rst.EOF Then
      buscarpespalet = cadbl(rst!pespalet)
        Else: buscarpespalet = 22
   End If
   Set rst = Nothing
End Function
Sub imprimir_frontals(vnumpalet As Byte, vnumpaletdeX As Byte, vtotalpalets As Byte, vistaprevia As Boolean)
 Dim oapp As CRAXDDRT.Application
 Dim vnumcopies As Byte
 Dim i As Byte
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "paperfrontalexpedicions.rpt", 1)
  'oreport.FormulaFields.GetItemByName("nomdirenvio").Text = "'" + treure_apostruf(etinfodelclient.tag) + "'"
  netejarformules oreport
  possar_dadesinformefrontals oreport, vnumpalet, vnumpaletdeX, vtotalpalets, vnumcopies
  
'  vnumcopies = cadbl(InputBox("Entra el numero de copies", "Copies", atrim(vnumcopies)))
 ' If vnumcopies < 1 Or vnumcopies > 4 Then vnumcopies = 1
 ' If existeix("c:\ordprog.ini") Then
  If vistaprevia Then
    Load veurereport
    veurereport.CRViewer.ReportSource = oreport
    veurereport.CRViewer.DisplayGroupTree = False
    veurereport.CRViewer.ViewReport
    veurereport.Show 1, Me
    Else
      For i = 1 To vnumcopies
        oreport.PrintOut False, 1
      Next i
  End If
  
End Sub


Private Sub Command8_Click()
   imprimir_albara
End Sub

Private Sub Command9_Click()
   If datalinies.Recordset.EOF Then Exit Sub
   actualitzar_bobinesent_vendes cadbl(datalinies.Recordset!lotinplacsa)
   carregar_bobinesentrega cadbl(datalinies.Recordset!lotinplacsa), cadbl(datalinies.Recordset!numalbara)
   MsgBox "Bobines actualitzades..."
End Sub
Sub actualitzar_bobinesent_vendes(numc As Double)
  If numc < 1 Then Exit Sub
  Set dbtmp = dbcomandes
  Set dbtmpb = dbbaixes
  ruta = rutadelproducte(numc)
  actualitzar_bobinesent numc
End Sub
Private Sub datacapcalera_Reposition()
   Dim vm3 As Double
   Dim vnumbases As Double
   If Not datacapcalera.Recordset.EOF Then
        datalinies.RecordSource = "select * from liniesalbara where numalbara=" + atrim(datacapcalera.Recordset!numalbara)
        datacapcalera.caption = "Alb: " + atrim(datacapcalera.Recordset.AbsolutePosition + 1) + "/" + atrim(datacapcalera.Recordset.RecordCount)
          Else: datalinies.RecordSource = "select * from liniesalbara where numalbara=-99999"
   End If
   Framecapcalera.Enabled = False
   activar_frames False
   ettanper100merma.tag = "<"
   datalinies.Refresh
   If Mid(ettanper100merma.tag + " ", 1, 1) = "<" Then ettanper100merma.tag = Mid(ettanper100merma.tag, 2)
   If Not datalinies.Recordset.EOF Then
       If Text1.MaxLength = 0 Then possar_amplesmax_linies
       vm3 = calcular_metrescubics_albara(datacapcalera.Recordset!numalbara, vnumbases)
       etmetrescubicscalculats.ForeColor = IIf(vm3 > 0, QBColor(2), QBColor(12))  'verd 2  i vermell 12
       If vm3 < 0 Then vm3 = vm3 * -1
       etmetrescubicscalculats = atrim(vm3) + " m3 " + "   " + atrim(vnumbases) + " Bases."
       If cadbl(vm3) > 0 Then
          vm3 = Redondejar(vm3, 2)
          datacapcalera.Database.Execute "update capcaleraalbara set metrescubicstransport=" + passaradecimalpunt(atrim(vm3)) + " where numalbara=" + atrim(datacapcalera.Recordset!numalbara)
          If Not escrops Then datacapcalera.Database.Execute "update capcaleraalbara set numbases=" + atrim(vnumbases) + " where numalbara=" + atrim(datacapcalera.Recordset!numalbara)
       End If
   End If
   possarcampscapçalera
   sumarkilosimetresseleccionats
   mirarsiclixesperfacturar
   If combotransportista.Text = "" Then bassignartransport.visible = True Else bassignartransport.visible = False
   possar_etiqueta_tanx100mermes datacapcalera.Recordset!numalbara
End Sub
Function calcular_metrescubics_albara(vnumalb As Double, vnumbases As Double) As Double
   Dim rst As Recordset
   Dim rstemb As Recordset
   Dim rstalb As Recordset
   Dim vcms As Double
   Dim vfaltenemb As Boolean
   Dim vpaletsbase As String
   Dim vllistaalbarans As String
   
   If escrops Then
        Set rstalb = dbvendes.OpenRecordset("select * from capcaleraalbara where codiclient=" + atrim(datacapcalera.Recordset!codiclient) + " and id_transport=" + atrim(datacapcalera.Recordset!id_transport) + " and id_direnvio=" + atrim(datacapcalera.Recordset!id_direnvio) + " and dataalbara=#" + Format(datacapcalera.Recordset!dataalbara, "mm/dd/yy") + "# order by numalbara desc")
         Else: Set rstalb = dbvendes.OpenRecordset("select * from capcaleraalbara where numalbara=" + atrim(vnumalb))
   End If
   If Not rstalb.EOF Then vnumalb = rstalb!numalbara
   While Not rstalb.EOF
        vllistaalbarans = vllistaalbarans + IIf(vllistaalbarans = "", "", " ,") + atrim(rstalb!numalbara)
        'dbvendes.Execute "update embolicarpalets set numalbara=0 where numalbara=" + atrim(vnumalb)
        Set rst = dbvendes.OpenRecordset("select distinct trim(comanda)+'/'+trim(numpalet) as NPalet,comanda,numpalet from bobinesent where numalbara=" + atrim(rstalb!numalbara))
        While Not rst.EOF
          Set rstemb = dbvendes.OpenRecordset("select * from embolicarpalets where numcomanda=" + atrim(rst!comanda) + " and numpalet=" + atrim(rst!numpalet))
          If Not rstemb.EOF Then
               vcms = vcms + cadbl(rstemb!metres) + IIf(rstemb!posicioalabase = 1, 16, 17)
               dbvendes.Execute "update embolicarpalets set numalbara=" + atrim(vnumalb) + " where numcomanda=" + atrim(rst!comanda) + " and numpalet=" + atrim(rst!numpalet)
               If InStr(1, vpaletsbase, atrim(rstemb!numreferenciagrup)) = 0 Then vnumbases = vnumbases + 1: vpaletsbase = vpaletsbase + " " + atrim(rstemb!numreferenciagrup)
               If IsNull(rstemb!Database) Then vfaltenemb = True
                Else: vfaltenemb = True
          End If
          rst.MoveNext
        Wend
        rstalb.MoveNext
   Wend
   vcms = vcms / 100
   calcular_metrescubics_albara = (0.8 * 1.2 * vcms) * IIf(vfaltenemb, -1, 1)
    ' aixó seria per si volem guardar els numeros de bases a l'ultim albarà CROPS
      'també s's'ha de posar a zero les bases de tots els altres albarans relacionats
   If escrops Then
        dbvendes.Execute "update capcaleraalbara set numbases=0 where numalbara in (" + atrim(vllistaalbarans) + ")"
        rstalb.MoveLast: rstalb.MoveFirst: rstalb.Edit: rstalb!numbases = vnumbases: rstalb.Update
   End If
   Set rst = Nothing
   Set rstemb = Nothing
End Function

Sub mirarsiclixesperfacturar(Optional preguntarperfacturar As Boolean, Optional vtreball As Double, Optional vordre As Byte)
   Dim vllistacomandes As String
   Dim vfacturarono As Boolean
   Dim rst2 As Recordset
   If vtreball = 0 Then Exit Sub
   If datalinies.Recordset.EOF Then Exit Sub
   Set rst2 = datalinies.Recordset.Clone
   rst2.MoveFirst
   While Not rst2.EOF
     comprovarsihihaclixesperenviar IIf(vtreball > 0, 0, rst2!lotinplacsa), vllistacomandes, False, vtreball, vordre
     If InStr(1, vllistacomandes, IIf(vtreball > 0, " 0", atrim(rst2!lotinplacsa))) > 0 And preguntarperfacturar Then
        vfacturarono = False
        If vtreball = 0 Then
           vfacturarono = IIf(MsgBox("Vols facturar els clixes de la comanda " + atrim(rst2!lotinplacsa) + "?", vbInformation + vbYesNo + vbDefaultButton2, "Albaranar Clixes") = vbYes, True, False)
             Else: vfacturarono = IIf(MsgBox("Vols facturar els clixes del treball " + atrim(vtreball) + "/" + atrim(vordre) + "?", vbInformation + vbYesNo + vbDefaultButton2, "Albaranar Clixes") = vbYes, True, False)
        End If
        If vfacturarono Then comprovarsihihaclixesperenviar IIf(vtreball > 0, 0, rst2!lotinplacsa), vllistacomandes, True, vtreball, vordre
        If vtreball > 0 Then GoTo fi
     End If
     rst2.MoveNext
   Wend
fi:
   Set rst2 = Nothing
   If atrim(vllistacomandes) <> "" And vtreball = 0 Then etmissatge = "Comandes amb clixes per facturar: " + vllistacomandes
End Sub
Sub possarcampscapçalera()
   Framecapcalera.Enabled = False
   If datacapcalera.Recordset.EOF Then Exit Sub
   carregar_dadesclient datacapcalera.Recordset!id_direnvio, datacapcalera.Recordset!codiclient

   possar_transportista_alacapcalera datacapcalera.Recordset!numalbara, cadbl(datacapcalera.Recordset!id_transport)
   escullirtransportista cadbl(datacapcalera.Recordset!id_transport)
   etmetrescubics = ""
   If datacapcalera.Recordset!papersfrontalsimpresos Then Checkpapersfrontalsimpresos.Value = 1 Else Checkpapersfrontalsimpresos.Value = 0
   If cadbl(datacapcalera.Recordset!metrescubicstransport) > 0 Then etmetrescubics = "[" + atrim(datacapcalera.Recordset!metrescubicstransport) + " m3" + "]"
   
End Sub
Sub possar_transportista_alacapcalera(vnumalb As Double, vidtransport As Double)
   Dim rstc As Recordset
   'Set rstc = datacapcalera.Database.OpenRecordset("SELECT linies_expedicions.albara, comandes_extres.transportista_albara FROM linies_expedicions LEFT JOIN comandes_extres ON linies_expedicions.comanda = comandes_extres.comanda where albara=" + atrim(vnumalb))
   Set rstc = datacapcalera.Database.OpenRecordset("SELECT comandes_extres.transportista_albara FROM liniesalbara LEFT JOIN comandes_extres ON liniesalbara.lotinplacsa = comandes_extres.comanda Where numalbara = " + atrim(vnumalb))
   If Not rstc.EOF Then
        If cadbl(rstc!transportista_albara) > 0 Then
             If atrim(datacapcalera.Recordset!tipusports) = "" Then datacapcalera.Database.Execute "update capcaleraalbara set tipusports='Pagats' where numalbara=" + atrim(vnumalb)
             If vidtransport <> cadbl(rstc!transportista_albara) Then
                 datacapcalera.Database.Execute "update capcaleraalbara set id_transport=" + atrim(rstc!transportista_albara) + " where numalbara=" + atrim(vnumalb)
                 wait 1
                 actualitzarpeudepagina
             End If
        End If
   End If
   Set rstc = Nothing
End Sub
Sub activaredicio(estat As Boolean)
   bdesbloquejarsap.visible = Not estat
'   estat = True  's'ha d'eliminar quan controlem la no edició cap a sap
   
   alta.Enabled = estat
   eliminar.Enabled = estat
   modificar.Enabled = estat
   Command1.Enabled = estat
   Command9.Enabled = estat
   Command4.Enabled = estat
   Command6.Enabled = estat
End Sub
Sub carregar_dadesclient(idenvio As Long, codiclient As Double)
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim rstclientcomptable As Recordset
   vimpostinclosalPVP = False
   etinfodelclient = ""
   ettraspasasap = ""
   ettraspasasap.BackStyle = 0
   ettraspasasap.BackColor = formvendes.BackColor
   etean128.visible = False
   Set rst = dbcomandes.OpenRecordset("select * from clients_envios where id=" + atrim(idenvio))
   If rst.EOF Then Exit Sub
   Set rstc = dbcomandes.OpenRecordset("select * from clients where codi=" + atrim(rst!codi))
   If rst.EOF Then Exit Sub
   Set rstclientcomptable = dbcomandes.OpenRecordset("select * from clients_codiscomptables where codicomptable=" + atrim(codiclient) + " and codifabricacio=" + atrim(rstc!codi))
   etfiproduccio = atrim(rst!avisfiproduccio)
   idiomaclient = atrim(rst!Idioma)
   etmissatge = ""
   If idiomaclient <> "CA" And idiomaclient <> "ES" And idiomaclient <> "EN" And idiomaclient <> "FR" Then
      etmissatge = "Aquesta direcció d'enviament no te idioma assignat."
      idiomaclient = "EN"
   End If
   etgrupclient = ""
   If atrim(rstc!grupdeclient) <> "" Then etgrupclient = "Grup: " + atrim(rstc!grupdeclient)
   etinfodelclient = atrim(rstc!nom) + Chr(10)
   etinfodelclient = etinfodelclient + atrim(rst!nome) + Chr(10)
   etinfodelclient = etinfodelclient + atrim(rst!poblacioe) + "  (" + atrim(rst!provinciae) + ")" + Chr(10)
   etinfodelclient = etinfodelclient + IIf(rst!pesnetbrut, "PES NET", "") + Chr(10)
   etinfodelclient = etinfodelclient + "Fact(" + atrim(rstclientcomptable!moneda) + ") " + atrim(codiclient) + " - " + atrim(rstclientcomptable!nomclient)
   etinfodelclient.tag = atrim(rst!nome) + " - " + atrim(rst!poblacioe) + " (" + atrim(rst!provinciae) + ") "
   If comprovarcodiclientexisteixasap(codiclient, logoinplacsa.visible) Then
      frame4.BackColor = &HE0E0E0
      frame4.caption = "Dades del Client"
       Else: frame4.BackColor = QBColor(12): frame4.caption = "Dades del Client (Client no existeix a SAP)"
   End If
   If IsDate(datacapcalera.Recordset!dataenvioasap) Then
      ettraspasasap = "Traspasat a SAP el " + Format(datacapcalera.Recordset!dataenvioasap, "dd/mm/yy hh:nn") + vbNewLine + atrim(datacapcalera.Recordset!numalbaraSAP) + "/" + atrim(datacapcalera.Recordset!numfacturasap)
      If NumCmrAssociat(datacapcalera.Recordset!numalbara) = 0 Then ettraspasasap.BackStyle = 1: ettraspasasap.BackColor = &H5C31DD
      activaredicio False
        Else
          activaredicio True
   End If
   If InStr(1, atrim(rst!estilfrontal) + " ", "128") > 0 Then etean128.visible = True
   Set rst = Nothing
   Set rstc = Nothing
   
End Sub
Function NumCmrAssociat(vnumalb As Double) As Double
   Dim rst As Recordset
   NumCmrAssociat = 0
   Set rst = dbvendes.OpenRecordset("select * from transportistes_avisos where numalbara=" + atrim(vnumalb))
   If Not rst.EOF Then NumCmrAssociat = cadbl(rst!numeroavis)
   Set rst = Nothing
End Function
Sub possar_amplesmax_linies()
   Dim objecte As Control
   For Each objecte In formvendes
     If TypeOf objecte Is TextBox Then
        If objecte.Container.Name = "framedadeslinia" Then
           If datalinies.Recordset.Fields(objecte.DataField).Type = 10 Then
                objecte.MaxLength = datalinies.Recordset.Fields(objecte.DataField).Size
           End If
        End If
     End If
   Next
End Sub

Private Sub datalinies_Reposition()
   
   If Not datalinies.Recordset.EOF Then
    If cadbl(datalinies.Recordset!lotinplacsa) <> 0 Then
            vimpostinclosalPVP = carregar_camp_PVPinclos(cadbl(datalinies.Recordset!lotinplacsa))
            dataliniespeu.RecordSource = "select * from liniespeu where numalbara=" + atrim(datacapcalera.Recordset!numalbara) + " order by ordre"
            possaretsobrepaper
            If nomdelcontrolactual <> "llistabobinessel" Then
                 carregar_bobinesentrega cadbl(datalinies.Recordset!lotinplacsa), cadbl(datalinies.Recordset!numalbara)
                 sumarkilosimetresseleccionats
                 sihihanbobinesassignades cadbl(datalinies.Recordset!lotinplacsa), atrim(datalinies.Recordset!tipusdeentrega)
            End If
     End If
      Else:
        etsobrepaper = ""
        dataliniespeu.RecordSource = "select * from liniespeu where numalbara=-9999"
   End If
   dataliniespeu.Refresh
   possartotalpaletsibobines
'   datacapcalera.Refresh
   
   possarinfoIMPOSTalBOX
   Set rstimpost = Nothing
   Set rstc = Nothing
   
End Sub
Function carregar_camp_PVPinclos(vnumc As Double) As Boolean
   Dim rst As Recordset
   vimpostinclosalPVP = False
   Set rst = dbcomandes.OpenRecordset("select PVPimpostinclos from comandes_extres where comanda=" + atrim(vnumc))
   If rst!PVPimpostinclos Then carregar_camp_PVPinclos = True
   Set rst = Nothing
End Function
Sub possarinfoIMPOSTalBOX(Optional vmerma As Boolean)
   Dim rstimpost As Recordset
   Dim rstc As Recordset
   Dim vlinia As String
   Dim vkgmerma As Double
   Dim vcalcultanx100merma As Double
   Dim vkgimpost As Double
   Dim vsumapackinglist As Double
   Dim vTEsp As Double
   Dim vTImp As Double
   Dim vTInt As Double
   Dim vTkm2 As Double
   Dim vTPk As Double
   Dim vTEspM As Double
   Dim vTImpM As Double
   Dim vTIntM As Double
   
   If datalinies.Recordset.EOF Then Exit Sub
   Set rstc = datacapcalera.Database.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(cadbl(datalinies.Recordset!lotinplacsa)))
   If rstc.EOF Then Exit Sub
   If datalinies.Recordset!tipusdeentrega <> "T" And vmerma Then MsgBox "No puc ensenyar la merma d'una comanda entregada PARCIAL.", vbCritical, "ATENCIÓ": Exit Sub
   Set rstimpost = datacapcalera.Database.OpenRecordset("select * from impostenvasos where (comanda=" + atrim(rstc!comanda) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2) + ") and numalbara=" + atrim(datalinies.Recordset!numalbara))
   If Not vmerma Then
            etimpostenvasos = "[T.PK_InT] [T.PK_Imp] [T.PK_Esp][T.PK_Tots][         ]  [Venta_Int][Venta_Imp][Venta_Esp][K/m2]" + vbNewLine
             Else: etimpostenvasos = "[T.PK_InT] [T.PK_Imp] [T.PK_Esp][T.PK_Tots][         ]  [Merma_Int][Merma_Imp][Merma_Esp][K/m2]" + vbNewLine
   End If
   With rstimpost
   While Not .EOF
         vkgimpost = vkgimpost + cadbl(!kgventaImp_mes_esp) + cadbl(!kgventaad_intracom) + cadbl(!kgventaEspanya)
         'vlinia = justificar(Redondejar(cadbl(!Imp_mes_Esp_TKg), 0), 7, "D") + "/" + justificar(Trim(Redondejar(cadbl(!Imp_mes_Esp_Ttanper100), 0)), 3, "E")
         'vlinia = vlinia + "" + justificar(Redondejar(!Imp_mes_Esp_KgNOIMPOST, 0), 7, "D") + "/" + justificar(Trim(Redondejar(![Imp_mes_Esp_%NOIMPOST], 0)), 3, "E")
         vlinia = vlinia + "" + justificar(Redondejar(!Ad_Intracom_KgIMPOST + !KgMermaIMPOST_AD_capa, 0), 8, "D") + "   " '+ "/" + justificar(Trim(Redondejar(![Ad_intracom_%impost], 0)), 3, "E")
         vlinia = vlinia + "" + justificar(Redondejar(!Imp_mes_Esp_KgIMPOST + !KgMermaIMPOST_IE_capa, 0), 8, "D") + "   " '+ "/" + justificar(Trim(Redondejar(![Imp_mes_Esp_%impost], 0)), 3, "E")
         vlinia = vlinia + "" + justificar(Redondejar(cadbl(!KgMermaIMPOST_ES_capa) + cadbl(!Espanya_KgIMPOST), 0), 8, "D") + "   " '+ "/" + justificar(Trim(Redondejar(![Ad_intracom_%impost], 0)), 3, "E")
         vlinia = vlinia + "" + justificar(Redondejar(cadbl(!KgMermaIMPOST_ES_capa) + cadbl(!KgMermaIMPOST_AD_capa) + cadbl(!KgMermaIMPOST_IE_capa) + cadbl(!Espanya_KgIMPOST) + cadbl(!Imp_mes_Esp_KgIMPOST) + cadbl(!Ad_Intracom_KgIMPOST), 0), 8, "D") + "   " '+ "/" + justificar(Trim(Redondejar(![Ad_intracom_%impost], 0)), 3, "E")
         vsumapackinglist = vsumapackinglist + (cadbl(!KgMermaIMPOST_ES_capa) + cadbl(!KgMermaIMPOST_AD_capa) + cadbl(!KgMermaIMPOST_IE_capa) + cadbl(!Espanya_KgIMPOST) + cadbl(!Imp_mes_Esp_KgIMPOST) + cadbl(!Ad_Intracom_KgIMPOST))
         vTEsp = vTEsp + cadbl(!kgventaEspanya): vTImp = vTImp + cadbl(!kgventaImp_mes_esp): vTInt = vTInt + cadbl(!kgventaad_intracom)
         vTPk = cadbl(!aD_iNTRACOM_KgNOIMPOST) + cadbl(!Imp_mes_Esp_KgNOIMPOST) + cadbl(!eSPANYA_KgNOIMPOST)
         vlinia = vlinia + "" + justificar(" ", 7, "D") + " " + justificar(Trim(" "), 3, "E")
        ' vlinia = vlinia + "" + justificar(Redondejar(!Ad_Intracom_KgNOIMPOST, 0), 7, "D") + "/" + justificar(Trim(Redondejar(![Ad_Intracom_%NOIMPOST], 0)), 3, "E")
        ' vlinia = vlinia + "" + justificar(Redondejar(!Ad_Intracom_KgIMPOST, 0), 7, "D") + "/" + justificar(Trim(Redondejar(![Ad_intracom_%impost], 0)), 3, "E")
         vTEspM = vTEspM + cadbl(!kgmermaespanya): vTImpM = vTImpM + cadbl(!kgmermaimp_mes_esp): vTIntM = vTIntM + cadbl(!kgmermaad_intracom)
         If Not vmerma Then
                vlinia = vlinia + "  " + justificar(Redondejar(cadbl(!kgventaad_intracom), 0), 10, "D")
                vlinia = vlinia + "  " + justificar(Redondejar(cadbl(!kgventaImp_mes_esp), 0), 8, "D")
                vlinia = vlinia + "  " + justificar(Redondejar(cadbl(!kgventaEspanya), 0), 10, "D")
                etimpostenvasos.BackColor = &H6BEBB1
                If datalinies.Recordset!tipusdeentrega <> "T" Then etimpostenvasos.BackColor = &HFFFF&
               Else
                  vlinia = vlinia + "  " + justificar(Redondejar(cadbl(!kgmermaad_intracom), 0), 10, "D")
                  vlinia = vlinia + "  " + justificar(Redondejar(cadbl(!kgmermaimp_mes_esp), 0), 8, "D")
                  vlinia = vlinia + "  " + justificar(Redondejar(cadbl(!kgmermaespanya), 0), 10, "D")
                  etimpostenvasos.BackColor = &H5C31DD
         End If
         
         vlinia = vlinia + "   " + justificar(!kgm2, 6, "E")
         vTkm2 = vTkm2 + !kgm2
         vkgmerma = vkgmerma + cadbl(!kgmermaimp_mes_esp) + cadbl(!kgmermaad_intracom) + cadbl(!kgmermaespanya)
         etimpostenvasos = etimpostenvasos + vlinia + vbNewLine
         vlinia = ""
         .MoveNext
   Wend
   For i = 1 To 3 - .RecordCount
      etimpostenvasos = etimpostenvasos + vbNewLine
   Next i
   End With
   
   If vkgimpost > 0 Then vcalcultanx100merma = calculartanxcentmermadellot(datalinies.Recordset!lotinplacsa)
   vTPk = vTPk + vsumapackinglist
   ettanper100merma = "(" + atrim(Redondejar(vcalcultanx100merma, 1)) + "%)"
   etimpostenvasos = etimpostenvasos + "         TOTAL PACKINGLIST " + justificar(atrim(Redondejar(vsumapackinglist, 0)) + "Kg", 7, "E") + "   Merma del " + justificar(atrim(Redondejar(vcalcultanx100merma, 1)), 5, "D") + "%"
   If Not vmerma Then
         etimpostenvasos = etimpostenvasos + "    " + justificar(Redondejar(vTInt, 0), 10, "D") + "" + justificar(Redondejar(vTImp, 0), 10, "D") + "  " + justificar(Redondejar(vTEsp, 0), 10, "D") + "   " + justificar(Redondejar(vTkm2, 5), 6, "E")
          Else: etimpostenvasos = etimpostenvasos + "    " + justificar(Redondejar(vTIntM, 0), 10, "D") + "" + justificar(Redondejar(vTImpM, 0), 10, "D") + "  " + justificar(Redondejar(vTEspM, 0), 10, "D") + "   " + justificar(Redondejar(vTkm2, 5), 6, "E")
   End If
   'ettanper100merma = "(" + atrim(Redondejar(vcalcultanx100merma, 1)) + "%)"
   ettanx100mermalot = ""
   If vTPk > 0 And datalinies.Recordset!tipusdeentrega = "T" Then ettanx100mermalot = "VentaVsFab->   " + atrim(Redondejar((vTPk - vTkilosseleccionatsmesparcials) * 100 / vTPk, 1)) + "%"
   Set rstimpost = Nothing
   Set rstc = Nothing
End Sub

Sub BORRAT_possarinfoIMPOSTalBOX(Optional vmerma As Boolean)
   Dim rstimpost As Recordset
   Dim rstc As Recordset
   Dim vlinia As String
   Dim vkgmerma As Double
   Dim vcalcultanx100merma As Double
   Dim vkgimpost As Double
   If datalinies.Recordset.EOF Then Exit Sub
   Set rstc = datacapcalera.Database.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(cadbl(datalinies.Recordset!lotinplacsa)))
   If rstc.EOF Then Exit Sub
   If datalinies.Recordset!tipusdeentrega <> "T" And vmerma Then MsgBox "No puc ensenyar la merma d'una comanda entregada PARCIAL.", vbCritical, "ATENCIÓ": Exit Sub
   Set rstimpost = datacapcalera.Database.OpenRecordset("select * from impostenvasos where (comanda=" + atrim(rstc!comanda) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2) + ") and numalbara=" + atrim(datalinies.Recordset!numalbara))
   If Not vmerma Then
            etimpostenvasos = "[I_Tkg/%] [I_KgSI/%] [AI_Tkg/%][AI_KgNO/%][AI_KgSI/%] [Kg_VentaI][Kg_VentaAI] [Kg_VentaES][K/m2]" + vbNewLine
             Else: etimpostenvasos = "[I_Tkg/%] [I_KgSI/%] [AI_Tkg/%][AI_KgNO/%][AI_KgSI/%] [Kg_MermaI][Kg_MermaAI] [Kg_MermaES][K/m2]" + vbNewLine
   End If
   With rstimpost
   While Not .EOF
         vkgimpost = vkgimpost + cadbl(!kgventaImp_mes_esp) + cadbl(!kgventaad_intracom) + cadbl(!kgventaEspanya)
         vlinia = justificar(Redondejar(cadbl(!Imp_mes_Esp_TKg), 0), 7, "D") + "/" + justificar(Trim(Redondejar(cadbl(!Imp_mes_Esp_Ttanper100), 0)), 3, "E")
         'vlinia = vlinia + "" + justificar(Redondejar(!Imp_mes_Esp_KgNOIMPOST, 0), 7, "D") + "/" + justificar(Trim(Redondejar(![Imp_mes_Esp_%NOIMPOST], 0)), 3, "E")
         vlinia = vlinia + "" + justificar(Redondejar(!Imp_mes_Esp_KgIMPOST, 0), 7, "D") + "/" + justificar(Trim(Redondejar(![Imp_mes_Esp_%impost], 0)), 3, "E")
         
         vlinia = vlinia + "" + justificar(Redondejar(cadbl(!Ad_Intracom_TKg), 0), 7, "D") + "/" + justificar(Trim(Redondejar(cadbl(!Ad_intracom_Ttanper100), 0)), 3, "E")
         vlinia = vlinia + "" + justificar(Redondejar(!aD_iNTRACOM_KgNOIMPOST, 0), 7, "D") + "/" + justificar(Trim(Redondejar(![Ad_Intracom_%NOIMPOST], 0)), 3, "E")
         vlinia = vlinia + "" + justificar(Redondejar(!Ad_Intracom_KgIMPOST, 0), 7, "D") + "/" + justificar(Trim(Redondejar(![Ad_intracom_%impost], 0)), 3, "E")
         
         If Not vmerma Then
                vlinia = vlinia + "  " + justificar(Redondejar(!kgventaImp_mes_esp, 0), 8, "D")
                vlinia = vlinia + "  " + justificar(Redondejar(!kgventaad_intracom, 0), 10, "D")
                vlinia = vlinia + "  " + justificar(Redondejar(cadbl(!kgventaEspanya), 0), 10, "D")
                etimpostenvasos.BackColor = &H6BEBB1
                If datalinies.Recordset!tipusdeentrega <> "T" Then etimpostenvasos.BackColor = &HFFFF&
               Else
                  vlinia = vlinia + "  " + justificar(Redondejar(!kgmermaimp_mes_esp, 0), 8, "D")
                  vlinia = vlinia + "  " + justificar(Redondejar(!kgmermaad_intracom, 0), 10, "D")
                  vlinia = vlinia + "  " + justificar(Redondejar(cadbl(!kgmermaespanya), 0), 10, "D")
                  etimpostenvasos.BackColor = &H5C31DD
         End If
         
         vlinia = vlinia + "  " + justificar(!kgm2, 7, "D")
         vkgmerma = vkgmerma + cadbl(!kgmermaimp_mes_esp) + cadbl(!kgmermaad_intracom) + cadbl(!kgmermaespanya)
         etimpostenvasos = etimpostenvasos + vlinia + vbNewLine
         .MoveNext
   Wend
   End With
   If vkgimpost > 0 Then vcalcultanx100merma = ((vkgmerma * 100) / vkgimpost)
   ettanper100merma = "(" + atrim(Redondejar(calculartanxcentmermadellot(datalinies.Recordset!lotinplacsa), 1)) + "%)"
   'ettanper100merma = "(" + atrim(Redondejar(vcalcultanx100merma, 1)) + "%)"
   Set rstimpost = Nothing
   Set rstc = Nothing
End Sub
Sub possar_etiqueta_tanx100mermes(vnumalb As Double)
   Dim vcalcultanx100merma As Double
   Dim rstc As Recordset
   Dim rstimpost As Recordset
   Dim vkgmerma As Double
   Dim rstalb As Recordset
   
   ettanper100merma.tag = ""
   Set rstalb = datalinies.Recordset.Clone
   While Not rstalb.EOF
    'Set rstc = datacapcalera.Database.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(cadbl(rstalb!lotinplacsa)))
    'Set rstimpost = datacapcalera.Database.OpenRecordset("select * from impostenvasos where (comanda=" + atrim(rstc!comanda) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2) + ") and numalbara=" + atrim(rstalb!numalbara))
    'vkgmerma = 0
    'vcalcultanx100merma = 0
    'While Not rstimpost.EOF
    '    vkgmerma = vkgmerma + cadbl(rstimpost!kgmermaimp_mes_esp) + cadbl(rstimpost!kgmermaad_intracom) + cadbl(rstimpost!kgmermaespanya)
    '    rstimpost.MoveNext
    'Wend
    vcalcultanx100merma = calculartanxcentmermadellot(rstalb!lotinplacsa)
    If vcalcultanx100merma <= 0 Or vcalcultanx100merma > 17 Then ettanper100merma.tag = ettanper100merma.tag + " [" + atrim(rstalb!lotinplacsa) + "- Merma " + atrim(Redondejar(vcalcultanx100merma, 1)) + "%]" + vbNewLine
    rstalb.MoveNext
   Wend
  Set rstalb = Nothing
  Set rstc = Nothing
  Set rstimpost = Nothing
End Sub
Function sihihanbobinesassignades(numc As Double, tipusdeentrega As String) As Boolean
   Dim rstb As Recordset
   Dim bobines As String
   If tipusdeentrega = "P" Then Exit Function
   Set rstb = dbstocks.OpenRecordset("select * from parcials where not utilitzada and comanda='" + atrim(numc) + "'")
   bobines = ""
   While Not rstb.EOF
     If cadbl(rstb!idpalet) <> 0 Then
        bobines = bobines + "[" + atrim(rstb!idpalet) + "/" + atrim(rstb!idbobina) + "] "
          Else: dbstocks.Execute "delete * from parcials where idpalet=0 and comanda='" + atrim(numc) + "'"
     End If
     rstb.MoveNext
   Wend
   If atrim(bobines) <> "" Then
      MsgBox "No es podrà passar la comanda " + atrim(numc) + " a acabada fins que les següents bobines es passin a disponibles o gastades." + Chr(10) + Chr(13) + bobines, vbInformation, "Atenció"
      sihihanbobinesassignades = True
     Else: sihihanbobinesassignades = False
   End If
End Function

Function nomdelcontrolactual() As String
  On Error Resume Next
  nomdelcontrolactual = formvendes.ActiveControl.Name
End Function
Sub afegirsobrepaper(vlinia As String)
   etsobrepaper = etsobrepaper + IIf(etsobrepaper <> "", Chr(10), "") + vlinia
   llistasobrepaper.AddItem vlinia
End Sub
Sub possaretsobrepaper()
  Dim vmesura As String
  Dim rstdirenvio As Recordset
  Dim vdataproduccio As String
  Dim vmicroperforat As String
  Set rstdirenvio = dbcomandes.OpenRecordset("select * from clients_envios where id=" + atrim(datacapcalera.Recordset!id_direnvio))
  With datalinies.Recordset
  etsobrepaper = ""
  If InStr(1, atrim(!unitatmesura), "/") > 0 Then
     vmesura = Mid(atrim(!unitatmesura), InStr(1, atrim(!unitatmesura), "/") + 1)
    Else: vmesura = atrim(unitatmesura)
  End If
  If vmesura = "FIX" Then vmesura = ""  'excepcio si es preu FIX que no surti a pantalla
  llistasobrepaper.Clear
  vmicroperforat = IIf(atrim(!microperforat) = "S", traducciodeabreviatures("Avmicrop", idiomaclient), vmicroperforat)
  vmicroperforat = IIf(atrim(!macroperforat) = "S", traducciodeabreviatures("Avmacrop", idiomaclient), vmicroperforat)
  'etsobrepaper = "Codi Inplacsa: " + atrim(!codiproducte)
  afegirsobrepaper vmesura + " " + atrim(!descripcioproducte)
  etsobrepaper = etsobrepaper + "          Q: " + atrim(!quantitat) + "  PVP: " + atrim(!preuvenda)
  afegirsobrepaper atrim(!descripciomides) + " " + vmicroperforat
  If atrim(!marcailinia) <> "" And Not rstdirenvio.EOF Then
     If Not cabool(rstdirenvio!albarasensetexteimpresio) Then afegirsobrepaper atrim(!marcailinia)
  End If
  If (atrim(!numcomandacli) + atrim(!refclient)) <> "" Then afegirsobrepaper IIf(atrim(!numcomandacli) <> "", traducciodeabreviatures("COM:", idiomaclient) + atrim(!numcomandacli), "") + "  " + IIf(atrim(!refclient) <> "", " REF: " + atrim(!refclient), "")
  If (atrim(!numcomandaclideclient) + atrim(!refclientdeclient)) <> "" Then afegirsobrepaper IIf(atrim(!numcomandaclideclient) <> "", traducciodeabreviatures("COM:", idiomaclient) + " " + atrim(!numcomandaclideclient), "") + IIf(atrim(!refclientdeclient) <> "", " REF: " + atrim(!refclientdeclient), "")
  If cabool(rstdirenvio!albaraambdataproduccio) Then vdataproduccio = traducciodeabreviatures("AvDataproduccio", idiomaclient) + ": " + atrim(!datafabricacio)
  If atrim(!codibarres) <> "" Then afegirsobrepaper IIf(vdataproduccio + "  " + atrim(!codibarres) <> "", vdataproduccio + "  " + "EAN: " + atrim(!codibarres), "")
  If atrim(!numcalloff) <> "" Or atrim(!numcontracte) <> "" Then
     afegirsobrepaper IIf(atrim(!numcontracte) <> "", "CONT.:  " + atrim(!numcontracte), "") + "  " + IIf(atrim(!numcalloff) <> "", "CALLOFF: " + atrim(!numcalloff), "")
  End If
  
  afegirsobrepaper possar_linia_quantitats
  If cabool(rstdirenvio!albaratotaldetallbobines) Then afegirsobrepaper atrim(generarliniadepackinglist(0, cadbl(!numalbara), cadbl(!lotinplacsa), True))
  If atrim(!observacionslinia) <> "" Then afegirsobrepaper atrim(!observacionslinia)
  If elclientesESPANYOL(datacapcalera.Recordset!numalbara) Then afegirsobrepaper possar_linia_impost(cadbl(!kgimpostenvasos))
  End With
  Set rstdirenvio = Nothing
End Sub
Function ImpostEnv_regimfiscalREFINPLACSA(vrefinplacsa As String) As String
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("select * from tarifes_referencies where refinplacsa='" + atrim(vrefinplacsa) + "'")
  If Not rst.EOF Then ImpostEnv_regimfiscalREFINPLACSA = atrim(rst!impost_regimenfiscal)
  Set rst = Nothing
End Function
Function possar_linia_impost(vkgimpost As Double) As String
  Dim veurokgimpost As Double
  Dim vdescripcioimpost As String
  Dim vRegimFiscal As String
  vRegimFiscal = ImpostEnv_regimfiscalREFINPLACSA(datalinies.Recordset!codiproducte)
  veurokgimpost = cadbl(llegir_ini("General", "PreuImpostEnvasos", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"))
  vdescripcioimpost = "  según Ley 7/2022 de envases no reutilizables."
  If vRegimFiscal <> "" Then
       vdescripcioimpost = "Ley 7/2022 exento por régimen fiscal (Letra " + vRegimFiscal + ")"
       vkgimpost = 0
  End If
  
  possar_linia_impost = IIf(vimpostinclosalPVP = True, "Incluido en PVP: " + atrim(Format(vkgimpost, "#,##0")) + "Kg Base imponible " + vdescripcioimpost, atrim(Format(vkgimpost, "#,##0")) + "Kg Base imp:(" + atrim(veurokgimpost) + "€/Kg) " + vdescripcioimpost)
End Function
Function possar_linia_quantitats() As String
   Dim rst As Recordset
   Dim vespesnet As Boolean
   Dim vlinia As String
   Dim vmoneda As String
   Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda, mesures.unitatinterna, Clients_envios.packinglistalbara, Clients_envios.pesnetbrut FROM (comandes INNER JOIN Clients_envios ON comandes.direnvio = Clients_envios.id) INNER JOIN mesures ON comandes.mesurapvp = mesures.codi WHERE (((comandes.comanda)=" + atrim(cadbl(datalinies.Recordset!lotinplacsa)) + "));")
   If rst.EOF Then Exit Function
   vmoneda = atrim(datacapcalera.Recordset!moneda)
   vespesnet = rst!pesnetbrut
   With datalinies.Recordset
   If atrim(!tipusproducte) <> "" Then
       'formats
        If rst!unitatinterna = "€/U" Then
            vlinia = atrim(!numbobs) + " " + traducciodeabreviatures(UCase(atrim(!tipusproducte)), idiomaclient) + "  " + traducciodeabreviatures("AvLot", idiomaclient) + ": " + atrim(!lotinplacsa)
            'vlinia = atrim(!numbobs) + " " + traducciodeabreviatures(UCase(atrim(!tipusproducte)), idiomaclient) + "  " + Format(!metreslineals, "##0") + " m" + "  " + traducciodeabreviatures("AvLot", idiomaclient) + ": " + atrim(!lotinplacsa)
             Else
                vlinia = atrim(!numbobs) + " " + traducciodeabreviatures(UCase(atrim(!tipusproducte)), idiomaclient) + "  " + IIf(!unitats > 0, formatar(!unitats, "##0") + traducciodeabreviatures("AvPeces", idiomaclient), "") + "  " + traducciodeabreviatures("AvLot", idiomaclient) + ": " + atrim(!lotinplacsa)
        End If
          Else
           'bobines
            If rst!unitatinterna = "€/M" Then
             vlinia = atrim(!numbobs) + " " + traducciodeabreviatures("AvBobines", idiomaclient) + "  " + formatar(Redondejar(cadbl(!kgtotalsbruts), 0), "##0") + " Kg" + "  " + IIf(cadbl(!unitats) > 0, formatar(cadbl(!unitats), "##0") + "  " + traducciodeabreviatures("AvPeces", idiomaclient), "") + "  " + traducciodeabreviatures("AvLot", idiomaclient) + ": " + atrim(!lotinplacsa)
            End If
            If (rst!unitatinterna = "€/K" And vespesnet) Or rst!unitatinterna = "€/B" Then
             vlinia = atrim(!numbobs) + " " + traducciodeabreviatures("AvBobines", idiomaclient) + "  " + formatar(Redondejar(cadbl(!kgtotalsbruts) + cadbl(!pespalets), 1), "##0") + " Kg " + traducciodeabreviatures("Avpesbrut", idiomaclient) + "  " + formatar(!metreslineals, "##0") + " m" + "  " + IIf(!unitats > 0, formatar(cadbl(!unitats), "##0") + "  " + traducciodeabreviatures("AvPeces", idiomaclient), "") + "  " + traducciodeabreviatures("AvLot", idiomaclient) + ": " + atrim(!lotinplacsa)
            End If
            If rst!unitatinterna = "€/M2" Then
             vlinia = atrim(!numbobs) + " " + traducciodeabreviatures("AvBobines", idiomaclient) + "  " + formatar(IIf(cadbl(!kgtotalsnets) = 0, cadbl(!kgtotalsbruts) + IIf(vmoneda = "Dolars", cadbl(!pespalets), 0), cadbl(!kgtotalsnets)), "##0") + " Kg " + "  " + formatar(!metreslineals, "##0") + " m" + "  " + IIf(!unitats > 0, formatar(!unitats, "##0") + "  " + traducciodeabreviatures("AvPeces", idiomaclient), "") + "  " + traducciodeabreviatures("AvLot", idiomaclient) + ": " + atrim(!lotinplacsa)
            End If
            If rst!unitatinterna = "€/K" Or rst!unitatinterna = "€/PROVA" And Not vespesnet Or rst!unitatinterna = "€/FIX" Then
             vlinia = atrim(!numbobs) + " " + traducciodeabreviatures("AvBobines", idiomaclient) + "  " + formatar(cadbl(!metreslineals), "##0") + " m" + "  " + IIf(cadbl(!unitats) > 0, formatar(cadbl(!unitats), "##0") + " " + traducciodeabreviatures("AvPeces", idiomaclient), "") + "  " + traducciodeabreviatures("AvLot", idiomaclient) + ": " + atrim(!lotinplacsa)
            End If
            If rst!unitatinterna = "€/KM" Then
             vlinia = atrim(!numbobs) + " " + traducciodeabreviatures("AvBobines", idiomaclient) + "  " + formatar(cadbl(!kgtotalsbruts), "##0") + " Kg" + "  " + IIf(cadbl(!unitats) > 0, formatar(cadbl(!unitats), "##0") + " " + traducciodeabreviatures("AvPeces", idiomaclient), "") + "  " + traducciodeabreviatures("AvLot", idiomaclient) + ": " + atrim(!lotinplacsa)
            End If
            If rst!unitatinterna = "€/U" Or rst!unitatinterna = "€/1000U" Then
             vlinia = atrim(!numbobs) + " " + traducciodeabreviatures("AvBobines", idiomaclient) + "  " + formatar(!kgtotalsbruts, "##0") + " Kg" + "  " + formatar(!metreslineals, "##0") + " m" + "  " + traducciodeabreviatures("AvLot", idiomaclient) + ": " + atrim(!lotinplacsa)
             If vespesnet Then vlinia = atrim(!numbobs) + " " + traducciodeabreviatures("AvBobines", idiomaclient) + "  " + formatar(cadbl(!kgtotalsnets), "##0") + " Kg  (" + formatar(Redondejar(cadbl(!kgtotalsbruts) + cadbl(!pespalets), 1), "##0") + " Kg " + traducciodeabreviatures("Avpesbrut", idiomaclient) + ")  " + formatar(cadbl(!metreslineals), "##0") + " m" + "  " + traducciodeabreviatures("AvLot", idiomaclient) + ": " + atrim(!lotinplacsa)
            End If
   End If
   
   End With
   Set rst = Nothing
   possar_linia_quantitats = vlinia
End Function
Function formatar(valor As Double, vformat As String) As String
   Dim i As Double
   Dim valorenter As String
   Dim valordecimal As String
   Dim vsimboldecimal As String
   Dim vsimbolmiler As String
   
   vsimboldecimal = "."
   vsimbolmiler = ","
   If InStr(1, Trim(CDbl(1 / 2)), ",") Then vsimbolmiler = ".": vsimboldecimal = ","
   
   'valor = 1234.56
   valorenter = Int(valor)
   valordecimal = valor - Int(valor)
   For i = Len(valorenter) To 1 Step -1
      If (Len(valorenter) - i) Mod 3 = 0 And (Len(valorenter) - i) > 0 Then
        valorfinal = Mid(valorenter, i, 1) + vsimbolmiler + valorfinal
          Else:
            valorfinal = Mid(valorenter, i, 1) + valorfinal
      End If
   Next i
   valorfinal = valorfinal + IIf(valordecimal > 0, vsimboldecimal + Mid(valor, Len(valorenter) + 2), "")
   formatar = valorfinal
End Function
Private Sub eliminar_Click()
   Dim vnumalbara As Double
  If mirarsieliminaralbara Then Exit Sub
  If datalinies.Recordset.EOF Then MsgBox "Escull primer una linia per eliminar", vbCritical, "Error": Exit Sub
  If datalinies.Recordset.EditMode <> 0 Then Exit Sub
  If InputBox("Segur que vols eliminar la linia d'albarà que fa referència a la comanda: " + atrim(datalinies.Recordset!lotinplacsa) + "?" + Chr(10) + "Escriu el número de comanda per eliminar-la", "Borrar linia d'albarà") <> atrim(datalinies.Recordset!lotinplacsa) Then Exit Sub
  vnumalbara = cadbl(cnumalbara)
'actualitzo les bobines d'entrega per treure el vincle amb aquest albarà
  dbbaixes.Execute "update bobinesent set numalbara=null where numalbara=" + atrim(vnumalbara) + " and comanda=" + atrim(datalinies.Recordset!lotinplacsa) + " and numalbara=" + atrim(cadbl(datalinies.Recordset!numalbara))
'actualitzo Impost d'envasos per treure el vincle amb aquest albarà i lot inplacsa
  datalinies.Database.Execute "delete * from impostenvasos where idliniaalbara=" + atrim(datalinies.Recordset!ID)
'elimino la linia de l'albarà
  datalinies.Recordset.Delete
  datalinies.Refresh
  mirarsieliminaralbara
  refrescarnumerosdalbara
End Sub
Function mirarsieliminaralbara() As Boolean
   Dim vnumalbara As Double
   If datalinies.Recordset.EOF Then
    If MsgBox("No hi ha comandes afegides a aquest albarà vols eliminar-lo", vbInformation + vbYesNo + vbDefaultButton2, "Eliminar albarà") = vbYes Then
        vnumalbara = cadbl(datacapcalera.Recordset!numalbara)
        datacapcalera.Database.Execute "delete * from liniespeu where numalbara=" + atrim(vnumalbara)
        datacapcalera.Recordset.Delete
        datacapcalera.Refresh
        datalinies.Database.Execute "delete * from impostenvasos where numalbara=" + atrim(vnumalbara)
        mirarsieliminaralbara = True
    End If
  End If
End Function

Function comprovarcodiclientexisteixasap(codicomptableclient As Double, esinplacsa As Boolean) As Boolean
  Dim rst As Recordset
  If esinplacsa Then
       Set rst = dbcomandes.OpenRecordset("select * from clients_codissap where codisap=" + atrim(codicomptableclient))
        Else: Set rst = dbcomandes.OpenRecordset("select * from Clients_CodisSAPPlasel where codisap=" + atrim(codicomptableclient))
  End If
  If rst.EOF Then
     comprovarcodiclientexisteixasap = False
      Else: comprovarcodiclientexisteixasap = True
  End If
  Set rst = Nothing
End Function

Sub assignardecimalipunt()
  Dim LocalID As Long
  If existeix("c:\ordprog.ini") Then Exit Sub
  LocalID = GetUserDefaultLCID()
  SetLocaleInfo LocalID, LOCALE_SDECIMAL, "."
  SetLocaleInfo LocalID, LOCALE_STHOUSAND, ","
End Sub

Sub possartotalpaletsibobines()
   Dim rst As Recordset
   Dim vsumapalets As Double
   Dim vkgtotalsbruts As Double
   Dim vbases As String
   Dim vNbases As Double
   Dim vx As String
   Dim rstc As Recordset
   Dim vkg As Double
   Dim vdifkg As Double
   Dim vkgparcials As Double
   
   vkgtotalsbruts = 0
   etdemanats.tag = ""
   etdemanats = ""
   etentregaparcial = ""
   etresumbobinesipalets = ""
   If cadbl(cnumalbara) = 0 Then Exit Sub
   Set rst = dbbaixes.OpenRecordset("SELECT DISTINCT [comanda] & [numpalet] AS Expr1 From bobinesent Where (((bobinesent.numalbara) = " + cnumalbara + "))")
   If Not rst.EOF Then
      rst.MoveLast
      vsumapalets = rst.RecordCount
   End If
   Set rst = datacapcalera.Database.OpenRecordset("SELECT liniesalbara.numalbara, Sum(liniesalbara.kgtotalsbruts) AS kgtotals From liniesalbara where numalbara=" + cnumalbara + " GROUP BY liniesalbara.numalbara;")
   If Not rst.EOF Then vkgtotalsbruts = Redondejar(rst!kgtotals, 0)
   
   vbases = IIf(cadbl(datacapcalera.Recordset!numbases) > 0, " NºBases: " + atrim(cadbl(datacapcalera.Recordset!numbases)), "")
   etresumbobinesipalets = atrim(vkgtotalsbruts) + " Kg Totals  -  " + atrim(vsumapalets) + " Palets " + vbases
   If escrops Then vNbases = 0 Else vNbases = cadbl(datacapcalera.Recordset!numbases)
   vresum = IIf(cadbl(vNbases) > 0, atrim(vNbases) + " BASES (" + atrim(vsumapalets) + " PALETS) ", atrim(vsumapalets) + " PALETS ") + "   " + atrim(vkgtotalsbruts) + " KG"
   datacapcalera.Database.Execute "update capcaleraalbara set numpalets=" + atrim(vsumapalets) + " where numalbara=" + atrim(cnumalbara)
   datacapcalera.Database.Execute "update capcaleraalbara set kg=" + atrim(vkgtotalsbruts) + " where numalbara=" + atrim(cnumalbara)
   If Not datalinies.Recordset.EOF Then Set rstc = dbcomandes.OpenRecordset("select * from comandesmesextres where comanda=" + atrim(datalinies.Recordset!lotinplacsa))
   If Not IsDate(datacapcalera.Recordset!dataenvioasap) Then
        If vresum <> "" And cobservacionstransport = "" Or (Mid(cobservacionstransport + " ", 1, 1) = "[" And InStr(1, cobservacionstransport + " ", "]") > 0) Then
          vx = cobservacionstransport + " "
          vx = Mid(vx, InStr(1, vx, "]") + 1)
          vx = atrim("[" + vresum + "] " + vx)
          datacapcalera.Database.Execute "update capcaleraalbara set observacionsports='" + atrim(vx) + "' where numalbara=" + atrim(cnumalbara)
          If atrim(vx) <> cobservacionstransport Then actualitzarpeudepagina
        End If
   End If
   datacapcalera.UpdateControls
   If Not datalinies.Recordset.EOF Then
        vkg = cadbl(rstc!totalspesMesTA)
        If vkg = 0 Then vkg = cadbl(rstc!tubbaseext)
        If InStr(rstc!ruta, "S") > 0 Then vkg = Redondejar(calcularpesxrpeça(rstc) * cadbl(rstc!cantitatsol), 0)
        vkgparcials = buscarkgentreguesparcials(rstc!comanda, cnumalbara)
        etdemanats = "Kg Comanda: " + atrim(Redondejar(vkg, 1))
        If vkgparcials > 0 Then etentregaparcial = "Ent. parcials: " + atrim(Redondejar(vkgparcials, 0)) + "Kg"
        etdemanats.tag = atrim(vkg)
        
        vdifkg = vkg - (cadbl(datalinies.Recordset!kgtotalsbruts) + vkgparcials)
        vTkilosseleccionatsmesparcials = cadbl(datalinies.Recordset!kgtotalsbruts) + vkgparcials
        Framebobines.BackColor = &HC0FFC0
        If vkg > 0 Then If vdifkg * 100 / vkg > 10 Or vdifkg * 100 / vkg < -10 Then Framebobines.BackColor = &H8080FF
   End If
   Set rst = Nothing
   Set rstc = Nothing

End Sub
Function buscarkgentreguesparcials(vnumc As Double, vnumalb As Double) As Double
   Dim rst As Recordset
   Set rst = datacapcalera.Database.OpenRecordset("select sum(kgtotalsbruts) as TKg from liniesalbara where lotinplacsa=" + atrim(vnumc) + " and numalbara<>" + atrim(vnumalb))
   buscarkgentreguesparcials = cadbl(rst!TKg)
   Set rst = Nothing
End Function

Function escrops() As Boolean
   Dim rst As Recordset
   Set rst = datacapcalera.Database.OpenRecordset("select codi from clients_envios where id=" + atrim(cadbl(datacapcalera.Recordset!id_direnvio)))
   If Not rst.EOF Then If rst!codi = 6841 Then escrops = True
   Set rst = Nothing
End Function
Private Sub etdataenviament_DblClick()
    Dim vdata
    vdata = InputBox("Entra la data d'entrega del material." + Chr(10) + "   Escriu [CAP] per treure la data.", "Canvi de data", Format(Now, "dd/mm/yy"))
    If vdata = "" Then Exit Sub
    If Not IsDate(vdata) And UCase(vdata) <> "CAP" Then MsgBox "Data erronea.", vbCritical, "Error": Exit Sub
    If UCase(vdata) = "CAP" Then
        dbbaixes.Execute "update bobinesent set dataentrega=null where numalbara=" + atrim(cadbl(cnumalbara))
      Else: dbbaixes.Execute "update bobinesent set dataentrega=#" + atrim(Format(vdata, "mm/dd/yy")) + "# where numalbara=" + atrim(cadbl(cnumalbara))
    End If
    datacapcalera.Recordset.Move 0
End Sub

Private Sub etfiltreactivat_DblClick()
  datacapcalera.RecordSource = "select * from capcaleraalbara order by numalbara desc"
  datacapcalera.Refresh
  etfiltreactivat = ""
End Sub

Private Sub etfiltreactivat_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If etfiltreactivat <> "" Then
         etfiltreactivat.MousePointer = 99
          Else: etfiltreactivat.MousePointer = 0
   End If
End Sub

Private Sub etfiproduccio_DblClick()
  MsgBox etfiproduccio
End Sub

Private Sub etimpostenvasos_DblClick()
    If etimpostenvasos.BackColor <> &H5C31DD Then
          possarinfoIMPOSTalBOX True
            Else: possarinfoIMPOSTalBOX False
    End If
End Sub

Private Sub etmetrescubics_DblClick()
   demanar_metrescubicstransport
   Command6_Click
End Sub

Private Sub etmetrescubicscalculats_DblClick()
Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
  formseleccio.Data1.RecordSource = "SELECT * FROM embolicarpalets Where numalbara=" + atrim(datacapcalera.Recordset!numalbara) + " order by numreferenciagrup"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).visible = False
  'formseleccio.DBGrid2.Columns(2).width = 900
  formseleccio.width = 10000
  If formseleccio.Data1.Recordset.RecordCount > 1 Then
     formseleccio.Show 1
    Else: MsgBox "No hi ha informació per ensenyar", vbCritical, "Error"
  End If
  Unload formseleccio
End Sub

Private Sub Form_Activate()
   If vObrirAlbaràSAP > 0 Then
        datacapcalera.Recordset.FindFirst "numalbaraSAP=" + atrim(vObrirAlbaràSAP) + ""
        vObrirAlbaràSAP = 0
   End If
   'Dim rstc As Recordset
   'Dim rstalb As Recordset
   'Set rstalb = datacapcalera.Database.OpenRecordset("select * from liniesalbara where kgimpostenvasos>0")
   'Set rstc = dbcomandes.OpenRecordset("select * from comandes")
   'While Not rstalb.EOF
   '  rstc.FindFirst "comanda=" + atrim(rstalb!lotinplacsa)
   '  If Not rstc.NoMatch Then
   '     datacapcalera.Database.Execute "update impostenvasos set idliniaalbara=" + atrim(rstalb!ID) + " where numalbara=" + atrim(rstalb!numalbara) + " and (comanda=" + atrim(rstc!comanda) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2) + ")"
   '  End If
   '  rstalb.MoveNext
   'Wend
End Sub
Function calculartanxcentmermadellot(vnumc As Double) As Double
   Dim rstimpost2 As Recordset
   Dim rstc As Recordset
   Dim vkgtotalimpostPK As Double
   Dim vkgtotalimpostVENTA As Double
   Dim vidliniaalbaraPRIMERA As Double
   
   Set rstc = datacapcalera.Database.OpenRecordset("select linkcomanda1,linkcomanda2,comanda from comandes where comanda=" + atrim(vnumc))
   'Set rstimpost2 = datacapcalera.Database.OpenRecordset("select * from impostenvasos where (kgmermaad_intracom>0 or kgmermaimp_mes_esp>0 or kgmermaespanya>0) and (comanda=" + atrim(vnumc) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2) + ")")
   Set rstimpost2 = datacapcalera.Database.OpenRecordset("select * from impostenvasos where  (comanda=" + atrim(vnumc) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2) + ") order by comanda ASC")
   
   If Not rstimpost2.EOF Then vidliniaalbaraPRIMERA = cadbl(rstimpost2!idliniAaLbara)
   'Set rstimpost2 = datacapcalera.Database.OpenRecordset("select * from impostenvasos where idliniaalbara=" + atrim(cadbl(rstimpost2!idliniaalbara)))
   While Not rstimpost2.EOF
        vkgmerma = vkgmerma + cadbl(rstimpost2!kgmermaimp_mes_esp) + cadbl(rstimpost2!kgmermaad_intracom) + cadbl(rstimpost2!kgmermaespanya)
        vkgtotalimpostVENTA = vkgtotalimpostVENTA + cadbl(rstimpost2!kgventaEspanya) + cadbl(rstimpost2!kgventaImp_mes_esp) + cadbl(rstimpost2!kgventaad_intracom)   'vkgtotalimpost + cadbl(rstimpost2!Imp_mes_Esp_KgIMPOST) + cadbl(rstimpost2!Ad_Intracom_KgIMPOST) + cadbl(rstimpost2!Espanya_KgIMPOST) + cadbl(rstimpost2!KgMermaIMPOST_IE_capa) + cadbl(rstimpost2!KgMermaIMPOST_AD_capa) + cadbl(rstimpost2!KgMermaIMPOST_ES_capa)
        'vkgtotalimpostPK = vkgtotalimpostPK + (cadbl(rstimpost2!Espanya_KgIMPOST) + cadbl(rstimpost2!eSPANYA_KgNOIMPOST) + rstimpost2!Imp_mes_Esp_KgIMPOST + rstimpost2!Imp_mes_Esp_KgNOIMPOST + rstimpost2!Ad_Intracom_KgIMPOST + rstimpost2!aD_iNTRACOM_KgNOIMPOST) + cadbl(rstimpost2!KgMermaIMPOST_AD_capa) + cadbl(rstimpost2!KgMermaIMPOST_IE_capa) + cadbl(rstimpost2!KgMermaIMPOST_ES_capa)
        If rstimpost2!idliniAaLbara = vidliniaalbaraPRIMERA Then vkgtotalimpostPK = vkgtotalimpostPK + cadbl(rstimpost2!Espanya_KgIMPOST) + rstimpost2!Imp_mes_Esp_KgIMPOST + rstimpost2!Ad_Intracom_KgIMPOST + cadbl(rstimpost2!KgMermaIMPOST_AD_capa) + cadbl(rstimpost2!KgMermaIMPOST_IE_capa) + cadbl(rstimpost2!KgMermaIMPOST_ES_capa)
        rstimpost2.MoveNext
     Wend
   If vkgtotalimpostVENTA > 0 Then calculartanxcentmermadellot = 100 * (vkgtotalimpostPK - vkgtotalimpostVENTA) / vkgtotalimpostPK
   Set rstimpost2 = Nothing
   Set rstc = Nothing
End Function
Sub comprovarsihihafitxerspendentsdimportarasap(Optional vhihainplacsa As Boolean, Optional vhihaplasel As Boolean, Optional vnoavisar As Boolean, Optional vseidor As Boolean)
   Dim vruta As String
   Dim vcontador As Integer
   Dim vdir As String
   Dim vmsg As String
   On Error GoTo errorsap
   vruta = llegir_ini("Vendes", "rutasap_INPLACSA", "comandes.ini")
   If llegir_ini("Vendes", "rutaSapSeidor_INPLACSA", "comandes.ini") = "{[}]" Then
     escriure_ini "Vendes", "rutaSapSeidor_INPLACSA", "\\servidorsap\seidor_COMUNICADOR\ENTALBVENTAS\INPLACSA", "comandes.ini"
     escriure_ini "Vendes", "rutaSapSeidor_PLASEL", "\\servidorsap\seidor_COMUNICADOR\ENTALBVENTAS\PLASEL", "comandes.ini"
   End If
   If vseidor Then vruta = llegir_ini("Vendes", "rutaSapSeidor_INPLACSA", "comandes.ini")
   mirarsihihafitxers vruta + "\V-*.csv", vcontador
   If vcontador > 0 Then vhihainplacsa = True: vmsg = "  Hi ha fitxers de INPLACSA pendents d'importar a SAP."
   vruta = llegir_ini("Vendes", "rutasap_PLASEL", "comandes.ini")
   If vseidor Then vruta = llegir_ini("Vendes", "rutaSapSeidor_PLASEL", "comandes.ini")
   mirarsihihafitxers vruta + "\V-*.csv", vcontador
   If vcontador > 0 Then vhihaplasel = True: vmsg = vmsg + Chr(10) + "  Hi ha fitxers de PLASEL pendents d'importar a SAP."
   If vmsg <> "" And Not vnoavisar Then MsgBox vmsg, vbCritical, "Atenció"
   
   Exit Sub
errorsap:
     MsgBox "L'usuari " + Environ("username") + " no te acces al servidor de SAP no es podran enviar albarans."
     Command2.Enabled = False
End Sub
Sub mirarsihihafitxers(vruta As String, vcontador As Integer)
   vdir = Dir(vruta)
   vcontador = 0
   While vdir <> ""
     If vdir <> "." And vdir <> ".." Then vcontador = vcontador + 1
     vdir = Dir
   Wend
End Sub

Private Sub Form_Click()

'  generar_CSV_dArta datacapcalera.Recordset!numalbara
'If Not toteslesbobinesdonadesdebaixaocanviunitatpvp Then Exit Sub

'Command1_Click
'comprovarsiesTotalihihabobinessenseseleccionar
 ' actualitzar_numerosalbaraSAP_a_produccio
 ' actualitzar_numerosfacturaSAP_a_produccio
'activaredicio True
'   mirar_resultat_importacio "\\servidorsap\seidor_COMUNICADOR\LOG\Inplacsa"
 '  Dim v As Double
 '  v = InputBox("Comanda")
 '  comprovarsihihaclixesperenviar v
 'comprovarsicalpackinglist
End Sub
Sub generar_CSV_dArta(vnumalbara As String)
  Dim rst As Recordset
  Dim rstpf As Recordset
  Dim rstc As Recordset
  Dim vsql As String
  Dim vnomfitxer As String
  Dim vlinia As String
  vnomfitxer = "c:\temp\CSV_ARTA_" + vnumalbara + ".CSV"
  If existeix(vnomfitxer) Then Kill vnomfitxer
  vsql = "SELECT bobinesent.comanda, bobinesent.numpalet FROM bobinesent "
  vsql = vsql + " WHERE (((bobinesent.numalbara)=" + vnumalbara + ")) GROUP BY bobinesent.comanda, bobinesent.numpalet ORDER BY bobinesent.comanda, bobinesent.numpalet;"
  Set rstc = dbvendes.OpenRecordset("select * from comandes")
  Set rst = dbvendes.OpenRecordset(vsql)
  If Not rst.EOF Then
      Open vnomfitxer For Output As #2
      Print #2, "CustomerOrderNumber;MaterialNumber;Description;SSCC_Code;ProductionDate;BestBeforeDate;Quantity;UnitOfMeasure;LotNumber;;;;"
      Else: GoTo fi
  End If
  While Not rst.EOF
    Set rstpf = dbvendes.OpenRecordset("select * from papersfrontals where numlotinplacsa=" + atrim(rst!comanda) + " and numpalet=" + atrim(rst!numpalet))
    rstc.FindFirst "comanda=" + atrim(rst!comanda)
    While Not rstpf.EOF
       vlinia = Mid(rstc!comandaclient, 1, 10) + ";" + Mid(rstc!refclient, 1, 40) + ";" + Mid(rstpf!texteimp, 1, 40) + ";" + Mid(rstpf!scc, 1, 20) + ";" + atrim(rstpf!datafabricacio) + ";" + atrim(rstpf!datacaducitat)
       vlinia = vlinia + ";" + passaradecimalpunt(atrim(rstpf!pesnet)) + ";KG;" + atrim(rstc!comanda) + ";;;;"
       Print #2, vlinia
       rstpf.MoveNext
    Wend
    rst.MoveNext
  Wend
  Close #2
  If existeix(vnomfitxer) Then enviaremailgenericambadjunt "recepcio@inplacsa.com", "CSV per enviar a d'ARTA", "", vnomfitxer, True
        
fi:
  Set rstc = Nothing
  Set rst = Nothing
  Set rstpf = Nothing
End Sub
Private Sub Form_DblClick()
'activar_frames True
'activaredicio True
End Sub

Private Sub Form_GotFocus()
  assignardecimalipunt
End Sub

Private Sub Form_Initialize()
   
   If Not datacapcalera.Recordset.EOF Then datacapcalera.Recordset.MoveLast: datacapcalera.Recordset.MoveFirst
End Sub

Private Sub Form_Load()
assignardecimalipunt
arguments = ObtenerLíneaComando
fitxerini = "comandes.ini"
If InStr(1, LCase(atrim(arguments(1))), ".ini") > 0 Then fitxerini = atrim(arguments(1))
  cami = llegir_ini("General", "cami", fitxerini)
  camistock = rutadelfitxer(cami) + "palets.mdb"
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
  If existeix("c:\ordprog.ini") Then
     cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
     Command15.visible = True
  End If
 If LCase(arguments(1)) = "obriralbara" Then vObrirAlbaràSAP = cadbl(arguments(2))
  centerscreen Me
  Set dbbaixes = DBEngine.OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
  Set dbcomandes = DBEngine.OpenDatabase(cami)
  Set dbclixes = DBEngine.OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  Set dbstocks = DBEngine.OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
  Set dbvendes = DBEngine.OpenDatabase(rutadelfitxer(cami) + "vendes.mdb")
  Set dbimpost = DBEngine.OpenDatabase(rutadelfitxer(cami) + "ImpostEnvasos.mdb")
  datalinies.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
  datacapcalera.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
  dataliniespeu.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
  'datacapcalera.RecordSource = "select * from capcaleraalbara where numalbara=-99999"
  activar_frames False
  escullirtransportista
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set dbbaixes = Nothing
  Set dbcomandes = Nothing
  Set dbclixes = Nothing
End Sub

Private Sub Framecontrolslinia_DblClick()
  If UCase(InputBox("Escriu el password")) = "INPLACSA" Then activaredicio True
End Sub

Private Sub llistabobinessel_Click()
  If seltot.tag = "1" Then Exit Sub
  
  sumarkilosimetresseleccionats
  actualitzartotalsdelinia
End Sub
Sub actualitzartotalsdelinia()
  Dim rstc As Recordset
  If datalinies.Recordset.EOF Then MsgBox "No hi ha cap linia sel.leccionada", vbCritical, "Error": Exit Sub
  Set rstc = dbcomandes.OpenRecordset("SELECT comandes.*, comandes_extres.codicomptable,comandes_extres.solpesgrmcm2  FROM comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda where comandes.comanda =" + atrim(datalinies.Recordset!lotinplacsa))
  If rstc.EOF Then Exit Sub
  If datalinies.Recordset.EditMode = 0 Then Exit Sub 'datalinies.Recordset.Edit
  'possar_quantitats_bobines rstc
  possar_kilos_metres_unitats_etc rstc
  datalinies.Recordset!quantitat = triarelvalordepenguentdelaunitat
  datalinies.Recordset.Update
  datalinies.Recordset.Move 0
  datalinies.Recordset.Edit
  possaretsobrepaper
  Set rstc = Nothing
End Sub

Sub sumarkilosimetresseleccionats()
   Dim i As Integer
   Dim tkilos As Double
   Dim tmetres As Double
   Dim rst As Recordset
   Dim vmesura As String
   Dim vnumcalloff As String
   vTkilosseleccionats = 0
   vTkilosseleccionatsmesparcials = 0
   If cadbl(llistabobinessel.tag) = 0 Then Exit Sub
   If Not datalinies.Recordset.EOF Then vnumcalloff = atrim(datalinies.Recordset!numcalloff)
   dbbaixes.Execute "update bobinesent set numalbara=null where comanda=" + atrim(cadbl(llistabobinessel.tag)) + " and numalbara=" + atrim(cadbl(cnumalbara))
   Set rst = dbbaixes.OpenRecordset("select * from bobinesent where comanda=" + atrim(cadbl(llistabobinessel.tag)) + " and (numalbara=" + atrim(cadbl(cnumalbara)) + " or numalbara=null or numalbara=0) order by numbob")
   For i = 0 To llistabobinessel.ListCount - 1
     If llistabobinessel.Selected(i) Then
        rst.FindFirst "numbob=" + atrim(llistabobinessel.ItemData(i))
        If Not rst.NoMatch Then
            tmetres = tmetres + cadbl(rst!metresisacs)
            tkilos = tkilos + Redondejar(cadbl(rst!kilosiunitats), 1)
            rst.Edit
            rst!numalbara = cnumalbara
            rst!numcalloff = Mid(vnumcalloff, 1, 30)
            rst.Update
        End If
     End If
   Next i
   Set rst = dbbaixes.OpenRecordset("SELECT  productes.ruta FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi where comanda=" + atrim(cadbl(llistabobinessel.tag)))
   vmesura = IIf(InStr(1, rst!ruta, "S"), "U", "M")
   ettotals = "Total: " + formatar(tmetres, "##0") + vmesura + "   " + formatar(tkilos, "##0.0") + "Kg"
   vTkilosseleccionats = tkilos
   Set rst = Nothing
   If vmesura = "U" Then
      If cadbl(datalinies.Recordset!quantitat) <> tmetres Then
          vcancelarSAP = True
          If vObrirAlbaràSAP = 0 Then MsgBox ("Comanda " + atrim(datalinies.Recordset!lotinplacsa) + vbNewLine + "La quantitat de Unitats seleccionades son diferents de les entrades a l'albarà. " + vbNewLine + "Edita i guarda una altra vegada aquesta linia d'albarà.")
      End If
   End If
End Sub

Sub activar_frames(activar As Boolean)
  framedadeslinia.Enabled = activar
  frameliniesalpaper.Enabled = activar
  'Framelinies.Enabled = activar
  Framepeualbara.Enabled = activar
  Framebobines.Enabled = activar
End Sub

Private Sub llistatentreguestransport_Click()
  Dim vdata As Date
  Dim vresp As String
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  
  vresp = InputBox("Entra la data que vols treure el llistat.", "Data entregues", Format(proximdianatural, "dd/mm/yy"))
  If StrPtr(vresp) = 0 Then Exit Sub
  If Not IsDate(vresp) Then Exit Sub
  vdata = vresp
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistatpackinglistentregues.rpt", 1)
  'oreport.RecordSelectionFormula = "{capcaleraalbara.numalbara} = " + numalb + Chr(13) + " and {liniesalbara.numalbara} = " + numalb + Chr(13) + " and {bobinesent.numalbara}=" + numalb
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "vendes.mdb"
  oreport.DiscardSavedData
  oreport.VerifyOnEveryPrint = True
  oreport.FormulaFields.GetItemByName("dataentrega").Text = "#" + Format(vdata, "mm/dd/yy") + "#"
  oreport.VerifyOnEveryPrint = False
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1, Me
    Else
      oreport.PrintOut False, 1
  End If
     
End Sub

Private Sub malbaransSAPsegellats_Click()
   Unload formescanejaralbaransproveidor
   Load formescanejaralbaransproveidor
   formescanejaralbaransproveidor.ettipusescaneig.caption = "Esperant Albarans SAP o CMR segellats..."
   formescanejaralbaransproveidor.tag = "albaransSAP"
   formescanejaralbaransproveidor.checktotselsproveidors.visible = False
   formescanejaralbaransproveidor.bcmr.visible = True
   formescanejaralbaransproveidor.caption = formescanejaralbaransproveidor.ettipusescaneig.caption
   DoEvents
   formescanejaralbaransproveidor.Show 1
'   passar_albara_a_enviat vnumalb, Now
End Sub

Private Sub malbcontenidors_Click()
  formalbaracontenidors.Show 1
End Sub

Private Sub memailstransportistes_Click()
   
End Sub

Private Sub malbSAP_Click()
  Dim vnomfitxer As String
  vnomfitxer = "\\ord_copies\AlbaransSAPClients\" + atrim(datacapcalera.Recordset!numalbaraSAP) + ".pdf"
  If existeix(vnomfitxer) Then
      obrir_document vnomfitxer
        Else: MsgBox "No trobo l'albarà escanejat.", vbExclamation, "Atenció"
  End If
End Sub

Private Sub mcertificaciolotsproveidor_Click()
   Unload formescanejaralbaransproveidor
   Load formescanejaralbaransproveidor
   formescanejaralbaransproveidor.ettipusescaneig.caption = "Esperant Certificats dels Albarans del Proveïdor..."
   formescanejaralbaransproveidor.tag = "certificats"
   formescanejaralbaransproveidor.caption = formescanejaralbaransproveidor.ettipusescaneig.caption
   DoEvents
   formescanejaralbaransproveidor.Show 1
End Sub

Private Sub mCMR_Click()
  Dim vnomfitxer As String
  Dim rst As Recordset
  Dim vnumalb As String
  vnumalb = atrim(datacapcalera.Recordset!numalbara)
  Set rst = datacapcalera.Database.OpenRecordset("select numeroavis from transportistes_avisos where numalbara=" + vnumalb)
  If rst.EOF Then MsgBox "No trobo numero de CMR per aquest albarà.", vbCritical, "Error": GoTo fi
  vnomfitxer = "\\ord_copies\AlbaransSAPClients\CMR_" + atrim(rst!numeroavis) + ".pdf"
  If Not existeix(vnomfitxer) Then vnomfitxer = "\\ord_copies\AlbaransSAPClients\CMRs\CMR_" + atrim(rst!numeroavis) + ".pdf"
  If existeix(vnomfitxer) Then
      obrir_document vnomfitxer
        Else: MsgBox "No trobo el CMR escanejat.", vbExclamation, "Atenció"
  End If
fi:
  Set rst = Nothing
End Sub


Sub CMRsNoEscanejats(vcriteri As String)
Dim rst As Recordset
   Dim vcmrsnotrobats As String
   Dim vruta As String
   Dim vmes As String
   Dim vany As String
   vmes = InputBox("Entra el mes que vols revisar:", "Mes", Month(Now))
   vany = InputBox("Entra l'any que vols revisar:", "Any", Year(Now))
   If vcriteri = "Tots" Then Set rst = dbvendes.OpenRecordset("Select * from transportistes_avisos where month([datarecullida])=" + atrim(cadbl(vmes)) + " and year([datarecullida])=" + atrim(cadbl(vany)))
   If vcriteri = "Marcats" Then Set rst = dbvendes.OpenRecordset("Select * from transportistes_avisos where escanejat=true and month([datarecullida])=" + atrim(cadbl(vmes)) + " and year([datarecullida])=" + atrim(cadbl(vany)))
   While Not rst.EOF
      vruta = "\\ord_copies\AlbaransSAPClients\CMRs\CMR_" + atrim(rst!numeroavis) + ".pdf"
      If Not existeix(vruta) Then
          If InStr(1, vcmrsnotrobats, atrim(rst!numeroavis)) = 0 Then
           vcmrsnotrobats = vcmrsnotrobats + IIf(vcmrsnotrobats <> "", ",", "") + "'" + atrim(rst!numeroavis) + "'"
          End If
      End If
      rst.MoveNext
   Wend
   If vcmrsnotrobats <> "" Then
        MsgBox vcmrsnotrobats
        Clipboard.Clear
        Clipboard.SetText vcmrsnotrobats
        Shell "notepad.exe", vbNormalFocus
    'Send the keys CTRL+V To Notepad (i.e the window that has focus)
        SendKeys "^V"
        wait 2
        Clipboard.Clear
        If MsgBox("Vols passar tots els CMRs que no trobo escanejats com a pendents d'escanejar?", vbExclamation + vbDefaultButton2 + vbYesNo, "PDF dels CMRs no trobats") = vbYes Then dbvendes.Execute "update Transportistes_avisos set escanejat=false where numeroavis in (" + vcmrsnotrobats + ")"
      Else: MsgBox "Tots els CMRs estan escanejats.", vbInformation, "OK"
   End If
   Set rst = Nothing
End Sub

Private Sub menurma_Click()
   formrecepciorma.Show 1
End Sub

Private Sub menviatsasap_Click()
   datacapcalera.RecordSource = "select * from capcaleraalbara where dataenvioasap<>null  order by numalbara desc"
   datacapcalera.Refresh
   etfiltreactivat = "Filtre: Passat a SAP"
End Sub

Private Sub mescanejaralbaransproveidor_Click()

   Unload formescanejaralbaransproveidor
   Load formescanejaralbaransproveidor
   formescanejaralbaransproveidor.ettipusescaneig.caption = "Esperant Albarans del Proveïdor..."
   formescanejaralbaransproveidor.tag = "albarans"
   formescanejaralbaransproveidor.caption = formescanejaralbaransproveidor.ettipusescaneig.caption
   DoEvents
   formescanejaralbaransproveidor.Show 1
End Sub

Private Sub mfacturarclixes_Click()
   Dim vtreball As Double
   Dim vordre As Byte
   vtreball = cadbl(InputBox("Entra el treball que vols albaranar.", "Albaranar clixes"))
   If vtreball > 0 Then
    vordre = cadbl(InputBox("Entra la versió que vols comprovar", "Albaranar clixes"))
    If vordre = 0 Then vtreball = 0
   End If
   If vtreball > 0 And vordre > 0 Then mirarsiclixesperfacturar True, vtreball, vordre
End Sub

Private Sub modificar_Click()
  If datalinies.Recordset.EOF Then Exit Sub
  activar_frames True
  actualitzar_bobinesent_vendes cadbl(datalinies.Recordset!lotinplacsa)
  carregar_bobinesentrega cadbl(datalinies.Recordset!lotinplacsa), cadbl(datalinies.Recordset!numalbara)
  datalinies.Recordset.Edit
  If atrim(datacapcalera.Recordset!codiclient) <> "43000006841" Then
      llistabobinessel.Enabled = True
        Else: llistabobinessel.Enabled = False
  End If
End Sub

Private Sub mpendentsdesap_Click()
   datacapcalera.RecordSource = "select * from capcaleraalbara where dataenvioasap=null  order by numalbara desc"
   datacapcalera.Refresh
   etfiltreactivat = "Filtre: Pendent -> SAP"
End Sub

Private Sub mpendentsenviats_Click()
   ratoli "espera"
   etfiltreactivat = ""
   DoEvents
   datacapcalera.RecordSource = " select * from capcaleraalbara where numalbara in (SELECT  bobinesent.numalbara From bobinesent WHERE (((bobinesent.dataentrega) Is Null) AND ((bobinesent.numalbara)>0)))"
   datacapcalera.Refresh
   
   If Not datacapcalera.Recordset.EOF Then
      datacapcalera.Recordset.MoveLast: datacapcalera.Recordset.MoveFirst
      etfiltreactivat = "Filtre: Pendent ENVIAR"
       Else: datacapcalera.RecordSource = "select * from capcaleraalbara order by numalbara desc"
   End If
   ratoli "normal"
End Sub

Private Sub mpersonalitzarbases_Click()
   imprimirbases 1, True
End Sub

Private Sub mprogramadembolicarpalets_Click()
   Shell """" + App.Path + "\Baixes enflajar.exe" + """", vbNormalFocus
End Sub

Private Sub mpwdsmpt_Click()
Dim usr As String
   usr = InputBoxEx("Entra la contrasenya d'enviament del correu:" + Chr(10) + llegir_ini("Enviomails", "usuari", "comandes.ini") + Chr(10) + "(Respecteu majúscules i minúscules)", "Contrasenya", , , , , , SPassword)
   If usr <> "" Then
      escriure_ini "Enviomails", "contrasenya", usr, "comandes.ini"
      MsgBox "Contrasenya canviada.", vbInformation, "D'acord"
   End If
End Sub

Private Sub mtotsCmrNoescanejats_Click()
   CMRsNoEscanejats "Tots"
End Sub

Private Sub mtotsdeunclient_Click()
  Dim id_direnvio As Long
  id_direnvio = triar_direnvio_client_busqueda
  If id_direnvio = 0 Then Exit Sub
  datacapcalera.RecordSource = "select * from capcaleraalbara where id_direnvio=" + atrim(id_direnvio) + "  order by numalbara desc"
  datacapcalera.Refresh
  If datacapcalera.Recordset.EOF Then
      MsgBox "No hi ha cap albarà amb aquest destí", vbCritical, "Error"
      datacapcalera.RecordSource = "select * from capcaleraalbara order by numalbara desc"
      datacapcalera.Refresh
      Exit Sub
  End If
  etfiltreactivat = "Filtre: Client escullit"
End Sub

Private Sub mtransportistes_Click()
   formtransportistes.Show 1
End Sub

Private Sub musrsmtp_Click()
 Dim usr As String
   usr = InputBox("Entra l'usuari d'enviament del correu:" + Chr(10) + "Ex: usuari@inplacsa.com", "Usuari", llegir_ini("Enviomails", "usuari", "comandes.ini"))
   If usr <> "" Then
      escriure_ini "Enviomails", "usuari", usr, "comandes.ini"
      MsgBox "Usuari canviat.", vbInformation, "D'acord"
   End If
   
End Sub

Private Sub mveurebases_Click()
etmetrescubicscalculats_DblClick
End Sub

Private Sub reixalinies_DblClick()
   'actualitzar_totals_a_comandes reixalinies.Text
End Sub

Private Sub reixalinies_GotFocus()
   actualitzartotalsdelinia
End Sub

Private Sub reixapeu_OnAddNew()
  Dim rst As Recordset
  Dim ultim As Double
  If dataliniespeu.Recordset.RecordCount > 5 Then MsgBox "Nomes es poden possar 6 peus d'albarà com a màxim", vbCritical, "Error": dataliniespeu.Recordset.CancelUpdate: Exit Sub
  Set rst = datacapcalera.Database.OpenRecordset("select ordre from liniespeu where numalbara=" + atrim(datacapcalera.Recordset!numalbara) + " order by ordre DESC")
  reixapeu.col = 0
  If Not rst.EOF Then
    ultim = cadbl(rst!ordre)
  End If
  dataliniespeu.Recordset!numalbara = datacapcalera.Recordset!numalbara
  dataliniespeu.Recordset!ordre = ultim + 5
  reixapeu.Text = ultim + 5
  reixapeu.col = 1
  Set rst = Nothing
End Sub

Private Sub seltot_Click()
  Dim i As Integer
  Dim nhihauna As Boolean
  Dim vmarcarlestotes As Boolean
  If seltot.tag = "totes" Then vmarcarlestotes = True
  seltot.tag = "1"
  If Not vmarcarlestotes Then
        For i = 0 To llistabobinessel.ListCount - 1
           If llistabobinessel.Selected(i) Then nhihauna = True
        Next i
        Else: nhihauna = False
  End If
  For i = 0 To llistabobinessel.ListCount - 1
     llistabobinessel.Selected(i) = IIf(nhihauna, False, True)
  Next i
  sumarkilosimetresseleccionats
  actualitzartotalsdelinia
  seltot.tag = "0"
End Sub


Sub imprimirpackinglist(numalb As String)
    Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "packinglistexpedicions.rpt", 1)
  oreport.RecordSelectionFormula = "{capcaleraalbara.numalbara} = " + numalb + Chr(13) + " and {liniesalbara.numalbara} = " + numalb + Chr(13) + " and {bobinesent.numalbara}=" + numalb
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "vendes.mdb"
  oreport.Database.Tables.Item(2).Location = rutadelfitxer(cami) + "vendes.mdb"
  oreport.Database.Tables.Item(3).Location = rutadelfitxer(cami) + "baixes.mdb"
  oreport.Database.Tables.Item(4).Location = rutadelfitxer(cami) + "comandes.mdb"
  oreport.DiscardSavedData
  oreport.VerifyOnEveryPrint = True
  
  
  
  'oreport.FormulaFields.GetItemByName("detallbobines").Text = "'N'"
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
    Else
      oreport.PrintOut False, 1
  End If
  
End Sub

Private Sub Text32_KeyDown(KeyCode As Integer, Shift As Integer)
  Text32.tag = "1"
End Sub

Private Sub Timer1_Timer()
   If datalinies.Recordset.EditMode > 0 Then
       Command1.BackColor = QBColor(12)
         Else: Command1.BackColor = &H8000000F
   End If
   If datacapcalera.Recordset.EditMode > 0 Then
       Command6.BackColor = QBColor(12)
         Else: Command6.BackColor = &H8000000F
   End If
End Sub

