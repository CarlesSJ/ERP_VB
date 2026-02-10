VERSION 5.00
Object = "{8C45F041-B87C-11D1-96EF-845C0FC10100}#1.3#0"; "SCROLLBOX.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form formclients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manteniment de Clients"
   ClientHeight    =   7215
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11280
   ControlBox      =   0   'False
   Icon            =   "clients.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame comandesafectades 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Comandes Afectades"
      Height          =   5340
      Left            =   7755
      TabIndex        =   76
      Top             =   7365
      Visible         =   0   'False
      Width           =   7065
      Begin MSFlexGridLib.MSFlexGrid reixa 
         Height          =   4965
         Left            =   150
         TabIndex        =   77
         Top             =   225
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   8758
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
      End
   End
   Begin VB.Data comandes 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9225
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select comanda,texteimpressio,'            ' as RISC, puntrisc from comandes "
      Top             =   810
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   11130
      Begin VB.CommandButton Command9 
         Height          =   450
         Left            =   9495
         Picture         =   "clients.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   225
         Width           =   525
      End
      Begin VB.CommandButton bmodificacions 
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   2
         Left            =   1935
         Picture         =   "clients.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   175
         TabStop         =   0   'False
         ToolTipText     =   "Llista de canvis realitzats a la comanda."
         Top             =   150
         Width           =   315
      End
      Begin Crystal.CrystalReport report 
         Left            =   3165
         Top             =   270
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Comandes Afectades"
         Height          =   465
         Left            =   7800
         TabIndex        =   75
         Top             =   225
         Width           =   915
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   525
         Picture         =   "clients.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Modificació Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton gravar 
         Height          =   450
         Left            =   10035
         Picture         =   "clients.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Guardar Registres"
         Top             =   225
         Width           =   450
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   1425
         Picture         =   "clients.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "clients.frx":213C
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.Data clients 
         Caption         =   "                     Clients"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   390
         Left            =   3990
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "clients"
         Top             =   225
         Width           =   3765
      End
      Begin VB.CommandButton sortir 
         Height          =   450
         Left            =   10575
         Picture         =   "clients.frx":26C6
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "Sortir"
         Top             =   225
         Width           =   450
      End
      Begin VB.CommandButton consultar 
         Height          =   360
         Left            =   975
         Picture         =   "clients.frx":2C50
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   "Busqueda de Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.Label estattaula 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3810
         TabIndex        =   60
         Top             =   300
         Width           =   105
      End
   End
   Begin VB.Frame areadatos 
      Enabled         =   0   'False
      Height          =   6555
      Left            =   75
      TabIndex        =   2
      Top             =   675
      Width           =   11160
      Begin ScrollBoxCtl.ScrollBox scroll 
         Height          =   5595
         Left            =   360
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   570
         Width           =   10830
         _ExtentX        =   19103
         _ExtentY        =   9869
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Frame Frame2 
            Caption         =   "Enviaments"
            Height          =   10800
            Left            =   0
            TabIndex        =   82
            Top             =   -15
            Width           =   10425
            Begin VB.Frame Frame8 
               Caption         =   "Pel CMR o Carta de ports"
               Height          =   1935
               Left            =   4455
               TabIndex        =   204
               Top             =   8460
               Width           =   5865
               Begin VB.CheckBox Check33 
                  Caption         =   "Incloure el Nº de comanda del client a les instruccions del remitent."
                  DataField       =   "cmr_comandaclient"
                  DataSource      =   "envios"
                  Height          =   225
                  Left            =   135
                  TabIndex        =   207
                  Top             =   285
                  Width           =   5625
               End
               Begin VB.TextBox Text46 
                  DataField       =   "cmr_observacions"
                  DataSource      =   "envios"
                  Height          =   1020
                  Left            =   90
                  MaxLength       =   255
                  MultiLine       =   -1  'True
                  TabIndex        =   205
                  Top             =   810
                  Width           =   5685
               End
               Begin VB.Label Label28 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Instruccions del remitent:"
                  Height          =   225
                  Left            =   180
                  TabIndex        =   206
                  Top             =   585
                  Width           =   5190
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "Passar avís abans d'acabar al producció"
               Height          =   1275
               Left            =   4485
               TabIndex        =   183
               Top             =   5475
               Width           =   5880
               Begin VB.TextBox Text12 
                  DataField       =   "avisfiproduccio"
                  DataSource      =   "envios"
                  Height          =   855
                  Left            =   120
                  MaxLength       =   255
                  TabIndex        =   184
                  Top             =   300
                  Width           =   5655
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "Direccions d'enviament"
               Height          =   3795
               Left            =   60
               TabIndex        =   98
               Top             =   540
               Width           =   4380
               Begin VB.TextBox Text13 
                  DataField       =   "observacionscomandaalalbara"
                  DataSource      =   "envios"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   855
                  MaxLength       =   100
                  TabIndex        =   185
                  Top             =   3435
                  Width           =   3405
               End
               Begin VB.ComboBox combopais 
                  DataField       =   "pais"
                  DataSource      =   "envios"
                  Height          =   315
                  ItemData        =   "clients.frx":31DA
                  Left            =   855
                  List            =   "clients.frx":31EA
                  TabIndex        =   166
                  Top             =   1500
                  WhatsThisHelpID =   1
                  Width           =   660
               End
               Begin VB.ComboBox comboquifactura 
                  DataField       =   "empresa"
                  DataSource      =   "envios"
                  Height          =   315
                  ItemData        =   "clients.frx":31FE
                  Left            =   1065
                  List            =   "clients.frx":3208
                  TabIndex        =   161
                  ToolTipText     =   "Qui facturarà a aquest client? Predeterminat Inplacsa."
                  Top             =   2865
                  Width           =   2475
               End
               Begin VB.TextBox Text41 
                  DataField       =   "arxiuexp"
                  DataSource      =   "envios"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   675
                  Locked          =   -1  'True
                  MouseIcon       =   "clients.frx":321E
                  MousePointer    =   99  'Custom
                  TabIndex        =   158
                  Top             =   2505
                  Width           =   1215
               End
               Begin VB.TextBox Text42 
                  DataField       =   "arxiuult"
                  DataSource      =   "envios"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   2610
                  Locked          =   -1  'True
                  MouseIcon       =   "clients.frx":3370
                  MousePointer    =   99  'Custom
                  TabIndex        =   157
                  Top             =   2505
                  Width           =   1365
               End
               Begin VB.CommandButton Command1 
                  Height          =   315
                  Left            =   1950
                  Picture         =   "clients.frx":34C2
                  Style           =   1  'Graphical
                  TabIndex        =   156
                  Top             =   2505
                  Width           =   315
               End
               Begin VB.CommandButton Command2 
                  Height          =   315
                  Left            =   4035
                  Picture         =   "clients.frx":3888
                  Style           =   1  'Graphical
                  TabIndex        =   155
                  Top             =   2505
                  Width           =   315
               End
               Begin VB.TextBox Text8 
                  DataField       =   "faxe"
                  DataSource      =   "envios"
                  Height          =   285
                  Left            =   2145
                  TabIndex        =   108
                  Top             =   1845
                  WhatsThisHelpID =   1
                  Width           =   1605
               End
               Begin VB.TextBox Text9 
                  DataField       =   "telefone"
                  DataSource      =   "envios"
                  Height          =   285
                  Left            =   840
                  TabIndex        =   107
                  Top             =   1845
                  WhatsThisHelpID =   1
                  Width           =   1320
               End
               Begin VB.TextBox Text33 
                  DataField       =   "provinciae"
                  DataSource      =   "envios"
                  Height          =   285
                  Left            =   2175
                  TabIndex        =   106
                  Top             =   1215
                  WhatsThisHelpID =   1
                  Width           =   2145
               End
               Begin VB.TextBox Text35 
                  DataField       =   "codipostale"
                  DataSource      =   "envios"
                  Height          =   285
                  Left            =   840
                  TabIndex        =   105
                  Top             =   1215
                  WhatsThisHelpID =   1
                  Width           =   1320
               End
               Begin VB.TextBox Text39 
                  DataField       =   "domicilie"
                  DataSource      =   "envios"
                  Height          =   285
                  Left            =   840
                  TabIndex        =   104
                  Top             =   645
                  WhatsThisHelpID =   1
                  Width           =   3495
               End
               Begin VB.TextBox nomenvio 
                  DataField       =   "nome"
                  DataSource      =   "envios"
                  Height          =   285
                  Left            =   840
                  TabIndex        =   103
                  Top             =   360
                  WhatsThisHelpID =   1
                  Width           =   3495
               End
               Begin VB.ComboBox Combo_peuimprenta 
                  Height          =   315
                  Left            =   1260
                  TabIndex        =   102
                  Top             =   2160
                  WhatsThisHelpID =   1
                  Width           =   2550
               End
               Begin VB.TextBox Text37 
                  DataField       =   "poblacioe"
                  DataSource      =   "envios"
                  Height          =   285
                  Left            =   840
                  TabIndex        =   101
                  Top             =   915
                  WhatsThisHelpID =   1
                  Width           =   3495
               End
               Begin VB.ComboBox Combo1 
                  DataField       =   "idioma"
                  DataSource      =   "envios"
                  Height          =   315
                  ItemData        =   "clients.frx":3C4E
                  Left            =   2205
                  List            =   "clients.frx":3C61
                  TabIndex        =   100
                  Text            =   "Combo1"
                  Top             =   1515
                  WhatsThisHelpID =   1
                  Width           =   675
               End
               Begin VB.CheckBox Check24 
                  Caption         =   "Primera etiqueta verificada."
                  DataField       =   "verificatclientnouxretiquetes"
                  DataSource      =   "envios"
                  Height          =   225
                  Left            =   2040
                  TabIndex        =   99
                  Top             =   120
                  Width           =   2265
               End
               Begin VB.Label etnumenviament 
                  BackStyle       =   0  'Transparent
                  Caption         =   "(0000)"
                  DataField       =   "id"
                  DataSource      =   "envios"
                  ForeColor       =   &H00FF0000&
                  Height          =   330
                  Left            =   90
                  TabIndex        =   213
                  Top             =   195
                  Width           =   870
               End
               Begin VB.Label Label1 
                  Caption         =   "Obs. Comanda a l'albarà:"
                  DataSource      =   "clients"
                  Height          =   585
                  Index           =   28
                  Left            =   60
                  TabIndex        =   186
                  Top             =   3150
                  Width           =   720
               End
               Begin VB.Label Label9 
                  Caption         =   "País"
                  Height          =   240
                  Left            =   45
                  TabIndex        =   167
                  Top             =   1530
                  Width           =   510
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Qui factura?"
                  Height          =   345
                  Index           =   20
                  Left            =   90
                  TabIndex        =   162
                  ToolTipText     =   "Qui facturarà a aquest client? Predeterminat Inplacsa."
                  Top             =   2925
                  Width           =   1155
               End
               Begin VB.Label Label1 
                  Caption         =   "ArxExp:"
                  DataSource      =   "clients"
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   160
                  Top             =   2505
                  Width           =   720
               End
               Begin VB.Label Label1 
                  Caption         =   " Ult:"
                  DataSource      =   "clients"
                  Height          =   255
                  Index           =   26
                  Left            =   2265
                  TabIndex        =   159
                  Top             =   2505
                  Width           =   510
               End
               Begin VB.Label Label4 
                  Caption         =   "Nom"
                  Height          =   270
                  Index           =   0
                  Left            =   30
                  TabIndex        =   115
                  Top             =   375
                  Width           =   735
               End
               Begin VB.Label Label4 
                  Caption         =   "Peu Imp i Data"
                  Height          =   270
                  Index           =   1
                  Left            =   75
                  TabIndex        =   114
                  Top             =   2190
                  Width           =   1185
               End
               Begin VB.Label Label4 
                  Caption         =   "Tlf/Fax:"
                  Height          =   270
                  Index           =   2
                  Left            =   30
                  TabIndex        =   113
                  Top             =   1875
                  Width           =   735
               End
               Begin VB.Label Label4 
                  Caption         =   "CP. Provincia:"
                  Height          =   270
                  Index           =   3
                  Left            =   30
                  TabIndex        =   112
                  Top             =   1245
                  Width           =   1095
               End
               Begin VB.Label Label4 
                  Caption         =   "Població:"
                  Height          =   270
                  Index           =   4
                  Left            =   30
                  TabIndex        =   111
                  Top             =   945
                  Width           =   735
               End
               Begin VB.Label Label4 
                  Caption         =   "Domicili:"
                  Height          =   270
                  Index           =   5
                  Left            =   30
                  TabIndex        =   110
                  Top             =   660
                  Width           =   735
               End
               Begin VB.Label Idioma 
                  Caption         =   "Idioma"
                  Height          =   240
                  Left            =   1635
                  TabIndex        =   109
                  Top             =   1560
                  Width           =   510
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Per la Rebobinadora "
               Height          =   1485
               Left            =   4470
               TabIndex        =   153
               Top             =   6960
               Width           =   5865
               Begin VB.TextBox cavisrebobinadora 
                  DataField       =   "avisrebobinadora"
                  DataSource      =   "envios"
                  Height          =   300
                  Left            =   90
                  MaxLength       =   80
                  TabIndex        =   202
                  Top             =   510
                  Width           =   5685
               End
               Begin VB.Label Label27 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Avís per a la secció de Rebobinadora (Al començar la baixa)"
                  Height          =   225
                  Left            =   90
                  TabIndex        =   203
                  Top             =   285
                  Width           =   5190
               End
            End
            Begin VB.Frame palets 
               Caption         =   "Indicar al Palet"
               Height          =   4125
               Left            =   4470
               TabIndex        =   119
               Top             =   105
               Width           =   5925
               Begin VB.CheckBox Check32 
                  Caption         =   "Imprimir Ref. Inplacsa"
                  DataField       =   "paletreferenciainplacsa"
                  DataSource      =   "envios"
                  Height          =   330
                  Left            =   4005
                  TabIndex        =   201
                  Top             =   3765
                  Width           =   1905
               End
               Begin VB.ComboBox combo_conosprotectors 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   132
                  Top             =   3495
                  Width           =   4500
               End
               Begin VB.ComboBox combo_guardarmostres 
                  Height          =   315
                  Left            =   1305
                  TabIndex        =   131
                  Top             =   2775
                  Width           =   4500
               End
               Begin VB.ComboBox combo_cert_qualitat 
                  Height          =   315
                  Left            =   1305
                  TabIndex        =   130
                  Top             =   2460
                  Width           =   4500
               End
               Begin VB.ComboBox combo_emb_anonim 
                  Height          =   315
                  Left            =   1305
                  TabIndex        =   129
                  Top             =   2130
                  Width           =   4500
               End
               Begin VB.ComboBox combo_protecciospr 
                  Height          =   315
                  ItemData        =   "clients.frx":3C79
                  Left            =   1305
                  List            =   "clients.frx":3C7B
                  TabIndex        =   128
                  Top             =   1815
                  Width           =   4500
               End
               Begin VB.ComboBox combo_protecciop 
                  Height          =   315
                  ItemData        =   "clients.frx":3C7D
                  Left            =   1305
                  List            =   "clients.frx":3C7F
                  TabIndex        =   127
                  Top             =   1500
                  Width           =   4500
               End
               Begin VB.ComboBox combo_protecciob 
                  Height          =   315
                  ItemData        =   "clients.frx":3C81
                  Left            =   1305
                  List            =   "clients.frx":3C83
                  TabIndex        =   126
                  Top             =   1185
                  Width           =   4500
               End
               Begin VB.ComboBox combo_alcadapalet 
                  Height          =   315
                  ItemData        =   "clients.frx":3C85
                  Left            =   1305
                  List            =   "clients.frx":3C87
                  TabIndex        =   125
                  Top             =   840
                  Width           =   1170
               End
               Begin VB.ComboBox combo_tipuspalet 
                  Height          =   315
                  ItemData        =   "clients.frx":3C89
                  Left            =   1305
                  List            =   "clients.frx":3C8B
                  TabIndex        =   124
                  Top             =   525
                  Width           =   4500
               End
               Begin VB.CheckBox Check20 
                  Caption         =   "Adjuntar Albarà al palet"
                  DataField       =   "albaraalpalet"
                  DataSource      =   "envios"
                  Height          =   300
                  Left            =   180
                  TabIndex        =   123
                  Top             =   3765
                  Width           =   1965
               End
               Begin VB.CheckBox Check21 
                  Caption         =   "Adjuntar Packing List"
                  DataField       =   "packingalpalet"
                  DataSource      =   "envios"
                  Height          =   330
                  Left            =   2145
                  TabIndex        =   122
                  Top             =   3750
                  Width           =   1905
               End
               Begin VB.TextBox Text10 
                  DataField       =   "pesmaxpalet"
                  DataSource      =   "envios"
                  Height          =   300
                  Left            =   2955
                  MaxLength       =   10
                  TabIndex        =   121
                  Top             =   840
                  Width           =   1065
               End
               Begin VB.ComboBox combo_guardarmostressol 
                  Height          =   315
                  Left            =   1305
                  TabIndex        =   120
                  Top             =   3105
                  Width           =   4500
               End
               Begin VB.Label missatgeenviament 
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
                  Left            =   180
                  TabIndex        =   144
                  Top             =   195
                  Width           =   4305
               End
               Begin VB.Label Label18 
                  Caption         =   "Cono protector:"
                  Height          =   435
                  Left            =   165
                  TabIndex        =   143
                  Top             =   3540
                  Width           =   1380
               End
               Begin VB.Label Label17 
                  Caption         =   "G. mostra Reb:"
                  Height          =   345
                  Left            =   165
                  TabIndex        =   142
                  Top             =   2820
                  Width           =   1380
               End
               Begin VB.Label Label16 
                  Caption         =   "Cert Qualitat:"
                  Height          =   330
                  Left            =   165
                  TabIndex        =   141
                  Top             =   2505
                  Width           =   1380
               End
               Begin VB.Label Label15 
                  Caption         =   "Emb Anonim:"
                  Height          =   300
                  Left            =   165
                  TabIndex        =   140
                  Top             =   2175
                  Width           =   1380
               End
               Begin VB.Label Label14 
                  Caption         =   "Protec Superior:"
                  Height          =   285
                  Left            =   165
                  TabIndex        =   139
                  Top             =   1860
                  Width           =   1380
               End
               Begin VB.Label Label13 
                  Caption         =   "Protec Pisos:"
                  Height          =   315
                  Left            =   165
                  TabIndex        =   138
                  Top             =   1515
                  Width           =   1275
               End
               Begin VB.Label Label12 
                  Caption         =   "Protec Base:"
                  Height          =   315
                  Left            =   165
                  TabIndex        =   137
                  Top             =   1200
                  Width           =   1305
               End
               Begin VB.Label Label11 
                  Caption         =   "Alçada Palet:"
                  Height          =   315
                  Left            =   180
                  TabIndex        =   136
                  Top             =   855
                  Width           =   1035
               End
               Begin VB.Label Label10 
                  Caption         =   "Tipus Palet:"
                  Height          =   315
                  Left            =   180
                  TabIndex        =   135
                  Top             =   540
                  Width           =   1035
               End
               Begin VB.Label Label5 
                  Caption         =   "Pes Max:"
                  Height          =   480
                  Left            =   2535
                  TabIndex        =   134
                  Top             =   810
                  Width           =   630
               End
               Begin VB.Label Label6 
                  Caption         =   "G. mostra Sol:"
                  Height          =   345
                  Left            =   165
                  TabIndex        =   133
                  Top             =   3150
                  Width           =   1380
               End
            End
            Begin VB.CommandButton Command4 
               Height          =   315
               Left            =   3090
               Picture         =   "clients.frx":3C8D
               Style           =   1  'Graphical
               TabIndex        =   118
               TabStop         =   0   'False
               ToolTipText     =   "Eliminacio Envio (Eliminar aquest envio pot supusar la perdua de traçabilitat amb la comanda)"
               Top             =   195
               Width           =   390
            End
            Begin VB.CommandButton Command6 
               Height          =   315
               Left            =   2700
               Picture         =   "clients.frx":4217
               Style           =   1  'Graphical
               TabIndex        =   117
               ToolTipText     =   "Actualitzar o guardar enviament"
               Top             =   195
               Width           =   390
            End
            Begin VB.CommandButton Command5 
               Height          =   315
               Left            =   2325
               Picture         =   "clients.frx":47A1
               Style           =   1  'Graphical
               TabIndex        =   116
               TabStop         =   0   'False
               ToolTipText     =   "Alta  Envio"
               Top             =   195
               Width           =   390
            End
            Begin VB.Data envios 
               Caption         =   "Envios"
               Connect         =   "Access"
               DatabaseName    =   ""
               DefaultCursorType=   0  'DefaultCursor
               DefaultType     =   2  'UseODBC
               Exclusive       =   0   'False
               Height          =   345
               Left            =   75
               Options         =   0
               ReadOnly        =   0   'False
               RecordsetType   =   1  'Dynaset
               RecordSource    =   ""
               Top             =   180
               Width           =   2040
            End
            Begin VB.Frame Frame4 
               Caption         =   "Indicar a L'Albarà"
               Height          =   6315
               Left            =   60
               TabIndex        =   85
               Top             =   4440
               Width           =   4380
               Begin VB.TextBox Text47 
                  DataField       =   "avisalbaragenerat"
                  DataSource      =   "envios"
                  Height          =   285
                  Left            =   1125
                  MaxLength       =   60
                  TabIndex        =   220
                  Top             =   5520
                  Width           =   2970
               End
               Begin VB.ComboBox comboincoterms 
                  DataField       =   "incoterm"
                  DataSource      =   "envios"
                  Height          =   315
                  ItemData        =   "clients.frx":4D2B
                  Left            =   1140
                  List            =   "clients.frx":4D2D
                  Locked          =   -1  'True
                  TabIndex        =   219
                  Top             =   5175
                  Width           =   2970
               End
               Begin VB.CheckBox Check36 
                  Caption         =   "La quantitat entregada ha de ser igual a la demanada."
                  DataField       =   "entregaigualademanada"
                  DataSource      =   "envios"
                  Height          =   225
                  Left            =   135
                  TabIndex        =   217
                  Top             =   6060
                  Width           =   4200
               End
               Begin VB.CheckBox Check35 
                  Caption         =   "Impost Llei 7/2022 inclòs en el PVP"
                  DataField       =   "impostinclosalPVP"
                  DataSource      =   "envios"
                  Height          =   225
                  Left            =   135
                  TabIndex        =   216
                  Top             =   5820
                  Width           =   3885
               End
               Begin VB.ComboBox combotransportfavorit 
                  DataField       =   "nom_transportFAVORIT"
                  DataSource      =   "envios"
                  Height          =   315
                  Left            =   1335
                  TabIndex        =   214
                  Top             =   4845
                  Width           =   2970
               End
               Begin VB.CheckBox Check31 
                  Caption         =   "Possar volum del bulto a l'albarà."
                  DataField       =   "volumalalbara"
                  DataSource      =   "envios"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   200
                  Top             =   3375
                  Width           =   3405
               End
               Begin VB.ComboBox combotipus_sscc 
                  DataField       =   "tipusscc"
                  DataSource      =   "envios"
                  Height          =   315
                  ItemData        =   "clients.frx":4D2F
                  Left            =   1005
                  List            =   "clients.frx":4D39
                  TabIndex        =   198
                  Top             =   3870
                  Width           =   1605
               End
               Begin VB.CheckBox Check30 
                  Caption         =   "SSCC a l'albarà. (Implica tipus SSCC [albarà]"
                  DataField       =   "sccalacomanda"
                  DataSource      =   "envios"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   197
                  Top             =   3615
                  Width           =   3885
               End
               Begin VB.CheckBox Check29 
                  Caption         =   "Possar informació caducitat del material."
                  DataField       =   "albaracaducitatmaterial"
                  DataSource      =   "envios"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   181
                  Top             =   3165
                  Width           =   3405
               End
               Begin VB.CheckBox Check26 
                  Caption         =   "Totalitzar bobines metres i unitats a l'albarà"
                  DataField       =   "albaratotaldetallbobines"
                  DataSource      =   "envios"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   177
                  Top             =   2925
                  Width           =   3405
               End
               Begin VB.CommandButton Command12 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Linies peu albarà"
                  Height          =   420
                  Left            =   2880
                  Style           =   1  'Graphical
                  TabIndex        =   174
                  Top             =   165
                  Width           =   1425
               End
               Begin VB.TextBox Text11 
                  DataField       =   "observacionsalbara"
                  DataSource      =   "envios"
                  Height          =   285
                  Left            =   1020
                  MaxLength       =   60
                  TabIndex        =   172
                  Top             =   4545
                  Width           =   3330
               End
               Begin VB.Frame framepesnet 
                  Height          =   1785
                  Left            =   2085
                  TabIndex        =   87
                  Top             =   825
                  Visible         =   0   'False
                  Width           =   2250
                  Begin MSDBGrid.DBGrid reixapesnet 
                     Bindings        =   "clients.frx":4D4C
                     Height          =   1290
                     Left            =   30
                     OleObjectBlob   =   "clients.frx":4D61
                     TabIndex        =   89
                     Top             =   120
                     Width           =   2175
                  End
                  Begin VB.CommandButton Command8 
                     Caption         =   "*"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   60
                     TabIndex        =   88
                     ToolTipText     =   "Possar tots els pesos iguals"
                     Top             =   1425
                     Width           =   285
                  End
               End
               Begin VB.CheckBox Check17 
                  Caption         =   "Possar Data Producció"
                  DataField       =   "albaraambdataproduccio"
                  DataSource      =   "envios"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   171
                  Top             =   2400
                  Width           =   2220
               End
               Begin VB.CheckBox Check17 
                  Caption         =   "Sense texte d'impresió"
                  DataField       =   "albarasensetexteimpresio"
                  DataSource      =   "envios"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   170
                  Top             =   2160
                  Width           =   1875
               End
               Begin VB.CheckBox Check25 
                  Caption         =   "Agrupar Comandes client al albaranar"
                  DataField       =   "agruparcomandesdeclient"
                  DataSource      =   "envios"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   165
                  Top             =   2685
                  Width           =   3405
               End
               Begin VB.ComboBox combopackinglistalbara 
                  DataField       =   "packinglistalbara"
                  DataSource      =   "envios"
                  Height          =   315
                  ItemData        =   "clients.frx":5744
                  Left            =   990
                  List            =   "clients.frx":5751
                  TabIndex        =   163
                  Top             =   4200
                  Width           =   1995
               End
               Begin VB.CheckBox peces 
                  Caption         =   "Arrodonir Kg a l'albarà"
                  DataField       =   "albaraarrodonirkg"
                  DataSource      =   "envios"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   97
                  Top             =   210
                  Width           =   1980
               End
               Begin VB.CheckBox Check11 
                  Caption         =   "Codi de Barres"
                  DataField       =   "codibarres"
                  DataSource      =   "envios"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   96
                  Top             =   405
                  Width           =   1920
               End
               Begin VB.CheckBox Check12 
                  Caption         =   "Data de Fabricació"
                  DataField       =   "datafabricacio"
                  DataSource      =   "envios"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   95
                  Top             =   645
                  Width           =   3615
               End
               Begin VB.CheckBox Check13 
                  Caption         =   "Pes Net "
                  DataField       =   "pesnetbrut"
                  DataSource      =   "envios"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   94
                  Top             =   1155
                  Width           =   960
               End
               Begin VB.CheckBox Check14 
                  Caption         =   "Detall Bob Palet Albarà"
                  DataField       =   "detallbobalpalet"
                  DataSource      =   "envios"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   93
                  Top             =   1425
                  Width           =   3615
               End
               Begin VB.CheckBox Check15 
                  Caption         =   "Detall Bob Palet PF"
                  DataField       =   "detallbobalfrontal"
                  DataSource      =   "envios"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   92
                  Top             =   1665
                  Width           =   1800
               End
               Begin VB.CheckBox Check16 
                  Caption         =   "Albarà Valorat"
                  DataField       =   "albaravalorat"
                  DataSource      =   "envios"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   91
                  Top             =   900
                  Width           =   3615
               End
               Begin VB.CheckBox Check17 
                  Caption         =   "Demanar conf. envio"
                  DataField       =   "okenvio"
                  DataSource      =   "envios"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   90
                  Top             =   1920
                  Width           =   1875
               End
               Begin VB.Data datapesnet 
                  Caption         =   "datapesnet"
                  Connect         =   "Access"
                  DatabaseName    =   ""
                  DefaultCursorType=   0  'DefaultCursor
                  DefaultType     =   2  'UseODBC
                  Exclusive       =   0   'False
                  Height          =   345
                  Left            =   735
                  Options         =   0
                  ReadOnly        =   0   'False
                  RecordsetType   =   1  'Dynaset
                  RecordSource    =   "taulapesnet"
                  Top             =   2400
                  Visible         =   0   'False
                  Width           =   1140
               End
               Begin VB.CheckBox pesnetstd 
                  Caption         =   "Std ----->"
                  DataField       =   "pesnetstd"
                  DataSource      =   "envios"
                  ForeColor       =   &H00FF0000&
                  Height          =   240
                  Left            =   1140
                  TabIndex        =   86
                  Top             =   1215
                  Visible         =   0   'False
                  Width           =   930
               End
               Begin VB.Label Label31 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Email avís albarà FET:"
                  Height          =   405
                  Left            =   135
                  TabIndex        =   221
                  Top             =   5430
                  Width           =   900
               End
               Begin VB.Label Label30 
                  BackStyle       =   0  'Transparent
                  Caption         =   "INCOTERMS:"
                  Height          =   225
                  Left            =   60
                  TabIndex        =   218
                  Top             =   5250
                  Width           =   1065
               End
               Begin VB.Label Label29 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Transport Favorit:"
                  Height          =   225
                  Left            =   60
                  TabIndex        =   215
                  Top             =   4920
                  Width           =   1410
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipus SSCC:"
                  Height          =   345
                  Index           =   29
                  Left            =   75
                  TabIndex        =   199
                  ToolTipText     =   "Qui facturarà a aquest client? Predeterminat Inplacsa."
                  Top             =   3915
                  Width           =   1155
               End
               Begin VB.Label Label19 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Obs. Albarà:"
                  Height          =   225
                  Left            =   75
                  TabIndex        =   173
                  Top             =   4590
                  Width           =   900
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Packing-List"
                  Height          =   345
                  Index           =   22
                  Left            =   60
                  TabIndex        =   164
                  ToolTipText     =   "Qui facturarà a aquest client? Predeterminat Inplacsa."
                  Top             =   4140
                  Width           =   1155
               End
            End
            Begin VB.ComboBox combo_paperfrontal 
               Height          =   315
               ItemData        =   "clients.frx":5780
               Left            =   5145
               List            =   "clients.frx":5782
               TabIndex        =   84
               Top             =   4650
               Width           =   3285
            End
            Begin VB.CommandButton Command10 
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               Caption         =   "Et. Bobina Reb."
               Height          =   420
               Left            =   3495
               Style           =   1  'Graphical
               TabIndex        =   83
               Top             =   120
               Width           =   915
            End
            Begin VB.Frame framepaperfrontal 
               Caption         =   "Paper Frontal"
               Height          =   1200
               Left            =   4470
               TabIndex        =   145
               Top             =   4275
               Width           =   5895
               Begin VB.CheckBox Check37 
                  Caption         =   "Packing Xr Palet (a Reb)"
                  DataField       =   "pfpackinglistXpalet"
                  DataSource      =   "envios"
                  Height          =   240
                  Left            =   1650
                  TabIndex        =   222
                  ToolTipText     =   "Sortirà a Rebobinadores al fer el paper frontal."
                  Top             =   930
                  Width           =   2295
               End
               Begin VB.ComboBox combocopiespaperfrontal 
                  DataField       =   "copiespaperfrontal"
                  DataSource      =   "envios"
                  Height          =   315
                  ItemData        =   "clients.frx":5784
                  Left            =   105
                  List            =   "clients.frx":5794
                  TabIndex        =   179
                  Text            =   "1"
                  Top             =   375
                  Width           =   555
               End
               Begin VB.ComboBox estilpaperfrontal 
                  DataField       =   "estilfrontal"
                  DataSource      =   "envios"
                  Height          =   315
                  ItemData        =   "clients.frx":57A4
                  Left            =   4005
                  List            =   "clients.frx":57B1
                  TabIndex        =   146
                  Top             =   375
                  Width           =   1695
               End
               Begin VB.CheckBox Check19 
                  Caption         =   "Pes Net"
                  DataField       =   "pfpesnet"
                  DataSource      =   "envios"
                  Height          =   240
                  Left            =   4080
                  TabIndex        =   150
                  Top             =   690
                  Width           =   1500
               End
               Begin VB.CheckBox Check18 
                  Caption         =   "Data fabricació - Data consum preferent"
                  DataField       =   "pfdatafab"
                  DataSource      =   "envios"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   149
                  Top             =   690
                  Width           =   3165
               End
               Begin VB.CheckBox Check22 
                  Caption         =   "Codi de Barres"
                  DataField       =   "pfcodibarres"
                  DataSource      =   "envios"
                  Height          =   195
                  Left            =   4080
                  TabIndex        =   148
                  Top             =   930
                  Width           =   2040
               End
               Begin VB.CheckBox Check23 
                  Caption         =   "Packing List"
                  DataField       =   "pfpacking"
                  DataSource      =   "envios"
                  Height          =   240
                  Left            =   180
                  TabIndex        =   147
                  Top             =   915
                  Width           =   1395
               End
               Begin VB.Label Label20 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Còpies"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   180
                  Top             =   180
                  Width           =   585
               End
               Begin VB.Label Label7 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Estil Paper Frontal"
                  Height          =   195
                  Left            =   4245
                  TabIndex        =   152
                  Top             =   165
                  Width           =   1500
               End
               Begin VB.Label Label8 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Col.locacio Paper Frontal"
                  Height          =   195
                  Left            =   1395
                  TabIndex        =   151
                  Top             =   165
                  Width           =   2430
               End
            End
         End
      End
      Begin VB.ComboBox combogrupclient 
         BackColor       =   &H00FFC0C0&
         DataField       =   "grupdeclient"
         DataSource      =   "clients"
         Height          =   315
         Left            =   4530
         TabIndex        =   169
         Top             =   180
         Width           =   2115
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Codis Comptables"
         Height          =   360
         Left            =   8640
         TabIndex        =   154
         Top             =   345
         Width           =   1485
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Enviaments"
         Height          =   360
         Left            =   7095
         TabIndex        =   80
         Top             =   345
         Width           =   1440
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Var"
         DataField       =   "var_com_rep"
         DataSource      =   "clients"
         Height          =   195
         Left            =   7590
         TabIndex        =   79
         Tag             =   "protegits"
         Top             =   3840
         Width           =   585
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Fix"
         DataField       =   "fix_com_rep"
         DataSource      =   "clients"
         Height          =   195
         Left            =   7590
         TabIndex        =   21
         Tag             =   "protegits"
         Top             =   3660
         Width           =   585
      End
      Begin VB.TextBox Text44 
         BackColor       =   &H80000018&
         DataField       =   "com_representant"
         DataSource      =   "clients"
         Height          =   285
         Left            =   7260
         TabIndex        =   20
         Tag             =   "protegits"
         Top             =   3705
         Width           =   330
      End
      Begin VB.TextBox Text14 
         DataField       =   "provinciae"
         DataSource      =   "clients"
         Height          =   285
         Left            =   10545
         TabIndex        =   16
         Top             =   1335
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Possar comanda de clixes del client a l'albarà."
         DataField       =   "impagats"
         DataSource      =   "clients"
         Height          =   375
         Left            =   3570
         TabIndex        =   74
         Top             =   3630
         Width           =   2250
      End
      Begin VB.TextBox riscpla 
         DataField       =   "importriscpla"
         DataSource      =   "clients"
         Height          =   330
         Left            =   5760
         TabIndex        =   72
         Top             =   3225
         Width           =   840
      End
      Begin VB.ComboBox companyiarisc 
         DataField       =   "companyiacredit"
         DataSource      =   "clients"
         Height          =   315
         ItemData        =   "clients.frx":57F0
         Left            =   1425
         List            =   "clients.frx":57FA
         TabIndex        =   70
         Top             =   3225
         Width           =   2715
      End
      Begin VB.TextBox risc 
         DataField       =   "importrisc"
         DataSource      =   "clients"
         Height          =   330
         Left            =   4500
         TabIndex        =   69
         Top             =   3225
         Width           =   840
      End
      Begin VB.TextBox Text43 
         BackColor       =   &H80000018&
         DataField       =   "obsultima"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   31
         Tag             =   "protegits"
         Top             =   6075
         Width           =   7335
      End
      Begin VB.CheckBox Check7 
         Caption         =   "1 REF. x PALET"
         DataField       =   "refpalet"
         DataSource      =   "clients"
         Height          =   255
         Left            =   8355
         TabIndex        =   66
         Top             =   3690
         Width           =   1815
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Paper Frontal Palet"
         DataField       =   "paperfrontal"
         DataSource      =   "clients"
         Height          =   255
         Left            =   8355
         TabIndex        =   62
         Top             =   4269
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox Text25 
         DataField       =   "email"
         DataSource      =   "clients"
         Height          =   285
         Left            =   855
         TabIndex        =   61
         Top             =   2895
         Width           =   4785
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   8700
         Top             =   150
      End
      Begin VB.TextBox Text40 
         BackColor       =   &H80000018&
         DataField       =   "obssol1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   30
         Tag             =   "protegits"
         Top             =   5730
         Width           =   7335
      End
      Begin VB.TextBox Text38 
         BackColor       =   &H80000018&
         DataField       =   "obsreb1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   29
         Tag             =   "protegits"
         Top             =   5385
         Width           =   7335
      End
      Begin VB.TextBox Text36 
         BackColor       =   &H80000018&
         DataField       =   "obslam1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   28
         Tag             =   "protegits"
         Top             =   5040
         Width           =   7335
      End
      Begin VB.TextBox Text34 
         BackColor       =   &H80000018&
         DataField       =   "obsimp1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   27
         Tag             =   "protegits"
         Top             =   4710
         Width           =   7335
      End
      Begin VB.TextBox Text32 
         BackColor       =   &H80000018&
         DataField       =   "obsext1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   26
         Tag             =   "protegits"
         Top             =   4380
         Width           =   7335
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000018&
         DataField       =   "observacions1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   25
         Tag             =   "protegits"
         Top             =   4050
         Width           =   7335
      End
      Begin VB.TextBox Text31 
         DataField       =   "horaridesc"
         DataSource      =   "clients"
         Height          =   285
         Left            =   7920
         TabIndex        =   17
         Top             =   2385
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Certificat Qualitat"
         DataField       =   "certqualitat"
         DataSource      =   "clients"
         Height          =   255
         Left            =   8355
         TabIndex        =   24
         Top             =   3945
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Albara Valorat"
         DataField       =   "albvalorat"
         DataSource      =   "clients"
         Height          =   255
         Left            =   8355
         TabIndex        =   23
         Top             =   3675
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text30 
         DataField       =   "numproveidor"
         DataSource      =   "clients"
         Height          =   285
         Left            =   7920
         TabIndex        =   18
         Top             =   2670
         Width           =   2415
      End
      Begin VB.TextBox Text29 
         BackColor       =   &H80000018&
         DataSource      =   "clients"
         Height          =   285
         Left            =   7920
         TabIndex        =   33
         Tag             =   "protegits"
         Top             =   3000
         Width           =   2445
      End
      Begin VB.TextBox Text28 
         BackColor       =   &H80000018&
         DataField       =   "formapag"
         DataSource      =   "clients"
         Height          =   285
         Left            =   7245
         TabIndex        =   22
         Tag             =   "protegits"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text27 
         BackColor       =   &H80000018&
         DataSource      =   "clients"
         Height          =   285
         Left            =   7920
         TabIndex        =   32
         Tag             =   "protegits"
         Top             =   3345
         Width           =   2460
      End
      Begin VB.TextBox Text26 
         BackColor       =   &H80000018&
         DataField       =   "representant"
         DataSource      =   "clients"
         Height          =   285
         Left            =   7245
         TabIndex        =   19
         Tag             =   "protegits"
         Top             =   3345
         Width           =   615
      End
      Begin VB.TextBox Text24 
         DataField       =   "obsfax2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   2760
         TabIndex        =   15
         Top             =   2595
         Width           =   3855
      End
      Begin VB.TextBox Text23 
         DataField       =   "fax2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   2595
         Width           =   1815
      End
      Begin VB.TextBox Text22 
         DataField       =   "obsfax1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   2760
         TabIndex        =   13
         Top             =   2310
         Width           =   3855
      End
      Begin VB.TextBox Text21 
         DataField       =   "fax1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Top             =   2310
         Width           =   1815
      End
      Begin VB.TextBox Text20 
         DataField       =   "obstel2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   2760
         TabIndex        =   11
         Top             =   2025
         Width           =   3855
      End
      Begin VB.TextBox Text19 
         DataField       =   "telefon2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Top             =   2025
         Width           =   1815
      End
      Begin VB.TextBox Text18 
         DataField       =   "obstel1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   2760
         TabIndex        =   9
         Top             =   1740
         Width           =   3855
      End
      Begin VB.TextBox Text17 
         DataField       =   "telefon1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   1740
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         DataField       =   "provincia"
         DataSource      =   "clients"
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   1455
         Width           =   4455
      End
      Begin VB.TextBox Text5 
         DataField       =   "codipostal"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   1455
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         DataField       =   "poblacio"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   1170
         Width           =   5775
      End
      Begin VB.TextBox Text3 
         DataField       =   "domicili"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   885
         Width           =   5775
      End
      Begin VB.TextBox Text2 
         DataField       =   "nom"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox Text1 
         DataField       =   "codi"
         DataSource      =   "clients"
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox Check27 
         Caption         =   "Avisar al Client per venir a revisar impressions Noves o Modificades."
         DataField       =   "clientvindraarevisarimpresio"
         DataSource      =   "clients"
         Height          =   405
         Left            =   780
         TabIndex        =   178
         Top             =   3585
         Width           =   2865
      End
      Begin VB.CheckBox Check28 
         Caption         =   "Obligar a demanar Quantitat demanada a la comanda"
         DataField       =   "obligatquantitatdemanada"
         DataSource      =   "clients"
         Height          =   405
         Left            =   6690
         TabIndex        =   182
         Top             =   1905
         Width           =   2565
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Valors espessor laminat i tolerancia"
         Height          =   975
         Left            =   6945
         TabIndex        =   187
         Top             =   900
         Width           =   3885
         Begin VB.TextBox Text15 
            DataField       =   "espessortinta"
            DataSource      =   "clients"
            Height          =   285
            Left            =   570
            TabIndex        =   190
            Top             =   225
            Width           =   675
         End
         Begin VB.TextBox Text16 
            DataField       =   "espessorcola"
            DataSource      =   "clients"
            Height          =   285
            Left            =   2370
            TabIndex        =   189
            Top             =   255
            Width           =   675
         End
         Begin VB.TextBox Text45 
            DataField       =   "espessortolerancia"
            DataSource      =   "clients"
            Height          =   285
            Left            =   930
            TabIndex        =   188
            Top             =   615
            Width           =   495
         End
         Begin VB.Label Label21 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Tinta:"
            Height          =   180
            Left            =   90
            TabIndex        =   196
            Top             =   300
            Width           =   510
         End
         Begin VB.Label Label22 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Cola:"
            Height          =   180
            Left            =   1950
            TabIndex        =   195
            Top             =   315
            Width           =   510
         End
         Begin VB.Label Label23 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Tolerancia:"
            Height          =   180
            Left            =   90
            TabIndex        =   194
            Top             =   660
            Width           =   870
         End
         Begin VB.Label Label24 
            BackColor       =   &H00EAD9CE&
            Caption         =   "µm"
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
            Left            =   1275
            TabIndex        =   193
            Top             =   225
            Width           =   510
         End
         Begin VB.Label Label25 
            BackColor       =   &H00EAD9CE&
            Caption         =   "µm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3090
            TabIndex        =   192
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label26 
            BackColor       =   &H00EAD9CE&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1470
            TabIndex        =   191
            Top             =   630
            Width           =   510
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FDDECE&
         Caption         =   "Rappels"
         Height          =   1890
         Left            =   8145
         TabIndex        =   208
         Top             =   4500
         Width           =   2970
         Begin VB.CommandButton Command13 
            Height          =   285
            Left            =   90
            Picture         =   "clients.frx":5811
            Style           =   1  'Graphical
            TabIndex        =   209
            ToolTipText     =   "Eliminacio Registres"
            Top             =   210
            Width           =   345
         End
         Begin VB.Data Datarappels 
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
         Begin MSDBGrid.DBGrid reixarappels 
            Bindings        =   "clients.frx":5D9B
            Height          =   1320
            Left            =   60
            OleObjectBlob   =   "clients.frx":5DB1
            TabIndex        =   210
            Top             =   510
            Width           =   2775
         End
         Begin VB.Label etgrupdeclients 
            BackStyle       =   0  'Transparent
            DataSource      =   "Datarappels"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H005C31DD&
            Height          =   300
            Left            =   675
            TabIndex        =   211
            Top             =   210
            Width           =   2070
         End
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Guarda Mostra"
         DataField       =   "guardarmostra"
         DataSource      =   "clients"
         Height          =   255
         Left            =   8355
         TabIndex        =   63
         Top             =   4566
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Embalatge Anònim"
         DataField       =   "anonim"
         DataSource      =   "clients"
         Height          =   255
         Left            =   8355
         TabIndex        =   64
         Top             =   4863
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Palet Europeu"
         DataField       =   "europeu"
         DataSource      =   "clients"
         Height          =   255
         Left            =   8370
         TabIndex        =   65
         Top             =   5730
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox Check34 
         Caption         =   "Expedicions enviar albarà a oficines"
         DataField       =   "albaraalpalet"
         DataSource      =   "clients"
         Height          =   465
         Left            =   5820
         TabIndex        =   212
         Tag             =   "protegits"
         Top             =   3570
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.Label Label1 
         Caption         =   "Nom del Grup de Clients:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   24
         Left            =   2715
         TabIndex        =   168
         Top             =   210
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "% Com. Repr:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   18
         Left            =   6255
         TabIndex        =   78
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Pla:"
         Height          =   315
         Left            =   5475
         TabIndex        =   73
         Top             =   3300
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Ipc:"
         Height          =   315
         Left            =   4200
         TabIndex        =   71
         Top             =   3300
         Width           =   390
      End
      Begin VB.Label Label1 
         Caption         =   "Companyia Credit:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   16
         Left            =   150
         TabIndex        =   68
         Top             =   3300
         Width           =   1890
      End
      Begin VB.Label Label1 
         Caption         =   "Obs.Ult."
         DataSource      =   "clients"
         Height          =   255
         Index           =   27
         Left            =   60
         TabIndex        =   67
         Top             =   6150
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Obs.Sol."
         DataSource      =   "clients"
         Height          =   255
         Index           =   25
         Left            =   60
         TabIndex        =   53
         Top             =   5775
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Obs.Reb."
         DataSource      =   "clients"
         Height          =   255
         Index           =   23
         Left            =   60
         TabIndex        =   52
         Top             =   5415
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Obs.Lam"
         DataSource      =   "clients"
         Height          =   255
         Index           =   21
         Left            =   60
         TabIndex        =   51
         Top             =   5070
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Obs.Imp."
         DataSource      =   "clients"
         Height          =   255
         Index           =   19
         Left            =   60
         TabIndex        =   50
         Top             =   4725
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Obs.Ext."
         DataSource      =   "clients"
         Height          =   255
         Index           =   17
         Left            =   60
         TabIndex        =   49
         Top             =   4395
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Obs.  "
         DataSource      =   "clients"
         Height          =   255
         Index           =   15
         Left            =   75
         TabIndex        =   48
         Top             =   4050
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Codi:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   5
         Left            =   75
         TabIndex        =   47
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Num Prov:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   4
         Left            =   6795
         TabIndex        =   46
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Horari Entrega"
         DataSource      =   "clients"
         Height          =   255
         Index           =   14
         Left            =   6765
         TabIndex        =   45
         Top             =   2460
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "F.Pag:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   13
         Left            =   6735
         TabIndex        =   44
         Top             =   3015
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Repr:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   12
         Left            =   6735
         TabIndex        =   43
         Top             =   3345
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail"
         DataSource      =   "clients"
         Height          =   255
         Index           =   11
         Left            =   150
         TabIndex        =   42
         Top             =   2955
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Fax2"
         DataSource      =   "clients"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   41
         Top             =   2670
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Fax1"
         DataSource      =   "clients"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   40
         Top             =   2385
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Telf2"
         DataSource      =   "clients"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   39
         Top             =   2100
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Telf1"
         DataSource      =   "clients"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   38
         Top             =   1815
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cp/Pr:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   37
         Top             =   1530
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Pob:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   1170
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Dom:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   885
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Nom:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "formclients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub camp1_Change()

End Sub

Private Sub alta_Click()
alta_registre
Frame2.Tag = ""
End Sub

Private Sub bmodificacions_Click(Index As Integer)
  If Index = 2 Then
     llistademodificacions
  End If
  
End Sub
Sub llistademodificacions()
  Unload formseleccio
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = clients.DatabaseName
  formseleccio.Data1.RecordSource = "select * from clients_controlcanvis where codiclient=" + atrim(clients.Recordset!codi)
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(1).Width = 1000
  'formseleccio.DBGrid2.Columns(0).Visible = False
  'formseleccio.DBGrid2.Columns(2).NumberFormat = "dd/mm/yy"
  formseleccio.DBGrid2.Columns(1).Width = 1000
  formseleccio.DBGrid2.Columns(2).Width = 1200
  formseleccio.DBGrid2.Columns(3).Width = 1700
  formseleccio.DBGrid2.Columns(4).Width = 1700
  formseleccio.DBGrid2.Columns(5).Width = 1500
  formseleccio.DBGrid2.Columns(6).Width = 1500
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Visible = False
  formseleccio.DBGrid2.Columns(2).NumberFormat = "dd/mm/yy hh:nn"
  If formseleccio.Data1.Recordset.EOF Then Exit Sub
  formseleccio.Width = 10000
  formseleccio.Show 1
  Unload formseleccio
End Sub
Private Sub Check10_Click()
 Check9.Value = 0
 Text44.Text = "0"
End Sub

Private Sub Check13_Click()
  If Check13.Value > 0 Then
     pesnetstd.Visible = True
     framepesnet.Visible = True
      carregar_pesosnets
       Else: framepesnet.Visible = False: pesnetstd.Visible = False
  End If
End Sub
Sub carregar_pesosnets()
  Dim idenv As Double
  idenv = cadbl(envios.Tag)
  If pesnetstd.Value = 1 Then idenv = -9999
  datapesnet.RecordSource = "select * from taulapesnet where idenvio=" + atrim(idenv)
  datapesnet.Refresh
  If datapesnet.Recordset.EOF Then
   If formclients.ActiveControl.Name = "Check13" Or formclients.ActiveControl.Name = "pesnetstd" Then
    If MsgBox("Vols crear una taula de pesos de canutu per aquest client?", vbYesNo, "Pesos Canutu") = vbYes Then
     For i = 0 To 120 Step 5
      datapesnet.Recordset.AddNew
      datapesnet.Recordset!idenvio = idenv
      datapesnet.Recordset!mida = i
      datapesnet.Recordset.Update
     Next i
    End If
    nomenvio.SetFocus
   End If
    Else: datapesnet.Recordset.MoveFirst
  End If
End Sub

Private Sub Check30_Click()
 If Check30.Value = 1 Then combotipus_sscc = "Albarà"
End Sub

Private Sub Check35_Click()
   If Check35.Value = 1 And Screen.ActiveControl.Name = "Check35" Then
        If combopais <> "ES" Then MsgBox "Aquest enviament no ès ESPANYOL, es correcte que marquis IMPOST INCLÒS?", vbCritical, "A T E N C I Ó"
   End If
End Sub

Private Sub Check9_Click()
Check10.Value = 0
End Sub

Private Sub clients_Reposition()
  If r <> "envio" Then carregar_lookups
  clients.Caption = "Clients:  " + atrim(cadbl(clients.Recordset.AbsolutePosition) + 1) + " de " + atrim(clients.Recordset.RecordCount)
  If clients.EditMode = 0 Then areadatos.Enabled = False
End Sub

Sub triarrepresentant()
  Load formseleccio
  formseleccio.Data1.DatabaseName = clients.DatabaseName
  formseleccio.Data1.RecordSource = "select * from representants"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text26.Text = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   Text27.Text = atrim(formseleccio.Data1.Recordset!nom)
  End If
  Unload formseleccio
  
End Sub

Sub triarformapag()
  Load formseleccio
  formseleccio.Data1.DatabaseName = clients.DatabaseName
  formseleccio.Data1.RecordSource = "select * from [formes de pagament]"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
    If cadbl((formseleccio.Data1.Recordset!codi)) > 0 Then
      Text28.Text = atrim((formseleccio.Data1.Recordset!codi))
      Text29.Text = atrim(formseleccio.Data1.Recordset!descripcio)
       Else
        Text28.Text = "0"
        Text29.Text = ""
    End If
  End If
  Unload formseleccio
  
End Sub


Private Sub combo_alcadapalet_Click()
 Dim combo As Object
 Dim vcamp As String
 Set combo = combo_alcadapalet
 vcamp = "alcadapalet"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_alcadapalet_KeyDown(KeyCode As Integer, Shift As Integer)
borrar_valor_combo KeyCode, "alcadapalet"
End Sub

Private Sub combo_cert_qualitat_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_cert_qualitat
 vcamp = "cert_qualitat"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_cert_qualitat_KeyDown(KeyCode As Integer, Shift As Integer)
borrar_valor_combo KeyCode, "cert_qualitat"
End Sub

Private Sub combo_conosprotectors_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_conosprotectors
 vcamp = "conosprotectors"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_conosprotectors_KeyDown(KeyCode As Integer, Shift As Integer)
borrar_valor_combo KeyCode, "conosprotectors"
End Sub

Private Sub combo_emb_anonim_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_emb_anonim
 vcamp = "emb_anonim"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_emb_anonim_KeyDown(KeyCode As Integer, Shift As Integer)
borrar_valor_combo KeyCode, "emb_anonim"
End Sub

Private Sub combo_guardarmostres_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_guardarmostres
 vcamp = "guardarmostres"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_guardarmostres_KeyDown(KeyCode As Integer, Shift As Integer)
borrar_valor_combo KeyCode, "guardarmostres"
End Sub

Private Sub combo_guardarmostressol_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_guardarmostressol
 vcamp = "guardarmostressol"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_guardarmostressol_KeyDown(KeyCode As Integer, Shift As Integer)
borrar_valor_combo KeyCode, "guardarmostressol"
End Sub

Private Sub combo_paperfrontal_Click()
Dim combo As Object
 Dim vcamp As String
 
 If atrim(combo_paperfrontal.Text) = "" Then
     framepaperfrontal.Enabled = False
       Else: framepaperfrontal.Enabled = True
 End If
 
 Set combo = combo_paperfrontal
 vcamp = "pfpaperfrontal"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_paperfrontal_KeyDown(KeyCode As Integer, Shift As Integer)
borrar_valor_combo KeyCode, "pfpaperfrontal"
End Sub

Private Sub Combo_peuimprenta_Click()
  If Combo_peuimprenta.ListIndex <> -1 Then envios.Recordset!peuimprenta = Combo_peuimprenta.ItemData(Combo_peuimprenta.ListIndex)
End Sub

Private Sub Combo_peuimprenta_KeyDown(KeyCode As Integer, Shift As Integer)
  borrar_valor_combo KeyCode, "peuimprenta"
 
End Sub
Sub borrar_valor_combo(KeyCode As Integer, campid As String)
If KeyCode = 46 Then envios.Recordset.Fields(campid) = 0: ActiveControl.Text = ""

End Sub

Private Sub combo_protecciob_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_protecciob
 vcamp = "tipusprotecciob"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)

End Sub

Private Sub combo_protecciob_KeyDown(KeyCode As Integer, Shift As Integer)
borrar_valor_combo KeyCode, "tipusprotecciob"
End Sub

Private Sub combo_protecciop_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_protecciop
 vcamp = "tipusprotecciop"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_protecciop_KeyDown(KeyCode As Integer, Shift As Integer)
borrar_valor_combo KeyCode, "tipusprotecciop"
End Sub

Private Sub combo_protecciospr_Click()
Dim combo As Object
 Dim vcamp As String
 Set combo = combo_protecciospr
 vcamp = "tipusprotecciospr"
 If combo.ListIndex <> -1 Then envios.Recordset.Fields(vcamp) = combo.ItemData(combo.ListIndex)
End Sub

Private Sub combo_protecciospr_KeyDown(KeyCode As Integer, Shift As Integer)
borrar_valor_combo KeyCode, "tipusprotecciospr"
End Sub

Private Sub combo_tipuspalet_Click()
 If combo_tipuspalet.ListIndex <> -1 Then envios.Recordset!tipuspalet = combo_tipuspalet.ItemData(combo_tipuspalet.ListIndex)
End Sub

Private Sub combo_tipuspalet_KeyDown(KeyCode As Integer, Shift As Integer)
borrar_valor_combo KeyCode, "tipuspalet"
End Sub

Sub copiarclientaenviament()
 If envios.Recordset.EOF Then
    'If envios.Recordset.EOF Then
     MsgBox "Crearé una adreça d'enviament amb les dades principals del client"
     
     envios.Recordset.AddNew
     envios.Recordset!codi = clients.Recordset!codi
     envios.Recordset!nome = atrim(clients.Recordset!nom)
     envios.Recordset!domicilie = atrim(clients.Recordset!domicili)
     envios.Recordset!poblacioe = atrim(clients.Recordset!poblacio)
     envios.Recordset!codipostale = atrim(clients.Recordset!codipostal)
     envios.Recordset!provinciae = atrim(clients.Recordset!provincia)
     envios.Recordset!telefone = atrim(clients.Recordset!telefon1)
     envios.Recordset!faxe = atrim(clients.Recordset!fax1)
     If cadbl(clients.Recordset!albvalorat) <> 0 Then envios.Recordset!albaravalorat = 1
     If cadbl(clients.Recordset!certqualitat) <> 0 Then envios.Recordset!cert_qualitat = 1
     If cadbl(clients.Recordset!paperfrontal) <> 0 Then envios.Recordset!pfpaperfrontal = 1
     If cadbl(clients.Recordset!guardarmostra) <> 0 Then envios.Recordset!guardarmostres = 4
     If cadbl(clients.Recordset!anonim) <> 0 Then envios.Recordset!emb_anonim = 1
     If cadbl(clients.Recordset!europeu) <> 0 Then envios.Recordset!tipuspalet = 2
     envios.Recordset.Update
    'End If
    DoEvents
 End If
End Sub

Sub escullirpais()
   Load formseleccio
   formseleccio.Caption = "Escull un pais"
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codipais,nompais from paisos "
  formseleccio.refrescar
'  formseleccio.DBGrid2.Columns(2).Width = 900
  formseleccio.Width = 9000
  'formseleccio.Left = formseleccio.Left - 3000
  formseleccio.Show 1
  
   If seleccioret = 1 Then
       combopais = formseleccio.DBGrid2.Columns("codipais")
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub combocopiespaperfrontal_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub combogrupclient_DropDown()
    carregar_grupdeclients
End Sub
Sub carregar_grupdeclients()
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select distinct(grupdeclient) as nomgrup from clients order by grupdeclient")
  combogrupclient.Clear
  combogrupclient.AddItem ""
  While Not rst.EOF
    If atrim(rst!nomgrup) <> "" Then
        combogrupclient.AddItem atrim(rst!nomgrup)
    End If
    rst.MoveNext
  Wend
End Sub

Private Sub comboincoterms_DropDown()
  Dim v As String
   v = InputBox("Escriu els INCOTERMS que pot utilitzar aquest client separats per espais." + vbNewLine + " Ex: DAP EXW CIF FCA", "INCOTERMS", atrim(comboincoterms))
   If StrPtr(v) = 0 Then Exit Sub
   If atrim(v) = "" Then Exit Sub
   comboincoterms = atrim(v)
End Sub

Private Sub Combopais_DropDown()
escullirpais
End Sub

Private Sub Combopais_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub Combopais_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub combotipus_sscc_DropDown()
  If Check30.Value <> 0 Then SendKeys "{TAB}"
End Sub

Private Sub combotipus_sscc_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub combotransportfavorit_DropDown()
   Dim vnomtransport As String
   Dim v As Long
   v = Menu.escullir_transportista(vnomtransport)
   If cadbl(v) > 0 Then
    envios.Recordset.id_transportFAVORIT = v
    envios.Recordset!nom_transportFAVORIT = vnomtransport
    envios.UpdateControls
      Else
        If MsgBox("Vols treure el transportista FAVORIT d'aquest client?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
            envios.Recordset.id_transportFAVORIT = 0
            envios.Recordset!nom_transportFAVORIT = ""
            envios.UpdateControls
        End If
   End If
End Sub

Private Sub Command10_Click()
  Unload etbobina
  Load etbobina
  etbobina.Visible = True
  etbobina.SetFocus
End Sub

Private Sub Command11_Click()
  Load formaltarep
  formaltarep.Caption = "Manteniment Codis Comptables"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from clients_codiscomptables where codifabricacio=" + atrim(formclients.clients.Recordset!codi)
  formaltarep.refrescar
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.DBGrid1.Columns(1).Visible = False
  formaltarep.DBGrid1.Columns(2).Width = 2000
  formaltarep.DBGrid1.Columns(3).Width = 3500
  formaltarep.Width = 8500
  formaltarep.DBGrid1.Refresh
  formaltarep.alta.Tag = "codifabricacio"
  formaltarep.alta.HelpContextID = clients.Recordset!codi
  formaltarep.Show 1
End Sub

Private Sub Command12_Click()
   Load formliniespeualbara
   formliniespeualbara.Tag = atrim(envios.Recordset!ID)
   formliniespeualbara.Show 1
   Unload formliniespeualbara
End Sub

Private Sub Command13_Click()
If Datarappels.Recordset.EOF Then MsgBox "Escull el rappel per eliminar", vbCritical, "Error": Exit Sub
   If MsgBox("Segur que vols eliminar aquest rappel?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
        Datarappels.Recordset.Delete
        Datarappels.Refresh
   End If
End Sub

Private Sub Command3_Click()
  comandes_afectades
End Sub
Sub comandes_afectades(Optional where As String)
 If where = "" Then where = "client = " + atrim(cadbl(clients.Recordset!codi))
If Not comandesafectades.Visible Then
    comandesafectades.Left = 1725
    comandesafectades.Top = 1275
    comandesafectades.ZOrder 0
    comandesafectades.Visible = True
    comandesafectades.Caption = "Buscant comandes un moment sisplau..."
    DoEvents
    ratoli "espera"
    comandes.RecordSource = "select comanda,texteimpressio,puntrisc from comandes where proximaseccio<>'T' and " + where + " order by comanda DESC"
'    comandes.RecordSource = "select comanda,texteimpressio,'            ' as RISC, puntrisc from comandes "
    comandes.Refresh
    While comandes.Recordset.RecordCount = 1
      DoEvents
    Wend
    reixa.Rows = 500
    i = 0
    reixa.row = 0
    reixa.ColWidth(0) = 100 * 9: reixa.ColWidth(1) = 1000 * 4: reixa.ColWidth(2) = 300 * 4
    reixa.col = 2
    DoEvents
    reixa.TextMatrix(0, 0) = "COMANDA": reixa.TextMatrix(0, 1) = "TEXTE": reixa.TextMatrix(0, 2) = "RISC"
    While Not comandes.Recordset.EOF
       'If Not comandes.Recordset.EOF Then reixa.Rows = comandes.Recordset.RecordCount+1
       i = comandes.Recordset.AbsolutePosition + 1
       reixa.TextMatrix(i, 1) = atrim(comandes.Recordset!texteimpressio)
       reixa.TextMatrix(i, 0) = atrim(comandes.Recordset!comanda)
       reixa.TextMatrix(i, 2) = IIf(comandes.Recordset!puntrisc = 1, "Vermell", IIf(comandes.Recordset!puntrisc = 2, "Verd", ""))
       reixa.row = i
       reixa.CellBackColor = IIf(reixa.Text = "Vermell", QBColor(12), IIf(reixa.Text = "Verd", QBColor(10), QBColor(15)))
       comandes.Recordset.MoveNext
    Wend
    ratoli "normal"
    comandesafectades.Caption = "Comandes Afectades"
   Else: comandesafectades.Visible = False
  End If

End Sub
Private Sub Command1_Click()
 r = obre_fitxer(ruta_relativa_docs, 2)
 If atrim(r) = "" And atrim(Text41) <> "" Then If MsgBox("Vols borrar el fitxer relacionat?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
 Text41 = Mid(r, Len(ruta_relativa_docs) + 2)
End Sub

Private Sub Command2_Click()
 r = obre_fitxer(ruta_relativa_docs, 2)
 If atrim(r) = "" And atrim(Text42) <> "" Then If MsgBox("Vols borrar el fitxer relacionat?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
 Text42 = Mid(r, Len(ruta_relativa_docs) + 2)
End Sub

Private Sub Command4_Click()
   Dim rst As Recordset
   Dim dbtmpt As Database
   
   
   If missatgeenviament <> "" Then MsgBox "Aquestes dades són genèriques del client hauries d'eliminar el client": Exit Sub
   If Not envios.Recordset.EOF Then
    
     DoEvents
     Set dbtmpt = OpenDatabase(cami)
     
     MsgBox "Primer miraré si hi ha alguna comanda relacionada.", vbInformation, "Atenció"
     ratoli "espera"
     Set rst = dbtmpt.OpenRecordset("select direnvio from comandes where client=" + atrim(envios.Recordset!codi))
     While Not rst.EOF
       If rst!direnvio = envios.Recordset!ID Then
          ratoli "normal"
          MsgBox "Aquesta direcció d'enviament no es pot borrar perquè tè comandes relacionades", vbCritical, "Atenció": GoTo fi
       End If
       rst.MoveNext
     Wend
     ratoli "normal"
     r = InputBox("Per eliminar aquest envio has d'escriure [eliminar]", "Atenció")
      '+ Chr(10) + Chr(13) + "RECORDA QUE TOTES LES COMANDES EFECTADES QUEDERAN DESVINCULADES DE LA DIRECCIÓ D'ENVIO
     If r = "eliminar" Then
       dbtmp.Execute "delete * from taulapesnet where idenvio=" + atrim(cadbl(envios.Tag))
       envios.Recordset.Delete
       envios.Refresh
       If envios.Recordset.EOF Then Command7_Click
      End If
     'End If
   End If
   comandesafectades.Visible = False
fi:
   Set dbtmpt = Nothing
   Set rst = Nothing
   
End Sub

Private Sub Command5_Click()
  If envios.Recordset.EOF Then
      copiarclientaenviament
      envios.Refresh
     Else
       If MsgBox("Vols copiar alguna direcció d'enviament?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then
                buscant = True
                envios.Recordset.AddNew
                buscant = False
                envios.Recordset!codi = clients.Recordset!codi
                nomenvio.SetFocus
                  Else: copiardirecciodenviament
       End If
   End If
End Sub
Sub copiardirecciodenviament()
   Dim direnvioescullida As Long
   Dim rste As Recordset
   Dim i As Integer
   envios.Recordset.MoveLast
   direnvioescullida = 0
   If envios.Recordset.RecordCount <> 1 Then
       direnvioescullida = triar_client_direnvio
       If direnvioescullida = 0 Then Exit Sub
        Else: direnvioescullida = envios.Recordset!ID
   End If
   Set rste = dbtmp.OpenRecordset("select * from clients_envios where id=" + atrim(direnvioescullida))
   If rste.EOF Then Exit Sub
   buscant = True
   envios.Recordset.AddNew
   buscant = False
   For i = 0 To rste.Fields.Count - 1
     If rste.Fields(i).Name <> "id" Then
      envios.Recordset.Fields(i) = rste.Fields(i)
     End If
   Next i
   envios.Recordset.Update
   envios.Recordset.MoveLast
   Set rste = Nothing
End Sub
Function triar_client_direnvio() As Long
   Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select id ,domicilie,poblacioe,provinciae from clients_envios where codi=" + atrim(cadbl(clients.Recordset!codi))
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(2).Width = 900
  formseleccio.Width = 9000
  formseleccio.Left = formseleccio.Left - 3000
  If formseleccio.Data1.Recordset.EOF Then MsgBox "Aquest client no te cap DIRECCIO D'ENVIO ASSIGNADA.": Exit Function
  formseleccio.Data1.Recordset.MoveLast
  formseleccio.Data1.Recordset.MoveFirst
  If formseleccio.Data1.Recordset.RecordCount > 1 Then
                                                                                                                
     formseleccio.Show 1
     While formseleccio.Visible And seleccioret = 0
       DoEvents
     Wend
    Else: seleccioret = 1
  End If
  
  
  
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           triar_client_direnvio = cadbl(formseleccio.DBGrid2.Columns("id"))
           'campid_treball.Tag = cadbl(formseleccio.DBGrid2.Columns("ordremodificacio"))
        End If
   End If
    If seleccioret = 9 Then
       triar_client_direnvio = 0

        'campid_treball.Tag = ""
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Function


Private Sub Command6_Click()
  Dim idactual As Integer
  If envios.Recordset.EditMode <> 0 Then
     envios.Recordset.Update
  End If
  If Not envios.Recordset.EOF Then idactual = envios.Recordset!ID
  missatgeenviament.Caption = "": envios.RecordSource = " select * from Clients_envios where codi=" + atrim(cadbl(clients.Recordset!codi))
  envios.Refresh
  envios.Recordset.MoveLast
  If Not envios.Recordset.EOF Then envios.Recordset.FindFirst "id=" + atrim(idactual)
End Sub

Private Sub Command7_Click()
If buscant Then Exit Sub
scroll.Visible = Not scroll.Visible
If scroll.Visible Then
  scroll.Left = 75
  scroll.Top = 900
  scroll.ZOrder 0
End If
palets.Visible = scroll.Visible
If Not scroll.Visible Then Exit Sub
nomenvio.SetFocus
If clients.Recordset.EditMode > 0 Then r = "envio": clients.Recordset.Update: clients.Recordset.Edit: areadatos.Enabled = True
envios.RecordSource = " select * from Clients_envios where codi=" + atrim(cadbl(clients.Recordset!codi))
envios.Refresh
If Not envios.Recordset.EOF Then envios.Recordset.MoveLast
missatgeenviament.Caption = ""
'If envios.Recordset.EOF Then
'     missatgeenviament.Caption = "Enviament únic. DADES GENÈRIQUES"
'     Set envios.Recordset = clients.Recordset
'     envios.RecordSource = clients.RecordSource
'     envios.Tag = cadbl(clients.Recordset!codi) - (cadbl(clients.Recordset!codi) * 2)
'   Else:
'      missatgeenviament.Caption = ""
'      envios.RecordSource = " select * from Clients_envios where codi=" + atrim(cadbl(clients.Recordset!codi))
'      envios.Refresh
'
'End If
If scroll.Visible Then
   carregar_combos
  ' kg.Value = 0: mtrs.Value = 0: unitats.Value = 0: peces.Value = 0: mt2.Value = 0: km.Value = 0: emiler.Value = 0
End If
'si s'ha accedit desde comandes al mirar el client carrega l'id d'envio de la comanda
If cadbl(Frame2.Tag) > 0 Then envios.Recordset.FindFirst "id=" + atrim(cadbl(Frame2.Tag))
If clients.EditMode > 0 And Not envios.Recordset.EOF Then envios.Recordset.Edit
End Sub
Sub carregar_combos()
   Dim dbcomandes As Database
   Dim rstc As Recordset
   Dim combo As Object
   Dim Combo2 As Object
   Dim Combo3 As Object
   Set dbcomandes = OpenDatabase(llegir_ini("General", "cami", fitxerini))
   'descripcio families productes
'   Set rstc = dbcomandes.OpenRecordset("select distinct familia from productes")
'   Set combo = qproducte
'   r = combo.Text: combo.Clear: combo.Text = r
'   While Not rstc.EOF
'     If atrim(rstc!familia) <> "" Then
'        r = atrim(rstc!familia)
'        Select Case (atrim(rstc!familia))
 '         Case "B"
 '            r = "Bosses"
 '         Case "F"
 '            r = "Formats"
 '       End Select
 '       combo.AddItem r
 '    End If
 '    'combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
 '    rstc.MoveNext
 '  Wend
   
  'tipus paper frontal
   Set rstc = dbcomandes.OpenRecordset("tipuspaperfrontal")
   Set combo = combo_paperfrontal
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
   If atrim(combo_paperfrontal.Text) = "" Then
     framepaperfrontal.Enabled = False
       Else: framepaperfrontal.Enabled = True
   End If
 

'peu imprenta i data
   Set rstc = dbcomandes.OpenRecordset("peuimprenta")
   Set combo = Combo_peuimprenta
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend

'tipus palets
   Set rstc = dbcomandes.OpenRecordset("tipuspalets")
   Set combo = combo_tipuspalet
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
'alçades palet
   Set rstc = dbcomandes.OpenRecordset("select * from alcadespalets order by descripcio")
   Set combo = combo_alcadapalet
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
 'tipus proteccio
   Set rstc = dbcomandes.OpenRecordset("tipusproteccions")
   Set combo = combo_protecciob: Set Combo2 = combo_protecciop: Set Combo3 = combo_protecciospr
   r = combo.Text: combo.Clear: combo.Text = r
   r = Combo2.Text: Combo2.Clear: Combo2.Text = r
   r = Combo3.Text: Combo3.Clear: Combo3.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio): combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     Combo2.AddItem atrim(rstc!descripcio): Combo2.ItemData(Combo2.NewIndex) = cadbl(rstc!codi)
     Combo3.AddItem atrim(rstc!descripcio): Combo3.ItemData(Combo3.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
'embalatges
'alçades palet
   Set rstc = dbcomandes.OpenRecordset("embalatgesanonims")
   Set combo = combo_emb_anonim
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
'certificat qualitat
'alçades palet
   Set rstc = dbcomandes.OpenRecordset("cert_qualitat")
   Set combo = combo_cert_qualitat
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
'guardarmostres reb
   'alçades palet
   Set rstc = dbcomandes.OpenRecordset("guardarmostres")
   Set combo = combo_guardarmostres
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
   
'guardarmostres reb
   'alçades palet
   Set rstc = dbcomandes.OpenRecordset("guardarmostres")
   Set combo = combo_guardarmostressol
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
   
'conos protectors
'alçades palet
   Set rstc = dbcomandes.OpenRecordset("conosprotectors")
   Set combo = combo_conosprotectors
   r = combo.Text: combo.Clear: combo.Text = r
   While Not rstc.EOF
     combo.AddItem atrim(rstc!descripcio)
     combo.ItemData(combo.NewIndex) = cadbl(rstc!codi)
     rstc.MoveNext
   Wend
End Sub

Private Sub Command8_Click()
  r = InputBox("Entra el pes que vols iguals a totes les mides.", "Entrada pes igual")
  If cadbl(r) >= 0 Then
    datapesnet.Recordset.MoveFirst
    While Not datapesnet.Recordset.EOF
       datapesnet.Recordset.Edit
       datapesnet.Recordset!pes = cadbl(r)
       datapesnet.Recordset.Update
       datapesnet.Recordset.MoveNext
    Wend
    datapesnet.Recordset.MoveFirst
  End If
End Sub

Private Sub Command9_Click()
 Dim taulatemp As String
 If clients.Tag = "" Then clients.Tag = " where codi>0"
  taulatemp = "c:\temporal.mdb"
  If existeix(taulatemp) Then Kill taulatemp
  DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
  dbtmp.Execute ("select * into temporal in '" + taulatemp + "' from clients " + clients.Tag)
 report.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "llistatclients1.rpt"
 report.DataFiles(0) = taulatemp
 report.Destination = crptToWindow
 report.Action = 1

End Sub

Private Sub consultar_Click()
  If clients.Recordset.EOF Then clients.RecordSource = "clients": clients.Refresh
  buscant = True
  alta_registre
  deixartotblanc
  Frame2.Tag = ""
  
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Datarappels_Reposition()
'  etgrupdeclients = ""
'  If Not Datarappels.Recordset.EOF And Not Datarappels.Recordset.BOF Then
'     If atrim(Datarappels.Recordset!grupdeclients) <> "" Then etgrupdeclients = "Grup: " + Datarappels.Recordset!grupdeclients
'  End If
End Sub

Private Sub eliminar_Click()
 On Error GoTo err
 Dim rst As Recordset
 Set rst = dbtmp.OpenRecordset("select * from comandes where client=" + atrim(clients.Recordset!codi))
 If Not rst.EOF Then MsgBox "Hi ha comandes fetes amb aquest client no es pot eliminar mentre hi hagui comandes fetes", vbCritical, "Error": Exit Sub
  If LCase(InputBox("Segur que vols Eliminar?" + Chr(10) + " Escriu [eliminar] per acceptar l'eliminació d'aquest client." + Chr(10) + "L'eliminació d'un client que te historial de comandes pot significar perdre informació d'aquestes.", "Atenció")) = "eliminar" Then
    envios.Refresh
    While Not envios.Recordset.EOF
      envios.Recordset.Delete
      envios.Recordset.MoveNext
    Wend
    clients.Recordset.Delete
    clients.Recordset.MoveNext
    If clients.Recordset.EOF Then clients.Recordset.MovePrevious
  End If
 Exit Sub
err:
  MsgBox "No s'ha pogut eliminar possiblement perque tingui registres relacionats. O bé no hi ha res per eliminar."
End Sub

Private Sub emiler_Click()
If Screen.ActiveControl.Name = "emiler" Then gravar_quantitat
End Sub

Private Sub envios_Reposition()
  If buscant Then Exit Sub
  carregar_altres_envio
  If clients.EditMode > 0 And Not envios.Recordset.EOF Then envios.Recordset.Edit
  If missatgeenviament <> "" Then
     envios.Caption = "-----": envios.Enabled = False
    Else:
       If Not envios.Recordset.EOF Then
         envios.Tag = envios.Recordset!ID
        Else: envios.Tag = cadbl(clients.Recordset!codi) - (cadbl(clients.Recordset!codi) * 2)
       End If
       envios.Enabled = True: envios.Caption = Trim(envios.Recordset.AbsolutePosition + 1) + "/" + Trim(envios.Recordset.RecordCount)
       If Check13.Value <> 0 Then carregar_pesosnets
  End If
     
End Sub

Sub obreenvios()
Set dbenvios = OpenDatabase(envios.DatabaseName)
End Sub

Sub carregar_altres_envio()
  obreenvios
  If envios.Recordset.EOF Then Exit Sub
  'qproducte.Text = ""
  combo_tipuspalet = possar_descripcio("tipuspalets", "descripcio", "codi", cadbl(envios.Recordset!tipuspalet))
  combo_alcadapalet = possar_descripcio("alcadespalets", "descripcio", "codi", cadbl(envios.Recordset!alcadapalet))
  combo_protecciospr = possar_descripcio("tipusproteccions", "descripcio", "codi", cadbl(envios.Recordset!tipusprotecciospr))
  combo_protecciob = possar_descripcio("tipusproteccions", "descripcio", "codi", cadbl(envios.Recordset!tipusprotecciob))
  combo_protecciop = possar_descripcio("tipusproteccions", "descripcio", "codi", cadbl(envios.Recordset!tipusprotecciop))
  combo_emb_anonim = possar_descripcio("embalatgesanonims", "descripcio", "codi", cadbl(envios.Recordset!emb_anonim))
  combo_cert_qualitat = possar_descripcio("cert_qualitat", "descripcio", "codi", cadbl(envios.Recordset!cert_qualitat))
  combo_guardarmostres = possar_descripcio("guardarmostres", "descripcio", "codi", cadbl(envios.Recordset!guardarmostres))
  combo_guardarmostressol = possar_descripcio("guardarmostres", "descripcio", "codi", cadbl(envios.Recordset!guardarmostressol))
  combo_conosprotectors = possar_descripcio("conosprotectors", "descripcio", "codi", cadbl(envios.Recordset!conosprotectors))
  combo_paperfrontal = possar_descripcio("tipuspaperfrontal", "descripcio", "codi", cadbl(envios.Recordset!pfpaperfrontal))
  Combo_peuimprenta = possar_descripcio("peuimprenta", "descripcio", "codi", cadbl(envios.Recordset!peuimprenta))
  Set dbenvios = Nothing
End Sub
Function possar_descripcio(vtaula As String, vdescripcio As String, vbuscara As String, vvalorbuscat As String)
  Dim rstenvio As Recordset
  Set rstenvio = dbenvios.OpenRecordset("Select " + vdescripcio + " from " + vtaula + " where " + vbuscara + "=" + atrim(cadbl(vvalorbuscat)))
  If Not rstenvio.EOF Then
        possar_descripcio = atrim(rstenvio.Fields(vdescripcio))
     Else: possar_descripcio = ""
  End If
  
End Function
Private Sub Form_Activate()
 Dim codiclient As Double
 codiclient = cadbl(llegir_ini("General", "clienttmp", fitxerini))
  If codiclient <> 0 Then
    clients.RecordSource = "select * from clients where codi=" + atrim(codiclient)
    clients.Refresh
    escriure_ini "General", "clienttmp", "", fitxerini
  End If
  colocarbloqueig
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 65 Then alta_registre: KeyCode = 0
'If KeyCode = 69 Then buscar_registre
If KeyCode = 27 Then cancelar_registre
If KeyCode = 112 Then gravar_registre
If KeyCode = 13 Then SendKeys "{TAB}": KeyCode = 0

End Sub
Sub buscar_registre()

End Sub
Sub alta_registre()
 If areadatos.Enabled = False Then
      areadatos.Enabled = True
      
      clients.Recordset.AddNew
      DoEvents
      Text1.Enabled = True
      'busco el mes gran i el poso a codi +1
      If Not buscant Then
        Set rsttmp = dbtmp.OpenRecordset("select max(codi) as [grancodi] from clients")
        If Not rsttmp.EOF Then
          Text1 = atrim(cadbl(rsttmp!grancodi) + 1)
         Else: Text1 = "1"
        End If
      End If
      Text1.SetFocus
 End If
End Sub
Sub gravar_registre()
 If areadatos.Enabled And Not buscant Then
    copiarclientaenviament
    Command6_Click 'gravar els enviaments
    scroll.Visible = False
    Text1.Enabled = False
    sortir.SetFocus
    DoEvents
    If Screen.ActiveControl.Name = "sortir" Then
      If envios.Recordset.EditMode > 0 Then envios.Recordset.Update
      If clients.EditMode > 0 Then clients.Recordset.Update
      areadatos.Enabled = False
      clients.Recordset.Bookmark = clients.Recordset.LastModified
    End If
    control_de_modificacions
 End If
 If buscant Then finalitzarbusqueda
End Sub


Sub cancelar_registre()
  If clients.Recordset.EditMode > 0 Then
   If envios.Recordset.EditMode > 0 Then envios.Recordset.CancelUpdate
   If clients.Recordset.EditMode > 0 Then clients.Recordset.CancelUpdate
   areadatos.Enabled = False
   Text1.Enabled = False
   buscant = False
   carregar_lookups
     Else: Unload Me
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
  If KeyAscii > 50 Then KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub Form_Load()
clients.DatabaseName = cami
comandes.DatabaseName = cami
envios.DatabaseName = cami
datapesnet.DatabaseName = cami
Datarappels.DatabaseName = cami

centerscreen Me
hora = Now
Set dbtmp = OpenDatabase(clients.DatabaseName)
clients.RecordSource = "clients"
clients.Refresh

possarvalordcamps

End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Set dbtmp = Nothing
End Sub

Private Sub gravar_Click()
gravar_registre
End Sub

Private Sub kg_Click()
'If Screen.ActiveControl.Name = "kg" Then gravar_quantitat
End Sub
Sub gravar_quantitat()
If qproducte = "" Then Exit Sub
  
  Set rsttmp = dbtmp.OpenRecordset("select * from unitatsxproducte where idproducte='" + atrim(qproducte) + "' and idenvio=" + atrim(cadbl(envios.Tag)))
  If rsttmp.EOF Then
     rsttmp.AddNew
     rsttmp!idenvio = envios.Tag
     rsttmp!idproducte = qproducte
   Else
     rsttmp.Edit
  End If
    rsttmp!kg = kg.Value
    rsttmp!mtrs = mtrs.Value
    rsttmp!pcs = peces.Value
    rsttmp!unts = unitats.Value
    rsttmp!mt2 = mt2.Value
    rsttmp!km = km.Value
    rsttmp!emiler = emiler.Value
  rsttmp.Update
End Sub

Private Sub km_Click()
If Screen.ActiveControl.Name = "km" Then gravar_quantitat
End Sub

Private Sub modificar_Click()
   If Not existeix("c:\ordprog.ini") Then
    If UCase(InputBox("Entra la contrasenya Inplacsa per poder editar.", "Contrasenya edició")) <> "INPLACSA" Then Exit Sub
   End If
   areadatos.Enabled = True
   guardar_controlcanvisclients
   envios.Refresh
   clients.Recordset.Edit
   
   Text2.SetFocus
End Sub
Sub control_de_modificacions()
   Dim i As Long
   If rstcontrolcanvis.EOF Then Exit Sub
   For i = 0 To clients.Recordset.Fields.Count - 1
      If rstcontrolcanvis.Fields(i) <> clients.Recordset.Fields(i) Then
         'guardar el control de canvis
         dbtmp.Execute "insert into clients_controlcanvis (codiclient,usuari,campafectat,valoranterior,valoractual) values (" + atrim(clients.Recordset!codi) + ",'" + nomordinador + "','" + rstcontrolcanvis.Fields(i).Name + "','" + atrim(rstcontrolcanvis.Fields(i)) + "','" + atrim(clients.Recordset.Fields(i)) + "')"
        ' MsgBox "insert into clients_controlcanvis (codiclient,usuari,campafectat,valoranterior,valoractual) values (" + atrim(clients.Recordset!codi) + ",'" + nomordinador + "','" + rstcontrolcanvis.Fields(i).Name + "','" + atrim(rstcontrolcanvis.Fields(i)) + "','" + atrim(clients.Recordset.Fields(i)) + "')"
      End If
   Next i
   envios.Refresh
   While Not envios.Recordset.EOF
    Set rstcontrolcanvis = dbcontrolcanvis.OpenRecordset("select * from clients_envios where id=" + atrim(envios.Recordset!ID))
    If Not rstcontrolcanvis.EOF Then
         For i = 0 To envios.Recordset.Fields.Count - 1
           If rstcontrolcanvis.Fields(i) <> envios.Recordset.Fields(i) Then
              'guardar el control de canvis
              dbtmp.Execute "insert into clients_controlcanvis (codiclient,usuari,campafectat,valoranterior,valoractual) values (" + atrim(envios.Recordset!codi) + ",'" + nomordinador + "','[Envio " + atrim(envios.Recordset.AbsolutePosition + 1) + "] " + rstcontrolcanvis.Fields(i).Name + "','" + atrim(rstcontrolcanvis.Fields(i)) + "','" + atrim(envios.Recordset.Fields(i)) + "')"
           End If
        Next i
    End If
    envios.Recordset.MoveNext
   Wend
fi:
   Set rstcontrolcanvis = Nothing
   Set dbcontrolcanvis = Nothing
End Sub
Sub guardar_controlcanvisclients()
   Set rstcontrolcanvis = Nothing
   Set dbcontrolcanvis = Nothing
   If existeix("c:\temp\~canvisclients.mdb") Then Kill "c:\temp\~canvisclients.mdb"
   DBEngine.CreateDatabase "c:\temp\~canvisclients.mdb", dbLangGeneral, DatabaseTypeEnum.dbVersion30
   Set dbcontrolcanvis = OpenDatabase("c:\temp\~canvisclients.mdb")
   clients.Database.Execute "select * into clients IN 'c:\temp\~canvisclients.mdb' from clients where codi=" + atrim(clients.Recordset!codi)
   clients.Database.Execute "select * into clients_envios IN 'c:\temp\~canvisclients.mdb' from clients_envios where codi=" + atrim(clients.Recordset!codi)
   Set rstcontrolcanvis = dbcontrolcanvis.OpenRecordset("select * from clients where codi=" + atrim(clients.Recordset!codi))
End Sub
Private Sub mt2_Click()
If Screen.ActiveControl.Name = "mt2" Then gravar_quantitat
End Sub

Private Sub mtrs_Click()
If Screen.ActiveControl.Name = "mtrs" Then gravar_quantitat
End Sub

Private Sub paperfrontal_Click()
 
End Sub

Private Sub peces_Click()
 On Error GoTo fi
If Screen.ActiveControl.Name = "peces" Then gravar_quantitat
fi:
End Sub

Private Sub pesnetstd_Click()
  carregar_pesosnets
End Sub

Private Sub qproducte_Click()
  If qproducte.ListIndex > -1 Then
    Set rsttmp = dbtmp.OpenRecordset("select * from unitatsxproducte where idproducte='" + atrim(qproducte) + "' and idenvio=" + atrim(cadbl(envios.Tag)))
    If Not rsttmp.EOF Then
      kg.Value = rsttmp!kg
      mtrs.Value = rsttmp!mtrs
      peces.Value = rsttmp!pcs
      unitats.Value = rsttmp!unts
      km.Value = rsttmp!km
      mt2.Value = rsttmp!mt2
      emiler.Value = rsttmp!emiler
        Else
         kg.Value = 0
         mtrs.Value = 0
         peces.Value = 0
         unitats.Value = 0
         km.Value = 0
         mt2.Value = 0
         emiler.Value = 0
    End If
  End If
End Sub

Private Sub reixapesnet_GotFocus()
  If pesnetstd.Value = 1 Then MsgBox "Si modifiques aquests valors canviarant tots els clients amb pes Standard.": Exit Sub
End Sub

Private Sub reixarappels_BeforeUpdate(Cancel As Integer)
   If combogrupclient = "" Then
        Datarappels.Recordset!codiclient = clients.Recordset!codiclient
       Else
         Datarappels.Recordset!codiclient = 0
         Datarappels.Recordset!grupdeclients = combogrupclient
   End If
   
End Sub

Private Sub risc_Change()
If risc = "" And Not buscant Then risc = ",00"
End Sub

Private Sub riscpla_Change()
If riscpla = "" And Not buscant Then riscpla = ",00"
End Sub

Private Sub sortir_Click()
  If clients.Recordset.EditMode > 0 Then MsgBox "Primer has de guardar canvis", vbCritical, "Error": Exit Sub
 Unload Me
End Sub

Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_LostFocus()
 If Not buscant And clients.Recordset.EditMode > 0 Then
   Set rsttmp = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(Text1.Text)))
   If rsttmp.RecordCount > 0 Then MsgBox "Aquest codi ja existeix haurieu de canviar-lo": If areadatos.Enabled Then Text1.SetFocus
 End If
End Sub

Private Sub Text15_LostFocus()
  Text15 = atrim(cadbl(Text15))
End Sub

Private Sub Text16_LostFocus()
  Text16 = atrim(cadbl(Text16))
End Sub

Private Sub Text26_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then triarrepresentant
End Sub

Private Sub Text27_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then triarrepresentant
End Sub

Private Sub Text28_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then triarformapag
End Sub

Private Sub Text29_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then triarformapag
End Sub

Private Sub Text41_Click()
 
' r = "cmd /c "
' If existeix("c:\windows\command\start.exe") Then r = "start "
' r = Shell(r + Chr$(34) + ruta_relativa_docs + "\" + ActiveControl.Text + Chr$(34), vbMinimizedFocus)



  Dim ruta As String
  Dim nomfitxer As String
  
  Text41.SetFocus
  nomfitxer = treure_apostruf(ActiveControl.Text)
  
  numcarpetaclient = Format(clients.Recordset!codi, "000000")
  If cadbl(Mid(nomfitxer, 1, 6)) = 0 Then nomfitxer = numcarpetaclient + " " + Trim(nomfitxer)
  ruta = ruta_relativa_docs + "\" + nomfitxer ' + Chr$(34)
  If existeix(ruta) Then
     obrir_document ruta
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + ruta, vbCritical, "Error"
  End If

End Sub

Private Sub Text42_Click()
 'r = "cmd /c "
 'If existeix("c:\windows\command\start.exe") Then r = "start "
 'r = Shell(r + Chr$(34) + ruta_relativa_docs + "\" + ActiveControl.Text + Chr$(34), vbMinimizedFocus)
 
  Dim ruta As String
  Dim nomfitxer As String
  
  Text42.SetFocus
  nomfitxer = treure_apostruf(ActiveControl.Text)
  
  numcarpetaclient = Format(clients.Recordset!codi, "000000")
  If cadbl(Mid(nomfitxer, 1, 6)) = 0 Then nomfitxer = numcarpetaclient + " " + Trim(nomfitxer)
  ruta = ruta_relativa_docs + "\" + nomfitxer ' + Chr$(34)
  If Not existeix(ruta) Then ruta = ruta + "x"
  If existeix(ruta) Then
     obrir_document ruta
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + ruta, vbCritical, "Error"
  End If
 
 
End Sub


Private Sub Text44_LostFocus()
 If Text44 = "" Then Text44.Text = "0"
End Sub

Private Sub Text45_LostFocus()
Text45 = atrim(cadbl(Text45))
End Sub

Private Sub Timer1_Timer()
  estattaula.Caption = textestattaula(clients.EditMode)
  If estattaula.ForeColor <> QBColor(0) Then
     estattaula.ForeColor = QBColor(0)
    Else: estattaula.ForeColor = QBColor(14)
  End If
End Sub


Sub recorregutregistres()
 Dim objecte As Object
 Dim protegir As Boolean
 protegir = IIf(llegir_ini("general", "protegircamps", fitxerini) = "si", True, False)
 If Not protegir Then escriure_ini "general", "protegircamps", "no", fitxerini
 queryorder = ""
 querywhere = ""
 'On Error Resume Next
 For Each objecte In Me
    If TypeOf objecte Is TextBox Or TypeOf objecte Is ComboBox Then
      If objecte.DataField <> "" Then ' Si Texto es igual "Hola".
        If objecte.Text <> "" Then evaluarcontingut objecte.DataField, objecte.Text, clients.Recordset.Fields(objecte.DataField).Type
     End If
     
    End If
Next

End Sub
Sub colocarbloqueig()
 Dim objecte As Object
 Dim protegir As Boolean
 protegir = IIf(llegir_ini("general", "protegircamps", fitxerini) = "no", False, True)
 If protegir Then escriure_ini "general", "protegircamps", "si", fitxerini
 
 queryorder = ""
 querywhere = ""
 'On Error Resume Next
 For Each objecte In Me
    If TypeOf objecte Is TextBox Then
     If objecte.Tag = "protegits" Then
        If protegir Then
          objecte.Locked = True
         Else
          objecte.BackColor = QBColor(15)
        End If
     End If
     End If
     
   
Next

End Sub

Function evaluarcontingut(Camp As String, valor As String, tipusdato As Byte) As String
  Dim rest As String
  rest = ""
  evaluarcontingut = ""
  If triarordre(Camp, valor) Then Exit Function
  If tipusdato = 10 Then
   If InStr(1, valor, "*") Or InStr(1, valor, "?") Then
      rest = " like '" + valor + "'"
     Else
       If InStr(1, valor, ">") Or InStr(1, valor, "<") Or InStr(1, valor, "=") Then
           rest = "='" + valor + "'"
        Else: rest = "=" + "'" + IIf(valor = " ", "", valor) + "'"
       End If
   End If
  End If
  If tipusdato <> 10 Then
    If InStr(1, valor, ">") Or InStr(1, valor, "<") Or InStr(1, valor, "=") Then
           rest = atrim(cadbl(valor))
        Else: rest = "=" + atrim(cadbl(valor))
    End If
  End If
  rest = Camp + rest
  evaluarcontingut = rest
  
  If querywhere = "" Then
     querywhere = rest
    Else
     querywhere = querywhere + " and " + rest + " "
  End If
End Function

Function triarordre(Camp As String, valorord As String) As Boolean
  Dim ord As String
  triarordre = False
  If InStr(1, valorord, "<<") Then ord = Camp + " " + " ASC"
  If InStr(1, valorord, ">>") Then ord = Camp + " " + " DESC"
  If ord <> "" Then
      triarordre = True
    Else: Exit Function
  End If
  If queryorder = "" Then
     queryorder = ord
   Else: queryorder = queryorder + ", " + ord
  End If
  
End Function
Sub finalitzarbusqueda()
 ratoli "espera"
 recorregutregistres
 If clients.Recordset.EditMode > 0 Then clients.Recordset.CancelUpdate
 buscant = False
 Text1.Enabled = False
 areadatos.Enabled = False
 If queryorder <> "" Then queryorder = " Order By " + queryorder
 If querywhere <> "" Then querywhere = " Where " + querywhere
 clients.RecordSource = "select * from clients " + querywhere + queryorder
 clients.Tag = querywhere + queryorder
 clients.Refresh
 If Not clients.Recordset.EOF Then clients.Recordset.MoveLast
 ratoli "normal"
End Sub

Sub deixartotblanc()
scroll.Visible = False
palets.Visible = False
 For Each objecte In Me
    If TypeOf objecte Is TextBox Then
      If objecte.DataField <> "" Then ' Si Texto es igual "Hola".
        objecte.Text = ""
     End If
    End If
Next



End Sub

Sub carregar_lookups()
 Dim vsql As String
 
 scroll.Visible = False: 'palets.Visible = False
 
 If clients.Recordset.EOF And clients.Recordset.BOF Then Exit Sub
 ' carrego els envios
 'If envios.RecordSource <> clients.RecordSource Then
  envios.RecordSource = " select * from Clients_envios where codi=" + atrim(cadbl(clients.Recordset!codi))
 'End If
 envios.Refresh
If Not envios.Recordset.EOF Then envios.Recordset.MoveLast
 'LOOKUP DE REPRESENTANT
  Set rsttmp = dbtmp.OpenRecordset("select nom from representants where codi=" + atrim(cadbl(clients.Recordset!representant)))
  If Not rsttmp.EOF Then
     Text27 = rsttmp!nom
    Else: Text27 = ""
  End If
  'carrega els rappels del client o del grup de clients
  If combogrupclient = "" Then
        vsql = "codiclient=" + atrim(cadbl(clients.Recordset!codi))
       Else
         vsql = "grupdeclients='" + combogrupclient + "'"
  End If
   
  Datarappels.RecordSource = "select * from clients_rappels where " + vsql
  Datarappels.Refresh
  
  'LOOKUP DE formade pag
  Set rsttmp = dbtmp.OpenRecordset("select descripcio from [formes de pagament] where codi='" + atrim(cadbl(clients.Recordset!formapag)) + "'")
  If Not rsttmp.EOF Then
     Text29 = rsttmp!descripcio
    Else: Text29 = ""
  End If
  
  Set rsttmp = Nothing
End Sub
Sub possarvalordcamps()
On Error Resume Next
  Dim vrecordset As Recordset
 For Each objecte In Me
    If TypeOf objecte Is TextBox Then
      'Set vrecordset = IIf(objecte.Parent.Name = "formenvios", envios.Recordset, clients.Recordset)
      Set vrecordset = IIf(objecte.WhatsThisHelpID = 1, envios.Recordset, clients.Recordset)
      If objecte.DataField <> "" Then
        ' MsgBox objecte.Name
         objecte.MaxLength = vrecordset.Fields(objecte.DataField).Size
      End If
    End If
Next

End Sub

Private Sub unitats_Click()
If Screen.ActiveControl.Name = "unitats" Then gravar_quantitat
End Sub
