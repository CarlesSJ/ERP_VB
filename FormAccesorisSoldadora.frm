VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormAccessorisSoldadora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accessoris Soldadora"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11610
   ControlBox      =   0   'False
   Icon            =   "FormAccesorisSoldadora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bfotos 
      BackColor       =   &H005C31DD&
      Caption         =   "Foto"
      Height          =   420
      Index           =   0
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   78
      Tag             =   "F"
      Top             =   915
      Width           =   795
   End
   Begin VB.CommandButton bfotos 
      Caption         =   "Posició"
      Height          =   420
      Index           =   1
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   77
      Tag             =   "P"
      Top             =   915
      Width           =   795
   End
   Begin VB.CommandButton bfotos 
      Caption         =   "Ubicació"
      Height          =   420
      Index           =   2
      Left            =   9030
      Style           =   1  'Graphical
      TabIndex        =   76
      Tag             =   "U"
      Top             =   915
      Width           =   795
   End
   Begin VB.Frame Framedadesaccessori 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Dades de l'accessori"
      Height          =   5760
      Left            =   2535
      TabIndex        =   50
      Top             =   4680
      Visible         =   0   'False
      Width           =   11220
      Begin VB.CommandButton cbuscaraccessori 
         Height          =   450
         Left            =   3180
         Picture         =   "FormAccesorisSoldadora.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   85
         TabStop         =   0   'False
         ToolTipText     =   "Busqueda de Registres"
         Top             =   2115
         Width           =   450
      End
      Begin VB.TextBox caccessorirelacionat 
         DataField       =   "accessorirelacionat"
         DataSource      =   "datadetall"
         Height          =   315
         Left            =   2010
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   83
         Top             =   2175
         Width           =   1155
      End
      Begin VB.CommandButton bduplicardetall 
         Caption         =   "Duplicar"
         Height          =   705
         Left            =   10185
         Picture         =   "FormAccesorisSoldadora.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   225
         Width           =   930
      End
      Begin VB.Timer Timerrefresc 
         Interval        =   1000
         Left            =   7905
         Top             =   3930
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Documentació  (Només PDF)"
         Height          =   2160
         Left            =   315
         TabIndex        =   71
         Top             =   3465
         Width           =   5085
         Begin VB.CommandButton Command7 
            Height          =   345
            Left            =   45
            Picture         =   "FormAccesorisSoldadora.frx":110B
            Style           =   1  'Graphical
            TabIndex        =   82
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar la documentació sel.leccionada"
            Top             =   225
            Width           =   375
         End
         Begin VB.FileListBox cfitxersdocumentacio 
            BackColor       =   &H00FDDECE&
            Height          =   1845
            Left            =   450
            OLEDropMode     =   1  'Manual
            TabIndex        =   72
            Top             =   195
            Width           =   4575
         End
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cancelar"
         Height          =   480
         Left            =   7935
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   5130
         Width           =   1530
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H006BEBB1&
         Caption         =   "Guardar Canvis"
         Height          =   480
         Left            =   9615
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   5130
         Width           =   1530
      End
      Begin VB.TextBox cobservacions 
         DataField       =   "observacions"
         DataSource      =   "datadetall"
         Height          =   315
         Left            =   225
         MaxLength       =   255
         TabIndex        =   66
         Top             =   3120
         Width           =   9585
      End
      Begin VB.TextBox cubicacio 
         DataField       =   "ubicacio"
         DataSource      =   "datadetall"
         Height          =   315
         Left            =   240
         MaxLength       =   20
         TabIndex        =   64
         Top             =   2190
         Width           =   1680
      End
      Begin VB.TextBox cdatadebaixa 
         DataField       =   "databaixa"
         DataSource      =   "datadetall"
         Height          =   315
         Left            =   3675
         MaxLength       =   10
         TabIndex        =   62
         Top             =   1290
         Width           =   1335
      End
      Begin VB.TextBox cdataarribada 
         DataField       =   "dataarribada"
         DataSource      =   "datadetall"
         Height          =   315
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   60
         Top             =   1305
         Width           =   1335
      End
      Begin VB.TextBox cdatacompra 
         DataField       =   "datacompra"
         DataSource      =   "datadetall"
         Height          =   315
         Left            =   210
         MaxLength       =   10
         TabIndex        =   58
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox crefproveidor 
         DataField       =   "refproveidor"
         DataSource      =   "datadetall"
         Height          =   315
         Left            =   6060
         MaxLength       =   20
         TabIndex        =   56
         Top             =   570
         Width           =   2010
      End
      Begin VB.TextBox crefinplacsa 
         BackColor       =   &H00C0C0C0&
         DataField       =   "refinplacsa"
         DataSource      =   "datadetall"
         Height          =   315
         Left            =   165
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   55
         Top             =   585
         Width           =   1335
      End
      Begin VB.ComboBox comboproveidor 
         DataField       =   "proveidor"
         DataSource      =   "datadetall"
         Height          =   315
         Left            =   1605
         TabIndex        =   53
         Top             =   570
         Width           =   4185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Accessori relacionat"
         Height          =   240
         Index           =   24
         Left            =   2025
         TabIndex        =   84
         Top             =   1920
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Observacions"
         Height          =   240
         Index           =   22
         Left            =   225
         TabIndex        =   65
         Top             =   2880
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicació"
         Height          =   240
         Index           =   21
         Left            =   255
         TabIndex        =   63
         Top             =   1935
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data de baixa"
         Height          =   240
         Index           =   20
         Left            =   3690
         TabIndex        =   61
         Top             =   1035
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data d'arribada"
         Height          =   240
         Index           =   19
         Left            =   1935
         TabIndex        =   59
         Top             =   1050
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data de compra"
         Height          =   240
         Index           =   18
         Left            =   225
         TabIndex        =   57
         Top             =   1065
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. del proveïdor"
         Height          =   240
         Index           =   17
         Left            =   6120
         TabIndex        =   54
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveïdor"
         Height          =   240
         Index           =   16
         Left            =   1860
         TabIndex        =   52
         Top             =   315
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. Inplacsa"
         Height          =   240
         Index           =   15
         Left            =   180
         TabIndex        =   51
         Top             =   330
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   180
      TabIndex        =   7
      Top             =   30
      Width           =   11205
      Begin VB.Timer Timer1 
         Interval        =   800
         Left            =   9465
         Top             =   225
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8460
         Top             =   210
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Data dataaccessoris 
         Caption         =   "Accessoris"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   435
         Left            =   4050
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Accessoris_soldadora"
         Top             =   180
         Width           =   2790
      End
      Begin VB.CommandButton consultar 
         Height          =   450
         Left            =   990
         Picture         =   "FormAccesorisSoldadora.frx":1695
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Busqueda de Registres"
         Top             =   195
         Width           =   450
      End
      Begin VB.CommandButton sortir 
         Height          =   450
         Left            =   10665
         Picture         =   "FormAccesorisSoldadora.frx":1C1F
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Sortir"
         Top             =   195
         Width           =   450
      End
      Begin VB.CommandButton alta 
         Height          =   450
         Left            =   90
         Picture         =   "FormAccesorisSoldadora.frx":21A9
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   195
         Width           =   450
      End
      Begin VB.CommandButton eliminar 
         Height          =   450
         Left            =   1440
         Picture         =   "FormAccesorisSoldadora.frx":2733
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   195
         Width           =   450
      End
      Begin VB.CommandButton gravar 
         Height          =   450
         Left            =   10170
         Picture         =   "FormAccesorisSoldadora.frx":2CBD
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Guardar Registres"
         Top             =   195
         Width           =   450
      End
      Begin VB.CommandButton modificar 
         Height          =   450
         Left            =   540
         Picture         =   "FormAccesorisSoldadora.frx":3247
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Modificació Registres"
         Top             =   195
         Width           =   450
      End
      Begin VB.Label etestat 
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
         ForeColor       =   &H00F1B75F&
         Height          =   330
         Left            =   2040
         TabIndex        =   48
         Top             =   270
         Width           =   2040
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
         TabIndex        =   14
         Top             =   300
         Width           =   105
      End
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "FormAccesorisSoldadora.frx":37D1
      Height          =   2655
      Left            =   210
      OleObjectBlob   =   "FormAccesorisSoldadora.frx":37EA
      TabIndex        =   6
      Top             =   6630
      Width           =   11115
   End
   Begin VB.Frame Framedades 
      Enabled         =   0   'False
      Height          =   5775
      Left            =   285
      TabIndex        =   0
      Top             =   705
      Width           =   11235
      Begin VB.CheckBox checkcontroltraçabilitat 
         Caption         =   "Control de Traçabilitat del consumible  a baixes."
         DataField       =   "control_traçabilitat"
         DataSource      =   "dataaccessoris"
         Height          =   420
         Left            =   4170
         TabIndex        =   86
         Top             =   2775
         Visible         =   0   'False
         Width           =   2580
      End
      Begin VB.TextBox cText 
         DataField       =   "observacions"
         DataSource      =   "dataaccessoris"
         Height          =   570
         Index           =   13
         Left            =   6120
         MaxLength       =   255
         TabIndex        =   80
         Top             =   3405
         Width           =   4980
      End
      Begin VB.CommandButton bcombotipussoldadura 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3645
         TabIndex        =   75
         Top             =   2970
         Width           =   330
      End
      Begin VB.TextBox combotipussoldadura 
         DataField       =   "tipussoldadura"
         DataSource      =   "dataaccessoris"
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   2970
         Width           =   3420
      End
      Begin VB.ListBox llistasoldadures 
         Height          =   2085
         Left            =   10605
         Style           =   1  'Checkbox
         TabIndex        =   73
         Top             =   3780
         Visible         =   0   'False
         Width           =   5700
      End
      Begin VB.CommandButton beliminardetall 
         Height          =   375
         Left            =   45
         Picture         =   "FormAccesorisSoldadora.frx":488F
         Style           =   1  'Graphical
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   4515
         Width           =   375
      End
      Begin VB.CommandButton bafegiraccessori 
         Height          =   375
         Left            =   45
         Picture         =   "FormAccesorisSoldadora.frx":4E19
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   4140
         Width           =   375
      End
      Begin VB.CommandButton botoull 
         Height          =   360
         Left            =   480
         Picture         =   "FormAccesorisSoldadora.frx":53A3
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Modificació Registres"
         Top             =   4350
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EEE4D7&
         Caption         =   "Mides de l'accessori"
         Height          =   1290
         Left            =   240
         TabIndex        =   31
         Top             =   1425
         Width           =   6750
         Begin VB.TextBox cText 
            DataField       =   "alçadatroquel"
            DataSource      =   "dataaccessoris"
            Height          =   315
            Index           =   12
            Left            =   915
            TabIndex        =   42
            Top             =   825
            Width           =   510
         End
         Begin VB.TextBox cText 
            DataField       =   "longitudtroquel"
            DataSource      =   "dataaccessoris"
            Height          =   315
            Index           =   11
            Left            =   2535
            TabIndex        =   41
            Top             =   825
            Width           =   510
         End
         Begin VB.TextBox cText 
            DataField       =   "diametretroquel"
            DataSource      =   "dataaccessoris"
            Height          =   315
            Index           =   10
            Left            =   4185
            TabIndex        =   40
            Top             =   825
            Width           =   510
         End
         Begin VB.TextBox cText 
            DataField       =   "longitudsol"
            DataSource      =   "dataaccessoris"
            Height          =   315
            Index           =   9
            Left            =   6060
            TabIndex        =   38
            Top             =   360
            Width           =   510
         End
         Begin VB.TextBox cText 
            DataField       =   "alçadafuelle"
            DataSource      =   "dataaccessoris"
            Height          =   315
            Index           =   8
            Left            =   4200
            TabIndex        =   36
            Top             =   360
            Width           =   510
         End
         Begin VB.TextBox cText 
            DataField       =   "amplebossa"
            DataSource      =   "dataaccessoris"
            Height          =   315
            Index           =   7
            Left            =   2550
            TabIndex        =   34
            Top             =   360
            Width           =   510
         End
         Begin VB.TextBox cText 
            DataField       =   "amplesol"
            DataSource      =   "dataaccessoris"
            Height          =   315
            Index           =   3
            Left            =   930
            TabIndex        =   32
            Top             =   360
            Width           =   510
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Totes les mides en mil.limetres"
            ForeColor       =   &H00A6A58E&
            Height          =   240
            Index           =   13
            Left            =   2055
            TabIndex        =   46
            Top             =   105
            Width           =   2235
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Alçada Troquel:"
            Height          =   420
            Index           =   12
            Left            =   105
            TabIndex        =   45
            Top             =   735
            Width           =   930
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Long.Troquel:"
            Height          =   240
            Index           =   11
            Left            =   1515
            TabIndex        =   44
            Top             =   885
            Width           =   1290
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Diam.Troquel:"
            Height          =   240
            Index           =   10
            Left            =   3165
            TabIndex        =   43
            Top             =   885
            Width           =   1305
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Long. Soldador:"
            Height          =   240
            Index           =   9
            Left            =   4890
            TabIndex        =   39
            Top             =   420
            Width           =   1305
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Alçada fuelle:"
            Height          =   240
            Index           =   8
            Left            =   3180
            TabIndex        =   37
            Top             =   420
            Width           =   1305
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Ample bossa:"
            Height          =   240
            Index           =   7
            Left            =   1530
            TabIndex        =   35
            Top             =   420
            Width           =   1305
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Ample Sol:"
            Height          =   240
            Index           =   5
            Left            =   90
            TabIndex        =   33
            Top             =   420
            Width           =   930
         End
      End
      Begin VB.Data datadetall 
         Caption         =   "Datadetall"
         Connect         =   "Access"
         DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   7395
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from Accessoris_soldadora_detall"
         Top             =   2925
         Visible         =   0   'False
         Width           =   1500
      End
      Begin MSDBGrid.DBGrid reixadetall 
         Bindings        =   "FormAccesorisSoldadora.frx":592D
         Height          =   1545
         Left            =   435
         OleObjectBlob   =   "FormAccesorisSoldadora.frx":5942
         TabIndex        =   30
         Top             =   4110
         Width           =   9870
      End
      Begin VB.TextBox cText 
         DataField       =   "maquinescompatibles"
         DataSource      =   "dataaccessoris"
         Height          =   315
         Index           =   6
         Left            =   10395
         TabIndex        =   27
         Top             =   4410
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.ComboBox Combomaquines 
         Height          =   315
         ItemData        =   "FormAccesorisSoldadora.frx":69F5
         Left            =   225
         List            =   "FormAccesorisSoldadora.frx":69F7
         TabIndex        =   26
         Top             =   3600
         Width           =   5685
      End
      Begin VB.ListBox llistamaquines 
         Height          =   2085
         Left            =   10710
         Style           =   1  'Checkbox
         TabIndex        =   25
         Top             =   3345
         Visible         =   0   'False
         Width           =   5700
      End
      Begin VB.CommandButton Command4 
         Height          =   345
         Left            =   10695
         Picture         =   "FormAccesorisSoldadora.frx":69F9
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2685
         Width           =   420
      End
      Begin VB.CommandButton Command3 
         Height          =   345
         Left            =   6660
         Picture         =   "FormAccesorisSoldadora.frx":6AE5
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Busqueda de Registres"
         Top             =   420
         Width           =   390
      End
      Begin VB.CommandButton Command2 
         Height          =   345
         Left            =   3555
         Picture         =   "FormAccesorisSoldadora.frx":706F
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Busqueda de Registres"
         Top             =   420
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   10440
         Picture         =   "FormAccesorisSoldadora.frx":75F9
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox cText 
         DataField       =   "subfamilia"
         DataSource      =   "dataaccessoris"
         Height          =   315
         Index           =   5
         Left            =   4035
         TabIndex        =   3
         Top             =   465
         Width           =   2625
      End
      Begin VB.TextBox cText 
         DataField       =   "familia"
         DataSource      =   "dataaccessoris"
         Height          =   315
         Index           =   4
         Left            =   1215
         TabIndex        =   2
         Top             =   465
         Width           =   2325
      End
      Begin VB.TextBox cText 
         DataField       =   "numaccessori"
         DataSource      =   "dataaccessoris"
         Height          =   315
         Index           =   2
         Left            =   195
         TabIndex        =   1
         Top             =   480
         Width           =   795
      End
      Begin VB.TextBox cText 
         DataField       =   "descripcio_llarga"
         DataSource      =   "dataaccessoris"
         Height          =   315
         Index           =   1
         Left            =   2610
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox cText 
         DataField       =   "descripcio_curta"
         DataSource      =   "dataaccessoris"
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1080
         Width           =   2325
      End
      Begin VB.Image fotoaccessori 
         Height          =   2760
         Left            =   7185
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3930
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Observacions"
         Height          =   240
         Index           =   23
         Left            =   6360
         TabIndex        =   79
         Top             =   3180
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipus de soldadura:"
         Height          =   240
         Index           =   14
         Left            =   270
         TabIndex        =   47
         Top             =   2745
         Width           =   2385
      End
      Begin VB.Label Label3 
         Caption         =   "Accessoris a fàbrica"
         Height          =   435
         Left            =   375
         TabIndex        =   29
         Top             =   3900
         Width           =   2760
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Màquines compatibles"
         Height          =   240
         Index           =   6
         Left            =   255
         TabIndex        =   28
         Top             =   3345
         Width           =   2280
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripció llarga"
         Height          =   240
         Index           =   4
         Left            =   3585
         TabIndex        =   19
         Top             =   855
         Width           =   2280
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripció curta (Oficines)"
         Height          =   240
         Index           =   3
         Left            =   495
         TabIndex        =   18
         Top             =   855
         Width           =   2280
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SubFamilia"
         Height          =   240
         Index           =   2
         Left            =   4995
         TabIndex        =   17
         Top             =   225
         Width           =   1110
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Familia"
         Height          =   240
         Index           =   1
         Left            =   1965
         TabIndex        =   16
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Accessori"
         Height          =   240
         Index           =   0
         Left            =   105
         TabIndex        =   15
         Top             =   240
         Width           =   930
      End
      Begin VB.Shape Shape1 
         Height          =   2940
         Left            =   7110
         Top             =   180
         Width           =   4065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Foto del accessori."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A6A58E&
         Height          =   330
         Left            =   8055
         TabIndex        =   23
         Top             =   1215
         Width           =   2430
      End
   End
End
Attribute VB_Name = "FormAccessorisSoldadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim vrutaaccessoris As String

Private Sub alta_Click()
    Dim gran As Long
    dataaccessoris.Recordset.AddNew
    Framedades.Enabled = True
    cText(2) = buscar_elmesgran
    cText(2).SetFocus
End Sub
Function buscar_elmesgran() As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("Select max(numaccessori) as ElGran from accessoris_soldadora")
  buscar_elmesgran = atrim(cadbl(rst!elgran) + 1)
  Set rst = Nothing
End Function

Sub ensenyar_combomaquines()
  
  llistamaquines.Top = Combomaquines.Top - (llistamaquines.Height)
  llistamaquines.Left = Combomaquines.Left
  llistamaquines.Width = Combomaquines.Width
  llistamaquines.ZOrder 0
  carregar_llistamaquines
  llistamaquines.Visible = True
  llistamaquines.SetFocus
End Sub
Sub carregar_llistamaquines()
  Dim rst As Recordset
  Dim vmaquines As String
  Dim vsqlmaquines As String
  Dim vmaquinestipussoldadura As String
  vmaquinestipussoldadura = BuscarMaquinesTipusSoldadura(atrim(dataaccessoris.Recordset!tipussoldadura))
  vsqlmaquines = IIf(vmaquinestipussoldadura <> "", " and codi in (" + atrim(vmaquinestipussoldadura) + ")", "")
  vmaquines = " ," + atrim(cText(6)) + ","
  llistamaquines.Clear
  Combomaquines.Text = ""
  Set rst = dbtmp.OpenRecordset("select * from maquines where maquina='S' and donadadebaixa=null " + vsqlmaquines + " order by codi")
  While Not rst.EOF
     llistamaquines.AddItem rst!descripcio
     llistamaquines.ItemData(llistamaquines.NewIndex) = rst!codi
     If InStr(1, vmaquines, "," + atrim(rst!codi) + ",") > 0 Then
         llistamaquines.Selected(llistamaquines.NewIndex) = True
         Combomaquines.Text = Combomaquines.Text + "[" + rst!descripcio + "] "
     End If
     rst.MoveNext
  Wend
  Set rst = Nothing
End Sub

Private Sub bafegiraccessori_Click()
   Dim v As String
   Dim rst As Recordset
   If dataaccessoris.Recordset.EditMode = dbEditAdd Then MsgBox "No pots afegir referències si estas afegint l'accessori, primer guarda'l.", vbCritical, "Atenció": Exit Sub
   v = UCase(InputBox("Escriu la referencia d'inplacsa del nou accessori.", "Nou accessori"))
   If atrim(v) = "" Then Exit Sub
   Set rst = dbtmp.OpenRecordset("Select * from Accessoris_soldadora_detall where refinplacsa='" + atrim(v) + "'")
   If Not rst.EOF Then MsgBox "Aquest codi d'inplacsa de l'accessori nou ja existeix.", vbCritical, "Error": GoTo fi
   datadetall.Recordset.AddNew
   datadetall.Recordset!id_accessori = dataaccessoris.Recordset!numaccessori
   datadetall.Recordset!refinplacsa = atrim(v)
   datadetall.Recordset.Update
   datadetall.Refresh
   datadetall.Recordset.FindFirst "refinplacsa='" + atrim(v) + "'"
   If Not datadetall.Recordset.NoMatch Then botoull_Click
fi:
   Set rst = Nothing
   
End Sub

Private Sub bcombotipussoldadura_Click()
   If Not llistasoldadures.Visible Then
         ensenyar_combotipussoldadures
          Else: llistasoldadures.Visible = False
   End If
End Sub

Private Sub bduplicardetall_Click()
  Dim vnoucodiinplacsa As String
  Dim rst As Recordset
  Dim vrutaaccessoris2 As String
  Dim i As Integer
  
  vnoucodiinplacsa = UCase(treure_apostruf(InputBox("Escriu la nova referència d'inplacsa pel nou accessori.", "Nou accessori")))
  If atrim(vnoucodiinplacsa) = "" Then Exit Sub
  Set rst = dbtmp.OpenRecordset("Select * from Accessoris_soldadora_detall where refinplacsa='" + vnoucodiinplacsa + "'")
  If Not rst.EOF Then
        MsgBox "Aquesta referencia ja està creada i no es pot utilitzar pel nou accessori.", vbCritical, "Error"
        GoTo fi
  End If
  rst.AddNew
  For i = 0 To rst.Fields.Count - 1
    rst.Fields(i) = datadetall.Recordset.Fields(i)
  Next i
  rst!refinplacsa = vnoucodiinplacsa
  rst.Update
  If MsgBox("Vols duplicar la documentació annexada a aquest accessori?", vbExclamation + vbDefaultButton2 + vbYesNo, "Documentació") = vbYes Then
     vrutaaccessoris2 = vrutaaccessoris + atrim(vnoucodiinplacsa)
     If Not existeix(vrutaaccessoris2) Then MkDir vrutaaccessoris2
     For i = 0 To cfitxersdocumentacio.ListCount - 1
       Copiar_Fitxer cfitxersdocumentacio.path + "\" + cfitxersdocumentacio.List(i), vrutaaccessoris2 + "\" + cfitxersdocumentacio.List(i)
     Next i
  End If
  datadetall.Refresh
  datadetall.Recordset.FindFirst "refinplacsa='" + vnoucodiinplacsa + "'"
fi:
  Set rst = Nothing
  
End Sub

Private Sub beliminardetall_Click()
 Dim i As Integer
 If MsgBox("Segur que vols eliminar aquest accessori." + vbNewLine + " SI JA L'HAS UTILITZAT POSSA DATA DE BAIXA I NO L'ELIMINIS.", vbCritical + vbYesNo + vbDefaultButton2, "ATENCIÓ") = vbNo Then Exit Sub
 If Not datadetall.Recordset.EOF Then
     For i = 0 To cfitxersdocumentacio.ListCount - 1
        Kill cfitxersdocumentacio.path + "\" + cfitxersdocumentacio.List(i)
     Next i
     RmDir cfitxersdocumentacio.path
     datadetall.Recordset.Delete: datadetall.Refresh
 End If
End Sub

Private Sub bfotos_Click(Index As Integer)
   bfotos(0).BackColor = &H8000000F: bfotos(1).BackColor = &H8000000F: bfotos(2).BackColor = &H8000000F
   bfotos(Index).BackColor = &H5C31DD
   carregar_fotoactiva
End Sub
Function fotoactiva() As String
  Dim i As Byte
  For i = 0 To 2
   If bfotos(i).BackColor = &H5C31DD Then fotoactiva = bfotos(i).Tag
  Next i
End Function

Private Sub botoull_Click()
  If datadetall.Recordset.EOF Then Exit Sub
   Framedadesaccessori.Left = Framedades.Left
   Framedadesaccessori.Top = Framedades.Top
   Framedadesaccessori.ZOrder 0
   Framedadesaccessori.Visible = True
   sortir.Enabled = False
   refrescar_documentacio
   datadetall.Recordset.Edit
End Sub

Private Sub cbuscaraccessori_Click()
  Dim rst As Recordset
  Dim vsql As String
  Unload formseleccio
  Load formseleccio
  vsql = GenerarFiltreSQL("maquinescompatibles", FormAccessorisSoldadora.dataaccessoris.Recordset!maquinescompatibles)
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = dataaccessoris.DatabaseName
  formseleccio.Data1.RecordSource = "SELECT Accessoris_soldadora_detall.refinplacsa, Accessoris_soldadora.descripcio_curta, Accessoris_soldadora_detall.databaixa FROM Accessoris_soldadora LEFT JOIN Accessoris_soldadora_detall ON Accessoris_soldadora.numaccessori = Accessoris_soldadora_detall.id_accessori WHERE (((Accessoris_soldadora_detall.refinplacsa) Is Not Null) AND ((Accessoris_soldadora_detall.databaixa) Is Null) and " + vsql + ");"
'  Clipboard.Clear
'  Clipboard.SetText "SELECT Accessoris_soldadora_detall.refinplacsa, Accessoris_soldadora.descripcio_curta, Accessoris_soldadora_detall.databaixa FROM Accessoris_soldadora LEFT JOIN Accessoris_soldadora_detall ON Accessoris_soldadora.numaccessori = Accessoris_soldadora_detall.id_accessori WHERE (((Accessoris_soldadora_detall.refinplacsa) Is Not Null) AND ((Accessoris_soldadora_detall.databaixa) Is Null) and maquinescompatibles in (" + FormAccessorisSoldadora.dataaccessoris.Recordset!maquinescompatibles + "));"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 2000
  formseleccio.DBGrid2.Columns(1).Width = 3000
  formseleccio.Width = 7000
  formseleccio.Show 1
  If seleccioret = 1 Then
      caccessorirelacionat = UCase(atrim(formseleccio.Data1.Recordset!refinplacsa))
  End If
  Unload formseleccio
fi:
  Set rst = Nothing
End Sub
Public Function GenerarFiltreSQL(ByVal Camp As String, ByVal EntradaUsuari As String) As String
    Dim pos As Integer
    Dim seguentPos As Integer
    Dim valor As String
    Dim sql As String
    
    ' Afegim una coma al final per facilitar el bucle
    EntradaUsuari = EntradaUsuari & ","
    pos = 1
    
    ' Busquem cada coma per separar els números
    Do While InStr(pos, EntradaUsuari, ",") > 0
        seguentPos = InStr(pos, EntradaUsuari, ",")
        valor = Trim(Mid(EntradaUsuari, pos, seguentPos - pos))
        
        If valor <> "" Then
            If sql <> "" Then sql = sql & " OR "
            ' Sintaxi LIKE per a Access (Jet)
            sql = sql & "(',' & " & Camp & " & ',' LIKE '*," & valor & ",*')"
        End If
        
        pos = seguentPos + 1
    Loop
    
    GenerarFiltreSQL = "(" & sql & ")"
End Function
Private Sub cfitxersdocumentacio_DblClick()
   obrir_document cfitxersdocumentacio.path + "\" + cfitxersdocumentacio.FileName
End Sub

Private Sub cfitxersdocumentacio_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim vnomfitxer As String
   vnomfitxer = UCase(Data.Files(1))
   If InStr(1, vnomfitxer, ".PDF") > 0 Then
         vnomfitxer = Mid(vnomfitxer, Len(rutadelfitxer(vnomfitxer) + "\"))
         v = UCase(InputBox("Escriu el nom que vols que tingui el fitxer PDF", "Canvi de nom", Mid(vnomfitxer, 1, InStr(1, vnomfitxer, ".PDF") - 1)))
         If v = "" Then Exit Sub
         If InStr(1, v, ".PDF") = 0 Then v = v + ".PDF"
         Copiar_Fitxer Data.Files(1), cfitxersdocumentacio.path + "\" + v
         'Command5_Click
        Else
          MsgBox "El fitxer ha de ser .PDF", vbCritical, "ERROR"
   End If
End Sub

Private Sub Combomaquines_DropDown()
  ensenyar_combomaquines
End Sub

Private Sub Combomaquines_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub Combomaquines_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub comboproveidor_DropDown()
  triar_proveidor
End Sub
Sub triar_proveidor()
  Load formseleccio
  formseleccio.sortirs.Tag = "filtre"
  'formseleccio.Data1.DatabaseName = cami
  Set formseleccio.Data1.Recordset = dbtmp.OpenRecordset("select * from proveidors where databaixa=null")
  'formseleccio.Data1.RecordSource = "select * from proveidors"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   comboproveidor = atrim(cadbl(formseleccio.Data1.Recordset!codi)) + " - " + atrim(formseleccio.Data1.Recordset!nom)
   Unload formseleccio
  End If
  Unload formseleccio
End Sub


Private Sub comboproveidor_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub comboproveidor_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub combotipussoldadura_DropDown()
 

End Sub
Sub ensenyar_combotipussoldadures()
  
  llistasoldadures.Top = combotipussoldadura.Top - (llistasoldadures.Height)
  llistasoldadures.Left = combotipussoldadura.Left
  llistasoldadures.Width = combotipussoldadura.Width
  llistasoldadures.ZOrder 0
  carregar_llistatipussoldadures
  llistasoldadures.Visible = True
  llistasoldadures.SetFocus
End Sub
Sub carregar_llistatipussoldadures()
  Dim rst As Recordset
  Dim vtipus As String
  vtipus = combotipussoldadura.Text
  llistasoldadures.Clear
 ' combotipussoldadura.Text = ""
  Set rst = dbtmp.OpenRecordset("select * from tipussoldadura order by codi")
  While Not rst.EOF
     llistasoldadures.AddItem rst!codi + " - " + rst!descripcio
     'llistasoldadures.ItemData(llistasol dadures.NewIndex) =  rst!codi
     If InStr(1, vtipus, "[" + atrim(rst!codi) + "]") > 0 Then
         llistasoldadures.Selected(llistasoldadures.NewIndex) = True
        ' combotipussoldadura.Text = combotipussoldadura.Text + "[" + rst!codi + "] "
     End If
     rst.MoveNext
  Wend
  Set rst = Nothing
End Sub








Private Sub Command1_Click()
  Dim vnomfitxer As String
  Dim vnomfitxerdesti As String
  Dim vnum As String
  Dim vfoto As String
  vfoto = fotoactiva
  vnum = atrim(cadbl(cText(2)))
  vnomfitxer = SeleccionarArxiuJPG
  If existeix(vnomfitxer) Then
      dataaccessoris.Recordset.Update
      vnomfitxerdesti = vrutaaccessoris + "FotosAccessoris\" + vfoto + "_" + vnum + ".jpg"
      If existeix(vnomfitxerdesti) Then Kill vnomfitxerdesti
      If Not existeix(vrutaaccessoris) Then MkDir vrutaaccessoris + "FotosAccessoris\"
      Copiar_Fitxer vnomfitxer, vnomfitxerdesti
      dataaccessoris.Recordset.FindFirst "numaccessori=" + vnum
  End If
End Sub
Public Function SeleccionarArxiuJPG() As String
    With CommonDialog1
        .Filter = "Archivos JPG (*.jpg)|*.jpg|Todos los archivos (*.*)|*.*"
        .FilterIndex = 1 ' Selecciona per defecte el filtre de JPG
        .DialogTitle = "Selecciona un fitxer JPG"
        .flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
        
        ' Mostra el diàleg d'obrir fitxer
        On Error Resume Next ' Per gestionar si l'usuari cancel·la el diàleg
        .ShowOpen
        On Error GoTo 0
        
        ' Retorna el nom de fitxer seleccionat
        SeleccionarArxiuJPG = .FileName
    End With
End Function

Private Sub Command2_Click()
  Dim rst As Recordset
  Unload formseleccio
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = dataaccessoris.DatabaseName
  formseleccio.Data1.RecordSource = "select ucase(familia) as Familia_ from accessoris_soldadora  group by familia order by familia "
  formseleccio.refrescar
  formseleccio.alta.Visible = True
  formseleccio.DBGrid2.Columns(0).Width = 5000
  formseleccio.Width = 7000
  If formseleccio.Data1.Recordset.EOF Then GoTo crearnova
  formseleccio.Show 1
  If seleccioret = 1 Then
      cText(4) = UCase(atrim(formseleccio.Data1.Recordset!familia_))
      cText(5) = ""
  End If
  If seleccioret = 2 Then
crearnova:
    v = atrim(InputBox("Escriu el nom de la familia NOVA", "Nova familia"))
    If v = "" Then Exit Sub
    v = UCase(treure_apostruf(v))
    Set rst = dbtmp.OpenRecordset("select distinct ucase(familia) as Familia_ from accessoris_soldadora")
    rst.FindFirst "Familia_='" + atrim(v) + "'"
    If Not rst.NoMatch Then MsgBox "Aquesta familia ja existeix.", vbCritical, "Error": GoTo fi
    cText(4) = v
    cText(5) = ""
   ' dataaccessoris.Recordset.Update
    'modificar_Click
    
  End If
  Unload formseleccio
fi:
  Set rst = Nothing
End Sub

Private Sub Command3_Click()
 Dim rst As Recordset
 If cText(4) = "" Then MsgBox ("Primer escull la familia."): GoTo fi
 Unload formseleccio
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = dataaccessoris.DatabaseName
  formseleccio.Data1.RecordSource = "select ucase(subfamilia) as SubFamilia_ from accessoris_soldadora where familia='" + cText(4) + "' group by subfamilia order by subfamilia "
  formseleccio.alta.Visible = True
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 5000
  formseleccio.Width = 7000
  If formseleccio.Data1.Recordset.EOF Then GoTo crearnova
  formseleccio.Show 1
  If seleccioret = 1 Then
      cText(5) = UCase(formseleccio.Data1.Recordset!subfamilia_)
  End If
  If seleccioret = 2 Then
crearnova:
    v = atrim(InputBox("Escriu el nom de la Subfamilia NOVA", "Nova Subfamilia"))
    If v = "" Then Exit Sub
    v = UCase(treure_apostruf(v))
    Set rst = dbtmp.OpenRecordset("select ucase(subfamilia) as SubFamilia_ from accessoris_soldadora where familia='" + cText(4) + "'")
    rst.FindFirst "SubFamilia_='" + atrim(v) + "'"
    If Not rst.NoMatch Then MsgBox "Aquesta SUBfamilia ja existeix.", vbCritical, "Error": GoTo fi
    cText(5) = v
  End If
  Unload formseleccio
fi:
  Set rst = Nothing
End Sub

Private Sub Command4_Click()
  Dim vnomfitxerdesti As String
  Dim vnum As String
  If MsgBox("Vols borrar aquesta foto?", vbCritical + vbYesNo, "Atenció") = vbYes Then
      dataaccessoris.Recordset.Update
      vnum = atrim(cadbl(cText(2)))
      vnomfitxerdesti = vrutaaccessoris + "FotosAccessoris\" + fotoactiva + "_" + vnum + ".jpg"
      If existeix(vnomfitxerdesti) Then Kill vnomfitxerdesti
      dataaccessoris.Recordset.FindFirst "numaccessori=" + vnum
  End If
End Sub

Private Sub Command5_Click()
  Framedadesaccessori.Visible = False
  If datadetall.Recordset.EditMode > 0 Then datadetall.Recordset.Update
  sortir.Enabled = True
End Sub

Private Sub Command6_Click()
  Framedadesaccessori.Visible = False
  If datadetall.Recordset.EditMode > 0 Then datadetall.Recordset.CancelUpdate
  sortir.Enabled = True
End Sub

Private Sub Command7_Click()
   Dim vnomfitxer As String
   If cfitxersdocumentacio.ListIndex < 0 Then MsgBox "Primer has d'escullir un arxiu per eliminar.", vbCritical, "Atenció": Exit Sub
   vnomfitxer = cfitxersdocumentacio.path + "\" + cfitxersdocumentacio.List(cfitxersdocumentacio.ListIndex)
   If existeix(vnomfitxer) Then
       If MsgBox("Segur que vols eliminar l'arxiu " + cfitxersdocumentacio.List(cfitxersdocumentacio.ListIndex) + "?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
           On Error GoTo errorborrar
           Kill vnomfitxer
           'Command5_Click
       End If
   End If
   Exit Sub
errorborrar:
   MsgBox "Hi ha hagut un error al eliminar l'arxiu, comprova que no el tinguis obert.", vbCritical, "Error"
End Sub

Private Sub consultar_Click()
   Dim b As String
   b = InputBox("Entra la Descripcio o NumAccessori a buscar " + Chr(10) + " No escriguis res per treure els filtres", "Busqueda")
   b = treure_apostruf(b)
   If cadbl(b) > 0 Then
     dataaccessoris.RecordSource = "select * from accessoris_soldadora where numaccessori=" + atrim(cadbl(b))
     dataaccessoris.Refresh
     b = ""
      Else
       If b <> "" Then
        dataaccessoris.RecordSource = "select * from accessoris_soldadora where descripcio_curta like '*" + b + "*' or numaccessori = " + atrim(cadbl(b))
        dataaccessoris.Refresh
          Else
             dataaccessoris.RecordSource = "select * from accessoris_soldadora order by numaccessori"
             dataaccessoris.Refresh
       End If
   End If
End Sub

Function jaexisteix(vnum As Long) As Boolean
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from accessoris_soldadora where numaccessori=" + atrim(vnum))
   If Not rst.EOF Then jaexisteix = True
   Set rst = Nothing
End Function

Private Sub crefinplacsa_DblClick()
 Dim v As String
 Dim rst As Recordset
 v = InputBox("Entra el numero de Referència d'inplacsa que vols posar.", "Canvi referencia", crefinplacsa)
 If v = "" Then Exit Sub
 Set rst = dbtmp.OpenRecordset("Select * from Accessoris_soldadora_detall where refinplacsa='" + atrim(v) + "'")
 If Not rst.EOF Then MsgBox "Aquest codi d'inplacsa de l'accessori nou ja existeix.", vbCritical, "Error": GoTo fi
 crefinplacsa = v
fi:
 Set rst = Nothing
End Sub

Private Sub cText_GotFocus(Index As Integer)
   cText(Index).SelStart = 0
   cText(Index).SelLength = Len(cText(Index))
End Sub

Private Sub cText_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 4 Or Index = 5 Then KeyCode = 0
End Sub

Private Sub cText_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 4 Or Index = 5 Then KeyAscii = 0
End Sub

Private Sub cText_LostFocus(Index As Integer)
  cText(Index) = UCase(cText(Index))
  If Index = 2 Then
      If cadbl(cText(2)) <> cadbl(dataaccessoris.Recordset!numaccessori) Then
           If jaexisteix(cadbl(cText(2))) Then MsgBox "Aquest numero d'accessori ja existeix.", vbCritical, "Error": cText(2).SetFocus: cText(2) = ""
      End If
  End If
End Sub

Private Sub dataaccessoris_Reposition()
   Dim vnomfitxer As String
   Framedades.Enabled = False
   
   If Screen.ActiveControl.Name <> "bfotos" Then bfotos_Click 0
   carregar_llistamaquines
'   combotipussoldadura = atrim(dataaccessoris.Recordset!tipussoldadura)
   If Not dataaccessoris.Recordset.EOF Then
        datadetall.RecordSource = " select * from Accessoris_soldadora_detall where id_accessori=" + atrim(dataaccessoris.Recordset!numaccessori) + " order by databaixa desc"
        datadetall.Refresh
        carregar_fotoactiva
   End If
   Framedadesaccessori.Visible = False
   sortir.Enabled = True
End Sub
Sub carregar_fotoactiva()
  Dim vfoto As String
  Dim vnomfitxer As String
  vfoto = fotoactiva
  vnomfitxer = vrutaaccessoris + "FotosAccessoris\" + vfoto + "_" + atrim(cadbl(dataaccessoris.Recordset!numaccessori)) + ".jpg"
  fotoaccessori.Picture = LoadPicture("")
  fotoaccessori.Tag = ""
  If existeix(vnomfitxer) Then
     On Error Resume Next
     fotoaccessori.Picture = LoadPicture(vnomfitxer)
     fotoaccessori.Tag = vnomfitxer
  End If
  vnomfitxer = vrutaaccessoris + "FotosAccessoris\F_" + atrim(cadbl(dataaccessoris.Recordset!numaccessori)) + ".jpg"
  If existeix(vnomfitxer) Then bfotos(0).FontUnderline = True Else bfotos(0).FontUnderline = False
  vnomfitxer = vrutaaccessoris + "FotosAccessoris\P_" + atrim(cadbl(dataaccessoris.Recordset!numaccessori)) + ".jpg"
  If existeix(vnomfitxer) Then bfotos(1).FontUnderline = True Else bfotos(1).FontUnderline = False
  vnomfitxer = vrutaaccessoris + "FotosAccessoris\U_" + atrim(cadbl(dataaccessoris.Recordset!numaccessori)) + ".jpg"
  If existeix(vnomfitxer) Then bfotos(2).FontUnderline = True Else bfotos(2).FontUnderline = False
End Sub
Function BuscarMaquinesTipusSoldadura(vtipussoldadura) As String
  Dim rst As Recordset
  Dim vtipus As String
  vtipus = substituir(" " + combotipussoldadura, "[", "'")
  vtipus = substituir(" " + vtipus, "]", "',") + "''"
 ' If InStr(1, vtipussoldadura, " - ") > 0 Then
   Set rst = dbtmp.OpenRecordset("select * from tipussoldadura where codi In (" + vtipus + ")")
   While Not rst.EOF
      If atrim(rst!maquinescompatibles) <> "" Then BuscarMaquinesTipusSoldadura = BuscarMaquinesTipusSoldadura + IIf(BuscarMaquinesTipusSoldadura <> "", ",", "") + atrim(rst!maquinescompatibles)
      rst.MoveNext
   Wend
 ' End If
  Set rst = Nothing
End Function

Private Sub datadetall_Reposition()
  refrescar_documentacio
End Sub
Sub refrescar_documentacio()
  Dim vrutaaccessoris2 As String
  If Not datadetall.Recordset.EOF Then
   vrutaaccessoris2 = vrutaaccessoris + atrim(datadetall.Recordset!refinplacsa)
   If Not existeix(vrutaaccessoris2) Then MkDir vrutaaccessoris2
   cfitxersdocumentacio.path = vrutaaccessoris2
'   Me.Caption = cfitxersdocumentacio.path
   cfitxersdocumentacio.Refresh
  End If
End Sub
Private Sub eliminar_Click()
  Dim vcodi As String
  Dim vnomfitxer As String
  vcodi = cadbl(dataaccessoris.Recordset!numaccessori)
  If datadetall.Recordset.EOF And datadetall.Recordset.BOF Then
        If MsgBox("Segur que vols borrar aquest accessori?", vbCritical + vbYesNo + vbDerfaultButton2, "Atenció") = vbYes Then
           If UCase(InputBox("Escriu la paraula [ELIMINAR] accessori " + atrim(vcodi) + "-" + atrim(dataaccessoris.Recordset!descripcio_curta) + vbNewLine + " per fer efectiu l'eliminació", "Control de seguretat")) = "ELIMINAR" Then
               dbtmp.Execute ("delete * from accessoris_solDadora where numaccessori=" + atrim(cadbl(vcodi)))
               vnomfitxer = rutadelfitxer(cami) + "FotosAccessoris\F_" + vcodi + ".jpg"
               If existeix(vnomfitxer) Then Kill vnomfitxer
               vnomfitxer = rutadelfitxer(cami) + "FotosAccessoris\P_" + vcodi + ".jpg"
               If existeix(vnomfitxer) Then Kill vnomfitxer
               vnomfitxer = rutadelfitxer(cami) + "FotosAccessoris\U_" + vcodi + ".jpg"
               If existeix(vnomfitxer) Then Kill vnomfitxer
               dataaccessoris.Refresh
           End If
        End If
         Else: MsgBox "Per poder eliminar un accessori primer s'han d'eliminar totes les linies d'accesori.", vbCritical, "Error"
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then gravar_Click
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
  vrutaaccessoris = "\\ord_copies\DadesProduccio\Arxius Produccio\DadesGenerals\DocumentacióAccessoris\"
  dataaccessoris.DatabaseName = rutadelfitxer(cami) + "comandes.mdb"
  dataaccessoris.RecordSource = "select * from accessoris_soldadora order by numaccessori"
  
End Sub

Private Sub fotoaccessori_DblClick()
  obrir_document fotoaccessori.Tag
End Sub

Private Sub gravar_Click()
  Dim vnum As String
  If dataaccessoris.Recordset.EditMode > 0 Then
      vnum = cadbl(cText(2))
      dataaccessoris.Recordset.Update
      dataaccessoris.Recordset.FindFirst "numaccessori=" + atrim(vnum)
  End If
  sortir.Enabled = True
End Sub

Private Sub llistamaquines_Click()
  Dim vmaquines As String
  vmaquines = ""
  For i = 0 To llistamaquines.ListCount - 1
     If llistamaquines.Selected(i) = True Then vmaquines = vmaquines + IIf(vmaquines = "", "", ",") + atrim(llistamaquines.ItemData(i))
  Next i
  cText(6) = vmaquines
End Sub

Private Sub llistamaquines_LostFocus()
  llistamaquines.Visible = False
  carregar_llistamaquines
End Sub

Private Sub llistasoldadures_Click()
  Dim vtipus As String
  vtipus = ""
  For i = 0 To llistasoldadures.ListCount - 1
     If llistasoldadures.Selected(i) = True Then vtipus = vtipus + "[" + atrim(Mid(llistasoldadures.List(i), 1, InStr(1, llistasoldadures.List(i), " -"))) + "] "
  Next i
  
  combotipussoldadura = vtipus
End Sub

Private Sub llistasoldadures_LostFocus()
  If Screen.ActiveControl.Name = "bcombotipussoldadura" Then Exit Sub
  llistasoldadures.Visible = False
  carregar_llistatipussoldadures
End Sub

Private Sub modificar_Click()
   If Not dataaccessoris.Recordset.EOF Then
     dataaccessoris.Recordset.Edit
     Framedades.Enabled = True
     cText(2).SetFocus
   End If
End Sub

Private Sub reixadetall_LostFocus()
  If ActiveControl.Name <> "botoull" Then botoull.Visible = False
End Sub

Private Sub reixadetall_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Dim vtop As Double
  If reixadetall.row < 0 Then botoull.Visible = False: Exit Sub
  botoull.Top = reixadetall.RowTop(reixadetall.row) + reixadetall.Top
  botoull.Left = reixadetall.Left + 20
  'botoull.Left = reixadetall.Columns(reixadetall.col).Left + reixadetall.Left
  botoull.Visible = True
End Sub

Private Sub sortir_Click()
 Unload Me
End Sub

Private Sub Timer1_Timer()
  If dataaccessoris.Recordset.EditMode > 0 Then
       etestat = "Editant..."
         Else: etestat = ""
  End If
End Sub

Private Sub Timerrefresc_Timer()
  If LCase(Screen.ActiveControl.Name) <> "cfitxersdocumentacio" Then refrescar_documentacio
 ' Timerrefresc.Enabled = False
End Sub
