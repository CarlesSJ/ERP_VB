VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{3D20F47F-E818-4A03-AD52-45B708ACCF23}#1.0#0"; "FoxitReaderOCX.ocx"
Begin VB.Form form1 
   Caption         =   "Baixes de Laminadora"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11895
   ClipControls    =   0   'False
   Icon            =   "baixes laminadores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framebobentrada 
      Caption         =   "Bobines Entrada"
      Height          =   3885
      Left            =   6795
      TabIndex        =   68
      Top             =   3915
      Visible         =   0   'False
      Width           =   3525
      Begin VB.CommandButton Command26 
         Height          =   330
         Left            =   3045
         Picture         =   "baixes laminadores.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Ubicació d'una bobina a magatzem."
         Top             =   3495
         Width           =   420
      End
      Begin VB.CommandButton botoensenyarpacking 
         Height          =   480
         Left            =   90
         Picture         =   "baixes laminadores.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Sel.lecciona la bobina del Packinglist"
         Top             =   3315
         Width           =   645
      End
      Begin VB.CommandButton Command20 
         Height          =   480
         Left            =   720
         Picture         =   "baixes laminadores.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Afegir manualment el Palet/Bobina d'entrada"
         Top             =   3315
         Width           =   645
      End
      Begin VB.CommandButton eliminarbobentrada 
         Height          =   480
         Left            =   1920
         Picture         =   "baixes laminadores.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Eliminar bobina d'entrada"
         Top             =   3315
         Width           =   645
      End
      Begin VB.CheckBox ensenyartoteslesbobines 
         Caption         =   "Totes"
         Height          =   195
         Left            =   2610
         TabIndex        =   83
         Top             =   3300
         Width           =   795
      End
      Begin MSDBGrid.DBGrid bobentrada 
         Bindings        =   "baixes laminadores.frx":1634
         Height          =   3075
         Left            =   90
         OleObjectBlob   =   "baixes laminadores.frx":1649
         TabIndex        =   70
         Top             =   225
         Width           =   3330
      End
   End
   Begin VB.Frame framepantones 
      Caption         =   "Adhesius"
      Height          =   3930
      Left            =   6885
      TabIndex        =   29
      Top             =   3915
      Visible         =   0   'False
      Width           =   3450
      Begin VB.CommandButton Command25 
         Height          =   375
         Left            =   2850
         Picture         =   "baixes laminadores.frx":21DF
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Informació de la resina i enduridor."
         Top             =   1305
         Width           =   525
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg4"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   3
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   41
         Tag             =   "1"
         Top             =   2070
         Width           =   550
      End
      Begin VB.CommandButton Command24 
         Height          =   315
         Left            =   375
         Picture         =   "baixes laminadores.frx":2769
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Guardar el lot d'enduridor com a predeterminat."
         Top             =   1290
         Width           =   300
      End
      Begin VB.CommandButton Command23 
         Height          =   315
         Left            =   405
         Picture         =   "baixes laminadores.frx":27C4
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Guardar lot de cola com a predeterminat."
         Top             =   645
         Width           =   300
      End
      Begin VB.CommandButton Command22 
         Height          =   285
         Left            =   1770
         Picture         =   "baixes laminadores.frx":281F
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Escanejar el codi de Lot de l'Enduridor"
         Top             =   1260
         Width           =   555
      End
      Begin VB.CommandButton Command19 
         Height          =   285
         Left            =   1770
         Picture         =   "baixes laminadores.frx":30E9
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Escanejar el codi de Lot de la Resina"
         Top             =   660
         Width           =   555
      End
      Begin VB.CommandButton Command18 
         Height          =   270
         Left            =   2595
         Picture         =   "baixes laminadores.frx":39B3
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Canviar de cola i enduridor"
         Top             =   360
         Width           =   270
      End
      Begin VB.TextBox Text1 
         DataField       =   "observacions"
         DataSource      =   "imppantones"
         Height          =   555
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   84
         Top             =   3330
         Width           =   3210
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg5"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   4
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   44
         Tag             =   "1"
         Top             =   2175
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg3"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   2
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   38
         Tag             =   "1"
         Top             =   1770
         Width           =   550
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg2"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   1
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   35
         Tag             =   "1"
         Top             =   990
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot2"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   1
         Left            =   675
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   34
         Tag             =   "888"
         Top             =   1275
         Width           =   1100
      End
      Begin VB.TextBox pantone 
         BackColor       =   &H00EAD9CE&
         DataField       =   "pantone2"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   1
         Left            =   255
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   33
         Tag             =   "888"
         Text            =   "LIOFOL 6020"
         Top             =   990
         Width           =   2580
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg1"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   0
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   32
         Tag             =   "1"
         Top             =   375
         Width           =   550
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot1"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   0
         Left            =   675
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   31
         Tag             =   "888"
         Top             =   660
         Width           =   1100
      End
      Begin VB.TextBox pantone 
         BackColor       =   &H00EEE4D7&
         DataField       =   "pantone1"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   0
         Left            =   255
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   30
         Tag             =   "888"
         Text            =   "LIOFOL 7724"
         Top             =   375
         Width           =   2370
      End
      Begin VB.CommandButton Command16 
         Caption         =   "+Re +En"
         Height          =   570
         Left            =   2355
         TabIndex        =   87
         Top             =   1785
         Width           =   450
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone3"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   2
         Left            =   255
         MaxLength       =   40
         TabIndex        =   36
         Tag             =   "888"
         Top             =   930
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot3"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   2
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   37
         Tag             =   "888"
         Top             =   930
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone4"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   3
         Left            =   270
         MaxLength       =   40
         TabIndex        =   39
         Tag             =   "888"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot4"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   3
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   40
         Tag             =   "888"
         Top             =   1200
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
         TabIndex        =   42
         Tag             =   "888"
         Top             =   1470
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot5"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   4
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   43
         Tag             =   "888"
         Top             =   1470
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
         TabIndex        =   45
         Tag             =   "888"
         Top             =   1755
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot6"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   5
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   46
         Tag             =   "888"
         Top             =   1755
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg6"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   5
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   47
         Tag             =   "1"
         Top             =   1755
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone7"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   6
         Left            =   255
         MaxLength       =   40
         TabIndex        =   48
         Tag             =   "888"
         Top             =   2025
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot7"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   6
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   49
         Tag             =   "888"
         Top             =   2025
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg7"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   6
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   50
         Tag             =   "1"
         Top             =   2025
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone8"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   7
         Left            =   255
         MaxLength       =   40
         TabIndex        =   51
         Tag             =   "888"
         Top             =   2310
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot8"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   7
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   52
         Tag             =   "888"
         Top             =   2310
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg8"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   7
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   53
         Tag             =   "1"
         Top             =   2310
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone10"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   9
         Left            =   255
         MaxLength       =   40
         TabIndex        =   57
         Tag             =   "888"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot10"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   9
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   58
         Tag             =   "888"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg10"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   9
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   59
         Tag             =   "1"
         Top             =   2520
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox pantone 
         DataField       =   "pantone9"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   8
         Left            =   255
         MaxLength       =   40
         TabIndex        =   54
         Tag             =   "888"
         Top             =   2565
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox compantone 
         DataField       =   "lot9"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   8
         Left            =   1755
         MaxLength       =   12
         TabIndex        =   55
         Tag             =   "888"
         Top             =   2565
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg9"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   8
         Left            =   2850
         MaxLength       =   8
         TabIndex        =   56
         Tag             =   "1"
         Top             =   2565
         Visible         =   0   'False
         Width           =   550
      End
      Begin MSDBGrid.DBGrid dblots 
         Bindings        =   "baixes laminadores.frx":3F3D
         Height          =   2805
         Left            =   780
         OleObjectBlob   =   "baixes laminadores.frx":3F4C
         TabIndex        =   77
         Top             =   180
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Re          En"
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
         TabIndex        =   61
         Top             =   300
         Width           =   225
      End
      Begin VB.Label Label5 
         Caption         =   "Comentari lots adhesiu."
         Height          =   255
         Left            =   105
         TabIndex        =   98
         Top             =   3030
         Width           =   3240
      End
      Begin VB.Label Label3 
         Caption         =   "Observacions"
         Height          =   210
         Left            =   495
         TabIndex        =   85
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NOM            LOT               KG"
         Height          =   255
         Left            =   1065
         TabIndex        =   60
         Top             =   150
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command29 
      Height          =   390
      Left            =   10710
      Picture         =   "baixes laminadores.frx":492A
      Style           =   1  'Graphical
      TabIndex        =   109
      ToolTipText     =   "Calcul diametre"
      Top             =   900
      Width           =   375
   End
   Begin FOXITREADEROCXLib.FoxitReaderOCX AcroPDF1 
      Height          =   1770
      Left            =   9060
      TabIndex        =   108
      Top             =   1695
      Visible         =   0   'False
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   3122
      _StockProps     =   0
      SRC             =   ""
   End
   Begin VB.CommandButton Command21 
      Height          =   390
      Left            =   11085
      Picture         =   "baixes laminadores.frx":4EB4
      Style           =   1  'Graphical
      TabIndex        =   99
      ToolTipText     =   "Calcul diametre"
      Top             =   900
      Width           =   375
   End
   Begin VB.CommandButton botodescansrelleu 
      Height          =   390
      Left            =   11460
      Picture         =   "baixes laminadores.frx":543E
      Style           =   1  'Graphical
      TabIndex        =   96
      ToolTipText     =   "Control Descans i Relleu"
      Top             =   900
      Width           =   375
   End
   Begin VB.CommandButton maquina 
      BackColor       =   &H00FF8080&
      Caption         =   "Maq: 0"
      Height          =   390
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   75
      Width           =   765
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
      Picture         =   "baixes laminadores.frx":59C8
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "Ensenya els valors de la cola."
      Top             =   5385
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
      Left            =   10470
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "lotslam"
      Top             =   3555
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H008080FF&
      Caption         =   "No Acabada"
      Height          =   660
      Left            =   9915
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   60
      Width           =   1005
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0080FF80&
      Caption         =   "Acabar Comanda"
      Height          =   645
      Left            =   8835
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   75
      Width           =   1080
   End
   Begin VB.Frame calculant 
      Height          =   2580
      Left            =   3570
      TabIndex        =   72
      Top             =   8340
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
         TabIndex        =   73
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
      Left            =   7545
      Picture         =   "baixes laminadores.frx":6AEA
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   825
      Width           =   495
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
      RecordSource    =   "bobinesentlam"
      Top             =   6975
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
      Left            =   10515
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "lamempalmes"
      Top             =   7200
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Calcular Totals"
      Height          =   390
      Left            =   5925
      Picture         =   "baixes laminadores.frx":6F64
      TabIndex        =   63
      Top             =   930
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
      Width           =   2010
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   0
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
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
      Left            =   10935
      Picture         =   "baixes laminadores.frx":7C3A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   45
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
      Left            =   10725
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "bobineslam"
      Top             =   7470
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Frame Frame2 
      Caption         =   "Totals"
      Height          =   765
      Left            =   120
      TabIndex        =   9
      Top             =   7665
      Width           =   11580
      Begin VB.CommandButton Command17 
         Height          =   330
         Left            =   11070
         Picture         =   "baixes laminadores.frx":983C
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Opcions d'Encarregat."
         Top             =   390
         Width           =   450
      End
      Begin VB.CheckBox comandaacavada 
         Caption         =   "Acabada"
         Enabled         =   0   'False
         Height          =   225
         Left            =   10410
         TabIndex        =   76
         Top             =   195
         Width           =   1125
      End
      Begin VB.TextBox hmaquina 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   540
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox hfunc 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox tkilos 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4485
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   375
         Width           =   840
      End
      Begin VB.TextBox tmetres 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3570
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   390
         Width           =   840
      End
      Begin VB.TextBox kiloshora 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   5430
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   840
      End
      Begin VB.TextBox tbob 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   405
         Width           =   840
      End
      Begin VB.Label ettipuscola 
         BackStyle       =   0  'Transparent
         Caption         =   "Cola especial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   540
         Left            =   6990
         TabIndex        =   97
         Top             =   165
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "H. Màquina"
         Height          =   210
         Left            =   570
         TabIndex        =   21
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Hores Func."
         Height          =   195
         Left            =   1605
         TabIndex        =   20
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bob."
         Height          =   210
         Left            =   2670
         TabIndex        =   19
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Metres"
         Height          =   210
         Left            =   3540
         TabIndex        =   18
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Metres/Min"
         Height          =   210
         Left            =   5400
         TabIndex        =   17
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Kilos"
         Height          =   210
         Left            =   4500
         TabIndex        =   16
         Top             =   180
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
      Top             =   975
      Width           =   3675
   End
   Begin VB.Timer rellotge 
      Left            =   345
      Top             =   420
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2055
      TabIndex        =   5
      Top             =   165
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Funcionament"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   150
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Canvi Màquina"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   150
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Capçalera"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   150
      Width           =   1335
   End
   Begin VB.Data laminadores 
      Caption         =   "laminadores"
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
      RecordSource    =   "laminadores"
      Top             =   735
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "baixes laminadores.frx":9DC6
      Height          =   2235
      Left            =   165
      OleObjectBlob   =   "baixes laminadores.frx":9DDC
      TabIndex        =   6
      Top             =   1305
      Width           =   11610
   End
   Begin VB.TextBox linkcomanda 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   735
      TabIndex        =   81
      TabStop         =   0   'False
      Tag             =   "888"
      Top             =   510
      Width           =   1065
   End
   Begin VB.TextBox comanda 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   735
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "888"
      Text            =   "135716"
      Top             =   180
      Width           =   1050
   End
   Begin VB.Frame framebobines 
      Caption         =   "Bobines"
      Height          =   3840
      Left            =   120
      TabIndex        =   22
      Top             =   3810
      Width           =   11655
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   10305
         Picture         =   "baixes laminadores.frx":BD65
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   2955
         Width           =   915
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
         Picture         =   "baixes laminadores.frx":D437
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Ensenya Pantones utilitzats"
         Top             =   2265
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
         Height          =   735
         Left            =   10290
         TabIndex        =   24
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton Command7 
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
         Left            =   10305
         Picture         =   "baixes laminadores.frx":E481
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   900
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
         Left            =   11130
         TabIndex        =   25
         Top             =   270
         Width           =   375
      End
      Begin MSDBGrid.DBGrid reixabobines 
         Bindings        =   "baixes laminadores.frx":10083
         Height          =   3570
         Left            =   120
         OleObjectBlob   =   "baixes laminadores.frx":10095
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
         TabIndex        =   64
         Top             =   2820
         Width           =   6315
      End
   End
   Begin VB.Frame frameempalmes 
      Caption         =   "Senyals"
      Height          =   3795
      Left            =   5625
      TabIndex        =   65
      Top             =   3930
      Visible         =   0   'False
      Width           =   4725
      Begin VB.CommandButton beliminarempalme 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Eliminar l'Empalme"
         Height          =   480
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Eliminar bobina d'entrada"
         Top             =   3255
         Width           =   1995
      End
      Begin MSDBGrid.DBGrid reixaempalmes 
         Bindings        =   "baixes laminadores.frx":11621
         Height          =   3015
         Left            =   60
         OleObjectBlob   =   "baixes laminadores.frx":11634
         TabIndex        =   66
         Top             =   195
         Width           =   4515
      End
   End
   Begin VB.Label hora 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   450
      TabIndex        =   7
      Top             =   870
      Width           =   1815
   End
   Begin VB.Shape reciclarmaterial2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   1830
      Shape           =   3  'Circle
      Top             =   555
      Width           =   225
   End
   Begin VB.Shape reciclarmaterial1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   1830
      Shape           =   3  'Circle
      Top             =   195
      Width           =   225
   End
   Begin VB.Label primerproces 
      BackColor       =   &H008080FF&
      Caption         =   "Primer Proces"
      Height          =   165
      Left            =   1905
      TabIndex        =   94
      Top             =   -15
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label numgrup 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6915
      TabIndex        =   92
      Top             =   3510
      Width           =   2985
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
      TabIndex        =   91
      ToolTipText     =   "Si es 'E' son bobines d'Estoc i si es 'P' de Packing-list"
      Top             =   75
      Width           =   330
   End
   Begin VB.Label proces 
      Height          =   165
      Left            =   0
      TabIndex        =   86
      Top             =   -15
      Width           =   390
   End
   Begin VB.Label Label18 
      Caption         =   "Lot2:"
      Height          =   285
      Left            =   150
      TabIndex        =   82
      Top             =   585
      Width           =   525
   End
   Begin VB.Label Label17 
      Caption         =   "Lot1:"
      Height          =   285
      Left            =   150
      TabIndex        =   80
      Top             =   225
      Width           =   525
   End
   Begin VB.Label firmat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7215
      TabIndex        =   78
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
      Height          =   240
      Left            =   180
      TabIndex        =   67
      Top             =   3585
      Width           =   6090
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
      TabIndex        =   28
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
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Sub obrestocks(Optional noobrirbd As Boolean)
 Dim camistocks As String
' Set ws = DBEngine.CreateWorkspace("", "admin", "")
 ' If estaobertstocks Then dbtemp.Execute "delete * from selecciobobentrada": Exit Sub
camistocks = llegir_ini("General", "ruta_stocks", "comandes.ini")
'If camistocks = "{[}]" Then camistocks = "\\Ser2\documentos\Stock Reclamaciones\Estoc inplacsa.mdb"
'If Not existeix(camistocks) Then
'    MsgBox "Error obrint la la base de dades de Estocs (Palets) intentarem obrir la BD per defecte", vbCritical, "Error"
'    camistocks = "\\serverprodu\dades\progcomandes\dades\palets.mdb"
'End If

If camistocks = "{[}]" Then escriure_ini "General", "ruta_stocks", rutadelfitxer(cami) + "palets.mdb", "comandes.ini"
camistocks = llegir_ini("General", "ruta_stocks", "comandes.ini")
If Not noobrirbd Then
   Set dbstocks = OpenDatabase(camistocks)
 '  dbtemp.Execute "delete * from selecciobobentrada"
End If
  
End Sub


Sub calcular_totals(Optional obrint As Boolean)
  Dim total As Double
  Dim hores As Double
  Dim bkimp As Double
  Dim bkbob As Double
  barraestat.Caption = "Calculant els totals..."
  'calculant.Visible = True
  fcalculant.Show 0, Me
  calculant.Top = 2222
  DoEvents
  
  
  'On Error GoTo fi
  reixa.EditActive = False
  reixabobines.EditActive = False
  If laminadores.Recordset.EOF Or cadbl(laminadores.Recordset!id) = 0 Then GoTo fi
  
  '---- guardo la posicio de linies imp i de bobina x recuperarlames avall
  If laminadores.Recordset!tipus = "F" Then bkimp = atrim(cadbl(laminadores.Recordset!id))
  If Not bobines.Recordset.EOF Then bkbob = atrim(cadbl(bobines.Recordset!numerodebobina))
  '------
  
  On Error Resume Next
  laminadores.Recordset.MoveLast
  laminadores.Recordset.MoveFirst
  i = 0
  While Not laminadores.Recordset.EOF And i < 100
   'On Error GoTo 0
   i = i + 1
   If laminadores.Recordset!tipus = "F" Then
    If laminadores.Recordset.EditMode = 0 Then laminadores.Recordset.Edit
    Set rsttmp = dbtmpb.OpenRecordset("select count(*) as elgran from bobineslam where controlid=" + atrim(laminadores.Recordset!id))
    If Not rsttmp.EOF Then laminadores.Recordset!totalbobines = rsttmp!elgran
  
    Set rsttmp = dbtmpb.OpenRecordset("select sum(kilos) as elgran from bobineslam where controlid=" + atrim(laminadores.Recordset!id))
    If Not rsttmp.EOF Then laminadores.Recordset!totalkilos = rsttmp!elgran
  
    Set rsttmp = dbtmpb.OpenRecordset("select sum(metres) as elgran from bobineslam where controlid=" + atrim(laminadores.Recordset!id))
    If Not rsttmp.EOF Then laminadores.Recordset!totalmetres = rsttmp!elgran
  
    Set rsttmp = dbtmpb.OpenRecordset("select id,metres from bobineslam where metres=0 and controlid=" + atrim(laminadores.Recordset!id))
    If Not rsttmp.EOF Then
     If rsttmp!id <> bobines.Recordset!id Then MsgBox "Hi ha bobines sense metres"
    End If
    If laminadores.Recordset.RecordCount - 1 = laminadores.Recordset.AbsolutePosition Then
        laminadores.Recordset!totallitresresina = kbpantone(0)
        laminadores.Recordset!totallitresenduridor = kbpantone(1)
        If laminadores.Recordset!totallitresresina <> kbpantone(0) Then laminadores.Recordset!totallitresresina = IIf(InStr(1, kbpantone(0), ",") > 0, substituir(kbpantone(0), ",", "."), substituir(kbpantone(0), ".", ","))
        If laminadores.Recordset!totallitresenduridor <> kbpantone(1) Then laminadores.Recordset!totallitresresina = IIf(InStr(1, kbpantone(1), ",") > 0, substituir(kbpantone(1), ",", "."), substituir(kbpantone(1), ".", ","))
    End If
    laminadores.Recordset.Update
   End If
  
   
   With laminadores.Recordset
    total = 0
    'On Error Resume Next
     If Not IsDate(CVDate(atrim(!datainici))) Or Not IsDate(CVDate(atrim(!horainici))) Or Not IsDate(atrim(!horafi)) Or Not IsDate(CVDate(atrim(!datafi))) Then
      If Not obrint And laminadores.Recordset!id <> bkimp Then MsgBox "Error d'hora d'inici o final de funcionament. Corretgeix l'error per poder continuar correctament."
       Else
            total = DateDiff("n", CVDate(atrim(!datainici) + " " + atrim(!horainici)), CVDate(atrim(!datafi) + " " + atrim(!horafi)))
            total = Format(total / 60, "#,##0.00")
            
     End If
    If laminadores.Recordset.EditMode = 0 Then laminadores.Recordset.Edit
     laminadores.Recordset!totalhores = total
     laminadores.Recordset.Update
   End With
  laminadores.Recordset.MoveNext
 Wend
 If i >= 100 Then MsgBox "Hi ha algun error de dades no puc calcular correctament.": GoTo fi
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
     laminadores.Recordset.FindFirst "id=" + atrim(bkimp)
     bobines.Recordset.FindFirst "numerodebobina=" + atrim(bkbob)
   Else: laminadores.Recordset.MoveLast
  End If
  '---
fi:
'calculant.Visible = False
barraestat.Caption = ""
Unload fcalculant
form1.SetFocus
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
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalbobines) as elgran from laminadores totalbobines where comanda=" + atrim(cadbl(comanda.Text)))
  If Not rsttmp.EOF Then tbob = cadbl(rsttmp!elgran)

  
'hores func
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from laminadores totalbobines where comanda=" + atrim(cadbl(comanda.Text)) + " and tipus='F'")
  If Not rsttmp.EOF Then hfunc = cadbl(rsttmp!elgran)
  

'hores maquina
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalhores) as elgran from laminadores totalbobines where comanda=" + atrim(cadbl(comanda.Text)) + " and tipus='C'")
  If Not rsttmp.EOF Then hmaquina = cadbl(rsttmp!elgran)

'total kilos
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalkilos) as elgran from laminadores  where comanda=" + atrim(cadbl(comanda.Text)))
  If Not rsttmp.EOF Then tkilos = cadbl(rsttmp!elgran)
  
'total metres
  Set rsttmp = dbtmpb.OpenRecordset("select sum(totalmetres) as elgran from laminadores totalbobines where comanda=" + atrim(cadbl(comanda.Text)))
  If Not rsttmp.EOF Then tmetres = cadbl(rsttmp!elgran)
  

  
  guarda_totals
  ensenya_totals
End Sub

Sub guarda_totals()
 Set rsttmp = dbtmpb.OpenRecordset("select * from laminadorestot where comanda=" + atrim(cadbl(comanda)))
  If rsttmp.EOF Then
      rsttmp.AddNew
    Else: rsttmp.Edit
  End If
  With rsttmp
    !firmat = atrim(firmat.Caption)
    !comanda = cadbl(comanda)
    !hcanvi = cadbl(hmaquina)
    !hfuncio = cadbl(hfunc)
    !tbobines = cadbl(tbob)
    !tkilos = cadbl(tkilos)
    !tmetres = cadbl(tmetres)
    !metresmin = cadbl(kiloshora)
    
    !acavada = cadbl(comandaacavada.Value)
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
Set rsttmp = dbtmpb.OpenRecordset("select * from laminadorestot where comanda=" + atrim(cadbl(comanda)))

  With rsttmp
    'comanda = atrim(!comanda)
    firmat = atrim(!firmat)
    hmaquina = atrim(!hcanvi)
    hfunc = atrim(!hfuncio)
    tbob = atrim(!tbobines)
    'tprova = atrim(!tprova)
    tkilos = atrim(!tkilos)
    tmetres = atrim(!tmetres)
    kiloshora = atrim(!metresmin)
    comandaacavada.Value = cadbl(!acavada)
    'If Not (bobines.Recordset.EOF Or bobines.Recordset.BOF) Then
    ' !kilostinta = cadbl(bobines.Recordset!kgtinta)
    ' If Not IsNull(bobines.Recordset!datafi) Then !dataimpressio = bobines.Recordset!datafi
     '!impressora = cadbl(impresores.Recordset!numeromaquina)
     '!operari = cadbl(bobines.Recordset!operari)
    'End If
  
  End With

End Sub


Private Sub beliminarempalme_Click()
  Dim vnumempalmes As Long
  If empalmes.Recordset.EOF Then Exit Sub
  If MsgBox("Segur que vols eliminar aquest empalme?", vbYesNo + vbInformation + vbDefaultButton2, "Error") = vbYes Then
    empalmes.Recordset.Delete
    empalmes.Refresh
    vnumempalmes = 0
    If Not empalmes.Recordset.EOF Then
       empalmes.Recordset.MoveLast
       vnumempalmes = empalmes.Recordset.RecordCount
    End If
    If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
    bobines.Recordset!numempalmes = vnumempalmes
    bobines.Recordset.Update
  End If
End Sub

Private Sub bobentrada_DblClick()
  Dim numoptmp As Integer
  Dim nomoptmp As String
  Dim rsttmpbob As Recordset
  Dim rsttmpbobimp As Recordset
  Dim rsttmpimp As Recordset
  Dim ensenyar As String
  Dim carregataulatmp As Boolean
  Exit Sub
  If r = "carregartaulatmp" Then carregartaulatmp = True
  ratoli "esperar"
  On Error Resume Next
  Unload formseleccio
  On Error GoTo 0
  If Not carregartaulatmp And cadbl(bobentrada.Columns(0).Text) = 0 Then
     If MsgBox("Desbobinador 1", vbYesNo, "Selecció de Desbobinador") = vbYes Then
         bobentrada.Columns(0).Text = "1"
          Else: bobentrada.Columns(0).Text = "2"
     End If
  End If
  If framebobentrada.Visible Then bobentrada.SetFocus
  If ensenyartoteslesbobines <> 1 Then
     ensenyar = "not utilitzadaabaixa and "
   Else: ensenyar = ""
  End If
  If carregartaulatmp Then ensenyar = ""
  If r = "carregartaulatmp2" Then carregartaulatmp = True
  obrestocks
  If proces <> "PC2" Then
    r = "SELECT DISTINCTROW numcom, Idpalet, Idbobina FROM bobines where " + ensenyar + " (bobines.Numcom) = '" + atrim(cadbl(linkcomanda)) + "' "
      Else: r = "SELECT DISTINCTROW numcom, Idpalet, Idbobina FROM bobines where " + ensenyar + " (bobines.Numcom) = '" + atrim(cadbl(comanda)) + "' "
  End If
  crear_taula_bobentrada
  Set rststocks = dbstocks.OpenRecordset(r)
  Set rsttmpbob = dbtmpb.OpenRecordset("bobentradatmplam")
  If proces <> "PC2" Then
     Set rsttmpimp = dbtmpb.OpenRecordset("select * from impressores where tipus='F' and comanda=" + atrim(cadbl(comanda)))
    Else: Set rsttmpimp = dbtmpb.OpenRecordset("select * from laminadores where tipus='F' and comanda=" + atrim(cadbl(comanda)))
  End If
  While Not rststocks.EOF
    rsttmpbob.AddNew
    rsttmpbob!idbobina = 0
    rsttmpbob!numlot = rststocks!numcom
    rsttmpbob!numpalet = rststocks!idpalet
    rsttmpbob!numbobent = rststocks!idbobina
    rsttmpbob.Update
    rststocks.MoveNext
  Wend
  While Not rsttmpimp.EOF
    If proces <> "PC2" Then
       Set rsttmpbobimp = dbtmpb.OpenRecordset("select * from bobinesimp where " + ensenyar + " controlid=" + atrim(cadbl(rsttmpimp!id)))
      Else: Set rsttmpbobimp = dbtmpb.OpenRecordset("select * from bobineslam where " + ensenyar + " controlid=" + atrim(cadbl(rsttmpimp!id)))
    End If
    While Not rsttmpbobimp.EOF
     rsttmpbob.AddNew
     rsttmpbob!idbobina = cadbl(rsttmpbobimp!id)
     rsttmpbob!numlot = cadbl(rsttmpimp!comanda)
     rsttmpbob!numpalet = cadbl(comanda.Tag)
     rsttmpbob!numbobent = cadbl(rsttmpbobimp!numerodebobina)
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
 ' dbtmpb.Close
 ' Set dbtmpb = OpenDatabase(laminadores.DatabaseName)
  'MsgBox bobinesent.EditMode
  'wait (3)
 
  Set rsttmp = dbtmpb.OpenRecordset("bobentradatmplam")
  If rsttmp.EOF Then MsgBox "No hi ha bobines d'entrada per escullir": dbstocks.Close: Exit Sub
   Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "bobentradatmplam"
   formseleccio.Caption = "Selecció bobina d'entrada"
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns(0).Visible = False
   formseleccio.DBGrid2.Columns(1).Visible = False
   formseleccio.DBGrid2.Columns(2).Width = 2500
   formseleccio.DBGrid2.Columns(3).Width = 2500
   ratoli "normal"
   formseleccio.Show 1
  If seleccioret = 1 Then
   If cadbl(formseleccio.Data1.Recordset!idbobina) = 0 Then
       bobentrada.Columns(1) = cadbl(formseleccio.Data1.Recordset!numpalet)
       bobentrada.Columns(2) = cadbl(formseleccio.Data1.Recordset!numbobent)
       If bobinesent.Recordset.EditMode = 0 Then bobinesent.Recordset.Edit
       'si es final gravo amb majuscula si no amb minuscula per saber si estava acavada o no
       If MsgBox("Ès final de bobina?", vbYesNo, "Bobina") = vbYes Then
          r = "P": dbstocks.Execute "update  bobines set utilitzadaabaixa=True where idpalet=" + bobentrada.Columns(1) + " and idbobina=" + bobentrada.Columns(2)
            Else: r = "p": dbstocks.Execute "update  bobines set utilitzadaabaixa=False where idpalet=" + bobentrada.Columns(1) + " and idbobina=" + bobentrada.Columns(2)
       End If
       bobinesent.Recordset!paletobobina = r
       bobinesent.Recordset!idbobina = 0
        Else
          bobentrada.Columns(1) = cadbl(formseleccio.Data1.Recordset!numpalet)
          bobentrada.Columns(2) = cadbl(formseleccio.Data1.Recordset!numbobent)
          If bobinesent.Recordset.EditMode = 0 Then bobinesent.Recordset.Edit
          'si es final gravo amb majuscula si no amb minuscula per saber si estava acavada o no
          If MsgBox("Ès final de bobina?", vbYesNo, "Bobina") = vbYes Then
            r = "B"
            If proces <> "PC2" Then
                dbtmpb.Execute "update  bobinesimp set utilitzadaabaixa=True where id=" + atrim(cadbl(formseleccio.Data1.Recordset!idbobina))
               Else: dbtmpb.Execute "update  bobineslam set utilitzadaabaixa=True where id=" + atrim(cadbl(formseleccio.Data1.Recordset!idbobina))
            End If
              Else:
               r = "b"
               If proces <> "PC2" Then
                 dbtmpb.Execute "update  bobinesimp set utilitzadaabaixa=False where id=" + atrim(cadbl(formseleccio.Data1.Recordset!idbobina))
                  Else: dbtmpb.Execute "update  bobineslam set utilitzadaabaixa=False where id=" + atrim(cadbl(formseleccio.Data1.Recordset!idbobina))
               End If
          End If
          bobinesent.Recordset!paletobobina = r
          bobinesent.Recordset!idbobina = cadbl(formseleccio.Data1.Recordset!idbobina)
   End If
  End If
  If bobinesent.EditMode > 0 Then bobinesent.Recordset.Update
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
End Sub
Function existeixlataula(vnomtaula As String) As Boolean
  Dim rstp As Recordset
  On Error GoTo errortaula
  existeixlataula = True
  Set rstp = dbtmpb.OpenRecordset("select * from " + vnomtaula)
  Set rstp = Nothing
  On Error GoTo 0
  Exit Function
errortaula:
  existeixlataula = False
  On Error GoTo 0
End Function
Sub crear_taula_bobentrada()
  Dim camps As String
  'On Error GoTo borrar
  ' dbtmpb.Execute "drop table bobentradatmplam"
  'On Error GoTo 0
  If Not existeixlataula("bobentradatmplam") Then
    camps = "idbobina double,numlot double,numpalet double,numbobent double"
    'ample double,plegat double,solapa double,espessor double,metres double,kilos double)"
    dbtmpb.Execute ("create table bobentradatmplam (" + camps) + ")"
      Else: dbtmpb.Execute "delete * from bobentradatmplam"
  End If
  Exit Sub
borrar:
  On Error Resume Next
  dbtmpb.Execute ("delete * from bobentradatmplam")
  
  
End Sub

Private Sub bobentrada_KeyDown(KeyCode As Integer, Shift As Integer)
tempseditant = Now
End Sub

Private Sub bobentrada_KeyUp(KeyCode As Integer, Shift As Integer)
If bobentrada.col = 1 And Len(bobentrada.Text) = 5 And KeyCode > 46 Then bobentrada.col = 2

End Sub

Private Sub bobentrada_LostFocus()
 ' SI FAIG UN LOSTFOCUS DONA ERROR AL COMPROVAR COSES AL ESCULLIR LES BOBINES D'ENTRADA
 
 'On Error Resume Next
 ' bobinesent.UpdateRecord
 ' si
End Sub

Private Sub bobentrada_OnAddNew()
 bobinesent.Recordset!id = bobines.Recordset!id
 bobentrada.col = 0
End Sub

Private Sub bobentrada_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  'If cadbl(bobentrada.Columns(0).Text) <> 1 And cadbl(bobentrada.Columns(0).Text) <> 2 And cadbl(bobentrada.Columns(0).Text) <> 0 Then
  '   MsgBox "El numero de desbobinador està malament"
  'End If
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
       bobinesent.RecordSource = "select * from bobinesentlam where id=99999999"
     Else
       bobinesent.RecordSource = "select * from bobinesentlam where id=" + atrim(cadbl(bobines.Recordset!id)) '+ " ORDER BY desb"
   End If
   bobinesent.Refresh
 End If
 
End Sub

Private Sub clixes_Click()
 
End Sub
Sub finalitza_seccio()
  On Error GoTo fi
  If laminadores.Recordset.EOF Then Exit Sub
  On Error Resume Next
  laminadores.Recordset.MoveLast
  If IsDate(laminadores.Recordset!datafi) Then r = "no": Exit Sub
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
  laminadores.Recordset.Edit
  laminadores.Recordset!datafi = Date: laminadores.Recordset!horafi = Time
  Select Case laminadores.Recordset!tipus
   Case "C"
   Case "M"
   Case "A"
   Case "F"
  End Select
  laminadores.Recordset.Update
calcular_totals
fi:
End Sub



Private Sub canvienfilada_DblClick()
If canvienfilada = "Si" Then
   canvienfilada = "No"
 Else: canvienfilada = "Si"
End If
End Sub

Private Sub botodescansrelleu_Click()
 Load formdescansirelleu
   If Not laminadores.Recordset.EOF Then
        laminadores.Recordset.MoveLast
        If Not laminadores.Recordset.EOF Then
           If Not IsDate(laminadores.Recordset!datafi) And Not IsDate(laminadores.Recordset!horafi) Then
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
    If Not laminadores.Recordset.EOF Then
        laminadores.Recordset.MoveLast
        If Not IsDate(atrim(laminadores.Recordset!datafi)) Or Not IsDate(atrim(laminadores.Recordset!horafi)) Then
           'If MsgBox("No hi ha la hora de fi de funcionament, Vols que el col.loqui automàticament?", vbInformation + vbYesNo, "Atenció") = vbYes Then
            laminadores.Recordset.Edit
            laminadores.Recordset!datafi = Date
            laminadores.Recordset!horafi = Time
            laminadores.Recordset.Update
           'End If
        End If
    End If
    command4_click
    wait 1
    capcalera.Hide
End Sub


Private Sub botoensenyarpacking_Click()
 Dim palet As Double
 Dim bobina As Double
 Dim utilitzades As String
 utilitzades = "noutilitzades"
 If ensenyartoteslesbobines.Value = 1 Then utilitzades = ""
 carregar_bobinesdentrada "ensenyar" + utilitzades, 1, palet, bobina, ncomanda, , ncomanda2, IIf(primerproces.Tag = "invertit", True, False), True

 If palet > 0 And bobina > 0 And Not bobines.Recordset.EOF Then
    'bobentrada.Columns("Palet") = atrim(palet): bobentrada.Columns("Bobina") = atrim(bobina)
    seleccionardesb.Show 1
    afegir_labobinadentrada palet, bobina, desb, IIf(primerproces.Tag = "invertit", True, False)
    If palet < (cadbl(comanda) - 3) Or palet > (cadbl(comanda) + 3) Then
       bobinesent.Recordset.FindFirst "palet=" + atrim(palet) + " and bobina=" + atrim(bobina)
       demanar_verificacio_espesoritractat palet, bobina
       'mirar_imprimir_controlqualitatPE palet, bobina
       imprimir_controlqualitatbobinaentrada palet, bobina, desb
    End If
    
    'demanar si es final de bobina
'    carregar_bobinesdentrada "marcarutilitzadademanar", , palet, bobina, ncomanda, False, ncomanda2
 End If
 Unload bobinesdentrada
 botoensenyarpacking.Tag = ""
 bobinesent.UpdateRecord
 ratoli "normal"
End Sub

Sub imprimir_controlqualitatbobinaentrada(palet As Double, bobina As Double, desb As Byte)
   Dim ultimalinia As String
   Dim esp As Double
   Dim cont As Byte
  If bobinesent.Recordset.EOF Then Exit Sub
   While cadbl(bobinesent.Recordset!verificacioespesor) = 0 And cont < 6
    bobinesent.Refresh
    bobinesent.Recordset.FindFirst "palet=" + atrim(palet) + " and bobina=" + atrim(bobina)
    wait 1
    cont = cont + 1
   Wend
  llistat.DataFiles(0) = ""
   llistat.DataFiles(1) = ""
   ultimalinia = "Lam-" + atrim(nummaq) + " Op: " + atrim(numop) + " Com: " + atrim(capcalera.Controls("lotdesb" + atrim(desb))) + " Fecha: " + Format(Now, "dd/mm/yy")
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
   llistat.Formulas(5) = "valorvalidespesor='Marge: >" + atrim(Redondejar(esp - (esp / 10), 1)) + " i <" + atrim(Redondejar(esp + (esp / 10), 1)) + "'"
   
   llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "verificacioqualitatlaminadoresbobinesentrada.rpt"
   llistat.Destination = crptToPrinter
    llistat.CopiesToPrinter = 1
   llistat.DiscardSavedData = True
' llistat.PrinterName = llegir_ini("Impressores", "nomfulla", "baixesimpressora.ini")
' llistat.PrinterPort = llegir_ini("Impressores", "portfulla", "baixesimpressora.ini")
' llistat.PrinterDriver = llegir_ini("Impressores", "driverfulla", "baixesimpressora.ini")
   DoEvents
   If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
   llistat.Action = 1
   'MsgBox "ATENCIÓ CONTROL DE VERIFICACIÓ DE QUALITAT." + Chr(10) + "VERIFICA LA IMPRESIÓ AMB L'ETIQUETA IMPRESA", vbInformation, "VERIFICACIÓ QUALITAT"
   llistat.SelectionFormula = ""
   llistat.DataFiles(0) = ""

End Sub
Function degramsamicres(codimat As Double) As Double
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("Select micresdelsgrm2 from materials where codi=" + atrim(codimat))
   If Not rst.EOF Then degramsamicres = IIf(IsNull(rst!micresdelsgrm2), 0, rst!micresdelsgrm2)
End Function
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
   obrestocks
   Set rstb = dbbaixes.OpenRecordset("select * from bobinesentlam where id=" + atrim(cadbl(bobines.Recordset!id)) + " and palet=" + atrim(palet) + " and bobina=" + atrim(bobina))
   If Not rstb.EOF Then
      Set rstp = dbstocks.OpenRecordset("select codimatprognou,grmsm2,micres from palets where idpalet=" + atrim(palet))
      If rstp.EOF Then Exit Sub
      Set rstm = dbtmp.OpenRecordset("SELECT materials.codi, familiescolorants.descripcio FROM familiescolorants RIGHT JOIN materials ON familiescolorants.codi = materials.familiacol where materials.codi=" + atrim(rstp!codimatprognou))
      If rstm.EOF Then Exit Sub
      colormat = UCase(atrim(rstm!descripcio))
      If cadbl(rstp!micres) > 0 Then espesorcorrecte = cadbl(rstp!micres): sonmicres = True: pregunta = "Entra el valor de l'espesor del micrometre." + Chr(10) + " +-10% de " + atrim(espesorcorrecte) + " Micres. VALOR MICROMETRE --> " + atrim(espesorcorrecte * 4)
      'If cadbl(rstp!grmsm2) Then espesorcorrecte = cadbl(rstp!grmsm2): pregunta = "Entra l'espesor en GRAMS PER METRE QUADRAT." + Chr(10) + " +-5% de " + atrim(espesorcorrecte) + " Grm/m2"
      If cadbl(rstp!grmsm2) > 0 Then
         espesorcorrecte = degramsamicres(rstp!codimatprognou)
         If espesorcorrecte = 0 Then espesorcorrecte = cadbl(rstp!grmsm2)
         sonmicres = True 'espesorcorrecte = cadbl(rstp!grmsm2):
      End If
      pregunta = " +-10% de " + atrim(espesorcorrecte) + " Micres" ' + Chr(10) + "  (Pensa que són 4 fulles de material)"
      While espesor = 0 Or Not correcte
         If nummaq = 3 Then
             espesor = demanar_valor_micrometre(pregunta, atrim(palet) + "/" + atrim(bobina))
            Else: espesor = cadbl(InputBox("Entra el valor de l'espesor del micrometre." + vbNewLine + pregunta + " VALOR MICROMETRE -->" + atrim(espesorcorrecte) + vbNewLine + "Bobina: " + atrim(palet) + "/" + atrim(bobina), "Espesor bobina entrada"))
         End If
         If sonmicres Then
            espesor = Redondejar(espesor, 1)
            correcte = ((espesor <= (espesorcorrecte + (espesorcorrecte * 10 / 100))) And (espesor >= (espesorcorrecte - (espesorcorrecte * 10 / 100))))
           Else: correcte = ((espesor < (espesorcorrecte + (espesorcorrecte * 5 / 100))) And (espesor > (espesorcorrecte - (espesorcorrecte * 5 / 100))))
         End If
      Wend
      tractat = IIf(MsgBox("Comprova que el tractat està a la cara correcte." + Chr(10) + "ES CORRECTE LA CARA DEL TRACTAT?", vbInformation + vbYesNo + vbDefaultButton2, "Comprovació del tractat") = vbYes, True, False)
      While MsgBox("Comprova que el COLOR del material sigui correcte." + Chr(10) + atrim(colormat), vbInformation + vbYesNo, "Verificacio") = vbNo
        DoEvents
      Wend
      
      rstb.Edit
      'si son grm/m2 ho passo amb negatiu
      If Not sonmicres Then espesorcorrecte = espesorcorrecte * -1
      rstb!espesorteoric = espesorcorrecte
      rstb!verificacioespesor = IIf(sonmicres, espesor, espesor * -1)
      rstb!verificaciotractat = tractat
      rstb!verificaciocolor = Mid(colormat, 1, 15)
      rstb.Update
   End If
   
End Sub
Function demanar_valor_micrometre(vpregunta As String, vnumpalet As String) As Double
   Load formcapturarmicrometre
   formcapturarmicrometre.etbobina = vnumpalet
   formcapturarmicrometre.ettolerancia = vpregunta
   formcapturarmicrometre.Show 1
   demanar_valor_micrometre = cadbl(formcapturarmicrometre.vmicrometre)
   Unload formcapturarmicrometre
End Function
Sub mirar_imprimir_controlqualitatPE(palet As Double, bobina As Double)
   Dim rsttmp3 As Recordset
   Dim rsttmp2 As Recordset
   Dim descmat As String
    Set rsttmp3 = dbstocks.OpenRecordset("select codimatprognou from palets where idpalet=" + atrim(palet))
    If rsttmp3.EOF Then Exit Sub
    Set rsttmp2 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rsttmp3!codimatprognou)))
    If rsttmp2.EOF Then Exit Sub
    descmat = capcalera.descripciomaterial(rsttmp2)
    If InStr(1, UCase(descmat), "PEBD") > 0 Then
       imprimir_controlqualitatPE numerocomandaambPE(cadbl(comanda), cadbl(linkcomanda)), numop, palet, bobina
    End If
    Set rsttmp3 = Nothing
    Set rsttmp2 = Nothing
End Sub
Function numerocomandaambPE(numc As Double, numc2 As Double) As Double
    Dim desc As String
    numerocomandaambPE = 0
    desc = generardadescomanda(numc)
    If InStr(1, desc, "PEBD") Then numerocomandaambPE = numc
    If numerocomandaambPE = 0 Then
      desc = generardadescomanda(numc2)
      If InStr(1, desc, "PEBD") Then numerocomandaambPE = numc2
    End If
End Function
Function generardadescomanda(numc As Double) As String
    Dim rstc As Recordset
    Dim rstd1 As Recordset
    Dim rstd2 As Recordset
    Dim rstmat As Recordset
    Dim codimat1 As Double
    Dim codimat2 As Double
    Dim nommaterial As String
    Dim descmicres As String
    Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
    
    If Not rstc.EOF Then
      '  Set rstd1 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(rstc!lotmatdesb1)))
      '  Set rstd2 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(rstc!lotmatdesb2)))
      '  If rstd1.EOF Or rstd2.EOF Then Exit Sub
      '  If rstd!refilatd <> 1 Then
           Set rstmat = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstc!materialex)))
           If Not rstmat.EOF Then nommaterial = capcalera.descripciomaterial(rstmat)
           generardadescomanda = nommaterial
           'Set rstmat = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstd2!materialex)))
           'If Not rstmat.EOF Then nommaterial = descripciomaterial(rstmat, True)
           'generardadescomanda = generardadescomanda + " + " + nommaterial
        'End If
        
    End If
    Set rstmat = Nothing
End Function
Sub carregarvalorsespessor(vcodiclient As Double, vt As Double, vc As Double, vtole As Double)
   Dim rst2 As Recordset
   Set rst2 = dbtmp.OpenRecordset("select * from clients where codi=" + atrim(vcodiclient), , ReadOnly)
   If Not rst2.EOF Then
      vt = cadbl(rst2!espessortinta)
      vc = cadbl(rst2!espessorcola)
      vtole = cadbl(rst2!espessortolerancia)
      'If vt = 0 Or vc = 0 Or vtole = 0 Then vt = 1.5: vc = 1.5: vtole = 10 ' si algun es zero posso valors per defecte
   End If
   Set rst2 = Nothing
End Sub
Function calcular_espesorteorica_material(vnumc As Double, vesptinta As Double, vespcola As Double, vesptolerancia As Double) As Double
   Dim rst As Recordset
   Dim petit As Double
   Dim vmicres As Double
   Dim vmicresmes As Double
   Dim rst2 As Recordset
   Set rst = dbtmp.OpenRecordset("select client,comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
     carregarvalorsespessor rst!client, vesptinta, vespcola, vesptolerancia
     petit = cadbl(vnumc)
     If cadbl(rst!linkcomanda1) < petit And cadbl(rst!linkcomanda1) > 0 Then petit = cadbl(rst!linkcomanda1)
     If cadbl(rst!linkcomanda2) < petit And cadbl(rst!linkcomanda2) > 0 Then petit = cadbl(rst!linkcomanda2)
     Set rst = dbtmp.OpenRecordset("SELECT mesuraesp,tubolam,comandes.espessor,comandes.comanda, comandes.linkcomanda1, comandes.linkcomanda2, productes.ruta FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi Where comanda = " + atrim(vnumc))
     vmicres = micresmaterial(cadbl(rst!mesuraesp), rst!espessor, atrim(rst!tubolam))
     If InStr(1, rst!ruta, "I") > 0 Then vmicresmes = vesptinta
     If vmicres < 0 Then vmicres = vmicres * -1
     If cadbl(rst!linkcomanda1) > 0 Then
        Set rst2 = dbtmp.OpenRecordset("SELECT mesuraesp,tubolam,comandes.espessor,comandes.comanda, comandes.linkcomanda1, comandes.linkcomanda2, productes.ruta FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi Where comanda = " + atrim(rst!linkcomanda1))
        vmicres = vmicres + micresmaterial(cadbl(rst2!mesuraesp), rst2!espessor, atrim(rst2!tubolam))
        vmicresmes = vmicresmes + vespcola
     End If
     If vmicres < 0 Then vmicres = vmicres * -1
     If cadbl(rst!linkcomanda2) > 0 Then
       Set rst2 = dbtmp.OpenRecordset("SELECT mesuraesp,tubolam,comandes.espessor,comandes.comanda, comandes.linkcomanda1, comandes.linkcomanda2, productes.ruta FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi Where comanda = " + atrim(rst!linkcomanda2))
       vmicres = vmicres + micresmaterial(cadbl(rst2!mesuraesp), rst2!espessor, atrim(rst2!tubolam))
       vmicresmes = vmicresmes + vespcola
     End If
     If vmicres < 0 Then vmicres = vmicres * -1
   End If
   calcular_espesorteorica_material = vmicres + vmicresmes
   Set rst = Nothing
   Set rst2 = Nothing
End Function
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
Function estriplex(numc As Double) As Boolean
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select linkcomanda2 from comandes where comanda=" + atrim(numc), , ReadOnly)
  If Not rst.EOF Then If cadbl(rst!linkcomanda2) > 0 Then estriplex = True
  Set rst = Nothing
End Function
Sub imprimir_controlqualitat(numc As Double, op As Byte, numbob As Double, vdata As Date)
   Dim ultimalinia As String
   Dim vmicresmaterial As Integer
   Dim vcont As Byte
   Dim vesp As Double
   Dim vtanx100 As Double
   Dim vesptolerancia As Double
   Dim vesptinta As Double
   Dim vespcola As Double
   Dim v As String
   Dim vcomprovarmicres As Boolean
   Dim pregunta As String

   vcont = 0
   vestriplex = estriplex(numc)
   If primerproces.Visible = True And Not vestriplex Then vcomprovarmicres = True
   If vestriplex And primerproces.Visible = False Then vcomprovarmicres = True
   If vcomprovarmicres Then
     vesp = calcular_espesorteorica_material(numc, vesptinta, vespcola, vesptolerancia)
     While vmicresmaterial = 0
       vtanx100 = (vesp * (vesptolerancia / 100))
       pregunta = "Min:" + atrim(Redondejar(vesp - vtanx100, 0)) + " Max:" + atrim(Redondejar(vesp + vtanx100, 0))
       If nummaq = 3 Then
           vmicresmaterial = demanar_valor_micrometre(pregunta, "")
            Else: vmicresmaterial = cadbl(InputBox("Comprova les micres del material laminat." + Chr(10) + "Escriu les micres del material.", "Control Micres de sortida"))
       End If
       
       If vmicresmaterial <> 0 Then
            If vmicresmaterial < Redondejar(vesp - vtanx100, 0) Or vmicresmaterial > Redondejar(vesp + vtanx100, 0) Then
                MsgBox "Aquest valor no està dins del marge correcte de tolerancia." + Chr(10) + "Min: " + atrim(Redondejar(vesp - vtanx100, 0)) + "    Max: " + atrim(Redondejar(vesp + vtanx100, 0)), vbCritical, "Error"
                vmicresmaterial = 0
            End If
       End If
       vcont = vcont + 1
       If vcont = 5 Then Exit Sub
     Wend
   End If
   ultimalinia = "Lam-" + atrim(nummaq) + " Op: " + atrim(op) + " NºBob.Salida: " + atrim(numbob) + " Fecha: " + Format(vdata, "dd/mm/yy")
   For i = 0 To 20
     llistat.Formulas(i) = ""
   Next i
   llistat.DataFiles(0) = ""
   llistat.DataFiles(1) = ""
   id = " +"
'   llistat.SelectionFormula = "{bobinesentlam.id}=" + atrim(cadbl(bobines.Recordset!id)) + " and {bobinesentlam.paletobobina}='p'"
   If vmicresmaterial > 0 Then
      v = "'Marge: >" + atrim(Redondejar(vesp - vtanx100, 0)) + " i <" + atrim(Redondejar(vesp + vtanx100, 0)) + "'"
       Else: v = ""
   End If
   llistat.Formulas(0) = "lot=" + atrim(numc)
   llistat.Formulas(1) = "ultimalinia='" + atrim(ultimalinia) + "'"
   llistat.Formulas(2) = "nummaq='" + atrim(nummaq) + "'"
   llistat.Formulas(3) = "micres=" + atrim(vmicresmaterial)
   llistat.Formulas(4) = "valorvalidespesor=" + v
   calcularvalorsreducciocilindre comanda.Tag, nummaq, 5   'comanda.tag hauria de ser la comanda impresa
   llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "verificacioqualitatlaminadores.rpt"
   llistat.Destination = crptToPrinter
    llistat.CopiesToPrinter = 1
   llistat.DiscardSavedData = True
' llistat.PrinterName = llegir_ini("Impressores", "nomfulla", "baixesimpressora.ini")
' llistat.PrinterPort = llegir_ini("Impressores", "portfulla", "baixesimpressora.ini")
' llistat.PrinterDriver = llegir_ini("Impressores", "driverfulla", "baixesimpressora.ini")
   DoEvents
   If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
   llistat.Action = 1
   MsgBox "ATENCIÓ CONTROL DE VERIFICACIÓ DE QUALITAT." + Chr(10) + "VERIFICA L'ETIQUETA IMPRESA", vbInformation, "VERIFICACIÓ QUALITAT"
   llistat.SelectionFormula = ""
   llistat.DataFiles(0) = ""
End Sub
Function maquinaquehaimpres(numc As Double) As Byte
   Dim rst As Recordset
   maquinaquehaimpres = 0
   Set rst = dbtmpb.OpenRecordset("select * from impressores where comanda=" + atrim(numc))
   If Not rst.EOF Then maquinaquehaimpres = cadbl(rst!numeromaquina)
   
End Function
Sub calcularvalorsreducciocilindre(numc As Double, ByVal numerodemaquina As Byte, numformula As Byte)
   Dim rstc As Recordset
   Dim rstclixes As Recordset
   Dim dbclixes As Database
   Dim rstmodifi As Recordset
   Dim desarrollteoric As Double
   Dim desarrollreal As Double
   Dim valorrealmostra As Double
   Dim motius As Double
   numerodemaquina = maquinaquehaimpres(numc)
   If numerodemaquina < 7 Then Exit Sub
   Set rstc = dbtmp.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(numc))
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
   If numformula = 0 Then GoTo fi
   llistat.Formulas(numformula) = "reducciopermetrelineal=" + passaradecimalpunt(atrim(rstclixes!reduccioxmetre))
   numformula = numformula + 1
   llistat.Formulas(numformula) = "parametrereduccio=" + passaradecimalpunt(atrim((IIf(numerodemaquina = 7, rstclixes!redcilindrefw, rstclixes!redcilindref2))))
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
Sub imprimir_controlqualitatPE(numc As Double, op As Byte, palet As Double, bobina)
   Dim ultimalinia As String
   ultimalinia = "Op: " + atrim(op) + "  Bob.Entrada: " + atrim(palet) + "/" + atrim(bobina) + "   Fecha: " + Format(Now, "dd/mm/yy")
   For i = 0 To 100
     llistat.Formulas(i) = ""
   Next i
   llistat.Formulas(0) = "lot=" + atrim(numc)
   llistat.Formulas(1) = "ultimalinia='" + atrim(ultimalinia) + "'"
   llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "verificacioqualitatlaminadoresPE.rpt"
   llistat.Destination = crptToPrinter
    llistat.CopiesToPrinter = 1
   llistat.DataFiles(0) = ""
   llistat.DiscardSavedData = True
' llistat.PrinterName = llegir_ini("Impressores", "nomfulla", "baixesimpressora.ini")
' llistat.PrinterPort = llegir_ini("Impressores", "portfulla", "baixesimpressora.ini")
' llistat.PrinterDriver = llegir_ini("Impressores", "driverfulla", "baixesimpressora.ini")
   DoEvents
   If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
   llistat.Action = 1
End Sub

Private Sub comanda_GotFocus()
   Dim vnumc As String
  vnumc = cadbl(InputBox("Entra la nova comanda", "Comanda"))
  If cadbl(vnumc) > 0 Then
     comanda = vnumc
     command4_click
     aviscomprovarcomplexamesimpresionormal comanda
  End If
End Sub

Private Sub comanda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then reixa.SetFocus
End Sub

Private Sub comanda_LostFocus()
   escriure_ini "Baixes", "ultimacomanda", comanda, "comandes.ini"
  ' command4_click
End Sub

Private Sub Command1_Click()
  carregar_capcalera
End Sub
Sub carregar_capcalera(Optional ensenyar As Boolean = True)
    Load capcalera
    capcalera.capcalera.DatabaseName = laminadores.DatabaseName
    capcalera.capcalera.RecordSource = "select * from laminadorestot where comanda=" + atrim(cadbl(comanda))
    capcalera.capcalera.Refresh
    If capcalera.capcalera.Recordset.EOF Then
        capcalera.capcalera.Recordset.AddNew
        capcalera.capcalera.Recordset!comanda = cadbl(comanda)
        capcalera.capcalera.Recordset.Update
    End If
    capcalera.capcalera.Refresh
    If AcroPDF1.Visible Then AcroPDF1.Visible = False: AcroPDF1.Visible = True
    tamany_visualitzadorpdf True
    If ensenyar Then
       capcalera.capcalera.Recordset.Edit
       capcalera.Show 1
        Else
         'capcalera.capcalera.Recordset.Edit
         capcalera.Show 0
    End If
    If form1.laminadores.Recordset.EOF And form1.laminadores.Recordset.BOF Then Command2.SetFocus: Command2_Click
    reixa.col = 5
    reixa.SetFocus
    If AcroPDF1.Visible Then AcroPDF1.Visible = False: AcroPDF1.Visible = True

End Sub


Private Sub Command10_Click()
Dim i As Double
client.ToolTipText = client.Caption

If Not laminadores.Recordset.EOF Then
    If Not IsDate(atrim(laminadores.Recordset!datafi)) Or Not IsDate(atrim(laminadores.Recordset!horafi)) Then
       If MsgBox("No hi ha la hora de fi de funcionament, Vols que el col.loqui automàticament?", vbInformation + vbYesNo, "Atenció") = vbYes Then
        laminadores.Recordset.Edit
        laminadores.Recordset!datafi = Date
        laminadores.Recordset!horafi = Time
        laminadores.Recordset.Update
       End If
    End If
End If
'mirar_bobinesdentrada_noacavades   'ho he tret per veure si es aixó que marca les bobines com acabades
command4_click: capcalera.Hide
comandaacavada.Value = 0
guarda_totals
Command8_Click
i = cadbl(InputBox("Entra la nova comanda", "Canvi de comanda"))
If i > 0 Then comanda.Text = i: command4_click
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
  Dim r2 As Double
If MsgBox("Eliminar aquesta pot suposar eliminar informació de bobines.", vbCritical + vbYesNo, "Atenció") = vbYes Then
     'reixa_BeforeDelete 0
     If MsgBox("Segur que vols borrar aquesta linia i tot el seu contingut?", vbYesNo, "Atenció") = vbNo Then Cancel = 1
    If Cancel <> 1 Then
    r = 0
    r2 = cadbl(laminadores.Recordset!id)
    If atrim(laminadores.Recordset!tipus) = "F" Then r = atrim(cadbl(laminadores.Recordset!id))
    dbtmpb.Execute "delete * from bobineslam where controlid=" + r
'    dbtmpb.Execute "delete * from impressores where id=" + atrim(r2)
    laminadores.Recordset.Delete
    command4_click
    laminadores.Recordset.MoveLast
  End If
End If
End Sub

Function quedenbobinesentrada() As Boolean
   Dim rsttmp2 As Recordset
   quedenbobinesentrada = False
   
   r = "carregartaulatmp2": bobentrada_DblClick
   Set rsttmp2 = dbtmpb.OpenRecordset("select * from bobentradatmplam")
   r = ""
   While Not rsttmp2.EOF
     quedenbobinesentrada = True
     r = r + " " + atrim(rsttmp2!numpalet) + "/" + atrim(rsttmp2!numbobent)
     rsttmp2.MoveNext
   Wend
   Set rsttmp2 = Nothing
End Function
Sub mirar_bobinesdentrada_noacavades()
 Dim metres As Double
 Dim metresant As Double
 Dim palet As Double
 Dim bobina As Double
 Dim rstconsulta2 As Recordset
 noespota0 = True
   carregar_bobinesdentrada "carregarbobinesnoutilitzades", , , , ncomanda, , ncomanda2, IIf(primerproces.Tag = "invertit", True, False)
   If Not rstconsulta.EOF Or Not rstconsulta.BOF Then rstconsulta.MoveFirst
   Set rstconsulta2 = rstconsulta.Clone
   mantenimentbobina.checknoimprimirparcial = 1
   While Not rstconsulta2.EOF
      On Error Resume Next
      palet = rstconsulta2!idpalet
      bobina = rstconsulta2!idbobina
      PoB = IIf(rstconsulta2!taula = "parcials", "p", "b")
      On Error GoTo 0
      If palet > 0 And bobina > 0 And UCase(PoB) = "P" Then
         'es una bobina d'estock
         
         estatdelabobina palet, bobina, 0, ncomanda, ncomanda2
         
         metres = bobinesdentrada.calcular_mtrsdispreals(palet, bobina)
         If metres > 0 Then ajustar_diametre_real atrim(palet) + "/" + atrim(bobina)
         Else
            'es una bobina feta a inplacsa
               carregar_bobinesdentrada "marcarutilitzadademanar", , palet, bobina, ncomanda, True, ncomanda2
      End If
      rstconsulta2.MoveNext
   Wend
   noespota0 = False
End Sub
Sub separarpaletibobina(vnumbob As String, vpalet As String, vbob As String)
    If vnumbob = "" Then Exit Sub
    If InStr(1, vnumbob, "/") = 0 Then Exit Sub
    vpalet = cadbl(Mid(vnumbob, 1, InStr(1, vnumbob, "/") - 1))
    vbob = cadbl(substituir(vnumbob, vpalet + "/", ""))
End Sub

Sub ajustar_diametre_real(vbobina As String)
   Dim vpalet As String
   Dim vbob As String
   Dim vdiametrenou As String
   Dim vmetresanteriors As Double
   Dim vmetres As Double
   Dim rstbob As Recordset
   Set dbstocks = dbtmpb
   separarpaletibobina vbobina, vpalet, vbob
   If cadbl(vpalet) = 0 Or cadbl(vbob) = 0 Then Exit Sub
   Set rstbob = dbstocks.OpenRecordset("select * from bobines where idbobina=" + atrim(vbobina) + " and idpalet=" + atrim(vpalet))
   If rstbob.EOF Then MsgBox "No he trobat la bobina " + vbobina, vbCritical, "Atenció": Exit Sub
   vdiametrenou = InputBox("Entra el diametre actual de la bobina " + vpalet + "/" + vbob, "Nou diametre")
   If cadbl(vdiametrenou) = 0 Then Exit Sub
   dbstocks.Execute "delete * from parcials where comanda='444' and idpalet=" + atrim(cadbl(vpalet)) + " and idbobina=" + atrim(cadbl(vbob))
   vmetresanteriors = bobinesdentrada.calcular_mtrsdispreals(cadbl(vpalet), cadbl(vbob))
   vmetres = Redondejar(calcular_metresambdiametre(cadbl(vpalet), cadbl(vbob), cadbl(vdiametrenou)), cadbl(rstbob!tamanycanutu))
   If vmetres <> 0 Then
       actualitzar_metresxrdiametre vpalet, vbob, vmetresanteriors, vmetres
       If vmetres < 500 Then mantenimentbobina.passarbobinaaacavada cadbl(vpalet), cadbl(vbob)
       wait 2
   End If
   bobinesdentrada.imprimir_bobinaparcial cadbl(vpalet), cadbl(vbob), , 1
End Sub
Sub actualitzar_metresxrdiametre(vpalet As String, vbob As String, vmetresanteriors As Double, vmetresnous As Double)
    Dim rst As Recordset
    Dim vmetresbob As Double
    Dim vValues As String
    Dim vmetresactualitzar As Double
   ' vmetresbob = bobinesdentrada.calcular_mtrsdispreals(cadbl(vpalet), cadbl(vbob))
    vmetresactualitzar = Redondejar(vmetresanteriors - vmetresnous, 0)
    dbstocks.Execute "delete * from parcials where comanda='444' and idpalet=" + atrim(cadbl(vpalet)) + " and idbobina=" + atrim(cadbl(vbob))
    vValues = "(" + atrim(vpalet) + "," + atrim(vbob) + ",True,'444',now," + atrim(cadbl(numop)) + ",'L','Actualització metres per diametre.')"
    dbstocks.Execute "insert into parcials (idpalet,idbobina,utilitzada,comanda,data,operari,seccio,observacions) values " + vValues
    Set rst = dbstocks.OpenRecordset("select * from parcials where comanda='444' and idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbob))
    If Not rst.EOF Then
       rst.Edit: rst!metres = vmetresactualitzar: rst.Update
       bobinesdentrada.actualitzar_metres_disponibles cadbl(vpalet), cadbl(vbob)
    End If
    Set rst = Nothing
End Sub
Function calcular_metresambdiametre(palet As Double, bobina As Double, vdiametre As Double, Optional canutu As Double) As Double
     Dim rstp As Recordset
  Dim rstb As Recordset
  Dim metres As Double
  Dim micres As Double
  Dim diametre As Double
  Dim pi As Double
  If cadbl(canutu) = 0 Then canutu = 15.2
  If canutu < 10 Then canutu = canutu + 2 'afegeixo l'amplada del cartrò del canutu
  If canutu >= 10 Then canutu = canutu + 2.8 'afegeixo l'amplada del cartrò del canutu
  '3,1416*(Diametro maximo^2-Diametro corazon^2)/(4*Espesor)
  
  Set rstp = dbstocks.OpenRecordset("select micres,grmsm2,codimatprognou from palets where idpalet=" + atrim(palet))
  If rstp.EOF Then GoTo fi
  Set rstb = dbstocks.OpenRecordset("select * from materials where codi=" + atrim(rstp!codimatprognou))
  If Not rstp.EOF Then
    pi = 4 * Atn(1)
    vdiametre = vdiametre / 100
    canutu = canutu / 100
    micres = cadbl(rstp!micres)
    If micres = 0 Then micres = cadbl(rstb!micresdelsgrm2)
    If micres = 0 Then GoTo fi
    micres = (micres * 0.0001) / 100
    diametre = (((vdiametre * vdiametre) - (canutu * canutu)) * pi) / (4 * micres)
    'diametre = Sqr(((metres * micres) / pi) + (canutu * canutu)) * 200
    calcular_metresambdiametre = Redondejar(diametre, 0)
    'If cadbl(calcular_metresambdiametre) < 9 Then calcular_metresambdiametre = "0"
  End If
fi:
  Set rstp = Nothing
  Set rstb = Nothing
End Function



Function metresfetsinferiorsacomanda(numc As Double) As Boolean
   Dim metresc As Double
   If cadbl(tmetres) < (cadbl(tmetres.Tag) - ((cadbl(tmetres.Tag) / 100) * 4)) Then
          If UCase(InputBox("Aquesta comanda es de " + tmetres.Tag + " metres i tu has fet " + tmetres + " metres" + Chr(10) + "PASSARÉ LA COMANDA A NO ACABADA. ESCRIU ACABADA SI ESTÀ REALMENT ACABADA", "ATENCIÓ")) = "ACABADA" Then
              metresfetsinferiorsacomanda = False
               Else: metresfetsinferiorsacomanda = True
          End If
   End If
End Function

Function elslitrosdeadhesiuestanentrats() As Boolean
  If (cadbl(kbpantone(0)) = 0 And pantone(0) <> "-") Or (cadbl(kbpantone(1)) = 0 And pantone(1) <> "-") Then
      MsgBox "Falta entrar els Kilos d'adhesiu utilitzat.", vbCritical, "Error"
      elslitrosdeadhesiuestanentrats = False
        Else: elslitrosdeadhesiuestanentrats = True
  End If
End Function
Function MetrosPackinglistInferiorsaLaminats() As Boolean
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select comanda,ruta from comandesmesextres where comanda=" + atrim(cadbl(comanda)) + " or comanda=" + atrim(cadbl(linkcomanda)))
  While Not rst.EOF
    If InStr(1, atrim(rst!ruta), "L") Then vnumc = cadbl(rst!comanda)
    rst.MoveNext
  Wend
  If vnumc = 0 Then GoTo fi
  Set rst = dbtmpb.OpenRecordset("Select sum(metres) as Tmetres from parcials where comanda='" + atrim(vnumc) + "'")
  If cadbl(rst!tmetres) < cadbl(tmetres) Then MetrosPackinglistInferiorsaLaminats = True
fi:
  Set rst = Nothing
End Function
Private Sub Command15_Click()
Dim com As Double

'If quedenbobinesentrada Then
'   If MsgBox("Hi ha bobines d'entrada pendents de laminar." + Chr$(10) + r + Chr$(10) + "Vols acabar igualment?", vbCritical + vbYesNo, "Atenció") = vbNo Then
'        ratoli "normal"
'        Exit Sub
'   End If
'End If
If MetrosPackinglistInferiorsaLaminats Then If MsgBox("Hi ha menys metres anotats al packinglist que laminats." + vbNewLine + "Vols continuar igualment?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
If Not elslitrosdeadhesiuestanentrats Then Exit Sub
client.ToolTipText = client.Caption
'If comprovarsifaltencamps Then Exit Sub
ratoli "espera"
laminadores.Recordset.MoveLast
If Not laminadores.Recordset.EOF Then
    If Not IsDate(atrim(laminadores.Recordset!datafi)) Or Not IsDate(atrim(laminadores.Recordset!horafi)) Then
        laminadores.Recordset.Edit
        laminadores.Recordset!datafi = Date
        laminadores.Recordset!horafi = Time
        laminadores.Recordset.Update
        wait 1
        
    End If
End If
calcular_totals
wait 1
mirar_bobinesdentrada_noacavades
If metresfetsinferiorsacomanda(cadbl(comanda)) Then Command10_Click: Exit Sub
passar_comanda_a_acavada
If comandaacavada.Value = 0 Then restarkiloscolaalscontenidors
comandaacavada.Value = 1
calcular_totals
guarda_totals

Command15.Tag = "imprimint"
command4_click
capcalera.Hide
Command15.Tag = ""
ratoli "espera"
wait (3)
If cadbl(kiloshora) = 0 Then
  command4_click
  capcalera.Hide
  wait (3)
End If
verificacio_netejaidespeje
imprimir_fulla
If stockopacking = "E" Then imprimir_packinglist ncomanda2
'Command8_Click
wait (3)
ratoli "normal"
com = cadbl(InputBox("Entra la nova comanda", "Fi de comanda"))
If com = 0 Then Exit Sub
comanda.Text = atrim(com)
ratoli "espera"
command4_click
ratoli "normal"
If cadbl(comanda.Text) = 0 Then Exit Sub
'trentats = InputBox("Quants tinters has rentat?", "Nova Comanda")
'pclixers = InputBox("Quants portaclixers?", "Nova Comanda")
'canvienfilada = InputBox("Has fet canvi d'enfilada?   S o N ", "Nova Comanda", "N")
'If Mid(canvienfilada, 1, 1) = "N" Then
'   canvienfilada = "No"
'    Else: canvienfilada = "Si"
'End If
guarda_totals

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
Sub restarkiloscolaalscontenidors()
   'compantone 0 i 1
   'kbpantone 0 i 1
   
End Sub
Sub imprimir_packinglist(numc As Double)
   If numc < 100000 Then Exit Sub
   Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "palets.exe comandes.ini " + atrim(numc), vbNormalFocus
End Sub

Sub passar_comanda_a_acavada()


laminadores.Recordset.MoveLast
  'posso la data als totals de seccio
  If IsDate(laminadores.Recordset!datafi) Then
   dbtmpb.Execute "update laminadorestot set datalaminacio=#" + Format(laminadores.Recordset!datafi, "yy/mm/dd") + "# where comanda=" + atrim(cadbl(comanda))
   dbtmpb.Execute "update laminadorestot set operari=" + atrim(cadbl(laminadores.Recordset!operari)) + " where comanda=" + atrim(cadbl(comanda))
   dbtmpb.Execute "update laminadorestot set laminadora=" + atrim(cadbl(laminadores.Recordset!numeromaquina)) + " where comanda=" + atrim(cadbl(comanda))
  End If

'si hi ha alguna bobina passo l'estat de la comanda a la proxima seccio
   'passo l'estat de comanda a la proxima
   proximaseccio ncomanda, False
   wait 2
   proximaseccio ncomanda2, False
   
End Sub
Function proximaseccio(numc As Double, nogravarcanvi As Boolean) As String
Dim estat As String
Dim ruta As String

   Set rsttmp = dbtmp.OpenRecordset("select producte,proximaseccio from comandes where comanda=" + atrim(numc))
   If Not rsttmp.EOF Then
     estat = atrim(rsttmp!proximaseccio)
     If estat = "" Then estat = "V"
   End If
   Set rsttmp = dbtmp.OpenRecordset("select ruta from productes where codi='" + rsttmp!producte + "'")
   If Not rsttmp.EOF Then ruta = rsttmp!ruta + "   "
   If InStr(1, "EIL", estat) > 0 Then
     seccio = Mid(ruta, InStr(1, ruta, "L") + 1, 1)
     If Trim(seccio) = "" Or Trim(ruta) = "E" Then seccio = "V"
     If Not nogravarcanvi Then
       dbtmp.Execute "update comandes set proximaseccio='" + seccio + "' where comanda=" + atrim(numc)
       dbtmp.Execute "update comandes set seccioactual='" + seccio + "' where comanda=" + atrim(numc)
     End If
     proximaseccio = seccio
   End If
   Set rsttmp = Nothing
   
End Function
Function comprovarsifaltencamps() As Boolean
  Dim faltenpatones As Boolean
  Dim faltenmtrs As Boolean
 ' For i = 0 To 9
 '   If atrim(pantone(i)) <> "" And cadbl(kbpantone(i)) = 0 Then
 '      faltenpantones = True
 '   End If
 ' Next i
  laminadores.Recordset.FindFirst "tipus='F'"
  While Not laminadores.Recordset.NoMatch
    If cadbl(laminadores.Recordset!mtrsminut) = 0 Then
      laminadores.Recordset.Edit
      laminadores.Recordset!mtrsminut = InputBox("Falten els Mtrs/Min.", "Atenció")
      laminadores.Recordset.Update
    End If
    laminadores.Recordset.FindNext "tipus='F'"
  Wend
  'If faltenpantones Then MsgBox "Falta entrar els Kg de tinta.": comprovarsifaltencamps = True
  
End Function

Private Sub Command16_Click()
  kbpantone(0) = cadbl(kbpantone(0)) + cadbl(kbpantone(2))
  kbpantone(1) = cadbl(kbpantone(1)) + cadbl(kbpantone(3))
End Sub

Private Sub Command17_Click()
   If UCase(InputBoxEx("Escriu la contrasenya d'Encarregat.", "Atenció", , , , , , SPassword)) <> "INPLACSA" Then Exit Sub
   
   formencarregat.Show 1
   
End Sub



Private Sub Command18_Click()
   Dim vnomcola As String
   Dim vnomenduridor As String
   escullircola vnomcola, vnomenduridor, cadbl(pantone(0).Tag), cadbl(pantone(1).Tag)
   If vnomcola <> "" Then
     pantone(0) = vnomcola
     pantone(1) = vnomenduridor
   End If
End Sub
Sub escullircola(vnomcola As String, vnomenduridor As String, vfam As Double, vsubfam As Double)
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "select resina,enduridor from adhesius where novisiblealaminadora=false and idfamilia=" + atrim(vfam) + " and idsubfamilia=" + atrim(vsubfam)
  formseleccio.Caption = "Selecció d'Adhesiu"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Visible = True
  formseleccio.DBGrid2.Columns(0).Width = 4000
  formseleccio.DBGrid2.Columns(1).Visible = True
  formseleccio.DBGrid2.Columns(1).Width = 4000
  formseleccio.DBGrid2.Font.Size = 10
  formseleccio.Show 1
  If seleccioret = 1 Then
   vnomcola = atrim(formseleccio.Data1.Recordset!resina)
   vnomenduridor = atrim(formseleccio.Data1.Recordset!enduridor)
  End If
  Unload formseleccio
End Sub

Private Sub Command19_Click()
   Dim vnumlot As String
   If atrim(compantone(0)) = "" Then vnumlot = llegir_ini("Laminadores", "lot1", "comandes.ini")
   vnumlot = InputBox("Escaneja o entra el numero de Lot de la Resina.", "Lot de la Resina", vnumlot)
   If vnumlot <> "" Then
     If Not esunlotvalid(vnumlot, "R") Then
        MsgBox "Aquest Lot no es vàlid o el contenidor ja està donat de baixa.", vbCritical, "Error": GoTo fi
     End If
     compantone(0) = vnumlot
     If compantone(1) = "" Then Command22_Click
   End If
fi:
End Sub

Private Sub Command2_Click()
  If comprovarsidescansorelleu Then Exit Sub
  aviscomprovarcomplexamesimpresionormal comanda
 If Not laminadores.Recordset.EOF Then
  laminadores.Recordset.MoveLast
  If laminadores.Recordset!tipus = "C" Then
      numop = escullir_operari
      nomoperari = UCase(r)
  End If
 End If
 crearseccio "C"
 reixa.SetFocus
End Sub
Sub aviscomprovarcomplexamesimpresionormal(numc As Double)
 Dim rstc As Recordset
 Dim rstclixes As Recordset
 Dim dbclixes As Database
 Set rstc = dbtmp.OpenRecordset("Select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(numc))
 If rstc.EOF Then GoTo fi
 If rstc!linkcomanda1 = 0 And rstc!linkcomanda2 = 0 Then GoTo fi
 Set rstc = dbtmp.OpenRecordset("Select producte,numtreball,numordremodificacio from comandes where comanda>0 and (comanda=" + atrim(rstc!comanda) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2) + ") order by comanda")
 If InStr(1, rstc!producte, "PC") = 0 Then
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
   Set rstclixes = dbclixes.OpenRecordset("select formaimpresio from modificacions where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordre=" + atrim(cadbl(rstc!numordremodificacio)))
   If rstclixes.EOF Then GoTo fi
   If atrim(rstclixes!formaimpresio) = "N" Then MsgBox "ULL VIGILAR MATERIAL BICAPA O TRICAPA AMB IMPRESSIÓ NORMAL, VIGILAR CARES A LAMINAR", vbExclamation, "Atenció"
 End If
fi:
 Set rstc = Nothing
 Set rstclixes = Nothing
 Set dbclixes = Nothing
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
  If Not laminadores.Recordset.EOF Then
      canvicamisa = atrim(laminadores.Recordset!canvicamisa)
      finalitza_seccio
      com = cadbl(laminadores.Recordset!comanda)
  End If
  r = ""
  If com = 0 Then Exit Sub
  laminadores.Recordset.AddNew
  laminadores.Recordset!comanda = com
  laminadores.Recordset!numeromaquina = nummaq
  laminadores.Recordset!operari = numop
  laminadores.Recordset!tipus = tipus
  laminadores.Recordset!datainici = Date
  laminadores.Recordset!horainici = Time
  If tipus = "C" Then
     laminadores.Recordset!canvicamisa = canvicamisa
  End If
  'laminadores.Recordset!texteimpresio = rsttmpcs!texteimpressio
  r = laminadores.Recordset!id
  Command3.Tag = r
  laminadores.Recordset.Update
  laminadores.Recordset.MoveLast
     Set rsttmpcs = Nothing
     
End Sub

Private Sub Command20_Click()
' Dim desb As Byte
Dim palet As Double
  Dim bobina As Double
  Dim rst As Recordset
  Dim inssql As String
  Dim jaexisteix As Boolean
  Dim numc As Double
  Dim utili As Boolean
  demanar_paletibobina palet, bobina, desb
  
  numc = ncomanda2
  If palet > 0 And bobina > 0 Then
    obrestocks
    'inssql = "SELECT CDbl([comanda]) AS Expr1, Parcials.idpalet, Parcials.idbobina From Parcials WHERE (((CDbl([comanda]))<10000) and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + ");"
    'Set rst = dbstocks.OpenRecordset(inssql)
    'If rst.EOF Then
    ' inssql = "select * from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and comanda='" + atrim(numc) + "'"
    ' Set rst = dbstocks.OpenRecordset(inssql)
    'End If
    'If rst.EOF Then
    '  MsgBox "El Palet: " + atrim(palet) + "/" + atrim(bobina) + " no està assignat per utilitzar-lo.", vbCritical, "Palet/Bobina equivocat"
    ' Else
    '   carregar_bobinesdentrada "mirarsiutilitzada", , palet, bobina, ncomanda, utili, ncomanda2
    '   If utili Then
    '      MsgBox "Aquesta bobina ja està marcada com utilitzada.", vbInformation + vbOKOnly, "bobina utilitzada"
    '       Else
    '        'demanar si es final de bobina
     '       carregar_bobinesdentrada "marcarutilitzadademanar", , palet, bobina, ncomanda, False, ncomanda2
          If comprovarsiesdelestoccorrecte(palet, bobina, cadbl(numgrup.Tag)) Then
            afegir_labobinadentrada palet, bobina, desb, IIf(primerproces.Tag = "invertit", True, False)
            
            If palet < (cadbl(comanda) - 3) Or palet > (cadbl(comanda) + 3) Then
                'mirar_imprimir_controlqualitatPE palet, bobina
                    bobinesent.Recordset.FindFirst "palet=" + atrim(palet) + " and bobina=" + atrim(bobina)
                    demanar_verificacio_espesoritractat palet, bobina
                    imprimir_controlqualitatbobinaentrada palet, bobina, desb
            End If
            dbstocks.Execute "insert into parcials (idpalet,idbobina,metres,comanda,orcomassignacio) values (" + atrim(palet) + "," + atrim(bobina) + ",0," + atrim(cadbl(ncomanda2)) + "," + atrim(cadbl(ncomanda2)) + ")"
                 Else: MsgBox "Aquesta bobina no es del grup " + numgrup.Tag, vbCritical, "Error"
          End If
          
      ' End If
    'End If
  End If
 ratoli "normal"
 
End Sub
Function comprovarsiesdelestoccorrecte(palet As Double, bobina As Double, grup As Double) As Boolean
   Dim rstp As Recordset
    If grup < 2000 Then Exit Function
    Set rstp = dbstocks.OpenRecordset("select * from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and cdbl(comanda)=" + atrim(grup))
    If rstp.EOF Then
      comprovarsiesdelestoccorrecte = False
        Else: comprovarsiesdelestoccorrecte = True
    End If
End Function
Function comprovarsidescansorelleu() As Boolean
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("select * from controldescansrelleu where (hores=0 or hores=null) and nummaq=" + atrim(nummaq) + " and operari=" + atrim(numop) + " and seccio='" + atrim(lletraseccio) + "'")
  If rst.EOF Then Exit Function
  comprovarsidescansorelleu = True
  MsgBox UCase(nomoperari) + " en aquest moment està fent " + atrim(rst!tipus) + Chr(10) + "Primer dona per acabada la incidència.", vbExclamation, "Atenció"
End Function

Private Sub Command21_Click()
  Load calculdiametre
  calculdiametre.micres = micrescomanda
  
  calculdiametre.Show 1
End Sub

Private Sub Command22_Click()
   Dim vnumlot As String
   'Dim vnumlotanterior As String
   If atrim(compantone(1)) = "" Then vnumlot = llegir_ini("Laminadores", "lot2", "comandes.ini")
   'vnumlotanterior = compantone(1)
   vnumlot = InputBox("Escaneja o entra el numero de Lot de l'Enduridor.", "Lot del Enduridor", vnumlot)
   If vnumlot <> "" Then
     If Not esunlotvalid(vnumlot, "E") Then
        MsgBox "Aquest Lot no es vàlid o el contenidor ja està donat de baixa.", vbCritical, "Error": GoTo fi
     End If
     compantone(1) = vnumlot
     'If vnumlotanterior <> vnumlot Then
     '    If MsgBox("Vols donar de baixa el lot anterior? " + vnumlotanterior, vbExclamation + vbDefaultButton2 + vbYesNo, "Baixa numero de lot") = vbYes Then
     '        donardebaixalallauna vnumlotanterior
     '    End If
     'End If
   End If
fi:
End Sub

Function esunlotvalid(vnumlot As String, Optional vResinaoEnduridor As String) As Boolean
  Dim rst As Recordset
  esunlotvalid = True
  If Len(vnumlot) > 4 And Len(vnumlot) < 8 Then
     If Mid(UCase(atrim(vnumlot)), 1, 1) = "A" Then
        Set rst = dbtmpb.OpenRecordset("select activa from llaunes where numllauna='" + atrim(vnumlot) + "'")
        If Not rst.EOF Then
           If Not rst!activa Then esunlotvalid = False
             Else: esunlotvalid = False
        End If
        If vResinaoEnduridor <> "" Then
            Set rst = dbtmpb.OpenRecordset("SELECT Llaunes.numllauna, familiestintes.descripcio as nomfamilia FROM (Llaunes INNER JOIN tintes ON Llaunes.idtinta = tintes.idtinta) INNER JOIN familiestintes ON tintes.idfamilia = familiestintes.codi where NUMllauna='" + vnumlot + "'")
            If rst.EOF Then
                 esunlotvalid = False = False
                Else
                  If vResinaoEnduridor = "R" Then If atrim(rst!nomfamilia) <> "RESINA" Then esunlotvalid = False
                  If vResinaoEnduridor = "E" Then If atrim(rst!nomfamilia) <> "ENDURIDOR" Then esunlotvalid = False
            End If
        End If
     End If
  End If
  Set rst = Nothing
End Function

Private Sub Command23_Click()
  Dim vnumlot As String
  'vnumlot = InputBox("Escaneja o entra el numero de Lot de la Resina per predeterminar.", "Lot de la Resina", vnumlot)
   'If vnumlot <> "" Then
   '  If Not esunlotvalid(vnumlot) Then
   '     MsgBox "Aquest Lot no es vàlid o el contenidor ja està donat de baixa.", vbCritical, "Error": GoTo fi
   '  End If
   
  If MsgBox("Vols possar aquestes dades de cola com a predeterminades?", vbInformation + vbYesNo, "Atenció") = vbYes Then
       If atrim(compantone(0)) <> "" Then gravar_valorxrdefecte_adhesius "C"
   End If
'fi:
End Sub

Private Sub Command24_Click()
  Dim vnumlot As String
  ' vnumlot = InputBox("Escaneja o entra el numero de Lot de l'Enduridor per predeterminar.", "Lot de la Resina", vnumlot)
  ' If vnumlot <> "" Then
  '   If Not esunlotvalid(vnumlot) Then
  '      MsgBox "Aquest Lot no es vàlid o el contenidor ja està donat de baixa.", vbCritical, "Error": GoTo fi
  '   End If
  If MsgBox("Vols possar aquestes dades d'Enduridor com a predeterminades?", vbInformation + vbYesNo, "Atenció") = vbYes Then
     If atrim(compantone(1)) <> "" Then gravar_valorxrdefecte_adhesius "E"
   End If
'fi:
End Sub

Private Sub Command25_Click()
   Dim rst As Recordset
   Dim vmsg As String
   Dim i As Byte
   Set rst = dbtmp.OpenRecordset("select * from adhesius where resina='" + pantone(0) + "' and enduridor='" + pantone(1) + "'")
   If Not rst.EOF Then
       'For i = 5 To 23
       '  If InStr(1, UCase(rst.Fields(i).Name), "EURO") = 0 Then
       '   vmsg = vmsg + Chr(10) + UCase(rst.Fields(i).Name) + " - " + atrim(rst.Fields(i))
       '  End If
       'Next i
       vmsg = vmsg + "Resina: " + atrim(rst![%resina]) + "%" + vbNewLine
       vmsg = vmsg + "Enduridor: " + atrim(rst![%enduridor]) + "%" + vbNewLine
       vmsg = vmsg + "Temperatura R1: " + atrim(rst!tempaigua1) + " ºC" + vbNewLine
       vmsg = vmsg + "Temperatura R2: " + atrim(rst!tempaigua2) + " ºC" + vbNewLine
       vmsg = vmsg + "Temperatura Pre-Heating: " + atrim(rst!temppreheating) + " ºC" + vbNewLine
       vmsg = vmsg + "Temperatura R4: " + atrim(rst!tempaigua4) + " ºC" + vbNewLine
       vmsg = vmsg + "Temperatura Prensa: " + atrim(rst!tempprensa) + " ºC" + vbNewLine
       vmsg = vmsg + "Temperatura Enduridor: " + atrim(rst!tempenduridor) + " ºC" + vbNewLine
       vmsg = vmsg + "Temperatura Manguera: " + atrim(rst!temptubo) + " ºC" + vbNewLine
       vmsg = vmsg + "Temperatura Resina: " + atrim(rst!tempresina) + " ºC" + vbNewLine
       vmsg = vmsg + "Aportació Material EVOH: " + atrim(rst!aportcola_EVOH) + " gr/mt2" + vbNewLine
       vmsg = vmsg + "Aportació Material Anònim: " + atrim(rst!aportcola_Anonim) + " gr/mt2" + vbNewLine
       vmsg = vmsg + "Aportació Material Imprès: " + atrim(rst!aportcola_Impres) + " gr/mt2" + vbNewLine
       vmsg = vmsg + "Aportació Material Tricapa imprès: " + atrim(rst!aportcola_Tricapa_Impres) + " gr/mt2" + vbNewLine
   End If
   If vmsg <> "" Then MsgBox vmsg, vbInformation, "Informació Adhesiu"
End Sub

Private Sub Command26_Click()
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

Private Sub Command29_Click()
    tamany_visualitzadorpdf True
End Sub
Sub tamany_visualitzadorpdf(vtamanygran As Boolean)
     AcroPDF1.Visible = Not AcroPDF1.Visible
     AcroPDF1.Width = 11000
     AcroPDF1.Height = 6500
     AcroPDF1.Left = 700
     AcroPDF1.ZOrder
     framebobentrada.Visible = Not AcroPDF1.Visible
End Sub

Private Sub Command3_Click()
 Dim mtrsprova As String
 Dim mtrsparcials As Double
 Dim opantic As Byte
 Dim idbobina As Long
 idbobina = 0
 If comprovarsidescansorelleu Then Exit Sub
 aviscomprovarcomplexamesimpresionormal comanda
 If Not laminadores.Recordset.EOF Then
    laminadores.Recordset.MoveLast
    If laminadores.Recordset!tipus = "A" Then
        mtrsprova = InputBox("Entra els Metres de prova.", "Atenció")
        laminadores.Recordset.FindLast "tipus='A'"
        If Not laminadores.Recordset.NoMatch Then
         laminadores.Recordset.Edit
         laminadores.Recordset!mtrsprova = cadbl(mtrsprova)
         laminadores.Recordset.Update
        End If
    
    End If
    Else: Exit Sub
 End If
 'firmar_fulla
 If laminadores.Recordset!tipus = "F" Then
 
    opantic = numop
    numop = escullir_operari
    nomoperari = UCase(r)
 End If
 If Not bobines.Recordset.EOF Then
   bobines.Recordset.MoveLast
   If cadbl(bobines.Recordset!metres) = 0 Then
     mtrsprova = InputBox("Entra els metres parcials de bobina.", "Bobina no acabada")
     If cadbl(mtrsprova) <> 0 Then
        mtrsparcials = cadbl(mtrsprova)
        laminadores.Recordset.Edit
        laminadores.Recordset!metresparcial = mtrsparcials
        laminadores.Recordset.Update
        bobines.Recordset.Edit
        bobines.Recordset!metresparcial = mtrsparcials
        bobines.Recordset!operari1 = numop
        bobines.Recordset!operari2 = opantic
        bobines.Recordset.Update
        idbobina = bobines.Recordset!id
        
     End If
   End If
 End If
 Command3.Tag = ""
 crearseccio "F"
 r = Command3.Tag
 If idbobina > 0 Then dbtmpb.Execute "update bobineslam set controlid=" + r + " where id=" + atrim(idbobina)
 mtrsparcials = 0
 laminadores.Recordset.MoveLast
 While bobines.Recordset.RecordCount = 0 And mtrsparcials < 200
   DoEvents
   bobines.Refresh
   mtrsparcials = mtrsparcials + 1
 Wend
 If bobines.Recordset.RecordCount = 0 And idbobina = 0 Then Command5_Click
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
Function hihaseccio(s As String, producte As String) As Boolean
Dim rsttmpp As Recordset
Set rsttmpp = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(producte) + "'")
 If InStr(1, rsttmpp!ruta, s) = 0 Then
    hihaseccio = False
   Else: hihaseccio = True
 End If
 Set rsttmpp = Nothing
End Function
Function comprovarsiesprimer(numc As Double, pc2 As Double, comandaactual As Double) As Boolean
    Dim rst As Recordset
    comprovarsiesprimer = True
    If cadbl(pc2) = 0 Then Exit Function
    Set rst = dbtmp.OpenRecordset("select refilatd from comandes where comanda=" + atrim(numc))
    If rst.EOF Then Exit Function
    If numc = comandaactual Then
       If cadbl(rst!refilatd) <> 1 And Not estacomençada(pc2) Then comprovarsiesprimer = False
    End If
    If pc2 = comanaactual Then
        If cadbl(rst!refilatd) = 1 And Not estacomençada(numc) Then comprovarsiesprimer = False
    End If
End Function
Function estacomençada(numc As Double) As Boolean
   Dim rst As Recordset
   estacomençada = True
   Set rst = dbbaixes.OpenRecordset("select * from laminadores where comanda=" + atrim(numc) + " and tipus='F'")
   If rst.EOF Then estacomençada = False
   Set rst = Nothing
End Function

Sub passarcomandaacomençada()
 dbtmp.Execute "update comandes set seccioactual='I' where comanda=" + atrim(comanda)
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
Function comandavalida(numc As Double, Optional nocomprovarllista As Boolean, Optional vpararcomanda As Boolean) As Boolean
   Dim rst As Recordset
   Dim proximaseccio As String
   comandavalida = False
   If numc = 0 Then Exit Function
   If Not nocomprovarllista Then
     Set rst = dbbaixes.OpenRecordset("select * from muntadora_ordremuntatge where comanda=" + atrim(numc))
     If Not rst.EOF Then MsgBox "La comanda " + atrim(numc) + " ja està a la llista.": Exit Function
   End If
   Set rst = dbtmp.OpenRecordset("SELECT comandes.numtreball,comandes.numordremodificacio,comandes.comanda, productes.ruta, comandes.proximaseccio,comandes.impressio FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(numc) + "));")
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
          comandavalida = False
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

Function nomdelacola(id As Long) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select resina,color,predeterminada from adhesius where codi=" + atrim(cadbl(id)))
   If rst.EOF Then Set rst = dbtmp.OpenRecordset("select resina,color,predeterminada from adhesius where predeterminada<>''")
   If Not rst.EOF Then
      nomdelacola = IIf(UCase(atrim(rst!Color)) = "VERD", atrim(rst!resina), "@ " + atrim(rst!resina))
   End If
   If Mid(nomdelacola, 1, 1) = "@" Then MsgBox "Atenció aquesta comanda porta una cola especial." + Chr(10) + Mid(nomdelacola, 2), vbInformation, "Atenció"
   Set rst = Nothing
End Function
Sub comprovarsiestriplexiavisar()
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select linkcomanda1,linkcomanda2 from comandes where producte='PC' and (comanda=" + atrim(cadbl(comanda)) + " or comanda=" + atrim(cadbl(linkcomanda)) + ")")
   If Not rst.EOF Then
      If cadbl(rst!linkcomanda1) And cadbl(rst!linkcomanda2) Then
       MsgBox "Aquest pedido es de tres capes i ara faras la del material del mig, has d'assegurar que el material es TRACTAT DUES CARES.", vbCritical, "ATENCIÓ"
      End If
   End If
   Set rst = Nothing
End Sub
Private Sub command4_click()
  Dim rst As Recordset
  Dim petit As Double
  Dim resp As String
  Dim vcolor As Double
  Dim vnomcola As String
  Dim vnomenduridor As String
  Dim vpararcomanda As Boolean
  If nummaq = 0 Then maquina_Click
  If nummaq = 0 Then Exit Sub
  primerproces.Visible = False
  reciclarmaterial1.BackColor = &H8000000F
  reciclarmaterial2.BackColor = &H8000000F
  'gravo la ultima comanda
  escriure_ini "Baixes", "ultimacomanda", comanda, "comandes.ini"

  'comprovo si existeix la comanda
  Set rsttmp = dbtmp.OpenRecordset("select refilatd,amplereb,producte,linkcomanda1,linkcomanda2,lotmatdesb1,lotmatdesb2,laminadora,codibarras,espessor,comanda,refclient,comandaclient,texteimpressio,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(cadbl(comanda)))
  If rsttmp.EOF Or cadbl(comanda) = 0 Then
      MsgBox "No hi ha numero de comanda vàlida"
         Command1.Enabled = False:   Command2.Enabled = False:   Command3.Enabled = False: Exit Sub
  End If
  'comprova que no estigui parada la comanda
  If Not comandavalida(cadbl(comanda), True, vpararcomanda) Then
    If vpararcomanda Then comanda = "0": Exit Sub
    If MsgBox("Aquesta comanda ESTÀ PARADA O HI HA ALGUN MOTIU PER PARAR-LA." + Chr(10) + "Vols continuar igualment?", vbCritical + vbYesNo + vbDefaultButton2, "ATENCIÓ") = vbNo Then Exit Sub
  End If
  
  proces = rsttmp!producte
  amplereb = cadbl(rsttmp!amplereb)
  petit = cadbl(comanda)
  If cadbl(rsttmp!linkcomanda1) < petit And cadbl(rsttmp!linkcomanda1) > 0 Then petit = cadbl(rsttmp!linkcomanda1)
  If cadbl(rsttmp!linkcomanda2) < petit And cadbl(rsttmp!linkcomanda2) > 0 Then petit = cadbl(rsttmp!linkcomanda2)
  comanda.Tag = petit
  If Not hihaseccio("L", rsttmp!producte) Then
      MsgBox "No hi ha seccio de laminadora en aquesta comanda"
      Command1.Enabled = False:   Command2.Enabled = False:   Command3.Enabled = False: Exit Sub
  End If
  
  
  If cadbl(rsttmp!refilatd) <> 0 Then primerproces.Visible = True
  If Not comprovarsiesprimer(petit, cadbl(linkcomanda2), atrim(comanda)) Then MsgBox "Aquest no es el primer procès d'aquesta comanda." + Chr(10) + "Comença per l'altra": Exit Sub
  
 ' If rsttmp!producte = "PC2" Then
 '     resp = UCase(InputBox("Vols laminar amb l'anonim o amb l'imprès? [A] o [I]", "Escull laminació", "I"))
 '     If resp = "A" Then linkcomanda = cadbl(rsttmp!linkcomanda2)
 '     If resp = "I" Then linkcomanda = cadbl(rsttmp!linkcomanda1)
 '     If cadbl(linkcomanda) = 0 Then MsgBox "Valor entrat no vàlid.", vbCritical, "Error": Exit Sub
 '    Else
        If cadbl(rsttmp!lotmatdesb1) = cadbl(comanda) Then
           linkcomanda = cadbl(rsttmp!lotmatdesb2)
          Else: linkcomanda = cadbl(rsttmp!lotmatdesb1)
        End If
  'End If
  
  If cadbl(comanda) > cadbl(linkcomanda) Then
     ncomanda2 = cadbl(comanda): ncomanda = cadbl(linkcomanda)
    Else: ncomanda = cadbl(comanda): ncomanda2 = cadbl(linkcomanda)
  End If
  primerproces.Tag = "invertit"
  If proces = "PC2" And primerproces.Visible = False Then primerproces.Tag = ""
  If proces <> "PC2" And primerproces.Visible = True Then primerproces.Tag = ""
  
  'aviso si es triplex que el material del mig ha de ser tractat dues cares
  comprovarsiestriplexiavisar
  
  possarelgrupdepalets
  ensenya_totals
  calcular_totals True
  bobines.RecordSource = "select * from bobineslam where controlid=-1"
  bobines.Refresh
  
  'miro si es stock o packing
  stockopacking = "P"
  Set rst = dbtmp.OpenRecordset("SELECT comandes_extres.assignarstock as estoc frOM comandes_extres WHERE comanda=" + atrim(ncomanda2) + ";")
  If Not rst.EOF Then If rst!estoc Then stockopacking = "E"
  
  ettipuscola.Tag = ""
  Set rsttmp = dbtmp.OpenRecordset("select rebmacroperforat,microperforat,tipusadhesiu,marcailinia,cantitatex,codibarras,espessor,mesuraesp,comanda,refclient,comandaclient,texteimpressio from comandes where comanda=" + atrim(cadbl(comanda)))
  'diria que aixó ja no funciona If Mid(nomdelacola(cadbl(rsttmp!tipusadhesiu)), 1, 1) = "@" Then ettipuscola.Tag = "1"
  mesuraespcomanda = ""
  If Not rsttmp.EOF Then
     tmetres.Tag = cadbl(rsttmp!cantitatex)
     Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rsttmp!mesuraesp)))
     If Not rsttmp2.EOF Then mesuraespcomanda = rsttmp2!descripcio
  End If
  refclient = "": comandaclient = ""
  texteimpresio = ""
  refclient = atrim(rsttmp!refclient)
  comandaclient = atrim(rsttmp!comandaclient)
   'clixes.Enabled = True
  texteimpresio = IIf(atrim(rsttmp!marcailinia) = "", atrim(rsttmp!texteimpressio), atrim(rsttmp!marcailinia))
  micrescomanda = cadbl(rsttmp!espessor)
  codibarras = atrim(rsttmp!codibarras)
  Command1.Enabled = True: Command2.Enabled = True: Command3.Enabled = True
  mirarsiMICROoMACRO rsttmp
  Set rsttmp = Nothing
  'fins aqui comprovo comanda
  laminadores.RecordSource = "select * from laminadores where comanda=" + atrim(cadbl(comanda)) + " order by datainici,horainici"
  imppantones.RecordSource = "select * from laminadoresadhesius where comanda=" + atrim(cadbl(comanda))
  laminadores.Refresh
  imppantones.Refresh
  carregar_families_cola vnomcola, vnomenduridor
  If imppantones.Recordset.EOF Then
     crear_pantones vnomcola, vnomenduridor
     imppantones.RecordSource = "select * from laminadoresadhesius where comanda=" + atrim(cadbl(comanda))
  End If
  carregar_client_ntintersialtres
  reixa.ReBind
  calcular_totals True
  'If laminadores.Recordset.EOF And laminadores.Recordset.BOF And Command1.Enabled Then Command1_Click
  framebobines.Enabled = False: framepantones.Visible = False
  'If laminadores.Recordset.EOF Then Command1_Click
  vcolor = comprovarsireciclarmaterial(cadbl(ncomanda))
  reciclarmaterial1.BackColor = vcolor
  vcolor = comprovarsireciclarmaterial(cadbl(ncomanda2))
  reciclarmaterial2.BackColor = vcolor
  carregar_capcalera IIf(Command15.Tag = "imprimint", False, True)
'  If impresores.Recordset.EOF Then MsgBox "Baixa nova es començarà amb edició de Clixes.": Command4.Tag = "nou": crearseccio "C": Command4.Tag = ""
  ratoli "normal"
End Sub
Sub mirarsiMICROoMACRO(rsttmp As Recordset)
  If atrim(rsttmp!microperforat) <> "" And atrim(rsttmp!microperforat) <> "N" Then MsgBox "Aquesta comanda porta Microperforat en " + IIf(atrim(rsttmp!microperforat) = "C", "Calent", "Fred") + vbNewLine + vbNewLine + "NO FER BOBINES DE MES DE 80cm DE DIAMETRE", vbInformation, "ATENCIÓ": vperforat = True
  If atrim(rsttmp!rebmacroperforat) <> "N" And atrim(rsttmp!rebmacroperforat) <> "" Then MsgBox "Aquesta comanda porta MACROPERFORAT" + vbNewLine + vbNewLine + "NO FER BOBINES DE MES DE 80cm DE DIAMETRE", vbInformation, "ATENCIÓ": vperforat = True
End Sub
Sub carregar_families_cola(vnomcola As String, vnomenduridor As String)
  Dim rst As Recordset
  
  Dim vcolaexacte As Boolean
  
  Set rst = dbtmpb.OpenRecordset("SELECT comandes.comanda, comandes.tipusadhesiu, comandes_extres.colaexacte FROM comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda WHERE (((comandes.comanda)=" + atrim(cadbl(comanda)) + "));", , ReadOnly)
  If Not rst.EOF Then
      vcolaexacte = rst!colaexacte
      Set rst = dbtmp.OpenRecordset("select resina,enduridor,idfamilia,idsubfamilia from adhesius where codi=" + atrim(cadbl(rst!tipusadhesiu)), , ReadOnly)
      If Not rst.EOF Then
         pantone(0).Tag = atrim(cadbl(rst!idfamilia))
         pantone(1).Tag = atrim(cadbl(rst!idsubfamilia))
         vnomcola = atrim(rst!resina)
         vnomenduridor = atrim(rst!enduridor)
      End If
      If vcolaexacte Then
         pantone(0).Tag = "9999"
         pantone(1).Tag = "9999"
      End If
  End If
  
End Sub
 
Sub possarelgrupdepalets()
   Dim rstopcions As Recordset
   obrestocks
   numgrup = ""
   numgrup.Tag = ""
   Set rstopcions = dbstocks.OpenRecordset("select * from opcionsdajust where comanda=" + atrim(cadbl(ncomanda2)))
   If rstopcions.EOF Then Exit Sub
   numgrup = "Grup de palets: " + atrim(cadbl(rstopcions!grupdestoc))
   numgrup.Tag = atrim(cadbl(rstopcions!grupdestoc))
   Set dbstocks = Nothing
End Sub
Sub carregar_client_ntintersialtres()
  Dim rstnt As Recordset
  Dim codicli As Double
  client.Caption = ""
  Set rstnt = dbtmp.OpenRecordset("select client,proximaseccio,cilindres,numerotintes from comandes where comanda=" + atrim(cadbl(comanda)))
  If Not rstnt.EOF Then
       ntintes = cadbl(rstnt!numerotintes)
       ncilindre = cadbl(rstnt!cilindres)
       framepantones.Tag = atrim(rstnt!proximaseccio)
       codicli = cadbl(rstnt!client)
       Set rstnt = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(codicli))
       If Not rstnt.EOF Then client.Caption = rstnt!nom
  End If
End Sub
Sub gravar_valorxrdefecte_adhesius(Optional vColaoEnduridor As String)
'On Error GoTo fi
 If Not imppantones.Recordset.EOF Then
  If vColaoEnduridor = "C" Or vColaoEnduridor = "" Then
   escriure_ini "Laminadores", "lot1", compantone(0), "comandes.ini"
   escriure_ini "Laminadores", "nomadhesiu1", pantone(0), "comandes.ini"
  End If
  If vColaoEnduridor = "E" Or vColaoEnduridor = "" Then
   escriure_ini "Laminadores", "lot2", compantone(1), "comandes.ini"
   escriure_ini "Laminadores", "nomadhesiu2", pantone(1), "comandes.ini"
  End If
 End If
fi:
End Sub
Sub crear_pantones(vnomcola As String, vnomenduridor As String)
  r = " comanda "
  For i = 1 To 2
    r = r + ",tinta" + atrim(i) + "a "
  Next i
  Set rsttmp = dbtmp.OpenRecordset("select " + r + " from comandes where comanda=" + atrim(comanda))
  If Not rsttmp.EOF Then
   imppantones.Recordset.AddNew
   imppantones.Recordset!comanda = comanda
   imppantones.Recordset!pantone1 = llegir_ini("Laminadores", "nomadhesiu1", "comandes.ini") 'vnomcola
   imppantones.Recordset!pantone2 = llegir_ini("Laminadores", "nomadhesiu2", "comandes.ini") 'vnomenduridor
   imppantones.Recordset!lot1 = "" 'llegir_ini("Laminadores", "lot1", "comandes.ini")
   imppantones.Recordset!lot2 = "" 'llegir_ini("Laminadores", "lot2", "comandes.ini")
   imppantones.Recordset!comanda = comanda
   imppantones.Recordset.Update
  End If
  imppantones.Refresh
  imppantones.UpdateControls
End Sub
Private Sub Command5_Click()
'  If Not clixes.Enabled Then Exit Sub
i = 0
While barraestat.Caption = "Calculant els totals..."
  DoEvents
Wend
  dblots.Visible = False
  framepantones.Visible = False
  frameempalmes.Visible = False
  framebobentrada.Visible = False

  bobines.UpdateRecord
 If laminadores.Recordset!tipus = "F" Then
     If cadbl(reixabobines.Columns(4).Text) = 0 And Not bobines.Recordset.EOF Then reixabobines.col = 4: reixabobines.SetFocus: MsgBox "Falten els metres a la bobina": Exit Sub
     calcular_totals
     While barraestat.Caption = "Calculant els totals..."
       DoEvents
     Wend
     nova_bobina
     Command13_Click
     copiarbobentanterior
     crearunempalmerestomalo
     Else: MsgBox "Has d'escullir una linia de FUNCIONAMENT."
  End If
  ' calcular_totals
End Sub
Sub crearunempalmerestomalo()
  empalmes.Recordset.AddNew
  empalmes.Recordset!id = bobines.Recordset!id
  empalmes.Recordset!observacions = "RESTO MALO"
  empalmes.Recordset.Update
End Sub
Sub copiarbobentanterior()
 Dim rsttmp1 As Recordset
 Dim primer As Boolean
 Dim rsttmp2 As Recordset
 Dim utili As Boolean
 Set rsttmp1 = dbtmpb.OpenRecordset("select * from bobinesentlam where id=" + atrim(cadbl(bobinesent.Tag))) ' + " and paletobobina='B'")
 obrestocks
 primer = True
 While Not rsttmp1.EOF
  'If rsttmp1!paletobobina <> "P" And rsttmp1!paletobobina <> "B" Then
   'If primer Then r = "carregartaulatmp": bobentrada_DblClick: primer = False: r = ""
   carregar_bobinesdentrada "mirarsiutilitzada", , rsttmp1!palet, rsttmp1!bobina, ncomanda, utili, ncomanda2
   If Not utili Then
    bobinesent.Recordset.AddNew
    bobinesent.Recordset!id = bobines.Recordset!id
    bobinesent.Recordset!desb = rsttmp1!desb
    bobinesent.Recordset!palet = rsttmp1!palet
    bobinesent.Recordset!bobina = rsttmp1!bobina
    bobinesent.Recordset!paletobobina = rsttmp1!paletobobina
    bobinesent.Recordset.Update
    bobinesent.Refresh
    bobinesent.Recordset.FindFirst "palet=" + atrim(rsttmp1!palet) + " and bobina=" + atrim(rsttmp1!bobina)
    'demanar si es final de bobina
'    carregar_bobinesdentrada "marcarutilitzadademanar", , cadbl(rsttmp1!palet), cadbl(rsttmp1!bobina), ncomanda, False, ncomanda2
      Else: MsgBox "La bobina " + atrim(rsttmp1!palet) + "/" + atrim(rsttmp1!bobina) + " ja està utilitzada per aquesta comanda i no s'afegirà a la bobina actual.", vbCritical, "Atenció"
   End If
  'End If
  rsttmp1.MoveNext
 Wend
 bobinesent.Refresh
 Set rsttmp1 = Nothing
 Set rsttmp2 = Nothing
 dbstocks.Close
End Sub
Sub nova_bobina()
  Dim rstmp As Recordset
  Dim rsttmp2 As Recordset
  Dim col As Byte
  Dim elgran As Double
  reixabobines.Tag = "afegint"
  If Not bobines.Recordset.EOF Then
   If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
   bobines.Recordset.Update
  End If
  Set rsttmp2 = dbtmpb.OpenRecordset("select id  from laminadores where comanda=" + atrim(laminadores.Recordset!comanda))
  elgran = 0
  While Not rsttmp2.EOF
   Set rstmp = dbtmpb.OpenRecordset("select max(numerodebobina) as elgran from bobineslam where controlid=" + atrim(rsttmp2!id))
   If Not rstmp.EOF Then
      If cadbl(rstmp!elgran) > elgran Then elgran = cadbl(rstmp!elgran)
   End If
   rsttmp2.MoveNext
  Wend
  Set rstmp = dbtmpb.OpenRecordset("select * from bobineslam where controlid=" + atrim(laminadores.Recordset!id) + " and numerodebobina=" + atrim(elgran))
  bobines.Recordset.AddNew
  bobines.Recordset!numerodebobina = elgran + 1
  bobines.Recordset!controlid = atrim(laminadores.Recordset!id)
  bobines.Recordset!numempalmes = 0
  bobines.Recordset!datafab = Date
  col = 0
  bobines.Recordset!operari1 = numop
  bobinesent.Tag = atrim(rstmp!id)
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
  reixabobines.SetFocus
  Set rstmp = Nothing
  Set rstmp2 = Nothing
If reixabobines.Text = "0" Then reixabobines.SelLength = Len(reixabobines.Text)
reixabobines.Tag = ""
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
 
 If MsgBox("Segur que vols borrar aquesta bobina?", vbCritical + 4, "Atenció") = vbYes Then
     If Not bobines.Recordset.EOF Then
       dbtmpb.Execute "delete * from lamempalmes where id=" + atrim(cadbl(bobines.Recordset!id))
       dbtmpb.Execute "delete * from bobinesentlam where id=" + atrim(cadbl(bobines.Recordset!id))
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
If cadbl(bobines.Recordset!metres) = 0 Then
  mtrs = cadbl(InputBox("Entra els Metres de la bobina", "Atenció"))
  If mtrs = 0 Then Exit Sub
  If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
  bobines.Recordset!metres = cadbl(mtrs)
  bobines.Recordset.Update
End If
  If cont = 3 Then cont = 0: form1.Caption = "Baixes Comandes (Laminadores)"
  If InStr(1, form1.Caption, "Imprimint la bobina") <> 0 Then cont = cont + 1: Exit Sub
  form1.Caption = "Imprimint la bobina"
  bobines.UpdateRecord
  If Not bobines.Recordset.EOF Then numb = bobines.Recordset!numerodebobina
  calcular_totals
  'wait (2)
  form1.Caption = "Imprimint la bobina."
  bobines.Recordset.FindFirst "numerodebobina=" + atrim(cadbl(numb))
  imprimir_bobina
  imprimir_controlqualitat numerobobina, bobines.Recordset!operari1, bobines.Recordset!numerodebobina, bobines.Recordset!datafab
  form1.Caption = "Baixes Comandes (Laminadores)"
  
End Sub
Function numerobobina() As Double
 numerobobina = cadbl(comanda)
 If numerobobina - 2 = cadbl(linkcomanda) Then numerobobina = cadbl(linkcomanda)
End Function


Private Sub imprimir_bobina()
  
netejarreport llistat
form1.Caption = "Imprimint la bobina.."
crear_taula_lam_empalmes
form1.Caption = "Imprimint la bobina..."
possar_valors_taula_lam_empalmes
form1.Caption = "Imprimint la bobina...."
llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "etempalmeslam.rpt"
'llistat.Destination = crptToWindow
llistat.Destination = crptToPrinter
llistat.CopiesToPrinter = 1
llistat.DataFiles(0) = cami
DoEvents
wait (5)
For i = 1 To 10
  llistat.Formulas(i) = ""
Next i
 llistat.Formulas(0) = "proximaseccio='" + buscarproximaseccio(ncomanda, ncomanda2) + "'"
'llistat.DiscardSavedData = True
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 
 form1.Caption = "Imprimint la bobina....."
llistat.Action = 1
llistat.Formulas(0) = ""
form1.Caption = "Imprimint la bobina......"
netejarreport llistat
form1.Caption = "Imprimint la bobina......."
DoEvents
End Sub
Function buscarproximaseccio(ncomanda As Double, ncomanda2 As Double) As String
   Dim rst As Recordset
   Dim rst2 As Recordset
   Set rst = dbtmp.OpenRecordset("SELECT comandes.comanda, comandes.linkcomanda2,productes.ruta, comandes.proximaseccio,comandes.impressio FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(ncomanda) + "));")
   Set rst2 = dbtmp.OpenRecordset("SELECT comandes.comanda, productes.ruta, comandes.proximaseccio,comandes.impressio FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(cadbl(rst!linkcomanda2)) + "));")
   If Not rst2.EOF Then If InStr(1, rst2!ruta, "L") > 0 And rst2!proximaseccio = "L" Then buscarproximaseccio = "L": GoTo fi
   If InStr(1, rst!ruta, "L") > 0 And rst!proximaseccio = "L" Then buscarproximaseccio = Mid(rst!ruta, InStr(1, rst!ruta, "L") + 1, 1)
fi:
   Set rst = Nothing
   Set rst2 = Nothing
End Function
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

Function buscarmicrescomanda(comanda1 As Double, comanda2 As Double) As Double
   Dim rstc1 As Recordset
   Dim rstc2 As Recordset
   Dim rstc3 As Recordset
   Dim comanda3 As Double
   Dim espesor2 As Double
   Dim espesor1 As Double
   Dim espesor3 As Double
   
   Set rstc1 = dbtmp.OpenRecordset("select espessor,refilatd,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(comanda1))
   Set rstc2 = dbtmp.OpenRecordset("select espessor,refilatd from comandes where comanda=" + atrim(comanda2))
   If rstc1.EOF Or rstc2.EOF Then Exit Function
   comanda3 = IIf(rstc1!linkcomanda1 = comanda2, rstc1!linkcomanda2, rstc1!linkcomanda1)
   If comanda3 > 0 Then espesor3 = micresdelmaterialcomanda(comanda3)
   espesor1 = micresdelmaterialcomanda(comanda1)
   espesor2 = micresdelmaterialcomanda(comanda2)
   If cadbl(rstc1!refilatd) = 0 Then
       buscarmicrescomanda = espesor1 + espesor2 + espesor3
        Else: buscarmicrescomanda = espesor1 + espesor2
   End If
   Set rstc1 = Nothing
   Set rstc2 = Nothing
   Set rstc3 = Nothing
End Function

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
 Set rs2 = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2,producte from comandes where comanda=" + atrim(comanda.Tag))
 r = atrim(capcalera.capcalera.Recordset!matdesb1) + " + " + atrim(capcalera.capcalera.Recordset!matdesb2)
 If rs2.EOF Then MsgBox "No s'ha trobat la comanda": Exit Sub
 rs.AddNew
 rs!numlotbobina = cadbl(comanda)
 If rs!numlotbobina - 2 = cadbl(linkcomanda) Then rs!numlotbobina = cadbl(linkcomanda)
 rs!numlot1 = cadbl(rs2!comanda)
 rs!numlot2 = cadbl(rs2!linkcomanda1)
 rs!numlot3 = cadbl(rs2!linkcomanda2)
 Set rs2 = Nothing
 rs!numbobsort = cadbl(bobines.Recordset!numerodebobina)
 rs!numop = cadbl(bobines.Recordset!operari1)
 rs!numop2 = cadbl(bobines.Recordset!operari2)
 rs!datafab = Format(bobines.Recordset!datafab, "dd/mm/yy")
 rs!client = client.Caption + " "
 rs!texteimpressio = texteimpresio + " "
 rs!refclient = refclient + " "
 rs!observacio = atrim(bobines.Recordset!observacio) + " "
 rs!comandaclient = comandaclient + " "
 rs!material = r + " "
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
 rs!espessor = buscarmicrescomanda(cadbl(comanda), cadbl(linkcomanda))
 'actualitzo les dades de la bobina
    bobines.Recordset.Edit
    bobines.Recordset!ample = rs!ample
    bobines.Recordset!espessor = rs!espessor
    bobines.Recordset.Update
 'fins aqui actualitzo
 'llistat.Formulas(0) = "mesuraesp='(" + mesuraespcomanda + ")'"
 llistat.Formulas(0) = "mesuraesp='(micres)'"
 rs!codibarres = codibarras + " "
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
Sub crear_taula_lam_empalmes()
  Dim camps As String
  Dim camps2 As String
  Dim camps3 As String
  Dim camps4 As String
  'On Error Resume Next
  'dbtmpb.Execute "drop table tmp_lam_empalmes"
  'On Error GoTo 0
  If Not existeixlataula("tmp_lam_empalmes") Then
        camps = "numlotbobina double,numlot1 double,numlot2 double, numlot3 double,numbobsort double, numop double,datafab string,numbobent1 string,numbobent2 string,client string,refclient string,comandaclient string,texteimpressio string,material string,ample double,plegat double,"
        camps2 = "solapa double,espessor double,metres double,kilos double, empalme1 string,mtrs1 double,empalme2 string, mtrs2 double,empalme3 string, mtrs3 double,empalme4 string"
        camps3 = " ,mtrs4 double,empalme5 string,mtrs5 double,numbobent3 string,numbobent4 string, codibarres string, observacio string, numop2 double,empalme6 string,mtrs6 double,empalme7 string,mtrs7 double,empalme8 string,mtrs8 double"
        camps4 = ",dist1 double, dist2 double, dist3 double, dist4 double, dist5 double, dist6 double, dist7 double, dist8 double"
        'ample double,plegat double,solapa double,espessor double,metres double,kilos double)"
        dbtmpb.Execute ("create table tmp_lam_empalmes (" + camps + camps2 + camps3 + camps4) + ")"
      Else:
        dbtmpb.Execute "delete * from tmp_lam_empalmes"
        crear_camp_numlotbobina
        wait 2
  End If
End Sub
Sub crear_camp_numlotbobina()
 On Error Resume Next
 'dbtmpb.Execute "alter table tmp_lam_empalmes add column numlotbobina double"
 On Error GoTo 0
End Sub
Function cabool(valor As Variant)
  If IsNull(valor) Then cabool = 0: Exit Function
  If valor Then
    cabool = 1
   Else: cabool = 0
  End If
End Function


Sub emplenar_capcalera_imp(rsttemp As Recordset)
 Dim rst As Recordset
 Set rst = dbtmpb.OpenRecordset("select * from laminadorestot where comanda=" + comanda.Text)
 If Not rst.EOF Then
   rsttemp!tipusmat1 = atrim(rst!matdesb1)
   rsttemp!tipusmat2 = atrim(rst!matdesb2)
   rsttemp!micres1 = atrim(rst!micres1)
   rsttemp!micres2 = atrim(rst!micres2)
   rsttemp!matimpres1 = cabool(rst!matimpres1)
   rsttemp!matimpres2 = cabool(rst!matimpres2)
   rsttemp!matcm1 = atrim(rst!amplada1)
   rsttemp!matcm2 = atrim(rst!amplada2)
   rsttemp!caratarctada1 = cabool(rst!tractatabaix)
   rsttemp!caratractada2 = cabool(rst!tractatadalt)
   rsttemp!tensio2 = atrim(rst!tensio2)
   rsttemp!tensio1 = atrim(rst!tensio1)
   rsttemp!observacions = atrim(rst!observacions)
   rsttemp!camisa = atrim(rst!camisa)
   rsttemp!cilindrecola = atrim(rst!cilincola)
   rsttemp!adhesiu = atrim(rst!adhesiu)
   rsttemp!impresvisual = atrim(rst!impresvisual)
   rsttemp!tensioreb = atrim(rst!tensioreb)
   
 End If
 Set rst = Nothing
End Sub



Sub imprimir_fulla()
  Dim mtrsparcialanteriors As Double
  Dim rst As Recordset
  Dim rsttemp As Recordset
  Dim relaciodpalets As String
   Dim rsttmp2 As Recordset
   Dim nb As String
   Dim np As Double
   Dim linia As Double
   Dim rsttmpbob As Recordset
   Dim canvicam As String
   Dim instsql2 As String
   Dim instsql As String
   Dim rstdr As Recordset
   Dim vcarpetadesti As String
   
   form1.Caption = "Imprimint..."
   nample = 0
   netejarreport llistat
 'carregar_client_ntintersialtres
   'panelimprimir.Visible = True
'panelimprimir.Top = Frame3.Top
  crear_taula_laminadora_baixa
  obrestocks
  Set rsttemp = dbtmpb.OpenRecordset("tmp_lam_baixa")
  imppantones.Refresh
  rsttemp.AddNew

  ' busco l'ample
   'ample_palet
  '-----------
  

  
  With rsttemp
  !comanda = atrim(comanda.Text)
  '!client = atrim(client.Caption)
  !client = client.ToolTipText
  !firmat = atrim(firmat.Caption)
  !nomfirmat = possarnomfirmat
  '!tintersrentats = cadbl(trentats)
  '!portaclixers = cadbl(pclixers)
  '!canvienfilada = atrim(canvienfilada)
  '!numtintes = cadbl(ntintes)
  '!cilindre = cadbl(ncilindre)
  !comandaacavada = IIf(comandaacavada.Value, 1, 0)
  'prep clixe
  emplenar_capcalera_imp rsttemp
  Set rst = dbtmpb.OpenRecordset("select id,operari,datainici,horainici,datafi,horafi,observacio,canvicamisa from laminadores where comanda=" + comanda.Text + " and tipus='C'")
  If Not rst.EOF Then
   rst.MoveLast
   If Not rst.BOF Then rst.MovePrevious:
   If rst.BOF Then
      rst.MoveNext
    Else: rst.MovePrevious: If rst.BOF Then rst.MoveNext
   End If
  End If
  i = 1
  canvicam = ""
  While Not rst.EOF
    .Fields("prepmaquina_data" + Trim(i)) = Format(atrim(rst!datainici), "dd/mm/yy")
    .Fields("prepmaquina_op" + Trim(i)) = cadbl(rst!operari)
    .Fields("prepmaquina_de" + Trim(i)) = Format(atrim(rst!horainici), "hh:nn")
    .Fields("prepmaquina_fins" + Trim(i)) = Format(atrim(rst!horafi), "hh:nn")
    .Fields("prepmaquina_observacions" + Trim(i)) = atrim(rst!observacio)
    If canvicam = "" And rst!canvicamisa = "Sí" Then canvicam = "Sí"
    i = i + 1
    rst.MoveNext
    If i > 2 Then rst.MoveLast: rst.MoveNext
  Wend
  
  If canvicam = "" Then canvicam = "No"
'relleus i descans
   i = 1
   Set rstdr = dbtmpb.OpenRecordset("select * from controldescansrelleu where seccio='" + atrim(lletraseccio) + "' and comanda=" + atrim(ncomanda) + " or comandafi=" + atrim(ncomanda))
   While Not rstdr.EOF And i < 4
        .Fields("prepdr_data" + Trim(i)) = Format(atrim(rstdr!datainici), "dd/mm/yy")
        .Fields("prepdr_op" + Trim(i)) = cadbl(rstdr!operari)
        .Fields("prepdr_de" + Trim(i)) = Format(atrim(rstdr!horainici), "hh:nn")
        .Fields("prepdr_fins" + Trim(i)) = Format(atrim(rstdr!horafi), "hh:nn")
        .Fields("prepdr_observacions" + Trim(i)) = atrim(cadbl(rstdr!hores)) + " Hores de " + atrim(rstdr!tipus)
         i = i + 1
        rstdr.MoveNext
   Wend
  
  'temps funcionament
  Set rst = dbtmpb.OpenRecordset("select id,operari,datainici,horainici,datafi,horafi,observacio,mtrsminut,totalmetres, metresparcial,mtrsminutcola from laminadores where comanda=" + comanda.Text + " and tipus='F'")
  If Not rst.EOF Then
    rst.MoveLast
   If Not rst.BOF Then rst.MovePrevious:
   If rst.BOF Then
      rst.MoveNext
    Else: rst.MovePrevious: If rst.BOF Then rst.MoveNext Else rst.MovePrevious: If rst.BOF Then rst.MoveNext
   End If
  End If
  i = 1
  While Not rst.EOF
    .Fields("tempslam_data" + Trim(i)) = Format(atrim(rst!datainici), "dd/mm/yy")
    .Fields("tempslam_op" + Trim(i)) = cadbl(rst!operari)
    .Fields("tempslam_de" + Trim(i)) = Format(atrim(rst!horainici), "hh:nn")
    .Fields("tempslam_fins" + Trim(i)) = Format(atrim(rst!horafi), "hh:nn")
    .Fields("tempslam_observacio" + Trim(i)) = atrim(rst!observacio)
    .Fields("tempslam_mtrsmin" + Trim(i)) = cadbl(rst!mtrsminut)
    .Fields("tempslam_mtrscola" + Trim(i)) = cadbl(rst!mtrsminutcola)
    .Fields("tempslam_mtrslaminats" + Trim(i)) = cadbl(rst!totalmetres) - mtrsparcialanteriors + cadbl(rst!metresparcial)
    mtrsparcialanteriors = cadbl(rst!metresparcial)
    i = i + 1
    rst.MoveNext
  Wend
  
  'acavar comandes
  Set rst = dbtmpb.OpenRecordset("select * from laminadoresadhesius where comanda=" + atrim(comanda))
  i = 1
  If Not rst.EOF Then
   For i = 1 To 2
    .Fields("desc_adhesiu" + Trim(i)) = atrim(rst.Fields("pantone" + atrim(i)))
    .Fields("numlot_ad" + Trim(i)) = atrim(rst.Fields("lot" + atrim(i)))
    .Fields("litres_ad" + Trim(i)) = IIf(atrim(rst.Fields("kg" + atrim(i))) = "", 0, atrim(rst.Fields("kg" + atrim(i))))
    .Fields("observacions1") = atrim(rst!observacions)
   Next i
  End If

  'posso els camps de totals
    !hclixe = cadbl(hclixe): !hmaquina = cadbl(hmaquina): !hajusts = cadbl(hajusts): !hfunc = cadbl(hfunc): !tprova = cadbl(tprova): !tbob = cadbl(tbob): !tmtrs = cadbl(tmetres): !tkilos = cadbl(tkilos): !mtrsmin = cadbl(kiloshora)
  '!acavada = comandaacavada
  Set rstbob = Nothing
  Set rst = Nothing
  
  
    
  End With
  
  'passo les bobines a la taula del llistat
  Set rst = dbtmpb.OpenRecordset("select id,operari,datainici,horainici,datafi,horafi,observacio,mtrsminut from laminadores where comanda=" + comanda.Text + " and tipus='F'")
  If rst.EOF Then dbtmpb.Execute "insert into tmp_lam_baixa_bob (operari ,operari2 ,palet1,bobent1,bobsort,kilos,metres) values (0,0,0,'0',0,0,0)"
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
        Set rsttmp2 = dbtmpb.OpenRecordset("select * from bobineslam where controlid=" + atrim(cadbl(rst!id)))
        
        With rsttmp2
        If Not rsttmp2.EOF Then
         rsttmp2.MoveLast
         rsttmp2.MoveFirst
          Else: dbtmpb.Execute "insert into tmp_lam_baixa_bob (operari,operari2,palet1,bobent1,bobsort,kilos,metres) values (0,0,0,'0',0,0,0)"
        End If
        While Not rsttmp2.EOF
          'If rsttmp2.AbsolutePosition + 1 = rsttmp2.RecordCount Then
              If Not rsttmp2.EOF Then Set rsttmpbob = dbtmpb.OpenRecordset("select * from bobinesentlam where id=" + atrim(cadbl(rsttmp2!id)) + " order by paletobobina ASC")
              nb = 0
              np = 0
              If Not rsttmpbob.EOF Then
                 rsttmpbob.MoveLast
                 rsttmpbob.MoveFirst
                 np = rsttmpbob!palet
                 nb = rsttmpbob!bobina
                 'If rsttmpbob.RecordCount > 1 Then nb = "*" + nb
                 'aprofito per buscar lamplada del palet
                 Set rststocks = dbstocks.OpenRecordset("select ample from palets where idpalet=" + atrim(np))
                 If Not rststocks.EOF Then nample = rststocks!ample
              End If
              relaciodpalets = ""
              While Not rsttmpbob.EOF
                relaciodpalets = relaciodpalets + IIf(relaciodpalets <> "", " - ", "") + atrim(rsttmpbob!palet) + "/" + atrim(rsttmpbob!bobina)
                rsttmpbob.MoveNext
              Wend
              
              dbtmpb.Execute "insert into tmp_lam_baixa_bob (operari,operari2,palet1,bobent1,bobsort,kilos,metres,relaciodpalets,observacions) values (" + atrim(cadbl(!operari1)) + "," + atrim(cadbl(!operari2)) + "," + atrim(np) + ",'" + atrim(nb) + "'," + atrim(cadbl(!numerodebobina)) + "," + atrim(cadbl(!kilos)) + "," + atrim(cadbl(!metres)) + ",'" + atrim(relaciodpalets) + "','" + atrim(!observacio) + "')"
              
           ' Else: dbtmpb.Execute "insert into tmp_imp_baixa_bob (operari,palet,bobent,bobsort,kilos,metres) values (" + atrim(cadbl(!operari1)) + "," + "0" + "," + "0" + "," + atrim(cadbl(!numerodebobina)) + "," + atrim("0") + "," + atrim("0") + ")"`'          End If
          rsttmp2.MoveNext
          rsttemp!ample = nample
        Wend
    ''    rsttmp.MoveNext
     ''Wend
     rst.MoveNext
     End With
  Wend
  
  rsttemp.Update
  
  dbtmpb.Close
  Set dbtmpb = laminadores.Database
  
  Set rsttmp2 = Nothing
  Set rsttmpbob = Nothing
  'imprimir llistat
  
  'ATENCIÓ QUE FAIG SERVIR BAIXESLAMINADORA.RPT PERÒ LA QUE S'IMPRIMEIX ES LA BAIXESLAMINADORA_PDF perquè també fa el pdf
 '   i amb la versió que estava fet no es podia genera el PDF
 
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "baixeslaminadora.rpt"
' llistat.Destination = crptToWindow
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 2
 llistat.DataFiles(0) = cami
 'llistat.DiscardSavedData = True
 llistat.Formulas(0) = "canvicamisa='" + canvicam + "'"
 llistat.Formulas(1) = "texteimpresio='" + treure_apostruf(texteimpresio) + "'"
 llistat.Formulas(2) = "nommaquina='" + nommaquina(laminadores.Recordset!numeromaquina, "L") + "'"
' llistat.PrinterName = llegir_ini("Impressores", "nomfulla", "baixesimpressora.ini")
' llistat.PrinterPort = llegir_ini("Impressores", "portfulla", "baixesimpressora.ini")
' llistat.PrinterDriver = llegir_ini("Impressores", "driverfulla", "baixesimpressora.ini")
  DoEvents
  wait (4)
' If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
' For i = 1 To 2
'   llistat.Action = 1
'   wait 1
' Next i
  escriure_ini "General", "exportantpdfs", "si", llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\organitzar.ini"
  crearlacarpetaperexportar cadbl(comanda.Text), vcarpetadesti
  exportarllistatapdf llistat, llegir_ini("General", "rutallistats", "comandes.ini") + "baixeslaminadora_PDF.rpt", cadbl(comanda.Text), vcarpetadesti
  escriure_ini "General", "exportantpdfs", "no", llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\organitzar.ini"
  netejarreport llistat
  Set rsttmp = Nothing
  Set rst = Nothing
  Set dbstocks = Nothing
 'panelimprimir.Visible = False
 form1.Caption = "Baixes Comandes (Laminadores)"
 
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
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat
  oreport.ExportOptions.DiskFileName = vcarpetadesti + "\" + atrim(vnumc) + "_BaixaLaminadores.pdf"
  oreport.ExportOptions.PDFExportAllPages = True
  oreport.Export False
  For i = 1 To vllistat.PrinterCopies
     oreport.PrintOut False
     wait 1
  Next i
End Sub
Function nommaquina(nmaq, seccio) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select descripcio from maquines where codi=" + atrim(nmaq) + " and maquina='" + seccio + "'")
   If Not rst.EOF Then nommaquina = atrim(rst!descripcio)
End Function
Function treure_apostruf(ByVal n As String) As String
   While InStr(n, "'")
     n = Mid(n, 1, InStr(1, n, "'") - 1) + "´" + Mid(n, InStr(1, n, "'") + 1)
   Wend
   If n = "{[}]" Then n = ""
   treure_apostruf = n
End Function
Function possarnomfirmat() As String
  Dim rsttmp As Recordset
  Set rsttmp = dbtmp.OpenRecordset("select descripcio from operaris where maquina='L' and codi=" + atrim(cadbl(firmat)))
  If Not rsttmp.EOF Then
     possarnomfirmat = rsttmp!descripcio
  End If
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
  
  'On Error Resume Next
  ' dbtmpb.Execute "drop table tmp_lam_baixa"
  ' dbtmpb.Execute "drop table tmp_lam_baixa_bob"
  'On Error GoTo 0
  
  If Not existeixlataula("tmp_lam_baixa") Or Not existeixlataula("tmp_lam_baixa_bob") Then
        campsextra = " nomfirmat text,firmat text,"
        campscapcalera = " comanda double, comanda2 double, client string,comandaacavada byte,"
        campscapcalera = campscapcalera + "tipusmat1 text, micres1 double, matcm1 double, matimpres1 byte, caratarctada1 byte, tensio2 double, tipusmat2 text, micres2 double, matcm2 double, matimpres2 byte, "
        campscapcalera2 = "caratractada2 byte,tensio1 double, obsdesb2 text,camisa double, cilindrecola double, adhesiu double, impresvisual double,tensioreb double, observacions text,ample double,"
        camps = camps + " prepmaquina_data1 string,prepmaquina_op1 byte,prepmaquina_de1 string,prepmaquina_fins1 string,prepmaquina_observacions1 string ,"
        camps = camps + " prepmaquina_data2 string,prepmaquina_op2 byte,prepmaquina_de2 string,prepmaquina_fins2 string,prepmaquina_observacions2 string ,"
        camps = camps + " canvicamisa byte, "
        
        camps3 = camps3 + "tempslam_data1 string,tempslam_op1 byte,tempslam_de1 string, tempslam_fins1 string,tempslam_mtrsmin1 double,tempslam_mtrscola1 double,tempslam_mtrslaminats1 double,tempslam_observacio1 string,"
        camps3 = camps3 + "tempslam_data2 string,tempslam_op2 byte,tempslam_de2 string, tempslam_fins2 string,tempslam_mtrsmin2 double,tempslam_mtrscola2 double, tempslam_mtrslaminats2 double,tempslam_observacio2 string,"
        camps2 = "tempslam_data3 string,tempslam_op3 byte,tempslam_de3 string, tempslam_fins3 string,tempslam_mtrsmin3 double,tempslam_mtrscola3 double, tempslam_mtrslaminats3 double,tempslam_observacio3 string,"
        camps2 = camps2 + "tempslam_data4 string,tempslam_op4 byte,tempslam_de4 string, tempslam_fins4 string,tempslam_mtrsmin4 double,tempslam_mtrscola4 double, tempslam_mtrslaminats4 double,tempslam_observacio4 string,"
        
        camps4 = " prepdr_data1 string,prepdr_op1 byte,prepdr_de1 string,prepdr_fins1 string,prepdr_observacions1 string ,"
        camps4 = camps4 + " prepdr_data2 string,prepdr_op2 byte,prepdr_de2 string,prepdr_fins2 string,prepdr_observacions2 string ,"
        camps4 = camps4 + " prepdr_data3 string,prepdr_op3 byte,prepdr_de3 string,prepdr_fins3 string,prepdr_observacions3 string ,"
        
        
        campspantone = " desc_adhesiu1 string, numlot_ad1 string,litres_ad1 double, observacions1 string, "
        campspantone = campspantone + " desc_adhesiu2 string, numlot_ad2 string,litres_ad2 double, observacions2 string, "
        campspantone = campspantone + " fi string "
        
        'creo els camps de total
        campstotal = ",hclixe double, hmaquina double, hajusts double, hfunc double, tprova double, tbob double,tmtrs double, tkilos double, mtrsmin double "
        
        'ample double,plegat double,solapa double,espessor double,metres double,kilos double)"
        'escriure_ini "a", "b", campsextra + camps + camps3 + camps2 + campspantone + campspantone2 + campstotal, "prova.ini"
        dbtmpb.Execute ("create table tmp_lam_baixa (" + campsextra + campscapcalera + campscapcalera2 + camps + camps3 + camps2 + camps4 + campspantone + campstotal) + ")"
        dbtmpb.Execute ("create table tmp_lam_baixa_bob (idbob integer,operari byte,operari2 byte,palet1 double,bobent1 string,palet2 double,bobent2 string,bobsort integer,kilos double,metres double,relaciodpalets string,observacions string)")
           Else
              dbtmpb.Execute "delete * from tmp_lam_baixa"
              dbtmpb.Execute "delete * from tmp_lam_baixa_bob"
   End If
   
End Sub



Private Sub Command8_Click()
calcular_totals
wait 4
imprimir_fulla
End Sub

Private Sub Command9_Click()
 tempseditant = 0
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
  'gravar_valorxrdefecte_adhesius
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

Private Sub dblots_LostFocus()
  If lots.Recordset.EditMode = 0 Then lots.Recordset.Edit
  lots.Recordset.Update
End Sub

Private Sub eliminarbobentrada_Click()
  If bobinesent.Recordset.EOF Then Exit Sub
  If MsgBox("Segur que vols eliminar la bobina d'entrada " + atrim(bobinesent.Recordset!palet) + "/" + atrim(bobinesent.Recordset!bobina), vbExclamation + vbYesNo, "Borrar bobina d'entrada") = vbYes Then
    'carregar_bobinesdentrada "marcarutilitzada", , bobinesent.Recordset!palet, bobinesent.Recordset!bobina, ncomanda, False, ncomanda2
    bobinesent.Recordset.Delete
    bobinesent.Refresh
    bobinesent.UpdateControls
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
assignardecimalipunt
If cadbl(numop) = 0 Then nomoperari_Click
End Sub
Sub modificar_bobines_malcalculades()
Dim rst As Recordset
  Dim vp As Double
  Dim vb As Double
  Dim mtrsrestants As Double
  Dim vdata As Date
  Dim numop As Double
  obrestocks
  Set rst = dbstocks.OpenRecordset("SELECT parcials.data,parcials.operari, Parcials.idpalet, Parcials.idbobina, Parcials.comanda, Parcials.metres, Parcials.data, Parcials.seccio From parcials WHERE seccio='L' and (((Parcials.comanda)='100') AND ((Parcials.metres)>500)) order by data desc;")
  While Not rst.EOF
     vp = rst!idpalet
     vb = rst!idbobina
     'If vp = 46770 Then Stop
     numop = rst!operari
     If IsNull(rst!Data) Then GoTo proxima
     vdata = rst!Data
     rst.Delete
     mtrsrestants = bobinesdentrada.calcular_mtrsdispreals(vp, vb)
     dbstocks.Execute "insert into parcials (idpalet,idbobina,operari,metres,comanda,data,seccio,utilitzada,orcomassignacio) values (" + atrim(vp) + "," + atrim(vb) + "," + atrim(cadbl(numop)) + "," + atrim(cadbl(mtrsrestants)) + ",100,#" + atrim(vdata) + "#,'L',true,0)"
     mtrsrestants = bobinesdentrada.calcular_mtrsdispreals(vp, vb)
'     bobinesdentrada.actualitzar_metres_disponibles vp, vb
     dbstocks.Execute "update bobines set disponible=0 where idpalet=" + atrim(vp) + " and idbobina=" + atrim(vb)
proxima:
     rst.MoveNext
  Wend
End Sub
Private Sub Form_Click()

'  MsgBox buscarproximaseccio(ncomanda, ncomanda2)
'  modificar_bobines_malcalculades
'MsgBox IIf(bobinesdentrada.esrestu(53542, 1), "R", "")
'ajustar_diametre_real "53000/1"
'demanar_verificacio_espesoritractat 53569, 2

'carregar_bobinesdentrada "marcarutilitzada", , cadbl(bobinesent.Recordset!palet), (bobinesent.Recordset!bobina), ncomanda, True, ncomanda2, IIf(primerproces.Tag = "invertit", True, False)
' appac = Shell("C:\Archivos de programa\swetiq.exe c:\prova.swe")
' wait (5)
' AppActivate appac
' wait (1)
' SendKeys ("%d")
' SendKeys ("{RIGHT}")
' SendKeys ("{ENTER}")
' SendKeys ("{ENTER}")
' SendKeys ("{ENTER}")
'Me.Caption = Me.Caption + "  (" + atrim(nummaq) + ")"
End Sub


'Sub obrestocks(Optional noobrirbd As Boolean)
'camistocks = llegir_ini("General", "ruta_stocksmdb", "comandes.ini")
'If camistocks = "{[}]" Then camistocks = "\\Ser2\documentos\Stock Reclamaciones\Estoc inplacsa.mdb"
'If Not existeix(camistocks) Then camistocks = "\\serverprodu\dades\progcomandes\dades\copiaestocinplacsa.mdb"
'If Not noobrirbd Then Set dbstocks = OpenDatabase(camistocks)
  
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If Chr$(KeyAscii) = "'" Then KeyAscii = Asc("´")
  tempseditant = Now
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


Sub posarimpresorapredeterminada()
  Dim impresora As String
  Dim r As String
  impresora = llegir_ini("Baixes", "impresoraa4", fitxerini)
  If impresora = "{[}]" Then escriure_ini "Baixes", "impresoraa4", "", fitxerini: Exit Sub
  On Error GoTo verror
  r = Shell("rundll32 PRINTUI.dll,PrintUIEntry /y /n """ + impresora + """", vbHide)
  Exit Sub
verror:
  escriure_ini "Baixes", "impresoraa4", "", fitxerini
End Sub

Private Sub Form_Load()
  Dim camistocks As String
  fitxerini = "comandes.ini"
  Shell "c:\windows\regedit.exe /s \\serverprodu\dades\progcomandes\aplicacio\desactivarctrl.reg"
  Shell ("net time \\serverprodu /set /y")
  posarimpresorapredeterminada
  camicomandes = llegir_ini("General", "cami", "comandes.ini")
  cami = llegir_ini("General", "camibaixes", "comandes.ini")
  'cami = "C:\Users\Usuari_Prog\Desktop\baixes.mdb"
  obrestocks True
  lletraseccio = "L"
  If cami = "{[}]" Then
    escriure_ini "General", "camibaixes", InputBox("Entra la ruta de baixes", "Atenció", "y:\comandes\baixes.mdb"), "comandes.ini"
  End If
  
  comanda = cadbl(llegir_ini("Baixes", "ultimacomanda", "comandes.ini"))
  r = cadbl(llegir_ini("Baixes", "nummaq", "comandes.ini"))
  nummaq = cadbl(r)
   assignardecimalipunt
  If LCase(App.EXEName) <> "baixes laminadores" And LCase(App.EXEName) <> "baixeslaminadora" Then form1.BackColor = &HFF80FF
  centerscreen Me
  'cami = "\\SERVERprodu\dades\progcomandes\dades\baixesprova.mdb"
  
  laminadores.DatabaseName = cami
  imppantones.DatabaseName = cami
  bobines.DatabaseName = cami
  empalmes.DatabaseName = cami
  bobinesent.DatabaseName = cami
  lots.DatabaseName = cami
  
  Set dbtmpb = OpenDatabase(cami)
  Set dbtmp = OpenDatabase(camicomandes)
  assignar_dbbaixes dbbaixes
  'On Error Resume Next
  'dbtmpb.Execute ("create table lotslam (nomlot string,codilot string)")
  'dbtmpb.Execute "drop table bobentradatmplamlam"
  'On Error GoTo 0
  crear_taulatemp_bobinesdentrada
  
 
  lots.Refresh
  'Set dbtmpb = OpenDatabase(laminadores.DatabaseName)
  rellotge.Enabled = True
  rellotge.Interval = 900
  
  'nummaq = 7
  'aquesta linia es per si s'obre desde un altre ordinador
  If nummaq = 0 Then
    maquina.Visible = True
   ' imprimir.Visible = True
   Else: maquina.Visible = False ': imprimir.Visible = False
  End If
  
  
  'If nummaq = 0 Then nummaq = 1
  'If nummaq > 90 Or nummaq = 0 Then
  '  i = nummaq
  '  nummaq = cadbl(InputBox("Entra el numero de maquina", "Atenció", "7"))
  '  If nummaq <> 7 And nummaq <> 5 And nummaq <> 2 Then MsgBox "Nomes hi ha la 2 la 5 i la 7": End
  '  If i = 0 Then escriure_ini "Baixes", "nummaq", atrim(nummaq), "comandes.ini"
  'End If
  
  
  laminadores.RecordSource = "select * from laminadores where comanda=-1"
  laminadores.Refresh
  
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
      If InStr(1, objecte.Name, "reciclarmaterial") = 0 And objecte.Name <> "AcroPDF1" And objecte.Name <> "nomoperari" And objecte.Name <> "Line1" And objecte.Name <> "rellotge" And objecte.Name <> "llistat" Then
        objecte.Enabled = False
      End If
     Next objecte
     
     
  frameempalmes.ZOrder 0
    framepantones.ZOrder 0
    framebobentrada.ZOrder 0
    

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  'If Shift = 2 Then MsgBox Trim(App.Major) + "." + Trim(App.Minor) + "." + Trim(App.Revision)
  If Shift = 4 Then MsgBox demanar_valor_micrometre("Prova de micrometre.", "99999/99")
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
 form1.Tag = "tancant"
 Unload capcalera
 Cancel = 0
 End
End Sub

Private Sub impresores_Reposition()
 


End Sub
Sub ensenya_les_bobines()
  Dim bk As String
  If Me.Name = "reixabobines" Then Exit Sub
  r = "-1"
  If laminadores.Recordset!tipus = "F" Then r = atrim(cadbl(laminadores.Recordset!id))
  If Not bobines.Recordset.EOF Then bk = bobines.Recordset!numerodebobina
  bobines.RecordSource = "select * from bobineslam where controlid=" + r + " order by numerodebobina"
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

Private Sub kbpantone_LostFocus(Index As Integer)
  kbpantone(Index) = comprovareldecimal(kbpantone(Index))
  If cadbl(kbpantone(Index)) > 99 Then MsgBox "Aquest valor es molt alt comprova que sigui correcte sisplau, potser el punt decimal es la coma.", vbInformation, "Atenció"
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
Function comprovareldecimal(v As String) As String
   If elsimboldecimal = "," Then comprovareldecimal = substituir(v, ".", ",")
   If elsimboldecimal = "." Then comprovareldecimal = substituir(v, ",", ".")
End Function
Function elsimboldecimal() As String
    Dim v As Double
    elsimboldecimal = ","
    If InStr(1, Trim(1 / 2), ".") > 0 Then elsimboldecimal = "."
End Function
Private Sub laminadores_Reposition()
If Not laminadores.Recordset.EOF Then
      ensenya_les_bobines
 End If
End Sub

Private Sub maquina_Click()
   nummaq = escullir_laminadora
   maquina.Caption = "Maq: " + atrim(nummaq)
   maquina.Tag = nummaq
End Sub
Function escullir_laminadora()
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "select codi,descripcio from maquines where maquina='L' and isnull(donadadebaixa)"
  formseleccio.Caption = "Selecció de Màquina"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   escullir_laminadora = cadbl(formseleccio.Data1.Recordset!codi)
'   nomoptmp = atrim(formseleccio.Data1.Recordset!descripcio)
   'If InStr(1, nomoperari.Caption, "MARTINEZ") Then
   '    Command12.Visible = True
   '   Else: Command12.Visible = False
   'End If
  End If
  Unload formseleccio
End Function
Private Sub nomoperari_Click()
 Dim numoptmp As Integer
 Dim nomoptmp As String
 If barraestat.Caption = "Calculant els totals..." Then Exit Sub
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "select codi,descripcio from operaris where maquina='L' and actiu<>0 order by codi"
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
      If InStr(1, objecte.Name, "reciclarmaterial") = 0 And objecte.Name <> "AcroPDF1" And objecte.Name <> "llistat" And objecte.Name <> "Line1" And objecte.Name <> "comandaacavada" Then
        objecte.Enabled = True
      End If
     Next objecte
      Else: If cadbl(numop) = 0 Then MsgBox "Has d'escullir un operari per treballar": Exit Sub
  End If
   command4_click
   aviscomprovarcomplexamesimpresionormal comanda
End Sub

Private Sub pantone_LostFocus(Index As Integer)
imppantones.Refresh
End Sub

Private Sub proces_Change()
Dim rsttmpp As Recordset
 
 Set rsttmpp = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(proces) + "'")
 If InStr(1, rsttmpp!ruta, "R") = 0 Then proces.Tag = "": Exit Sub
 If Not rsttmpp.EOF Then proces.Tag = Mid(rsttmpp!ruta, InStr(1, rsttmpp!ruta, "R") - 1, 1)
End Sub

Private Sub reixa_AfterUpdate()
  'calcular_totals
End Sub

Private Sub reixa_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
tempseditant = 0
End Sub

Private Sub reixa_BeforeDelete(Cancel As Integer)
  If Screen.ActiveControl.Name <> "Command14" Then
   If MsgBox("Segur que vols borrar aquesta linia i tot el seu contingut?", vbYesNo, "Atenció") = vbNo Then Cancel = 1
  End If
  If Cancel <> 1 Then
    If laminadores.Recordset!tipus = "F" Then r = atrim(cadbl(laminadores.Recordset!id))
    dbtmpb.Execute "delete * from bobineslam where controlid=" + r
  End If
End Sub

Private Sub reixa_ColEdit(ByVal ColIndex As Integer)
tempseditant = 0
End Sub

Private Sub reixa_DblClick()
If reixa.col = 14 Then
   r = triar_observacio(laminadores.Recordset!tipus)
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
If reixa.col = 9 Then
    mtrsparcials = cadbl(InputBox("Entra els metres parcials.", "Metres parcials"))
    If mtrsparcials > 0 Then
        If MsgBox("Segur que vols modificar els metres parcials?", vbCritical + vbYesNo, "Atenció") = vbYes Then
          reixa.EditActive = True
          reixa.Text = atrim(mtrsparcials)
          reixa.EditActive = False
          calcular_totals
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
  Dim ultimfocus As String
  If atrim(reixa.Columns(12)) = "" And reixa.Columns(5) = "C" And reixa.Tag <> "reixa" Then
      reixa.Tag = "reixa"
      If MsgBox("Has canviat de camisa en aquesta comanda?", vbYesNo, "Atenció") = vbYes Then
        reixa.Columns(12) = "Sí"
          Else: reixa.Columns(12) = "No"
      End If
      If laminadores.Recordset.EditMode > 0 Then laminadores.Recordset.Update: reixa.EditActive = False
      reixa.Tag = ""
  End If

End Sub

Private Sub reixa_KeyDown(KeyCode As Integer, Shift As Integer)
  tempseditant = Now
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
If laminadores.Recordset.EOF Then Exit Sub
For i = 0 To 14
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
If laminadores.Recordset!tipus = "C" Then reixa.Columns(12).Locked = False: reixa.Columns(14).Locked = False  ': reixa.Columns(11).Locked = False:reixa.Columns(7).Locked = False
If laminadores.Recordset!tipus = "F" Then reixa.Columns(11).Locked = False: reixa.Columns(10).Locked = False: reixa.Columns(14).Locked = False




End Sub
Private Sub reixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 Dim valtmp As String
 
    If reixa.col = 0 Then reixa.EditActive = False
 '-------
 bloquejar_camps_innecesaris
 If Not laminadores.Recordset.EOF Then
 'texteimpresio = atrim(impresores.Recordset!texteimpresio)
  If atrim(laminadores.Recordset!tipus) = "F" Then
     framebobines.Enabled = True
       Else: framebobines.Enabled = False: framepantones.Visible = False
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
 
 frameempalmes.Visible = False
 framepantones.Visible = False
 
End Sub

Private Sub reixabobines_AfterColUpdate(ByVal ColIndex As Integer)
 If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
 On Error Resume Next
 bobines.Recordset.Fields(reixabobines.Columns(ColIndex).DataField) = reixabobines.Columns(ColIndex).Text
 reixabobines.EditActive = False
'bobines.Recordset.Update
End Sub

Private Sub reixabobines_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    tempseditant = 0
End Sub

Private Sub reixabobines_ColEdit(ByVal ColIndex As Integer)
    tempseditant = 0
End Sub

Private Sub reixabobines_DblClick()
  Dim mtrsparcials As Double
 If reixabobines.col = 7 Then
   r = triar_observacio("B")
   If r <> "" Then reixabobines.Text = r
 End If
 If reixabobines.col = 1 Or reixabobines.col = 0 Then
  reixabobines.Text = escullir_operari
  If reixabobines.col = 0 Then
   nomoperari = UCase(r)
   numop = cadbl(reixabobines.Text)
  End If
  
 End If

 If reixabobines.col = 6 Then
    mtrsparcials = cadbl(InputBox("Entra els metres parcials.", "Metres parcials"))
    If mtrsparcials > 0 Then
        If MsgBox("Segur que vols modificar els metres parcials?", vbCritical + vbYesNo, "Atenció") = vbYes Then
          reixabobines.EditActive = True
          reixabobines.Text = atrim(mtrsparcials)
          reixabobines.EditActive = False
          calcular_totals
        End If
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
   formseleccio.Data1.RecordSource = "select codi,descripcio from operaris where maquina='L' and actiu<>0"
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
 frameempalmes.Visible = False
 framepantones.Visible = False
 If reixabobines.col <> 7 Then
     framepantones.Visible = False
     frameempalmes.Visible = False
     framebobentrada.Visible = True
   Else: framebobentrada.Visible = False
 End If
End Sub

Private Sub reixabobines_KeyDown(KeyCode As Integer, Shift As Integer)
    tempseditant = Now
End Sub

Private Sub reixabobines_LostFocus()
Dim camps As String
camps = "Command7Command9Command12Command13Command3Command5Command6"

If reixabobines.col > 1 And InStr(1, camps, Screen.ActiveControl.Name) = 0 Then calcular_totals
End Sub

Private Sub reixabobines_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Static fila As Double
If IsNull(fila) Then fila = 0
If fila <> reixabobines.row Then
 'calcular_totals
End If
fila = reixabobines.row
If reixabobines.col <> 8 Then
     framepantones.Visible = False
     frameempalmes.Visible = False
     framebobentrada.Visible = True
   Else: framebobentrada.Visible = False
 End If
End Sub

Private Sub reixaempalmes_AfterColEdit(ByVal ColIndex As Integer)
  tempseditant = 0
End Sub

Private Sub reixaempalmes_AfterUpdate()
  If bobines.Recordset.EditMode = 0 And Not bobines.Recordset.EOF Then
    bobines.Recordset.Edit
    bobines.Recordset!numempalmes = empalmes.Recordset.RecordCount
    bobines.Recordset.Update
  End If
End Sub

Private Sub reixaempalmes_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
  'tempseditant = 0
End Sub

Private Sub reixaempalmes_ColEdit(ByVal ColIndex As Integer)
'tempseditant = 0
End Sub

Private Sub reixaempalmes_DblClick()
If reixaempalmes.col = 1 Then
   r = triar_observacio("S")
   If r <> "" Then reixaempalmes.Text = r
End If
End Sub

Private Sub reixaempalmes_KeyDown(KeyCode As Integer, Shift As Integer)
    tempseditant = Now
End Sub

Private Sub reixaempalmes_OnAddNew()
 empalmes.Recordset!id = bobines.Recordset!id
 'reixa.col = 0
 
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

Sub vigilartipusdecola()
   If ettipuscola.Tag = "1" Then
        ettipuscola.Visible = Not ettipuscola.Visible
          Else: ettipuscola.Visible = False
   End If
End Sub
Private Sub rellotge_Timer()
  Static tempsoperari As Byte
'  Static ultimarow As Double
'  If ultimarow = 0 Then ultimarow = reixa.Row
'  If ultimarow <> reixa.Row Then
'     ultimarow = reixa.Row: calcular_totals
 ' End If
 mirarsiparar
 comprovareditantbobines
 vigilartipusdecola
 On Error GoTo error_screen
 If Screen.ActiveControl.Name = "akjdfks" Then Me.Caption = Me.Caption
 On Error GoTo 0
 If client.Caption = "" And (laminadores.Recordset.BOF And laminadores.Recordset.EOF) Then
   carregar_client_ntintersialtres
 End If
 
 If numop = 0 And Not formseleccio.Visible And reixa.Enabled Then
   numop = escullir_operari
   nomoperari = UCase(r)
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
  rellotge.Tag = cadbl(rellotge.Tag) + 1
  If rellotge.Tag = "10" Then
    'calcular_totals
    rellotge.Tag = "0"
  End If
  
  If Not laminadores.Recordset.EOF Then
    Select Case atrim(laminadores.Recordset!tipus)
       
       Case "C"
          Command1.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor
          Command2.BackColor = &HFF8080
       Case "F"
          Command1.BackColor = Command4.BackColor: Command2.BackColor = Command4.BackColor
          Command3.BackColor = &HFF8080
        Case Else
           Command1.BackColor = Command4.BackColor: Command2.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor
    End Select
    If Screen.ActiveForm.Name = "capcalera" Then
      Command2.BackColor = Command4.BackColor: Command3.BackColor = Command4.BackColor
          Command1.BackColor = &HFF8080
    End If
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
  '
  'End If
  Exit Sub
error_screen:
'MsgBox "Error d'Screen en el Timer"
'End
End Sub
Sub comprovareditantbobines()
  If DateDiff("s", tempseditant, Now) > 15 And tempseditant > 0 Then
   On Error Resume Next
   If Not bobines.Recordset.EOF Then
    If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
    bobines.Recordset.Update
   End If
   bobines.UpdateControls
   If Not laminadores.Recordset.EOF Then
    If laminadores.Recordset.EditMode = 0 Then laminadores.Recordset.Edit
    
    laminadores.Recordset.Update
   End If
   laminadores.UpdateControls
   If Not empalmes.Recordset.EOF Then
    If empalmes.Recordset.EditMode = 0 Then empalmes.Recordset.Edit
    empalmes.Recordset.Update
   End If
   empalmes.UpdateControls
      
   tempseditant = 0
  End If
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

Sub carregar_pdf(vnumtreball As Double, vordre As Double)
   Dim generarfitxer_pdf As String
   Dim ruta_documentacio_clixes As String
   ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
   generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(vnumtreball, "00000") + "\pdf" + Format(vnumtreball, "00000") + "-" + Format(vordre, "000") + ".pdf"
  
   If existeix(generarfitxer_pdf) Then
       AcroPDF1.OpenFile generarfitxer_pdf
       AcroPDF1.ZOrder 0
        Else
          AcroPDF1.OpenFile rutadelfitxer(cami) + "pdfblanc.pdf"
   End If
   
End Sub
