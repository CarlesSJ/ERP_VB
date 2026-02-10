VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmclixes 
   Caption         =   "Manteniment de Clixes"
   ClientHeight    =   9090
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11940
   Icon            =   "clixes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framealb 
      Caption         =   "        Albarans"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   585
      TabIndex        =   52
      Top             =   4575
      Width           =   11310
      Begin MSDBGrid.DBGrid reixaalbarans 
         Bindings        =   "clixes.frx":74F2
         Height          =   3600
         Left            =   135
         OleObjectBlob   =   "clixes.frx":7505
         TabIndex        =   54
         Tag             =   "albarans"
         Top             =   210
         Width           =   10890
      End
   End
   Begin VB.CommandButton sortir 
      Height          =   390
      Left            =   11295
      Picture         =   "clixes.frx":8A92
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Sortir"
      Top             =   165
      Width           =   390
   End
   Begin MSComCtl2.DTPicker picker 
      Height          =   315
      Left            =   645
      TabIndex        =   58
      Top             =   4350
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/mm/yy"
      Format          =   662831107
      CurrentDate     =   41303
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   30
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command8 
      Height          =   360
      Left            =   75
      Picture         =   "clixes.frx":901C
      Style           =   1  'Graphical
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   8205
      Width           =   375
   End
   Begin VB.Data albarans 
      Caption         =   "albarans"
      Connect         =   "Access"
      DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\Clixes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4425
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"clixes.frx":95A6
      Top             =   4305
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data modifis 
      Caption         =   "modifis"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2265
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clixes_modifi"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame framebotons1 
      Height          =   1155
      Left            =   30
      TabIndex        =   48
      Top             =   4590
      Width           =   435
      Begin VB.CommandButton Command10 
         Height          =   360
         Left            =   45
         Picture         =   "clixes.frx":964D
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Albarans del Treball"
         Top             =   645
         Width           =   345
      End
      Begin VB.CommandButton Command9 
         Height          =   360
         Left            =   45
         Picture         =   "clixes.frx":9BD7
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Modificacions Treballs"
         Top             =   210
         Width           =   345
      End
   End
   Begin VB.Frame dadesclixes 
      Caption         =   "Dades Clixes"
      Height          =   3720
      Left            =   90
      TabIndex        =   9
      Top             =   675
      Width           =   11670
      Begin VB.CommandButton Command14 
         Height          =   345
         Left            =   10740
         Picture         =   "clixes.frx":A161
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Intercanviar dates si ho hi ha data d'Entrega"
         Top             =   3150
         Width           =   330
      End
      Begin VB.TextBox Text11 
         DataField       =   "datamodificacio"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   9045
         TabIndex        =   73
         Top             =   3315
         Width           =   1350
      End
      Begin VB.CommandButton Command13 
         Height          =   285
         Left            =   10410
         Picture         =   "clixes.frx":A6EB
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Borrar la data de Modificació."
         Top             =   3360
         Width           =   270
      End
      Begin VB.TextBox Text10 
         DataField       =   "dataaprovdissenys"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   9045
         TabIndex        =   70
         Top             =   2415
         Width           =   1350
      End
      Begin VB.CommandButton Command12 
         Height          =   285
         Left            =   10410
         Picture         =   "clixes.frx":AC75
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Borrar la data d'Entrega"
         Top             =   2415
         Width           =   270
      End
      Begin VB.TextBox Text9 
         DataField       =   "dataprevclixes"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   9045
         TabIndex        =   67
         Top             =   2700
         Width           =   1350
      End
      Begin VB.CommandButton Command11 
         Height          =   285
         Left            =   10410
         Picture         =   "clixes.frx":B1FF
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Borrar la data d'Entrega"
         Top             =   2715
         Width           =   270
      End
      Begin VB.CommandButton Command3 
         Height          =   285
         Left            =   4680
         Picture         =   "clixes.frx":B789
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Borrar la data d'Inici de Treball"
         Top             =   555
         Width           =   270
      End
      Begin VB.CommandButton borrardataentrega 
         Height          =   285
         Left            =   10410
         Picture         =   "clixes.frx":BD13
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Borrar la data d'Entrega"
         Top             =   3045
         Width           =   270
      End
      Begin MSMask.MaskEdBox bandes 
         DataField       =   "bandesclixes"
         DataSource      =   "clixes"
         Height          =   300
         Left            =   9585
         TabIndex        =   63
         Top             =   1395
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         _Version        =   327681
         Format          =   "0"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox sistemaimpresio 
         DataField       =   "sistemadimpresio"
         DataSource      =   "clixes"
         Height          =   315
         ItemData        =   "clixes.frx":C29D
         Left            =   7605
         List            =   "clixes.frx":C2AA
         TabIndex        =   22
         Top             =   1380
         Width           =   1680
      End
      Begin VB.TextBox idestatclixe 
         DataField       =   "id_estatclixe"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   7320
         TabIndex        =   57
         Top             =   255
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Timer Timer1 
         Interval        =   400
         Left            =   165
         Top             =   240
      End
      Begin VB.TextBox Text8 
         DataField       =   "forma_imprimir"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   6705
         TabIndex        =   50
         Top             =   240
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tenim el Clixe Nosaltres."
         DataField       =   "tenimclixe"
         DataSource      =   "clixes"
         Height          =   510
         Left            =   10230
         TabIndex        =   47
         Top             =   1230
         Width           =   1410
      End
      Begin VB.TextBox observacio 
         DataField       =   "observacio"
         DataSource      =   "clixes"
         Height          =   330
         Left            =   1305
         MaxLength       =   255
         TabIndex        =   33
         Top             =   3195
         Width           =   6375
      End
      Begin VB.TextBox Text7 
         DataField       =   "dataentrega"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   9045
         TabIndex        =   32
         Top             =   3000
         Width           =   1350
      End
      Begin VB.TextBox Text6 
         DataField       =   "codibarres"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   4980
         TabIndex        =   31
         Top             =   2820
         Width           =   2325
      End
      Begin VB.TextBox Text5 
         DataField       =   "montadora"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   3735
         TabIndex        =   30
         Top             =   2820
         Width           =   1080
      End
      Begin VB.CommandButton Command7 
         Height          =   315
         Left            =   3255
         Picture         =   "clixes.frx":C2CE
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2820
         Width           =   315
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         DataField       =   "link_pdf"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2820
         Width           =   1590
      End
      Begin VB.CommandButton Command6 
         Height          =   315
         Left            =   2910
         Picture         =   "clixes.frx":C858
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2820
         Width           =   315
      End
      Begin VB.TextBox numsuportdigital 
         DataField       =   "numarxiusop"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   9045
         TabIndex        =   28
         Top             =   2115
         Width           =   1455
      End
      Begin VB.ComboBox nomproveidor 
         Height          =   315
         Left            =   5580
         Locked          =   -1  'True
         TabIndex        =   27
         Tag             =   "proveidor"
         Top             =   2100
         Width           =   2895
      End
      Begin VB.ComboBox nomrepresentant 
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   61
         Tag             =   "representant"
         Top             =   2085
         Width           =   2895
      End
      Begin VB.ComboBox liniaproducte 
         DataField       =   "id_liniaproducte"
         DataSource      =   "clixes"
         Height          =   315
         Left            =   1320
         TabIndex        =   25
         Top             =   1725
         Width           =   6180
      End
      Begin VB.ComboBox marcaproducte 
         DataField       =   "id_marca"
         DataSource      =   "clixes"
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Top             =   1380
         Width           =   6180
      End
      Begin VB.ComboBox nomclient 
         Height          =   315
         ItemData        =   "clixes.frx":CDE2
         Left            =   1320
         List            =   "clixes.frx":CDE4
         Locked          =   -1  'True
         TabIndex        =   23
         Tag             =   "id_client"
         Text            =   "nomclient"
         Top             =   945
         Width           =   6180
      End
      Begin VB.CheckBox actiu 
         Caption         =   "Actiu?"
         DataField       =   "actiu"
         DataSource      =   "clixes"
         Height          =   255
         Left            =   10710
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   195
         Width           =   870
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "clixes.frx":CDE6
         Left            =   7575
         List            =   "clixes.frx":CDE8
         Locked          =   -1  'True
         TabIndex        =   21
         Tag             =   "id_estatclixe"
         Top             =   540
         Width           =   2625
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "clixes.frx":CDEA
         Left            =   5070
         List            =   "clixes.frx":CDF4
         TabIndex        =   3
         Tag             =   "forma_imprimir"
         Top             =   540
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         DataField       =   "datainicitreball"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   3345
         TabIndex        =   2
         Top             =   540
         Width           =   1305
      End
      Begin VB.TextBox Text2 
         DataField       =   "arxiuclixe"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   2340
         TabIndex        =   1
         Top             =   540
         Width           =   885
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         DataField       =   "id_treball"
         DataSource      =   "clixes"
         Height          =   285
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   540
         Width           =   885
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Modificació:"
         Height          =   345
         Left            =   7785
         TabIndex        =   74
         Top             =   3345
         Width           =   1380
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Data aprovació:"
         Height          =   345
         Left            =   7785
         TabIndex        =   71
         Top             =   2445
         Width           =   1380
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Data prevista:"
         Height          =   345
         Left            =   7785
         TabIndex        =   68
         Top             =   2730
         Width           =   1380
      End
      Begin VB.Label Label18 
         Caption         =   "Bandes:"
         Height          =   210
         Left            =   9525
         TabIndex        =   60
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label Label17 
         Caption         =   "Sistema d'impresió:"
         Height          =   285
         Left            =   7770
         TabIndex        =   59
         Top             =   1110
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Data d'Entrega:"
         Height          =   345
         Left            =   7785
         TabIndex        =   46
         Top             =   3030
         Width           =   1380
      End
      Begin VB.Label Label15 
         Caption         =   "Codi de Barres"
         Height          =   270
         Left            =   5265
         TabIndex        =   45
         Top             =   2565
         Width           =   1275
      End
      Begin VB.Label Label14 
         Caption         =   "Montadora"
         Height          =   210
         Left            =   3855
         TabIndex        =   44
         Top             =   2580
         Width           =   1005
      End
      Begin VB.Label Label13 
         Caption         =   "Nº Suport Digital"
         Height          =   240
         Left            =   9150
         TabIndex        =   41
         Top             =   1815
         Width           =   1425
      End
      Begin VB.Label Label12 
         Caption         =   "Fotogravador:"
         Height          =   240
         Left            =   4530
         TabIndex        =   40
         Top             =   2130
         Width           =   1260
      End
      Begin VB.Label Label11 
         Caption         =   "Observació:"
         Height          =   285
         Left            =   45
         TabIndex        =   39
         Top             =   3270
         Width           =   945
      End
      Begin VB.Label Label10 
         Caption         =   "Ruta fitxer PDF:"
         Height          =   315
         Left            =   45
         TabIndex        =   38
         Top             =   2805
         Width           =   1365
      End
      Begin VB.Label Label9 
         Caption         =   "Representant:"
         Height          =   240
         Left            =   60
         TabIndex        =   37
         Top             =   2145
         Width           =   1260
      End
      Begin VB.Label Label8 
         Caption         =   "Linia Producte:"
         Height          =   285
         Left            =   45
         TabIndex        =   36
         Top             =   1785
         Width           =   1500
      End
      Begin VB.Label Label7 
         Caption         =   "Marca Producte:"
         Height          =   300
         Left            =   45
         TabIndex        =   35
         Top             =   1395
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Client Principal:"
         Height          =   285
         Left            =   45
         TabIndex        =   34
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Estat del Clixé"
         Height          =   345
         Left            =   8130
         TabIndex        =   19
         Top             =   270
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Forma d'Imprimir"
         Height          =   285
         Left            =   5280
         TabIndex        =   18
         Top             =   270
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Data d'Inici del Treball"
         Height          =   360
         Left            =   3330
         TabIndex        =   17
         Top             =   270
         Width           =   1875
      End
      Begin VB.Label Label2 
         Caption         =   "Arxiu del Clixe"
         Height          =   225
         Left            =   2235
         TabIndex        =   16
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "NºTreball"
         Height          =   225
         Left            =   1440
         TabIndex        =   15
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   11700
      Begin VB.CommandButton Command16 
         Caption         =   "Posar Treball a Comanda"
         Height          =   285
         Left            =   8715
         TabIndex        =   77
         Top             =   210
         Width           =   2415
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Intern -netejar dates fi de modificacions que no hhi ha i no es la ultima"
         Height          =   330
         Left            =   7875
         TabIndex        =   76
         Top             =   195
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton consultar 
         Height          =   360
         Left            =   1905
         Picture         =   "clixes.frx":CE0F
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "Buscar Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   390
         Left            =   10020
         TabIndex        =   51
         Top             =   135
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Data clixes 
         Caption         =   "Clixes"
         Connect         =   "Access"
         DatabaseName    =   "M:\progcomandes\dades\Clixes.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4785
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   195
         Width           =   3150
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "clixes.frx":D399
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   989
         Picture         =   "clixes.frx":D923
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   532
         Picture         =   "clixes.frx":DEAD
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Modificar Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   1446
         Picture         =   "clixes.frx":E437
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Actualitzar Registres"
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
         Left            =   2655
         TabIndex        =   8
         Top             =   195
         Width           =   2025
      End
   End
   Begin VB.Frame framebotons2 
      Height          =   1440
      Left            =   45
      TabIndex        =   11
      Top             =   7230
      Width           =   450
      Begin VB.CommandButton Command5 
         Height          =   360
         Left            =   30
         Picture         =   "clixes.frx":E9C1
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   570
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Height          =   360
         Left            =   30
         Picture         =   "clixes.frx":EF4B
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame framemodi 
      Caption         =   "           Modificacions"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   630
      TabIndex        =   10
      Top             =   4650
      Width           =   11295
      Begin MSDBGrid.DBGrid reixamodifis 
         Bindings        =   "clixes.frx":F4D5
         Height          =   3750
         Left            =   90
         OleObjectBlob   =   "clixes.frx":F4E7
         TabIndex        =   53
         Tag             =   "modifis"
         Top             =   225
         Width           =   11145
      End
   End
   Begin VB.Menu m_manteniments 
      Caption         =   "Manteniments"
      Begin VB.Menu m_fotogravadors 
         Caption         =   "Fotogravadors"
      End
      Begin VB.Menu mestatdclixes 
         Caption         =   "Estats de clixés"
      End
   End
   Begin VB.Menu mllistats 
      Caption         =   "Llistats"
      Begin VB.Menu m_pendentsdefacturar 
         Caption         =   "Pendents de facturar"
      End
      Begin VB.Menu llistatmodifiacionspendentsacabar 
         Caption         =   "Modificacions pendents d'acabar."
      End
   End
End
Attribute VB_Name = "frmclixes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function rutadelfitxer(cam As String) As String
   Dim c As Byte
   c = 0
   While InStr(c + 1, cam, "\") <> 0
    c = InStr(c + 1, cam, "\")
   Wend
   If c = 0 Then c = Len(cam)
   rutadelfitxer = Mid(cam, 1, c)
End Function
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



Private Sub alta_Click()
  If nouregistre Then
     Text1 = elclixemesgran + 1
     'gravar_canvis
     'modificar_Click
  End If
End Sub
Function elclixemesgran() As Integer
   Dim rstcli As Recordset
    elclixemesgran = 0
    Set rstcli = dbtmp.OpenRecordset("select max(id_treball) as gran from clixes")
    If Not rstcli.EOF Then
      elclixemesgran = rstcli!gran
    End If
    Set rstcli = Nothing
End Function
Sub activarframes(estat As Boolean)
  dadesclixes.Enabled = estat
  'framemodi.Enabled = estat
  framebotons2.Enabled = estat
  'framebotons1.Enabled = True
  bloquejarreixa reixaalbarans
  bloquejarreixa reixamodifis
  buscant = False
  Text1.Locked = True
End Sub
Sub bloquejarreixa(reixa As Object)
   Dim estat As Boolean
   Dim pos As Double
   estat = dadesclixes.Enabled
   For i = 0 To reixa.Columns.Count - 1
     reixa.Columns(i).Locked = Not estat
   Next i
End Sub

Private Sub borrardataentrega_Click()
  borrardata "dataentrega"
End Sub
Sub borrardata(camp As String)
  gravar_canvis
  dbtmp.Execute "update clixes set " + camp + "=null where id_treball=" + atrim(clixes.Recordset!id_treball)
  clixes.UpdateControls
  modificar_Click
  
End Sub
Private Sub clixes_Reposition()
If Not clixes.Recordset.EOF Then
   clixes.Caption = "Clixes " + atrim(clixes.Recordset.AbsolutePosition + 1) + " / " + atrim(clixes.Recordset.RecordCount)
  Else: clixes.Caption = "Clixes"
End If
activarframes False
  actualizar_links
'carregar_lookupalbarans
End Sub

Sub possar_color_tipusimpresio(tipus As String)
  Dim color As Long
  Dim colorclassic As Long
  Dim colorkodak As Long
  Dim coloroffset As Long
  colorclassic = &H8000000F
  colorkodak = &HC0FFFF
  coloroffset = &HFDD7FD
  If tipus = "Flexo Std" Then color = colorclassic
  If tipus = "Flexo Kodak" Then color = colorkodak
  If tipus = "Offset" Then color = coloroffset
  If color = 0 Then color = colorclassic
  possar_color_frames color
  
End Sub
Sub possar_color_frames(color As Long)
   
   dadesclixes.BackColor = color
   For i = 0 To frmclixes.Controls.Count - 1
      If TypeOf frmclixes.Controls(i) Is CheckBox Then frmclixes.Controls(i).BackColor = color
   Next i
   
End Sub


'Sub carregar_lookupalbarans()
'  Dim rstdesc As Recordset
'  If albarans.Recordset.EOF Then Exit Sub
' albarans.Recordset.MoveFirst
''  While Not albarans.Recordset.EOF
'      Set rstdesc = dbtmp.OpenRecordset("select * from clixes_detallsalb where id_detall=" + atrim(cadbl(albarans.Recordset!id_detall)))
'      If Not rstdesc.EOF Then
'          reixaalbarans.Columns("descripcio") = rstdesc!descripcio
'      End If
'      albarans.Recordset.MoveNext
 ' Wend
'End Sub
Sub actualizar_links()
   Dim rstlinks As Recordset
   
   modifis.RecordSource = "select * from clixes_modifi where id_treball=" + atrim(cadbl(clixes.Recordset!id_treball)) + " order by ordre"
   modifis.Refresh
   dbtmp.Execute "delete * from clixes_albarans where id_detall is null"
   r = "SELECT Clixes_albarans.*, Clixes_detallsalb.descripcio FROM Clixes_detallsalb INNER JOIN Clixes_albarans ON Clixes_detallsalb.id_detall = Clixes_albarans.id_detall"
   albarans.RecordSource = r + "  where id_treball=" + atrim(cadbl(clixes.Recordset!id_treball)) + " order by ordre"
   albarans.Refresh
   'If nomcontrolactiu <> "reixaalbarans" And nomcontrolactiu <> "reixamodifi" And clixes.Recordset.EditMode = 0 Then
   ' If Not modifis.Recordset.EOF Then ensenyar_frame "framemodi"
   ' If Not albarans.Recordset.EOF Then ensenyar_frame "framealb"
   'End If
   
   ' estats clixes
   Set rstlinks = dbtmp.OpenRecordset("select * from clixes_estats where id_estat=" + atrim(cadbl(clixes.Recordset!id_estatclixe)))
   If Not rstlinks.EOF Then
       Combo2.Text = atrim(rstlinks!descripcio)
         Else: Combo2.Text = ""
   End If
DoEvents
   ' nom client
   Set rstlinks = dbtmpb.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(clixes.Recordset!id_client)))
   If Not rstlinks.EOF Then
       nomclient.Text = atrim(rstlinks!nom)
         Else: nomclient.Text = ""
   End If
   'nom representant
   Set rstlinks = dbtmpb.OpenRecordset("select nom from representants where codi=" + atrim(cadbl(clixes.Recordset!representant)))
   If Not rstlinks.EOF Then
       nomrepresentant.Text = atrim(rstlinks!nom)
         Else: nomrepresentant.Text = ""
   End If
   'nom fotogravador
   Set rstlinks = dbtmp.OpenRecordset("select nomfotogravador from fotogravadors where codi=" + atrim(cadbl(clixes.Recordset!proveidor)))
   If Not rstlinks.EOF Then
       nomproveidor.Text = atrim(rstlinks!nomfotogravador)
         Else: nomproveidor.Text = ""
   End If
   
End Sub
Function nomcontrolactiu() As String
  On Error Resume Next
  nomcontrolactiu = Screen.ActiveControl.Name
End Function
Private Sub Combo1_Click()
  Text8.Text = Mid(Combo1, 1, 1)
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_DropDown()
Load formseleccio
   formseleccio.Data1.DatabaseName = clixes.DatabaseName
   formseleccio.Data1.RecordSource = "select * from clixes_estats"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           Combo2.Text = formseleccio.DBGrid2.Columns("descripcio")
           idestatclixe.Text = formseleccio.DBGrid2.Columns("id_estat")
        End If
   End If
   If seleccioret = 9 Then
           Combo2.Text = ""
           idestatclixe.Text = Null
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub Command1_Click()
   gravar_canvis
End Sub
Function comprovarcamps() As Boolean
  comprovarcamps = True
  If (atrim(Text7) <> "" And Not IsDate(Text7)) Or (atrim(Text3) <> "" And Not IsDate(Text3)) Then
    comprovarcamps = False
    MsgBox "Hi ha una Data errònea", vbCritical, "Atenció"
  End If
End Function
Sub gravar_canvis()
  
If Not comprovarcamps Then Exit Sub
gravant = True
 If clixes.Recordset.EditMode > 0 Then
  If Not buscant Then
   'If albarans.Recordset.EditMode > 0 Then albarans.Recordset.Update
   'If modifis.Recordset.EditMode > 0 Then modifis.Recordset.Update
   If Not gravar_reixa(reixaalbarans) Then Exit Sub
   If Not gravar_reixa(reixamodifis) Then Exit Sub
   If Not (modifis.Recordset.EOF And modifis.Recordset.BOF) Then modifis.Recordset.MoveLast
   If Not modifis.Recordset.EOF Then
      clixes.Recordset!id_estatclixe = modifis.Recordset!id_estatclixe
   End If
   If bandes = "" Then bandes = 0
    'clixes.Recordset.Fields(Text7.DataField).Value = Null
   clixes.Recordset.Update
   clixes.Recordset.Bookmark = clixes.Recordset.LastModified
   
   activarframes False
     Else: finalitzarbusqueda
  End If
 End If

 gravant = False
End Sub

Private Sub Command10_Click()
   ensenyar_frame "framealb"
End Sub

Private Sub Command11_Click()
  borrardata "dataprevclixes"
End Sub

Private Sub Command12_Click()
borrardata "dataaprovdissenys"
End Sub

Private Sub Command13_Click()
borrardata "datamodificacio"
End Sub

Private Sub Command14_Click()
  Dim datatmp As String
  If Len(Text7) > 0 Then
     If Len(Text11) > 0 Then If MsgBox("Hi ha una data de modificació." + Chr(10) + "Vols sobre escriure-la amb la d'entrega?", vbCritical + vbYesNo, "Atenció") = vbNo Then Exit Sub
     datatmp = Text7
     gravar_canvis
     
     dbtmp.Execute "update clixes set datamodificacio=#" + Format(datatmp, "mm/dd/yy") + "# where id_treball=" + atrim(clixes.Recordset!id_treball)
     'clixes.UpdateControls
     borrardata "dataentrega"
     Exit Sub
  End If
  If Len(Text11) > 0 And Len(Text7) = 0 Then
     datatmp = Text11
     gravar_canvis
     dbtmp.Execute "update clixes set dataentrega=#" + Format(datatmp, "mm/dd/yy") + "# where id_treball=" + atrim(clixes.Recordset!id_treball)
     'clixes.UpdateControls
     borrardata "datamodificacio"
     
     Exit Sub
  End If
     
End Sub

Private Sub Command15_Click()
   clixes.Recordset.MoveLast
   While Not clixes.Recordset.BOF
     If Not modifis.Recordset.EOF Then
       modifis.Recordset.MoveLast
       If Not modifis.Recordset.EOF Then modifis.Recordset.MovePrevious
       While Not modifis.Recordset.BOF
        
         If IsNull(modifis.Recordset!data_fi) Then
           modifis.Recordset.Edit
            modifis.Recordset!data_fi = modifis.Recordset!data_inici
           modifis.Recordset.Update
         End If
         modifis.Recordset.MovePrevious
       Wend
     End If
     clixes.Recordset.MovePrevious
      DoEvents
   Wend
End Sub

Private Sub Command16_Click()
   Dim col As Column
   Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select numtreball,comanda,codibarras,obsreb2 as triar from comandes where codibarras='" + atrim(Text6) + "' and (numtreball=0 or numtreball=null)"
   formseleccio.DBGrid2.AllowDelete = False
   
   formseleccio.DBGrid2.MarqueeStyle = 6
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns(0).Visible = False
   formseleccio.DBGrid2.Columns(1).Locked = True
   formseleccio.DBGrid2.Columns(2).Locked = True
   formseleccio.DBGrid2.Columns(3).Width = 500
   formseleccio.DBGrid2.AllowUpdate = True
   formseleccio.Show 1
   If seleccioret = 1 Then
        formseleccio.Data1.Recordset.MoveFirst
        While Not formseleccio.Data1.Recordset.EOF
            If UCase(atrim(formseleccio.Data1.Recordset!triar)) <> "N" Then
               formseleccio.Data1.Recordset.Edit
               formseleccio.Data1.Recordset!numtreball = clixes.Recordset!id_treball
'               MsgBox formseleccio.Data1.Recordset!comanda
               formseleccio.Data1.Recordset.Update
            End If
            formseleccio.Data1.Recordset.MoveNext
        Wend
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub Command2_Click()
   Dim rst As Recordset
    nomfitxer = "c:\clixes\traspasclixes.txt"
    Kill nomfitxer
    Open nomfitxer For Output As #1
    MsgBox "Aquest proces actualitza els clients antics a els nous, fa falta tenir a la bd clixes una taula de clients de la taula antiga"
    clixes.Refresh
   While Not clixes.Recordset.EOF
     'Set rst = dbtmp.OpenRecordset("select * from clients where campo1=" + atrim(cadbl(clixes.Recordset!id_client)))
     'Set rsttmp = dbtmpb.OpenRecordset("select * from clients where codi=" + atrim(cadbl(rst!campo29)))
     'If Not rsttmp.EOF Then
     '   clixes.Recordset.Edit
     '     clixes.Recordset!id_client = cadbl(rsttmp!codi)
     '   clixes.Recordset.Update
     'End If
     Set rsttmp = dbtmpb.OpenRecordset("select * from clients where codi=" + atrim(cadbl(clixes.Recordset!id_client)))
     If rsttmp.EOF Then
       r = atrim(cadbl(clixes.Recordset!id_treball)) + " "
       r = r + atrim(cadbl(clixes.Recordset!id_client))
       Print #1, r
       DoEvents
     End If
     
     'Set rst = dbtmpb.OpenRecordset("select * from representants where codiintern=" + atrim(cadbl(clixes.Recordset!representant)))
     'If Not rst.EOF Then
     '   clixes.Recordset.Edit
     '   clixes.Recordset!representant = cadbl(rst!codi)
     '   clixes.Recordset.Update
     'End If
     clixes.Recordset.MoveNext
   Wend
    MsgBox "Si tot ha anat be elimina la taula clients de la bd de clixes"
    Close #1
End Sub

Private Sub Command3_Click()
 borrardata "datainicitreball"
End Sub

Private Sub Command4_Click()
  If MsgBox("Segur que vols borrar aquesta linia?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
      If framealb.Visible Then
          If Not albarans.Recordset.EOF Then
              albarans.Recordset.Delete
              albarans.Refresh
              gravar_canvis
          End If
           Else
                If Not modifis.Recordset.EOF Then
                       modifis.Recordset.Delete
                       modifis.Refresh
                       gravar_canvis
                End If
      End If
  End If
  
End Sub

Private Sub Command5_Click()
 Dim Control As String
 Dim bk As Double
 Control = frmclixes.ActiveControl.Name
 

 If frmclixes.Controls(Control).Visible Then frmclixes.Controls(Control).SetFocus
 If frmclixes.ActiveControl.Name = "reixaalbarans" And cadbl(albarans.Recordset!ordre) > 0 Then
  bk = albarans.Recordset!ordre
  'If albarans.Recordset.EditMode = 0 Then albarans.Recordset.Edit
  'albarans.Recordset.Update
  If Not gravar_reixa(reixaalbarans) Then Exit Sub
  
  If clixes.Recordset.EditMode > 0 Then clixes.Recordset.Update
  modificar_Click
  If frmclixes.Controls(Control).Visible Then frmclixes.Controls(Control).SetFocus
  albarans.Recordset.FindFirst "ordre=" + atrim(cadbl(bk))
  
 End If
 If frmclixes.ActiveControl.Name = "reixamodifis" And cadbl(modifis.Recordset!ordre) > 0 Then
   bk = modifis.Recordset!ordre
   'If modifis.Recordset.EditMode = 0 Then modifis.Recordset.Edit
   'modifis.Recordset.Update
   If Not gravar_reixa(reixamodifis) Then Exit Sub
   If Not (modifis.Recordset.EOF And modifis.Recordset.BOF) Then modifis.Recordset.MoveLast
   If Not modifis.Recordset.EOF Then
      clixes.Recordset!id_estatclixe = modifis.Recordset!id_estatclixe
   End If
   If clixes.Recordset.EditMode > 0 Then clixes.Recordset.Update
   modificar_Click
   If frmclixes.Controls(Control).Visible Then frmclixes.Controls(Control).SetFocus
   modifis.Recordset.FindFirst "ordre=" + atrim(cadbl(bk))
   
 End If
 
End Sub
Function gravar_reixa(reixa As DBGrid) As Boolean
    Dim fila As Double
    gravar_reixa = True
    If reixa.Visible And reixa.Row > 0 Then
     fila = reixa.Row
     reixa.SetFocus
     SendKeys "{down}"
     DoEvents
     If reixa.Row <> fila + 1 Then gravar_reixa = False
     reixa.Row = fila
    End If
    
End Function
Private Sub Command6_Click()
 Dim nomfitxer As String
 r = Format(clixes.Recordset!id_client, "000000")
 
 r = busca_nomdirectori_codiclient(ruta_relativa_docs, r)
 
 If Not existeix(ruta_relativa_docs + "\" + r) Then r = ""
 r = obre_fitxer(ruta_relativa_docs + "\" + r + "\pdfs", 2)
 If Trim(r) <> "" Then
  nomfitxer = Mid(r, Len(ruta_relativa_docs) + 2)
  If Len(nomfitxer) < 255 Then
    Text4 = nomfitxer
      Else: MsgBox "El nom del fitxer es massa gran.", vbCritical, "Atenció"
  End If
 End If
 Text5.SetFocus
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
Private Sub Command7_Click()
 Dim ru As String
 If Text4.Text = "" Then Exit Sub
 ru = ""
 If InStr(1, Text4.Text, "\pdfs") = 0 Then
    ru = busca_nomdirectori_codiclient(ruta_relativa_docs, Format(clixes.Recordset!id_client, "000000"))
    ru = ru + "\pdfs\"
 End If
 r = ""
obrir_document r + Chr$(34) + ruta_relativa_docs + "\" + ru + Text4.Text + Chr$(34)
'MsgBox r + Chr$(34) + ruta_relativa_docs + "\" + Text4.Text + Chr$(34)

 
End Sub

Private Sub Command8_Click()
  'Set dbtmp = OpenDatabase(camiclixes)
  crear_taules_tmp
  If framealb.Visible Then imprimir_albarans
  If framemodi.Visible Then imprimir_modificacions
  'dbtmp.Close
End Sub
Sub imprimir_modificacions()
Dim rstimp As Recordset
  Set rstimp = dbtmp.OpenRecordset("tmp_clixes_capcalera")
  If Not rstimp.EOF Then
     While Not rstimp.EOF
       rstimp.Delete
       rstimp.MoveNext
     Wend
  End If
  rstimp.AddNew
  rstimp!id_treball = clixes.Recordset!id_treball
  rstimp!arxiuclixe = clixes.Recordset!arxiuclixe
  rstimp!datainici = clixes.Recordset!datainicitreball
  rstimp!formaimp = Combo1.Text
  rstimp!estatclixe = Combo2.Text
  rstimp!client = atrim(clixes.Recordset!id_client) + " - " + nomclient.Text
  rstimp!marca = clixes.Recordset!id_marca
  rstimp!linia = clixes.Recordset!id_liniaproducte
  rstimp!representant = nomrepresentant
  rstimp!proveidor = nomproveidor
  rstimp!montadora = clixes.Recordset!montadora
  rstimp!codibarres = clixes.Recordset!codibarres
  rstimp!dataentrega = clixes.Recordset!dataentrega
  rstimp!observacions = clixes.Recordset!observacio
  rstimp!sistemaimpresio = clixes.Recordset!sistemadimpresio
  rstimp!bandesclixes = clixes.Recordset!bandesclixes
  rstimp.Update
  
  Set rstimp = dbtmp.OpenRecordset("tmp_clixes_modifis_linies")
  If Not rstimp.EOF Then
     While Not rstimp.EOF
       rstimp.Delete
       rstimp.MoveNext
     Wend
  End If
  modifis.Recordset.MoveFirst
  While Not modifis.Recordset.EOF
   If atrim(modifis.Recordset!imprimiracomanda) <> "Sí" Then
    rstimp.AddNew
    rstimp!id_treball = modifis.Recordset!id_treball
    rstimp!descripcio = modifis.Recordset!descripcio
    rstimp!inici = modifis.Recordset!data_inici
    rstimp!fi = modifis.Recordset!data_fi
    rstimp.Update
   End If
   modifis.Recordset.MoveNext
  Wend
  llistat.Formulas(0) = ""
  llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "clixes_modificacions.rpt"
 llistat.DataFiles(0) = camiclixes
  llistat.DiscardSavedData = True
 llistat.Destination = crptToPrinter
 wait 1
 llistat.Action = 1
  
End Sub
Sub imprimir_albarans()
  Dim rstimp As Recordset
  Dim facturats As Boolean
  facturats = True
  If MsgBox("Vols imprimir els albarans NO FACTURATS?" + Chr(10) + Chr(13) + "Si prems No es faran els FACTURATS", vbYesNo + vbDefaultButton1, "Escull") = vbYes Then
      facturats = False
  End If
  Set rstimp = dbtmp.OpenRecordset("tmp_clixes_capcalera")
  If Not rstimp.EOF Then
     While Not rstimp.EOF
       rstimp.Delete
       rstimp.MoveNext
     Wend
  End If
  rstimp.AddNew
  rstimp!id_treball = clixes.Recordset!id_treball
  rstimp!arxiuclixe = clixes.Recordset!arxiuclixe
  rstimp!datainici = clixes.Recordset!datainicitreball
  rstimp!formaimp = Combo1.Text
  rstimp!estatclixe = Combo2.Text
  rstimp!client = atrim(clixes.Recordset!id_client) + " - " + nomclient.Text
  rstimp!marca = clixes.Recordset!id_marca
  rstimp!linia = clixes.Recordset!id_liniaproducte
  rstimp!representant = nomrepresentant
  rstimp!proveidor = nomproveidor
  rstimp!montadora = clixes.Recordset!montadora
  rstimp!codibarres = clixes.Recordset!codibarres
  rstimp!dataentrega = clixes.Recordset!dataentrega
  rstimp!observacions = clixes.Recordset!observacio
  rstimp!sistemaimpresio = clixes.Recordset!sistemadimpresio
  rstimp!bandesclixes = clixes.Recordset!bandesclixes
  rstimp.Update
  
  Set rstimp = dbtmp.OpenRecordset("tmp_clixes_albarans_linies")
  If Not rstimp.EOF Then
     While Not rstimp.EOF
       rstimp.Delete
       rstimp.MoveNext
     Wend
  End If
  albarans.Recordset.MoveFirst
  While Not albarans.Recordset.EOF
   If albarans.Recordset!facturat = facturats Then
    rstimp.AddNew
    rstimp!id_treball = albarans.Recordset!id_treball
    rstimp!Data = albarans.Recordset!Data
    rstimp!numalb = albarans.Recordset!num_alb
    rstimp!quantitat = albarans.Recordset!quantitat
    rstimp!descripcio = albarans.Recordset!descripcio
    rstimp!import = albarans.Recordset!import
    rstimp!facturat = IIf(albarans.Recordset!facturat, "Sí", "No")
    rstimp.Update
   End If
    albarans.Recordset.MoveNext
  Wend
  
  
   llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "clixes_albarans.rpt"
 llistat.DataFiles(0) = camiclixes
 If facturats Then
    llistat.Formulas(0) = "facturatsono='Albarans Facturats'"
   Else
     llistat.Formulas(0) = "facturatsono='Albarans NO Facturats'"
 End If
 llistat.DiscardSavedData = True
 llistat.Destination = crptToPrinter
 wait 1
 llistat.Action = 1
  
End Sub


Sub crear_taules_tmp()
  Dim camps(100, 2) As String
  taula_tmp = "tmp_clixes_capcalera"
  On Error Resume Next
   dbtmp.Execute "drop table " + taula_tmp
  On Error GoTo 0
  i = 1
  camps(i, 1) = "id_treball": camps(i, 2) = "integer": i = i + 1
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
  
  dbtmp.Execute ("create table " + taula_tmp + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbtmp.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
  'creo la taula de linies d'albarans
  taula_tmp = "tmp_clixes_albarans_linies"
  On Error Resume Next
   dbtmp.Execute "drop table " + taula_tmp
  On Error GoTo 0
  i = 1
  camps(i, 1) = "id_treball": camps(i, 2) = "integer": i = i + 1
  camps(i, 1) = "data": camps(i, 2) = "date": i = i + 1
  camps(i, 1) = "numalb": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "quantitat": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "descripcio": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "import": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "facturat": camps(i, 2) = "string": i = i + 1
  
  dbtmp.Execute ("create table " + taula_tmp + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbtmp.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
  
  'creo la taula de linies
  taula_tmp = "tmp_clixes_modifis_linies"
  On Error Resume Next
   dbtmp.Execute "drop table " + taula_tmp
  On Error GoTo 0
  i = 1
  camps(i, 1) = "descripcio": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "inici": camps(i, 2) = "date": i = i + 1
  camps(i, 1) = "fi": camps(i, 2) = "date": i = i + 1
  camps(i, 1) = "id_treball": camps(i, 2) = "integer": i = i + 1
  
  dbtmp.Execute ("create table " + taula_tmp + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbtmp.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
  
End Sub
Private Sub Command9_Click()
  ensenyar_frame "framemodi"
  
End Sub
Sub ensenyar_frame(nomf As String)
If nomf = "framemodi" Then
  framemodi.Visible = True
  framemodi.Top = 4440
  framemodi.Left = 465
  framemodi.ZOrder 0
  framealb.Visible = False
End If

If nomf = "framealb" Then
  framealb.Visible = True
  framealb.Top = 4440
  framealb.Left = 465
  framealb.ZOrder 0
  framemodi.Visible = False
End If

End Sub

Private Sub consultar_Click()
 If clixes.Recordset.EditMode > 0 Then clixes.Recordset.CancelUpdate
 If albarans.Recordset.EditMode > 0 Then albarans.Recordset.CancelUpdate
  If modifis.Recordset.EditMode > 0 Then modifis.Recordset.CancelUpdate
  If nouregistre Then
   Text1.Locked = False
   Text1.Text = ""
   Text1.SetFocus
   buscant = True
  End If
End Sub
Function nouregistre() As Boolean
  clixes.RecordSource = "clixes"
  clixes.Refresh
  nouregistre = True
  'If clixes.Recordset.EOF Then
  '  clixes.RecordSource = "clixes"
  '  clixes.Refresh
  '  If clixes.Recordset.EOF Then nouregistre = False: Exit Function
  'End If
  clixes.Recordset.AddNew
  dadesclixes.Enabled = True
  Text2.SetFocus
End Function

Private Sub eliminar_Click()
 Dim cli As Long
If clixes.Recordset.EOF Then Exit Sub
cli = clixes.Recordset!id_treball
 If Not modifis.Recordset.EOF Or Not albarans.Recordset.EOF Then
   MsgBox "No pots eliminar un clixe si te linies d'albarà i/o modificacions. Eliminales primer.", vbCritical, "Eliminar Clixe": Exit Sub
 End If
  If InputBox("Eliminar aquest clixe implica eliminar totes les relacions." + Chr(13) + Chr(10) + " ESCRIU [Eliminar] SI ESTAS SEGUR QUE HO VOLS FER.", "ATENCIO") = "Eliminar" Then
     If Not modifis.Recordset.EOF Then modifis.Recordset.MoveFirst
     While Not modifis.Recordset.EOF
       modifis.Recordset.Delete
       modifis.Recordset.MoveNext
     Wend
     If Not albarans.Recordset.EOF Then albarans.Recordset.MoveFirst
     While Not albarans.Recordset.EOF
       albarans.Recordset.Delete
       albarans.Recordset.MoveNext
     Wend
     If clixes.Recordset!id_treball = cli Then
      clixes.Recordset.Delete
     End If
     clixes.RecordSource = "clixes"
     clixes.Refresh
     If Not clixes.Recordset.EOF Then clixes.Recordset.MoveLast
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 112 Then
    gravant = True
     r = frmclixes.ActiveControl.Name
    If r = "reixaalbarans" Or r = "reixamodifis" Then
       Command5_Click
      Else: gravar_canvis
    End If
    gravant = False
  End If
  If KeyCode = 27 Then
      If albarans.Recordset.EditMode > 0 Then albarans.Recordset.CancelUpdate: reixaalbarans.ReBind
      If modifis.Recordset.EditMode > 0 Then modifis.Recordset.CancelUpdate
      If clixes.Recordset.EditMode > 0 Then clixes.Recordset.CancelUpdate
      
      
      activarframes False
  End If
  If Shift = 2 And KeyCode = Asc("0") Then
      consultar_Click
  End If
End Sub
Sub finalitzarbusqueda(Optional tipus As Byte)
 ratoli "espera"
 
 If cadbl(tipus) = 1 Then GoTo ficonsulta
 recorregutregistres
 If clixes.Recordset.EditMode > 0 Then clixes.Recordset.CancelUpdate
ficonsulta:
 activarframes False
 buscant = False
 If queryorder <> "" Then
     queryorder = " Order By " + queryorder
    Else: queryorder = " order by id_treball"
 End If
 If querywhere <> "" Then querywhere = " Where " + querywhere
 clixes.RecordSource = "select * from clixes " + querywhere + queryorder
 clixes.Refresh
 If Not clixes.Recordset.EOF Then clixes.Recordset.MoveLast: clixes.Recordset.MoveFirst
 ratoli "normal"
 'Unload subbusqueda
End Sub
Function triarordre(camp As String, valorord As String) As Boolean
  Dim ord As String
  triarordre = False
  If InStr(1, valorord, "<<") Then ord = camp + " " + " ASC"
  If InStr(1, valorord, ">>") Then ord = camp + " " + " DESC"
  If ord <> "" Then
      triarordre = True
    Else: Exit Function
  End If
  If queryorder = "" Then
     queryorder = ord
   Else: queryorder = queryorder + ", " + ord
  End If
  
End Function
Sub recorregutregistres()
 Dim objecte As Object
 queryorder = ""
 querywhere = ""
 'On Error Resume Next
 For Each objecte In Me
    If TypeOf objecte Is TextBox Or TypeOf objecte Is ComboBox Then
     If objecte.Tag = "9" Or objecte.Text <> "" Then
       If objecte.DataField <> "" Then
         If objecte.Text <> "" Then
           evaluarcontingut objecte.DataField, objecte.Text, clixes.Recordset.Fields(objecte.DataField).Type
           objecte.Text = ""
         End If
      End If
     End If
    End If
 Next
 
'exepcions
  If nomclient <> "" Then
    If querywhere <> "" Then querywhere = querywhere + " and "
    querywhere = querywhere + " id_client=" + atrim(cadbl(clixes.Recordset!id_client)) + " "
  End If
 
  If nomproveidor <> "" Then
    If querywhere <> "" Then querywhere = querywhere + " and "
    querywhere = querywhere + " proveidor=" + atrim(cadbl(clixes.Recordset!proveidor)) + " "
  End If
 
  If nomrepresentant <> "" Then
    If querywhere <> "" Then querywhere = querywhere + " and "
    querywhere = querywhere + " representant=" + atrim(cadbl(clixes.Recordset!representant)) + " "
  End If
 
End Sub

Function evaluarcontingut(camp As String, valor As String, tipusdato As Byte) As String
  Dim rest As String
  rest = ""
  evaluarcontingut = ""
  If triarordre(camp, valor) Then Exit Function
  If tipusdato = 10 Then
   If InStr(1, valor, "*") Or InStr(1, valor, "?") Then
      rest = " like '" + valor + "'"
     Else
       If InStr(1, valor, ">") Or InStr(1, valor, "<") Or InStr(1, valor, "=") Then
           If Mid(valor, 1, 2) = "<>" Then
             valor = Mid(valor, 1, 2) + "'" + Mid(valor, 3) + "'"
            Else: valor = Mid(valor, 1, 1) + "'" + Mid(valor, 2) + "'"
           End If
           rest = "" + valor + ""
        Else: rest = "=" + "'" + valor + "'"
       End If
   End If
  End If
  If tipusdato = 8 Then
    i = 1
    While Not IsNumeric(Mid(valor, i, 1))
     rest = rest + Mid(valor, i, 1)
     i = i + 1
    Wend
    If rest = "" Then rest = "="
    rest = rest + "#" + Format(Mid(valor, i, 50), "d/m/yyyy") + "#"
  End If
  If tipusdato <> 10 And tipusdato <> 8 Then
    valor = passaradecimalpunt(valor)
    If InStr(1, valor, ">") Or InStr(1, valor, "<") Or InStr(1, valor, "=") Then
           rest = atrim((valor))
        Else: rest = "=" + atrim((valor))
    End If
  End If
 
  evaluarcontingut = camp + rest
  
  rest = evaluarcontingut
  
  If querywhere = "" Then
     querywhere = rest
    Else
     querywhere = querywhere + " and " + rest + " "
  End If
  
End Function



Private Sub Form_KeyPress(KeyAscii As Integer)
  If Chr$(KeyAscii) = "'" Then KeyAscii = Asc("´")
End Sub

Private Sub Form_Load()
Dim arguments As Variant
arguments = ObtenerLíneaComando
fitxerini = "comandes.ini"
If atrim(arguments(1)) <> "" Then fitxerini = atrim(arguments(1))
On Error Resume Next
  Kill "c:\temporal.mdb"
  DBEngine.CreateDatabase "c:\temporal.mdb", dbLangGeneral
On Error GoTo 0

If llegir_ini("General", "parar", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini") = "si" Then MsgBox "Ara no es pot entrar al programa s'està actualitzant, espera 5 MINUTS, Gràcies", vbCritical, "Actualització": End
  assignardecimalipunt
  cami = llegir_ini("General", "cami", fitxerini)
  ruta_relativa_docs = "\\ser2\documentos\Pautacli"
  '"c:\misdoc~1\commandes\comandes.mdb"
  If existeix("c:\ordprog.ini") Then cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
  hora = Now
  centerscreen Me
  camiclixes = rutadelfitxer(cami) + "clixes.mdb"
  clixes.DatabaseName = camiclixes
  modifis.DatabaseName = camiclixes
  Set dbtmp = DBEngine.OpenDatabase(camiclixes)
  Set dbtmpb = DBEngine.OpenDatabase(cami)
  If cadbl(arguments(2)) > 0 Then
     clixes.RecordSource = "select * from clixes where id_treball=" + atrim(cadbl(arguments(2)))
     Frame3.Enabled = False
     m_manteniments.Enabled = False
     mllistats.Enabled = False
      Else: clixes.RecordSource = "select * from clixes "
  End If
  clixes.Refresh
  If Not clixes.Recordset.EOF Then clixes.Recordset.MoveLast
  activarframes False
  ensenyar_frame "framemodi"
  'DESACTIVAR EL FONS DE TOTS ELS LABELS
 For i = 0 To frmclixes.Controls.Count - 1
      If TypeOf frmclixes.Controls(i) Is Label Then frmclixes.Controls(i).BackStyle = 0
   Next i
  buscar_clixes_perduts_amb_liniescreades
End Sub
Sub buscar_clixes_perduts_amb_liniescreades()
  Set rsttmp = dbtmp.OpenRecordset("SELECT DISTINCTROW Clixes_albarans.id_treball, Clixes_albarans.data FROM Clixes_albarans LEFT JOIN Clixes ON Clixes_albarans.id_treball = Clixes.id_treball WHERE (((Clixes_albarans.id_treball)>3000) AND ((Clixes.id_treball) Is Null));")
  Set rstconsulta = dbtmp.OpenRecordset("SELECT DISTINCTROW Clixes_modifi.id_treball FROM Clixes_modifi LEFT JOIN Clixes ON Clixes_modifi.id_treball = Clixes.id_treball WHERE (((Clixes_modifi.id_treball)>3000) AND ((Clixes.id_treball) Is Null));")

  If Not rsttmp.EOF Or Not rstconsulta.EOF Then
     If MsgBox("Hi ha clixes que s'han perdut i que tenen albarans o modificiacions assignades." + Chr(10) + Chr(13) + "Vols crear els clixes sense dades corresponents?", vbYesNo, "Atenció") = vbYes Then
        treballs = ""
        While Not rsttmp.EOF
           clixes.Recordset.AddNew
             clixes.Recordset!id_treball = cadbl(rsttmp!id_treball)
             treballs = treballs + "[" + atrim(cadbl(rsttmp!id_treball)) + "]"
           clixes.Recordset.Update
           rsttmp.MoveNext
        Wend
        While Not rstconsulta.EOF
          Set rsttmp = dbtmp.OpenRecordset("select * from clixes where id_treball=" + atrim(rstconsulta!id_treball))
          If rsttmp.EOF Then
           clixes.Recordset.AddNew
             clixes.Recordset!id_treball = cadbl(rstconsulta!id_treball)
             treballs = treballs + "[" + atrim(cadbl(rstconsulta!id_treball)) + "]"
           clixes.Recordset.Update
          End If
          rstconsulta.MoveNext
        Wend
        MsgBox "Aquests son els clixes creats: " + treballs
        Set rstconsulta = Nothing
        Set rsttmp = Nothing
     End If
  End If

End Sub
Sub carregar_reixaalbarans()
  Dim rec As Recordset
  Set rec = albarans.Recordset.Clone
    With freixaalbarans
        .Rows = 2
         .Cols = rec.Fields.Count
    ' en el encabezado del Grid ponemos los nombres de los campos y ajustamos el ancho
    ' de la columna
        For i = 0 To .Cols - 1
         .TextMatrix(0, i) = rec.Fields(i).Name
         .ColWidth(i) = rec.Fields(i).Size * 120
        Next i
        .Rows = rec.RecordCount + 1
        .Row = 1
        .col = 0
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .Clip = rec.RecordCount
        .Visible = True
        .Row = 1
     End With

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > dadesclixes.Left + Command7.Left And X < dadesclixes.Left + Command7.Left + Command7.Width Then
     If Y > dadesclixes.Top + Command7.Top And Y < dadesclixes.Top + Command7.Top + Command7.Height Then
         Command7_Click
     End If
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > dadesclixes.Left + Command7.Left And X < dadesclixes.Left + Command7.Left + Command7.Width Then
     If Y > dadesclixes.Top + Command7.Top And Y < dadesclixes.Top + Command7.Top + Command7.Height Then
         frmclixes.MouseIcon = Command7.Picture
         frmclixes.MousePointer = 99
         Exit Sub
     End If
  End If
  frmclixes.MousePointer = 0
End Sub

Private Sub liniaproducte_DropDown()
Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   r = " and id_marca='" + marcaproducte + "'"
   formseleccio.Data1.RecordSource = "select distinct id_liniaproducte from clixes where id_client=" + atrim(cadbl(clixes.Recordset!id_client)) + r
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns("id_liniaproducte").Caption = "           Linia de producte"
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           marcaproducte = formseleccio.DBGrid2.Columns("id_liniaproducte")
        End If
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub llistatmodifiacionspendentsacabar_Click()
   imprimirllistat
End Sub
Sub imprimirllistat()
  
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application

  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistatclixesmodificacionspendents.rpt", 1)
  oreport.Database.Tables.Item(1).Location = clixes.DatabaseName
  oreport.Database.Tables.Item(2).Location = clixes.DatabaseName
  oreport.Database.Tables.Item(3).Location = clixes.DatabaseName
  oreport.Database.Tables.Item(4).Location = clixes.DatabaseName
  oreport.Database.Tables.Item(5).Location = clixes.DatabaseName
  oreport.DiscardSavedData
  '
  oreport.RecordSelectionFormula = " ISNULL({Clixes_modifi.data_fi}) "
  'oreport.FormulaFields.GetItemByName("titol").Text = "'Llistat de la planificacio de la maquina: " + atrim(nummaquina) + " - " + nommaquina(nummaquina) + "'"
  
  'If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
  '  Else
  '    oreport.PrintOut False, 1
 ' End If
  
  

End Sub


Private Sub m_fotogravadors_Click()
  fFotogravadors.Show 1
End Sub

Sub crear_taula_tmp_llistatpntfact()
  Dim taula_tmp As String
Dim camps(100, 2) As String
'creo la taula de linies d'albarans
  taula_tmp = "tmp_albarans_pendents"
  On Error Resume Next
   dbtmp.Execute "drop table " + taula_tmp
  On Error GoTo 0
  i = 1
  camps(i, 1) = "id_treball": camps(i, 2) = "integer": i = i + 1
  camps(i, 1) = "data": camps(i, 2) = "date": i = i + 1
  camps(i, 1) = "client": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "producte": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "total": camps(i, 2) = "double": i = i + 1
  
  dbtmp.Execute ("create table " + taula_tmp + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbtmp.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
  
End Sub

Private Sub m_pendentsdefacturar_Click()
   Dim rstlli As Recordset
   Dim rstmod As Recordset
   Dim rsttreball As Recordset
   Dim rstclient
   Dim nomclient As String
   crear_taula_tmp_llistatpntfact
   Set rstlli = dbtmp.OpenRecordset("tmp_albarans_pendents")
   r = "SELECT Clixes_albarans.id_treball, Sum(Clixes_albarans.import) AS total From Clixes_albarans"
   r = r + " GROUP BY Clixes_albarans.id_treball, Clixes_albarans.facturat"
   r = r + " HAVING (((Clixes_albarans.facturat)=False));"
   Set rstmod = dbtmp.OpenRecordset(r)
   While Not rstmod.EOF
     Set rsttreball = dbtmp.OpenRecordset("select * from clixes where id_treball=" + atrim(cadbl(rstmod!id_treball)))
     If Not rsttreball.EOF Then
         Set rstclient = dbtmpb.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(rsttreball!id_client)))
         If Not rstclient.EOF Then nomclient = atrim(rstclient!nom)
         rstlli.AddNew
           rstlli!id_treball = rsttreball!id_treball
           rstlli!Data = rsttreball!datainicitreball
           rstlli!client = nomclient
           rstlli!producte = rsttreball!id_liniaproducte
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

Private Sub marcaproducte_DropDown()
Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "select distinct id_marca from clixes where id_client=" + atrim(cadbl(clixes.Recordset!id_client))
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns("id_marca").Caption = "              Marca"
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           marcaproducte = formseleccio.DBGrid2.Columns("id_marca")
        End If
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub mestatdclixes_Click()
 Load formaltarep
  formaltarep.Caption = "Estat de clixés"
'  formaltarep.autonum = "transportistes"
  formaltarep.Data1.DatabaseName = clixes.DatabaseName
  formaltarep.Data1.RecordSource = "select * from clixes_estats"

  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  
  formaltarep.Width = formaltarep.Width + 700
  formaltarep.DBGrid1.Width = formaltarep.DBGrid1.Width + 700
  formaltarep.Show
End Sub

Private Sub modificar_Click()
If clixes.Recordset.EOF Then Exit Sub
  activarframes True
  clixes.Recordset.Edit
  Text2.SetFocus
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub nomclient_DropDown()
Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select * from clients"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           nomclient = formseleccio.DBGrid2.Columns("nom")
           clixes.Recordset!id_client = formseleccio.DBGrid2.Columns("codi")
        End If
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub nomproveidor_DropDown()
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "select codi,nomfotogravador from fotogravadors"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           nomproveidor = formseleccio.DBGrid2.Columns("nomfotogravador")
           clixes.Recordset!proveidor = formseleccio.DBGrid2.Columns("codi")
        End If
   End If
    If seleccioret = 9 Then
        nomproveidor = ""
        clixes.Recordset!proveidor = Null
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub nomrepresentant_DropDown()
Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select codi,nom from representants"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           nomrepresentant = formseleccio.DBGrid2.Columns("nom")
           clixes.Recordset!representant = formseleccio.DBGrid2.Columns("codi")
        End If
   End If
   If seleccioret = 9 Then
           nomrepresentant = ""
           clixes.Recordset!representant = Null
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio

End Sub

Private Sub picker_Change()
 frmclixes.Controls(picker.Tag) = picker.Value
 frmclixes.Controls(picker.Tag).SetFocus
End Sub

Private Sub reixaalbarans_ButtonClick(ByVal ColIndex As Integer)
  If reixaalbarans.Columns(ColIndex).Locked Then Exit Sub
   If Not validarliniaalb Then Exit Sub
   Load formseleccio
   formseleccio.Data1.DatabaseName = clixes.DatabaseName
   formseleccio.Data1.RecordSource = "select * from clixes_detallsalb"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns("id_detall").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           reixaalbarans.Columns("id_detall") = formseleccio.DBGrid2.Columns("id_detall")
        End If
   End If
  ' If seleccioret = 9 Then
  '         reixaalbarans.Columns("id_detall") = Null
  ' End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
'   albarans.Refresh
 'guardar_reg reixaalbarans, albarans
   'If albarans.Recordset.EditMode = 0 Then albarans.Recordset.Edit
   'albarans.Recordset.Update
   gravar_reixa reixaalbarans
End Sub
Function validarliniaalb() As Boolean
  Dim bk As Integer
  validarliniaalb = True
  If Not IsDate(reixaalbarans.Columns("data")) Then
      MsgBox "Data erronea canvia-la sisplau"
      reixaalbarans.col = reixaalbarans.Columns("data").ColIndex
      validarliniaalb = False
      Exit Function
  End If
  'On Error GoTo fi
  bk = reixaalbarans.Columns("ordre")
  If albarans.Recordset.EditMode = 0 Then albarans.Recordset.Edit
  albarans.Recordset.Update
  
  If bk > 0 Then albarans.Recordset.FindFirst "ordre=" + atrim(bk)
  Exit Function
fi:
  validarliniaalb = False
End Function
Sub guardar_reg(reixa As Control, dbdata As Control)
    Dim i As Byte
    Dim camp As String
    If reixa.Row = -1 Then Exit Sub
    i = 0
    If dbdata.Recordset.EditMode = 0 Then dbdata.Recordset.Edit
    i = 0
    While i < dbdata.Recordset.Fields.Count
     'reixabobines.col = i
     'camp = reixabobines.Columns(i + 1).DataField
     camp = dbdata.Recordset.Fields(i).Name
     If existeixelcamp(camp, reixa) Then dbdata.Recordset.Fields(camp) = reixa.Columns(camp)
     i = i + 1
    Wend
    dbdata.Recordset.Update
End Sub
Function existeixelcamp(camp As String, reixa As Control) As Boolean
  For i = 0 To reixa.Columns.Count - 1
     If reixa.Columns(i).DataField = camp Then existeixelcamp = True
  Next i
End Function


Sub comprovarsiestafacturat()
  Dim bk As Double
  If reixaalbarans.Columns("data").Locked Or albarans.Recordset.EOF Then Exit Sub
 If reixaalbarans.Columns(reixaalbarans.col).DataField = "facturat" Then
     If reixaalbarans.Columns("facturat") = "Sí" Then
         reixaalbarans.Columns("facturat") = False
        Else:
          If MsgBox("Vols passar tots els pendents de facturar a facturats?", vbInformation + vbYesNo, "Atenció") = vbYes Then
            dbtmp.Execute "update clixes_albarans set facturat=true where id_treball=" + atrim(cadbl(clixes.Recordset!id_treball))
            bk = albarans.Recordset!ordre
            albarans.Refresh
            reixaalbarans.Refresh
            albarans.Recordset.FindFirst "ordre=" + atrim(bk)
              Else: reixaalbarans.Columns("facturat") = True
          End If
     End If
     'If albarans.Recordset.EditMode = 0 Then albarans.Recordset.Edit
     'albarans.Recordset.Update
     gravar_reixa reixaalbarans
                 

 End If
End Sub

Private Sub reixaalbarans_DblClick()
   'If albarans.Recordset.EditMode = 0 Then albarans.Recordset.Edit
   'albarans.Recordset.Update
   If Not gravar_reixa(reixaalbarans) Then Exit Sub
   ensenyar_picker reixaalbarans, albarans
   comprovarsiestafacturat
   
End Sub

 Private Sub reixaalbarans_Error(ByVal DataError As Integer, Response As Integer)
   If 16389 = DataError Then
      Response = 0
      MsgBox "No hi ha una descripció sel.leccionada", vbCritical, "Atenció"
      reixaalbarans.SetFocus
   End If
End Sub

Private Sub reixaalbarans_GotFocus()
    reixaalbarans.Columns("facturat").Locked = True
    If clixes.Recordset.EditMode > 0 Then clixes.UpdateRecord: clixes.Recordset.Edit ': activarframes True
End Sub

Private Sub reixaalbarans_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 32 Then comprovarsiestafacturat
   
End Sub

Private Sub reixaalbarans_OnAddNew()
   Dim gran As Integer
   gran = albmaxordre(albarans)
   albarans.Recordset!id_treball = clixes.Recordset!id_treball
   albarans.Recordset!ordre = gran + 1
   reixaalbarans.Columns("ordre") = gran + 1
   'reixaalbarans.Columns("data") = Format(Now, "dd/mm/yy")
End Sub
Function albmaxordre(dbctrl As Control) As Integer
   Dim rs As Recordset
    'If albarans.Recordset.EOF Then albmaxordre = 0
   Set rs = dbctrl.Recordset.Clone
   If Not rs.EOF Then
     rs.MoveLast
     albmaxordre = cadbl(rs!ordre)
    Else: albmaxordre = 0
   End If
End Function

Sub ensenyar_picker(reixa As DBGrid, taula As Data)
      If reixa.Columns(reixa.col).Locked = True Then Exit Sub
  r = ""
  If taula.Recordset.Fields(reixa.Columns(reixa.col).DataField).Type = 8 Then
       picker.Visible = True
       If IsDate(reixa) Then
          picker.Value = reixa
         Else: picker.Value = Now
       End If
       picker.Move reixa.Container.Left + reixa.Columns(reixa.col).Left, reixa.Container.Top + reixa.Columns(reixa.col).Top
       picker.SetFocus
       SendKeys ("%{down}")
       picker.Tag = reixa.Name
         Else: picker.Visible = False
  End If
   
End Sub

Private Sub reixaalbarans_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 If Not gravant Then
    If (reixaalbarans.col = 0 Or reixaalbarans.col = 1) And reixaalbarans.Text = "" And clixes.Recordset.EditMode > 0 Then reixaalbarans.Text = copiarvaloranterior(reixaalbarans.Columns(reixaalbarans.col).DataField)
 End If
End Sub
Function copiarvaloranterior(camp As String) As String
  Dim rstcp As Recordset
  Set rstcp = albarans.Recordset.Clone
  If rstcp.EOF Then Exit Function
  rstcp.MoveLast ': rstcp.MovePrevious
  If rstcp.EOF Then Exit Function
  If atrim(rstcp!ordre) = reixaalbarans.Columns("ordre") Then rstcp.MovePrevious
  copiarvaloranterior = atrim(rstcp.Fields(camp))
End Function

Private Sub reixamodifis_ButtonClick(ByVal ColIndex As Integer)
   Dim ordre As Long
   Dim nomdata As String
   nomdata = reixamodifis.Columns(ColIndex).DataField
   If IsDate(reixamodifis.Columns(ColIndex)) Then
     If UCase(InputBox("Segur que vols eliminar aquesta data?" + Chr(10) + "Escriu ELIMINAR per eliminar-la.", "Eliminar data")) = "ELIMINAR" Then
         ordre = modifis.Recordset!ordre
         modifis.Database.Execute "update clixes_modifi set " + nomdata + "=null where id_treball=" + atrim(modifis.Recordset!id_treball) + " and ordre=" + atrim(modifis.Recordset!ordre)
         modifis.Refresh
         modifis.Recordset.FindFirst "ordre=" + atrim(ordre)
     End If
   End If
   If nomdata = "descripcioestat" Then
       demanaridestat
   End If
End Sub

Sub demanaridestat()
   Load formseleccio
   formseleccio.Data1.DatabaseName = clixes.DatabaseName
   formseleccio.Data1.RecordSource = "select * from clixes_estats"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           reixamodifis.Columns("descripcioestat") = formseleccio.DBGrid2.Columns("descripcio")
           reixamodifis.Columns("id_estatclixe") = formseleccio.DBGrid2.Columns("id_estat")
        End If
   End If
   If seleccioret = 9 Then
            reixamodifis.Columns("descripcioestat") = ""
           reixamodifis.Columns("id_estatclixe") = Null
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub reixamodifis_ColEdit(ByVal ColIndex As Integer)
   If reixamodifis.Columns("Estat del Clixé") = "" Then
      If Not possardatafialanteriormodificacio Then Exit Sub
      reixamodifis.col = 1
      demanaridestat
      reixamodifis.col = 2
   End If
End Sub

Private Sub reixamodifis_DblClick()
 'If modifis.Recordset.EditMode = 0 Then modifis.Recordset.Edit
 'modifis.Recordset.Update
 If Not gravar_reixa(reixamodifis) Then Exit Sub
 ensenyar_picker reixamodifis, modifis
 comprovarsiestaseleccionat
 
End Sub

Private Sub reixamodifis_GotFocus()
 reixamodifis.Columns("imprimiracomanda").Locked = True
 If clixes.Recordset.EditMode > 0 Then clixes.UpdateRecord: clixes.Recordset.Edit ': activarframes True
End Sub

Sub comprovarsiestaseleccionat()
  If reixamodifis.Columns("descripcio").Locked Then Exit Sub
 If reixamodifis.Columns(reixamodifis.col).DataField = "imprimiracomanda" Then
     If reixamodifis.Columns("imprimiracomanda") = "Sí" Then
         reixamodifis.Columns("imprimiracomanda") = "No"
        Else:
          If MsgBox("Vols passar tots els pendents de marcar a fets?", vbInformation + vbYesNo, "Atenció") = vbYes Then
            dbtmp.Execute "update clixes_modifi set imprimiracomanda='Sí' where id_treball=" + atrim(cadbl(clixes.Recordset!id_treball))
            modifis.Refresh
            modifis.Recordset.MoveLast
            reixamodifis.Refresh
              Else:
                 reixamodifis.Columns("imprimiracomanda") = "Sí"
                 'If modifis.Recordset.EditMode = 0 Then modifis.Recordset.Edit
                 'modifis.Recordset.Update
                 gravar_reixa reixamodifis

          End If
            
     End If
    
 End If

End Sub
Private Sub reixamodifis_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 32 Then comprovarsiestaseleccionat

End Sub

Private Sub reixamodifis_OnAddNew()
    Dim gran As Integer
   If Not possardatafialanteriormodificacio Then Exit Sub
   gran = albmaxordre(modifis)
   'modifis.Recordset.AddNew
   modifis.Recordset!id_treball = clixes.Recordset!id_treball
   modifis.Recordset!ordre = gran + 1
   'reixamodifis.Columns("data_inici") = Format(Now, "dd/mm/yy")
End Sub
Function possardatafialanteriormodificacio() As Boolean
  Dim rst As Recordset
  possardatafialanteriormodificacio = True
  Set rst = modifis.Recordset.Clone
  If Not rst.EOF Then
    rst.MoveLast
    If Not IsDate(rst!data_fi) Then
       possardatafialanteriormodificacio = False
       If modifis.Recordset.EditMode > 0 Then modifis.Recordset.CancelUpdate
       modifis.Refresh
       modifis.Recordset.FindFirst "ordre=" + atrim(rst!ordre)
       Data = InputBox("Entra la data de finalització de la modificació" + Chr(10) + atrim(rst!descripcio), "Data Fi", Format(Now, "dd/mm/yy"))
       If Not IsDate(Data) Then
          Exit Function
         Else
           modifis.Recordset.Edit
           modifis.Recordset!data_fi = Format(Data, "dd/mm/yy")
           modifis.Recordset.Update
           'modifis.Recordset.Move modifis.Recordset.RecordCount
           reixamodifis.SetFocus
       End If
    End If
  End If
  Set rst = Nothing
End Function
Private Sub sistemaimpresio_Change()
  possar_color_tipusimpresio sistemaimpresio
End Sub

Private Sub sistemaimpresio_Click()
possar_color_tipusimpresio sistemaimpresio
End Sub

Private Sub sortir_Click()
  If MsgBox("Segur que vols sortir?", vbCritical + vbYesNo, "Atenció") = vbYes Then End
End Sub

Private Sub Text1_DblClick()
   Text1.Locked = False
End Sub

Private Sub Text10_LostFocus()
If atrim(Text10) > "" Then
  If Not IsDate(Text10) Then MsgBox "Data errònea", vbCritical, "Atenció"
 End If
End Sub

Private Sub Text7_LostFocus()
 If atrim(Text7) > "" Then
  If Not IsDate(Text7) Then MsgBox "Data errònea", vbCritical, "Atenció"
 End If
End Sub

Private Sub Text8_Change()
  combo8 = ""
  If Text8 = "N" Then Combo1 = "Normal"
  If Text8 = "T" Then Combo1 = "Transparencia"

End Sub

Private Sub Text9_LostFocus()
If atrim(Text9) > "" Then
  If Not IsDate(Text9) Then MsgBox "Data errònea", vbCritical, "Atenció"
 End If
End Sub

Private Sub Timer1_Timer()
 Select Case clixes.Recordset.EditMode
    Case 0
       estatedicio = ""
    Case 1
       estatedicio = "Editant..."
       framealb.Enabled = True
       framemodi.Enabled = True
    Case 2
       estatedicio = "Agegint..."
       framealb.Enabled = False
       framemodi.Enabled = False
 End Select
 If buscant Then estatedicio = "Buscant..."
    canviarelscolorsdelscontrolsalentrar
 mirarsiparar
End Sub

Sub mirarsiparar()
 Static contar
 Dim paraula As String
 paraula = llegir_ini("General", "parar", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini")
  If paraula = "si" Or InStr(1, paraula, "[clixes]") > 0 Then
    contar = contar + 1
     If contar = 1 Then MsgBox2 "El programa es pararà d'aqui a 1 minut. TANCA TOT I ESPERA CINC MINUTS.", 5, "Actualització", vbCritical
     If contar = 15 Then MsgBox2 "El programa es pararà d'aqui a 30 segons. TANCA TOT I ESPERA CINC MINUTS.", 5, "Actualització, vbCritical"
     If contar = 27 Then MsgBox2 "El programa es pararà d'aqui a 5 segons. TANCA TOT I ESPERA CINC MINUTS.", 3, "Actualització", vbCritical
     If contar > 30 Then End
   Else: contar = 0
  End If
  If paraula = "ja" Then End
End Sub
