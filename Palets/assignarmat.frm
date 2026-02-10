VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form assignarmat 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assignar material a la comanda"
   ClientHeight    =   9090
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   11850
   Icon            =   "assignarmat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox etinformacio 
      Alignment       =   2  'Center
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
      Height          =   525
      Left            =   -7905
      MultiLine       =   -1  'True
      TabIndex        =   71
      Top             =   1245
      Visible         =   0   'False
      Width           =   8685
   End
   Begin VB.CommandButton modificar 
      Height          =   315
      Left            =   8220
      Picture         =   "assignarmat.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   78
      TabStop         =   0   'False
      ToolTipText     =   "Box de subfamilies de materials compatibles"
      Top             =   180
      Width           =   390
   End
   Begin VB.CheckBox ordenatperpalet 
      Height          =   270
      Left            =   2355
      TabIndex        =   66
      Top             =   2865
      Width           =   165
   End
   Begin VB.CommandButton botoajust 
      BackColor       =   &H0080FF80&
      Caption         =   "Ajust"
      Height          =   375
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Aquesta comanda s'assignarà material d'Stock."
      Top             =   2505
      Width           =   660
   End
   Begin VB.CommandButton assignarstock 
      BackColor       =   &H0080FF80&
      Caption         =   "Stock"
      Height          =   375
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Aquesta comanda s'assignarà material d'Stock."
      Top             =   2505
      Width           =   585
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   405
      Left            =   10755
      TabIndex        =   59
      Top             =   1815
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CheckBox partirmanual 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Partir bob. manual"
      Height          =   195
      Left            =   180
      TabIndex        =   58
      ToolTipText     =   "La bobina que sel.leccionis ultima serà la que es partirà."
      Top             =   2865
      Width           =   2115
   End
   Begin MSFlexGridLib.MSFlexGrid reixacomandes 
      Height          =   5205
      Left            =   8190
      TabIndex        =   54
      Top             =   3015
      Visible         =   0   'False
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   9181
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin VB.CommandButton desreservar 
      BackColor       =   &H000000FF&
      Caption         =   "Des-Reservar"
      Height          =   390
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2490
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton botocomprar 
      BackColor       =   &H009AA6FA&
      Height          =   390
      Left            =   7815
      Picture         =   "assignarmat.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Comprar"
      Top             =   2490
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Data datalat 
      Caption         =   "datalat"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   9405
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "percomandaoclient"
      Top             =   2595
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSDBGrid.DBGrid reixalat 
      Bindings        =   "assignarmat.frx":109E
      Height          =   5895
      Left            =   8250
      OleObjectBlob   =   "assignarmat.frx":10B0
      TabIndex        =   44
      Top             =   3000
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton Reserves 
      BackColor       =   &H000080FF&
      Caption         =   "Reserves"
      Height          =   420
      Left            =   10470
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   525
      Width           =   1230
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   7485
      Top             =   2475
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   6975
      Picture         =   "assignarmat.frx":1DEC
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Imprimir Packing-List"
      Top             =   2490
      Width           =   810
   End
   Begin VB.Frame missatgepantalla 
      BackColor       =   &H00C0FFC0&
      Height          =   960
      Left            =   4515
      TabIndex        =   19
      Top             =   4155
      Visible         =   0   'False
      Width           =   3360
      Begin VB.Label etprogres 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   495
         TabIndex        =   65
         Top             =   615
         Width           =   2430
      End
      Begin VB.Shape liniaprogres 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   195
         Top             =   600
         Width           =   105
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   315
         Left            =   165
         Top             =   570
         Width           =   3060
      End
      Begin VB.Label etmissatge 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carregant materials"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   30
         TabIndex        =   20
         Top             =   255
         Width           =   3285
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Atenció no tocar"
      Height          =   510
      Left            =   10530
      TabIndex        =   14
      Top             =   1125
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox mtrsnecessaris 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3015
      TabIndex        =   13
      Text            =   "0"
      Top             =   2505
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Ok... Assignar"
      Height          =   360
      Left            =   4185
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Palets"
      Height          =   420
      Left            =   10485
      TabIndex        =   10
      Top             =   60
      Width           =   1230
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   5985
      Left            =   75
      TabIndex        =   0
      Top             =   3120
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   10557
      _Version        =   393216
      GridColor       =   8421504
      GridLinesFixed  =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2550
      Left            =   135
      TabIndex        =   1
      Top             =   -45
      Width           =   10275
      Begin VB.Frame Framecompatibles 
         BackColor       =   &H00FDDECE&
         Height          =   1350
         Left            =   1680
         TabIndex        =   75
         Top             =   1590
         Visible         =   0   'False
         Width           =   5490
         Begin VB.ComboBox Combocompatibles 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00ED823A&
            Height          =   360
            Left            =   255
            TabIndex        =   76
            Top             =   465
            Width           =   4995
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Grup de materials compatibles"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   77
            Top             =   180
            Width           =   3270
         End
      End
      Begin VB.CommandButton parafiltre 
         BackColor       =   &H008080FF&
         Caption         =   "Parar Filtre"
         Height          =   570
         Left            =   8790
         Picture         =   "assignarmat.frx":2376
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   885
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CheckBox matprovproves 
         Caption         =   "Material proveïdor proves."
         Height          =   195
         Left            =   3195
         TabIndex        =   70
         Top             =   135
         Width           =   2235
      End
      Begin VB.ListBox llistaaltrescomandes 
         Height          =   255
         Left            =   3000
         Sorted          =   -1  'True
         TabIndex        =   57
         Top             =   195
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox altrescomandes 
         Height          =   285
         Left            =   1995
         TabIndex        =   56
         ToolTipText     =   "Altres comandes"
         Top             =   195
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Frame frameclient 
         Height          =   765
         Left            =   45
         TabIndex        =   46
         Top             =   630
         Visible         =   0   'False
         Width           =   3030
         Begin VB.ComboBox nomclient 
            Height          =   315
            Left            =   60
            TabIndex        =   48
            Top             =   390
            Width           =   2910
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Stock"
            Height          =   270
            Left            =   2205
            TabIndex        =   47
            Top             =   120
            Width           =   750
         End
         Begin VB.Label codiclient 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   675
            TabIndex        =   50
            Top             =   195
            Width           =   750
         End
         Begin VB.Label Label5 
            Caption         =   "Client:"
            Height          =   225
            Left            =   135
            TabIndex        =   49
            Top             =   180
            Width           =   510
         End
      End
      Begin VB.ComboBox famad 
         Height          =   315
         Left            =   3405
         TabIndex        =   28
         Top             =   1170
         Width           =   2580
      End
      Begin VB.ComboBox subfamad 
         Height          =   315
         Left            =   6015
         TabIndex        =   27
         Tag             =   "famad"
         Top             =   1155
         Width           =   2490
      End
      Begin VB.ComboBox famcol 
         Height          =   315
         Left            =   3405
         TabIndex        =   26
         Top             =   840
         Width           =   2580
      End
      Begin VB.ComboBox subfamcol 
         Height          =   315
         Left            =   6015
         TabIndex        =   25
         Tag             =   "famcol"
         Top             =   840
         Width           =   2490
      End
      Begin VB.CommandButton filtrar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Filtrar"
         Height          =   510
         Left            =   8805
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1275
      End
      Begin VB.ComboBox subfammat 
         Height          =   315
         Left            =   6015
         Sorted          =   -1  'True
         TabIndex        =   7
         Tag             =   "fammat"
         Top             =   510
         Width           =   2490
      End
      Begin VB.ComboBox fammat 
         Height          =   315
         Left            =   3405
         TabIndex        =   5
         Top             =   510
         Width           =   2580
      End
      Begin VB.TextBox comanda 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   165
         Width           =   705
      End
      Begin VB.Frame Frame2 
         Caption         =   "Informació de la comanda"
         Height          =   1125
         Left            =   45
         TabIndex        =   15
         Top             =   1395
         Width           =   10170
         Begin VB.CommandButton Command7 
            Caption         =   "Canviar Material de la comanda"
            Height          =   255
            Left            =   7500
            TabIndex        =   60
            Top             =   135
            Width           =   2625
         End
         Begin VB.TextBox isolapa 
            DataField       =   "solapa"
            DataSource      =   "palets"
            Height          =   285
            Left            =   3525
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   480
            Width           =   555
         End
         Begin VB.TextBox iample 
            DataField       =   "Ample"
            DataSource      =   "palets"
            Height          =   285
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   480
            Width           =   795
         End
         Begin VB.TextBox iplegat 
            DataField       =   "Plegat"
            DataSource      =   "palets"
            Height          =   285
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   480
            Width           =   555
         End
         Begin VB.TextBox iespesor 
            DataField       =   "micres"
            DataSource      =   "palets"
            Height          =   285
            Left            =   4125
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   495
            Width           =   660
         End
         Begin VB.ComboBox itl 
            DataField       =   "semielaborat"
            DataSource      =   "palets"
            Height          =   315
            ItemData        =   "assignarmat.frx":2900
            Left            =   45
            List            =   "assignarmat.frx":290A
            TabIndex        =   33
            Top             =   465
            Width           =   615
         End
         Begin VB.ComboBox icares 
            DataField       =   "carestractat"
            DataSource      =   "palets"
            Height          =   315
            ItemData        =   "assignarmat.frx":2914
            Left            =   675
            List            =   "assignarmat.frx":2921
            TabIndex        =   32
            Top             =   465
            Width           =   615
         End
         Begin VB.ComboBox iobert 
            DataField       =   "obert"
            DataSource      =   "palets"
            Height          =   315
            ItemData        =   "assignarmat.frx":292E
            Left            =   1485
            List            =   "assignarmat.frx":293B
            TabIndex        =   31
            Top             =   465
            Width           =   615
         End
         Begin VB.CheckBox imicrop 
            Caption         =   "Microperforat"
            DataField       =   "microperforat"
            DataSource      =   "palets"
            Height          =   285
            Left            =   75
            TabIndex        =   30
            Top             =   780
            Width           =   1365
         End
         Begin VB.CheckBox limitarample 
            Caption         =   "."
            Height          =   195
            Left            =   2100
            TabIndex        =   23
            ToolTipText     =   "Filtre d'amplades."
            Top             =   210
            Value           =   1  'Checked
            Width           =   210
         End
         Begin VB.Label infoajust 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   4080
            TabIndex        =   63
            Top             =   855
            Width           =   5985
         End
         Begin VB.Label lblLabels 
            Caption         =   "Solapa:"
            Height          =   255
            Index           =   0
            Left            =   3555
            TabIndex        =   53
            Top             =   210
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "Ample:"
            Height          =   255
            Index           =   2
            Left            =   2325
            TabIndex        =   43
            Top             =   180
            Width           =   630
         End
         Begin VB.Label lblLabels 
            Caption         =   "Plegat:"
            Height          =   255
            Index           =   3
            Left            =   2970
            TabIndex        =   42
            Top             =   210
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "Espesor:"
            Height          =   255
            Index           =   15
            Left            =   4170
            TabIndex        =   41
            Top             =   180
            Width           =   600
         End
         Begin VB.Label etmicres 
            Caption         =   "Micres"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   4860
            TabIndex        =   40
            Top             =   195
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "T/L"
            Height          =   300
            Index           =   3
            Left            =   165
            TabIndex        =   36
            Top             =   195
            Width           =   360
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cares Tractat"
            Height          =   300
            Index           =   2
            Left            =   525
            TabIndex        =   35
            Top             =   210
            Width           =   1020
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Obert"
            Height          =   300
            Index           =   1
            Left            =   1575
            TabIndex        =   34
            Top             =   225
            Width           =   540
         End
         Begin VB.Label infodescripciomat 
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
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   5055
            TabIndex        =   18
            Top             =   435
            Width           =   5070
         End
         Begin VB.Label infometres 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   1380
            TabIndex        =   17
            Top             =   780
            Width           =   4605
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripció del Material"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5700
            TabIndex        =   16
            Top             =   165
            Width           =   1800
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Prova nou Mètode filtratge"
         Height          =   600
         Left            =   8595
         TabIndex        =   72
         Top             =   585
         Value           =   1  'Checked
         Width           =   1620
      End
      Begin VB.TextBox cmargeamplada 
         Height          =   285
         Left            =   9765
         TabIndex        =   73
         Text            =   "15"
         Top             =   1140
         Width           =   330
      End
      Begin VB.Label etsubfamcompatible 
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
         Height          =   195
         Left            =   5985
         TabIndex        =   79
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Marge amplada"
         Height          =   360
         Left            =   8580
         TabIndex        =   74
         Top             =   1155
         Width           =   1350
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Mat:         Col:          Ad:"
         Height          =   1020
         Left            =   3090
         TabIndex        =   29
         Top             =   450
         Width           =   345
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         Height          =   225
         Left            =   1860
         TabIndex        =   55
         Top             =   225
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Familia "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4380
         TabIndex        =   4
         Top             =   300
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Comanda:         o Grups de Palets"
         Height          =   495
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   255
         Width           =   1365
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Subfamilia "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6075
         TabIndex        =   6
         Top             =   285
         Width           =   2520
      End
   End
   Begin VB.Label comandesperreservar 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   3915
      TabIndex        =   69
      Top             =   2880
      Width           =   7605
   End
   Begin VB.Label sumatori 
      BackStyle       =   0  'Transparent
      Caption         =   "õ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   -15
      TabIndex        =   68
      Top             =   540
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Ordenat per Palet"
      Height          =   225
      Left            =   2535
      TabIndex        =   67
      Top             =   2895
      Width           =   1425
   End
   Begin VB.Image check2 
      Height          =   165
      Left            =   10515
      Picture         =   "assignarmat.frx":2948
      Top             =   1785
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image check 
      Height          =   180
      Left            =   10410
      Picture         =   "assignarmat.frx":2B16
      Stretch         =   -1  'True
      Top             =   945
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label etdataimpresio 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF8080&
      Height          =   315
      Left            =   6600
      TabIndex        =   22
      Top             =   2850
      Width           =   5130
   End
   Begin VB.Label metressel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   165
      TabIndex        =   9
      ToolTipText     =   "Metres sel.leccionats"
      Top             =   2505
      Width           =   1215
   End
   Begin VB.Image nocheck 
      Height          =   180
      Left            =   9915
      Picture         =   "assignarmat.frx":2CE4
      Stretch         =   -1  'True
      Top             =   135
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Mtrs Sel.leccionats de"
      Height          =   330
      Left            =   1410
      TabIndex        =   12
      Top             =   2595
      Width           =   1680
   End
End
Attribute VB_Name = "assignarmat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vnofirmar As Boolean

Sub baixaseccioextrusora(comanda As String, Optional borrar As Boolean)

Dim dbbaixes As Database
Dim kilos As Double
Dim vkilostotals As Double
Dim vkilostotalsfinals As Double
Dim vnoutotal As Double
 Dim vmetres As Double
Dim vproducte As String

 Set dbstocks = dbtmp
 'Set dbstocks = OpenDatabase(camistocks, , True)
 'Set dbbaixes = OpenDatabase(camibaixes)
 Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", "comandes.ini"))

 'busco la informació de la comanda que necessito
 Set rsttmp = dbtmpb.OpenRecordset("select producte,proximaseccio from comandes where comanda=" + atrim(comanda))
   If Not rsttmp.EOF Then
     estat = atrim(rsttmp!proximaseccio)
     If estat = "" Then
         estat = "E"
     End If
      Else: GoTo noexisteixlacomanda
   End If
   vproducte = atrim(rsttmp!producte)
   Set rsttmp = dbtmpb.OpenRecordset("select ruta from productes where codi='" + rsttmp!producte + "'")
   estat = estat + " "
   ruta = rsttmp!ruta
   

 
 'borro la seccio d'extrussora de la comanda en concret
 Set rsttmp = dbbaixes.OpenRecordset("select id from extrussores where comanda=" + atrim(cadbl(comanda)))
 While Not rsttmp.EOF
     dbbaixes.Execute "delete * from bobinesext where controlid=" + atrim(rsttmp!id)
     rsttmp.MoveNext
 Wend
 dbbaixes.Execute "delete * from extrussores where comanda=" + atrim(cadbl(comanda))
' dbcomanda.Execute "update comandes set proximaseccio='E' where comanda=" + aTrim(cadbl(comanda))
 
 'fins aqui borra seccio baixa
 If borrar Then Exit Sub
 'faig l'alta de la seccio d'extrussores a baixes
 vdata = format(Now, "mm/dd/yy")
 vhora = format(Now, "hh:nn")
 Set rststocks = dbstocks.OpenRecordset("select idpalet,idbobina,metres from parcials where comanda='" + comanda + "'")
 dbbaixes.Execute "insert into extrussores (comanda,tipus,datainici,horainici,datafi,horafi) values (" + comanda + ",'F',#" + vdata + "#,#" + vhora + "#,#" + vdata + "#,#" + vhora + "#)"
     
  
  Set rsttmp = dbbaixes.OpenRecordset("select id from extrussores where comanda=" + atrim(cadbl(comanda)))
  id = cadbl(rsttmp!id)
  Set rsttmp = dbstocks.OpenRecordset("select * from palets where idpalet=" + atrim(cadbl(rststocks!idpalet)))
  Set rstbob = dbstocks.OpenRecordset("select * from bobines where idpalet=" + atrim(cadbl(rststocks!idpalet)))
  cont = 1
  While Not rststocks.EOF
     rstbob.FindFirst "idbobina=" + atrim(rststocks!idbobina)
     If Not rstbob.NoMatch Then
        If rstbob!mts = rststocks!metres Then
           vkilostotals = vkilostotals + cadbl(rstbob!pesdelproveidor)
            Else: vkilostotals = vkilostotals + compramat.conversiokilos(rsttmp!codimatprognou, rsttmp!ample, rststocks!metres, cadbl(rsttmp!micres), rsttmp!semielaborat, rsttmp!solapa)
        End If
          Else: vkilostotals = vkilostotals + compramat.conversiokilos(rsttmp!codimatprognou, rsttmp!ample, rststocks!metres, cadbl(rsttmp!micres), rsttmp!semielaborat, rsttmp!solapa)
     End If
     vmetres = vmetres + cadbl(rststocks!metres)
     'vkilostotals = vkilostotals + compramat.conversiokilos(rsttmp!codimatprognou, rsttmp!ample, rststocks!metres, cadbl(rsttmp!micres), rsttmp!semielaborat, rsttmp!solapa)
     cont = cont + 1
     rststocks.MoveNext
  Wend
  If cont > 1 Then rststocks.MoveFirst
  If vkilostotals > 0 And atrim(ruta) = "E" And vproducte <> "PC" And vproducte <> "PC2" And vproducte <> "PCP" Then vnoutotal = cadbl(InputBox("Aquest material va directa al client." + Chr(10) + "Entra el pes que vols facturar al client.", "Pes per facturar al client", vkilostotals))
  cont = 1
  While Not rststocks.EOF
     kilos = compramat.conversiokilos(rsttmp!codimatprognou, rsttmp!ample, rststocks!metres, cadbl(rsttmp!micres), rsttmp!semielaborat, rsttmp!solapa)
     If vnoutotal > 0 Then
        'kilos = Redondejar((vnoutotal * kilos) / vkilostotals, 3)
        kilos = rststocks!metres * (vkilostotals / vmetres)
     End If
     vkilostotalsfinals = vkilostotalsfinals + kilos
     dbbaixes.Execute "insert into bobinesext (controlid,metres,kilos,numerodebobina) values (" + atrim(id) + "," + atrim(cadbl(rststocks!metres)) + "," + passardecomaapunt(cadbl(kilos)) + "," + Trim(cont) + ")"
     cont = cont + 1
     rststocks.MoveNext
  Wend
  If vnoutotal > 0 Then MsgBox "El pes de les bobines abans d'assignar material era de " + atrim(vkilostotals) + "Kg i ara es de " + atrim(vkilostotalsfinals) + "Kg. ", vbInformation, "Informació"
 'fins aqui l'alta de seccio
 'si hi ha alguna bobina passo l'estat de la comanda a la proxima seccio
 'If cont > 1 Then
   'passo l'estat de comanda a la proxima
   
  If siesmaterialexacte(cadbl(comanda), True) <> "" Then If MsgBox("Es material ESPECÍFIC, Vols ASSIGNAR-LO definitiviament?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then vnofirmar = True: GoTo noexisteixlacomanda
   If Trim(estat) = "E" Then
     seccio = Mid(ruta, 2, 1)
     If Trim(seccio) = "" Then
        seccio = "V"
     End If
     dbtmpb.Execute "update comandes set seccioactual='E' where comanda=" + atrim(comanda)
     dbtmpb.Execute "update comandes set proximaseccio='" + seccio + "' where comanda=" + atrim(comanda)
   End If
noexisteixlacomanda:
' End If

End Sub
Function passardecomaapunt(valo As String) As String
   While InStr(1, valo, ",")
      valo = Mid(valo, 1, InStr(1, valo, ",") - 1) + "." + Mid(valo, InStr(1, valo, ",") + 1, Len(valo))
   Wend
   passardecomaapunt = valo
End Function




Private Sub altrescomandes_Change()
   altrescomandes.Width = (Len(altrescomandes) * 100) + 100
   If altrescomandes.Width < 650 Then altrescomandes.Width = 650
   If altrescomandes.Width > 2850 Then altrescomandes.Width = 2850
End Sub

Private Sub altrescomandes_GotFocus()
  If cadbl(comanda) = 0 Then comanda.SetFocus
End Sub

Private Sub altrescomandes_LostFocus()
   comprovar_altrescomandes
End Sub
Sub comprovar_altrescomandes()
   Dim p As String
   Dim c As Double
   Dim i As Double
   Dim j As Double
   Dim quantitat As Double
   Dim eliminades As String
   llistaaltrescomandes.Clear
   p = altrescomandes
   While p <> ""
      c = cadbl(Mid(p, 1, InStr(1, p, ",")))
      If InStr(1, p, ",") = 0 Then c = cadbl(p): p = ""
      p = Mid(p, InStr(1, p, ",") + 1)
      If c > 0 Then llistaaltrescomandes.AddItem atrim(c)
      If InStr(1, p, ",") = 0 Then If cadbl(p) = 0 Then p = ""
   Wend
   If llistaaltrescomandes.ListCount > 0 Then
    For i = 0 To llistaaltrescomandes.ListCount - 1
       For j = i + 1 To llistaaltrescomandes.ListCount - 1
          If llistaaltrescomandes.List(i) = llistaaltrescomandes.List(j) Then llistaaltrescomandes.List(j) = ""
       Next j
       If cadbl(llistaaltrescomandes.List(i)) = cadbl(comanda) Then
           llistaaltrescomandes.List(i) = ""
          Else
           If Not sonigualsquelaprimera(llistaaltrescomandes.List(i), quantitat) Then eliminades = eliminades + "," + llistaaltrescomandes.List(i): llistaaltrescomandes.List(i) = ""
       End If
    Next i
   End If
   p = ""
   If eliminades <> "" Then MsgBox "Les comandes " + eliminades + " no comparteixen les carecteristiques amb la primera.", vbCritical
   i = 0
   While i < llistaaltrescomandes.ListCount
     If atrim(llistaaltrescomandes.List(i)) <> "" Then
         p = p + IIf(p <> "", ",", "") + llistaaltrescomandes.List(i)
         i = i + 1
        Else: llistaaltrescomandes.RemoveItem i
     End If

   Wend
   altrescomandes = p
   If cadbl(altrescomandes.Tag) > 0 Then mtrsnecessaris = quantitat + cadbl(altrescomandes.Tag)
End Sub
Function sonigualsquelaprimera(numc As String, quantitat As Double) As Boolean
    Dim rstcom As Recordset
    Dim rstmat As Recordset
    Dim rstmat2 As Recordset
    Dim rstcom2 As Recordset
    sonigualsquelaprimera = False
    If cadbl(numc) = 0 Then Exit Function
    Set rstcom = dbtmpb.OpenRecordset("select * from comandes where comanda=" + atrim(comanda))
    Set rstcom2 = dbtmpb.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(numc)))
    If rstcom.EOF Or rstcom2.EOF Then Exit Function
    Set rstmat = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstcom!materialex)))
    Set rstmat2 = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstcom2!materialex)))
    If rstmat.EOF Or rstmat2.EOF Then Exit Function
    If aatrim(rstcom!tubolam) <> aatrim(rstcom2!tubolam) Then Exit Function
    If aatrim(rstcom!oberturaex) <> aatrim(rstcom2!oberturaex) Then Exit Function
    If cabool(rstcom!micropex) <> cabool(rstcom2!micropex) Then Exit Function
    If atrim(rstcom!tractatex) <> atrim(rstcom2!tractatex) Then Exit Function
    If atrim(rstcom!ampleesq) <> atrim(rstcom2!ampleesq) Then Exit Function
    If atrim(rstcom!plegatesq) <> atrim(rstcom2!plegatesq) Then Exit Function
    If atrim(rstcom!solapa) <> atrim(rstcom2!solapa) Then Exit Function
    If cadbl(rstcom!espessor) <> cadbl(rstcom2!espessor) Then Exit Function
    
    'comprovo els materials
    If rstmat!familia <> rstmat2!familia Or rstmat!subfamilia <> rstmat2!subfamilia Then Exit Function
    If rstmat!familiacol <> rstmat2!familiacol Or rstmat!subfamiliacol <> rstmat2!subfamiliacol Then Exit Function
    If rstmat!familiaad <> rstmat2!familiaad Or rstmat!subfamiliaad <> rstmat2!subfamiliaad Then Exit Function
    
    If cadbl(rstcom!mesuracantex) = cadbl(rstcom2!mesuracantex) Then quantitat = quantitat + rstcom2!cantitatex
    altrescomandes.Tag = atrim(rstcom!cantitatex)
    'atrim (rstinfo!cantitatex) + IIf(rstinfo!mesuracantex = 1, " Mtrs", " Kg")
    'iespesor = micresmaterial(rstinfo!mesuraesp, rstinfo!espessor, rstinfo!tubolam)
End Function

Private Sub assignarstock_Click()
Dim mtrsassignats As Double
If missatgepantalla.Visible Then Exit Sub
  If cadbl(comanda) > 10000 Then
   If cadbl(mtrsnecessaris) < 1 Then MsgBox "Has de possar els metres assignats que vols per aquesta comanda d'estoc.", vbCritical + vbOKOnly, "Atenció": Exit Sub
    If Not sicaltreurecomandadereclamades(cadbl(comanda)) Then Exit Sub
    assignarstock.BackColor = &H9AA6FA
    If opcionsmaterialajust(True) Then
     baixaseccioextrusora comanda
     r = "nopregunta"
     des_reservar comanda
     mtrsassignats = cadbl(mtrsnecessaris)
     dbtmpb.Execute ("update comandes_extres set assignarstock=true,mtrsassignatsestock=" + atrim(mtrsassignats) + " where comanda=" + atrim(CDbl(comanda)))
     firmar_fulla
     dbtmp.Execute ("delete * from historic_packinglist where comanda='" + atrim(cadbl(comanda)) + "'")
    End If
    If esimpresores(cadbl(comanda)) Then
       botoajust_Click
     Else: carregar_info_comanda atrim(cadbl(comanda))
    End If
  End If
End Sub
Function opcionsmaterialajust(Optional triarstock As Boolean) As Boolean
    Unload opcionsmatajust
    Load opcionsmatajust
    If assignarmat.assignarstock.BackColor <> &H80FF80 Then opcio = "estoc"
    If assignarstock.Tag = "PC" Or assignarstock.Tag = "PC2" Then triarstock = True
    
    If triarstock Then
      opcionsmatajust.b1.Enabled = False
       opcionsmatajust.b3.Enabled = False
      ' opcionsmatajust.possarframes "b2"
       opcionsmatajust.Frame4.Top = 1
       opcionsmatajust.Height = opcionsmatajust.Frame4.Height + 200
       opcionsmatajust.Frame4.ZOrder 0
      Else
       If opcio = "estoc" Then
        'opcionsmatajust.b1.Enabled = False
        'opcionsmatajust.b3.Enabled = False
        'opcionsmatajust.possarframes "b2"
        opcionsmatajust.dadesmatestoc.Enabled = False
        'opcionsmatajust.Frame4.Top = 1
        'opcionsmatajust.Height = opcionsmatajust.Frame4.Height + 100
       End If
    End If
'    If assignarstock.Tag = "PC" Or assignarstock.Tag = "PC2" Then
'       opcionsmatajust.Frame1.Enabled = False
'       opcionsmatajust.Frame2.Enabled = False
'       opcionsmatajust.framematespecific.Visible = False
 '      opcionsmatajust.dadesmatestoc.Visible = True
 '   End If
    opcionsmatajust.numcomanda = comanda
    opcionsmatajust.Show 1
    If botoajust.Tag = "1" Then
       opcionsmaterialajust = True
      Else: opcionsmaterialajust = False
    End If
End Function

Private Sub botoajust_Click()
  If missatgepantalla.Visible Then Exit Sub
   If esimpresores(cadbl(comanda)) Then
      opcionsmaterialajust
      carregar_info_comanda atrim(cadbl(comanda))
   End If
   infoajust = textedajust(cadbl(comanda))
End Sub

Private Sub botocomprar_Click()
'If cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("ample"))) = 0 And comanda <> "" Then MsgBox "Escull primer una mida, o treu la comanda sel.leccionada": Exit Sub
  compramat.compres.DatabaseName = camistock
  compramat.quantitatcomprar = cadbl(mtrsnecessaris) - metresareservar
  compramat.quantitatcomprar.Tag = metresareservar
  compramat.amplecompra = reixa.TextMatrix(reixa.row, columnadelcamp("ample"))
  compramat.descmat.Tag = crear_criteri_familia
  If reixa.Cols > 2 Then compramat.numreserva = cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("idreserva")))
  If cadbl(compramat.numreserva) = 0 Then
      compramat.compres.Tag = "select * from compresmaterial where  idreserva<0"
      Else: compramat.compres.Tag = "select * from compresmaterial where  idreserva=" + atrim(cadbl(compramat.numreserva))
  End If
  compramat.compres.RecordSource = compramat.compres.Tag + " order by data Desc"
  compramat.compres.Refresh
  compramat.Show
End Sub

Private Sub comanda_Change()
  
  If Len(comanda) >= 6 Then
     If mtrsnecessaris.Tag <> "reserva" Then
        mtrsnecessaris = buscar_metres_necessaris
       Else: mtrsnecessaris.Tag = ""
     End If
     If Screen.ActiveControl.Name = "comanda" Then
       Command2.Caption = IIf(Reserves.Caption = "Assignar", "Ok... Reservar", "Ok... Assignar")
       comandesperreservar = ""
       dbtmp.Execute "delete * from pendentsdereservar where entrat"
     End If
     carregar_info_comanda atrim(cadbl(comanda)), vfiltrebobinesdesdeimpresores
     'If vfiltrebobinesdesdeimpresores Then r = "no pregunta"
     If subfammat <> "" Or fammat <> "" Or Combocompatibles <> "" Then
        If vfiltrebobinesdesdeimpresores Then Check1.Value = 0
        filtrar_Click
     End If
  End If
End Sub
Sub demanar_nou_material(numc As String, codimat As Double, codicol As Double, codiad As Double)
   Dim rstmat As Recordset
   Dim noucodimat As Double
   Dim descmat As String
   Set rstmat = dbtmpb.OpenRecordset("Select * from materials where codi=" + atrim(codimat))
   If Not rstmat.EOF Then descmat = atrim(rstmat!descripcio)
   Set rstmat = dbtmpb.OpenRecordset("Select * from colorants where codi=" + atrim(codicol))
   If Not rstmat.EOF Then descmat = descmat + " + " + atrim(rstmat!descripcio)
   Set rstmat = dbtmpb.OpenRecordset("Select * from aditius where codi=" + atrim(codiad))
   If Not rstmat.EOF Then descmat = descmat + " + " + atrim(rstmat!descripcio)
   
   If descmat <> "" Then
      Load formseleccio
      formseleccio.Caption = "Material: " + descmat
      formseleccio.Data1.DatabaseName = cami
      formseleccio.Data1.RecordSource = "SELECT materials.codi as [Codi], materials.descripcio as [Descripcio], materials.refproducte as [RefProducte], proveidors.nom as [Proveidor] FROM materials LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE (((materials.codi)>499))"
      formseleccio.refrescar
      formseleccio.DBGrid2.Columns(0).Width = 500
      formseleccio.DBGrid2.Columns(1).Width = 2500
      formseleccio.DBGrid2.Columns(2).Width = 1000
      formseleccio.DBGrid2.Columns(3).Width = 1500
      formseleccio.Command2.Tag = "1"
      formseleccio.Width = formseleccio.Width + ((formseleccio.Width / 100) * 30)
      formseleccio.Show 1
      noucodimat = 0
      If seleccioret = 1 Then
        noucodimat = atrim(formseleccio.Data1.Recordset!codi)
      End If
      Unload formseleccio
      comprovarsidoscares cadbl(numc), noucodimat
   End If
   If noucodimat <> codimat And noucodimat > 0 Then
     dbtmpb.Execute "update comandes set colorex=0,aditiuex=0,materialex=" + atrim(noucodimat) + " where comanda=" + numc
     dbtmpb.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(numc) + ",'" + nomordinador + "','PALETS_materialex','" + atrim(codimat) + "','" + atrim(noucodimat) + "')"
   End If
End Sub
Sub comprovarsidoscares(numc As Double, noucodimat As Double)
 Dim rstm As Recordset
 Dim rstc As Recordset
 Set rstc = dbtmpb.OpenRecordset("select tractatex from comandes where comanda=" + atrim(numc))
 If rstc.EOF Then Exit Sub
 Set rstm = dbtmpb.OpenRecordset("select material2cares from materials where codi=" + atrim(noucodimat))
 If Not rstm.EOF Then If rstm!material2cares And atrim(rstc!tractatex) <> "2" Then MsgBox "OJU... AQUEST MATERIAL ESTÀ TRACTAT A DOS CARES I LA COMANDA POSSA MATERIAL QUE NO... FES ELS CANVIS OPORTUNS...", vbCritical, "ATENCIOOOOO"
 Set rstm = Nothing
 Set rstc = Nothing

End Sub
Sub comprovarsihihareservaoassignacio(numc As String)
   Dim rstreserva As Recordset
   Dim rstparcial As Recordset
   Dim rstcompres As Recordset
   Dim reservat As Boolean
   Dim resp As String
   
   Set rstparcial = dbtmp.OpenRecordset("select * from parcials where comanda='" + atrim(cadbl(comanda)) + "'")
   Set rstreserva = dbtmp.OpenRecordset("select sum(metres) as mtrs from percomandaoclient where (idcompra=0 or idcompra=null) and numcomanda=" + numc)
   Set rstcompres = dbcompres.OpenRecordset("SELECT comandesxlinia.numcomanda, liniescompra.kgentregats as kilos FROM liniescompra RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra WHERE (((comandesxlinia.numcomanda)=" + numc + "));")
   If Not rstreserva.EOF Then reservat = True
   If reservat Then
       If Not rstreserva.EOF Then r = "Info": des_reservar comanda
       'If rstreserva.EOF And Reserves.Caption = "Reserves" Then MsgBox "Hi ha material reservat per aquesta comanda."
   End If
   If Not rstcompres.EOF Then
     If Reserves.Caption <> "Reserves" And rstcompres!kilos = 0 Then
      While resp <> "SI"
       resp = InputBox("AQUESTA COMANDA TE UNA COMPRA FETA, ESCRIU [SI] PER CONTINUAR.", "ATENCIÓ")
      Wend
     End If
   End If
   If Not rstparcial.EOF And Reserves.Caption = "Reserves" Then
     MsgBox "La comanda " + numc + " ja te assignades bobines ", vbOKOnly, "Assignació de material"
   End If
End Sub
Sub mirarsihihaajust(numc As Double)
  Dim rstopcions As Recordset
  Set rstopcions = dbtmp.OpenRecordset("select * from opcionsdajust where comanda=" + atrim(numc))
  If Not rstopcions.EOF Then
     If cadbl(rstopcions!sistemadajust) > 0 Then botoajust.BackColor = &H9AA6FA
  End If
End Sub
Sub carregar_info_comanda(numc As String, Optional noveuremissatges As Boolean)
   Dim rstinfo As Recordset
   Dim rstcli As Recordset
   Dim rstmat As Recordset
   Dim rstparcial As Recordset
   
   Set rstinfo = dbtmpb.OpenRecordset("select * from comandes where comanda=" + numc)
   If Not rstinfo.EOF Then
     If rstinfo!materialex < 500 Then
       MsgBox "Aquesta comanda te el material asignat inferior al codi 500. S'ha de canviar, ESCULL UN MATERIAL NOU"
       demanar_nou_material numc, cadbl(rstinfo!materialex), cadbl(rstinfo!colorex), cadbl(rstinfo!aditiuex)
       Set rstinfo = dbtmpb.OpenRecordset("select * from comandes where comanda=" + numc)
     End If
   End If
   assignarstock.BackColor = &H80FF80
   botoajust.BackColor = &H80FF80
   mirarsihihaajust cadbl(comanda)
   infoajust = textedajust(cadbl(comanda))
   If Not rstinfo.EOF Then
       If Not noveuremissatges Then comprovarsihihareservaoassignacio numc
       Set rstmat = dbtmpb.OpenRecordset("Select * from materials where codi=" + atrim(cadbl(rstinfo!materialex)))
       'MsgBox rstinfo!materialex
       codiclient = atrim(rstinfo!client)
       nomclient = ""
       Framecompatibles.Visible = False
       assignarstock.Tag = ""
       Set rstcli = dbtmpb.OpenRecordset("select nom from clients where codi=" + codiclient)
       If Not rstcli.EOF Then nomclient = rstcli!nom
       assignarstock.Tag = atrim(rstinfo!producte)
       botoajust.Visible = IIf(atrim(rstinfo!producte) = "PC" Or atrim(rstinfo!producte) = "PC2", False, True)
       itl = aatrim(rstinfo!tubolam)
       iobert = aatrim(rstinfo!oberturaex)
       imicrop = IIf(cabool(rstinfo!micropex), 1, 0)
       infotuboobertmicro = aatrim(rstinfo!tubolam) + " - " + aatrim(rstinfo!oberturaex) + " - " + aatrim(rstinfo!micropex)
       imicrop.Tag = " and semielaborat='" + aatrim(rstinfo!tubolam) + "' and obert='" + aatrim(rstinfo!oberturaex) + "' and " + IIf(cabool(rstinfo!micropex), "", "not") + " microperforat"
       icares = atrim(rstinfo!tractatex)
       If icares = "0" Then icares = "N"
       infometres = atrim(rstinfo!cantitatex) + IIf(rstinfo!mesuracantex = 1, " Mtrs", " Kg")
       iample = atrim(rstinfo!ampleesq)
       iplegat = atrim(rstinfo!plegatesq)
       isolapa = atrim(rstinfo!solapa)
       'infoampleplegat = atrim(rstinfo!ampleesq) + " / " + atrim(rstinfo!plegatesq)
       iplegat.Tag = atrim(rstinfo!ampleesq)
       iespesor = micresmaterial(cadbl(rstinfo!mesuraesp), rstinfo!espessor, atrim(rstinfo!tubolam))
       etmicres = r: r = ""
       If Not rstmat.EOF Then
          'infodescripciomat = atrim(rstmat!descripcio)
          infodescripciomat = atrim(rstmat!codi) + "-" + atrim(rstmat!descripcio) + Chr(10) + Chr(13) + descripciomaterial(rstmat)
          infodescripciomat.Tag = rstmat!codi
          If rstmat!codi < 500 Then infodescripciomat.ForeColor = QBColor(12)
          infodescripciomat.ForeColor = infometres.ForeColor
          possar_familia_subfamilia cadbl(rstmat!familia), cadbl(rstmat!subfamilia), cadbl(rstmat!familiacol), cadbl(rstmat!subfamiliacol), cadbl(rstmat!familiaad), cadbl(rstmat!subfamiliaad), cadbl(rstmat!subfamiliacompatible)
       End If
       Set rstmat = dbtmpb.OpenRecordset("select assignarstock,dataimpresiopacking,codigrupmaterialcompatible from comandes_extres where comanda=" + atrim(numc))
       If Not rstmat.EOF Then
         If IsDate(rstmat!dataimpresiopacking) Then
           etdataimpresio = "<-- Data Packing-List: " + format(rstmat!dataimpresiopacking, "dd/mm/yy hh:nn")
          Else: etdataimpresio = ""
         End If
         If rstmat!assignarstock Then
           assignarstock.BackColor = &H9AA6FA
             Else: assignarstock.BackColor = &H80FF80
         End If
         If cadbl(rstmat!codigrupmaterialcompatible) > 0 Then
             For i = 0 To Combocompatibles.ListCount - 1
                If Combocompatibles.ItemData(i) = cadbl(rstmat!codigrupmaterialcompatible) Then
                    Combocompatibles.ListIndex = i
                    Framecompatibles.Visible = True
                End If
             Next i
         End If
       End If
       'If Reserves.Caption <> "Reserves" Then comprovar_si_ja_te_materialreservat
   End If
End Sub
Sub comprovar_si_ja_te_materialreservat()
   Dim rstxcomanda As Recordset
   Set rstxcomanda = dbtmp.OpenRecordset("select sum(metres) as suma from percomandaoclient where (idcompra=null or idcompra<1) and numcomanda=" + atrim(cadbl(comanda)))
   Command2.Enabled = True
   If Not rstxcomanda.EOF Then
      If cadbl(rstxcomanda!suma) > 0 Then
        If MsgBox("Aquesta comanda ja te resevats " + atrim(cadbl(rstxcomanda!suma)) + " Metres, si fas Si les reserves es sumaran.", vbYesNo, "Atenció") = vbNo Then
           Command2.Enabled = False
        End If
      End If
   End If
  
End Sub
   
Sub reserva_dades_capcalera(numreserva As Double, idpalet As Double)
  Dim rstinfo As Recordset
  Dim campespesor As String
  campespesor = "espesor"
  
  Set rstinfo = dbtmp.OpenRecordset("select * from reserves where idreserva=" + atrim(numreserva))
  If rstinfo.EOF Then
    Set rstinfo = dbtmp.OpenRecordset("select * from palets where idpalet=" + atrim(idpalet))
    campespesor = "micres": etmicres = "Micres"
    If cadbl(iespesor) <= 0 Then campespesor = "grmsm2"
  End If
  If rstinfo.EOF Then Exit Sub
  iample = cadbl(rstinfo!ample)
  iplegat = cadbl(rstinfo!plegat)
  icares = atrim(rstinfo!carestractat)
  If icares = "" Or icares = "0" Then icares = "N"
  iobert = atrim(rstinfo!obert)
  imicrop.Value = IIf(rstinfo!microperforat, 1, 0)
  itl = atrim(rstinfo!semielaborat)
  isolapa = atrim(rstinfo!solapa)
  iespesor = cadbl(rstinfo.Fields(campespesor))
  If campespesor = "grmsm2" Then iespesor = iespesor * -1: etmicres = "Grms/m2"
  If campespesor = "micres" Or campespesor = "grmsm2" Then Set rstinfo = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstinfo!codimatprognou)))
  If Not rstinfo.EOF Then possar_familia_subfamilia cadbl(rstinfo!familia), cadbl(rstinfo!subfamilia), cadbl(rstinfo!familiacol), cadbl(rstinfo!subfamiliacol), cadbl(rstinfo!familiaad), cadbl(rstinfo!subfamiliaad)
  
End Sub
Sub marcar_fila(palet As Double, bobina As Double, metres As Double, utilitzada As Boolean)
  Dim i As Double
  i = 1
  While reixa.Rows > i
     If cadbl(reixa.TextMatrix(i, columnadelcamp("palet"))) = palet And cadbl(reixa.TextMatrix(i, columnadelcamp("bobina"))) = bobina Then
        reixa.TextMatrix(i, columnadelcamp("mtrsassignats")) = formatreixa(metres)
        If Not utilitzada Then
          reixa.TextMatrix(i, columnadelcamp("mtrsdisponibles")) = formatreixa(cadbl(reixa.TextMatrix(i, columnadelcamp("mtrsdisponibles"))) + cadbl(metres))
          reixa.TextMatrix(i, columnadelcamp("mtrsdiferencia")) = formatreixa(cadbl(reixa.TextMatrix(i, columnadelcamp("mtrsdisponibles"))) - reixa.TextMatrix(i, columnadelcamp("mtrsassignats")))
        End If
        reixa.row = i
        reixa.col = columnadelcamp("seleccionat")
        
        reixa.Text = "1"
        If utilitzada Then
           reixa.Text = "2"
           Set reixa.CellPicture = check2.Picture
          Else: Set reixa.CellPicture = check.Picture
        End If
        i = reixa.Rows
     End If
    i = i + 1
  Wend
End Sub
Sub possar_familia_subfamilia(fammate As Double, sfammate As Double, famcolo As Double, sfamcolo As Double, famadi As Double, sfamadi As Double, Optional sfammatcompatible As Double)
  Dim rst As Recordset
 carregar_combo_families
 'posso families de materials
 For i = 0 To fammat.ListCount - 1
   If fammat.ItemData(i) = fammate Then
      fammat.ListIndex = i
      i = fammat.ListCount
   End If
 Next i
 carregar_subfamilies subfammat
 For i = 0 To subfammat.ListCount - 1
   If subfammat.ItemData(i) = sfammate Then
      subfammat.ListIndex = i
      i = subfammat.ListCount
   End If
 Next i
 'poso families de colorants
 For i = 0 To famcol.ListCount - 1
   If famcol.ItemData(i) = famcolo Then
      famcol.ListIndex = i
      i = famcol.ListCount
   End If
 Next i
 carregar_subfamilies subfamcol
 For i = 0 To subfamcol.ListCount - 1
   If subfamcol.ItemData(i) = sfamcolo Then
      subfamcol.ListIndex = i
      i = subfamcol.ListCount
   End If
 Next i
 'poso families de aditius
 For i = 0 To famad.ListCount - 1
   If famad.ItemData(i) = famadi Then
      famad.ListIndex = i
      i = famad.ListCount
   End If
 Next i
 carregar_subfamilies subfamad
 For i = 0 To subfamad.ListCount - 1
   If subfamad.ItemData(i) = sfamadi Then
      subfamad.ListIndex = i
      i = subfamad.ListCount
   End If
 Next i
 etsubfamcompatible = "": etsubfamcompatible.Tag = ""
 If sfammatcompatible > 0 Then
   Set rst = dbtmpb.OpenRecordset("select codi,descripcio from subfamiliesmaterials where codi=" + atrim(sfammatcompatible))
   If Not rst.EOF Then
    etsubfamcompatible = atrim(rst!descripcio)
    etsubfamcompatible.Tag = rst!codi
   End If
 End If
 Set rst = Nothing
End Sub
Function micresmaterial(codimesuralineal As Byte, espesor As Double, tubolam As String) As String
  Dim rstmesural As Recordset
  Dim descripcio As String
 ' Dim r As String
  Set rstmesural = dbtmpb.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(codimesuralineal))
  If rstmesural.EOF Then Exit Function
  descripcio = rstmesural!descripcio
  r = espesor
  If descripcio = "GALGUES" Then
            If tubolam = "T" Then
                 r = format(espesor / 4, "#,##0.0")
                  Else: r = format(espesor / 2, "#,##0.0")
            End If
  End If
  'If InStr(1, descripcio, "GR/") > 0 Then
  '  micresmaterial = espesor * -1
  'End If
  descripcio = IIf(descripcio = "MICRES", "Mic", descripcio)
  descripcio = IIf(descripcio = "GALGUES", "Mic", descripcio)
  If InStr(1, descripcio, "GR/") > 0 Then
     descripcio = "GR/MT2"
     r = cadbl(r) * -1
  End If
     
  micresmaterial = r
  r = descripcio
End Function
Function aatrim(va As Variant) As String
   aatrim = atrim(va)
   If aatrim = "" Then aatrim = "N"
End Function
Function buscar_metres_necessaris() As Double
  Dim rstmtrs As Recordset
  Dim bandes As Byte
  Set rstmtrs = dbtmpb.OpenRecordset("select producte,rebmtrs,cantitatex,amplereb,simulteneitatreb,ampleesq from comandes where comanda=" + atrim(cadbl(comanda)))
  If Not rstmtrs.EOF Then
    bandes = IIf(cadbl(rstmtrs!simulteneitatreb) > 0, cadbl(rstmtrs!simulteneitatreb), 1)
    buscar_metres_necessaris = IIf(rstmtrs!rebmtrs > 0, rstmtrs!rebmtrs / bandes, cadbl(rstmtrs!cantitatex))
      
      If rstmtrs!producte = "PC" Or rstmtrs!producte = "PC2" Then buscar_metres_necessaris = rstmtrs!cantitatex
      mtrsnecessaris.Tag = IIf(cadbl(rstmtrs!amplereb) > 0, cadbl(rstmtrs!amplereb), cadbl(rstmtrs!ampleesq))
  End If
End Function

Private Sub comanda_LostFocus()
 'If Reserves.Caption = "Reserves" Then
  If Not sonnumeros(comanda) Then
     If comanda <> "" Then MsgBox "El numero de comanda o grup no pot ser zero ni alfanumèric.", vbCritical, "Error"
     Exit Sub
  End If
  If cadbl(comanda) < 10000 Then
    carregar_grupdepalets
  End If
 'End If
End Sub
Function sonnumeros(vn As String) As Boolean
   Dim i As Byte
   If Len(vn) = 0 Then Exit Function
   sonnumeros = True
   For i = 1 To Len(vn)
      If Not IsNumeric(Mid(vn, i, 1)) Then sonnumeros = False
   Next i
End Function
Sub carregar_grupdepalets()
  Dim rstinfo As Recordset
  Dim rstmat As Recordset
  Set rstinfo = dbtmp.OpenRecordset("select * from grupsdepalets where numerogrup=" + atrim(cadbl(comanda)))
  If Not rstinfo.EOF Then
       Set rstmat = dbtmpb.OpenRecordset("Select * from materials where codi=" + atrim(cadbl(rstinfo!codimatprognou)))
       'MsgBox rstinfo!materialex
       codiclient = ""
       nomclient = ""
       'Set rstcli = dbtmpb.OpenRecordset("select nom from clients where codi=" + codiclient)
       'If Not rstcli.EOF Then nomclient = rstcli!nom
       itl = aatrim(rstinfo!semielaborat)
       iobert = aatrim(rstinfo!obert)
       imicrop = cabool(rstinfo!microperforat)
       infotuboobertmicro = aatrim(rstinfo!semielaborat) + " - " + aatrim(rstinfo!obert) + " - " + aatrim(rstinfo!microperforat)
       imicrop.Tag = " and semielaborat='" + aatrim(rstinfo!semielaborat) + "' and obert='" + aatrim(rstinfo!obert) + "' and " + IIf(cabool(rstinfo!microperforat), "", "not") + " microperforat"
       icares = atrim(rstinfo!carestractat)
       If icares = "0" Then icares = "N"
       infometres = atrim(rstinfo!nomdelgrup)
       iample = atrim(rstinfo!ample)
       iplegat = atrim(rstinfo!plegat)
       isolapa = atrim(rstinfo!solapa)
       'infoampleplegat = atrim(rstinfo!ampleesq) + " / " + atrim(rstinfo!plegatesq)
       iplegat.Tag = atrim(rstinfo!ample)
       iespesor = IIf(cadbl(rstinfo!micres) > 0, cadbl(rstinfo!micres), cadbl(rstinfo!grmsm2) * -1)
       
       etmicres = IIf(cadbl(iespesor) >= 0, " Micres", " Grm/m2")
       If Not rstmat.EOF Then
          'infodescripciomat = atrim(rstmat!descripcio)
          infodescripciomat = atrim(rstmat!descripcio) + Chr(10) + Chr(13) + descripciomaterial(rstmat)
          infodescripciomat.Tag = rstmat!codi
          If rstmat!codi < 500 Then infodescripciomat.ForeColor = QBColor(12)
          infodescripciomat.ForeColor = infometres.ForeColor
          possar_familia_subfamilia cadbl(rstmat!familia), cadbl(rstmat!subfamilia), cadbl(rstmat!familiacol), cadbl(rstmat!subfamiliacol), cadbl(rstmat!familiaad), cadbl(rstmat!subfamiliaad), cadbl(rstmat!subfamiliacompatible)
       End If
       If cadbl(rstinfo!codigrupmaterialscompatibles) > 0 Then
          carregar_combo_compatibles rstinfo!codigrupmaterialscompatibles
          Framecompatibles.Visible = True
          'Combocompatibles = atrim(rstinfo!nomgrupmaterialscompatibles)
           Else: Framecompatibles.Visible = True = False
       End If
       'Set rstmat = dbtmpb.OpenRecordset("select dataimpresiopacking from comandes_extres where comanda=" + atrim(numc))
       'If Not rstmat.EOF Then
       '  If IsDate(rstmat!dataimpresiopacking) Then
       '    etdataimpresio = "<-- Data Packing-List: " + Format(rstmat!dataimpresiopacking, "dd/mm/yy hh:nn")
       '   Else: etdataimpresio = ""
       '  End If
       'End If
       'If Reserves.Caption <> "Reserves" Then comprovar_si_ja_te_materialreservat
       mtrsnecessaris = "999.999"
       Else: MsgBox "No hi ha cap grup de palets amb aquest codi"
   End If
End Sub

Private Sub Combocompatibles_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Command1_Click()
 If assignarmat.Visible = False Then Exit Sub
  'assignarmat.Hide
  Form1.Visible = True
  Form1.SetFocus
  Form1.WindowState = 0
  AppActivate "Manteniment de Palets"
  assignarmat.Hide
'  wait 1
'  Form1.Show
End Sub
Function columnadelcamp(nom As String) As Integer
  Dim i As Byte
  nom = UCase(nom)
  If reixa.Cols < 3 Then columnadelcamp = 0: Exit Function
  For i = 0 To reixa.Cols - 1
    If UCase(rstconsulta.Fields(reixa.ColData(i)).Name) = nom Then columnadelcamp = i: GoTo fi
  Next i
fi:
End Function
Function possartoteslescomandesxrreservar()
   Dim comandes As String
'escric les comandes si es que n'hi ha mes d'una a reservar
   Set rst = dbtmp.OpenRecordset("select * from pendentsdereservar where not reservar")
   If rst.EOF Then comandes = comanda
   While Not rst.EOF
     comandes = comandes + " " + atrim(rst!comanda)
     rst.MoveNext
   Wend
   possartoteslescomandesxrreservar = comandes
End Function
Private Sub Command2_Click()
   Dim numpalet As Integer
   Dim numbobina As Integer
   Dim ampleres As Double
   Dim numc As String
   Dim msg As String
   Dim i As Integer
   Dim comandes As String
   Dim rstc As Recordset
   If missatgepantalla.Visible Then Exit Sub
   
   If Not sonnumeros(comanda) Then MsgBox "El numero de comanda o grup no pot ser zero ni alfanumèric.", vbCritical, "Error": Exit Sub
   
   If Reserves.Caption <> "Reserves" Then
     comandes = possartoteslescomandesxrreservar
     Set rstc = dbtmp.OpenRecordset("select * from pendentsdereservar where not reservar")
     If cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("ample"))) > 0 Then
      If codiclient <> "" And cadbl(comanda) = 0 Then
         msg = "Reservare aquest material pel client " + nomclient
        Else: msg = "Reservare aquest material per la comanda: " + comandes
      End If
      If MsgBox(msg, vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
       ampleres = cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("ample")))
       If rstc.EOF Then
        'si nomes estic reservant la que tinc a pantalla
         metresareservar = cadbl(mtrsnecessaris)
         reservarmat comanda
           Else
            'si reservo totes les que tinc seleccionades a compra
             While Not rstc.EOF
               metresareservar = cadbl(rstc!metres)
               reservarmat atrim(rstc!comanda)
               rstc.MoveNext
             Wend
             dbtmp.Execute "delete * from pendentsdereservar"
       End If
       filtrar_materials
       ensenyar_comandes cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("idreserva")))
       If cadbl(mtrsnecessaris) > metresareservar Then
          compramat.mirarampleareserva ampleres
          botocomprar_Click
       End If
      End If
     End If
     Exit Sub
   End If
   
    
   If cadbl(comanda) > 0 Then
        If Not avis_vindraelclientBAT(comanda) Then Exit Sub
        'comprovo que hi hagi suficients metres sel.lecionats
        If cadbl(metressel) < cadbl(mtrsnecessaris) Then
           If MsgBox("Hi ha menys metres sel.leccionats dels necessaris." + Chr(10) + "Vols continuar igualment?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
        End If
      If midesdiferents Then If MsgBox("Atenció hi han mides diferents sel.leccionades, ES CORRECTE?", vbCritical + vbYesNo, "Atenció mides diferents") = vbNo Then Exit Sub
      If Not sicaltreurecomandadereclamades(cadbl(comanda)) Then Exit Sub
      If materialsdirefentsdeBACAICOA(cadbl(comanda)) Then Exit Sub
      assignar_material comanda
      
      If esimpresores(cadbl(comanda)) Then
         comprovarsihihaamuntadoraeltreballiavisar cadbl(comanda) 'aviso si hi ha un treball igual a muntadora
         opcionsmaterialajust
      End If
   End If
   If llistaaltrescomandes.ListCount > 0 Then
     If MsgBox("Hi ha multiples comandes sel.leccionades." + Chr(10) + Chr(13) + "Segur que vols assignar-les totes?", vbYesNo, "Atenció") = vbYes Then
       For i = 0 To llistaaltrescomandes.ListCount - 1
         numc = llistaaltrescomandes.List(i)
         If cadbl(numc) > 0 Then
            If Not sicaltreurecomandadereclamades(cadbl(comanda)) Then Exit Sub
            assignar_material numc
            comprovarsihihaamuntadoraeltreballiavisar cadbl(numc) 'aviso si hi ha un treball igual a muntadora
         End If
       Next i
     End If
   End If
   If matprovproves.Value = 1 Then
      If MsgBox("Vols actualitzar el llistat de Materials de proves al Drive?", vbInformation + vbYesNo + vbDefaultButton1, "Atenció") = vbYes Then
         ratoli "espera"
         treurellistatmaterialprovesdrive
         ratoli "normal"
      End If
   End If
'   re_reservar_bobines
End Sub
Function materialsdirefentsdeBACAICOA(vnumc As Double) As Boolean
  Dim i As Long
  Dim vsql As String
  Dim rst As Recordset
  For i = 0 To reixa.Rows - 1
     If reixa.TextMatrix(i, columnadelcamp("seleccionat")) <> "0" Then
       If InStr(1, reixa.TextMatrix(i, columnadelcamp("proveidor")), "BACAICOA") Then GoTo revisarbacaicoa
     End If
  Next i
  Set rst = Nothing
  Exit Function
revisarbacaicoa:
  Set rst = dbtmp.OpenRecordset("select linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then
      If rst!linkcomanda1 = 0 Then GoTo fi
      vs = "(Parcials.comanda='" + atrim(rst!linkcomanda1) + "'"
      vs = vs + IIf(rst!linkcomanda2 = 0, "", " or Parcials.comanda='" + atrim(rst!linkcomanda2) + "'")
      vs = vs + ")"
      vsql = "SELECT Parcials.comanda, proveidors.nom FROM (Palets LEFT JOIN Parcials ON Palets.Idpalet = Parcials.idpalet) LEFT JOIN (materials LEFT JOIN proveidors ON materials.proveidor = proveidors.codi) ON Palets.codimatprognou = materials.codi "
      vsql = vsql + " WHERE " + vs + " AND (proveidors.nom Like '*BACAICOA*');"
      Set rst = dbtmp.OpenRecordset(vsql)
      If Not rst.EOF Then
          materialsdirefentsdeBACAICOA = True
          If UCase(InputBox("AQUESTA COMANDA HI HA MATERIALS COMPLEXES DE BACAICOA NO ES RECOMENABLE LAMINAR-LOS JUNTS." + vbNewLine + vbNewLine + "Escriu [MATERIALBACAICOA] per fer-ho igualment.", "ATENCIÓ MATERIALS BACAICOA")) = "MATERIALBACAICOA" Then
             materialsdirefentsdeBACAICOA = False
          End If
      End If
  End If
fi:
  Set rst = Nothing
End Function
Sub comprovarsihihaamuntadoraeltreballiavisar(vcomanda As Double)
    Dim rst As Recordset
    Dim vresp As String
    Dim vcos As String
    Dim rstc As Recordset
    Dim vnumtreball As Double
    
    Set rst = dbcomandes.OpenRecordset("select numtreball from comandes where comanda=" + atrim(vcomanda))
    If rst.EOF Then GoTo fi
    vnumtreball = cadbl(rst!numtreball)
    'Set rst = dbbaixes.OpenRecordset("SELECT muntadora_ordremuntatge.comanda, comandes.numtreball FROM muntadora_ordremuntatge INNER JOIN comandes ON muntadora_ordremuntatge.comanda = comandes.comanda where numtreball=" + atrim(vnumtreball))
   ' Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", "comandes.ini"))
    Set rst = dbbaixes.OpenRecordset("SELECT muntadora_ordremuntatge.comanda FROM muntadora_ordremuntatge ")
    Set rstc = dbcomandes.OpenRecordset("select numtreball,comanda from comandes")
    While Not rst.EOF And vnumtreball <> 0
       rstc.FindFirst "comanda=" + atrim(rst!comanda)
       If Not rstc.NoMatch Then
           If rstc!numtreball = vnumtreball Then GoTo cont
       End If
       rst.MoveNext
    Wend
cont:
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
 End Sub

Function avis_vindraelclientBAT(vnumc As String) As Boolean
  Dim vresp As String
  Dim rst As Recordset
  avis_vindraelclientBAT = True
  Set rst = dbcomandes.OpenRecordset("select clientvindraarevisarimpresio from comandes_extres where comanda=" + atrim(vnumc))
  If Not rst.EOF Then
    If rst!clientvindraarevisarimpresio Then
       vresp = UCase(InputBox("ALERTA!!!" + Chr(10) + "AQUESTA COMANDA NECESSITA OK DEL CLIENT." + Chr(10) + "Escriu SI per afegir-la igualment.", "VINDRÀ EL CLIENT."))
       If vresp = "SI" Then
            avis_vindraelclientBAT = True
         Else: avis_vindraelclientBAT = False
       End If
    End If
  End If
  Set rst = Nothing
End Function
Function eselcomplexa(vproducte As String) As Boolean
   eselcomplexa = False
   If vproducte = "PC" Or vproducte = "PC2" Or vproducte = "PCP" Then eselcomplexa = True
End Function
Function sicaltreurecomandadereclamades(comanda As Double) As Boolean
    Dim rst As Recordset
    Dim vresp As String
    sicaltreurecomandadereclamades = True
    If comanda < 10000 Then Exit Function
    Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", "comandes.ini"))
    Set rst = dbcomandes.OpenRecordset("select passaraimpresores,producte from comandesmesextres where comanda=" + atrim(comanda))
    
    If Not rst.EOF Then
      If cadbl(rst!passaraimpresores) = 0 Then
       vresp = ""
       If Not eselcomplexa(rst!producte) Then
          vresp = UCase(InputBox("Aquesta comanda està amb StandBy, vols activar-la?" + Chr(10) + "Escriu SI per activar-la.", "Standby"))
           Else: vresp = "No"
       End If
       If vresp = "SI" Then
         dbbaixes.Execute "update planificacio_reclamades set reactivada=true where numcomanda=" + atrim(comanda)
         'també la marco a passar a fabrica per si estigues parada
         dbcomandes.Execute "update comandes_extres set passaraimpresores=1 where comanda=" + atrim(comanda)
         sicaltreurecomandadereclamades = True
       End If
      End If
           Else: sicaltreurecomandadereclamades = False
    End If
    Set rst = Nothing
End Function

Sub treurellistatmaterialprovesdrive()
    Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  If Not existeix("c:" + Environ("HOMEPATH") + "\Google drive\Llistat de material proves") Then
     If MsgBox("No trovo la carpeta del drive local... " + Chr(10) + "c:" + Environ("HOMEPATH") + "\google drive\Llistat de material proves", vbYesNo + vbDefaultButton2, "Atencio") = vbNo Then Exit Sub
  End If
  If existeix("c:" + Environ("HOMEPATH") + "\Google drive\Llistat de material proves\materialproves.xls") Then
       Kill "c:" + Environ("HOMEPATH") + "\Google drive\Llistat de material proves\materialproves.xls"
  End If
  If existeix("c:" + Environ("HOMEPATH") + "\Google drive\Llistat de material proves\materialproves.pdf") Then
       Kill "c:" + Environ("HOMEPATH") + "\Google drive\Llistat de material proves\materialproves.pdf"
  End If
  
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistatmaterialdeproves.rpt", 1)
  oreport.DiscardSavedData
  oreport.ExportOptions.DiskFileName = "c:" + Environ("HOMEPATH") + "\Google drive\Llistat de material proves\materialproves.pdf" '"c:\temp\materialproves.xls"
  oreport.ExportOptions.PDFExportAllPages = True
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat 'crEFTExcel80Tabular
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  
  oreport.EnableParameterPrompting = False
  
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "palets.mdb"
  oreport.Database.Tables.Item(2).Location = rutadelfitxer(cami) + "comandes.mdb"
  
  oreport.Export False
  MsgBox "Fitxer actualitzat.", vbOKOnly, "Google Drive"
  ratoli "normal"
End Sub
Function esimpresores(numc As Double) As Boolean
   Dim rstcom As Recordset
  ' Set rstcom = dbtmpb.OpenRecordset("SELECT comandes.comanda, productes.ruta as ruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(numc) + "));")
   Set rstcom = dbtmpb.OpenRecordset("SELECT comandes.comanda, producte FROM comandes where comandes.comanda=" + atrim(numc) + ";")
   If rstcom.EOF Then Exit Function
   Set rstcom = dbtmpb.OpenRecordset("select ruta from productes where codi='" + atrim(rstcom!producte) + "'")
   esimpresores = False
   If Not rstcom.EOF Then
      If InStr(1, rstcom!ruta, "I") > 0 Then esimpresores = True
   End If
End Function
Sub assignar_material(comanda As String)
   Dim numpalet As Double
   Dim numbobina As Double
   Dim almenysunaassignada As Boolean
   Dim i As Long
   Dim vcodicompatible As Double
   Dim borrarassignacio As Boolean
   Dim vmetres As Double
   Dim vcomandatractatdoscares As Boolean
   Dim vassignattractatdoscares As Boolean
   
   vcomandatractatdoscares = IIf(InStr(1, infodescripciomat, "TRACTAT 2 CARES") > 0, True, False)
   
   vnofirmar = False
   numpalet = columnadelcamp("palet")
   numbobina = columnadelcamp("bobina")
   
   eliminar_assignat cadbl(comanda)
   For i = 0 To reixa.Rows - 1
     If reixa.TextMatrix(i, columnadelcamp("seleccionat")) = "1" And cadbl(reixa.TextMatrix(i, columnadelcamp("mtrsassignats"))) > 0 Then
      assignar_bobina comanda, reixa.TextMatrix(i, numpalet), reixa.TextMatrix(i, numbobina), cadbl(reixa.TextMatrix(i, columnadelcamp("mtrsassignats")))
      vassignattractatdoscares = IIf(InStr(1, reixa.TextMatrix(reixa.row, columnadelcamp("FAMILIES")), "TRACTAT 2 CARES") > 0, True, vassignattractatdoscares)
      vmetres = vmetres + cadbl(reixa.TextMatrix(i, columnadelcamp("mtrsassignats")))
      almenysunaassignada = True
     End If
     If reixa.TextMatrix(i, columnadelcamp("seleccionat")) = "2" Then almenysunaassignada = True
   Next i
   If almenysunaassignada Then
    baixaseccioextrusora comanda
    vcodicompatible = 0
    If Framecompatibles.Visible Then
          If Combocompatibles.ListIndex > -1 Then vcodicompatible = Combocompatibles.ItemData(Combocompatibles.ListIndex)
    End If
    dbtmpb.Execute ("update comandes_extres set codigrupmaterialcompatible=" + atrim(vcodicompatible) + " where comanda=" + atrim(CDbl(comanda)))
    'borro la data de creacio d´impresiopackinglist per si estava fet
    dbtmpb.Execute ("update comandes_extres set dataimpresiopacking='' where comanda=" + atrim(comanda))
    dbtmpb.Execute ("update comandes_extres set metresassignatspackinglist=" + atrim(vmetres) + " where comanda=" + atrim(comanda))
    If InStr(1, infodescripciomat, " MDO ") > 0 And vassignattractatdoscares <> vcomandatractatdoscares Then enviaremailgeneric "tintes@inplacsa.com", "Material MDO ASSIGNAT tractat 2 cares i a la comanda no. " + atrim(comanda), atrim(Now) + vbNewLine + "REVISAR LA TINTA SI ESTÀ ASSIGNADA I AJUSTAR-LA. "
    If Not vnofirmar Then firmar_fulla
    vnofirmar = False
    etdataimpresio = ""
    If Not des_reservar(comanda) Then MsgBox "No s'ha DESRESERVAT correctament.", vbCritical + vbOKOnly, "ATENCIÓ": Exit Sub
    MsgBox "Assignació realitzada i DESRESERVAT." '+ Chr(10) + Chr(13) + "Ara recarregaré la consulta per verificar l'assignació."
      Else
        baixaseccioextrusora comanda, True
        dbtmpb.Execute "update comandes set proximaseccio='E' where comanda=" + atrim(cadbl(comanda))
        dbtmpb.Execute ("update comandes_extres set dataimpresiopacking='' where comanda=" + atrim(comanda))
        dbtmpb.Execute ("update comandes_extres set metresassignatspackinglist=0 where comanda=" + atrim(comanda))
        dbtmpb.Execute ("update comandes_extres set assignarstock=false,mtrsassignatsestock=0 where comanda=" + atrim(CDbl(comanda)))
        vcodicompatible = 0
        If Framecompatibles.Visible Then
          If Combocompatibles.ListIndex > -1 Then vcodicompatible = Combocompatibles.ItemData(Combocompatibles.ListIndex)
        End If
        dbtmpb.Execute ("update comandes_extres set codigrupmaterialcompatible=" + atrim(vcodicompatible) + " where comanda=" + atrim(CDbl(comanda)))
        dbstocks.Execute ("delete * from opcionsdajust where comanda=" + atrim(cadbl(comanda)))
        MsgBox "Assignació de bobines esborrada correctament."
        borrarassignacio = True
   End If
   dbtmpb.Execute ("update comandes_extres set assignarstock=false,mtrsassignatsestock=0 where comanda=" + atrim(CDbl(comanda)))
   If Not borrarassignacio Then comprovarseccionoE comanda
   carregar_info_comanda comanda, True
'el desactivo per no recarregar despres d'assignar   filtrar_materials
   
End Sub
Sub firmar_fulla()
   dbtmp.Execute "update comandes_firmes set anulada=true,dataanulacio=now where (tipus='PK1' or tipus='PK2') and comanda=" + atrim(comanda)
   dbtmp.Execute "insert into comandes_firmes (comanda,usuari,tipus,data) values (" + atrim(comanda) + ",'" + nomordinador + "','PK1',now)"
End Sub

Sub comprovarseccionoE(numc)
   Dim rstc As Recordset
   Set rstc = dbtmpb.OpenRecordset("select proximaseccio from comandes where comanda=" + atrim(numc))
   If Not rstc.EOF Then
      If rstc!proximaseccio = "E" And siesmaterialexacte(cadbl(numc), True) = "" Then
       MsgBox "La comanda " + atrim(numc) + " no ha canviat de secció. Repetiu el procés d'assignació i aviseu a en Miquel."
      End If
   End If
   Set rstc = Nothing
End Sub
Function midesdiferents() As Boolean
  Dim midaant As Double
  Dim mida As Double
  Dim i As Long
  For i = 0 To reixa.Rows - 1
     If reixa.TextMatrix(i, columnadelcamp("seleccionat")) = "1" Then
       If midaant = 0 Then midaant = cadbl(reixa.TextMatrix(i, columnadelcamp("ample")))
       mida = cadbl(reixa.TextMatrix(i, columnadelcamp("ample")))
       If mida <> midaant Then midesdiferents = True
       midaant = mida
     End If
  Next i
End Function


Sub reservarmat(numc As String)
   Dim numreserva As Double
   Dim rstreserva As Recordset
   Dim rstxcomanda As Recordset
   
   
   'comprovo si ja hi ha material assignat x aquesta comanda
   'aqui s'haura de fer per client per controlar quan no hi ha comanda creada encara
   
   
   
   numreserva = cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("idreserva")))
   If numreserva = 0 Then
      numreserva = crear_novareserva
   End If
   Set rstreserva = dbtmp.OpenRecordset("select * from reserves where idreserva=" + atrim(cadbl(numreserva)))
   
   rstreserva.Edit
   
   'If cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("disponible"))) < cadbl(mtrsnecessaris) Then
   '   metresareservar = cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("disponible")))
   '   MsgBox "Reservaré " + atrim(metresareservar) + " Metres i la resta s'ha de comprar."
   'End If
   
   rstreserva!metresreservats = cadbl(rstreserva!metresreservats) + metresareservar
   reixa.TextMatrix(reixa.row, columnadelcamp("reservat")) = formatreixa(rstreserva!metresreservats)
   rstreserva.Update
   Set rstxcomanda = dbtmp.OpenRecordset("select *  from percomandaoclient where numcomanda=" + atrim(cadbl(numc))) '+ " and idreserva=" + atrim(cadbl(numreserva)))
   afegir_reserva rstxcomanda, numreserva, cadbl(numc)
   'For i = 0 To llistaaltrescomandes.ListCount - 1
   '   afegir_reserva rstxcomanda, numreserva, cadbl(llistaaltrescomandes.List(i))
   'Next i
   
   comprovarsihihaunacomandasemblantafabricaiavisar cadbl(numc)
   
End Sub
Sub comprovarsihihaunacomandasemblantafabricaiavisar(numc As Double)
   Dim rstclixes As Recordset
   Dim rstc As Recordset
   Dim dbclixes As Database
   Dim rstclixesm As Recordset
   Dim rstordremuntatge As Recordset
   Dim avisarmsg As String
   Set rstc = dbcomandes.OpenRecordset("select passaraimpresores from comandes_extres where comanda=" + atrim(numc))
   If rstc.EOF Then Exit Sub
   If cadbl(rstc!passaraimpresores) = 0 Then Exit Sub
   Set rstc = dbcomandes.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(numc))
   If rstc.EOF Then Exit Sub
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
   Set rstclixes = dbclixes.OpenRecordset("SELECT Clixes.marca, Modificacions.numerodelinia FROM Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball where modificacions.id_treball=" + atrim(cadbl(rstc!numtreball)) + " and modificacions.ordre=" + atrim(cadbl(rstc!numordremodificacio)))
   If rstclixes.EOF Then Exit Sub
   Set rstordremuntatge = dbclixes.OpenRecordset("select * from muntadora_ordremuntatge")
   While Not rstordremuntatge.EOF
        Set rstc = dbcomandes.OpenRecordset("select numtreball,numordremodificacio,impressio from comandes where comanda=" + atrim(cadbl(rstordremuntatge!comanda)))
        If rstc.EOF Then GoTo proxima
        If atrim(rstc!impressio) = "R" Then
         Set rstclixesm = dbclixes.OpenRecordset("SELECT Clixes.marca, Modificacions.numerodelinia FROM Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball where modificacions.id_treball=" + atrim(cadbl(rstc!numtreball)) + " and modificacions.ordre=" + atrim(rstc!numordremodificacio))
         If atrim(rstclixes!marca) = atrim(rstclixesm!marca) And cadbl(rstclixes!numerodelinia) = cadbl(rstclixesm!numerodelinia) And cadbl(rstclixesm!numerodelinia) > 0 Then
            If cadbl(rstordremuntatge!comanda) <> numc Then
                avisarmsg = rstordremuntatge!comanda
                rstordremuntatge.MoveLast
            End If
         End If
        End If
proxima:
        rstordremuntatge.MoveNext
   Wend
   If avisarmsg <> "" Then
      enviaremailgeneric "liniesimpresio@inplacsa.com", "URGENT COMANDA APUNT PER IMPRIMIR", "La comanda " + atrim(numc) + " acaba d'entrar i hi ha una altra semblant a muntadora, la " + atrim(avisarmsg)
   End If
   
   Set dbclixes = Nothing
   Set rstclixes = Nothing
   Set rstclixesm = Nothing
   Set rstc = Nothing
End Sub


Sub afegir_reserva(rstxcomanda As Recordset, numreserva As Double, numc As Double)
   Dim quantitat As Double
   Dim rstcom As Recordset
   rstxcomanda.AddNew
   rstxcomanda!idreserva = numreserva
   If codiclient <> "" And numc = 0 Then
       rstxcomanda!numclient = codiclient: quantitat = metresareservar
     Else:
       rstxcomanda!numcomanda = numc
       'Set rstcom = dbtmpb.OpenRecordset("select cantitatex from comandes where comanda=" + atrim(numc))
       quantitat = metresareservar
   End If
   rstxcomanda!metres = quantitat
   rstxcomanda.Update
End Sub

Function crear_novareserva(Optional rample As Double) As Double
  Dim rstres As Recordset
  Set rstres = dbtmp.OpenRecordset("reserves")
  rstres.AddNew
  If rample = 0 Then
     rstres!ample = Redondejar(cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("ample"))), 1)
       Else: rstres!ample = Redondejar(rample, 1)
  End If
  rstres!plegat = cadbl(iplegat)
  rstres!carestractat = atrim(icares)
  rstres!obert = atrim(iobert)
  rstres!microperforat = cabool(imicrop.Value)
  rstres!semielaborat = itl
  rstres!espesor = cadbl(iespesor)
  If fammat.ListIndex <> -1 Then
   rstres!familia = fammat.ItemData(fammat.ListIndex)
   rstres!subfamilia = subfammat.ItemData(subfammat.ListIndex)
  End If
  If famcol.ListIndex <> -1 Then
   rstres!familiacol = famcol.ItemData(famcol.ListIndex)
   rstres!subfamiliacol = subfamcol.ItemData(subfamcol.ListIndex)
  End If
  If famad.ListIndex <> -1 Then
    rstres!familiaad = famad.ItemData(famad.ListIndex)
    rstres!subfamiliaad = subfamad.ItemData(subfamad.ListIndex)
  End If
  rstres.Update
  rstres.Bookmark = rstres.LastModified
  crear_novareserva = rstres!idreserva
  Set rstres = Nothing
End Function
Sub eliminar_assignat(numc As Double)
  Dim rstparcial As Recordset
  Dim nump As Double
  Dim numb As Integer
  dbtmp.Execute "delete * from historic_packinglist where comanda='" + atrim(numc) + "'"
  dbtmp.Execute "update comandes_firmes set anulada=true,dataanulacio=now where (tipus='PK1' or tipus='PK2') and comanda=" + atrim(numc)
  Set rstparcial = dbtmp.OpenRecordset("select * from parcials where comanda='" + atrim(cadbl(numc)) + "'")
  While Not rstparcial.EOF
    nump = cadbl(rstparcial!idpalet)
    numb = cadbl(rstparcial!idbobina)
    If Not rstparcial!utilitzada Then
     rstparcial.Delete
    End If
    rstparcial.MoveNext
    actualitzar_metres_disponibles nump, numb
  Wend
End Sub
Sub reservar_bobina(rstpaletvell As Recordset, palet As Double, bobina As Integer, metres As Double, numcom As String)
    Dim rstreservat As Recordset
    Dim rstmaterial As Recordset
    Dim rstreserva As Recordset
    Dim numc As Double
    Dim r As String
    Dim r2 As String
    Dim idreserva As Double
    numc = cadbl(Mid(numcom, 1, IIf(InStr(1, numcom, "/") > 0, InStr(1, numcom, "/") - 1, Len(numcom))))
    Set rstmaterial = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(rstpaletvell!codimatprognou))
    If Not rstmaterial.EOF Then
      r = "ample=" + adec(rstpaletvell!ample) + " and plegat=" + adec(rstpaletvell!plegat)
      r = r + " and solapa=" + adec(rstpaletvell!solapa) + " and carestractat='" + atrim(rstpaletvell!carestractat + "'")
      r = r + " and obert='" + atrim(rstpaletvell!obert) + "' and microperforat=" + IIf(cabool(rstpaletvell!microperforat), "True", "False")
      r = r + " and semielaborat='" + atrim(rstpaletvell!semielaborat) + "' and espesor=" + adec(rstpaletvell!micres)
      r2 = "and familia=" + atrim(cadbl(rstmaterial!familia)) + " and subfamilia=" + atrim(cadbl(rstmaterial!subfamilia))
      r2 = r2 + " and familiacol=" + atrim(cadbl(rstmaterial!familiacol)) + " and subfamiliacol=" + atrim(cadbl(rstmaterial!subfamiliacol))
      r2 = r2 + " and familiaad=" + atrim(cadbl(rstmaterial!familiaad)) + " and subfamiliaad=" + atrim(cadbl(rstmaterial!subfamiliaad))
      
      Set rstreservat = dbtmp.OpenRecordset("select * from reserves where " + r + r2)
      If Not rstreservat.EOF Then
         'sumar els metres a aquesta reserva
         rstreservat.Edit
          rstreservat!metresreservats = cadbl(rstreservat!metresreservats) + metres
          rstreservat.Update
          idreserva = rstreservat!idreserva
         Else
           'crear reserva nova amb els metres corresponents
           idreserva = novareservadimportacio(rstpaletvell, metres, rstmaterial)
      End If
      If idreserva > 0 Then
          Set rstreserva = dbtmp.OpenRecordset("select * from percomandaoclient where numcomanda=" + atrim(numc))
          If Not rstreserva.EOF Then
              rstreserva.Edit
              rstreserva!metres = cadbl(rstreserva!metres) + metres
              rstreserva.Update
            Else:
              dbtmp.Execute "insert into percomandaoclient (idreserva,numcomanda,metres) values (" + atrim(idreserva) + "," + atrim(numc) + "," + atrim(cadbl(metres)) + ")"
          End If
          Set rstreserva = Nothing
      End If
    End If
End Sub
Function novareservadimportacio(rstpaletvell As Recordset, metres As Double, rstmaterial As Recordset) As Double
  Dim rstres As Recordset
  Set rstres = dbtmp.OpenRecordset("reserves")
  rstres.AddNew
  rstres!ample = rstpaletvell!ample
  rstres!plegat = rstpaletvell!plegat
  rstres!carestractat = rstpaletvell!carestractat
  rstres!obert = atrim(rstpaletvell!obert)
  rstres!microperforat = cabool(rstpaletvell!microperforat)
  rstres!semielaborat = rstpaletvell!semielaborat
  rstres!espesor = cadbl(rstpaletvell!micres)
  rstres!metresreservats = metres
  rstres!familia = cadbl(rstmaterial!familia)
  rstres!subfamilia = cadbl(rstmaterial!subfamilia)
  rstres!familiacol = cadbl(rstmaterial!familiacol)
  rstres!subfamiliacol = cadbl(rstmaterial!subfamiliacol)
  rstres!familiaad = cadbl(rstmaterial!familiaad)
  rstres!subfamiliaad = cadbl(rstmaterial!subfamiliaad)
  rstres.Update
  rstres.Bookmark = rstres.LastModified
  novareservadimportacio = rstres!idreserva
  Set rstres = Nothing
End Function

Sub assignar_bobina(comanda As String, palet As Double, bobina As Integer, metres As Double)
  Dim parcial As Recordset
  Dim rstpalet As Recordset
  Set rstpalet = dbtmp.OpenRecordset("select * from palets where idpalet=" + atrim(palet), dbOpenSnapshot, dbReadOnly)
  Set parcial = dbtmp.OpenRecordset("select * from parcials where not utilitzada and idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)))
  
  parcial.AddNew
  parcial!idpalet = palet
  parcial!idbobina = bobina
  parcial!metres = metres
  parcial!comanda = comanda
  parcial!orcomassignacio = comanda
  parcial!material = rstpalet!codimatprognou
  parcial.Update
  'dbtmp.Execute "update bobines set numcom='" + comanda + "' where idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina))
  actualitzar_metres_disponibles palet, bobina
  Set rstpalet = Nothing
  Set parcial = Nothing
End Sub
Sub actualitzar_metres_disponibles(palet As Double, bobina As Integer)
  Dim rstparcial As Recordset
  Dim total As Double
  total = 0
  'Dim rstbobina As Recordset
  'Set rstbobina = dbtmp.OpenRecordset("select * from bobines where idpalet=" + atrim(palet) + " and idbobina=" + atrim(cadbl(bobina)))
  Set rstparcial = dbtmp.OpenRecordset("select sum(metres) as total from parcials where cdbl(comanda)>0 and idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)), dbOpenSnapshot, dbReadOnly)
  If Not rstparcial.EOF Then total = cadbl(rstparcial!total)
  dbtmp.Execute "update bobines set disponible=mts-" + atrim(Redondejar(total)) + " where idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina))
  Set rstparcial = Nothing
End Sub
Function esrestu(palet As Double, bobina As Integer) As Boolean
  Dim rstparcial As Recordset
  Dim mtsdisponibles As Double
  Dim mts As Double
  esretu = False
  Set rstparcial = dbtmp.OpenRecordset("select disponible,mts from bobines where idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)), dbOpenSnapshot, dbReadOnly)
  If Not rstparcial.EOF Then mtsdisponibles = cadbl(rstparcial!disponible): mts = cadbl(rstparcial!mts)
  Set rstparcial = dbtmp.OpenRecordset("select metres,utilitzada  from parcials where not utilitzada and idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)), dbOpenSnapshot, dbReadOnly)
  If Not rstparcial.EOF Then
   rstparcial.MoveLast
   If mtsdisponibles <= 0 And rstparcial.RecordCount = 1 And rstparcial!metres < mts Then esrestu = True
     Else: If mtsdisponibles < mts And mtsdisponibles > 0 Then esrestu = True
  End If
End Function
Function esparcial(palet As Double, bobina As Integer) As Boolean
  Dim rstparcial As Recordset
  Dim mtsdisponibles As Double
  Dim mts As Double
  esparcial = False
  Set rstparcial = dbtmp.OpenRecordset("select disponible,mts from bobines where idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)), dbOpenSnapshot, dbReadOnly)
  If Not rstparcial.EOF Then mtsdisponibles = cadbl(rstparcial!disponible): mts = cadbl(rstparcial!mts)
  Set rstparcial = dbtmp.OpenRecordset("select metres,utilitzada  from parcials where not utilitzada and idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)), dbOpenSnapshot, dbReadOnly)
  If Not rstparcial.EOF Then
    rstparcial.MoveLast
    If rstparcial.RecordCount > 1 Then esparcial = True
    If rstparcial.RecordCount = 1 And mtsdisponibles > 0 Then esparcial = True
  End If
End Function


Private Sub Command3_Click()
   If Reserves.Caption <> "Reserves" Then
     imprimir_packinglistreserva cadbl(comanda), llistat
     Exit Sub
   End If
   If cadbl(comanda) > 0 Then
      imprimir_packinglist cadbl(comanda), llistat, IIf(etdataimpresio = "", True, False)
      carregar_info_comanda atrim(cadbl(comanda))
   End If
End Sub
Sub imprimir_packinglistreserva(numcomanda As Double, llistat As CrystalReport)
  Dim rstpalet As Recordset
  Dim rstpro As Recordset
  Dim rstmat As Recordset
  Dim rstbobina As Recordset
  Dim rstmaterial As Recordset
  Dim rstparcials As Recordset
  Dim dataimpresio As String
  Dim nomdelclient As String
  Dim codimat As Long
  
  If numcomanda < 1 Then Exit Sub
  obrir_dbllistats
  dbllistat.Execute "delete * from packinglistxcomanda"
  Set rstllistat = dbllistat.OpenRecordset("Packinglistxcomanda")
  Set rstparcials = dbtmp.OpenRecordset("SELECT percomandaoclient.numcomanda,percomandaoclient.metres, Reserves.* FROM Reserves INNER JOIN percomandaoclient ON Reserves.idreserva = percomandaoclient.idreserva WHERE ((idcompra<1 or idcompra is null) and (percomandaoclient.numcomanda=" + Trim(numcomanda) + "))")
  If rstparcials.EOF Then MsgBox "No hi ha cap reserva per aquesta comanda.", vbInformation, "Reserves": Exit Sub
  While Not rstparcials.EOF
     Set rstmaterial = dbtmp.OpenRecordset(" select codi from materials where " + crear_criteri_familia)
     If Not rstmaterial.EOF Then
       guardar_registre_taulatmp_reserva rstparcials, rstmaterial!codi
     End If
    rstparcials.MoveNext
  Wend
  Set rstmat = dbtmpb.OpenRecordset("SELECT comandes.comanda, clients.nom as nomclient FROM comandes INNER JOIN clients ON comandes.client = clients.codi WHERE (((comandes.comanda)=" + atrim(numcomanda) + "));")
  If Not rstmat.EOF Then nomdelclient = rstmat!nomclient
  dbllistat.Close
   'imprimir llistat
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "packinglistreservaxcomanda.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 For i = 0 To 20
  llistat.Formulas(i) = ""
 Next i
 llistat.Formulas(1) = "dataimpresio='" + format(Now, "long date", vbMonday) + " " + format(Now, "hh:nn") + "'"
 llistat.Formulas(0) = "comanda='" + format(numcomanda, "#,##0") + "'"
 llistat.Formulas(2) = "nomdelclient='" + nomdelclient + "'"
 
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 llistat.Action = 1
  'Unload llistat
  obrir_dbllistats

  Set rstpalet = Nothing
  Set rstpro = Nothing
  Set rstmat = Nothing
  Set rstbobina = Nothing
  Set rstmaterial = Nothing
  Set rstparcials = Nothing
  For i = 1 To 10
   llistat.Formulas(i) = ""
  Next i
End Sub

Sub guardar_registre_taulatmp_reserva(rstpalet As Recordset, codimaterial As Long)
   Dim kg As Double
   Dim ample As Double
   Dim rstm As Recordset
   rstllistat.AddNew
   rstllistat!ample = rstpalet!ample
   rstllistat!plegat = rstpalet!plegat
   rstllistat!solapa = rstpalet!solapa
   rstllistat!obert = IIf(rstpalet!obert = "", "N", rstpalet!obert)
   rstllistat!microperforat = rstpalet!microperforat
   rstllistat!semielaborat = rstpalet!semielaborat
   rstllistat!carestractat = IIf(rstpalet!carestractat = "", "N", rstpalet!carestractat)
   If cadbl(rstpalet!espesor) >= 0 Then
       rstllistat!micres = atrim(rstpalet!espesor) + " µ"
      Else: rstllistat!micres = atrim(cadbl(rstpalet!espesor) * -1) + " Gr/m²"
   End If
     rstllistat!material = descripciomaterial(rstpalet)
     Set rstm = dbtmp.OpenRecordset("select descripcio from materials where codi=" + atrim(codimaterial))
     If Not rstm.EOF Then rstllistat!nommaterial = atrim(rstm!descripcio)
     'rstllistat!familia = rstmaterial!refproducte
     'rstllistat!proveidor = rstpro!nom
   'End If
   
   rstllistat!metres = rstpalet!metres
   rstllistat!kilos = compramat.conversiokilos(codimaterial, cadbl(rstpalet!ample), cadbl(rstpalet!metres), cadbl(rstpalet!espesor), rstpalet!semielaborat, cadbl(rstpalet!solapa))
   
   'persaber els grams mt2
  ' kg = ((cadbl(rstmaterial!grmcm3) / 0.000001) * (cadbl(rstpalet!micres) * 0.000001) / 1000)
   'ample = cadbl(rstpalet!ample)
   'If (rstpalet!semielaborat <> "L") Then ample = cadbl(rstpalet!ample) * 2 + cadbl(rstpalet!solapa)
   'ample = ample / 100
   'rstllistat!kilos = kg * ample * rstparcials!metres
   'rstllistat!mtrsdisponibles = rstbobina!disponible
   'rstllistat!observacionsp = rstpalet!Observ
   'rstllistat!observacionsb = rstbobina!Obser
   'rstllistat!resto = resto
   
   rstllistat.Update
End Sub

Sub crear_bobina(rstbobvell As Recordset, rstb As Recordset)
Dim i As Byte
  If Not rstbobvell.EOF Then
     rstb.AddNew
     For i = 0 To rstbobvell.Fields.Count - 1
        rstb.Fields(rstbobvell.Fields(i).Name) = rstbobvell.Fields(i).Value
     Next i
     rstb!numcom = " "
     rstb!impres = True
     rstb.Update
     rstb.Bookmark = rstb.LastModified
  End If
End Sub
Sub crear_palet_nou(rstp As Recordset, rstpalvell As Recordset)
  Dim i As Byte
  Dim c As Byte
  If Not rstpalvell.EOF Then
     rstp.AddNew
     For i = 0 To rstpalvell.Fields.Count - 1
       If rstpalvell.Fields(i).Name = "Grmt2" Then
          rstp!grmsm2 = rstpalvell.Fields(i).Value
         Else
           If rstpalvell.Fields(i).Name <> "Idfam" And rstpalvell.Fields(i).Name <> "Idprov" Then
             'If rstpalvell.Fields(i).Name = "Disponible" Then
             '  If Not rstpalvell.Fields(i).Value Then Stop
             'End If
             
             rstp.Fields(rstpalvell.Fields(i).Name) = rstpalvell.Fields(i).Value
           End If
       End If
     Next i
     rstp.Update
     rstp.Bookmark = rstp.LastModified
  End If
  
End Sub
Sub obrestocksvells(dbpaletsvells)
 Dim camistocksvell As String
camistocksvell = llegir_ini("General", "ruta_stocksmdb", "comandes.ini")
'If camistocksvell = "{[}]" Then
camistocks = "c:\stockvell.mdb"
Set dbpaletsvells = OpenDatabase(camistocks, , True)

  
End Sub
Private Sub Command4_Click()
  Dim dbpaletsvells As Database
  Dim rstpalvell As Recordset
  Dim rstbobvell As Recordset
  Dim rstt As Recordset
  Dim rstb As Recordset
  Dim rstp As Recordset
  Dim com As String
  Dim metres As Double
  Dim numbobant As Integer
  Exit Sub
  obrestocksvells dbpaletsvells
  Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", "comandes.ini"))
  dbtmp.Execute "delete * from parcials "
  dbtmp.Execute "delete * from bobines "
  dbtmp.Execute "delete * from palets "
  dbtmp.Execute "delete * from percomandaoclient "
  dbtmp.Execute "delete * from reserves "
  dbtmp.Execute "delete * from compresmaterial "
  
  r = "SELECT  Bobines.Idpalet,Bobines.Numcom from bobines where Bobines.Numcom='0' and Palets.Idpalet=Bobines.Idpalet "
  Set rstpalvell = dbpaletsvells.OpenRecordset("SELECT DISTINCTROW Palets.Idpalet, Palets.Idprod,Palets.grmsm2, Palets.micres, Palets.carestractat, Palets.obert, Palets.microperforat, Palets.semielaborat, Palets.Disponible, Productes.Idfam, Productes.Idprov, Productes.Grmt2, Palets.Ample, Palets.Plegat, Palets.codimatprognou, Palets.Solapa, Palets.Numalb, Palets.Numlot, Palets.Numpalet, Palets.Datarec, Palets.Datarev, Palets.Observ, Palets.Mostrasino, Palets.Numpaletpro, Palets.Tractat FROM Productes INNER JOIN (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) ON Productes.Idprod = Palets.Idprod WHERE Palets.codimatprognou>499  ORDER BY Palets.Idpalet ASC;") '(Palets.Disponible=True and exists (" + r + "))
  Set rstp = dbtmp.OpenRecordset("palets")
  If Not rstpalvell.EOF Then rstpalvell.MoveLast: rstpalvell.MoveFirst
  While Not rstpalvell.EOF
     metres = 0
     numbobant = 9999
  '   Me.Caption = "Falten Palets: " + atrim(cadbl(rstpalvell.RecordCount) - cadbl(rstpalvell.AbsolutePosition)) + "  " + atrim(rstpalvell!idpalet)
     crear_palet_nou rstp, rstpalvell
     Set rstb = dbtmp.OpenRecordset("select * from bobines where idpalet=" + atrim(cadbl(rstp!idpalet)) + " order by idpalet,idbobina")
     Set rstbobvell = dbpaletsvells.OpenRecordset("select * from bobines where idpalet=" + atrim(cadbl(rstpalvell!idpalet)) + " order by idpalet,idbobina")
     While Not rstbobvell.EOF
        If rstbobvell!idbobina <> numbobant Then
           If metres > 0 Then
              dbtmp.Execute "update bobines set mts=" + atrim(metres) + " where idpalet=" + atrim(rstpalvell!idpalet) + " and idbobina=" + atrim(numbobant)
           End If
           If numbobant <> 9999 Then actualitzar_metres_disponibles rstpalvell!idpalet, numbobant
           metres = 0
           crear_bobina rstbobvell, rstb
           numbobant = rstbobvell!idbobina
        End If
        If cadbl(rstbobvell!numcom) > 0 Then
           assignar_bobina rstbobvell!numcom, rstb!idpalet, rstb!idbobina, cadbl(rstbobvell!mts)
           comprovar_bobina_utilitzada rstbobvell!numcom, rstb!idpalet, rstb!idbobina
        End If
        If Len(rstbobvell!numcomrev) > 2 And cadbl(rstbobvell!numcom) = 0 Then
           reservar_bobina rstpalvell, rstb!idpalet, rstb!idbobina, cadbl(rstbobvell!mts), rstbobvell!numcomrev
        End If
        metres = metres + cadbl(rstbobvell!mts)
        rstbobvell.MoveNext
     Wend
     
     If metres > 0 Then
          dbtmp.Execute "update bobines set mts=" + atrim(metres) + " where idpalet=" + atrim(rstpalvell!idpalet) + " and idbobina=" + atrim(numbobant)
     End If
     If numbobant <> 9999 Then actualitzar_metres_disponibles rstpalvell!idpalet, numbobant
     
proxim:
     rstpalvell.MoveNext
     
     DoEvents
     Me.Caption = "Falten Palets: " + atrim(cadbl(rstpalvell.RecordCount) - cadbl(rstpalvell.AbsolutePosition)) + "  " + atrim(rstpalvell!idpalet)
  Wend
  
  
  
  
  
  
  
  
  
  
  
End Sub
Sub comprovar_bobina_utilitzada(numc As Double, nump As Double, numb As Double)
   Dim rstcom As Recordset
   Dim com1 As Double
   Dim com2 As Double
   Dim com3 As Double
   Dim seccio As String
   Dim utilitzada As Boolean
   Dim rstproducte As Recordset
   Dim rstbaixes As Recordset
   Dim vruta As String
   Dim op As Byte
   Dim data As Date
   seccio = ""
   Set rstcom = dbtmpb.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
   If Not rstcom.EOF Then
     com1 = cadbl(rstcom!comanda)
     com2 = cadbl(rstcom!linkcomanda1)
     com3 = cadbl(rstcom!linkcomanda2)
     ordenar com1, com2, com3
     
     Set rstproducte = dbtmpb.OpenRecordset("select ruta from productes where codi='" + atrim(rstcom!producte) + "'")
     If Not rstproducte.EOF Then
       
       vruta = rstproducte!ruta
       
     End If
     op = 0
     data = IIf(IsNull(rstcom!dataentrega), Now, rstcom!dataentrega)
     Select Case numc
       Case Is = com1
          If InStr(1, vruta, "I") Then
            utilitzada = mirarsiacavada("impressorestot", com1)
            seccio = "I"
            Set rstbaixes = dbbaixes.OpenRecordset("select max(datafi) as data,max(operari) as op, max(numeromaquina) from impressores where comanda=" + atrim(com1))
            If Not rstbaixes.EOF Then
              If Not IsNull(rstbaixes!data) And Not IsNull(rstbaixes!op) Then
               op = cadbl(rstbaixes!op): data = rstbaixes!data
              End If
            End If
          End If
          If vruta = "E" Or vruta = "ER" Or vruta = "ES" Then
            If vruta = "E" Then vruta = "EE"
            If rstcom!proximaseccio = "T" Or rstcom!proximaseccio = "V" Then utilitzada = True: seccio = Mid(vruta, 2, 1)
          End If
             
       Case Is = com2
         utilitzada = mirarsiacavada("laminadorestot", com1)
         seccio = "L"
         Set rstbaixes = dbbaixes.OpenRecordset("select max(datafi) as data,max(operari) as op, max(numeromaquina) from laminadores where comanda=" + atrim(com1))
         If Not rstbaixes.EOF Then
           If Not IsNull(rstbaixes!data) And Not IsNull(rstbaixes!op) Then
            op = cadbl(rstbaixes!op): data = rstbaixes!data
           End If
         End If
       Case Is = com3
         utilitzada = mirarsiacavada("laminadorestot", com3)
         seccio = "L"
         Set rstbaixes = dbbaixes.OpenRecordset("select max(datafi) as data,max(operari) as op, max(numeromaquina) from laminadores where comanda=" + atrim(com1))
         If Not rstbaixes.EOF Then
           If Not IsNull(rstbaixes!data) And Not IsNull(rstbaixes!op) Then
            op = cadbl(rstbaixes!op): data = rstbaixes!data
           End If
         End If
     End Select
      Else: If numc > 10000 Then seccio = "T": utilitzada = True
   End If
   
   If Not utilitzada And Not rstcom.EOF Then
     If seccio = "" Then seccio = "T"
     If rstcom!proximaseccio = "T" Or rstcom!proximaseccio = "V" Then utilitzada = True
   End If
   
   If seccio <> "" And utilitzada Then
      dbtmp.Execute "update parcials set seccio='" + seccio + "',utilitzada=" + IIf(utilitzada, "True", "False") + " where idpalet=" + atrim(nump) + " and idbobina=" + atrim(numb) + " and comanda='" + atrim(numc) + "'"
      dbtmp.Execute "update parcials set operari=" + atrim(op) + ",data=#" + format(data, "yy/mm/dd") + "# where idpalet=" + atrim(nump) + " and idbobina=" + atrim(numb) + " and comanda='" + atrim(numc) + "'"
   End If
End Sub
Function mirarsiacavada(seccio As String, numc As Double) As Boolean
  mirarsiacavada = False
  Set rstbaixes = dbbaixes.OpenRecordset("select acavada from " + seccio + " where comanda=" + atrim(numc))
  If Not rstbaixes.EOF Then
     mirarsiacavada = cabool(rstbaixes!acavada)
  End If
  
End Function
Sub ordenar(com1 As Double, com2 As Double, com3 As Double)
    Dim num(3) As Double
    Dim i As Byte
    Dim t As Double
    i = 0
    num(i) = com1: i = i + 1
     num(i) = com2: i = i + 1
    num(i) = com3
    For j = 0 To 2
     For i = 1 To 2
      If num(i) < num(i - 1) Then t = num(i): num(i) = num(i - 1): num(i - 1) = t
     Next i
    Next j
    j = 10
    For i = 0 To 2
      If num(i) > 0 And j > 3 Then j = i
    Next i
    If j < 3 Then
        com1 = 0: com2 = 0: com3 = 0
        If j < 3 And num(j) > 0 Then com1 = num(j): j = j + 1
        If j < 3 And num(j) > 0 Then com2 = num(j): j = j + 1
        If j < 3 And num(j) > 0 Then com3 = num(j): j = j + 1
    End If
    
End Sub

Private Sub Command5_Click()
  If codiclient = "9999" Then
     codiclient = "": nomclient = ""
    Else
     codiclient = "9999"
     nomclient = "Stock Inplacsa"
     comanda = ""
  End If
End Sub

Private Sub Command6_Click()
Dim dbpaletsvells As Database
  Dim rstpalvell As Recordset
  Dim rstbobvell As Recordset
  Dim rstt As Recordset
  Dim rstb As Recordset
  Dim rstp As Recordset
  Dim com As String
  Dim bobina As Double
  Dim palet As Double
  Dim metres As Double
  Dim numbobant As Integer
  Dim linia As String
  
  Dim r As String
    r = "c:\llistatsitstock.txt"
    If Not existeix(r) Then
              Open r For Output As 1
          Else: Open r For Append As 1
    End If
    
    
  
  
  obrestocksvells dbpaletsvells
  Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", "comandes.ini"))
  Set rstb = dbtmp.OpenRecordset("select * from bobines")
  If Not rstb.EOF Then rstb.MoveLast: rstb.MoveFirst
  While Not rstb.EOF
    bobina = cadbl(rstb!idbobina)
    palet = cadbl(rstb!idpalet)
'    palet = 34759
'    bobina = 1
    Set rstbobvell = dbpaletsvells.OpenRecordset("SELECT Bobines.Idpalet, Bobines.Idbobina, Bobines.Sit  AS situacio From bobines where (((Bobines.Idpalet)=" + atrim(palet) + ") AND (not BOBINES.entregat) and ((Bobines.Idbobina)=" + atrim(bobina) + "));")
    If Not rstbobvell.EOF Then
     If atrim(rstb!sit) <> atrim(rstbobvell!situacio) Then
         rstb.Edit
         rstb!sit = rstbobvell!situacio
         rstb.Update
         Me.Caption = atrim(palet)
     End If
    End If
    rstb.MoveNext
  Wend
  Close 1
  End
End Sub

Private Sub Command7_Click()
  If cadbl(comanda) > 10000 Then
   If MsgBox("El canvi de material de la comanda afecta a totes les reserves i assignacions fetes a aquest material, RECORDA CORRETGIR LES ASSIGNACIONS I RESERVES.", vbCritical + vbYesNo, "CANVIAR MATERIAL ASSIGNAT A COMANDA") = vbYes Then
      demanar_nou_material comanda, 0, 0, 0
      carregar_info_comanda atrim(cadbl(comanda))
      filtrar_materials
   End If
  End If
End Sub

Private Sub desreservar_Click()
   r = ""
   des_reservar comanda
End Sub
Function des_reservar(numc As String) As Boolean
   Dim rstdesr As Recordset
   Dim rstdesr2 As Recordset
   Dim rstres As Recordset
   Dim msgdesr As String
   'Set rstdesr = dbtmp.OpenRecordset("SELECT percomandaoclient.numcomanda, percomandaoclient.numclient, Reserves.Ample,reserves.idreserva, percomandaoclient.metres,percomandaoclient.idcompra FROM Reserves INNER JOIN percomandaoclient ON Reserves.idreserva = percomandaoclient.idreserva WHERE (((percomandaoclient.numcomanda)=" + atrim(cadbl(numc)) + "));")
   Set rstdesr2 = dbtmp.OpenRecordset("SELECT * FROM percomandaoclient  WHERE (((percomandaoclient.numcomanda)=" + atrim(cadbl(numc)) + "));")
   des_reservar = True
   If rstdesr2.EOF Then
     If r <> "nopregunta" Then
      MsgBox "No hi ha cap reserva d'aquesta comanda."
     End If
      Exit Function
   End If
   Set rstdesr = dbtmp.OpenRecordset("select * from reserves where idreserva=" + atrim(rstdesr2!idreserva))
   If rstdesr.EOF Then
      If MsgBox("Aquesta comanda o client te una reserva sense contingut." + Chr(10) + "VOLS ELIMINAR-LA?", vbCritical + vbDefaultButton2 + vbYesNo, "Error") = vbYes Then
           rstdesr2.Delete
      End If
      Exit Function
   End If
   While Not rstdesr.EOF
      msgdesr = msgdesr + "Ample: " + format(cadbl(rstdesr!ample), "#,##0.0") + " cm <---> " + format(cadbl(rstdesr2!metres), "#,###,##0") + " Mtrs" + Chr(13) + Chr(10)
      rstdesr.MoveNext
   Wend
   If r <> "Info" And r <> "nopregunta" Then
     r = InputBox("Entra el numero de comanda per des-reservar, ha de coincidir amb la consulta." + Chr(13) + Chr(10) + msgdesr, "Comfirmació Des-Reservar")
     msgdesr = ""
   End If
   If cadbl(r) = cadbl(numc) Or r = "Info" Or r = "nopregunta" Then
      's = "(select distinct(idreserva) from percomandaoclient where numcomanda=" + comanda + ")"
      'Set rstdesr = dbtmp.OpenRecordset("SELECT percomandaoclient.numcomanda, percomandaoclient.numclient, Reserves.Ample,reserves.idreserva, percomandaoclient.metres,percomandaoclient.idcompra FROM Reserves INNER JOIN percomandaoclient ON Reserves.idreserva = percomandaoclient.idreserva WHERE (((percomandaoclient.numcomanda)=" + atrim(cadbl(comanda)) + ") and percomandaoclient.idreserva in " + s + ");")
      rstdesr.MoveFirst
      msgdesr = msgdesr + "Compres afectades: " + Chr(10) + Chr(13)
      While Not rstdesr.EOF
       Set rstres = dbtmp.OpenRecordset("select * from compresmaterial where not entregada and idreserva=" + atrim(cadbl(rstdesr!idreserva)))
       While Not rstres.EOF
          msgdesr = msgdesr + atrim(rstres!codimat) + "-" + atrim(rstres!descmat) + "    --->  NºCompra: " + atrim(rstres!numcompra) + Chr(13) + Chr(10)
          rstres.MoveNext
       Wend
       If r <> "Info" And cadbl(rstdesr2!idcompra) < 1 And cadbl(rstdesr2!numclient) = 0 And cadbl(rstdesr2!numcomanda) > 0 Then
         dbtmp.Execute "update reserves set metresreservats=metresreservats-" + atrim(cadbl(rstdesr2!metres)) + " where idreserva=" + atrim(cadbl(rstdesr2!idreserva))
         dbtmp.Execute "delete * from percomandaoclient where numcomanda=" + atrim(numc)
       End If
       rstdesr.MoveNext
      Wend
      
      If r <> "Info" Then
         dbtmp.Execute "delete * from percomandaoclient where numcomanda=" + atrim(cadbl(numc)) '+ " and (idcompra<1 or idcompra=null)"
      End If
      If (msgdesr <> "" And r <> "nopregunta") And msgdesr <> ("Compres afectades: " + Chr(10) + Chr(13)) Then MsgBox "Aquestes compres queden afectades per la Des-Reserva." + Chr(10) + Chr(13) + msgdesr
        Else: des_reservar = False
   End If
End Function

Private Sub etreixa_Click()

End Sub

Private Sub fammat_Click()
 Dim rstfam As Recordset
 If assignarmat.ActiveControl.Name = "fammat" Then infodescripciomat.Tag = ""
 If fammat.ListIndex <> -1 Then
    Set rstfam = dbtmp.OpenRecordset("select * from materials where familia=" + atrim(cadbl(fammat.ItemData(fammat.ListIndex))))
    If Not rstfam.EOF Then
        If cadbl(rstfam!grmm2) > 0 Then
            etmicres = "Grm/m2"
             Else: etmicres = "Micres"
        End If
    End If
 End If
End Sub

Private Sub fammat_DropDown()
  'carregar_combo_families
End Sub
Sub carregar_combo_families()
  Dim rstfam As Recordset
  
  Set rstfam = dbtmpb.OpenRecordset("select * from familiesmaterials where codi>499")
  fammat.Clear
  While Not rstfam.EOF
    fammat.AddItem atrim(rstfam!descripcio)
    fammat.ItemData(fammat.NewIndex) = cadbl(rstfam!codi)
    rstfam.MoveNext
  Wend
  Set rstfam = dbtmpb.OpenRecordset("select * from familiescolorants where codi>499")
  famcol.Clear
  While Not rstfam.EOF
    famcol.AddItem atrim(rstfam!descripcio)
    famcol.ItemData(famcol.NewIndex) = cadbl(rstfam!codi)
    rstfam.MoveNext
  Wend
  Set rstfam = dbtmpb.OpenRecordset("select * from familiesaditius where codi>499")
  famad.Clear
  While Not rstfam.EOF
    famad.AddItem atrim(rstfam!descripcio)
    famad.ItemData(famad.NewIndex) = cadbl(rstfam!codi)
    rstfam.MoveNext
  Wend
End Sub
Sub possarmissatge(msg As String)
   If msg <> "" Then
      missatgepantalla.Visible = True
      etmissatge = msg
      missatgepantalla.Left = (assignarmat.Width / 2) - (missatgepantalla.Width / 2)
     Else: missatgepantalla.Visible = False
   End If
   DoEvents
End Sub
Private Sub filtrar_Click()
  
  parafiltre.Tag = "0"
  filtrar_materials
  parafiltre.Visible = False
  ratoli "normal"
End Sub
Sub filtrar_materials()
  Dim vmaterialexacte As String
'If cadbl(comanda) = 0 Then MsgBox "Primer entra un numero de comanda.": Exit Sub
  parafiltre.Visible = True
  ratoli "espera"
    possarmissatge "Carregant materials..."
    reixa.Clear
    DoEvents
  reixa.Redraw = False
  actualitzar_consulta
  'reixa.Visible = False
  If Reserves.Caption <> "Reserves" Then
   ' comprovar_reserves_orfes
    'compramat.comprovar_reserves_negatives
  End If
  If Reserves.Caption = "Reserves" Then
    vmaterialexacte = siesmaterialexacte(cadbl(comanda))
    If vmaterialexacte = "" Then
          assignarstock.Enabled = True
        Else
            assignarstock.Enabled = False
    End If
    If Check1.Value <> 0 Then
       carregar_rstconsulta
         Else: carregar_rstconsulta_rst
     End If
      Else: carregar_rstconsulta_rst
  End If
  If parafiltre.Tag = "1" Then GoTo fi
  barraprogres 100, 100
  If Reserves.Caption <> "Reserves" Then
     agrupar_reserves
  End If
  If parafiltre.Tag = "1" Then GoTo fi
  etprogres = "Configurant reixa... "
  configurar_reixa
  If Reserves.Caption <> "Reserves" Then
      'possar nom a la reixa
     possarnomdelacapçalerareixa
  End If
  etprogres = "Amples de reixa... "
  DoEvents
  carregar_amples_reixa
  If parafiltre.Tag = "1" Then GoTo fi
  etprogres = "Poblant la reixa... "
  DoEvents
  If vfiltrebobinesdesdeimpresores Then wait 2
  poblar_reixa
  If parafiltre.Tag = "1" Then GoTo fi
  etprogres = "Marcar palets reservats... "
  DoEvents
  marcar_paletsseleccionats
fi:
  'reixa.Visible = True
  reixa.Redraw = True
  ratoli "normal"
  possarmissatge ""
  parafiltre.Visible = False
End Sub
Sub possarnomdelacapçalerareixa()
 Dim noms(10) As String
 noms(0) = "ample"
 noms(1) = "reservat"
 noms(2) = "disponible"
 noms(3) = "t.terra"
 noms(4) = "compralk"
 noms(5) = "compraep"
 noms(6) = "t.comprat"
 noms(7) = "t.total"
 
 For i = 0 To 7
     reixa.col = i
     'reixa.ColData(i) = i
     reixa.TextMatrix(0, i + 1) = UCase(noms(i))
  Next i
End Sub

Sub actualitzar_consulta()
   imicrop.Tag = " and palets.semielaborat='" + aatrim(itl) + "' and obert='" + aatrim(iobert) + "' and " + IIf(cabool(imicrop), "", "not") + " microperforat"
End Sub
Sub carregarcompresdaquestareserva(rstcomprat As Recordset, idreserva As Double, compratlk As Double, compratep As Double)
   Dim rstreserva As Recordset
   Dim r As String
   Dim r2 As String
   Dim vcriterifam As String
   Dim rstl As Recordset
   Dim vcont As Byte
   
   Set rstreserva = dbtmp.OpenRecordset("select * from reserves where idreserva=" + atrim(idreserva))
   If rstreserva.EOF Then Exit Sub
   'If rstreserva!ample = 89 Then Stop
   r = "ample=" + passaradecimalpunt(rstreserva!ample) + " and plegat=" + passaradecimalpunt(rstreserva!plegat)
   r = r + " and solapa=" + passaradecimalpunt(rstreserva!solapa) + " and carestractat='" + atrim(rstreserva!carestractat + "'")
   'r = r + " and obert='" + atrim(rstreserva!obert) + "' and microperforat=" + IIf(cabool(rstreserva!microperforat), "True", "False")
     'he tret l'obert no se si serà correcte l'Alicia trobava a faltar material perque
       ' s'havia fet reserves amb tractats diferents i es barrejaven
   r = r + " and microperforat=" + IIf(cabool(rstreserva!microperforat), "True", "False")
   r = r + " and semielaborat='" + atrim(rstreserva!semielaborat) + "' and " + IIf(rstreserva!espesor > 0, "micres=", "grmm2=") + passaradecimalpunt(IIf(rstreserva!espesor > 0, rstreserva!espesor, rstreserva!espesor * -1))
   vcriterifam = crear_criteri_familia(vcont)
'   MsgBox vcriterifam
   If vcont = 5 Then
    r2 = " and familia=" + atrim(cadbl(rstreserva!familia)) + " and subfamilia=" + atrim(cadbl(rstreserva!subfamilia))
    r2 = r2 + " and familiacol=" + atrim(cadbl(rstreserva!familiacol)) + " and subfamiliacol=" + atrim(cadbl(rstreserva!subfamiliacol))
    r2 = r2 + " and familiaad=" + atrim(cadbl(rstreserva!familiaad)) + " and subfamiliaad=" + atrim(cadbl(rstreserva!subfamiliaad))
      Else: r2 = r2 + " and " + vcriterifam
   End If
compratlk = 0
compratep = 0
   'Clipboard.Clear
   'Clipboard.SetText "select * from liniescompra where not totentregat  and (" + r + r2 + ")"
   'r2 = substituirtots(r2, "materials.", "")
   Set rstcomprat = dbcompres.OpenRecordset("select * from liniescompra where not totentregat  and (" + r + r2 + ")")
   'Set rstcomprat = dbcompres.OpenRecordset("select * from liniescompra where not totentregat and kgentregats=0 and (" + r + r2 + ")")
   While Not rstcomprat.EOF
     Set rstl = dbcompres.OpenRecordset("select * from comandesxlinia where idliniacompra=" + atrim(cadbl(rstcomprat!idliniacompra)))
     While Not rstl.EOF
       If cadbl(rstl!numcomanda) = 0 Then
           compratep = compratep + rstl!kgcompra
          Else: compratlk = compratlk + rstl!kgcompra
       End If
       rstl.MoveNext
     Wend
     rstcomprat.MoveNext
   Wend
   Set rstcomprat = dbcompres.OpenRecordset("select * from liniescompra where not totentregat  and (" + r + r2 + ")")
   Set rstreserva = Nothing
   Set rstl = Nothing
End Sub
Sub agrupar_reserves()
   Dim rstreserves As Recordset
   Dim rstpalets As Recordset
   Dim rstcomprat As Recordset
   Dim rst2 As Recordset
   Dim ample As Double
   Dim comprat As Double
   Dim compratlk As Double
   Dim compratep As Double
   Dim vreservagenericaperfamilies As Boolean
   
   vreservagenericaperfamilies = mirarsireservaperfamilies
   dbllistat.Execute "delete * from reservamaterial"
   Set rstreserves = dbllistat.OpenRecordset("SELECT ample, Sum(mtrsdisponibles)AS disponible, first(palet) as idpalet  From assignaciomaterial GROUP BY ample;")
   Set rstconsulta = dbllistat.OpenRecordset("reservamaterial")
   While Not rstreserves.EOF
     rstconsulta.AddNew
      rstconsulta!ample = rstreserves!ample
      rstconsulta!disponible = rstreserves!disponible
      rstconsulta!idreserva = 0
      rstconsulta!idpalet = IIf(vreservagenericaperfamilies, 0, rstreserves!idpalet)
      rstconsulta!saldoterra = rstreserves!disponible
     rstconsulta.Update
     rstreserves.MoveNext
   Wend
   ample = cadbl(iplegat.Tag)
   

   If InStr(1, criteridebusqueda, "micres") Then
      substituir criteridebusqueda, "micres=", "espesor="
      substituir criteridebusqueda, "micres>", "espesor>"
      substituir criteridebusqueda, "micres>=", "espesor>="
   End If
   If InStr(1, criteridebusqueda, "grmsm2") Then
     substituir criteridebusqueda, "grmsm2>=", "espesor<=-"
     substituir criteridebusqueda, "grmsm2>", "espesor<-"
   End If
   While InStr(1, criteridebusqueda, "palets.")
     substituir criteridebusqueda, "palets.", ""
   Wend
   If InStr(1, criteridebusqueda, " codimatprognou=") > 0 Then
     'criteridebusqueda = Mid(criteridebusqueda, InStr(1, criteridebusqueda, " and ") + 4)
     MsgBox "NO ES POT RESERVAR UNA COMANDA AMB MATERIAL ESPECIFIC, S'HA D'ASSIGNAR DIRECTAMENT.", vbCritical, "ERROR"
     GoTo fi
   End If
   r = criteridebusqueda
   
   
   'MsgBox criteridebusqueda
salt_reservats:
   ' Clipboard.Clear
    'Clipboard.SetText "select * from reserves where " + r
    'r = substituirtots(r, "materials.", "")

   Set rstreserves = dbtmp.OpenRecordset("select * from reserves where " + r) 'ample>=" + atrim(ample) + infotuboobertmicro.Tag)
   Set rstpalets = dbllistat.OpenRecordset("select * from assignaciomaterial")
   Set rstconsulta = dbllistat.OpenRecordset("select * from reservamaterial order by ample")
   While Not rstreserves.EOF
     Set rstconsulta = dbllistat.OpenRecordset("select * from reservamaterial order by ample")
     'Set rstcomprat = dbtmp.OpenRecordset("select sum(kilos) as kg,sum(metres) as mtrs,sum(kgpendents) as comprat from compresmaterial where not entregada and idreserva=" + atrim(cadbl(rstreserves!idreserva)))
     compratlk = 0
     compratep = 0
     carregarcompresdaquestareserva rstcomprat, cadbl(rstreserves!idreserva), compratlk, compratep
     i = 0
     While Not rstconsulta.EOF And i = 0
       If cadbl(rstconsulta!ample) = cadbl(rstreserves!ample) Then
          i = 1
         Else: rstconsulta.MoveNext
       End If
     Wend
     If r <> criteridebusqueda Then i = 0
     If i = 1 Then
            rstconsulta.Edit
          Else:
             rstconsulta.AddNew
             rstconsulta!ample = rstreserves!ample
             rstconsulta!disponible = 0
     End If
     
      'comprat = passaramtrsxreglad3(cadbl(rstcomprat!comprat), cadbl(rstcomprat!kg), cadbl(rstcomprat!mtrs))
      'passo lo comprat de kilos a metres
      compratlk = compramat.conversiokilos(cadbl(rstcomprat!codimaterial), cadbl(rstcomprat!ample), compratlk * -1, IIf(cadbl(rstcomprat!grmm2) > 0, cadbl(rstcomprat!grmm2) * -1, cadbl(rstcomprat!micres)), atrim(rstcomprat!semielaborat), cadbl(rstcomprat!solapa))
      compratlk = Redondejar(compratlk, 0)
      compratep = compramat.conversiokilos(cadbl(rstcomprat!codimaterial), cadbl(rstcomprat!ample), compratep * -1, IIf(cadbl(rstcomprat!grmm2) > 0, cadbl(rstcomprat!grmm2) * -1, cadbl(rstcomprat!micres)), atrim(rstcomprat!semielaborat), cadbl(rstcomprat!solapa))
      compratep = Redondejar(compratep, 0)
     
     rstconsulta!compratlk = cadbl(rstconsulta!compratlk) + compratlk
     rstconsulta!compratep = cadbl(rstconsulta!compratep) + compratep
     rstconsulta!idreserva = cadbl(rstreserves!idreserva)
     rstconsulta!reservat = cadbl(rstconsulta!reservat) + cadbl(rstreserves!metresreservats)
     'rstconsulta!perreservar = cadbl(rstreserves!pendentsreservar)
     If r <> criteridebusqueda Then
        rstconsulta!estareservat = True
        Set rst2 = dbllistat.OpenRecordset("select * from reservamaterial where ample=" + atrim(adec(rstreserves!ample)))
        If Not rst2.EOF Then
          
          'rstconsulta!comprat = rst2!comprat
          rstconsulta!disponible = rst2!disponible
        End If
       Else: rstconsulta!estareservat = False
     End If
     'If Not rstcomprat.EOF Then rstconsulta!comprat = rstcomprat!comprat
     'rstconsulta!disponible = cadbl(rstconsulta!disponible) '+ cadbl(rstconsulta!comprat) - (cadbl(rstreserves!metresreservats) + cadbl(rstreserves!pendentsreservar))
     rstconsulta!saldoterra = cadbl(rstconsulta!disponible) - cadbl(rstreserves!metresreservats)
     rstconsulta!saldocomprat = cadbl(rstconsulta!compratlk) + cadbl(rstconsulta!compratep)
     rstconsulta!saldototal = cadbl(rstconsulta!saldoterra) + cadbl(rstconsulta!compratep)
     If cadbl(rstconsulta!saldototal) = 0 Then
        rstconsulta.CancelUpdate
         Else: rstconsulta.Update
     End If
     rstreserves.MoveNext
   Wend
   Set rst2 = Nothing
   If r = criteridebusqueda Then
       r = ""
       If cadbl(comanda) > 0 Then
           r = " idreserva in (select idreserva from percomandaoclient where numcomanda=" + atrim(cadbl(comanda)) + ")"
         Else: If cadbl(codiclient) > 0 Then r = " idreserva in (select idreserva from percomandaoclient where numclient=" + atrim(cadbl(codiclient)) + ")"
       End If
     If r <> "" Then GoTo salt_reservats
   End If
   Set rstconsulta = dbllistat.OpenRecordset("select * from reservamaterial order by estareservat,ample")
   Set rstcomprat = Nothing
fi:
End Sub
Function mirarsireservaperfamilies() As Boolean
  Dim vcriterifam As String
  Dim vcont As Byte

  vcriterifam = crear_criteri_familia(vcont)
  mirarsireservaperfamilies = IIf(vcont = 5, False, True)
End Function




'Sub substituir(cadena As String, buscar As String, canviar As String)
'   comença = InStr(1, cadena, buscar)
'   If comença < 1 Then Exit Sub
'   comença = comença - 1
'   acaba = comença + Len(buscar) + 1
'   cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
'   'MsgBox linia
'End Sub
Function passaramtrsxreglad3(comprat As Double, kilos As Double, metres As Double) As Double
 If kilos > 0 Then
  passaramtrsxreglad3 = (metres * comprat) / kilos
 End If
End Function

Sub marcar_paletsseleccionats()
  Dim rstparcial As Recordset
  Dim mtrs As Double
  Set rstparcial = dbtmp.OpenRecordset("select * from parcials where comanda='" + atrim(cadbl(comanda)) + "'")
  mtrs = 0
  While Not rstparcial.EOF
       marcar_fila rstparcial!idpalet, rstparcial!idbobina, rstparcial!metres, rstparcial!utilitzada
       mtrs = mtrs + rstparcial!metres
       rstparcial.MoveNext
  Wend
  metressel = mtrs
End Sub
Sub poblar_reixa()
  Dim row As Integer
  Dim col As Integer
  Dim vample As Double
  Dim vampleant As Double
  Dim valor As String
  Dim rstparcial As Recordset
  Dim nhihaalmenysun As Boolean
  Dim vcont As Integer
  Dim vestemfentreservaperfamilies As Boolean
  
  Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
 ' reixa.Visible = False
  reixa.Rows = 2
  reixa.FillStyle = flexFillRepeat
  reixa.BackColor = QBColor(15)
  reixa.Redraw = False
  vestemfentreservaperfamilies = mirarsireservaperfamilies
  If Reserves.Caption <> "Reserves" And vestemfentreservaperfamilies Then reixa.BackColor = QBColor(11)
  nhihaalmenysun = True
  row = 1
  If Not rstconsulta.EOF Then
     rstconsulta.MoveFirst
    Else: MsgBox ("No hi ha registres"): Exit Sub
  End If
  While Not rstconsulta.EOF
    DoEvents
    ' etprogres = "Poblant... " + atrim(rstconsulta.AbsolutePosition) + "/" + atrim(rstconsulta.RecordCount)
     vampleant = vample
     vample = rstconsulta.Fields("ample")
     'posso el color corresponent a la fila
     If vample <> vampleant Then
      If nhihaalmenysun Then
        If ultimcolor = &HE8BBBB Then
           ultimcolor = &HD29F7D
          Else: ultimcolor = &HE8BBBB
        End If
      End If
      nhihaalmenysun = False
    End If
    If Reserves.Caption <> "Reserves" And vestemfentreservaperfamilies Then ultimcolor = QBColor(11)
    possar_color_fila
    
    'nomes refresco pantalla cada 20 registres per no elentir lacarrega
    If vcont = 20 Then
       DoEvents: cont = 0
       Else: cont = cont + 1
    End If
    
     If Reserves.Caption = "Reserves" Then
       bobinesdentrada.actualitzar_metres_disponibles rstconsulta!palet, rstconsulta!bobina
       Set rstparcial = dbtmp.OpenRecordset("select * from parcials where idpalet=" + atrim(rstconsulta!palet) + " and idbobina=" + atrim(rstconsulta!bobina) + " and comanda='" + atrim(cadbl(comanda)) + "'", dbOpenSnapshot, dbReadOnly)
       If rstparcial.EOF And rstconsulta!mtrsdisponibles <= 0 Then GoTo seguentregistres
       If rstconsulta!parcial Then
          reixa.CellBackColor = QBColor(13)
         Else: nhihaalmenysun = True
       End If
       If Not IsDate(rstconsulta!datarec) Then
          reixa.CellBackColor = &H80FF&
       End If
     End If
     

   
     'posso els checks si el camp es check i si no el valor corresponent a cada camp
    For col = 0 To rstconsulta.Fields.Count - 1
     
     If rstconsulta.Fields(col).Type = 1 Then
        reixa.col = col
        reixa.row = row
        
        reixa.TextMatrix(row, col) = IIf(rstconsulta.Fields(col), "1", "0")
        Set reixa.CellPicture = IIf(reixa.TextMatrix(row, col) = "0", nocheck.Picture, check.Picture)
        reixa.CellForeColor = reixa.CellBackColor
       
         Else:
           valor = formatreixa(rstconsulta.Fields(col))
           
           reixa.TextMatrix(row, col) = valor
           reixa.col = col
           reixa.row = row
           If cadbl(reixa.TextMatrix(row, col)) < 0 Then
              reixa.CellForeColor = QBColor(12)
             Else: reixa.CellForeColor = QBColor(0)
           End If
      End If
     Next col

     'canvio el color de la linia si correspont
   
      If Reserves.Caption <> "Reserves" Then
         If reixa.TextMatrix(reixa.row, columnadelcamp("estareservat")) = "1" Then
              possarcolorfila reixa, reixa.row, QBColor(14)
         End If
      End If
      If Reserves.Caption = "Reserves" Then
         If reixa.TextMatrix(reixa.row, columnadelcamp("seleccionat")) = "3" Then
              possarcolorfila reixa, reixa.row, QBColor(14)
         End If
      End If
      If Reserves.Caption = "Reserves" Then
         If UCase(Mid(reixa.TextMatrix(reixa.row, columnadelcamp("situacio")) + "  ", 1, 1)) = "N" Then
              possarcolorfila reixa, reixa.row, QBColor(8)
         End If
         'If UCase(Mid(reixa.TextMatrix(reixa.row, columnadelcamp("situacio")) + "  ", 1, 1)) = "F" Then
         '     possarcolorfila reixa, reixa.row, QBColor(14)
         'End If
      End If
      'posso el camp de palet de diferent color si te impost d'envasos
      If Reserves.Caption = "Reserves" Then
        If rstconsulta!impostenvasos Then
             reixa.col = 2
            reixa.row = row
            reixa.CellBackColor = QBColor(12)
        End If
      End If
    
    'incremento la fila
    reixa.Rows = row + 2
    row = row + 1
    reixa.row = row
seguentregistres:

    rstconsulta.MoveNext
  Wend
  If reixa.TextMatrix(reixa.Rows - 1, 1) = "" Then reixa.Rows = reixa.Rows - 1
  reixa.Redraw = True
  'reixa.Visible = True
End Sub
Function formatreixa(ByVal valor) As String
   If cadbl(valor) <> 0 Then
         If (cadbl(valor) - Int(cadbl(valor))) <> 0 Then
            valor = format(valor, "#,##0.0")
           Else: valor = format(valor, "#,##0")
         End If
   End If
   If IsNull(valor) Then valor = ""
   formatreixa = valor
End Function

Sub possar_color_fila()
    Dim color As Long
    Dim colorr As Long

    color = 0
    If color <> 0 Then
          colorr = color
        Else: colorr = ultimcolor
    End If
    possarcolorfila reixa, reixa.row, colorr
End Sub
Sub possar_bobina_reixa(rstpalet As Recordset, rstbobina As Recordset, sel As String)
   Dim rstpro As Recordset
   Dim rstmaterial As Recordset
   Dim rstparcial As Recordset
   Dim resto As Boolean
   'If rstbobina!numcomrev <> "" Or rstbobina!numcom <> "" Then
   ' Set rstparcial = dbtmp.OpenRecordset("select * from parcials where  comanda<>''  and idpalet=" + atrim(cadbl(rstbobina!idpalet)) + " and idbobina=" + atrim(cadbl(rstbobina!idbobina)))
   ' If Not rstparcial.EOF Then resto = True
   ' While Not rstparcial.EOF
   '  If Not rstparcial!utilitzada Then resto = False
   '  rstparcial.MoveNext
   ' Wend
   'End If
   resto = esrestu(rstpalet!idpalet, rstbobina!idbobina)
   Set rstmaterial = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstpalet!codimatprognou)), dbOpenSnapshot, dbReadOnly)
   If Not rstmaterial.EOF Then Set rstpro = dbtmpb.OpenRecordset("select nom from proveidors where codi=" + atrim(cadbl(rstmaterial!proveidor)), dbOpenSnapshot, dbReadOnly)
   'If rstbobina!disponible > 0 Then
    guardar_registre_taulatmp2 rstpalet, rstpro, rstbobina, rstmaterial, resto, sel
   Set rstmaterial = Nothing
   Set rstpro = Nothing
End Sub
Sub guardar_registre_taulatmp2(rstpalet As Recordset, rstpro As Recordset, rstbobina As Recordset, rstmaterial As Recordset, resto As Boolean, sel As String)
   Dim espe As Double
   espe = IIf(atrim(iespesor) = "", 1, cadbl(iespesor))
   rstconsulta.AddNew
   rstconsulta!seleccionat = sel
   rstconsulta!palet = rstpalet!idpalet
   rstconsulta!bobina = rstbobina!idbobina
   rstconsulta!ample = Redondejar(rstpalet!ample, 1)
   If espe > 0 Then
          rstconsulta!micres = rstpalet!micres
            'AIXÓ ES PER SI SON GRMS/M2 HO POSO EN NEGATIU
          If cadbl(rstpalet!grmsm2) > 0 Then rstconsulta!micres = (rstpalet!grmsm2) * -1
        Else: rstconsulta!micres = (rstpalet!grmsm2) * -1
   End If
   rstconsulta!plegat = rstpalet!plegat
   rstconsulta!solapa = rstpalet!solapa
   rstconsulta!numpaletprov = rstpalet!numpaletpro
   rstconsulta!codimat = rstpalet!codimatprognou
   rstconsulta!numlot = rstpalet!numlot
   If Not rstmaterial.EOF Then
     rstconsulta!material = rstmaterial!descripcio
     If Not rstpro.EOF Then rstconsulta!proveidor = rstpro!nom
     rstconsulta!families = descripciomaterial(rstmaterial)
   End If
   rstconsulta!familia = fammat.Text
   rstconsulta!tractat = rstpalet!tractat
   rstconsulta!datarec = rstpalet!datarec
   rstconsulta!situacio = rstbobina!sit
   rstconsulta!reserva = rstbobina!numcomrev
   rstconsulta!comanda = rstbobina!numcom
   rstconsulta!metres = rstbobina!mts
   rstconsulta!kilos = compramat.conversiokilos(cadbl(rstpalet!codimatprognou), cadbl(rstpalet!ample), cadbl(rstbobina!mts), cadbl(rstconsulta!micres), atrim(rstpalet!semielaborat), cadbl(rstpalet!solapa))
   rstconsulta!mtrsdisponibles = rstbobina!disponible
   rstconsulta!resto = resto
   rstconsulta!observacionsp = rstpalet!observ
   rstconsulta!observacionsb = rstbobina!obser
   rstconsulta!parcial = esparcial(rstpalet!idpalet, rstbobina!idbobina)
   rstconsulta!impostenvasos = cabool(rstpalet!teimpost)
   rstconsulta.Update
   
End Sub
Sub carregar_rstreserves()
  Dim rstsel As Recordset
  Dim rstparcial As Recordset
  Dim rstselbob As Recordset
  Dim numfam As String
  Dim numsubfam As String
  Dim ample As Double

  ample = cadbl(iplegat.Tag)
  If limitarample.Value = 0 Then ample = 0
  
  'Set rstparcial = dbtmp.OpenRecordset("SELECT Min(Palets.Ample) AS minample, Parcials.comanda FROM Parcials INNER JOIN Palets ON Parcials.idpalet = Palets.Idpalet GROUP BY Parcials.comanda HAVING (((Parcials.comanda)='" + comanda + "'));")
  'If Not rstparcial.EOF Then ample = rstparcial!minample
  'Set rstparcial = Nothing
  
  numsubfam = ""
  numfam = "0"
  If fammat.ListIndex >= 0 Then numfam = Trim(fammat.ItemData(fammat.ListIndex))
  If subfammat.ListIndex >= 0 Then numsubfam = " and materials.subfamilia=" + Trim(cadbl(subfammat.ItemData(subfammat.ListIndex)))
  If infodescripciomat.Tag <> "" Then numsubfam = aondelmaterial(cadbl(infodescripciomat.Tag))
  'treure aquesta linia proxima per filtra per espesors obert micro t o b etc...
  imicrop.Tag = ""
  dbllistat.Execute ("delete * from assignaciomaterial")
  Set rstconsulta = dbllistat.OpenRecordset("assignaciomaterial")
  Set rstsel = dbtmp.OpenRecordset("SELECT Palets.*,materials.familia, materials.subfamilia FROM Palets INNER JOIN materials ON Palets.codimatprognou = materials.codi WHERE (((materials.familia)=" + numfam + ")" + numsubfam + ") and palets.ample>=" + atrim(ample) + imicrop.Tag + ";")
  
  While Not rstsel.EOF
    Set rstselbob = dbtmp.OpenRecordset("select * from bobines where numcom='0'  and idpalet=" + atrim(cadbl(rstsel!idpalet)))
    While Not rstselbob.EOF
      If rstsel!ample >= cadbl(mtrsnecessaris.Tag) Then
       actualitzar_metres_disponibles rstselbob!idpalet, rstselbob!idbobina
       possar_bobina_reixa rstsel, rstselbob, "0"
      End If
      rstselbob.MoveNext
    Wend
    rstsel.MoveNext
    'hauria de seleccionar tots els palets que tinguin la familia i subfamilia escullida
    'triar totes les bobines que estan lliures en aquests palets
    'buscar de cada palet el proveidor i altres taules necessaries per emplenar els datos
    'ianar afegint cada bobina a la taula d'assigaciomaterial
  Wend
  
  'posso 10 registres per provar
  ''For i = 1 To 10
  '' rstconsulta.AddNew
  '' rstconsulta.Update
  ''Next i
  'fins aqui
  
  Set rstconsulta = dbllistat.OpenRecordset("select * from assignaciomaterial order by ample asc,datarec Asc")
End Sub
Function crear_criteri_familia(Optional vcont As Byte) As String
   Dim d As String
   Dim rst As Recordset
   Dim vsql As String
   Dim vfamcolor As Double
   
   If Framecompatibles.Visible = False Then
        If fammat.ListIndex >= 0 Then
           d = " familia=" + atrim(fammat.ItemData(fammat.ListIndex))
         Else: d = " familia>0"
        End If
        If subfammat.ListIndex >= 0 Then vcont = vcont + 1: d = d + " and subfamilia=" + atrim(subfammat.ItemData(subfammat.ListIndex))
        If famcol.ListIndex >= 0 Then vcont = vcont + 1: d = d + " and familiacol=" + atrim(famcol.ItemData(famcol.ListIndex))
        If subfamcol.ListIndex >= 0 Then vcont = vcont + 1: d = d + " and subfamiliacol=" + atrim(subfamcol.ItemData(subfamcol.ListIndex))
        If famad.ListIndex >= 0 Then vcont = vcont + 1: d = d + " and familiaad=" + atrim(famad.ItemData(famad.ListIndex))
        If subfamad.ListIndex >= 0 Then vcont = vcont + 1: d = d + " and subfamiliaad=" + atrim(subfamad.ItemData(subfamad.ListIndex))
        If d = "" Then d = " familia=0 "
        crear_criteri_familia = siesmaterialexacte(cadbl(comanda), True)
        
        If crear_criteri_familia = "" Then
            crear_criteri_familia = d
             Else:
               If cadbl(etsubfamcompatible.Tag) > 0 Then
                 d = substituir(d, "subfamilia=" + atrim(subfammat.ItemData(subfammat.ListIndex)), "subfamilia=" + etsubfamcompatible.Tag)
                 crear_criteri_familia = "(" + crear_criteri_familia + " or (" + d + "))"
               End If
        End If
         Else
           If Combocompatibles.ListIndex > -1 Then
            Set rst = dbtmp.OpenRecordset("select * from grupsmaterialscompatibles where numerodegrup=" + atrim(Combocompatibles.ItemData(Combocompatibles.ListIndex)))
            If Not rst.EOF Then
                vsql = rst!sqlprincipal + rst!sqlsubfamilies + ")"
                If InStr(1, vsql, "familiacol=") > 0 Then vfamcolor = cadbl(Mid(vsql, InStr(1, vsql, "familiacol=") + 11, 4))
                If cadbl(famcol.ItemData(famcol.ListIndex)) <> vfamcolor Then MsgBox "EL COLOR DE MATERIAL DE LA FAMILIA DE COMPATIBLES ES DIFERENT QUE LA DE LA COMANDA." + vbNewLine + "REVISA QUE TOT SIGUI CORRECTE ABANS D'ASSINAR EL MATERIAL.", vbCritical, "ERROR": crear_criteri_familia = "sortir": GoTo fi
            End If
            crear_criteri_familia = vsql
            Set rst = Nothing
           End If
   End If
fi:
   
End Function
Function siesmaterialexacte(numc As Double, Optional noensenyarmissatge As Boolean) As String
   Dim rstc As Recordset
   siesmaterialexacte = ""
   Set rstc = dbtmp.OpenRecordset("SELECT comandes_extres.materialexacte, comandes.materialex FROM comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda Where comandes.comanda = " + atrim(numc))
   If Not rstc.EOF Then
      If cabool(rstc!materialexacte) Then
        siesmaterialexacte = " codimatprognou=" + atrim(rstc!materialex)
        If Not noensenyarmissatge Then
           MsgBox "Aquesta comanda te un material concret assignat, només ensenyaré les bobines amb aquest material." + Chr(10) + "EL MATERIAL ESPECIFIC NO ES POT RESERVAR", vbCritical, "Atenció"
'           If cadbl(InputBox("Si vols treure aquesta reserva igualment repeteix el numero de comanda.", "Des-reservar especific")) = numc Then
'                siesmaterialexacte = ""
'                  Else: MsgBox "No coincideix amb la comanda, no es farà cap canvi", vbCritical, "Atenció"
'           End If
        End If
      End If
   End If
   
End Function

Function altrescriteris() As String
  Dim criteri As String
  Dim desc As String
  desc = infodescripciomat
  If atrim(itl) <> "" Then criteri = criteri + " and palets.semielaborat='" + aatrim(itl) + "' "
  If InStr(1, "PEAD", desc) > 0 Or InStr(1, "PEMD", desc) > 0 Or InStr(1, "PEBD", desc) > 0 Then
    If atrim(icares) <> "" Then criteri = criteri + " and carestractat='" + aatrim(icares) + "' "
  End If
  If atrim(iobert) <> "" Then criteri = criteri + " and obert='" + aatrim(iobert) + "' "
  If atrim(iplegat) <> "" And cadbl(iplegat) > 0 Then criteri = criteri + " and plegat=" + atrim(passaradecimalpunt(cadbl(iplegat)))
  If atrim(isolapa) <> "" And cadbl(isolapa) > 0 Then criteri = criteri + " and solapa=" + atrim(passaradecimalpunt(cadbl(isolapa)))
  criteri = criteri + " and " + IIf(cabool(imicrop), "", "not") + " microperforat"
altrescriteris = criteri
End Function
Function hihaassignacio() As Boolean
  Dim rstparcial As Recordset
  If cadbl(comanda) > 0 Then
   Set rstparcial = dbtmp.OpenRecordset("select * from parcials where comanda='" + atrim(cadbl(comanda)) + "'")
   If rstparcial.EOF Then
     hihaassignacio = False
    Else: hihaassignacio = True
   End If
    Else: hihaassignacio = False
  End If
  Set rstpacial = Nothing
End Function

Sub carregar_rstconsulta()
  Dim criteribusqbob As String
  Dim rstsel As Recordset
  Dim rstsel2 As Recordset
  Dim rstparcial As Recordset
  Dim rstselbob As Recordset
  Dim numfam As String
  Dim numsubfam As String
  Dim ample As Double
  Dim criterifamilia As String
  Dim nomespesor As String
  Dim consultamicres As String
  Dim criteri As String
  Dim sel As String
  Dim vpassada As Byte
  Dim vultimabob As String
  Dim vcont As Integer
  
  vpassada = 1
  ample = cadbl(iample)
  If ample > 1 Then ample = ample - 1
  If limitarample.Value = 0 Then ample = 0
 ' If Reserves.Caption = "Reserves" Then
 '  If cadbl(comanda) > 0 Then
 '   Set rstparcial = dbtmp.OpenRecordset("SELECT Min(Palets.Ample) AS minample, Parcials.comanda FROM Parcials INNER JOIN Palets ON Parcials.idpalet = Palets.Idpalet GROUP BY Parcials.comanda HAVING (((Parcials.comanda)='" + comanda + "'));")
 '   If Not rstparcial.EOF Then ample = rstparcial!minample
 '   Set rstparcial = Nothing
 '  End If
 ' End If
  
  numsubfam = ""
  numfam = "0"
  dbllistat.Execute ("delete * from assignaciomaterial")
  Set rstconsulta = dbllistat.OpenRecordset("assignaciomaterial")
  criterifamilia = crear_criteri_familia
  If criterifamilia = "sortir" Then GoTo fi
  If criterifamilia = "" Then MsgBox "No hi ha cap seleccio de families per filtrar.", vbCritical, "Error": GoTo fi
  If fammat.ListIndex >= 0 Then numfam = Trim(fammat.ItemData(fammat.ListIndex))
  If subfammat.ListIndex >= 0 Then numsubfam = " and subfamilia=" + Trim(cadbl(subfammat.ItemData(subfammat.ListIndex)))
  If infodescripciomat.Tag <> "" Then numsubfam = aondelmaterial(cadbl(infodescripciomat.Tag))
  'treure aquesta linia proxima per filtra per espesors obert micro t o b etc...
  'imicrop.Tag = ""
  If atrim(iespesor) <> "" Then
     consultamicres = " and  micres=" + atrim(cadbl(iespesor))
    Else: consultamicres = " and  micres>1"
  End If
      'quan era grm/m2 abans estava amb grmsm2>= però l'alicia ara diu que només ha de ser igual.  18/03/24
  If cadbl(iespesor) < 0 Or etmicres = "Grm/m2" Then consultamicres = " and  grmsm2=" + atrim(cadbl(iespesor) * IIf(cadbl(iespesor) < 0, -1, 1))
  
  'criteridebusqueda = criterifamilia + IIf(criterifamilia <> "", " and", "") + " ample >= " + adec(atrim(ample)) + consultamicres + imicrop.Tag
  criteridebusqueda = criterifamilia + altrescriteris + passaradecimalpunt(consultamicres) + " and (ample>=" + passaradecimalpunt(atrim(ample)) + " AND ample<=" + (passaradecimalpunt(atrim(ample) + cadbl(cmargeamplada))) + ") "
  mtrsnecessaris.Tag = ample
  criteri = "SELECT Palets.*,materials.familia, materials.subfamilia,materials.proveidor  FROM Palets INNER JOIN materials ON Palets.codimatprognou = materials.codi WHERE " + criteridebusqueda + " and palets.ample>=" + atrim(passaradecimalpunt(mtrsnecessaris.Tag)) + " and proveidor" + IIf(matprovproves = 1, "=", "<>") + "581 "
  sel = "0"
  'MsgBox criteridebusqueda
saltrebuscar:
  'Set rstsel = dbtmp.OpenRecordset("SELECT Palets.*,materials.familia, materials.subfamilia FROM Palets INNER JOIN materials ON Palets.codimatprognou = materials.codi WHERE " + criteri + ";")
  'MsgBox criteri
  etprogres = "Filtrant ... 1 ": DoEvents
 ' Clipboard.Clear
 ' Clipboard.SetText criteri
  
  Set rstsel = dbtmp.OpenRecordset(criteri, dbOpenSnapshot, dbReadOnly)
  
  If Not rstsel.EOF Then
   rstsel.MoveLast
   rstsel.MoveFirst
  End If
  criteribusqbob = ""
  If Not hihaassignacio Then criteribusqbob = " disponible >0 and "
  'Clipboard.Clear
  'Clipboard.SetText "SELECT Bobines.*, Parcials.comanda FROM Bobines INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet) where parcials.comanda='" + atrim(comanda) + "' order by bobines.idpalet,bobines.idbobina "
  Set rstselbob = dbtmp.OpenRecordset("SELECT Bobines.*, Parcials.comanda FROM Bobines INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet) where parcials.comanda='" + atrim(comanda) + "' order by bobines.idpalet,bobines.idbobina ", dbOpenSnapshot, dbReadOnly)  'or disponible>0 order by disponible,bobines.idpalet,idbobina ", dbOpenSnapshot, dbReadOnly)
  
 ' Set rstselbob = dbtmp.OpenRecordset("select * from bobines where disponible>0 order by disponible,idpalet,idbobina ", dbOpenSnapshot, dbReadOnly)
passarelsregistres:
  
  While Not rstsel.EOF                                               'numcom='0'  and
   ' etprogres = "Afegint ...  " + atrim(rstsel.AbsolutePosition) + "/" + atrim(rstsel.RecordCount): DoEvents                                   'numcom='0'  and
    rstselbob.FindFirst "idpalet=" + atrim(cadbl(rstsel!idpalet))
    While Not rstselbob.NoMatch
          If vultimabob <> atrim(rstselbob!idpalet) + " " + atrim(rstselbob!idbobina) Then
                vultimabob = atrim(rstselbob!idpalet) + " " + atrim(rstselbob!idbobina)
                If parafiltre.Tag = "1" Then GoTo fi
                If rstsel!ample >= cadbl(mtrsnecessaris.Tag) Then
                 'actualitzar_metres_disponibles rstselbob!idpalet, rstselbob!idbobina
                 If sel <> "3" Then
                      possar_bobina_reixa rstsel, rstselbob, "0"
                    Else
                      If nohies(rstsel!idpalet, rstselbob!idbobina) Then
                        possar_bobina_reixa rstsel, rstselbob, "0"
                      End If
                 End If
                End If
          End If
          rstselbob.FindNext "idpalet=" + atrim(cadbl(rstsel!idpalet))
          
          If vpassada <> 1 And vcont > 20 Then barraprogres rstsel.AbsolutePosition, rstsel.RecordCount: vcont = 0: DoEvents
    Wend
    vcont = vcont + 1
    rstsel.MoveNext
    'hauria de seleccionar tots els palets que tinguin la familia i subfamilia escullida
    'triar totes les bobines que estan lliures en aquests palets
    'buscar de cada palet el proveidor i altres taules necessaries per emplenar els datos
    'ianar afegint cada bobina a la taula d'assigaciomaterial
  Wend
  If vpassada = 1 And Not (rstsel.EOF And rstsel.BOF) Then
     Set rstselbob = dbtmp.OpenRecordset("select * from bobines where  disponible>0 order by disponible,idpalet,idbobina ", dbOpenSnapshot, dbReadOnly)
     rstsel.MoveFirst
     vpassada = 2: GoTo passarelsregistres
  End If
  
  criteri = "SELECT Palets.*, materials.familia, materials.subfamilia as numc FROM (Palets INNER JOIN materials ON Palets.codimatprognou = materials.codi) INNER JOIN Parcials ON Palets.Idpalet = Parcials.idpalet where " + " parcials.comanda='" + atrim(cadbl(comanda)) + "' " + ";"
  If cadbl(comanda) = 0 Then sel = "3"
  If sel <> "3" Then
    sel = "3"
    GoTo saltrebuscar
  End If
  
  'posso 10 registres per provar
  ''For i = 1 To 10
  '' rstconsulta.AddNew
  '' rstconsulta.Update
  ''Next i
  'fins aqui
   etprogres = "Fi afegint ... ": DoEvents
  If ordenatperpalet <> 1 Then
    Set rstconsulta = dbllistat.OpenRecordset("select * from assignaciomaterial order by ample asc,datarec asc,palet Asc,bobina asc")
   ' rstconsulta.MoveLast
   ' rstconsulta.MoveFirst
   ' MsgBox atrim(rstconsulta.RecordCount)
      Else: Set rstconsulta = dbllistat.OpenRecordset("select * from assignaciomaterial order by palet asc,bobina asc,datarec asc")
  End If
fi:
  
End Sub
Sub carregar_rstconsulta_rst()
  Dim criteribusqbob As String
  Dim rstsel As Recordset
  Dim rstsel2 As Recordset
  Dim rstparcial As Recordset
  Dim rstselbob As Recordset
  Dim numfam As String
  Dim numsubfam As String
  Dim ample As Double
  Dim criterifamilia As String
  Dim nomespesor As String
  Dim consultamicres As String
  Dim criteri As String
  Dim sel As String
  ample = cadbl(iample)
  If ample > 1 Then ample = ample - 1
  If limitarample.Value = 0 Then ample = 0
 ' If Reserves.Caption = "Reserves" Then
 '  If cadbl(comanda) > 0 Then
 '   Set rstparcial = dbtmp.OpenRecordset("SELECT Min(Palets.Ample) AS minample, Parcials.comanda FROM Parcials INNER JOIN Palets ON Parcials.idpalet = Palets.Idpalet GROUP BY Parcials.comanda HAVING (((Parcials.comanda)='" + comanda + "'));")
 '   If Not rstparcial.EOF Then ample = rstparcial!minample
 '   Set rstparcial = Nothing
 '  End If
 ' End If
  
  numsubfam = ""
  numfam = "0"
  criterifamilia = crear_criteri_familia
  If crear_criteri_familia = "sortir" Then GoTo fi
  If fammat.ListIndex >= 0 Then numfam = Trim(fammat.ItemData(fammat.ListIndex))
  If subfammat.ListIndex >= 0 Then numsubfam = " and subfamilia=" + Trim(cadbl(subfammat.ItemData(subfammat.ListIndex)))
  If infodescripciomat.Tag <> "" Then numsubfam = aondelmaterial(cadbl(infodescripciomat.Tag))
  'treure aquesta linia proxima per filtra per espesors obert micro t o b etc...
  'imicrop.Tag = ""
  dbllistat.Execute ("delete * from assignaciomaterial")
  Set rstconsulta = dbllistat.OpenRecordset("assignaciomaterial")
  If atrim(iespesor) <> "" Then
     consultamicres = " and  micres=" + atrim(cadbl(iespesor))
    Else: consultamicres = " and  micres>1"
  End If
  If matprovproves = 1 Then consultamicres = ""
  If cadbl(iespesor) < 0 Or etmicres = "Grm/m2" Then consultamicres = " and  grmsm2>=" + atrim(cadbl(iespesor) * IIf(cadbl(iespesor) < 0, -1, 1))
  
  'criteridebusqueda = criterifamilia + IIf(criterifamilia <> "", " and", "") + " ample >= " + adec(atrim(ample)) + consultamicres + imicrop.Tag
  criteridebusqueda = criterifamilia + altrescriteris + passaradecimalpunt(consultamicres) + " and ample>=" + passaradecimalpunt(atrim(ample))
  mtrsnecessaris.Tag = ample
 ' MsgBox criteridebusqueda
  criteri = "SELECT Palets.*,materials.familia, materials.subfamilia,materials.proveidor  FROM Palets INNER JOIN materials ON Palets.codimatprognou = materials.codi WHERE " + criteridebusqueda + " and palets.ample>=" + atrim(passaradecimalpunt(mtrsnecessaris.Tag)) + " and proveidor" + IIf(matprovproves = 1, "=", "<>") + "581 "
'  Clipboard.Clear
'  Clipboard.SetText criteri
  sel = "0"
  'MsgBox criteridebusqueda
saltrebuscar:
  'Set rstsel = dbtmp.OpenRecordset("SELECT Palets.*,materials.familia, materials.subfamilia FROM Palets INNER JOIN materials ON Palets.codimatprognou = materials.codi WHERE " + criteri + ";")
 'MsgBox criteri
  etprogres = "Filtrant ... 1 ": DoEvents
  
  Set rstsel = dbtmp.OpenRecordset(criteri, dbOpenSnapshot, dbReadOnly)
  
  If Not rstsel.EOF Then
   rstsel.MoveLast
   rstsel.MoveFirst
  End If
  criteribusqbob = ""
  If Not hihaassignacio Then criteribusqbob = " disponible >0 and "
  'Set rstselbob = dbtmp.OpenRecordset("select * from bobines where disponible>0 order by disponible,idpalet,idbobina ", dbOpenSnapshot, dbReadOnly)
  
  While Not rstsel.EOF
    
    If sel <> "3" Then
        Set rstselbob = dbtmp.OpenRecordset("select * from bobines where  " + criteribusqbob + " idpalet=" + atrim(cadbl(rstsel!idpalet)), dbOpenSnapshot, dbReadOnly)
         Else: Set rstselbob = dbtmp.OpenRecordset("SELECT Bobines.*, Parcials.comanda FROM Bobines INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet) where parcials.comanda='" + atrim(comanda) + "' order by bobines.idpalet,bobines.idbobina ", dbOpenSnapshot, dbReadOnly)     'or disponible>0 order by disponible,bobines.idpalet,idbobina ", dbOpenSnapshot, dbReadOnly)
    End If
    While Not rstselbob.EOF
      If parafiltre.Tag = "1" Then GoTo fi
      If rstsel!ample >= cadbl(mtrsnecessaris.Tag) Then
       actualitzar_metres_disponibles rstselbob!idpalet, rstselbob!idbobina
       If sel <> "3" Then
            possar_bobina_reixa rstsel, rstselbob, "0"
          Else
            If nohies(rstsel!idpalet, rstselbob!idbobina) Then
              possar_bobina_reixa rstsel, rstselbob, "0"
            End If
       End If
      End If
      'rstselbob.FindNext "idpalet=" + atrim(cadbl(rstsel!idpalet))
      rstselbob.MoveNext
      barraprogres rstsel.AbsolutePosition, rstsel.RecordCount
    Wend
    rstsel.MoveNext
    'hauria de seleccionar tots els palets que tinguin la familia i subfamilia escullida
    'triar totes les bobines que estan lliures en aquests palets
    'buscar de cada palet el proveidor i altres taules necessaries per emplenar els datos
    'ianar afegint cada bobina a la taula d'assigaciomaterial
  Wend
  criteri = "SELECT Palets.*, materials.familia, materials.subfamilia as numc FROM (Palets INNER JOIN materials ON Palets.codimatprognou = materials.codi) INNER JOIN Parcials ON Palets.Idpalet = Parcials.idpalet where " + " parcials.comanda='" + atrim(cadbl(comanda)) + "' " + ";"
 
  If cadbl(comanda) = 0 Then sel = "3"
  If sel <> "3" Then
    sel = "3"
    GoTo saltrebuscar
  End If
  
  'posso 10 registres per provar
  ''For i = 1 To 10
  '' rstconsulta.AddNew
  '' rstconsulta.Update
  ''Next i
  'fins aqui
   etprogres = "Fi afegint ... ": DoEvents
  If ordenatperpalet <> 1 Then
    Set rstconsulta = dbllistat.OpenRecordset("select * from assignaciomaterial order by ample asc,datarec asc,palet Asc,bobina asc")
    If Not rstconsulta.EOF And Not rstconsulta.BOF Then
     rstconsulta.MoveLast
     rstconsulta.MoveFirst
    End If
   ' MsgBox atrim(rstconsulta.RecordCount)
      Else: Set rstconsulta = dbllistat.OpenRecordset("select * from assignaciomaterial order by palet asc,bobina asc,datarec asc")
  End If
fi:
  
End Sub

Sub barraprogres(actual As Double, gran As Double)
  Dim factor As Double
  If actual = gran Then
     etprogres = "Actualitzant reixa"
    Else: etprogres = ""
  End If
  factor = (actual * 100) / gran
  liniaprogres.Width = factor * (3000 / 100)
  DoEvents
End Sub
Function nohies(palet As Double, bobina As Double)
  Dim rstc As Recordset
  Set rstc = dbllistat.OpenRecordset("select palet from assignaciomaterial where palet=" + atrim(palet) + " and bobina=" + atrim(bobina))
  If Not rstc.EOF Then
     nohies = False
       Else: nohies = True
  End If
  Set rstc = Nothing
End Function
Function adec(v As String) As String
   v = atrim(cadbl(v))
   If InStr(1, v, ",") Then v = Mid(v, 1, InStr(1, v, ",") - 1) + "." + Mid(v, InStr(1, v, ",") + 1)
   adec = v
End Function
Function aondelmaterial(codimat As Long) As String
  Dim d As String
  Dim f As String
  Set rsttmp = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(codimat))
  d = ""
  
  If Not rsttmp Then
      If cadbl(rsttmp!subfamilia) > 0 Then d = d + " and subfamilia=" + atrim(cadbl(rsttmp!subfamilia))
      If cadbl(rsttmp!familiacol) > 0 Then d = d + " and familiacol=" + atrim(cadbl(rsttmp!familiacol))
      If cadbl(rsttmp!subfamiliacol) > 0 Then d = d + " and subfamiliacol=" + atrim(cadbl(rsttmp!subfamiliacol))
      If cadbl(rsttmp!familiaad) > 0 Then d = d + " and familiaad=" + atrim(cadbl(rsttmp!familiaad))
      If cadbl(rsttmp!subfamiliaad) > 0 Then d = d + " and subfamiliaad=" + atrim(cadbl(rsttmp!subfamiliaad))
      For i = 1 To 1000
       f = f + "prova"
      Next i
      '
      '
      '
      '
  End If
  aondelmaterial = d
  Set rsttmp = Nothing
End Function
Sub configurar_reixa()
  Dim col As Integer
  Dim i As Integer
  Dim espe As Double
  'reixa.Clear
  espe = IIf(atrim(iespesor) = "", 1, cadbl(iespesor))
  col = 0
  reixa.Rows = 2
  
  reixa.Cols = rstconsulta.Fields.Count
  reixa.FixedRows = 1
  reixa.FixedCols = 0
  For i = 0 To rstconsulta.Fields.Count - 1
     reixa.col = col
     reixa.ColData(i) = i
     reixa.TextMatrix(0, col) = UCase(rstconsulta.Fields(i).Name)
     If UCase(rstconsulta.Fields(i).Name) = "MICRES" And espe <= 0 Then
        reixa.TextMatrix(0, col) = "GRMSM2"
     End If
     reixa.ColWidth(col) = IIf(rstconsulta.Fields(i).Size > 25, 25, rstconsulta.Fields(i).Size) * 200
     col = col + 1
  Next i
  carregar_amples_reixa
End Sub

Sub refrescar_reixa()
 agrupar_reserves
 configurar_reixa
 poblar_reixa
End Sub

Sub comprovarmaterialexactependentdassignar()
  Dim rst As Recordset
  Dim rstp As Recordset
  Dim vsql As String
  Dim vmsg As String
  Dim r As String
  ratoli "espera"
  etinformacio.Top = 120
  etinformacio.Left = 195
  etinformacio.Text = "Comprovant compres de material especific."
  etinformacio.Visible = True
  DoEvents
  Me.Caption = "Comprovant compres de material especific."
  'vsql = "SELECT comandesmesextres.comanda, comandesmesextres.materialexacte, comandesmesextres.proximaseccio, comandesxlinia.idliniacompra, liniescompra.kgentregats "
  'vsql = vsql + " FROM liniescompra RIGHT JOIN (comandesmesextres INNER JOIN comandesxlinia ON comandesmesextres.comanda = comandesxlinia.numcomanda) ON liniescompra.idliniacompra = comandesxlinia.idliniacompra "
  'vsql = vsql + " WHERE (((comandesmesextres.materialexacte)=True) AND ((comandesmesextres.proximaseccio)<>'T') AND ((liniescompra.kgentregats)>0));"

  Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda, comandes.proximaseccio, comandes_extres.materialexacte FROM comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda WHERE (((comandes.proximaseccio)<>'T') AND ((comandes_extres.materialexacte)=True));")
  While Not rst.EOF
     Set rstp = dbcomandes.OpenRecordset("select * from comandesxlinia where numcomanda=" + atrim(rst!comanda))
     If Not rstp.EOF Then
      Set rstp = dbcomandes.OpenRecordset("select * from liniescompra where idliniacompra=" + atrim(rstp!idliniacompra))
      If Not rstp.EOF Then
        If cadbl(rstp!kgentregats) > 0 Then
           Set rstp = dbtmp.OpenRecordset("select comanda from parcials where comanda='" + atrim(rst!comanda) + "'")
           If rstp.EOF Then vmsg = vmsg + " - " + atrim(rst!comanda)
        End If
      End If
     End If
     rst.MoveNext
  Wend
  ratoli "normal"
  Set rst = Nothing
  Set rstp = Nothing
  If vmsg <> "" Then
     r = ""
     While r <> "D'ACORD"
        r = UCase(InputBox("Les comandes " + vmsg + " tenen material especific, ja ha arribat la compra i no està assignat." + Chr(10) + "ESCRIU [D'ACORD] PER CONTINUAR"))
     Wend
   End If
   Me.Caption = "Assignar material a la comanda."
   etinformacio.Text = ""
   etinformacio.Visible = False
   comanda.SetFocus
End Sub



Private Sub Form_Activate()
  
  
   assignardecimalipunt
   If vfiltrebobinesdesdeimpresores And cadbl(comanda) = 0 Then
    reixa.Top = 1
    reixa.Left = 10
    assignarmat.Height = reixa.Height + 600
    assignarmat.Width = reixa.Width + 100
    reixa.ZOrder 0
    missatgepantalla.ZOrder 0
    comanda = cadbl(llegir_ini("Baixes", "numcomanda", "comandes.ini"))
    
    'wait 1
    'filtrar_Click
     Else: reixa.ZOrder 1
   End If
   
   
End Sub
Sub carregar_combo_compatibles(Optional vcodiescullit As Double)
  Dim rst As Recordset
  Dim i As Integer
  
  Combocompatibles = Clear
  Set rst = dbtmp.OpenRecordset("select * from grupsmaterialscompatibles order by nomdelgrup")
  While Not rst.EOF
      Combocompatibles.AddItem rst!nomdelgrup
      Combocompatibles.ItemData(Combocompatibles.NewIndex) = rst!numerodegrup
      If vcodiescullit = cadbl(rst!numerodegrup) Then i = Combocompatibles.NewIndex
      rst.MoveNext
  Wend
  If i > -1 And i < Combocompatibles.ListCount Then Combocompatibles.ListIndex = i
  Set rst = Nothing
End Sub
Function hiharelacionsunacomanda(numc As Double) As String
  Dim dbt As Database
  Dim rstt As Recordset

  Set dbt = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
  
  Set rstt = dbt.OpenRecordset("select * from comandesxlinia where numcomanda=" + atrim(numc))
  If Not rstt.EOF Then hiharelacionsunacomanda = "-COMPRES"
  
  Set dbt = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
 
  Set rstt = dbt.OpenRecordset("select * from parcials where comanda='" + atrim(numc) + "'")
  If Not rstt.EOF Then hiharelacionsunacomanda = hiharelacionsunacomanda + "-ASSIGNACIO"
  
  Set rstt = dbt.OpenRecordset("select * from percomandaoclient where numcomanda=" + atrim(numc))
    
  If Not rstt.EOF Then hiharelacionsunacomanda = hiharelacionsunacomanda + "-RESERVA"
  
fi:
  Set dbt = Nothing
  Set rstt = Nothing

End Function
Function ensenyar_relacions(numc As Double) As String
   Dim rst As Recordset
   Dim msg As String
   If numc = 0 Then GoTo fi
   Set rst = dbcomandes.OpenRecordset("select comanda,linkcomanda1,linkcomanda2,dataactivacio from comandes where comanda=" + atrim(numc))
   If Not rst.EOF Then
       If Not IsNull(rst!dataactivacio) Then GoTo fi
       If cadbl(rst!comanda) > 0 Then msg = atrim(rst!comanda) + hiharelacionsunacomanda(cadbl(rst!comanda))
       If cadbl(rst!linkcomanda1) > 0 Then msg = msg + Chr(10) + atrim(rst!linkcomanda1) + hiharelacionsunacomanda(cadbl(rst!linkcomanda1))
       If cadbl(rst!linkcomanda2) > 0 Then msg = msg + Chr(10) + atrim(rst!linkcomanda2) + hiharelacionsunacomanda(cadbl(rst!linkcomanda2))
        Else: GoTo fi
   End If
   ensenyar_relacions = msg
fi:
   Set rst = Nothing
End Function
Sub comprovar_siencarahiharelacions()
   Dim msg As String
   Dim numc As Double
   numc = -1
   While Not formseleccio.Data1.Recordset.EOF
      numc = cadbl(formseleccio.Data1.Recordset!comandaoreferencia)
      msg = ensenyar_relacions(numc)
      If InStr(1, msg, "-") = 0 Then
           dbcomandes.Execute "update informaciodesactivades set actiu=false where comandaoreferencia='" + atrim(numc) + "'"
      End If
      formseleccio.Data1.Recordset.MoveNext
   Wend
   If numc <> -1 Then formseleccio.Data1.Recordset.MoveFirst
End Sub
Sub comprovarcomandesdesactivades()
   Dim msg As String
   Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select comandaoreferencia,nomclient,descripcio from informaciodesactivades WHERE tipus='P' and actiu order by data"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   comprovar_siencarahiharelacions
   formseleccio.refrescar
   If formseleccio.Data1.Recordset.EOF Then GoTo fi
   'formseleccio.Width = formseleccio.Width + (formseleccio.Width / 3)
   formseleccio.DBGrid2.Columns("comandaoreferencia").Width = 1200
   formseleccio.DBGrid2.Columns("nomclient").Width = 2500
   formseleccio.DBGrid2.Columns("descripcio").Width = 5000
   formseleccio.Width = 10500
   formseleccio.Caption = "Comandes desactivades"
   formseleccio.Left = (Screen.Width / 2) - (formseleccio.Width / 2)
   formseleccio.Command2.Tag = "1"
   formseleccio.Show 1
   If seleccioret = 1 Then
      msg = ensenyar_relacions(cadbl(formseleccio.Data1.Recordset!comandaoreferencia))
      If InStr(1, msg, "-") > 0 Then
       MsgBox msg
         Else
           dbcomandes.Execute "update informaciodesactivades set actiu=false where comandaoreferencia='" + atrim(numc) + "'"
      End If
   End If
   
fi:
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
    
    
End Sub
Sub carregarperreservar(numc As Double)
    Dim metres As Double
    passar_a_reservar
    Set dbcompres = DBEngine.OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
    comanda = atrim(numc)
    Set rst = dbtmp.OpenRecordset("select * from pendentsdereservar where not reservar")
    If rst.EOF Then Exit Sub
    Command2.Caption = "Ok Reservat Tot"
    comandesperreservar = "comandes: "
    If rst.EOF Then comandesperreservar = comandesperreservar + atrim(rst!comanda)
    While Not rst.EOF
      comandesperreservar = comandesperreservar + ", " + atrim(rst!comanda)
      metres = metres + atrim(cadbl(rst!metres))
      rst.MoveNext
    Wend
    comandesperreservar = atrim(metres) + " Mtrs-" + comandesperreservar
    
End Sub
Private Sub Form_Click()
   
'MsgBox Combocompatibles.ItemData(Combocompatibles.ListIndex)
'If materialsdirefentsdeBACAICOA(cadbl(comanda)) Then Exit Sub
  'comprovarsihihaunacomandasemblantafabricaiavisar 203779
  'comprovarsihihaamuntadoraeltreballiavisar 195737
  ' des_reservar 175456
  'comprovarsihihaunacomandasemblantafabricaiavisar 182927
 ' For i = 1 To 20
 '   imprimir_packinglist 150275, llistat
 ' Next i
 'r = "nopregunta"
 'Set rsttmp = dbtmp.OpenRecordset("SELECT Trim([numcomanda]) AS Expr1, percomandaoclient.idcompra, percomandaoclient.metres From percomandaoclient WHERE (((Trim([numcomanda])) In (select comanda from parcials)));")
 'While Not rsttmp.EOF
 '   des_reservar rsttmp!expr1
 '   rsttmp.MoveNext
 'Wend
'baixaseccioextrusora comanda
 
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  iniconfigreixa = "reixaassignarmat.ini"
  On Error Resume Next
  If Not existeix("c:\windows\" + iniconfigreixa) Then FileCopy "\\serverprodu\dades\progcomandes\aplicacio\" + iniconfigreixa, "c:\windows\" + iniconfigreixa
  reixa.Rows = 1
  carregar_combo_families
  passar_a_assignar
  etinformacio.Top = 90
  etinformacio.Left = 210
  Framecompatibles.Left = 3090
  Framecompatibles.Top = 120
  carregar_combo_compatibles
'  comanda_Change
  'comprovarcomandesdesactivades
End Sub

Function sumar_seleccionats() As Double
  Dim total As Double
  Dim metres As Byte
  Dim seleccionat As Byte
  Dim sel As String
  For i = 0 To reixa.Rows - 1
    sel = reixa.TextMatrix(i, columnadelcamp("seleccionat"))
    If sel = "1" Or sel = "2" Then
       total = total + cadbl(reixa.TextMatrix(i, columnadelcamp("mtrsassignats")))
    End If
  Next i
  sumar_seleccionats = total
End Function

Private Sub Form_Unload(Cancel As Integer)
guardar_amples_reixa
If Form1.Visible = False Then End
descarregarvariables
Unload Form1
End Sub
Sub descarregarvariables()
  Set dbtmp = Nothing
  Set dbtmpb = Nothing
  Set dbstocks = Nothing
  Set rsttmp = Nothing
  Set dbllistat = Nothing
  Set rstllistat = Nothing
  
End Sub
Sub guardar_amples_reixa()
If iniconfigreixa <> "" Then
  For j = 0 To reixa.Cols - 1
   escriure_ini "AmplesReixa", UCase(reixa.TextMatrix(0, j)), atrim(reixa.ColWidth(j)), iniconfigreixa
 Next j
End If
End Sub
Sub carregar_amples_reixa()
 Dim ample As String
 
 If iniconfigreixa <> "" Then ' existeix("c:\windows\" + iniconfigreixa) Then
  For j = 0 To reixa.Cols - 1
   ample = llegir_ini("AmplesReixa", UCase(reixa.TextMatrix(0, j)), iniconfigreixa)
   
   If ample <> "{[}]" Then reixa.ColWidth(j) = cadbl(ample)
   r = llegir_ini("NomsReixa", UCase(rstconsulta.Fields(reixa.ColData(j)).Name) + "-nom", iniconfigreixa)
   If r <> "{[}]" Then
      reixa.TextMatrix(0, j) = r
   End If
 Next j
End If
End Sub

Private Sub materialcomanda_Click()

End Sub

Private Sub infoampleplegat_Click()

End Sub

Private Sub Frame1_Click()
   'comprovarsihihaunacomandasemblantafabricaiavisar 183209
End Sub

Private Sub iample_DblClick()
iample = cadbl(InputBox("Entra l'amplada del material", "Canvi d'amplada"))
  filtrar_materials
End Sub

Private Sub iespesor_DblClick()
iespesor = cadbl(InputBox("Entra les micres del material", "Canvi d'espesor"))
  filtrar_materials
End Sub

Private Sub imicrop_Click()
  If Screen.ActiveControl.Name = "imicrop" Then filtrar_materials

End Sub

Private Sub iplegat_DblClick()
iplegat = cadbl(InputBox("Entra el plegat del material", "Canvi de plegat"))
  filtrar_materials
End Sub

Private Sub isolapa_DblClick()
isolapa = cadbl(InputBox("Entra la solapa del material", "Canvi de solapa"))
  filtrar_materials
End Sub

Private Sub modificar_Click()
  If Framecompatibles.Visible = False Then
    If siesmaterialexacte(cadbl(comanda), True) <> "" Then
        MsgBox "Aquesta comanda te material exacte assignat no es pot fer amb un material compatible", vbCritical, "Error": Exit Sub
         Else: GoTo cont
    End If
      Else: Combocompatibles.Text = "": Combocompatibles.ListIndex = -1
  End If
cont:
  Framecompatibles.Visible = Not Framecompatibles.Visible
End Sub

Private Sub mtrsnecessaris_Change()
  metresareservar = 0
End Sub

Private Sub mtrsnecessaris_GotFocus()
  mtrsnecessaris.SelStart = 0
  mtrsnecessaris.SelLength = Len(mtrsnecessaris)
End Sub

Private Sub re_reservar_Click()

End Sub

Sub ensenyar_comandes(idreserva As Double)
 '  reixalat.Visible = False
 Dim rst As Recordset
  Dim nomclient As String
   reixacomandes.Visible = False
   reixacomandes.FillStyle = flexFillRepeat
   datalat.DatabaseName = camistock
   datalat.RecordSource = "select * from percomandaoclient where  idreserva=" + atrim(idreserva) '(idcompra=null or idcompra<1) and
   datalat.Refresh
   reixacomandes.Clear
   reixacomandes.FormatString = "<id>|<Comanda|^Client |<Nom Client      |>Metres    "
   reixacomandes.Rows = 1
   With datalat.Recordset
   While Not datalat.Recordset.EOF
      nomclient = ""
      If cadbl(!numclient) > 0 Then
         Set rst = dbtmpb.OpenRecordset("select nom from clients where codi=" + atrim(!numclient))
         If Not rst.EOF Then nomclient = rst!nom
      End If
      reixacomandes.AddItem atrim(!idreserva) & Chr(9) & atrim(!numcomanda) & Chr(9) & atrim(!numclient) & Chr(9) & nomclient & Chr(9) & atrim(!metres)
      If (Not IsNull(!idcompra) And !idcompra > 0) Then
         possarcolorfila reixacomandes, reixacomandes.Rows - 1, &HC0C0FF
          Else
           If cadbl(idcompra) = 0 And cadbl(!numcomanda) = 0 Then
            possarcolorfila reixacomandes, reixacomandes.Rows - 1, QBColor(11)
           End If
      End If

     datalat.Recordset.MoveNext
   Wend
   End With
   reixacomandes.ColSel = 0
   datalat.Refresh
   If Not datalat.Recordset.EOF Then reixacomandes.Visible = True
'   reixalat.Refresh
'   If Not datalat.Recordset.EOF Then reixalat.Visible = True
End Sub
Sub possarcolorfila(reixac As MSFlexGrid, fila As Integer, color As Long)
         reixac.col = 0
         reixac.row = fila
         reixac.RowSel = fila
         reixac.ColSel = reixac.Cols - 1
         reixac.CellBackColor = color
End Sub
Sub escullir_client()
 Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select * from clients order by nom"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.Width = formseleccio.Width + (formseleccio.Width / 3)
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 9 Then
     codiclient = ""
     nomclient = ""
   End If
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           nomclient = formseleccio.DBGrid2.Columns("nom")
           codiclient = formseleccio.DBGrid2.Columns("codi")
        End If
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub nomclient_DropDown()
 escullir_client
   SendKeys "{TAB}"
End Sub

Private Sub parafiltre_Click()
  parafiltre.Tag = "1"
End Sub

Private Sub partirmanual_Click()
colocar_assignats
End Sub
Sub comprovarsiaquestabobinaescorrecteiavisar()
    Dim criteri As String
    Dim vcodimat As Double
    vcodimat = cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("codimat")))
    criteri = "SELECT Palets.*,materials.familia, materials.subfamilia,materials.proveidor  FROM Palets INNER JOIN materials ON Palets.codimatprognou = materials.codi WHERE " + criteridebusqueda + " and palets.codimatprognou=" + atrim(vcodimat) + " and palets.ample>=" + atrim(passaradecimalpunt(mtrsnecessaris.Tag)) + " and proveidor" + IIf(matprovproves = 1, "=", "<>") + "581 ;"
    Set rst = dbtmp.OpenRecordset(criteri)
    If rst.EOF Then
      MsgBox "Aquesta bobina escullida no es del material filtrat, REVISA QUE SIGUI CORRECTE", vbCritical + vbOKOnly, "A T E N C I Ó  !!!!!!"
    End If
    
End Sub
Private Sub reixa_Click()
Dim numreserva As Double
Dim idpalet As Double
If vfiltrebobinesdesdeimpresores Then Exit Sub
  metresareservar = 0
 If reixa.Rows < 2 Then Exit Sub
 If reixa.TextMatrix(reixa.row, columnadelcamp("ample")) = "" Then Exit Sub
 If Reserves.Caption <> "Reserves" Then
  ' If Not mirarsireservaperfamilies Then   'l'ALICIA M'HA DIT QUE HO TREGUI PER PROVAR 19/05/2022
    numreserva = cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("idreserva")))
    idpalet = cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("idpalet")))
    reserva_dades_capcalera numreserva, idpalet
    ensenyar_comandes numreserva
  'End If
 End If
 
  If rstconsulta.Fields(reixa.ColData(reixa.col)).Type = 1 And rstconsulta.Fields(reixa.ColData(reixa.col)).Name <> "resto" Then
   
    If reixa.CellPicture = check.Picture Then
       Set reixa.CellPicture = nocheck.Picture
       reixa.TextMatrix(reixa.row, reixa.col) = "0"
       reixa.TextMatrix(reixa.row, columnadelcamp("mtrsassignats")) = 0
       check.Tag = ""
        Else
         If reixa.CellPicture = nocheck.Picture Then
          comprovarsiaquestabobinaescorrecteiavisar
          If cadbl(metressel) >= cadbl(mtrsnecessaris) Then MsgBox "Ja passes dels metres necessaris": Exit Sub
          Set reixa.CellPicture = check.Picture: reixa.TextMatrix(reixa.row, reixa.col) = "1": check.Tag = atrim(reixa.row)
         End If
   End If
   reixa.CellForeColor = reixa.CellBackColor
   colocar_assignats
   metressel = format(sumar_seleccionats, "#,##0")
   mtrsnecessaris = format(mtrsnecessaris, "#,##0")
  End If
End Sub
Sub colocar_assignats()
  Dim total As Double
  Dim metres As Byte
  Dim seleccionat As Byte
  Dim mtrsdis As Double
  Dim srestu As Boolean
  Dim sparcial As Boolean
  Dim sel As String
  
  triada = cadbl(check.Tag)
  If partirmanual.Value <> 1 Then triada = 0
  
  'Primer escullo els restus
  For i = 1 To reixa.Rows - 1
    sel = reixa.TextMatrix(i, columnadelcamp("seleccionat"))
    srestu = cabool(reixa.TextMatrix(i, columnadelcamp("resto")))
    If (sel = "1" Or sel = "2") And srestu And i <> triada Then
       mtrsdis = cadbl(reixa.TextMatrix(i, columnadelcamp("mtrsdisponibles")))
       If sel = "2" Then mtrsdis = cadbl(reixa.TextMatrix(i, columnadelcamp("mtrsassignats")))
       If (total + mtrsdis) > cadbl(mtrsnecessaris) And sel <> "2" Then
            mtrsdis = cadbl(mtrsnecessaris) - total
       End If
       total = total + mtrsdis
       reixa.TextMatrix(i, columnadelcamp("mtrsassignats")) = formatreixa(mtrsdis)
    End If
  Next i
  
  'ara escullo totes les que no sol restus ni parcials
  For i = 1 To reixa.Rows - 1
    sel = reixa.TextMatrix(i, columnadelcamp("seleccionat"))
    srestu = cabool(reixa.TextMatrix(i, columnadelcamp("resto")))
    sparcial = cabool(reixa.TextMatrix(i, columnadelcamp("parcial")))
    If (sel = "1" Or sel = "2") And Not srestu And Not sparcial And i <> triada Then
       mtrsdis = cadbl(reixa.TextMatrix(i, columnadelcamp("mtrsdisponibles")))
       If sel = "2" Then mtrsdis = cadbl(reixa.TextMatrix(i, columnadelcamp("mtrsassignats")))
       If (total + mtrsdis) > cadbl(mtrsnecessaris) And sel <> "2" Then
            mtrsdis = cadbl(mtrsnecessaris) - total
       End If
       total = total + mtrsdis
       reixa.TextMatrix(i, columnadelcamp("mtrsassignats")) = formatreixa(mtrsdis)
    End If
  Next i
  
  'ara miro les parcials i tallaré una parcial si cal
  For i = 1 To reixa.Rows - 1
    sel = reixa.TextMatrix(i, columnadelcamp("seleccionat"))
    sparcial = cabool(reixa.TextMatrix(i, columnadelcamp("parcial")))
    If (sel = "1" Or sel = "2") And sparcial And i <> triada Then
       mtrsdis = cadbl(reixa.TextMatrix(i, columnadelcamp("mtrsdisponibles")))
       If sel = "2" Then mtrsdis = cadbl(reixa.TextMatrix(i, columnadelcamp("mtrsassignats")))
       If (total + mtrsdis) > cadbl(mtrsnecessaris) And sel <> "2" Then
            mtrsdis = cadbl(mtrsnecessaris) - total
       End If
       total = total + mtrsdis
       reixa.TextMatrix(i, columnadelcamp("mtrsassignats")) = formatreixa(mtrsdis)
    End If
  Next i
  
  'si no tinc els metres necessaris miro si haig de partir la ultima sel.leccionada
  If total < cadbl(mtrsnecessaris) And triada > 0 Then
     i = triada
     mtrsdis = cadbl(reixa.TextMatrix(i, columnadelcamp("mtrsdisponibles")))
     If (total + mtrsdis) > cadbl(mtrsnecessaris) Then
            mtrsdis = cadbl(mtrsnecessaris) - total
     End If
     total = total + mtrsdis
     reixa.TextMatrix(i, columnadelcamp("mtrsassignats")) = formatreixa(mtrsdis)
  End If
  
End Sub

Sub canviarnomcapcalera()
    r = InputBox("Entra el nom que vols a la capçalera", "Canvi de nom de columna")
    If atrim(r) <> "" Then
          escriure_ini "NomsReixa", UCase(rstconsulta.Fields(reixa.ColData(reixa.col)).Name) + "-nom", r, iniconfigreixa
          reixa.TextMatrix(0, reixa.col) = r
    End If
End Sub

Private Sub reixa_DblClick()
  Dim mtrsass As String
  Dim mtrsdif As Double
  If vfiltrebobinesdesdeimpresores Then
       ensenya_situacioamagatzem
       Exit Sub
  End If
  If Reserves.Caption <> "Reserves" Then
   ' botocomprar_Click
    Exit Sub
  End If
  
  If cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("seleccionat"))) <> 1 Then Exit Sub
   If reixa.col = columnadelcamp("mtrsassignats") Then
      mtrsass = InputBox("Entra els metres que vols assignar", "Assignar metres", reixa.TextMatrix(reixa.row, columnadelcamp("mtrsassignats")))
      If cadbl(mtrsass) > 0 Then
       mtrsdif = cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("mtrsdisponibles")) - cadbl(mtrsass))
       If mtrsdif < 0 Then
           MsgBox "No pots agafar tants metres d'aquesta bobina", vbCritical + vbOKOnly, "Atenció"
           Exit Sub
       End If
       reixa.Text = mtrsass
       
       reixa.TextMatrix(reixa.row, columnadelcamp("mtrsdiferencia")) = formatreixa(mtrsdif)
       Command2_Click
      End If
   End If
   
End Sub
Sub ensenya_situacioamagatzem()
   Dim numpalet As Double
   Dim numbobina As Double
   If MsgBox("Vols imprimir la fulla de bobina?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
      numpalet = cadbl(reixa.TextMatrix(reixa.row, bobinesdentrada.columnadelcamp("PALET")))
      numbobina = cadbl(reixa.TextMatrix(reixa.row, bobinesdentrada.columnadelcamp("BOBINA")))
      If numpalet > 0 And numbobina > 0 Then
        Set dbstocks = dbtmp
        bobinesdentrada.imprimir_bobinaparcial numpalet, numbobina, , 1
      End If
   End If
End Sub

Private Sub reixa_LostFocus()
  guardar_amples_reixa
End Sub

Private Sub reixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 2 Then canviarnomcapcalera: Exit Sub
End Sub

Sub eliminar_reserva()
   Dim rstreserva As Recordset
   Dim numreserva As Double
   numreserva = cadbl(reixa.TextMatrix(reixa.row, columnadelcamp("idreserva")))
   Set rstreserva = dbtmp.OpenRecordset("select * from reserves where idreserva=" + atrim(cadbl(numreserva)))
   If Not rstreserva.EOF Then
      rstreserva.Edit
      If cadbl(datalat.Recordset!metres) > 0 And cadbl(datalat.Recordset!comanda) <> 0 Then
         rstreserva!metresreservats = cadbl(rstreserva!metresreservats) - cadbl(datalat.Recordset!metres)
         reixa.TextMatrix(reixa.row, columnadelcamp("reservat")) = formatreixa(rstreserva!metresreservats)
       'Else: '
          'aquests els sumu però com que son negatius es resten
          'rstreserva!pendentsreservar = cadbl(rstreserva!pendentsreservar) + cadbl(datalat.Recordset!metres)
        '  reixa.TextMatrix(reixa.row, columnadelcamp("perreservar")) = cadbl(rstreserva!pendentsreservar)
      End If
      rstreserva.Update
   End If
   datalat.Recordset.Delete
   datalat.Refresh
   Set rstreserva = Nothing
   'refrescar_reixa
End Sub

Private Sub reixalat_DblClick()
If MsgBox("Vols eliminar aquesta reserva?", vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
       eliminar_reserva
   End If
End Sub

Private Sub Reserves_Click()
  If Reserves.Caption = "Reserves" Then
    If arguments(2) <> "comprant" Then comprovar_reserves_orfes
    passar_a_reservar
    Set dbcompres = DBEngine.OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
     Else
      passar_a_assignar
  End If
End Sub
Sub fer_where(vrst2 As Recordset, vwhere As String)
    Dim i As Byte
    Dim vrst As Recordset
    Set vrst = dbtmp.OpenRecordset("select * from reserves where idreserva=" + atrim(vrst2!midreserva))
    For i = 1 To vrst.Fields.Count - 1
       If vrst.Fields(i).Name <> "metresreservats" Then
           vwhere = vwhere + IIf(vwhere = "", "", " and ") + vrst.Fields(i).Name + "=" + tipusdato(vrst.Fields(i))
       End If
    Next i
End Sub
Function tipusdato(vf As Field) As String
   If vf.Type = 10 Then tipusdato = "'" + atrim(vf.Value) + "'"
   If vf.Type = 7 Or vf.Type = 6 Or vf.Type = 4 Then tipusdato = passaradecimalpunt(atrim(cadbl(vf.Value)))
   If vf.Type = 1 Then tipusdato = IIf(vf.Value, "True", "False")
   
End Function
Sub treure_reserves_duplicades()
   Dim i As Byte
   Dim rstreserves As Recordset
   Dim rst As Recordset
   Dim vsuma As Double
   Dim vborrats As Boolean
   Dim vwhere As String
   Dim vultimaidreserva As Long
   Dim vsql As String
   vsql = "SELECT Max(Reserves.idreserva) AS Midreserva From Reserves GROUP BY Reserves.Ample, Reserves.familia, Reserves.subfamilia, Reserves.familiacol, Reserves.subfamiliacol, Reserves.familiaad, Reserves.subfamiliaad, Reserves.espesor, Reserves.carestractat, Reserves.Plegat, Reserves.Solapa, Reserves.obert, Reserves.semielaborat, Reserves.semielaborat Having (((Count(Reserves.idreserva)) > 1)) ORDER BY Max(Reserves.idreserva)"

   
'   Set rstreserves = dbtmp.OpenRecordset("select * from reserves where idreserva in (" + vsql + ") order by idreserva")
   Set rstreserves = dbtmp.OpenRecordset(vsql)
   While Not rstreserves.EOF
    'If rstreserves!ample = 56 And rstreserves!familia = 578 Then Stop
    vwhere = ""
    vborrats = False
    vultimaidreserva = rstreserves!midreserva
    fer_where rstreserves, vwhere
     'Clipboard.Clear
     'Clipboard.SetText vwhere
     Set rst = dbtmp.OpenRecordset("select * from reserves where " + vwhere)
     vsuma = 0
     If rst.EOF Then GoTo cont
     rst.MoveLast: rst.MoveFirst
     For i = 1 To rst.RecordCount - 1
        vsuma = vsuma + cadbl(rst!metresreservats)
        rst.Delete
        vborrats = True
        rst.MoveNext
     Next i
     If vsuma > 0 Then rst.Edit: rst!metresreservats = vsuma: rst.Update
cont:
    If vborrats Then
         Set rstreserves = dbtmp.OpenRecordset(vsql)
         rstreserves.FindFirst "midreserva>=" + atrim(vultimaidreserva)
       Else: rstreserves.MoveNext
    End If
   Wend
End Sub
Sub comprovar_reserves_orfes()
   Dim rstre As Recordset
   Dim rstc As Recordset
   Dim llista As String
   Dim vdesreservar As Boolean
   'Set rstre = dbtmp.OpenRecordset("SELECT Reserves.idreserva FROM Reserves where reserves.idreserva not in (select idreserva from percomandaoclient);")
   'If Not rstre.EOF Then MsgBox "Hi han reserves orfes, sense comanda assignada", vbCritical, "Atenció"
   treure_reserves_duplicades
inici:
   llista = ""
   Set rstre = dbtmp.OpenRecordset("select * from reserves where metresreservats>0")
   While Not rstre.EOF
     Set rstc = dbtmp.OpenRecordset("select sum(metres) as tmetres from percomandaoclient where idreserva=" + atrim(rstre!idreserva))
     If Not rstc.EOF Then dbtmp.Execute "update reserves set metresreservats=" + atrim(cadbl(rstc!tmetres)) + " where idreserva=" + atrim(cadbl(rstre!idreserva))
     rstre.MoveNext
   Wend
   Set rstre = dbtmp.OpenRecordset("SELECT percomandaoclient.numcomanda as ncom, percomandaoclient.metres, percomandaoclient.idreserva, comandes.proximaseccio FROM percomandaoclient LEFT JOIN comandes ON percomandaoclient.numcomanda = comandes.comanda WHERE (((comandes.proximaseccio)<>'E'));")
   While Not rstre.EOF
    
    If rstre!ncom > 0 Then
       
       If vdesreservar Then
             r = "nopregunta"
            des_reservar rstre!ncom
             Else: llista = llista + " - " + atrim(rstre!ncom)
       End If
      ' eliminar_assignat cadbl(rstre!ncom)
    End If
    rstre.MoveNext
   Wend
   If llista <> "" And vdesreservar = False Then
        v = InputBox("Comandes assignades i encara reservades." + Chr(10) + llista + Chr(10) + " Comprova que l'estat no hagi passat de E a I." + vbNewLine + "ESCRIU [DESRESERVARLES] PER TREURE LA RESERVA.", "Comandes Assignades amb reserva feta encara.")
        If UCase(v) = "DESRESERVARLES" Then vdesreservar = True: GoTo inici
   End If
End Sub
Sub passar_a_reservar()
    assignarmat.BackColor = &H80C0FF 'taronja suau
    Reserves.BackColor = &HC0FFC0 'verd suau
    Command2.BackColor = &H80FF&
    Command2.Caption = "Ok... Reservar"
    Reserves.Caption = "Assignar"
    reixa.SelectionMode = flexSelectionByRow
    metressel.Visible = False
    frameclient.Visible = True
    'reixa.Width = 8000
    botocomprar.Visible = True
    Command2.Enabled = True
    desreservar.Visible = True
    partirmanual.Visible = False
    assignarstock.Visible = False
    botoajust.Visible = False
End Sub
Sub passar_a_assignar()
      assignarmat.BackColor = &HC0FFC0 'verdsuau
      Reserves.BackColor = &H80FF& 'taronja fort
      Command2.BackColor = &H80FF80 'verd fort
      Command2.Caption = "Ok... Assignar"
      Reserves.Caption = "Reserves"
      reixa.SelectionMode = flexSelectionFree
      reixa.Width = 11500
      metressel.Visible = True
      frameclient.Visible = False
      botocomprar.Visible = False
      Command2.Enabled = True
      desreservar.Visible = False
      partirmanual.Visible = True
      reixacomandes.Visible = False
      assignarstock.Visible = True
      botoajust.Visible = True
End Sub

Private Sub subfamad_Change()
carregar_subfamilies
End Sub

Private Sub subfamad_DropDown()
carregar_subfamilies
End Sub

Private Sub subfamcol_Change()
carregar_subfamilies
End Sub

Private Sub subfamcol_DropDown()
carregar_subfamilies
End Sub

Sub carregar_subfamilies(Optional combof As Control)
  Dim rstsub As Recordset
  Dim combo As Control
  Dim subfamilia As String
  
  Set combo = assignarmat.ActiveControl
  If Not combof Is Nothing Then Set combo = combof
  If assignarmat.Controls(combo.Tag).ListIndex = -1 And combof Is Nothing Then MsgBox "Primer has d'escullir la familia": Exit Sub
  'If combo.ListIndex = -1 Then combo.Clear: Exit Sub
  If combo.Name = "subfammat" And fammat.ListIndex <> -1 Then r = " codifam=" + atrim(cadbl(fammat.ItemData(fammat.ListIndex))): subfamilia = "subfamiliesmaterials"
  If combo.Name = "subfamcol" And famcol.ListIndex <> -1 Then r = " codifam=" + atrim(cadbl(famcol.ItemData(famcol.ListIndex))): subfamilia = "subfamiliescolorants"
  If combo.Name = "subfamad" And famad.ListIndex <> -1 Then r = " codifam=" + atrim(cadbl(famad.ItemData(famad.ListIndex))): subfamilia = "subfamiliesaditius"
    combo.Clear

  If subfamilia <> "" Then
     v = subfammat.Text
  '   If Mid(v + " ", 1, 4) = "### " Then
  '     v = Mid(v, 5)
  '     subfammat.Text = v
  '   End If
     Set rstsub = dbtmpb.OpenRecordset("select codi,descripcio" + IIf(subfamilia = "subfamiliesmaterials", ",matcompatible", "") + " from " + subfamilia + " where " + r) '+ " and descripcio like '*" + treure_apostrof(v) + "*'")
    Else: Exit Sub
  End If
  
  While Not rstsub.EOF
    If Reserves.Caption = "Assignar" And subfamilia = "subfamiliesmaterials" Then If rstsub!matcompatible = "S" Then vcompat = "### " Else vcompat = ""
    combo.AddItem vcompat + atrim(rstsub!descripcio)
    combo.ItemData(combo.NewIndex) = cadbl(rstsub!codi)
    rstsub.MoveNext
  Wend
  
  
End Sub

Private Sub subfammat_Change()
   If assignarmat.ActiveControl.Name <> "subfammat" And arguments(2) <> "comprant" Then carregar_subfamilies
End Sub

Private Sub subfammat_Click()
  If assignarmat.ActiveControl.Name = "subfammat" Then
       infodescripciomat.Tag = ""
'       If Mid(subfammat + " ", 1, 4) = "### " Then
'            v = Mid(subfammat, 5)
'
'            subfammat.Text = v
'       End If
  End If
End Sub

Private Sub subfammat_DropDown()
  carregar_subfamilies
End Sub

Private Sub txtFields_Change(Index As Integer)

End Sub


