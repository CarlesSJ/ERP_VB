VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00EBC5C5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manteniment de Palets"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11895
   Icon            =   "palets.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Height          =   360
      Left            =   10215
      Picture         =   "palets.frx":59D62
      Style           =   1  'Graphical
      TabIndex        =   90
      TabStop         =   0   'False
      ToolTipText     =   "Veure l'albarà del proveïdor escanejat."
      Top             =   1770
      Width           =   420
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Index           =   2
      Left            =   7485
      Picture         =   "palets.frx":5A2EC
      Style           =   1  'Graphical
      TabIndex        =   82
      TabStop         =   0   'False
      ToolTipText     =   "Llista de canvis realitzats als parcials"
      Top             =   4635
      Width           =   315
   End
   Begin VB.CommandButton infocompra 
      BackColor       =   &H00FF8080&
      Caption         =   "Info Compra"
      Height          =   270
      Left            =   10395
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   930
      Width           =   1230
   End
   Begin VB.Data bobinesent 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   10620
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8760
      Visible         =   0   'False
      Width           =   1275
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   -180
      Top             =   4095
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   120
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   51
      Text            =   "34530/2"
      Top             =   525
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Data palets 
      Caption         =   "palets"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   4665
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from Palets order by idpalet DESC"
      Top             =   60
      Width           =   3105
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   -15
      TabIndex        =   39
      Top             =   -90
      Width           =   11760
      Begin VB.CommandButton consultar 
         Height          =   360
         Left            =   1845
         Picture         =   "palets.frx":5A876
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   "Buscar Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton assignamaterial 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Assignar Material"
         Height          =   390
         Left            =   9405
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   150
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   1410
         Picture         =   "palets.frx":5AE00
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Acceptar canvis"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   11280
         Picture         =   "palets.frx":5B38A
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Sortir"
         Top             =   150
         Width           =   390
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   520
         Picture         =   "palets.frx":5B914
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Edicio del  Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   965
         Picture         =   "palets.frx":5BE9E
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   90
         Picture         =   "palets.frx":5C428
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
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
         Left            =   2490
         TabIndex        =   45
         Top             =   150
         Width           =   2025
      End
   End
   Begin VB.Frame framebobines 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Bobines"
      Height          =   4230
      Left            =   0
      TabIndex        =   1
      Top             =   4410
      Width           =   11880
      Begin VB.CommandButton betiquetabobinaprov 
         BackColor       =   &H0017D062&
         Height          =   390
         Left            =   60
         Picture         =   "palets.frx":5C9B2
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Etiqueta de la bobina de proveidor."
         Top             =   1455
         Width           =   345
      End
      Begin VB.CommandButton Command6 
         Height          =   390
         Left            =   60
         Picture         =   "palets.frx":5CDD4
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Passar llistat bobines a Excel."
         Top             =   1050
         Width           =   345
      End
      Begin VB.CommandButton Command4 
         Height          =   390
         Left            =   60
         Picture         =   "palets.frx":5D35E
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Duplicar la bobina sel.leccionada per X vegades."
         Top             =   645
         Width           =   345
      End
      Begin VB.CommandButton imprimirparcial 
         Height          =   390
         Left            =   60
         Picture         =   "palets.frx":5D8E8
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Imprimir l'etiqueta de bobina parcial."
         Top             =   240
         Width           =   345
      End
      Begin VB.Data parcials 
         Caption         =   "parcials"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   9195
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Parcials"
         Top             =   -135
         Visible         =   0   'False
         Width           =   2025
      End
      Begin MSDBGrid.DBGrid reixaparcials 
         Bindings        =   "palets.frx":5DE72
         Height          =   3885
         Left            =   7470
         OleObjectBlob   =   "palets.frx":5DE85
         TabIndex        =   46
         Top             =   225
         Width           =   4365
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "palets.frx":5F752
         Height          =   3885
         Left            =   420
         OleObjectBlob   =   "palets.frx":5F764
         TabIndex        =   2
         Top             =   195
         Width           =   7035
      End
      Begin VB.Data bobines 
         Caption         =   "bobines"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3315
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Bobines"
         Top             =   90
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label totalkilosteorics 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   6120
         TabIndex        =   79
         Top             =   0
         Width           =   45
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "å"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   5940
         TabIndex        =   78
         Top             =   0
         Width           =   180
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "å"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   2970
         TabIndex        =   74
         Top             =   0
         Width           =   180
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "å"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   1425
         TabIndex        =   73
         Top             =   -15
         Width           =   165
      End
      Begin VB.Label totalkilos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   3150
         TabIndex        =   72
         Top             =   0
         Width           =   45
      End
      Begin VB.Label totalmetres 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   1590
         TabIndex        =   71
         Top             =   -15
         Width           =   45
      End
      Begin VB.Label etparcials 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Parcials"
         Height          =   195
         Left            =   6975
         TabIndex        =   47
         Top             =   -15
         Width           =   555
      End
   End
   Begin VB.Frame framepalets 
      BackColor       =   &H00D29F7D&
      Caption         =   "Palet"
      Enabled         =   0   'False
      Height          =   3405
      Left            =   120
      TabIndex        =   0
      Top             =   705
      Width           =   11595
      Begin VB.CommandButton bdesactivarpalets 
         Height          =   285
         Left            =   8025
         Picture         =   "palets.frx":61DA3
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Desactivar aquest palet i tots els d'aquest albarà"
         Top             =   1740
         Width           =   465
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Dataprevistaarribada"
         DataSource      =   "palets"
         Height          =   285
         Index           =   16
         Left            =   10170
         TabIndex        =   85
         Top             =   1725
         Width           =   1275
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Dataactivacio"
         DataSource      =   "palets"
         Height          =   285
         Index           =   15
         Left            =   6690
         TabIndex        =   83
         Top             =   1725
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         Height          =   285
         Left            =   8070
         Picture         =   "palets.frx":6232D
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Edicio del  Registres"
         Top             =   2070
         Width           =   270
      End
      Begin VB.CheckBox matclient 
         BackColor       =   &H00D29F7D&
         Caption         =   "Material del Client"
         DataField       =   "materialdelclient"
         DataSource      =   "palets"
         Height          =   345
         Left            =   10170
         TabIndex        =   70
         Top             =   2670
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Height          =   285
         Left            =   8355
         Picture         =   "palets.frx":628B7
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Borrar la data recepció"
         Top             =   2070
         Width           =   270
      End
      Begin VB.CommandButton Command2 
         Height          =   360
         Left            =   2775
         Picture         =   "palets.frx":62E41
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Edicio del  Registres"
         Top             =   645
         Width           =   330
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Index           =   13
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1860
         Width           =   660
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Numpalet"
         DataSource      =   "palets"
         Height          =   285
         Index           =   5
         Left            =   6690
         TabIndex        =   18
         Top             =   2385
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Numalb"
         DataSource      =   "palets"
         Height          =   285
         Index           =   6
         Left            =   6690
         MaxLength       =   15
         TabIndex        =   14
         Top             =   1095
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Numlot"
         DataSource      =   "palets"
         Height          =   285
         Index           =   7
         Left            =   6690
         MaxLength       =   20
         TabIndex        =   15
         Top             =   1410
         Width           =   2145
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Datarev"
         DataSource      =   "palets"
         Height          =   285
         Index           =   8
         Left            =   10170
         TabIndex        =   16
         Top             =   1410
         Width           =   1275
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0C0C0&
         DataField       =   "Datarec"
         DataSource      =   "palets"
         Height          =   285
         Index           =   9
         Left            =   6690
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2055
         Width           =   1290
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Observ"
         DataSource      =   "palets"
         Height          =   285
         Index           =   10
         Left            =   6690
         MaxLength       =   50
         TabIndex        =   19
         Top             =   2700
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Numpaletpro"
         DataSource      =   "palets"
         Height          =   285
         Index           =   12
         Left            =   6690
         MaxLength       =   20
         TabIndex        =   13
         Top             =   780
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.TextBox txtFields 
         DataField       =   "micres"
         DataSource      =   "palets"
         Height          =   285
         Index           =   11
         Left            =   1620
         TabIndex        =   9
         Top             =   1845
         Width           =   780
      End
      Begin VB.CheckBox microp 
         BackColor       =   &H00D29F7D&
         Caption         =   "Microperforat"
         DataField       =   "microperforat"
         DataSource      =   "palets"
         Height          =   300
         Left            =   2910
         TabIndex        =   23
         Top             =   3015
         Width           =   1470
      End
      Begin VB.ComboBox Combo3 
         DataField       =   "obert"
         DataSource      =   "palets"
         Height          =   315
         ItemData        =   "palets.frx":633CB
         Left            =   2145
         List            =   "palets.frx":633D8
         TabIndex        =   22
         Top             =   3015
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "carestractat"
         DataSource      =   "palets"
         Height          =   315
         ItemData        =   "palets.frx":633E5
         Left            =   1335
         List            =   "palets.frx":633F2
         TabIndex        =   21
         Top             =   3015
         Width           =   615
      End
      Begin VB.ComboBox tractat 
         DataField       =   "tractat"
         DataSource      =   "palets"
         Height          =   315
         ItemData        =   "palets.frx":633FF
         Left            =   1605
         List            =   "palets.frx":6340F
         TabIndex        =   12
         Top             =   2445
         Width           =   1710
      End
      Begin VB.TextBox nommaterial 
         Height          =   285
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   1020
         Width           =   2505
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "semielaborat"
         DataSource      =   "palets"
         Height          =   315
         ItemData        =   "palets.frx":6343B
         Left            =   705
         List            =   "palets.frx":63445
         TabIndex        =   20
         Top             =   3015
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   120
         Top             =   165
      End
      Begin VB.CheckBox chkFields 
         BackColor       =   &H00D29F7D&
         DataField       =   "Disponible"
         DataSource      =   "palets"
         Height          =   285
         Index           =   14
         Left            =   8745
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkFields 
         BackColor       =   &H00D29F7D&
         DataField       =   "Mostrasino"
         DataSource      =   "palets"
         Height          =   285
         Index           =   11
         Left            =   8010
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3060
         Width           =   225
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Solapa"
         DataSource      =   "palets"
         Height          =   285
         Index           =   4
         Left            =   1620
         TabIndex        =   11
         Top             =   2145
         Width           =   780
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Plegat"
         DataSource      =   "palets"
         Height          =   285
         Index           =   3
         Left            =   3045
         TabIndex        =   8
         Top             =   1560
         Width           =   780
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Ample"
         DataSource      =   "palets"
         Height          =   285
         Index           =   2
         Left            =   1620
         TabIndex        =   7
         Top             =   1545
         Width           =   795
      End
      Begin VB.TextBox txtFields 
         DataField       =   "codimatprognou"
         DataSource      =   "palets"
         Height          =   285
         Index           =   1
         Left            =   1635
         TabIndex        =   6
         Top             =   1020
         Width           =   540
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0C0C0&
         DataField       =   "Idpalet"
         DataSource      =   "palets"
         Height          =   285
         Index           =   0
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   690
         Width           =   1125
      End
      Begin VB.TextBox preucompra 
         Height          =   315
         Left            =   9375
         TabIndex        =   69
         Top             =   3015
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtFields 
         DataField       =   "preucompra"
         DataSource      =   "palets"
         Height          =   285
         Index           =   14
         Left            =   9420
         TabIndex        =   66
         Top             =   3045
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label etimpostenv 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "etimpostenv"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7005
         TabIndex        =   91
         Top             =   165
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Data prevista arribada:"
         Height          =   255
         Index           =   19
         Left            =   8550
         TabIndex        =   86
         Top             =   1755
         Width           =   1905
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Activació:"
         Height          =   255
         Index           =   18
         Left            =   5460
         TabIndex        =   84
         Top             =   1755
         Width           =   1200
      End
      Begin VB.Label datacreacio 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         Height          =   570
         Left            =   10230
         TabIndex        =   77
         Top             =   510
         Width           =   1260
      End
      Begin VB.Label avg 
         BackStyle       =   0  'Transparent
         Caption         =   "Preu mig."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10110
         TabIndex        =   68
         Top             =   3060
         Width           =   1440
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00D29F7D&
         Caption         =   "Compra €/Kg:"
         DataField       =   "preucompra"
         Height          =   255
         Index           =   17
         Left            =   8385
         TabIndex        =   67
         Top             =   3090
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label refprod 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2100
         TabIndex        =   63
         Top             =   1320
         Width           =   2520
      End
      Begin VB.Label proveidor 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3135
         TabIndex        =   62
         Top             =   345
         Width           =   6645
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Grms/m2:"
         Height          =   255
         Index           =   16
         Left            =   3150
         TabIndex        =   59
         Top             =   1890
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D29F7D&
         Caption         =   "Micres"
         Height          =   195
         Left            =   2490
         TabIndex        =   58
         Top             =   1875
         Width           =   675
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Espesor:"
         Height          =   255
         Index           =   15
         Left            =   690
         TabIndex        =   57
         Top             =   1860
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Obert"
         Height          =   300
         Index           =   1
         Left            =   2235
         TabIndex        =   56
         Top             =   2775
         Width           =   540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cares Tractat"
         Height          =   300
         Index           =   0
         Left            =   1185
         TabIndex        =   55
         Top             =   2760
         Width           =   1020
      End
      Begin VB.Label numerodepalet 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   90
         TabIndex        =   52
         Top             =   150
         Width           =   1860
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "T/L"
         Height          =   300
         Index           =   3
         Left            =   810
         TabIndex        =   48
         Top             =   2760
         Width           =   765
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible"
         Height          =   255
         Index           =   14
         Left            =   8985
         TabIndex        =   37
         Top             =   2085
         Width           =   870
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Tractat per?:"
         Height          =   255
         Index           =   13
         Left            =   660
         TabIndex        =   36
         Top             =   2490
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Palet Prov.:"
         Height          =   255
         Index           =   12
         Left            =   5445
         TabIndex        =   35
         Top             =   825
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Mostra S/N:"
         Height          =   255
         Index           =   11
         Left            =   7050
         TabIndex        =   33
         Top             =   3105
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Observacions:"
         Height          =   255
         Index           =   10
         Left            =   5445
         TabIndex        =   32
         Top             =   2760
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Recepció:"
         Height          =   255
         Index           =   9
         Left            =   5445
         TabIndex        =   31
         Top             =   2100
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Alb. Prov:"
         Height          =   255
         Index           =   8
         Left            =   8925
         TabIndex        =   30
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Lot del Prov.:"
         Height          =   255
         Index           =   7
         Left            =   5445
         TabIndex        =   29
         Top             =   1455
         Width           =   1200
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Albarà Prov.:"
         Height          =   255
         Index           =   6
         Left            =   5430
         TabIndex        =   28
         Top             =   1140
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Q. de Palets:"
         Height          =   255
         Index           =   5
         Left            =   5445
         TabIndex        =   27
         Top             =   2445
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Solapa:"
         Height          =   255
         Index           =   4
         Left            =   675
         TabIndex        =   26
         Top             =   2190
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00D29F7D&
         Caption         =   "Plegat:"
         Height          =   255
         Index           =   3
         Left            =   2460
         TabIndex        =   25
         Top             =   1590
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ample:"
         Height          =   255
         Index           =   2
         Left            =   675
         TabIndex        =   24
         Top             =   1590
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Producte:"
         Height          =   255
         Index           =   1
         Left            =   675
         TabIndex        =   5
         Top             =   1065
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Palet:"
         Height          =   255
         Index           =   0
         Left            =   675
         TabIndex        =   3
         Top             =   735
         Width           =   540
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   90
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   870
      ScaleHeight     =   3855
      ScaleWidth      =   7980
      TabIndex        =   50
      Top             =   285
      Visible         =   0   'False
      Width           =   8040
   End
   Begin VB.Label botoensenyarpacking 
      Caption         =   "Label6"
      Height          =   90
      Left            =   45
      TabIndex        =   87
      Top             =   8715
      Width           =   300
   End
   Begin VB.Label comanda 
      Caption         =   "Label3"
      Height          =   30
      Left            =   30
      TabIndex        =   64
      Top             =   5205
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Menu m_compres 
      Caption         =   "Compres"
      Begin VB.Menu m_recepciomat 
         Caption         =   "Recepcio de material."
      End
      Begin VB.Menu albdecompres 
         Caption         =   "Albarans de compres."
      End
      Begin VB.Menu mdevmat 
         Caption         =   "Devolució de material"
      End
      Begin VB.Menu mactivaciodelspalets 
         Caption         =   "Activació de palets."
      End
      Begin VB.Menu mactivaciodepalets 
         Caption         =   "Arribada de palets."
      End
   End
   Begin VB.Menu mvendes 
      Caption         =   "Vendes"
      Visible         =   0   'False
      Begin VB.Menu mnpapersfrontals 
         Caption         =   "Papers Frontals"
      End
   End
   Begin VB.Menu m_opcions 
      Caption         =   "Opcions"
      Begin VB.Menu mprestatges 
         Caption         =   "Manteniment de Prestatges"
      End
      Begin VB.Menu mprestNOUS 
         Caption         =   "Manteniment prestatges NOUS"
      End
      Begin VB.Menu m_grupdepalets 
         Caption         =   "Grups de palets"
      End
      Begin VB.Menu m_grupsdecompatibles 
         Caption         =   "Grups de materials compatibles entre ells"
      End
      Begin VB.Menu matarpicos 
         Caption         =   "Passar picos <500 a acavats"
      End
      Begin VB.Menu MCANVISSITUACIO 
         Caption         =   "Canvis Situacio bobines TORERUS"
      End
      Begin VB.Menu membpalets 
         Caption         =   "Embolicar palets - Control"
      End
      Begin VB.Menu msincronitzacionstorerus 
         Caption         =   "Comptador de sincronitzacions de la Tablet de TORERUS"
      End
      Begin VB.Menu mbuscarbobproveidor 
         Caption         =   "Buscar numero de bobina del proveïdor"
      End
   End
   Begin VB.Menu mimprimir 
      Caption         =   "Imprimir"
      Begin VB.Menu packinglistcomanda 
         Caption         =   "Packing-List Comanda"
      End
      Begin VB.Menu hitoricpackinglist 
         Caption         =   "Historic Packing-List"
      End
      Begin VB.Menu mbobimp 
         Caption         =   "Bobines d'Impresores"
         Begin VB.Menu mbobperpujarimp 
            Caption         =   "Per Pujar"
         End
         Begin VB.Menu mbobperbaixarimp 
            Caption         =   "Per Baixar"
         End
         Begin VB.Menu mbaixaranivellar 
            Caption         =   "Per Baixar (per anivellar)"
         End
         Begin VB.Menu menuperrevisaraimp 
            Caption         =   "Per revisar a IMP"
         End
      End
      Begin VB.Menu mbobineslam 
         Caption         =   "Bobines de Laminadores"
      End
      Begin VB.Menu metiquetapalet 
         Caption         =   "Etiqueta-Palet"
         Begin VB.Menu mtotes 
            Caption         =   "Totes les bobines"
            Begin VB.Menu bobspendentsimp 
               Caption         =   "Pendents d'Imprimir"
            End
            Begin VB.Menu impbobdesdefins 
               Caption         =   "Desde a fins palet"
            End
         End
         Begin VB.Menu mllistatdiametre 
            Caption         =   "Parcials modificats per Diametre pendents d'imprimir."
         End
         Begin VB.Menu munabobina 
            Caption         =   "Una bobina"
         End
         Begin VB.Menu escullirimpetbob 
            Caption         =   "Escullir Impresora Et.Bobina"
         End
      End
      Begin VB.Menu llistatsensemoviment 
         Caption         =   "Llistat bobines sense moviment entre dates"
      End
      Begin VB.Menu mllistaperpantalla 
         Caption         =   "Llistar per pantalla"
      End
   End
   Begin VB.Menu invnou 
      Caption         =   "Inventari"
      Begin VB.Menu mllistinvn 
         Caption         =   "Llistats"
         Begin VB.Menu llistrealestocsenseassignar 
            Caption         =   "Llistat d'estoc real "
            Begin VB.Menu llestocrealdisp 
               Caption         =   "Disponible + Grups"
               Begin VB.Menu mdisponiblegrups 
                  Caption         =   "Disponible + Grups"
               End
               Begin VB.Menu mdisponiblegrupsdetallcompres 
                  Caption         =   "Disponible + Grups (Detall de compres)"
               End
            End
            Begin VB.Menu llestrassig 
               Caption         =   "Assignat"
            End
         End
         Begin VB.Menu mllistatalbaranscompres 
            Caption         =   "Llistat Albarans compres entre dates (Comptabilitat)"
         End
         Begin VB.Menu mllistatalbaransvsfacturesSAP 
            Caption         =   "Llistat Albarans Vs Factures SAP"
         End
         Begin VB.Menu llisestocproduccio 
            Caption         =   "Llistat d'estoc en producció"
         End
         Begin VB.Menu llistestocxrentregar 
            Caption         =   "Llistat d'estoc a punt per entregar"
         End
      End
      Begin VB.Menu ivprestat 
         Caption         =   "Inventari Prestatgeries"
         Begin VB.Menu invstoctteoric 
            Caption         =   "Inventari Estoc Teóric "
         End
         Begin VB.Menu invtotselsforats 
            Caption         =   "Inventari de Tots els forats"
         End
         Begin VB.Menu mbobinesparcials2 
            Caption         =   "Inventari de bobines parcials"
            Begin VB.Menu mbobinesparcials 
               Caption         =   "Tots els Parcials"
            End
            Begin VB.Menu mparcialssenseregularitzardiametre 
               Caption         =   "Només parcials sense regularitzar el diametre"
            End
         End
      End
      Begin VB.Menu llinventari 
         Caption         =   "Inventari Vell"
         Visible         =   0   'False
         Begin VB.Menu llistatinventarisensereserva 
            Caption         =   "Llistat d'estoc real sense assignar"
         End
         Begin VB.Menu llistatinventariambreserva 
            Caption         =   "Llistat de reservats"
         End
         Begin VB.Menu llestocimp 
            Caption         =   "Llistat d'estoc a impresores"
            Begin VB.Menu inventariimpanonim 
               Caption         =   "Llistat d'anònim"
            End
         End
         Begin VB.Menu llistestlam 
            Caption         =   "Llistat d'estoc a laminadores"
            Begin VB.Menu invlamimpres 
               Caption         =   "Llistat d'imprès "
            End
            Begin VB.Menu invlamanonim 
               Caption         =   "Llistat d'anònim"
            End
         End
         Begin VB.Menu llisestocreb 
            Caption         =   "Llistat d'estoc a rebobinadores"
            Begin VB.Menu invrebimpres 
               Caption         =   "Llistat d'imprès i/o Laminat"
            End
            Begin VB.Menu invrebanonim 
               Caption         =   "Llistat d'anònim"
            End
         End
      End
      Begin VB.Menu mcontrolestocdeseguretat 
         Caption         =   "Control Estoc de seguretat"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Dim querywhere As String
Dim queryorder As String
Dim buscant As Boolean
Dim grmcm3 As Double
Dim grmm2 As Double

Sub activarframes(estat As Boolean)
  framepalets.Enabled = estat
  'framebobines.Enabled = estat
  If reixaparcials.EditActive Then parcials.Recordset.Edit: parcials.Recordset.Update
  If DBGrid1.EditActive Then bobines.Recordset.Edit: bobines.Recordset.Update
  On Error Resume Next
  palets.UpdateControls
  bobines.UpdateControls
  parcials.UpdateControls
  On Error GoTo 0
  
  DBGrid1.AllowAddNew = estat
  DBGrid1.AllowDelete = estat
  DBGrid1.AllowUpdate = estat
  
  reixaparcials.AllowAddNew = estat
  reixaparcials.AllowDelete = estat
  reixaparcials.AllowUpdate = estat
  
  
  
End Sub

Private Sub albdecompres_Click()
  If Not comprovaraccessabip Then MsgBox "No tens access al servidor de beep no es pujarant les linies"
    albaranscompres.Show 1
End Sub

 Sub alta_Click()
  Dim vcomandapartirbobines As Double
  Dim rst As Recordset
  vcomandapartirbobines = cadbl(InputBox("Entra la comanda d'Inplacsa que has utilitzat per partir les bobines d'un palet." + vbNewLine + "SI EL QUE VOLS ES CREAR UN PALET NOU HAS DE FER UNA COMPRA I RECEPCIÓ DE MATERIAL.", "Partir bobines d'un palet"))
  If vcomandapartirbobines = 0 Then
        If MsgBox("Vols crear el palet manualment?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
        GoTo noupalet
  End If
  Set rst = dbtmpb.OpenRecordset("select * from comandes where comanda=" + atrim(vcomandapartirbobines))
  If rst.EOF Then MsgBox "Aquesta comanda no existeix.", vbCritical, "Error": Exit Sub
     
  If rst!client <> 7 Then
       MsgBox "Aquesta comanda no es de [7-Rebobinadora Inplacsa] no puc utilitzar-la per partir bobines."
        Else: partirbobinesicrearpalets cadbl(vcomandapartirbobines)
  End If
  Set rst = Nothing
  Exit Sub
noupalet:
  noupalet
End Sub
Sub partirbobinesicrearpalets(vcomanda As Double)
  Dim rst As Recordset
  Dim rstpkg As Recordset
  Dim rstreb As Recordset
  Dim vsql As String
  Dim vpalet As Double
  Dim vnumpaletnou As Double
  Dim vample As Double
  Dim vmidesample As String
  
  Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
  Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(vcomanda))
  If rst.EOF Then Exit Sub
  If rst!proximaseccio <> "V" And rst!proximaseccio <> "T" Then MsgBox "Aquesta comanda encara no està acabada de rebobinar.", vbCritical, "Error": GoTo fi
  Set rstpkg = dbtmp.OpenRecordset("select * from parcials where comanda='" + atrim(vcomanda) + "'")
  If rstpkg.EOF Then MsgBox "No hi ha bobines al packinglist d'aquesta comanda.", vbCritical, "Error": Exit Sub
  'Set rstreb = dbbaixes.OpenRecordset("SELECT rebobinadores.comanda, bobinesreb.numerodebobina, bobinesreb.kilos, bobinesreb.metres, bobinesreb.ample, bobinesentreb.palet, bobinesentreb.bobina FROM (bobinesreb LEFT JOIN rebobinadores ON bobinesreb.controlid = rebobinadores.Id) LEFT JOIN bobinesentreb ON bobinesreb.Id = bobinesentreb.id WHERE (((rebobinadores.comanda)=" + atrim(vcomanda) + "));")
  vsql = "SELECT Last(bobinesentreb.palet) AS Paletentrada, Last(bobinesentreb.bobina) AS bobinaentrada, bobinesreb.numerodebobina AS bobinasortida, First(bobinesreb.kilos) AS Tkilos, First(bobinesreb.metres) AS Tmetres,First(bobinesreb.ample) AS Tample, First(bobinesreb.palet) AS Tpalet FROM rebobinadores LEFT JOIN (bobinesreb LEFT JOIN bobinesentreb ON bobinesreb.Id = bobinesentreb.id) ON rebobinadores.Id = bobinesreb.controlid Where (((rebobinadores.comanda) = " + atrim(vcomanda) + ") And ((bobinesreb.controlid) Is Not Null)) GROUP BY bobinesreb.numerodebobina order by First(bobinesreb.ample),Last(bobinesentreb.palet),numerodebobina;"
  Set rstreb = dbbaixes.OpenRecordset(vsql)
  If rstreb.EOF Then MsgBox "No hi ha bobines de sortida de rebobinadora en aquesta comanda.", vbCritical, "Error": GoTo fi
  rstreb.MoveLast: rstreb.MoveFirst
  's'han de crear les bobines per el palet nou depenen de les bobines de la comanda i del palet d'origen
 
  vpalet = 0
  vample = 0
  While Not rstreb.EOF
    If vpalet <> cadbl(rstreb!paletentrada) Or vample <> cadbl(rstreb!tample) Then
         vpalet = rstreb!paletentrada: crearpaletnou vpalet, vnumpaletnou, vcomanda, cadbl(rstreb!tample)
         If vample <> cadbl(rstreb!tample) Then vmidesample = vmidesample + IIf(vmidesample <> "", " i ", "") + atrim(rstreb!tample)
    End If
    vample = cadbl(rstreb!tample)
    If vnumpaletnou = 0 Then GoTo fi
    crearbobinanova vnumpaletnou, rstreb
    ' dbtmp.Execute "Insert into bobines (idpalet,idbobina,"
    rstreb.MoveNext
  Wend
  rstreb.MoveFirst
  palets.Refresh
  palets.Recordset.MoveLast
  'palets.Recordset.FindFirst "idpalet=" + atrim(vnumpaletnou)
  MsgBox "S'han creat " + atrim(rstreb.RecordCount) + " bobines de sortida de REBOBINADORA de " + atrim(vmidesample) + "Cm." + vbNewLine + "Es correcte aquesta mida?", vbExclamation + vbDefaultButton2 + vbYesNo, "ATENCIÓ"
  dbtmp.Execute "update comandes set proximaseccio='T' where comanda=" + atrim(vcomanda)
fi:
  Set rstreb = Nothing
  Set rst = Nothing
End Sub
Sub crearbobinanova(vnumpaletnou As Double, rstreb As Recordset)
  Dim rst As Recordset
  Dim vultimnumbobina As Double
  Set rst = dbtmp.OpenRecordset("select * from bobines where idpalet=" + atrim(vnumpaletnou) + " order by idbobina desc")
  If rst.EOF Then
       vultimnumbobina = 0
        Else: vultimnumbobina = rst!idbobina
  End If
  rst.AddNew
  rst!idpalet = vnumpaletnou
  rst!idbobina = vultimnumbobina + 1
  rst!mts = rstreb!tmetres
  rst!kilos = rstreb!tkilos
  rst!pesdelproveidor = rstreb!tkilos
  rst!sit = "REB"
  rst!disponible = rstreb!tmetres
  rst!tamanycanutu = 15.2
  rst!numpaletpro = rstreb!tpalet
  rst!numbobina = atrim(rstreb!bobinasortida)
  rst.Update
fi:
  Set rst = Nothing
End Sub
Sub crearpaletnou(vpalet As Double, vnumpaletnou As Double, vcomanda As Double, vample As Double)
  Dim rst As Recordset
  Dim rstantic As Recordset
  Set rstantic = dbtmp.OpenRecordset("select * from palets where idpalet=" + atrim(vpalet))
  If rstantic.EOF Then Exit Sub
  Set rst = dbtmp.OpenRecordset("select * from palets order by idpalet")
  rst.MoveLast
  vnumpaletnou = rst!idpalet + 1
  rst.AddNew
  rst!idpalet = vnumpaletnou
  For i = 1 To rst.Fields.Count - 1
      rst.Fields(i) = rstantic.Fields(i)
  Next i
  rst!numlot = atrim(vcomanda)
  rst!numalb = "1"
  rst!ample = vample
  rst!observ = "BOBINA PARTIDA A REB: " + atrim(vcomanda)
  rst!link_numpalet = atrim(vpalet)
  rst!dataaltapalet = Date
  rst!dataactivacio = Date
  rst!datarev = Date
  rst!datarec = Date
  rst!disponible = False
  rst!Dataprevistaarribada = Date
  rst.Update
  Set rst = Nothing
  Set rstantic = Nothing
End Sub
Sub noupalet()
  Dim rstpalets As Recordset
  Dim elgran As Double
  'If palets.Recordset.EOF Then Exit Sub
  If palets.Recordset.EditMode > 0 Then MsgBox "Estas editant. Primer finalitza l'edicio.": Exit Sub
  Set rstpalets = dbtmp.OpenRecordset("select max(idpalet) as elgran from palets")
  activarframes True
  elgran = cadbl(rstpalets!elgran)
  palets.Recordset.AddNew
  refprod = ""
  If Not buscant Then
   palets.Recordset!idpalet = elgran + 1
   palets.Recordset!Idprod = 0
   txtFields(0) = elgran + 1
   chkFields(14).Value = 0
   txtFields(3) = "0"
   txtFields(4) = "0"
    palets.Recordset!ample = 1
    palets.Recordset!tractat = "IMPRIMIR"
  End If
  If Not buscant Then
    If palets.Recordset.EditMode > 0 Then palets.Recordset.Update
   palets.Recordset.Bookmark = palets.Recordset.LastModified
   palets.Recordset.Edit
   palets.Recordset!ample = 0
   palets.Recordset!tractat = ""
   txtFields(2) = "0"
   tractat = ""
  End If
  Set rstpalets = Nothing
  If Screen.ActiveForm.Name = "Form1" Then txtFields(1).SetFocus
End Sub
Private Sub assignamaterial_Click()
  
   obrir_assignarmaterial
End Sub
Sub obrir_assignarmaterial()
obrir_dbllistats
 crear_taules_tmp
 assignarmat.Show '1
 Form1.Visible = False
 assignarmat.comprovarcomandesdesactivades
 assignarmat.comprovarmaterialexactependentdassignar
End Sub

Private Sub bdesactivarpalets_Click()
  If txtFields(6) = "" Or txtFields(8) = "" Then MsgBox "No hi ha albarà de proveïdor o data d'albarà.", vbCritical, "Error": Exit Sub
  If UCase(InputBox("Estas segur que vols desactivar aquest palet i tots els relacionats amb la comanda " + txtFields(6) + " amb data " + txtFields(8) + "?" + Chr(10) + "Escriu [DESACTIVAR] per desactivar-los.", "Desactivació")) = "DESACTIVAR" Then
     dbtmp.Execute "update palets set dataactivacio=null,Dataprevistaarribada=null where numalb='" + txtFields(6) + "' and datarev=#" + format(txtFields(8), "mm/dd/yy") + "#"
     MsgBox "Aquest palet i tots els relacionats amb la comanda han estat DESACTIVATS.", vbInformation, "DESACTIVATS"
     palets.Recordset.Move 0
  End If
End Sub

Private Sub betiquetabobinaprov_Click()
 Dim vubicaciobobina As String
   If bobines.Recordset.EOF Then
      MsgBox "No hi ha cap bobines escullida, primer escull una.", vbCritical, "Error"
      Exit Sub
   End If
   vubicaciobobina = nomfitxer_fotoetiquetabobina(atrim(bobines.Recordset!idpalet) + "/" + atrim(bobines.Recordset!idbobina))
   If existeix(vubicaciobobina) Then obrir_document vubicaciobobina
End Sub

Private Sub bobines_Reposition()
  Dim vubicaciobobina As String
  If Not bobines.Recordset.EOF Then
     parcials.RecordSource = "select * from parcials where idpalet=" + atrim(cadbl(palets.Recordset!idpalet)) + " and idbobina=" + atrim(cadbl(bobines.Recordset!idbobina))
     parcials.Refresh
     etparcials = "Parcials " + atrim(cadbl(bobines.Recordset!idpalet)) + "-" + atrim(cadbl(bobines.Recordset!idbobina))
     vubicaciobobina = nomfitxer_fotoetiquetabobina(atrim(bobines.Recordset!idpalet) + "/" + atrim(bobines.Recordset!idbobina))
     If existeix(vubicaciobobina) Then
          betiquetabobinaprov.BackColor = &H17D062
            Else: betiquetabobinaprov.BackColor = &H5C31DD
     End If
     
       Else:
         parcials.RecordSource = " select * from parcials where idpalet=-99999999"
         parcials.Refresh
         betiquetabobinaprov.BackColor = &H5C31DD
  End If
End Sub

Private Sub bobspendentsimp_Click()
   imprimiretpalet 0, 999999999, , True
End Sub

Private Sub chkFields_Click(Index As Integer)
  Dim error As String
  On Error GoTo fi
  error = Form1.ActiveControl.Name
  If Form1.ActiveControl.Name = "chkFields" And Index = 14 Then If Not IsDate(txtFields(9)) Then MsgBox "No pots marcar disponible si no hi ha data de recepció": chkFields(14).Value = 0
fi:
End Sub

Private Sub Command1_Click()
  gravar_canvis
End Sub
Sub gravar_canvis()
 If Not buscant Then
   If palets.Recordset.EditMode > 0 Then
     'If IsDate(txtFields(9)) Then passarelmaterialarebut
     If cadbl(txtFields(14)) = 0 And matclient.Value = 0 Then MsgBox "No hi ha preu en aquest palet... primer asigna un preu de compra.", vbCritical + vbOKOnly, "Atenció": Exit Sub
     If txtFields(14).DataChanged Then palets.Recordset!preucompraavg = False
     On Error GoTo errors
     gravar_reixa_bobines
     possar_siteimpostono
     If palets.Recordset.EditMode > 0 Then palets.Recordset.Update
     activarframes False
     palets.Recordset.Bookmark = palets.Recordset.LastModified
   End If
    Else: finalitzarbusqueda
 End If
 
 txtFields(0).Locked = True
 Exit Sub
errors:
  MsgBox err.Description
  Resume Next
End Sub
Sub possar_siteimpostono()
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select kgimpostenvasos from albaransbip where numpalet=" + atrim(palets.Recordset!idpalet))
  If Not rst.EOF Then
      If cadbl(rst!kgimpostenvasos) > 0 Then palets.Recordset!teimpost = True
       Else
        palets.Recordset!teimpost = False
        If cadbl(palets.Recordset!link_numpalet) > 0 Then
                Set rst = dbtmp.OpenRecordset("select teimpost from palets where idpalet=" + atrim(cadbl(palets.Recordset!link_numpalet)))
                If Not rst.EOF Then palets.Recordset!teimpost = rst!teimpost
        End If
  End If
  Set rst = Nothing
End Sub
Sub gravar_reixa_bobines()
   If Screen.ActiveForm.Name = "form1" Then
     DBGrid1.SetFocus
'     SendKeys "{DOWN}"
     bobines.Recordset.Move 0
   End If
End Sub
Private Sub Command2_Click()
  'txtFields(0).Enabled = True
  txtFields(0).Locked = False
  
End Sub

Sub crearbmpnumpalet(nump As String)
 On Local Error Resume Next
    Text1.Text = nump
    Me.Picture1.Cls
    Me.Picture1.CurrentX = -100
    Me.Picture1.CurrentY = -200
    Me.Picture1.Print Text1.Text
    err.Clear
    If Not existeix("c:\temp") Then MkDir "c:\temp"
    If Not existeix("c:\temp\numpalet.bmp") Then Kill "c:\temp\numpalet.bmp"
    SavePicture Me.Picture1.Image, "c:\temp\numpalet.bmp"
    If err.Number <> 0 Then MsgBox "Error:" + err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub Command3_Click()
  Dim palet As Double
    If MsgBox("Eliminaras la data de recepció d'aquest palet", vbCritical + vbYesNo, "Atenció") = vbNo Then Exit Sub
   palet = cadbl(palets.Recordset!idpalet)
   gravar_canvis
   dbtmp.Execute "update palets set datarec=null where idpalet=" + atrim(palet)
   dbtmp.Execute "update palets set disponible=false where idpalet=" + atrim(palet)
   'txtFields(9).Text = ""
   palets.Refresh
   palets.Recordset.FindFirst "idpalet=" + atrim(palet)
End Sub

Private Sub Command4_Click()
    Dim clonarxvegades As Double
    Dim rsta As Recordset
    Dim i As Long
    Dim j As Long
    
    If Not bobines.Recordset.EOF Then
      If bobines.Recordset.EditMode = 0 Then bobines.Recordset.Edit
       bobines.Recordset.Update
    End If
    
    If palets.Recordset.EditMode = 0 Then MsgBox "Primer passa a mode edició.": Exit Sub
    If bobines.Recordset.EditMode > 0 Then MsgBox "Estas editant aquesta bobina primer gravala": Exit Sub
    If bobines.Recordset.EOF Then Exit Sub
    clonarxvegades = cadbl(InputBox("Entra les copies que vols de la bobina " + atrim(bobines.Recordset!idbobina), "Copiar la bobina seleccionada", "1"))
    If clonarxvegades = 0 Then Exit Sub
    Set rsta = bobines.Recordset.Clone
    rsta.MoveLast
    For i = 1 To clonarxvegades
      bobines.Recordset.AddNew
      bobines.Recordset!idbobina = labobinamesgran + 1
      For j = 0 To bobines.Recordset.Fields.Count - 1
         If bobines.Recordset.Fields(j).Name <> "idbobina" Then
             bobines.Recordset.Fields(j) = rsta.Fields(j)
             possarvaloralareixa rsta, j
         End If
         bobines.Recordset!disponible = bobines.Recordset!mts
         DBGrid1.Columns("disponible") = cadbl(bobines.Recordset!mts)
      Next j
      bobines.Recordset.Update
    Next i
    bobines.Recordset.Bookmark = bobines.Recordset.LastModified
    Set rsta = Nothing
End Sub
Sub possarvaloralareixa(rsta As Recordset, j As Long)
   On Error Resume Next
   DBGrid1.Columns(rsta.Fields(j).Name) = rsta.Fields(j)
End Sub

Private Sub Command5_Click()
  Dim vdata As String
'  If txtFields(9) <> "" Then
    vdata = InputBox("Entra la nova data de recepció.", "Nova data")
    If Not IsDate(vdata) Then MsgBox "Aquesta data no es vàlida", vbCritical, "Error": Exit Sub
    dbtmp.Execute "update palets set datarec=#" + format(vdata, "mm/dd/yy") + "# where idpalet=" + atrim(cadbl(palets.Recordset!idpalet))
    palets.Recordset.Move 0
 ' End If
End Sub

Private Sub Command6_Click()
  Dim vcontador As Long
  Dim vlinia As String
  Dim i As Integer
  Dim j As Integer
'On Error GoTo errorcrearfitxer
   Open "c:\temp\~exportaciobobines.csv" For Output As #1
   vlinia = "Palet;Ample;Micres"
   For j = 1 To bobines.Recordset.Fields.Count - 1
            vlinia = vlinia + ";" + atrim(bobines.Recordset.Fields(j).Name)
   Next j
   On Error GoTo 0
   Print #1, vlinia
   If DBGrid1.SelBookmarks.Count > 0 Then
      vcontador = DBGrid1.SelBookmarks.Count
       Else:
         bobines.Refresh
         vcontador = bobines.Recordset.RecordCount
   End If
   For i = 0 To vcontador - 1
         If DBGrid1.SelBookmarks.Count > 0 Then DBGrid1.Bookmark = DBGrid1.SelBookmarks(i)
         vlinia = atrim(palets.Recordset!idpalet) + ";" + atrim(palets.Recordset!ample) + ";" + atrim(palets.Recordset!micres)
         For j = 1 To bobines.Recordset.Fields.Count - 1
            vlinia = vlinia + ";" + atrim(bobines.Recordset.Fields(j).Value)
         Next j
         Print #1, vlinia
         If DBGrid1.SelBookmarks.Count = 0 Then bobines.Recordset.MoveNext
         If bobines.Recordset.EOF Then GoTo fi
   Next i
fi:
   Close #1
   If existeix("c:\temp\~exportaciobobines.csv") Then obrir_document ("c:\temp\~exportaciobobines.csv")
Exit Sub
errorcrearfitxer:
   MsgBox "Error al crear el fitxer d'Excel, mira que no el tinguis obert i torna-ho a provar", vbCritical, "Error"
   Close #1
End Sub

Sub llistademodificacions_parcials()
  Dim vpa As Long
  Dim vbo As Long
  vpa = atrim(bobines.Recordset!idpalet)
  vbo = atrim(bobines.Recordset!idbobina)
  Unload formseleccio
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = palets.DatabaseName
  formseleccio.Data1.RecordSource = "select * from Parcials_controlcanvis where palet=" + atrim(vpa) + " and bobina=" + atrim(vbo) + " order by data "
  formseleccio.Width = 10000
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(4).Width = 1000
  formseleccio.DBGrid2.Columns(5).Width = 1600
  formseleccio.DBGrid2.Columns(6).Width = 1700
  formseleccio.DBGrid2.Columns(7).Width = 1700
  formseleccio.DBGrid2.Columns(8).Width = 1700
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Visible = False
  formseleccio.DBGrid2.Columns(2).Visible = False
  formseleccio.DBGrid2.Columns(3).Visible = False
  formseleccio.DBGrid2.Columns("data").NumberFormat = "dd/mm/yy hh:nn"
  
  
  If formseleccio.Data1.Recordset.EOF Then Exit Sub
  formseleccio.Show 1
  Unload formseleccio
End Sub

Private Sub Command7_Click()
  
End Sub
Function nomfitxer_fotoetiquetabobina(vbobina As String) As String
   Dim vrutafotos As String
   Dim vpalet As Double
   Dim vnomfitxer As String
   vrutafotos = llegir_ini("ruta", "ruta_etiquetes_bobinaproveidor", rutadelfitxer(cami) + "valorsprograma.ini")
   If Not existeix(vrutafotos) Then GoTo fi
   vpalet = cadbl(Mid(vbobina, 1, InStr(1, vbobina + " ", "/") - 1))
   If cadbl(vpalet) = 0 Then GoTo fi
   vrutafotos = rutadelfitxer(cami) + "cache_EtiquetesBobinesProveidor"
   vnomfitxer = vrutafotos + "\Els_" + atrim(atrim(Int(cadbl(vpalet) / 1000)) + "000") + "\" + substituir(vbobina, "/", "_") + ".jpg"
   If existeix(vnomfitxer) Then
           nomfitxer_fotoetiquetabobina = vnomfitxer
         Else
           vrutafotos = llegir_ini("ruta", "ruta_etiquetes_bobinaproveidor", rutadelfitxer(cami) + "valorsprograma.ini")
           vnomfitxer = vrutafotos + "\Els_" + atrim(atrim(Int(cadbl(vpalet) / 1000)) + "000") + "\" + substituir(vbobina, "/", "_") + ".jpg"
           If existeix(vnomfitxer) Then nomfitxer_fotoetiquetabobina = vnomfitxer
   End If
fi:
End Function

Function treuresimbolsnovalidsnomfitxer(desc As String) As String
   desc = substituir(desc, "\", "_")
   desc = substituir(desc, "/", "_")
   desc = substituir(desc, "|", "_")
   desc = substituir(desc, ":", ";")
   desc = substituir(desc, "?", "¿")
   desc = substituir(desc, "*", "x")
   desc = substituir(desc, """", "'")
   desc = substituir(desc, ">", "+")
   desc = substituir(desc, "<", "-")
   treuresimbolsnovalidsnomfitxer = desc
End Function

Private Sub Command8_Click()
   Dim v As String
   Dim vnomfitxer As String
   vnomfitxer = "\\ord_copies\Albarans_Proveidors\" + atrim(palets.Recordset!numalb) + " [" + atrim(proveidor.Tag) + "]"
   vnomfitxer = treuresimbolsnovalidsnomfitxer(vnomfitxer) + "*.pdf"
   v = Dir(vnomfitxer)
   If v <> "" Then obrir_document "\\ord_copies\Albarans_Proveidors\" + v
   
   
End Sub

Private Sub Command9_Click(Index As Integer)
   llistademodificacions_parcials
End Sub

Private Sub consultar_Click()
'  Dim palet As Double
'    palet = cadbl(InputBox("Entra el palet que busques", "Buscant palet"))
'  palets.Recordset.FindFirst "idpalet=" + atrim(palet)
Dim objecte As Object
  If palets.Recordset.EditMode > 0 Then MsgBox "Estas editant. Primer finalitza l'edicio.": Exit Sub
  buscant = True
  'alta_Click
  noupalet
  If palets.Recordset.EditMode > 0 Then
   
   txtFields(0).Locked = False
   txtFields(0) = ""
   txtFields(0).SetFocus
   txtFields(14) = ""
   For Each objecte In Me
     If TypeOf objecte Is CheckBox Then
       objecte = 2
     End If
   Next
    Else: buscant = False
  End If
End Sub
Sub finalitzarbusqueda(Optional tipus As Byte)
 ratoli "espera"
 
 If cadbl(tipus) = 1 Then GoTo ficonsulta
 recorregutregistres
 If palets.Recordset.EditMode > 0 Then palets.Recordset.CancelUpdate
ficonsulta:
 activarframes False
 buscant = False
 If queryorder <> "" Then
     queryorder = " Order By " + queryorder
    Else: queryorder = " order by idpalet desc"
 End If
 If querywhere <> "" Then querywhere = " Where " + querywhere
 palets.RecordSource = "select * from palets " + querywhere + queryorder
 palets.Refresh
 If Not palets.Recordset.EOF Then palets.Recordset.MoveLast: palets.Recordset.MoveFirst
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
     If objecte.Tag = "9" Or objecte <> "" Then
       If objecte.DataField <> "" Then
         If objecte <> "" Then
           evaluarcontingut objecte.DataField, atrim(objecte), palets.Recordset.Fields(objecte.DataField).Type
           objecte = ""
         End If
      End If
     End If
    End If
   If TypeOf objecte Is CheckBox Then
    If objecte.Value <> 2 And objecte.DataField <> "" Then
       If querywhere <> "" Then querywhere = querywhere + " and "
       querywhere = objecte.DataField + "=" + IIf(objecte.Value = 1, "True", "False")
    End If
   End If
   
 Next
'exepcions


   
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
    rest = rest + "#" + format(Mid(valor, i, 50), "d/m/yyyy") + "#"
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

Private Sub DBGrid1_AfterDelete()
sumarmetresikilos
End Sub

Private Sub DBGrid1_AfterUpdate()
  sumarmetresikilos
End Sub
Sub sumarmetresikilos()
   Dim rsts As Recordset
   Set rsts = bobines.Database.OpenRecordset("select sum(mts) as sumametres,sum(kilos) as sumakilost,sum(pesdelproveidor) as sumakilos from bobines where idpalet=" + atrim(cadbl(palets.Recordset!idpalet)))
   If Not rsts.EOF Then
      
      totalmetres = format(cadbl(rsts!sumametres), "#,##0") + " Mtrs"
      totalkilos = format(cadbl(rsts!sumakilos), "#,##0") + " Kg"
      totalkilosteorics = format(cadbl(rsts!sumakilost), "#,##0") + " Kg"
   End If
   Set rsts = Nothing
End Sub
Private Sub DBGrid1_BeforeUpdate(Cancel As Integer)
  Dim kilos As Double
  Dim metres As Double
  metres = cadbl(DBGrid1.Columns("mts"))
  If metres <= 0 Then bobines.Recordset.CancelUpdate: MsgBox "Els metres de bobina no poden ser zero o negatius.", vbCritical + vbOKOnly, "Atenció": Exit Sub
  kilos = compramat.conversiokilos(palets.Recordset!codimatprognou, palets.Recordset!ample, metres, cadbl(palets.Recordset!micres), atrim(palets.Recordset!semielaborat), cadbl(palets.Recordset!solapa))
  DBGrid1.Columns("kilos") = format(kilos, "#,##0")
End Sub

'--------------------------
Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim metres As Double
   Dim valoranterior As Double
   If DBGrid1.Columns(DBGrid1.col).DataField = "Mts" And metrespartits > 0 Then
     valoranterior = cadbl(DBGrid1.Columns("Mts"))
      KeyCode = 0
        DBGrid1.Columns("Mts") = atrim(valoranterior)
        metres = cadbl(InputBox("Entra els metres a modificar", "Rectificacio de metres"))
        If metres >= metrespartits Then
          DBGrid1.Columns("Mts") = atrim(metres)
            Else: MsgBox ("Els metres assignats son superiors al entrats, rectifica primer la assignació"): Exit Sub
        End If
      KeyCode = 0
   End If
End Sub

Private Sub DBGrid1_OnAddNew()
  novabobina
End Sub
Sub novabobina()
  DBGrid1.Columns("idpalet") = palets.Recordset!idpalet
  
  DBGrid1.Columns("idbobina") = labobinamesgran + 1
  
End Sub
Function labobinamesgran() As Long
  Dim clonebob As Recordset
  Dim gran As Long
  gran = 0
  Set clonebob = bobines.Recordset.Clone
  If Not clonebob.EOF Then
    clonebob.MoveLast
    gran = cadbl(clonebob!idbobina)
  End If
  labobinamesgran = gran

End Function
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If DBGrid1.Columns(DBGrid1.col).DataField = "Mts" Then
    If metrespartits = 0 Then
      DBGrid1.Columns(DBGrid1.col).Locked = False
     Else: DBGrid1.Columns(DBGrid1.col).Locked = True
    End If
  End If
 If DBGrid1.Columns(DBGrid1.col).DataField <> "Sit" Then
  DBGrid1.SelStart = Len(DBGrid1.Text)
  DBGrid1.SelLength = 0
 End If
End Sub

Private Sub eliminar_Click()
  If palets.Recordset.EOF Then Exit Sub
  If UCase(InputBox("Eliminar aquest palet implica eliminar totes les seves bobines." + Chr(13) + Chr(10) + " ESCRIU [Eliminar palet] SI ESTAS SEGUR QUE HO VOLS FER.", "ATENCIO")) = UCase("Eliminar palet") Then
     If Not bobines.Recordset.EOF Then
      bobines.Recordset.MoveFirst
      While Not bobines.Recordset.EOF
        bobines.Recordset.Delete
        bobines.Recordset.MoveNext
      Wend
     End If
     palets.Recordset.Delete
     palets.Recordset.MovePrevious
  End If
End Sub

Private Sub escullirimpetbob_Click()
'Seleccionar impresora
 Dim nomimpresora As String
 Dim novaimpresora As String
 Dim objprinter As Printer
 If MsgBox("Estas segur que vols canviar la impresora per imprimir etiquetes de palet?", vbDefaultButton2 + vbYesNo + vbCritical, "Atenció") = vbNo Then Exit Sub
 
 novaimpresora = llegir_ini("Expedicions", "nomimpresoraetiquetes", fitxerini)
 If novaimpresora = "{[}]" Then novaimpresora = ""
 For Each objprinter In Printers
    resp = MsgBox(objprinter.DeviceName, vbInformation + vbYesNo, "Vols escullir aquesta impresora?")
    If resp = vbYes Then
      novaimpresora = objprinter.DeviceName
      GoTo fi
      
    End If
 Next
fi:
escriure_ini "Expedicions", "nomimpresoraetiquetes", novaimpresora, fitxerini
End Sub

Private Sub Form_Activate()
  If vfiltrebobinesdesdeimpresores Then Exit Sub
  If vexportantelllistat Then
     Form1.Visible = False
     invtotselsforats_Click
     End
  End If
  comprovarpaletssensematerialassignat
  actualitzar_vinculats
  comprovar_nodisponibles
  If palets.Recordset.EditMode = 1 Then
    Form1.DBGrid1.SetFocus
    Form1.DBGrid1.col = 2
  End If
  
  
End Sub
Sub comprovarpaletssensematerialassignat()
   Dim rstp As Recordset
   Dim paletsnuls As String
   If jaheentratavui Then Exit Sub
   Set rstp = palets.Database.OpenRecordset("select idpalet from palets where codimatprognou=null")
   While Not rstp.EOF
      paletsnuls = paletsnuls + format(rstp!idpalet, "#,##0") + " "
      rstp.MoveNext
   Wend
   If paletsnuls <> "" Then MsgBox "S'han trobat palets sense material assignat," + Chr(10) + "revisals i eliminals si cal." + Chr(10) + paletsnuls, vbCritical, "Atenció"
End Sub
Sub comprovar_nodisponibles()
   If Not jaheentratavui Then
     comprovarpaletsnodisponibles
   End If
End Sub
Function jaheentratavui() As Boolean
   If llegir_ini("Palets", "uncopaldianodisponibles", fitxerini) = Date Then
      jaheentratavui = True
        Else
          jaheentratavui = False
          escriure_ini "Palets", "uncopaldianodisponibles", atrim(Date), fitxerini
   End If
End Function
Private Sub Form_Click()

  'comprespalets.traspasdetotlarticleaSAPaunfitxerapartCSV 679, "segon article material", "M", "c:\temp\exportar.csv"
'imprimirllistatreferencies
 ' Dim rstp As Recordset
 '  Set dbstocks = OpenDatabase(camistock)
 ' Set rstp = dbstocks.OpenRecordset("select * from parcials where id in (SELECT parcials.id FROM comandes INNER JOIN Parcials ON comandes.comanda = cdbl(Parcials.comanda) WHERE (((CDbl(parcials.comanda))>120000) AND ((Parcials.utilitzada)=False) AND ((Parcials.metres)>0) and comandes.proximaseccio='T');)")
 ' rstp.MoveLast
 ' rstp.MoveFirst
 '
 ' While Not rstp.EOF
 '
 '   rstp.Edit
  '  rstp!operari = 0
  '  rstp!seccio = "S"
  '  rstp!data = "31/12/2013"
  '  rstp!utilitzada = True
  '  rstp.Update
  '  rstp.MoveNext
  'Wend
End Sub
Sub comprovarpaletsnodisponibles()
     Dim consulta As String
     Dim rstc As Recordset
     Dim vdatarev As Date
     Dim datarecepcio As String
      'consulta = "SELECT nom , Datarev , Numalb  FROM Palets INNER JOIN (materials INNER JOIN proveidors ON materials.proveidor = proveidors.codi) ON Palets.codimatprognou = materials.codi WHERE Disponible=False"
      consulta = "SELECT First(nom) as nomproveidor , First(Datarev) as vdatarev , Numalb FROM Palets INNER JOIN (materials LEFT JOIN proveidors ON materials.proveidor = proveidors.codi) ON Palets.codimatprognou = materials.codi Where Disponible = False GROUP BY Palets.Numalb"
      Set rstc = dbtmp.OpenRecordset(consulta)
      If rstc.EOF Then Set rstc = Nothing: Exit Sub
      Load formseleccio
      formseleccio.Caption = "Palets marcats com a no disponibles"
      formseleccio.Data1.DatabaseName = camistock
      ordre = " order by Proveidor,data_alb"
      formseleccio.Data1.RecordSource = consulta
      formseleccio.refrescar
      formseleccio.Width = formseleccio.Width + ((formseleccio.Width / 100) * 20)
      formseleccio.DBGrid2.Columns(0).Width = 2500
      formseleccio.DBGrid2.Columns(0).Caption = "Proveidor"
      formseleccio.DBGrid2.Columns(1).Width = 1000
      formseleccio.DBGrid2.Columns(1).Caption = "Data Alb."
      formseleccio.DBGrid2.Columns(2).Width = 1500
      formseleccio.DBGrid2.Columns(2).Caption = "Num Alb."
      'formseleccio.DBGrid2.Columns(3).Width = 1500
      formseleccio.Command2.Tag = "2"
      formseleccio.Show 1
      noucodimat = 0
      If seleccioret = 1 Then
        While Not IsDate(datarecepcio)
demanardata:
           datarecepcio = InputBox("Entra la data de recepció del material" + Chr(10) + "Escriu [Sortir] per cancelar." + Chr(10) + "Ex: " + format(Now, "dd/mm/yy"), "Data Recepció", format(Now, "dd") + "/" + format(Now, "mm") + "/" + format(Now, "yy"))
           If UCase(datarecepcio) = "SORTIR" Then Exit Sub
           If Not IsDate(datarecepcio) Then
              MsgBox "Data incorrecte.", vbCritical, "Error"
               Else
                 If DateDiff("d", Now, datarecepcio) > 0 Then
                    MsgBox "No pots posar un data de recepció que encara no ha arribat.", vbCritical, "Error"
                    datarecepcio = 0
                 End If
           End If
         Wend
         If Month(datarecepcio) <> Month(Now) Then
             If MsgBox("Has possat una data de recepció d'un mes diferent del que estem aixó significa un canvi en el càlcul d'existencies i s'haurà d'avisar a COMPTABILITAT." + vbNewLine + "Estàs segur que es la data que vols possar?", vbCritical + vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbNo Then
                     GoTo demanardata
                      Else: enviaremailgeneric "miquel.inplacsa@gmail.com", "ATENCIÓ!!! PALETS ARRIBATS EN EL CANVI DE MES, REVISAR-HO PER EXISTENCIES.", "Proveidor: " + atrim(formseleccio.Data1.Recordset!nomproveidor) + vbNewLine + "Revisar l'albarà de proveidor " + atrim(formseleccio.Data1.Recordset!numalb)
             End If
         End If
         ratoli "espera"
         vdatarev = CVDate(formseleccio.Data1.Recordset!vdatarev)
         dbtmp.Execute "update palets set disponible=true,datarec=#" + format(datarecepcio, "mm/dd/yy") + "# where numalb='" + atrim(formseleccio.Data1.Recordset!numalb) + "' and datarev=#" + atrim(format(vdatarev, "mm/dd/yy")) + "#"
         passarcomandesarebudes formseleccio.Data1.Recordset!numalb, vdatarev
         avisarramonsidataalbaradelmespassat formseleccio.Data1.Recordset!numalb
         palets.RecordSource = "select * from palets where numalb='" + atrim(formseleccio.Data1.Recordset!numalb) + "'"
         palets.Refresh
         ratoli "normal"
      End If
      Unload formseleccio
   
End Sub
Sub comprovarpaletssensedatadactivacio()
     Dim consulta As String
     Dim rstc As Recordset
     Dim vdatarev As Date
     Dim dataactivacio As String
     Dim dataprevista As String
      'consulta = "SELECT nom , Datarev , Numalb  FROM Palets INNER JOIN (materials INNER JOIN proveidors ON materials.proveidor = proveidors.codi) ON Palets.codimatprognou = materials.codi WHERE Disponible=False"
      consulta = "SELECT First(nom) , First(Datarev) as vdatarev , Numalb FROM Palets INNER JOIN (materials LEFT JOIN proveidors ON materials.proveidor = proveidors.codi) ON Palets.codimatprognou = materials.codi Where not disponible and dataactivacio is null GROUP BY Palets.Numalb"
      Set rstc = dbtmp.OpenRecordset(consulta)
      If rstc.EOF Then Set rstc = Nothing: Exit Sub
      Load formseleccio
      formseleccio.Caption = "Palets sense data d'activació"
      formseleccio.Data1.DatabaseName = camistock
      ordre = " order by Proveidor,data_alb"
      formseleccio.Data1.RecordSource = consulta
      formseleccio.refrescar
      formseleccio.Width = formseleccio.Width + ((formseleccio.Width / 100) * 30)
      formseleccio.DBGrid2.Columns(0).Width = 2500
      formseleccio.DBGrid2.Columns(0).Caption = "Proveidor"
      formseleccio.DBGrid2.Columns(1).Width = 1000
      formseleccio.DBGrid2.Columns(1).Caption = "Data Alb."
      formseleccio.DBGrid2.Columns(2).Width = 1500
      formseleccio.DBGrid2.Columns(2).Caption = "Num Alb."
      'formseleccio.DBGrid2.Columns(3).Width = 1500
      formseleccio.Command2.Tag = "2"
      formseleccio.Show 1
      noucodimat = 0
      If seleccioret = 1 Then
        While Not IsDate(dataactivacio)
           dataactivacio = InputBox("Entra la data d'activació del material" + Chr(10) + "Escriu [Sortir] per cancelar." + Chr(10) + "Ex: " + format(Now, "dd/mm/yy"), "Data Activació", format(Now, "dd") + "/" + format(Now, "mm") + "/" + format(Now, "yy"))
           If UCase(dataactivacio) = "SORTIR" Then Exit Sub
           If Not IsDate(dataactivacio) Then MsgBox "Data incorrecte.", vbCritical, "Error"
         Wend
         While Not IsDate(dataprevista)
           dataprevista = InputBox("Entra la data prevista d'arribada del material" + Chr(10) + "Escriu [Sortir] per cancelar." + Chr(10) + "Ex: " + format(Now, "dd/mm/yy"), "Data prevista d'arribada", format(Now, "dd") + "/" + format(Now, "mm") + "/" + format(Now, "yy"))
           If UCase(dataprevista) = "SORTIR" Then Exit Sub
           If Not IsDate(dataprevista) Then MsgBox "Data incorrecte.", vbCritical, "Error"
         Wend
         vdatarev = CVDate(formseleccio.Data1.Recordset!vdatarev)
         'dbtmp.Execute "update palets set disponible=true,datarec=#" + Format(datarecepcio, "mm/dd/yy") + "# where numalb='" + atrim(formseleccio.Data1.Recordset!numalb) + "' and datarev=#" + atrim(Format(vdatarev, "mm/dd/yy")) + "#"
         dbtmp.Execute "update palets set dataactivacio=#" + format(dataactivacio, "mm/dd/yy") + "#,Dataprevistaarribada=#" + format(dataprevista, "mm/dd/yy") + "# where numalb='" + atrim(formseleccio.Data1.Recordset!numalb) + "' and datarev=#" + atrim(format(vdatarev, "mm/dd/yy")) + "#"
        ' passarcomandesarebudes formseleccio.Data1.Recordset!numalb, vdatarev
        ' avisarramonsidataalbaradelmespassat formseleccio.Data1.Recordset!numalb
         palets.RecordSource = "select * from palets where numalb='" + atrim(formseleccio.Data1.Recordset!numalb) + "'"
         palets.Refresh
      End If
      Unload formseleccio
   
End Sub


Sub avisarramonsidataalbaradelmespassat(numalb As String)
   Dim rstalb As Recordset
   Dim llpalets As String
   Set rstalb = dbtmp.OpenRecordset("select * from palets where numalb='" + atrim(numalb) + "'")
   While Not rstalb.EOF
    If IsDate(rstalb!dataaltapalet) And IsDate(rstalb!datarev) And IsDate(rstalb!datarec) Then
     If Month(rstalb!dataaltapalet) = Month(Now) Then
         If Month(rstalb!datarev) = Month(DateAdd("m", -1, Now)) Then
            If Month(rstalb!datarec) = Month(Now) Then
               llpalets = llpalets + " | " + atrim(rstalb!idpalet)
            End If
         End If
     End If
    End If
    rstalb.MoveNext
   Wend
   If llpalets <> "" Then MsgBox "Els següents palets s'han activat aquest mes però l'albarà es del mes passat s'hauria d'avisar a en RAMON." + Chr(10) + llpalets, vbCritical, "Atenció"
   Set rstalb = Nothing
End Sub
Sub passarcomandesarebudes(numalb As String, vdatarev As Date)
   Dim rstalb As Recordset
   Dim rstc As Recordset
   'Set rstcompres = dbcompres.OpenRecordset("SELECT capcalera.numcomanda,capcalera.materialrebut, capcalera.data, capcalera.dataentrega, capcalera.codiproveidorcomercial,capcalera.empresa,capcalera.nomprov, liniescompra.* FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra where numcomanda=" + atrim(cadbl(comandacompra)) + ";")
   Set rstalb = dbcompres.OpenRecordset("select * from albaransbip where numalbaraprov='" + atrim(numalb) + "'" + " and data=#" + atrim(format(vdatarev, "mm/dd/yy")) + "#")
   If Not rstalb.EOF Then
       While Not rstalb.EOF
          Set rstc = dbcompres.OpenRecordset("SELECT capcalera.numcomanda as elnumdecomanda, liniescompra.* FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra Where idliniacompra = " + atrim(cadbl(rstalb!idliniacompra)))
          If Not rstc.EOF Then
            If rstc!totenviat Then
              rstc.Edit: rstc!totentregat = True: rstc.Update
              passarliniadecompraarebudes rstc!idliniacompra
            End If
               Else: MsgBox "No trobo la compra " + atrim(rstc!elnumdecomanda) + " amb data " + atrim(vdatarev) + vbNewLine + "La compra no quedarà com a Entregada TOTAL.", vbCritical, "Error"
          End If
          comprespalets.reservarlescomandesassociades rstc
          rstalb.MoveNext
       Wend
          Else: MsgBox "No he trobat l'albarà " + atrim(numalb) + " amd data " + atrim(vdatarev) + " les compres respectives no passaran a REBUDES.", vbCritical, "Error"
   End If
End Sub

Sub passarliniadecompraarebudes(idlinia As Double)
   Dim rstm As Recordset
   Dim rstt As Recordset
   Dim rstc As Recordset
   Dim numc As Double
   numc = 0
   Set rstc = dbcompres.OpenRecordset("SELECT capcalera.numcomanda, liniescompra.idliniacompra FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra where idliniacompra=" + atrim(cadbl(idlinia)))
   If Not rstc.EOF Then numc = cadbl(rstc!numcomanda) Else GoTo fi
   dbcompres.Execute "update liniescompra set totentregat=true where idliniacompra=" + atrim(idlinia)
   Set rstm = dbcompres.OpenRecordset("SELECT capcalera.numcomanda, liniescompra.totentregat FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE (((capcalera.numcomanda)=" + atrim(numc) + ") AND ((liniescompra.totentregat)=True));")
   Set rstt = dbcompres.OpenRecordset("SELECT capcalera.numcomanda, liniescompra.totentregat FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE (((capcalera.numcomanda)=" + atrim(numc) + ") );")
   If rstm.EOF Or rstt.EOF Then GoTo fi
   rstm.MoveLast
   rstt.MoveLast
   If rstm.RecordCount = rstt.RecordCount Then dbcompres.Execute "update capcalera set materialrebut=true where numcomanda=" + atrim(numc)
fi:
   Set rstm = Nothing
   Set rstt = Nothing
   Set rstc = Nothing
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 112 Then gravar_canvis
  If KeyCode = 27 Then cancelar_canvis
  If KeyCode = 13 Then SendKeys "{TAB}": KeyCode = 0
  If KeyCode = 48 And Shift = 2 Then consultar_Click
End Sub
Sub cancelar_canvis()
  If palets.Recordset.EditMode > 0 Then
    palets.Recordset.CancelUpdate
    If bobines.Recordset.EditMode > 0 Then bobines.Recordset.CancelUpdate
    If parcials.Recordset.EditMode > 0 Then parcials.Recordset.CancelUpdate
    DBGrid1.EditActive = False
    reixaparcials.EditActive = False
    txtFields(0).Locked = True
    activarframes False
    palets.Refresh
  End If
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



Private Sub Image1_Click()

End Sub


Sub mirarsihihalafontTTFdecodidebarres()
   Dim objshell As Variant
   Dim objFolderItem As Variant
   If existeix("c:\windows\fonts\free3of9.ttf") Then Exit Sub
   Copiar_Fitxer llegir_ini("General", "rutallistats", fitxerini) + "\free3of9.ttf", "c:\windows\fonts"
   'Set objshell = CreateObject("Shell.Application")
   'Set objFolder = objshell.Namespace("C:\windows\Fonts")
   'Set objFolderItem = objFolder.ParseName("free3of9.ttf")
   'objFolderItem.InvokeVerb ("Install")
End Sub

Private Sub Form_Load()
  palets.DatabaseName = camistock
  ensenyar_camps
  DoEvents
  bobines.DatabaseName = palets.DatabaseName
  parcials.DatabaseName = palets.DatabaseName
  palets.RecordSource = "select * from Palets order by idpalet ASC"
  
  ' poso l'ordre al rebes perque si s'està entrant palets es col.loca al registre actiu i tarda un rato a reaccionar
  palets.Refresh
 Set dbcompres = DBEngine.OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
 Set dbcomandes = OpenDatabase(cami)
  'If Not palets.Recordset.EOF Then palets.Recordset.MoveLast: palets.Recordset.MoveFirst
  DoEvents
  If llegir_ini("Compres", "numempresabip_inplacsa", "comandes.ini") = "{[}]" Then
      escriure_ini "Compres", "numempresabip_inplacsa", 2, "comandes.ini"
      escriure_ini "Compres", "numempresabip_plasel", 6, "comandes.ini"
  End If
  mirarsihihalafontTTFdecodidebarres
End Sub
Sub ensenyar_camps()
   Dim i As Byte
   Dim camp As String
   camp = ""
   i = 1
   While camp <> "{[}]"
      camp = llegir_ini("Palets", "ensenyarcamps" + atrim(i), fitxerini)
      If atrim(camp) <> "" And atrim(camp) <> "{[}]" Then
        ensenya camp
      End If
      i = i + 1
   Wend
   If i = 2 And camp = "{[}]" Then escriure_ini "Palets", "ensenyarcamps1", ";preucompra", fitxerini
End Sub
Sub ensenya(c As String)
  Dim Control As Object
  On Error Resume Next
  For Each Control In Form1
     If Control.DataField <> c Then
        Resume Next
       Else: Control.Visible = True
     End If
  Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Frame3_Click()
'ajustar_picos_alesbobin
End Sub

Private Sub Frame3_DblClick()
 Dim rst As Recordset
 Dim dbtmp2 As Database
 Dim rstp As Recordset
 Dim vcontadorimpresos As Double
 Exit Sub
 If MsgBox("Vols imprimir tots els parcials?", vbCritical + vbYesNo, "Atenció") = vbNo Then Exit Sub
 Set dbtmp2 = OpenDatabase(nomfitxertemporal) '("c:\temporal.mdb")
 Set rst = dbtmp2.OpenRecordset("select lletralleixa,lloclleixa,foratlleixa from llistatinventari where lletralleixa='C' or lletralleixa='D' or lletralleixa='E' order by lletralleixa,lloclleixa,foratlleixa")
 Set dbstocks = dbtmp
 While Not rst.EOF
   'MsgBox atrim(rst!lletralleixa) + atrim(rst!lloclleixa) + atrim(rst!foratlleixa)
   Set rstp = dbtmp.OpenRecordset("select * from bobines where sit='" + atrim(rst!lletralleixa) + atrim(rst!lloclleixa) + atrim(rst!foratlleixa) + "'")
   While Not rstp.EOF
        vcontadorimpresos = vcontadorimpresos + 1
        Me.Caption = "Imprimint " + atrim(rstp!idpalet) + "-" + atrim(rstp!idbobina)
        bobinesdentrada.imprimir_bobinaparcial rstp!idpalet, rstp!idbobina, True
        wait 1
        rstp.MoveNext
   Wend
   rst.MoveNext
 Wend
 MsgBox atrim(vcontadorimpresos) + " Fulls impresos"
 Set rst = Nothing
 Set dbtmp2 = Nothing
 Set dbstocks = Nothing
End Sub

Private Sub hitoricpackinglist_Click()
Dim numcomanda As Double
  numcomanda = cadbl(InputBox("Entra la Comanda que vols imprimir l'HISTORIC DEL PACKING-LIST", "Packing-List"))
  If numcomanda > 0 Then imprimir_packinglist numcomanda, llistat, False, "historic_packinglist"
End Sub

Private Sub impbobdesdefins_Click()
   Dim paletinici As Double
   Dim paletfi As Double
   paletinici = cadbl(InputBox("Escriu el palet d'Inici.", "Palet Inici", atrim(palets.Recordset!idpalet)))
   paletfi = cadbl(InputBox("Escriu el palet de Fi.", "Palet Fi", atrim(palets.Recordset!idpalet)))
   If paletinici > 0 And paletfi > 0 Then imprimiretpalet paletinici, paletfi
End Sub

Private Sub imprimirparcial_Click()
  Set dbstocks = dbtmp
  bobinesdentrada.imprimir_bobinaparcial bobines.Recordset!idpalet, bobines.Recordset!idbobina, True
  Set dbstocks = Nothing
End Sub

Private Sub infocompra_Click()
  Dim rstc As Recordset
  Dim rstprov As Recordset
  Dim vnumpalet As Double
  Dim v As String
  vnumpalet = cadbl(palets.Recordset!idpalet)
  If cadbl(palets.Recordset!link_numpalet) > 0 Then vnumpalet = cadbl(palets.Recordset!link_numpalet)
  Set rstc = dbcompres.OpenRecordset("select * from albaransbip where numpalet=" + atrim(vnumpalet), , ReadOnly)
  If rstc.EOF Then MsgBox "No hi ha informació de compra per aquest palet.", vbInformation, "Informació Compra"
  While Not rstc.EOF
    Set rstprov = dbcompres.OpenRecordset("SELECT proveidors.tipusproveidorIMPOST, proveidors_comercial.codicomptable FROM proveidors_comercial LEFT JOIN proveidors ON proveidors_comercial.codi = proveidors.codi WHERE (((proveidors_comercial.codicomptable)='" + atrim(rstc!codiproveidorcomercial) + "'));")
    v = "Comanda: " + atrim(rstc!numcomanda) + Chr(10) + "Data Entrada: " + atrim(rstc!data) + Chr(10) + "Descripcio: " + atrim(rstc!descripcio) + Chr(10) + "Kg entregats:" + atrim(rstc!quantitat) + Chr(10) + "Preu: " + atrim(rstc!preu) + vbNewLine
    If (cadbl(rstc!kgimpostenvasos)) > 0 Then v = v + "Kg Impost Envasos: " + atrim(rstc!kgimpostenvasos) + " X " + atrim(rstc!preuImpostEnvasos) + "€" + IIf(Not rstprov.EOF, " [" + atrim(rstprov!tipusproveidorIMPOST) + "]", "")
    If cadbl(palets.Recordset!link_numpalet) > 0 Then v = "ATENCIÓ BOBINA PARTIDA." + vbNewLine + "================" + vbNewLine + vbNewLine + v
    MsgBox v
    rstc.MoveNext
  Wend
  Set rstc = Nothing
End Sub

Private Sub inventariimpanonim_Click()
  Dim rstcom As Recordset
  Dim consulta As String
  Dim Grups As String
  Dim grups2 As String
  Dim sqlgrups2
   borrartaulallistatinventari
   buscarcomandesafabricaasignades
   creartaulatempllistat
   Grups = " select comanda from assignadesafabrica where seccio='I' and producte<>'PC'"
   'ELS GRUPS2 NOMES QUAN ES LAMINADORA
    'grups2 = "select trim(linkcomanda1) from assignadesafabrica where linkcomanda1>0 and seccio='L' and secciolink1='V'"
   'posso els palets i bobines al llistat
   If grups2 <> "" Then sqlgrups2 = "OR comanda in(" + grups2 + ")"
   'posso primer tot menys els laminats
   consulta = "Insert into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
   consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,parcials.metres as metres,palets.micres as espesor"
   consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (parcials.comanda in (" + Grups + "));" ' + sqlgrup2 + ");"
   dbtmp.Execute consulta
   'posso ara els laminats PER SUPUSAT QUAN FAIG REBOBINADORA,QUE ES QUAN CAL, SI NO NO CAL
   If grups2 <> "" Then
    consulta = "Insert into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
    consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,parcials.comanda as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,parcials.metres as metres,palets.micres as espesor"
    consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (parcials.comanda in (" + sqlgrups2 + "));"
   End If
   
   passoelscodimatlaminatsamatfulla1
   passodemetresakilos
   possoelsnomsalsmaterials
   
   
   Set rstcom = Nothing
   
  'faig el llistat
  llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistat inventari.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 llistat.Formulas(0) = "nomllistat='Llistat inventari a Impresores (Anònim)'"
 llistat.Formulas(1) = "hora='" + format(Now, "dd/mm/yy  hh:nn") + "'"
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 dbllistat.Close
 obrir_dbllistats
  
  
End Sub
Sub passoelscodimatlaminatsamatfulla1()
  Dim rstc As Recordset
  Dim rstcom As Recordset
  Dim rstmat As Recordset
  Dim petit As Double
  Set rstc = dbllistat.OpenRecordset("SELECT * from llistatinventari where trim(comanda)<>''")
  While Not rstc.EOF
    Set rstcom = dbtmpb.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(rstc!comanda))
    If Not rstcom.EOF Then
       petit = cadbl(rstcom!comanda)
       If cadbl(rstcom!linkcomanda1) < petit And cadbl(rstcom!linkcomanda1) > 0 Then petit = cadbl(rstcom!linkcomanda1)
       If cadbl(rstcom!linkcomanda2) < petit And cadbl(rstcom!linkcomanda2) > 0 Then petit = cadbl(rstcom!linkcomanda2)
       Set rstmat = dbtmpb.OpenRecordset("select materialex from comandes where comanda=" + atrim(petit))
       If Not rstmat.EOF Then
          rstc.Edit
           rstc!codimat = rstmat!materialex
          rstc.Update
       End If
    End If
    rstc.MoveNext
  Wend
  Set rstc = Nothing
  Set rstmat = Nothing
  Set rstcom = Nothing
End Sub
Sub creartaulatempllistatbobinessensemoviment()
  Dim consulta As String
   consulta = "SELECT Palets.Idpalet, Bobines.Idbobina, Palets.Ample, Palets.micres, materials.descripcio, bobines.kilos,bobines.mts,Bobines.disponible, Bobines.Sit, Palets.dataaltapalet"
   consulta = consulta + " into llistatinventari in '" + nomfitxertemporal + "' FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi"
   consulta = consulta + " WHERE (((Bobines.disponible)>0));"
   dbtmp.Execute consulta
End Sub

Sub creartaulatempllistat()
  Dim consulta As String
   
 'faig aquesta consulta per crear la mateixa taula que a l'inventari disponible i poder fer servir el mateix llistat
   consulta = "SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,bobines.disponible as metres,palets.micres as espesor, palets.preucompra as preucompra,'' as producte, '' as descripcio,'' as datacomanda, '' as fabricant "
   consulta = consulta + " into llistatinventari in '" + nomfitxertemporal + "'"
   consulta = consulta + " FROM Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet where palets.disponible and bobines.disponible>0;"
   dbtmp.Execute consulta
   dbllistat.Execute "delete * from llistatinventari"
End Sub
Sub treurenulsdellinkcomanda(rstp As Recordset)
  Set rstp = dbtmp.OpenRecordset("SELECT * From comandesamblinkcomandanull;")
  While Not rstp.EOF
     If IsNull(rstp!linkcomanda1) Then
        rstp.Edit
        rstp!linkcomanda1 = 0
        rstp.Update
          Else: rstp.MoveLast
     End If
     rstp.MoveNext
  Wend
  Set rstp = dbtmp.OpenRecordset("select linkcomanda2 from comandes  order by linkcomanda2 desc")
  While Not rstp.EOF
     If IsNull(rstp!linkcomanda2) Then
        rstp.Edit
        rstp!linkcomanda1 = 0
        rstp.Update
          Else: rstp.MoveLast
     End If
     rstp.MoveNext
  Wend
  Workspaces(1).CommitTrans
End Sub
Sub buscarcomandesafabricaasignades()
 Dim csql As String * 1000
 Dim rstp As Recordset
 Dim csql2 As String
  'treurenulsdellinkcomanda rstp
     
  'If Not rstp.EOF Then
   'dbtmpb.Execute "update comandes set linkcomanda2=0 where linkcomanda2=null"
   'dbtmpb.Execute "update comandes set linkcomanda1=0 where linkcomanda1=null"
  'End If
  Form1.Caption = "Manteniment de Palets (Triant bobines assignades)"
  csql = "SELECT DISTINCT Parcials.comanda, comandes.producte, productes.ruta, comandes.proximaseccio AS Seccio, comandes_1.proximaseccio AS SeccioLink1, comandes_2.proximaseccio AS Secciolink2 ,comandes_2.proximaseccio as seccioanterior,comandes.linkcomanda1,comandes.linkcomanda2,comandes.materialex,comandes.datacomanda "
  csql = Trim(csql) + " into assignadesafabrica " + "FROM (Palets INNER JOIN (((comandes INNER JOIN productes ON comandes.producte = productes.codi) INNER JOIN comandes AS comandes_1 ON comandes.linkcomanda1 = comandes_1.comanda) INNER JOIN (Bobines INNER JOIN Parcials ON (Bobines.Idpalet = Parcials.idpalet) AND (Bobines.Idbobina = Parcials.idbobina)) ON trim(comandes.comanda) = Parcials.comanda) ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN comandes AS comandes_2 ON comandes.linkcomanda2 = comandes_2.comanda " 'WHERE ((comandes.proximaseccio <>'P' and comandes.proximaseccio <>'T' and comandes.proximaseccio <>'V') AND ((Palets.Disponible)<>False)) order by proximaseccio,parcials.comanda;"
  csql2 = " Where ((( comandes.proximaseccio) <> 'V'  And (comandes.proximaseccio) <> 'T' And (comandes.proximaseccio) <> 'P')  and ((palets.Disponible) <> False)) " 'or (comandes.producte='PC' and comandes.proximaseccio='V' and (comandes_1.proximaseccio<>'T' and comandes_1.proximaseccio<>'V' and comandes_1.proximaseccio<>'P')
  csql2 = Trim(csql2) + " ORDER BY comandes.proximaseccio, Parcials.comanda; "
  
 borrartaulaassignades
    dbtmp.Execute Trim(csql + csql2)
    Form1.Caption = "Manteniment de Palets (Afegint camp posicioruta)"
    dbtmp.Execute "alter table assignadesafabrica add column posicioruta byte"
    crearindextaula
    Form1.Caption = "Manteniment de Palets (Col.locant secció anterior)"
 colocarseccioanterior
 
End Sub
Sub crearindextaula()
 
  dbtmp.Execute "CREATE INDEX 1 ON assignadesafabrica (comanda);"
  dbtmp.Execute "CREATE INDEX 2 ON assignadesafabrica (producte);"
  dbtmp.Execute "CREATE INDEX 3 ON assignadesafabrica (ruta);"
  dbtmp.Execute "CREATE INDEX 4 ON assignadesafabrica (seccio);"
  dbtmp.Execute "CREATE INDEX 5 ON assignadesafabrica (secciolink1);"
  dbtmp.Execute "CREATE INDEX 6 ON assignadesafabrica (secciolink2);"
  dbtmp.Execute "CREATE INDEX 7 ON assignadesafabrica (seccioanterior);"
  dbtmp.Execute "CREATE INDEX 8 ON assignadesafabrica (linkcomanda1);"
  dbtmp.Execute "CREATE INDEX 9 ON assignadesafabrica (linkcomanda2);"
  dbtmp.Execute "CREATE INDEX 106 ON assignadesafabrica (materialex);"
  dbtmp.Execute "CREATE INDEX 18 ON assignadesafabrica (datacomanda);"
  dbtmp.Execute "CREATE INDEX 19 ON assignadesafabrica (posicioruta);"
  
End Sub
Sub colocarseccioanterior()
  Dim rstc As Recordset
  Dim sec As String
  Dim posicio As Byte
  Set rstc = dbtmp.OpenRecordset("select * from assignadesafabrica")
  While Not rstc.EOF
    sec = rstc!ruta
    pos = InStr(1, sec, rstc!seccio) - 1
    posicio = InStr(1, sec, rstc!seccio)
    If posicio = 0 Then posicio = 1
    If rstc!producte = "PC" Then posicio = 9
    If pos < 1 Then pos = 1
    sec = Mid(sec, pos, 1)
    If sec = "L" And rstc!producte <> "PC" And rstc!producte <> "PC2" Then
      sec = Mid(rstc!ruta, pos - 1, 1)
    End If
    If rstc!seccio = "S" And sec = "R" Then
       sec = Mid(rstc!ruta, pos - 1, 1)
    End If
    
    rstc.Edit
    rstc!seccioanterior = sec
    rstc!posicioruta = posicio
    rstc.Update
    rstc.MoveNext
  Wend
  Set rstc = Nothing
End Sub
Sub borrartaulaassignades()
 On Error Resume Next
 dbtmp.Execute "drop table assignadesafabrica"
 On Error GoTo 0
End Sub
Sub llistat_inventari_impanonim()
  Dim consulta As String
  Dim rstinv As Recordset
  Dim rstmat As Recordset
  Dim rstres As Recordset
  Dim rstfam As Recordset
  Dim ultimpalet As Double
  Dim nomfamilia As String
  Dim nommaterial As String
   ' si tipus es  "sensereserva"  o "ambreserva"
   On Error Resume Next
   dbllistat.Execute "drop table llistatinventari "
   On Error GoTo 0
   'faig aquesta consulta per crear la mateixa taula que a l'inventari disponible i poder fer servir el mateix llistat
   consulta = "SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, Bobines.kilos AS kilos ,bobines.disponible as metres,palets.micres as espesor"
   consulta = consulta + " into llistatinventari in '" + nomfitxertemporal + "'"
   consulta = consulta + " FROM Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet where palets.disponible and bobines.disponible>0;"
   dbtmp.Execute consulta
   
   
End Sub

Private Sub invlamanonim_Click()
Dim consulta As String
  Dim Grups As String
  Dim grups2 As String
  Dim sqlgrups2
   borrartaulallistatinventari
   buscarcomandesafabricaasignades
   creartaulatempllistat
   Grups = " select comanda from assignadesafabrica where seccio='L' and seccioanterior='E' and producte<>'PC'"
   'ELS GRUPS2 NOMES QUAN ES LAMINADORA.. o sigui a rebobinadora i imprès
   grups2 = "select trim(linkcomanda1) from assignadesafabrica where linkcomanda1>0 and seccio='L' and secciolink1='V'"
   'posso els palets i bobines al llistat
   'posso primer tot menys els laminats
   consulta = "Insert into llistatinventari in '" + nomfitxertemporal + "'"
   consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,parcials.metres as metres,palets.micres as espesor"
   consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (parcials.comanda in (" + Grups + "));" ' + sqlgrup2 + ");"
   dbtmp.Execute consulta
   'afegeixo els PC perque la proximaseccio es posa amb V i queda fora del select
   If grups2 <> "" Then
    consulta = "Insert into llistatinventari in '" + nomfitxertemporal + "'"
    consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,parcials.metres as metres,palets.micres as espesor"
    consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (parcials.comanda in (" + grups2 + "));"
    dbtmp.Execute consulta
   End If
   
   passoelscodimatlaminatsamatfulla1
   passodemetresakilos
   possoelsnomsalsmaterials
   
   
   Set rstcom = Nothing
   
  'faig el llistat
  llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistat inventari.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 llistat.Formulas(0) = "nomllistat='Llistat inventari a Laminadora (Anònim)'"
 llistat.Formulas(1) = "hora='" + format(Now, "dd/mm/yy  hh:nn") + "'"
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 dbllistat.Close
 obrir_dbllistats
End Sub

Private Sub invlamimpres_Click()
Dim consulta As String
  Dim Grups As String
  Dim grups2 As String
  Dim sqlgrups2
   borrartaulallistatinventari
   buscarcomandesafabricaasignades
   creartaulatempllistat
   Grups = " select comanda from assignadesafabrica where seccio='L' and seccioanterior='I' and producte<>'PC'"
   'ELS GRUPS2 NOMES QUAN ES LAMINADORA.. o sigui a rebobinadora i imprès
   'grups2 = "select trim(linkcomanda1) from assignadesafabrica where linkcomanda1>0 and seccio='R' and secciolink1='V'"
   'posso els palets i bobines al llistat
   'posso primer tot menys els laminats
   consulta = "Insert into llistatinventari in '" + nomfitxertemporal + "'"
   consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,parcials.metres as metres,palets.micres as espesor"
   consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (parcials.comanda in (" + Grups + "));" ' + sqlgrup2 + ");"
   dbtmp.Execute consulta
   'posso ara els laminats PER SUPUSAT QUAN FAIG REBOBINADORA,QUE ES QUAN CAL, SI NO NO CAL
   If grups2 <> "" Then
    consulta = "Insert into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
    consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,parcials.comanda as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,parcials.metres as metres,palets.micres as espesor"
    consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (parcials.comanda in (" + grups2 + "));"
    dbtmp.Execute consulta
   End If
   
   passoelscodimatlaminatsamatfulla1
   passodemetresakilos
   possoelsnomsalsmaterials
   
   
   Set rstcom = Nothing
   
  'faig el llistat
  llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistat inventari.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 llistat.Formulas(0) = "nomllistat='Llistat inventari a Laminadora (Imprès)'"
 llistat.Formulas(1) = "hora='" + format(Now, "dd/mm/yy  hh:nn") + "'"
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 dbllistat.Close
 obrir_dbllistats
End Sub

Private Sub invrebanonim_Click()
Dim rstcom As Recordset
  Dim consulta As String
  Dim Grups As String
  Dim grups2 As String
  Dim sqlgrups2
   borrartaulallistatinventari
   buscarcomandesafabricaasignades
   creartaulatempllistat
   Grups = " select comanda from assignadesafabrica where seccio='R' and seccioanterior='E' and producte<>'PC'"
   'ELS GRUPS2 NOMES QUAN ES LAMINADORA.. o sigui a rebobinadora i imprès
   'grups2 = "select trim(linkcomanda1) from assignadesafabrica where linkcomanda1>0 and seccio='R' and secciolink1='V'"
   'posso els palets i bobines al llistat
   'posso primer tot menys els laminats
   consulta = "Insert into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
   consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,parcials.metres as metres,palets.micres as espesor"
   consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (parcials.comanda in (" + Grups + "));" ' + sqlgrup2 + ");"
   dbtmp.Execute consulta
   'posso ara els laminats PER SUPUSAT QUAN FAIG REBOBINADORA,QUE ES QUAN CAL, SI NO NO CAL
   If grups2 <> "" Then
    consulta = "Insert into llistatinventari in '" + nomfitxertemporal + "'"
    consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,parcials.comanda as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,parcials.metres as metres,palets.micres as espesor"
    consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (parcials.comanda in (" + grups2 + "));"
    dbtmp.Execute consulta
   End If
   
   passoelscodimatlaminatsamatfulla1
   passodemetresakilos
   possoelsnomsalsmaterials
   
   
   Set rstcom = Nothing
   
  'faig el llistat
  llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistat inventari.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 llistat.Formulas(0) = "nomllistat='Llistat inventari a Rebobinadores (Anònim)'"
 llistat.Formulas(1) = "hora='" + format(Now, "dd/mm/yy  hh:nn") + "'"
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 dbllistat.Close
 obrir_dbllistats
End Sub

Private Sub invrebimpres_Click()
  Dim rstcom As Recordset
  Dim consulta As String
  Dim Grups As String
  Dim grups2 As String
  Dim sqlgrups2
   borrartaulallistatinventari
   buscarcomandesafabricaasignades
   creartaulatempllistat
   Grups = " select comanda from assignadesafabrica where seccio='R' and producte<>'PC'"
   'ELS GRUPS2 NOMES QUAN ES LAMINADORA.. o sigui a rebobinadora
   grups2 = "select trim(linkcomanda1) from assignadesafabrica where linkcomanda1>0 and seccio='R' and secciolink1='V'"
   'posso els palets i bobines al llistat
   'posso primer tot menys els laminats
   consulta = "Insert into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
   consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,parcials.metres as metres,palets.micres as espesor"
   consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (parcials.comanda in (" + Grups + "));" ' + sqlgrup2 + ");"
   dbtmp.Execute consulta
   'posso ara els laminats PER SUPUSAT QUAN FAIG REBOBINADORA,QUE ES QUAN CAL, SI NO NO CAL
   If grups2 <> "" Then
    consulta = "Insert into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
    consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,parcials.comanda as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,parcials.metres as metres,palets.micres as espesor"
    consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (parcials.comanda in (" + grups2 + "));"
    dbtmp.Execute consulta
   End If
   
   passoelscodimatlaminatsamatfulla1
   passodemetresakilos
   possoelsnomsalsmaterials
   
   
   Set rstcom = Nothing
   
  'faig el llistat
  llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistat inventari.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 llistat.Formulas(0) = "nomllistat='Llistat inventari a Rebobinadores (Imprès i/o Laminat)'"
 llistat.Formulas(1) = "hora='" + format(Now, "dd/mm/yy  hh:nn") + "'"
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 dbllistat.Close
 obrir_dbllistats
  
End Sub
Sub afegircomandescomplexesalesfabricant()
 Dim csql As String * 1000
 Dim rstp As Recordset
 Dim rstp2 As Recordset
 Dim csql2 As String
 Dim csql3 As String
  csql = "SELECT  comanda, producte, productes.ruta, proximaseccio AS Seccio, '' AS SeccioLink1, '' AS Secciolink2 ,'' as seccioanterior,linkcomanda1,linkcomanda2,9 as posicioruta,materialex,datacomanda "
  csql = Trim(csql) + " FROM  comandes INNER JOIN productes ON comandes.producte = productes.codi "
  csql2 = " Where comanda in (select linkcomanda1 from assignadesafabrica where linkcomanda1>0 and producte<>'PC' and producte<>'PC2')"
  csql3 = " Where comanda in (select linkcomanda2 from assignadesafabrica where linkcomanda2>0 and producte<>'PC' and producte<>'PC2')"
 ' MsgBox Trim(Trim(csql) + csql2)
  dbtmp.Execute "insert into assignadesafabrica " + Trim(Trim(csql) + csql2)
  dbtmp.Execute "insert into assignadesafabrica " + Trim(Trim(csql) + csql3)
 ' Set rstp = dbtmp.OpenRecordset("select * from assignadesafabrica where producte='PC' or producte='PC2'")
 ' While Not rstp.EOF
 '   Set rstp2 = dbtmp.OpenRecordset("select seccioanterior from assignadesafabrica where comanda='" + atrim(rstp!linkcomanda1) + "'")
 '   If Not rstp2.EOF Then rstp.Edit: rstp!seccioanterior = rstp2!seccioanterior: rstp.Update
 '   rstp.MoveNext
 ' Wend
  
  
End Sub
Sub possarsituacio(ll1 As String, num As Double, ll2 As String, sit As String)
   Dim n As String
   Dim c As Byte
   ll1 = Mid(sit, 1, 1)
   c = 1
   While IsNumeric(Mid(sit, 2, c + 1)) And c < Len(sit)
     c = c + 1
   Wend
   num = cadbl(Mid(sit, 2, c))
   If num > 0 Then
      ll2 = Mid(sit, c + 2, 1)
      If ll2 = "" Then ll2 = "A"
       Else: num = 0: ll2 = ""
   End If
   
End Sub
Sub possardadesrestants(rstc As Recordset)
  Dim rstm As Recordset
  Dim rstp As Recordset
  Dim proveidor As String
  Dim tipusbobina As String
  Dim nomproducte As String
  Dim ll1sit  As String
  Dim numerosit As Double
  Dim ll2sit As String
  
  If selecciomicres > 0 Then If cadbl(rstc!micres) <> selecciomicres Then rstc.Delete: Exit Sub
  
  If atrim(selecciofam) <> "" Then
    Set rstm = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(rstc!codimatprognou) + " and " + selecciofam)
    If rstm.EOF Then rstc.Delete: Exit Sub
  End If
  
  possarsituacio ll1sit, numerosit, ll2sit, atrim(rstc!sit)
  tipusbobina = IIf(assignarmat.esparcial(rstc!idpalet, rstc!idbobina), "P", "")
  tipusbobina = IIf(assignarmat.esrestu(rstc!idpalet, rstc!idbobina), "R", tipusbobina)
  If tipusbobina = "" Then tipusbobina = "E"
  
  Set rstm = dbtmpb.OpenRecordset("SELECT materials.codi, materials.descripcio as nom,proveidors.nom as proveidor FROM materials INNER JOIN proveidors ON materials.proveidor = proveidors.codi WHERE (((materials.codi)=" + atrim(rstc!codimatprognou) + "));")
  If Not rstm.EOF Then
    nomproducte = rstm!nom
    proveidor = rstm!proveidor
    
  End If
  rstc.Edit
  rstc!diametre = bobinesdentrada.calcular_diametre(rstc!idpalet, rstc!idbobina)
  rstc!proveidor = proveidor
  rstc!descproducte = nomproducte
  rstc!enterapracialresto = tipusbobina
  rstc!disponible = bobinesdentrada.calcular_mtrsdispreals(rstc!idpalet, rstc!idbobina)
  If rstc!mts > 0 Then
     rstc!kilos = format((rstc!kilos / rstc!mts) * rstc!disponible, "#,##0")
    Else: rstc!kilos = 0
  End If
  rstc!sit = convertirsituacio(UCase(atrim(rstc!sit)))
  If numerosit = 0 Then ll2sit = ll1sit: ll1sit = ""
  rstc!ll1sit = UCase(ll1sit)
  rstc!numerosit = numerosit
  rstc!ll2sit = UCase(ll2sit)
  rstc.Update
End Sub
Function convertirsituacio(sit As String) As String
  sit = UCase(sit)
  If Not IsNumeric(Mid(sit, 2, 1)) Then
     sit = Mid(sit, 1, 1)
  End If
  convertirsituacio = sit
End Function

Private Sub invstoctteoric_Click()
   inventariteoric
End Sub

 Sub inventariteoric()
  Dim consulta As String
  Dim rstc As Recordset
  Dim rutastocks As String
  selecciofam = ""
  nomfiltrefam = ""
  borrartaulallistatinventari
  sel_families.Show 1
  If selecciofam = "NO" Then Exit Sub
  ratoli "espera"
'  llistatestocproduccio "nomesconsulta", "<3"
  creartaulatempllistat
  
  'subconsulta = " SELECT Parcials.id "
  'subconsulta = subconsulta + " FROM Parcials INNER JOIN Palets ON Parcials.idpalet = Palets.Idpalet "
  'subconsulta = subconsulta + " WHERE (Parcials.utilitzada=False AND Len([comanda])>3 AND (cdbl([comanda])>2000 and cdbl([comanda])<3000 or cdbl([comanda])>140000) and (parcials.idpalet=bobines.idpalet and parcials.idbobina=bobines.idbobina))"
  'consulta = " insert into llistatinventari in 'c:\temporal.mdb' "
  'consulta = consulta + "SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,bobines.disponible as metres,palets.micres as espesor,palets.preucompra "
  'consulta = consulta + " FROM Palets INNER JOIN Bobines  ON Palets.Idpalet = Bobines.Idpalet where exists (" + subconsulta + ")"
  'dbtmp.Execute consulta
  'MsgBox consulta
  consulta = " insert into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
  consulta = consulta + "SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,bobines.disponible as metres,palets.micres as espesor,palets.preucompra "
  consulta = consulta + " FROM Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet where palets.disponible and bobines.disponible>0;"
  dbtmp.Execute consulta
  
  On Error Resume Next
  dbllistat.Execute "drop table llistatinventariprestatgeries"
  On Error GoTo 0
'  consulta = "SELECT  Bobines.Sit, Bobines.Idpalet, Bobines.Idbobina, Bobines.Mts, Bobines.disponible, palets.codimatprognou,Palets.semielaborat, Palets.Ample, Palets.micres, Palets.grmsm2  "
'  consulta = consulta + "  into llistatinventariprestatgeries in 'c:\temporal.mdb' "
'  consulta = consulta + " FROM Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet "
'  consulta = consulta + " where (Trim(bobines.idpalet)+Trim(bobines.idbobina)) In (select  (Trim(llistatinventari.palet)+Trim(llistatinventari.bobina)) as idp from llistatinventari in 'c:\temporal.mdb');"
  
  consulta = "SELECT  Bobines.Sit, Bobines.Idpalet, Bobines.Idbobina, Bobines.Mts, Bobines.disponible, bobines.kilos,palets.codimatprognou,Palets.semielaborat, Palets.Ample, Palets.micres, Palets.grmsm2  "
  consulta = consulta + "  into llistatinventariprestatgeries in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
  consulta = consulta + " FROM Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet "
  consulta = consulta + " where exists (SELECT * FROM llistatinventari in '" + nomfitxertemporal + "' where palets.idpalet=llistatinventari.palet and bobines.idbobina=llistatinventari.bobina);"
  
  
  dbtmp.Execute consulta
  
  dbllistat.Execute "Alter table llistatinventariprestatgeries add column proveidor string"
  dbllistat.Execute "Alter table llistatinventariprestatgeries add column descproducte string"
  dbllistat.Execute "Alter table llistatinventariprestatgeries add column enterapracialresto string"
  dbllistat.Execute "Alter table llistatinventariprestatgeries add column diametre double"
  dbllistat.Execute "Alter table llistatinventariprestatgeries add column ll1sit string"
  dbllistat.Execute "Alter table llistatinventariprestatgeries add column numerosit double"
  dbllistat.Execute "Alter table llistatinventariprestatgeries add column ll2sit string "
  
  
  
  Set rstc = dbllistat.OpenRecordset("select * from llistatinventariprestatgeries")
  Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "Palets.mdb")
  While Not rstc.EOF
     possardadesrestants rstc
    rstc.MoveNext
  Wend
  wait 3
  dbllistat.Execute "delete * from llistatinventariprestatgeries where kilos<1"
  rutastocks = llegir_ini("General", "ruta_stocks", "comandes.ini")
  wait 2
  If filtrarprestatge = "Altes" Then dbllistat.Execute "delete * from llistatinventariprestatgeries where llistatinventariprestatgeries.ll1sit+trim(llistatinventariprestatgeries.numerosit) not in (select numlleixa from prestatges in '" + rutastocks + "' where prestatgealt)"
  If filtrarprestatge = "Baixes" Then dbllistat.Execute "delete * from llistatinventariprestatgeries where llistatinventariprestatgeries.ll1sit+trim(llistatinventariprestatgeries.numerosit) in (select numlleixa from prestatges in '" + rutastocks + "' where prestatgealt)"
'  If filtrarprestatge = "Baixes" Then dbllistat.Execute "delete * from llistatinventariprestatgeries where exists (select * from prestatgesalts in '" + rutastocks + "' where llistatinventariprestatgeries.ll1sit+trim(llistatinventariprestatgeries.numerosit)=prestatgesalts.numlleixa)"
  
  Set rstc = Nothing
   For i = 1 To 50
    llistat.Formulas(i) = ""
   Next i
  'faig el llistat
  llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "inventariprestatges.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 llistat.Formulas(0) = "nomllistat='Inventari de prestatgeries: Filtrant prestatges:" + filtrarprestatge + "'"
 llistat.Formulas(1) = "hora='" + format(Now, "dd/mm/yy  hh:nn") + "'"
 llistat.Formulas(2) = "descripciofiltre='" + nomfiltrefam + "'"
 llistat.Formulas(3) = ""
 llistat.Formulas(4) = ""
 
 
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 dbstocks.Close
 dbllistat.Close
 obrir_dbllistats
 ratoli "normal"
  
End Sub

Private Sub invtotselsforats_Click()
   Dim rstforats As Recordset
   Dim rstbob As Recordset
   Dim bob As String
   Dim pres As String
   Dim metresdis As Double
   Dim contador As Double
   Dim vnomfitxerxls As String
   Dim vforatsbuits As Double
   
   ratoli "espera"
   'obrestocks
   Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "Palets.mdb")
   obrir_dbllistats
   borrartaulallistatinventari
   dbtmp.Execute "select trim(columna) as nlleixa,estanteria as lletralleixa,fila as lloclleixa, 'A' as foratlleixa into llistatinventari in '" + nomfitxertemporal + "' from prestatgesnous" ' where foratsperlleixa>0"
   dbllistat.Execute "alter table llistatinventari add column bobines text"
   'dbtmp.Execute "insert into llistatinventari in 'c:\temporal.mdb'select numlleixa+'B' as nlleixa,mid(numlleixa,1,1) as lletralleixa,cdbl(mid(numlleixa,2)) as lloclleixa, 'B' as foratlleixa from prestatges where foratsperlleixa>1"
   'dbtmp.Execute "insert into llistatinventari in 'c:\temporal.mdb'select numlleixa+'C' as nlleixa,mid(numlleixa,1,1) as lletralleixa,cdbl(mid(numlleixa,2)) as lloclleixa, 'C' as foratlleixa  from prestatges where foratsperlleixa>2"
   'dbtmp.Execute "insert into llistatinventari in 'c:\temporal.mdb'select numlleixa+'D' as nlleixa,mid(numlleixa,1,1) as lletralleixa,cdbl(mid(numlleixa,2)) as lloclleixa, 'C' as foratlleixa  from prestatges where foratsperlleixa>3"
   'dbtmp.Execute "insert into llistatinventari in 'c:\temporal.mdb'select numlleixa+'E' as nlleixa,mid(numlleixa,1,1) as lletralleixa,cdbl(mid(numlleixa,2)) as lloclleixa, 'D' as foratlleixa  from prestatges where foratsperlleixa>4"
   'dbtmp.Execute "insert into llistatinventari in 'c:\temporal.mdb'select numlleixa+'F' as nlleixa,mid(numlleixa,1,1) as lletralleixa,cdbl(mid(numlleixa,2)) as lloclleixa, 'E' as foratlleixa  from prestatges where foratsperlleixa>5"
   Set rstforats = dbllistat.OpenRecordset("SELECT * from llistatinventari order by lletralleixa,nlleixa,lloclleixa")
   contador = 0
   vforatsbuits = 0
   While Not rstforats.EOF
     Set rstbob = dbtmp.OpenRecordset("select * from bobines where sit='" + atrim(rstforats!lletralleixa) + format(rstforats!nlleixa, "00") + atrim(rstforats!lloclleixa) + "'")
     bob = ""
     If rstbob.EOF Then vforatsbuits = vforatsbuits + 1
     While Not rstbob.EOF
       metresdis = bobinesdentrada.calcular_mtrsdispreals(rstbob!idpalet, rstbob!idbobina)
       If metresdis < 1 Then
          bob = bob + "##[" + atrim(rstbob!idpalet) + "/" + atrim(rstbob!idbobina) + "]## "
          contador = contador + 1
            Else: bob = bob + "[" + atrim(rstbob!idpalet) + "/" + atrim(rstbob!idbobina) + "] "
       End If
       rstbob.MoveNext
     Wend
     rstforats.Edit
     rstforats!bobines = Mid(bob, 1, 250)
     rstforats!nlleixa = atrim(rstforats!lletralleixa) + format(rstforats!nlleixa, "00") + atrim(rstforats!lloclleixa)
     rstforats.Update
     rstforats.MoveNext
     DoEvents
   Wend
     'faig el llistat
  llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "inventariperforat.rpt"
  llistat.Destination = crptToPrinter
  llistat.CopiesToPrinter = 1
  llistat.DataFiles(0) = nomfitxertemporal
  llistat.DiscardSavedData = True
  llistat.Formulas(0) = "hora='" + format(Now, "dd/mm/yy  hh:nn") + "'"
  llistat.Formulas(1) = "titol='Llistat de bobines parcials.'"
  llistat.Formulas(2) = "foratsbuits=" + atrim(vforatsbuits)
  llistat.Formulas(3) = ""
  DoEvents
  If vexportantelllistat Then
    llistat.Destination = crptToFile
    llistat.PrintFileType = crptExcel50
    vnomfitxerxls = "\\ser2\DOCUMENTOS\Agusti Feliu\PALETS situacio\"
    If Not existeix(vnomfitxerxls + Trim(Year(Now))) Then MkDir vnomfitxerxls + Trim(Year(Now))
    If Not existeix(vnomfitxerxls + Trim(Year(Now)) + "\" + Trim(format(Now, "m")) + "-" + UCase(atrim(format(Now, "mmmm")))) Then MkDir vnomfitxerxls + Trim(Year(Now)) + "\" + Trim(format(Now, "m")) + "-" + UCase(atrim(format(Now, "mmmm")))
    vnomfitxerxls = UCase(vnomfitxerxls + atrim(Year(Now)) + "\" + Trim(format(Now, "m")) + "-" + UCase(atrim(format(Now, "mmmm"))) + "\" + format(Now, "yymmdd") + ".xls")
    
    If existeix(vnomfitxerxls) Then Kill vnomfitxerxls
    llistat.PrintFileName = vnomfitxerxls
     Else
      If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
      If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
      If contador > 0 Then MsgBox "He trobat incidencies en els forats hi han " + atrim(contador) + " bobines marcades per revisar-les", vbCritical + vbOKOnly, "Atenció"
  End If
  llistat.Action = 1
  dbllistat.Close
  obrir_dbllistats
  Set rstbob = Nothing
  Set rstforats = Nothing
  ratoli "normal"
End Sub

Sub llistat_inventari_disponible_PALETSNODISPONIBLES()
  Dim i As Byte
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistatinventaripaletsnodisponibles.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = palets.DatabaseName
 For i = 1 To 8
   llistat.DataFiles(i) = cami
 Next i
 llistat.DiscardSavedData = True
 llistat.Formulas(0) = ""
 llistat.Formulas(1) = ""
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 For i = 0 To 8
   llistat.DataFiles(i) = ""
 Next i
End Sub

Private Sub llestrassig_Click()
   llistatestocproduccio "Llistat d'assignat que no s'està produint.", ">2"
End Sub

Private Sub llisestocproduccio_Click()
   llistatestocproduccio "Llistat d´estoc en produccio.", "<3"
End Sub
Sub llistatestocproduccio(nomllistat As String, criteri As String)
  Dim bobspreu0 As String
  Dim consulta As String
  Dim Grups As String
  Dim grups2 As String
  Dim rstc As Recordset
  Dim rstb As Recordset
  Dim numc As Double
  Dim numc2 As Double
  Dim codimat As String
  Dim sqlgrups2
  Dim rstpc As Recordset
  Dim comandesnoassignades As String
  ratoli "espera"
   borrartaulallistatinventari
   buscarcomandesafabricaasignades
   Form1.Caption = "Manteniment de Palets (Borrant bobines innecessaries)"
   dbtmp.Execute "delete * from assignadesafabrica where posicioruta" + criteri
   dbtmp.Execute "delete * from assignadesafabrica where producte='PC' or producte='PC2'"
   Form1.Caption = "Manteniment de Palets (Afegint comandes complexes)"
   afegircomandescomplexesalesfabricant
   Form1.Caption = "Manteniment de Palets (Creant el llistat)"
   creartaulatempllistat
   Grups = " select comanda from assignadesafabrica where posicioruta>2"
   'posso els palets i bobines al llistat
   
   'consulta = "Insert into llistatinventari in 'c:\temporal.mdb' "
   'consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,parcials.comanda as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,parcials.metres as metres,palets.micres as espesor,palets.preucompra as preucompra ,'' as producte, '' as descripcio "
   'consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (parcials.comanda in (" + Grups + "));" ' + sqlgrup2 + ");"
   'dbtmp.Execute consulta
   Set rstc = dbtmp.OpenRecordset("select distinct comanda,* from assignadesafabrica ")
   Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", fitxerini))
   Set rstb = dbtmp.OpenRecordset("select * from assignadesafabrica where linkcomanda1=-9999")
   While Not rstc.EOF
     numc = cadbl(rstc!comanda)
     
     'codimat = ""
     'Set rstpc = Nothing
     'If InStr(1, rstc!producte, "PC") > 0 Then
     '   numc = cadbl(rstc!linkcomanda1): numc2 = cadbl(rstc!comanda): codimat = atrim(rstc!materialex)
     '   Set rstapc = dbtmp.OpenRecordset("select * from parcials where comanda='" + atrim(numc2) + "'")
     'End If
     If atrim(rstc!seccioanterior) = "I" Then Set rstb = dbbaixes.OpenRecordset("SELECT First(bobinesentimp.palet) AS palet, impressores.comanda, First(bobinesentimp.bobina) AS bobina, bobinesimp.metres, bobinesimp.numerodebobina FROM bobinesentimp LEFT JOIN (bobinesimp LEFT JOIN impressores ON bobinesimp.controlid = impressores.Id) ON bobinesentimp.id = bobinesimp.Id GROUP BY impressores.comanda, bobinesimp.metres, bobinesimp.numerodebobina HAVING (((impressores.comanda)=" + atrim(numc) + "));")
     If atrim(rstc!seccioanterior) = "L" Then
        secant = ""
        secant = Mid(rstc!ruta, InStr(rstc!ruta, rstc!seccioanterior) - 1, 1)
        If secant = "I" Then
         Set rstb = dbbaixes.OpenRecordset("SELECT First(bobinesentimp.palet) AS palet, impressores.comanda, First(bobinesentimp.bobina) AS bobina, bobinesimp.metres, bobinesimp.numerodebobina FROM bobinesentimp LEFT JOIN (bobinesimp LEFT JOIN impressores ON bobinesimp.controlid = impressores.Id) ON bobinesentimp.id = bobinesimp.Id GROUP BY impressores.comanda, bobinesimp.metres, bobinesimp.numerodebobina HAVING (((impressores.comanda)=" + atrim(numc) + "));")
           Else
               Set rstb = dbtmp.OpenRecordset("select idpalet as palet,idbobina as bobina,metres from parcials where metres>0 and comanda='" + atrim(numc) + "'")
        End If
        'Set rstb = dbbaixes.OpenRecordset("SELECT laminadores.comanda, First(bobinesentlam.palet) AS palet, First(bobinesentlam.bobina) AS bobina, First(bobineslam.metres) AS metres, bobinesentlam.paletobobina, bobineslam.numerodebobina FROM (bobinesentlam INNER JOIN bobineslam ON bobinesentlam.id = bobineslam.Id) INNER JOIN laminadores ON bobineslam.controlid = laminadores.Id GROUP BY laminadores.comanda, bobinesentlam.paletobobina, bobineslam.numerodebobina HAVING (((laminadores.comanda)=" + atrim(numc) + ") AND ((bobinesentlam.paletobobina)='p'));")
     End If
     If atrim(rstc!seccioanterior) = "R" And (atrim(rstc!ruta) = "ER" Or atrim(rstc!ruta) = "ERS") Then Set rstb = dbbaixes.OpenRecordset("SELECT First(bobinesentreb.palet) AS palet, rebobinadores.comanda, First(bobinesentreb.bobina) AS bobina, bobinesreb.metres, bobinesreb.numerodebobina FROM bobinesentreb LEFT JOIN (bobinesreb LEFT JOIN rebobinadores ON bobinesreb.controlid = rebobinadores.Id) ON bobinesentreb.id = bobinesreb.Id GROUP BY rebobinadores.comanda, bobinesreb.metres, bobinesreb.numerodebobina HAVING (((rebobinadores.comanda)=" + atrim(numc) + "));")
     If InStr(1, rstc!producte, "PC") > 0 Or atrim(rstc!seccioanterior) = "E" Then Set rstb = dbtmp.OpenRecordset("select idpalet as palet,idbobina as bobina,metres from parcials where metres>0 and comanda='" + atrim(numc) + "'")
     If rstb.EOF Then comandesnoassignades = comandesnoassignades + ", " + atrim(numc)
     While Not rstb.EOF
        possarregistreallistat numc, rstb
        rstb.MoveNext
     Wend
     Set rstb = dbtmp.OpenRecordset("select * from assignadesafabrica where linkcomanda1=-9999")
     rstc.MoveNext
   Wend
   
   If nomllistat = "nomesconsulta" Then Exit Sub
   'passoelscodimatlaminatsamatfulla1
   passodemetresakilos
   possoelsnomsalsmaterials
   possocampsextresalllistat
   bobspreu0 = bobinesambpreu0
   
   Set rstcom = Nothing
   
  'faig el llistat
  llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistat estoc en produccio.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 llistat.Formulas(0) = "nomllistat='" + treure_apostruf(nomllistat) + "'"
 llistat.Formulas(1) = "hora='" + format(Now, "dd/mm/yy  hh:nn") + "'"
 llistat.Formulas(2) = "comandesnoassignades='" + Mid(atrim(comandesnoassignades), 1, 210) + "'"
 llistat.Formulas(3) = "bobinesambpreu0='" + atrim(bobspreu0) + "'"
 If criteri = "<3" Then
    llistat.Formulas(4) = "imp30%+='S'"
   Else: llistat.Formulas(4) = "imp30%+='N'"
 End If
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 dbllistat.Close
 obrir_dbllistats
 ratoli "normal"
 Form1.Caption = "Manteniment de Palets"
End Sub
Function bobinesambpreu0() As String
  Dim rstl As Recordset
  Dim c As String
  Set rstl = dbllistat.OpenRecordset("select palet,bobina from llistatinventari where preucompra=0")
  While Not rstl.EOF
     c = c + ", " + atrim(rstl!palet) + "/" + atrim(rstl!bobina)
     rstl.MoveNext
  Wend
  bobinesambpreu0 = c
End Function
Sub possarregistreallistat(numc As Double, rstb As Recordset)
  Dim consulta As String
  Dim strp As Recordset
  If cadbl(rstb!metres) < 0 Then Exit Sub
  Set strp = dbtmp.OpenRecordset("select comanda from parcials where comanda='" + atrim(numc) + "' and (idpalet = " + atrim(cadbl(rstb!palet)) + " and idbobina=" + atrim(cadbl(rstb!bobina)) + ")")
  If strp.EOF Then Set strp = Nothing: Exit Sub
'  If atrim(codimat) = "" Then
    codimat = "palets.codimatprognou"
    consulta = "Insert into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
    consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,'" + atrim(numc) + "' as comanda, ' ' as familia, " + codimat + " AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ," + atrim(rstb!metres) + " as metres,palets.micres as espesor,palets.preucompra as preucompra ,'' as producte, '' as descripcio "
    consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) WHERE palets.disponible and (bobines.idpalet = " + atrim(cadbl(rstb!palet)) + " and bobines.idbobina=" + atrim(cadbl(rstb!bobina)) + ")"
    'consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (parcials.idpalet = " + atrim(cadbl(rstb!palet)) + " and parcials.idbobina=" + atrim(cadbl(rstb!bobina)) + ")"
 
 '  Else
 '    If rstpc.EOF Then Exit Sub
 '    consulta = "Insert into llistatinventari in 'c:\temporal.mdb' "
 '    consulta = consulta + " SELECT 9999 AS palet,' ' as nommaterial,' ' as seccio,'" + atrim(numc) + "' as comanda, ' ' as familia, " + codimat + " AS codimat, 999 AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ," + atrim(rstb!metres) + " as metres,palets.micres as espesor,palets.preucompra as preucompra ,'' as producte, '' as descripcio "
 '    consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) WHERE (bobines.idpalet = " + atrim(cadbl(rstpc!idpalet)) + " and bobines.idbobina=" + atrim(cadbl(rstpc!idbobina)) + ")"
 '    'consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (parcials.idpalet = " + atrim(cadbl(rstb!palet)) + " and parcials.idbobina=" + atrim(cadbl(rstb!bobina)) + ")"
 ' End If
  'MsgBox consulta
  dbtmp.Execute consulta

End Sub

Sub possocampsextresalllistat()
   Dim rsttmp As Recordset
   Dim rstp As Recordset
   Dim producte As String
   Dim rstc As Recordset
   Dim desc As String
   Dim seccio As String
   Dim datac As String
'passo els metres a kilos
   Set rsttmp = dbllistat.OpenRecordset("SELECT distinct comanda from llistatinventari ")
   While Not rsttmp.EOF
      Set rstp = dbtmpb.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(rsttmp!comanda)))
      If Not rstp.EOF Then
        If cadbl(rstp!client) > 999 Then
         Set rstc = dbtmpb.OpenRecordset("select nom from clients where codi=" + atrim(rstp!client))
         If Not rstc.EOF Then desc = Mid(atrim(rstc!nom), 1, 15)
         desc = desc + " - " + Mid(atrim(rstp!texteimpressio), 1, 20)
         producte = atrim(rstp!producte)
         seccio = atrim(rstp!proximaseccio)
         datac = format(rstp!datacomanda, "dd/mm/yy")
            Else: borrarcomandesdelllistatclientmespetitde1000 cadbl(rstp!comanda), cadbl(rstp!linkcomanda1), cadbl(rstp!linkcomanda2)
        End If
      End If
      
      dbllistat.Execute "update llistatinventari set descripcio ='" + treure_apostruf(desc) + "',producte='" + atrim(producte) + "', seccio='" + atrim(seccio) + "',datacomanda='" + datac + "' where comanda='" + atrim(rsttmp!comanda) + "'"
      rsttmp.MoveNext
      datac = ""
      desc = ""
      producte = ""
      seccio = ""
   Wend
   marcarcomandaamaquina
   Set rsttmp = Nothing
   Set rstp = Nothing
   Set rstc = Nothing
End Sub
Sub borrarcomandesdelllistatclientmespetitde1000(comanda As Double, comanda2 As Double, comanda3 As Double)
   dbllistat.Execute "delete * from llistatinventari where comanda='" + atrim(comanda) + "' or comanda='" + atrim(comanda2) + "' or comanda='" + atrim(comanda3) + "'"
End Sub
Sub marcarcomandaamaquina()
  Dim rstc As Recordset
  Set rstc = dbbaixes.OpenRecordset("SELECT impressores.comanda, impressores.datainici, IsDate([impressores].[datainici]), IsDate([impressores].[datafi]) AS Expr1, impressorestot.acavada FROM impressores INNER JOIN impressorestot ON impressores.comanda = impressorestot.comanda Where (((IsDate([impressores].[datainici])) <> False) And ((IsDate([impressores].[datafi])) = False) And ((impressorestot.acavada) = '0')) ORDER BY impressores.datainici DESC;")
  While Not rstc.EOF
    dbllistat.Execute "update llistatinventari set fabricant ='(*)' where comanda='" + atrim(cadbl(rstc!comanda)) + "'"
    rstc.MoveNext
  Wend
  
  Set rstc = dbbaixes.OpenRecordset("SELECT laminadores.comanda, laminadores.datainici, IsDate([laminadores].[datainici]), IsDate([laminadores].[datafi]) AS Expr1, laminadorestot.acavada FROM laminadores INNER JOIN laminadorestot ON laminadores.comanda = laminadorestot.comanda Where (((IsDate([laminadores].[datainici])) <> False) And ((IsDate([laminadores].[datafi])) = False) And ((laminadorestot.acavada) = '0')) ORDER BY laminadores.datainici DESC;")
  While Not rstc.EOF
    dbllistat.Execute "update llistatinventari set fabricant ='(*)' where comanda='" + atrim(cadbl(rstc!comanda)) + "'"
    rstc.MoveNext
  Wend
  
  Set rstc = dbbaixes.OpenRecordset("SELECT rebobinadores.comanda, rebobinadores.datainici, IsDate([rebobinadores].[datainici]), IsDate([rebobinadores].[datafi]) AS Expr1, rebobinadorestot.acavada FROM rebobinadores INNER JOIN rebobinadorestot ON rebobinadores.comanda = rebobinadorestot.comanda Where (((IsDate([rebobinadores].[datainici])) <> False) And ((IsDate([rebobinadores].[datafi])) = False) And ((rebobinadorestot.acavada) = '0')) ORDER BY rebobinadores.datainici DESC;")
  While Not rstc.EOF
    dbllistat.Execute "update llistatinventari set fabricant ='(*)' where comanda='" + atrim(cadbl(rstc!comanda)) + "'"
    rstc.MoveNext
  Wend
  
  
End Sub
Private Sub llistatinventariambreserva_Click()
 ratoli "espera"
 If Not comprovar_quereservanoassignat Then MsgBox "No es pot continuar sense reparar aquest problema": GoTo fi
 llistat_inventari_reserva
fi:
 ratoli "normal"
End Sub
Function comprovar_quereservanoassignat() As Boolean
   Dim rstres As Recordset
   Dim rstass As Recordset
   Dim assignades As String * 1000
'   Set rstres = dbtmp.OpenRecordset("select * from percomandaoclient where ")
'   While Not rstres.EOF
   Set rstass = dbtmp.OpenRecordset("select numcomanda,idcompra from percomandaoclient where trim(numcomanda)  in (select trim(comanda) from parcials)")
   assignades = ""
   While Not rstass.EOF
      If cadbl(rstass!idcompra) = 0 Then
       assignades = Trim(assignades) + " | " + atrim(rstass!numcomanda)
      End If
      rstass.MoveNext
   Wend
   If Trim(assignades) <> "" Then
      MsgBox Trim(assignades), vbCritical + vbOKOnly, "Comandes que estan assignades i reservades"
      comprovar_quereservanoassignat = False
        Else: comprovar_quereservanoassignat = True
   End If
   
'      rstres.MoveNext
 '  Wend
   
End Function
Sub llistat_inventari_reserva()
  Dim consulta As String
  Dim rstinv As Recordset
  Dim rstmat As Recordset
  Dim rstres As Recordset
  Dim rstfam As Recordset
  Dim ultimpalet As Double
  Dim nomfamilia As String
  Dim nommaterial As String
   ' si tipus es  "sensereserva"  o "ambreserva"
   borrartaulallistatinventari
   'faig aquesta consulta per crear la mateixa taula que a l'inventari disponible i poder fer servir el mateix llistat
   consulta = "SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, Bobines.kilos AS kilos ,bobines.disponible as metres,palets.micres as espesor"
   consulta = consulta + " into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
   consulta = consulta + " FROM Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet where palets.disponible and bobines.disponible>0;"
   dbtmp.Execute consulta
   dbllistat.Execute "delete * from llistatinventari"
   Set rstinv = dbllistat.OpenRecordset("select * from llistatinventari")
   Set rstres = dbtmp.OpenRecordset("select * from reserves")
   While Not rstres.EOF
   
     Set rstmat = dbtmp.OpenRecordset("select * from percomandaoclient where idreserva=" + atrim(rstres!idreserva))
     If Not rstmat.EOF Then
        Set rstmat = dbtmpb.OpenRecordset("select materialex from comandes where comanda=" + atrim(cadbl(rstmat!numcomanda)))
        If Not rstmat.EOF Then
          Set rstmat = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(rstmat!materialex))
        End If
     End If
     If rstmat.EOF Then GoTo notrobat
     nomfamilia = ""
     Set rstfam = dbtmpb.OpenRecordset("select * from familiesmaterials where codi=" + atrim(cadbl(rstres!familia)))
     If Not rstfam.EOF Then nomfamilia = rstfam!descripcio
     nommaterial = descripciomaterial(rstres)
     rstinv.AddNew
      rstinv!espesor = rstres!espesor
      rstinv!metres = rstres!metresreservats
      rstinv!kilos = demetresakilos(rstres!ample, rstmat!grmcm3, rstres!espesor, rstres!semielaborat, rstres!solapa)
      rstinv!kilos = format(rstinv!kilos * (rstres!ample / 100) * rstres!metresreservats, "#,##0")
      rstinv!nommaterial = Trim(nommaterial)
      rstinv!familia = Trim(nomfamilia)
      rstinv.Update
notrobat:
      rstres.MoveNext
   Wend
   
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistat inventari.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 llistat.Formulas(0) = "nomllistat='Llistat inventari RESERVATS'"
 llistat.Formulas(1) = "hora='" + format(Now, "dd/mm/yy  hh:nn") + "'"
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 ratoli "normal"
 llistat.Action = 1
 dbllistat.Close
 obrir_dbllistats
End Sub

Private Sub llistatinventarisensereserva_Click()
   llistat_inventari_disponible
End Sub
Sub borrartaulallistatinventari()
   On Error Resume Next
   dbllistat.Execute "drop table llistatinventari "
   On Error GoTo 0
End Sub
Sub llistat_inventari_disponible(Optional compra As String)
Dim Grups As String
Dim grups2 As String
  Dim espesor As Double
  Dim consulta As String
  Dim rstmat As Recordset
  Dim rstfam As Recordset
  Dim ultimpalet As Double
  Dim rstp As Recordset
  Dim nomfamilia As String
  Dim i As Byte
  
   borrartaulallistatinventari
   consulta = "SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,bobines.disponible as metres,palets.micres as espesor,palets.preucompra,palets.teimpost "
   consulta = consulta + " into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
   consulta = consulta + " FROM Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet where bobines.disponible>0 and (Palets.Datarec) Is Not Null;" 'palets.disponible and
   'MsgBox consulta
   'Clipboard.Clear
   'Clipboard.SetText consulta
   dbtmp.Execute consulta
   'afegeixo els 2015 2016 2017 i altres grups
   Set rstmat = dbtmp.OpenRecordset("select * from grupsdepalets")
   Grups = ""
   grups2 = ""
   While Not rstmat.EOF
     Grups = Grups + "'" + atrim(rstmat!numerogrup) + "'"
     rstmat.MoveNext
     
     If Not rstmat.EOF Then Grups = Grups + ","
     If Len(Grups) > 245 Then grups2 = Grups: Grups = ""
   Wend
   consulta = "Insert into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
   consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,bobines.disponible as metres,palets.micres as espesor, palets.preucompra,palets.teimpost "
   consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet) WHERE (comanda in (" + grups2 + Grups + ") );"
   'MsgBox consulta
   dbtmp.Execute consulta
   
   passodemetresakilos
   possoelsnomsalsmaterials
  
   wait 2
  
   dbllistat.Execute "delete * from llistatinventari where espesor=0 or kilos<=35 or espesor=null"
   
   Set rstfam = Nothing
   Set rstmat = Nothing
   If compra = "compres" Then
      llistatdecompra
     Else
        llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistat inventari.rpt"
        llistat.Destination = crptToPrinter
        llistat.CopiesToPrinter = 1
        llistat.DataFiles(0) = nomfitxertemporal
        llistat.DiscardSavedData = True
        llistat.Formulas(0) = "nomllistat='Llistat inventari DISPONIBLE    >35Kg (AMB ELS GRUPS INCLOSOS)'"
        llistat.Formulas(1) = "hora='" + format(Now, "dd/mm/yy  hh:nn") + "'"
        For i = 2 To 20
          llistat.Formulas(i) = ""
        Next i
        DoEvents
        If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
        If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
        llistat.Action = 1
        
  End If
  dbllistat.Close
  obrir_dbllistats
   
End Sub
Sub calcular_estocactual()
  Dim rst As Recordset
  Dim vestocatual As Double
  Set rst = dbllistat.OpenRecordset("select distinct palet from llistatinventari")
  While Not rst.EOF
     
     vestocactual = calcular_kilos_disponibles_palet(cadbl(rst!palet))
     dbllistat.Execute "update llistatinventari set estocactual=" + atrim(cadbl(vestocactual)) + " where palet=" + atrim(cadbl(rst!palet))
     rst.MoveNext
  Wend
 Set rst = Nothing
End Sub
Sub llistatdecompra()
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  
  Set oapp = New CRAXDDRT.Application
  dbllistat.Execute "alter table llistatinventari add column estocactual double"
  calcular_estocactual
  
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistatinventari_detallcompres.rpt", 1)
  oreport.Database.Tables.Item(1).Location = nomfitxertemporal '"c:\temporal.mdb"
  oreport.Database.Tables.Item(2).Location = rutadelfitxer(camistock) + "compres.mdb"
  oreport.FormulaFields.GetItemByName("nomllistat").Text = "'Llistat inventari DISPONIBLE    >35Kg (AMB ELS GRUPS INCLOSOS)'"
  oreport.FormulaFields.GetItemByName("hora").Text = "'" + format(Now, "dd/mm/yy  hh:nn") + "'"

  oreport.DiscardSavedData
  'If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = True
   veurereport.CRViewer.ViewReport
   veurereport.Show 1
End Sub
Sub llistat_inventari_disponibleNOMESDEGRUPS()
Dim Grups As String
Dim grups2 As String
  Dim espesor As Double
  Dim consulta As String
  Dim rstmat As Recordset
  Dim rstfam As Recordset
  Dim ultimpalet As Double
  Dim rstp As Recordset
  Dim nomfamilia As String
   ' si tipus es  "sensereserva"  o "ambreserva"
  
   borrartaulallistatinventari
   consulta = "SELECT top 1 Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,bobines.disponible as metres,0.0 as mtrsassignats,palets.micres as espesor,palets.preucompra "
   consulta = consulta + " into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
   consulta = consulta + " FROM Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet where bobines.disponible>0;" 'palets.disponible and
   'MsgBox consulta
   
   dbtmp.Execute consulta
   dbllistat.Execute "delete * from llistatinventari"
   'afegeixo els 2015 2016 2017 i altres grups
   Set rstmat = dbtmp.OpenRecordset("select * from grupsdepalets")
   Grups = ""
   grups2 = ""
   While Not rstmat.EOF
     Grups = Grups + "'" + atrim(rstmat!numerogrup) + "'"
     rstmat.MoveNext
     
     If Not rstmat.EOF Then Grups = Grups + ","
     If Len(Grups) > 245 Then
       grups2 = Grups: Grups = ""
     End If
   Wend
   consulta = "Insert into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
   consulta = consulta + " SELECT Palets.Idpalet AS palet,' ' as nommaterial,' ' as seccio,' ' as comanda, ' ' as familia, Palets.codimatprognou AS codimat, Bobines.Idbobina AS bobina, bobines.mts as metresbob,Bobines.kilos AS kilos ,parcials.metres as metres,palets.micres as espesor, cdbl(parcials.comanda) as preucompra "
   consulta = consulta + " FROM (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)WHERE (comanda in (" + grups2 + Grups + ") );"
   'MsgBox consulta
   dbtmp.Execute consulta
   'Clipboard.Clear
   'Clipboard.SetText consulta
   
   
   possoelsmetresassignats
   passodemetresakilos
  
  ' dbllistat.Execute "delete * from llistatinventari where espesor=0 or kilos<=35 or espesor=null"
   
   Set rstfam = Nothing
   Set rstmat = Nothing
   
   llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistat inventari GRUPS.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DataFiles(1) = ""
 llistat.DiscardSavedData = True
 llistat.Formulas(0) = "nomllistat='Llistat inventari de TOTS ELS GRUPS (Ja inclosos en el de DISPONIBLE)'"
 llistat.Formulas(1) = "hora='" + format(Now, "dd/mm/yy  hh:nn") + "'"
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 dbllistat.Close
 obrir_dbllistats
   
End Sub

Sub possoelsmetresassignats()
  Dim rstm As Recordset
  Dim rstmpc As Recordset
  Dim rstll As Recordset
  Dim nommaterial As String
  Dim nomfamilia As String
  Dim vComandes As String
  Dim rstmat As Recordset
  'Set rstm = dbtmp.OpenRecordset("SELECT grupsdepalets.numerogrup, First(grupsdepalets.nomdelgrup) AS pnomdelgrup, Sum(comandes.cantitatex) AS smetres FROM (opcionsdajust INNER JOIN comandes ON opcionsdajust.comanda = comandes.comanda) INNER JOIN grupsdepalets ON opcionsdajust.grupdestoc = grupsdepalets.numerogrup Where comandes.producte<>'PC' and comandes.producte<>'PC2' and (comandes.proximaseccio='I' or comandes.proximaseccio='L') GROUP BY grupsdepalets.numerogrup;")
  'Set rstmpc = dbtmp.OpenRecordset("SELECT grupsdepalets.numerogrup, First(grupsdepalets.nomdelgrup) AS pnomdelgrup, Sum(comandes.cantitatex) AS smetres FROM (opcionsdajust INNER JOIN comandes ON opcionsdajust.comanda = comandes.comanda) INNER JOIN grupsdepalets ON opcionsdajust.grupdestoc = grupsdepalets.numerogrup Where (comandes.producte='PC' or comandes.producte='PC2') and (comandes.proximaseccio='I' or comandes.proximaseccio='L') GROUP BY grupsdepalets.numerogrup;")
  Set rstm = dbtmp.OpenRecordset("select * from grupsdepalets")
  While Not rstm.EOF
    vsqlLAM = "SELECT Sum([COMANDES].[cantitatex]) AS smetres FROM ((opcionsdajust LEFT JOIN comandes ON opcionsdajust.comanda = comandes.comanda) LEFT JOIN productes ON comandes.producte = productes.codi) LEFT JOIN comandes AS comandes_1 ON comandes.linkcomanda1 = comandes_1.comanda WHERE (((comandes.proximaseccio)<>'T') AND ((opcionsdajust.grupdestoc)=" + atrim(cadbl(rstm!numerogrup)) + ") AND ((comandes.producte)='PC' Or (comandes.producte)='PCP' Or (comandes.producte)='PC2') AND ((comandes_1.proximaseccio)='E' Or (comandes_1.proximaseccio)='I' Or (comandes_1.proximaseccio)='L'));"
    vsqlIMP = "SELECT Sum([COMANDES].[cantitatex]) AS smetres FROM ((opcionsdajust LEFT JOIN comandes ON opcionsdajust.comanda = comandes.comanda) LEFT JOIN productes ON comandes.producte = productes.codi) LEFT JOIN comandes AS comandes_1 ON comandes.linkcomanda1 = comandes_1.comanda WHERE (comandes.proximaseccio='E' Or comandes.proximaseccio='I')  AND (opcionsdajust.grupdestoc=" + atrim(cadbl(rstm!numerogrup)) + ");"
   ' vsqlIMP = "SELECT  Sum(comandes.cantitatex) AS smetres FROM opcionsdajust INNER JOIN comandes ON opcionsdajust.comanda = comandes.comanda Where (comandes.proximaseccio='I') and opcionsdajust.grupdestoc=" + atrim(cadbl(rstm!numerogrup)) + ";"
    vImp_o_Lam = IIf(rstm!seccio = "I", vsqlIMP, IIf(rstm!seccio = "L", vsqlLAM, ""))
    'posso el valor de metres de comandes assignades a aquest grup en el camp preucompra que es el que faig servir en el llistat per sumar els metres
    Set rstmpc = dbtmp.OpenRecordset(vImp_o_Lam)
    If Not rstmpc.EOF Then
     Set rstll = dbllistat.OpenRecordset("Select * from llistatinventari where metres>0 and preucompra=" + atrim(cadbl(rstm!numerogrup)))
     If Not rstll.EOF Then
         'If cadbl(rstm!numerogrup) = 2887 Then Stop
         rstll.Edit
          rstll!mtrsassignats = cadbl(rstll!mtrsassignats) + cadbl(rstmpc!smetres)
         rstll.Update
     End If
    End If
 '   vImp_o_Lam = IIf(rstm!seccio = "I", vsqlIMP, IIf(rstm!seccio = "L", vsqlLAM, ""))
 '   Set rstmpc = dbtmp.OpenRecordset(vImp_o_Lam)
 '   If Not rstmpc.EOF Then
 '    Set rstll = dbllistat.OpenRecordset("Select * from llistatinventari where metres>100 and preucompra=" + atrim(cadbl(rstm!numerogrup)))
 '    If Not rstll.EOF Then
 '        rstll.Edit
 '         rstll!mtrsassignats = cadbl(rstll!mtrsassignats) + cadbl(rstmpc!smetres)
 '        rstll.Update
 '    End If
 '   End If
    rstm.MoveNext
  Wend
 
  Set rstm = dbtmp.OpenRecordset("SELECT grupsdepalets.*, Palets.codimatprognou as codimat , palets.ample as ample ,palets.micres as espesor FROM grupsdepalets INNER JOIN Palets ON grupsdepalets.paletexemple = Palets.Idpalet;")
  While Not rstm.EOF
     Set rstmat = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(rstm!codimat))
       If Not rstmat.EOF Then
         nomfamilia = ""
         nomdelmaterial = ""
        Set rstfam = dbtmpb.OpenRecordset("select * from familiesmaterials where codi=" + atrim(cadbl(rstmat!familia)))
        If Not rstfam.EOF Then nomfamilia = rstfam!descripcio
        nomdelmaterial = descripciomaterial(rstmat)
        If InStr(1, nomdelmaterial, "EVOH") > 0 Then nomfamilia = Trim(nomfamilia) + "+EVOH"
        nomdelmaterial = nomdelmaterial + " " + atrim(cadbl(rstm!ample)) + " Cm " + atrim(cadbl(rstm!espesor)) + " Micres"
        dbllistat.Execute " update llistatinventari set familia='" + treure_apostruf(nomdelmaterial) + "' where preucompra=" + atrim(rstm!numerogrup)
       End If
    rstm.MoveNext
  Wend
  Set rstll = Nothing
  Set rstm = Nothing
  Set rstmpc = Nothing
  Set rstmat = Nothing
End Sub

Sub possoelsnomsalsmaterials()
  Dim rstfam As Recordset
  Dim rsttmp As Recordset
  Dim rstmat As Recordset
  Dim nomdelmaterial As String
  Dim ultimpalet As Double
'posso els noms als materials
   Set rsttmp = dbllistat.OpenRecordset("SELECT distinct llistatinventari.codimat, llistatinventari.palet From llistatinventari GROUP BY llistatinventari.codimat, llistatinventari.palet;")
   While Not rsttmp.EOF
      If rsttmp!codimat <> ultimpalet Then
       Set rstmat = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(rsttmp!codimat))
       If Not rstmat.EOF Then
         If rstmat!proveidor = 581 Then
            eliminaraquestmaterialdelllistat cadbl(rsttmp!codimat): GoTo seguir
         End If
         nomfamilia = ""
         nomdelmaterial = ""
        Set rstfam = dbtmpb.OpenRecordset("select * from familiesmaterials where codi=" + atrim(cadbl(rstmat!familia)))
        If Not rstfam.EOF Then nomfamilia = rstfam!descripcio
        nomdelmaterial = descripciomaterial(rstmat)
        If InStr(1, nomdelmaterial, "EVOH") > 0 Then nomfamilia = Trim(nomfamilia) + "+EVOH"
        dbllistat.Execute "update llistatinventari set nommaterial ='" + atrim(nomdelmaterial) + "', familia='" + atrim(nomfamilia) + "'  where codimat=" + atrim(cadbl(rsttmp!codimat))
seguir:
       End If
       ultimpalet = rsttmp!codimat
      End If
      
      rsttmp.MoveNext
   Wend
   Set rstfam = Nothing
   Set rsttmp = Nothing
End Sub
Sub eliminaraquestmaterialdelllistat(codimatperborrar As Double)
  dbllistat.Execute "delete * from llistatinventari where codimat=" + atrim(codimatperborrar)
End Sub
Sub passodemetresakilos()
   Dim grmm2 As Double
   Dim kilos As Double
   Dim espesor As Double
   Dim rsttmp As Recordset
   Dim metres As Double
   Dim rstp As Recordset
'passo els metres a kilos
   Set rsttmp = dbllistat.OpenRecordset("SELECT * from llistatinventari ")
   If Not rsttmp.EOF Then rsttmp.MoveLast: rsttmp.MoveFirst
   'MsgBox rsttmp.RecordCount
   'MsgBox rsttmp.AbsolutePosition
   While Not rsttmp.EOF
      Set rstp = dbtmp.OpenRecordset("select * from palets where idpalet=" + atrim(rsttmp!palet))
      kilos = 0
      If Not rstp.EOF Then
        espesor = cadbl(rstp!micres)
        If rstp!grmsm2 > 0 And espesor = 0 Then
           espesor = rstp!grmsm2 * -1
        End If
        metres = rsttmp!metres
        If metres < 0 Then metres = 0
        kilos = compramat.conversiokilos(cadbl(rstp!codimatprognou), rstp!ample, metres, espesor, atrim(rstp!semielaborat), cadbl(rstp!solapa))
      End If
      If kilos = 0 Then
         kilos = (cadbl(rsttmp!metres) * cadbl(rsttmp!kilos)) / cadbl(rsttmp!metresbob)
      End If
      kilos = Redondejar(kilos, 0)
      dbllistat.Execute "update llistatinventari set kilos =" + atrim(kilos) + ", espesor=" + passaradecimalpunt(atrim(espesor)) + " where metres=" + atrim(rsttmp!metres) + " and palet=" + atrim(rsttmp!palet) + " and bobina=" + atrim(rsttmp!bobina)
      
      rsttmp.MoveNext
   Wend
   Set rsttmp = Nothing
   Set rstp = Nothing
End Sub

Function buscarbobinesentdataentregaavui() As String
   Dim rst As Recordset
   Dim v As String
   Set rst = dbbaixes.OpenRecordset("select distinct comanda from bobinesent where format(data,'dd/mm/yy')=format(now,'dd/mm/yy')")
   While Not rst.EOF
     v = v + IIf(v <> "", ",", "") + atrim(rst!comanda)
     rst.MoveNext
   Wend
   buscarbobinesentdataentregaavui = v
   Set rst = Nothing
End Function

Private Sub llistatsensemoviment_Click()
   llistatsensemovimententredates
End Sub
Sub llistatsensemovimententredates()
  Dim consulta As String
  Dim rstc As Recordset
  Dim rst As Recordset
  Dim rutastocks As String
  Dim vinici As String
  vinici = InputBox("Entra la data limit de la consulta" + Chr(10) + " Bobines sense moviment abans de ..." + Chr(10) + "A LA VISTA PREVIA PODEU EXPORTAR AMB EXCEL DATA ONLY", "Data inici")
  If Not IsDate(vinici) Then MsgBox "Data no vàlida", vbCritical, "Error": Exit Sub
 
  borrartaulallistatinventari
  ratoli "espera"
  creartaulatempllistatbobinessensemoviment
  
  Set rstc = dbllistat.OpenRecordset("select * from llistatinventari")
  Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "Palets.mdb")
  While Not rstc.EOF
    Set rst = dbtmp.OpenRecordset("select * from parcials where idpalet=" + atrim(rstc!idpalet) + " and idbobina=" + atrim(rstc!idbobina) + " order by data desc")
    If Not rst.EOF Then
      If Not IsNull(rst!data) Then
        If CVDate(rst!data) >= CVDate(vinici) Then rstc.Delete: GoTo proxim
        rstc.Edit
        rstc!dataaltapalet = rst!data
        rstc.Update
      End If
    End If
    
    If IsNull(rstc!dataaltapalet) Then rstc.Delete: GoTo proxim
    If CVDate(rstc!dataaltapalet) >= CVDate(vinici) Then rstc.Delete
proxim:
    rstc.MoveNext
  Wend
  wait 3
  rutastocks = llegir_ini("General", "ruta_stocks", "comandes.ini")
  'dbllistat.Execute "delete * from llistatinventari where dataaltapalet=null"
  'dbllistat.Execute "delete * from llistatinventari where  (((llistatinventari.dataaltapalet)<#" + Format(vinici, "m/d/yy") + "#))"
  
  wait 2
  Set rstc = Nothing
   For i = 1 To 50
    llistat.Formulas(i) = ""
   Next i
  'faig el llistat
  llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistatbobinessensemovimententredates.rpt"
  llistat.Destination = crptToWindow
  llistat.CopiesToPrinter = 1
  llistat.DataFiles(0) = nomfitxertemporal
  llistat.DiscardSavedData = True
  llistat.Formulas(0) = "filtre=' " + format(vinici, "dd/mm/yy") + "'"
  llistat.Formulas(1) = ""
  llistat.Formulas(3) = ""
  llistat.Formulas(4) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 dbstocks.Close
 dbllistat.Close
 obrir_dbllistats
 ratoli "normal"
End Sub

Private Sub llistestocxrentregar_Click()
  Dim consulta As String
  Dim rstc As Recordset
  Dim comandesa02 As String
  Dim comandesa0 As String
   ratoli "espera"
   borrartaulallistatinventari
   Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes ", fitxerini))
   vllistacomandesdavui = buscarbobinesentdataentregaavui
   consulta = "SELECT comandes.comanda, COMANDES.ampleesq,comandes.refclient,comandes.proximaseccio, comandes.client,comandes.producte, productes.ruta, comandes.datacomanda, comandes.pvp, comandes.mesurapvp, [unitatinterna] , comandes.texteimpressio"
   consulta = consulta + " into llistatinventari in '" + nomfitxertemporal + "'" ''c:\temporal.mdb' "
   consulta = consulta + "FROM (comandes INNER JOIN mesures ON comandes.mesurapvp = mesures.codi) INNER JOIN productes ON comandes.producte = productes.codi WHERE ((((comandes.proximaseccio)='V' Or (comandes.proximaseccio)='P') AND ((comandes.producte)<>'PC' And (comandes.producte)<>'PC2')) )" + IIf(vllistacomandesdavui <> "", "or comandes.comanda in (" + vllistacomandesdavui + ");", "")
      
   dbtmpb.Execute consulta
   dbllistat.Execute "alter table llistatinventari add column metresfabricats double"
   dbllistat.Execute "alter table llistatinventari add column kilosfabricats double"
   dbllistat.Execute "alter table llistatinventari add column kilosentregats double"
   dbllistat.Execute "alter table llistatinventari add column metresentregats double"
   dbllistat.Execute "alter table llistatinventari add column pecesfabricades double"
   dbllistat.Execute "alter table llistatinventari add column unitatspvp double"
   dbllistat.Execute "alter table llistatinventari add column nomclient string"
   dbllistat.Execute "alter table llistatinventari add column pes1000mtrs double"
   dbllistat.Execute "alter table llistatinventari add column kiloscomanda double"
   dbllistat.Execute "alter table llistatinventari add column metrescomanda double"
   dbllistat.Execute "alter table llistatinventari add column unitatsxmetre double"
   dbllistat.Execute "alter table llistatinventari add column nomfamilies string"
   dbllistat.Execute "alter table llistatinventari add column pecescomanda double"
   dbllistat.Execute "alter table llistatinventari add column pecesvenudes double"
   dbllistat.Execute "alter table llistatinventari add column dessarroll double"
   dbllistat.Execute "alter table llistatinventari add column rutareal string"
   dbllistat.Execute "alter table llistatinventari add column entregaavui byte"
   
   
   Set rstc = dbllistat.OpenRecordset("llistatinventari")
   
   comandesa0 = ""
   comandesa02 = ""
   While Not rstc.EOF
     possarcampsextres rstc
     If InStr(1, vllistacomandesdavui, atrim(rstc!comanda)) > 0 Then
       rstc.Edit
       rstc!entregaavui = 1
       rstc.Update
     End If
     If cadbl(rstc!pvp) = 0 Then
        If Len(comandesa0) < 240 Then
           comandesa0 = comandesa0 + "[" + atrim(rstc!comanda) + "] "
          Else: comandesa02 = comandesa02 + "[" + atrim(rstc!comanda) + "] "
        End If
     End If
     rstc.MoveNext
   Wend
   If comandesa0 <> "" Then
      If MsgBox("He trobat les següents comandes amb PVP a zero." + Chr(10) + Chr(13) + comandesa0 + " " + comandesa02 + " VOLS CONTINUAR EL LLISTAT?", vbInformation + vbYesNo, "Atenció") = vbNo Then
         GoTo fi
      End If
   End If
   llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistatproducteacabat.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 'llistat.Formulas(0) = "nomllistat='" + treure_apostruf(nomllistat) + "'"
 llistat.Formulas(0) = "hora='" + format(Now, "dd/mm/yy  hh:nn") + "'"
 llistat.Formulas(1) = "comandesa0='" + Mid(IIf(comandesa0 <> "", "Comandes PVP=0: ", "") + comandesa0, 1, 253) + "'"
 llistat.Formulas(2) = "comandesa02='" + Mid(comandesa02, 1, 253) + "'"
 If Len(vllistacomandesdavui) > 250 Then vllistacomandesdavui = Mid(vllistacomandesdavui, 1, 200) + "..."
 llistat.Formulas(3) = "comandesambdatadavui='" + vllistacomandesdavui + "'"
 
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
fi:
 dbllistat.Close
 obrir_dbllistats
 ratoli "normal"
 Set dbbaixes = Nothing
End Sub

Function comprovarlaseccioenruta(numc As Double) As String
  Dim posicioruta As String
  Dim rsttmp As Recordset
    Set rsttmp = dbtmp.OpenRecordset("SELECT comandes.proximaseccio,comandes.linkcomanda1,comandes.linkcomanda2, comandes.producte,productes.ruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (comandes.comanda)=" + atrim(numc))
    If rsttmp.EOF Then Exit Function
    If rsttmp!producte = "PC" Or rsttmp!producte = "PC2" Then Exit Function
    posicioruta = posicioenlaruta(numc, rsttmp!proximaseccio, rsttmp!ruta)
    If posicioruta <> "" Then
        comprovarlaseccioenruta = posicioruta
        'dbllistat.Execute "update llistatinventari set rutareal='" + posicioruta + "' where comanda=" + atrim(cadbl(rsttmp!linkcomanda1)) + " or comanda=" + atrim(cadbl(rsttmp!linkcomanda2))
      Else: comprovarlaseccioenruta = ""
    End If
    
    Set rsttmp = Nothing
End Function
Function posicioenlaruta(numc As Double, seccioactual As String, laruta As String) As String
  Dim rstp As Recordset
  If InStr(1, "VPT", seccioactual) = 0 Then Exit Function
  Set rstp = dbbaixes.OpenRecordset("SELECT comandes.comanda, rebobinadorestot.acavada as acavadar, laminadorestot.acavada as acavadal, impressorestot.acavada as acavadai FROM ((comandes LEFT JOIN rebobinadorestot ON comandes.comanda = rebobinadorestot.comanda) LEFT JOIN laminadorestot ON comandes.comanda = laminadorestot.comanda) LEFT JOIN impressorestot ON comandes.comanda = impressorestot.comanda WHERE (((comandes.comanda)=" + atrim(numc) + "));")
  
  If Not rstp.EOF Then
     If InStr(1, laruta, "R") > 0 And cadbl(rstp!acavadar) = 0 Then posicioenlaruta = "R"
     If InStr(1, laruta, "L") > 0 And cadbl(rstp!acavadal) = 0 Then posicioenlaruta = "L"
     If InStr(1, laruta, "I") > 0 And cadbl(rstp!acavadai) = 0 Then posicioenlaruta = "I"
  End If
  
  Set rstp = Nothing
End Function
Sub possarcampsextres(rstc As Recordset)
   Dim rstcli As Recordset
   Dim rstcom As Recordset
   Dim rstcomextres As Recordset
   Dim metres As Double
   Dim kilos As Double
   Dim kilosentregats As Double
   Dim ultimaseccio As String
   Dim pesde1000 As Double
   Dim metrescomanda As Double
   Dim dessarrollcomanda As Double
   Dim rstmat As Recordset
   Dim solpesgrm2 As Double
   
   If atrim(rstc!nomclient) = "" Then
     Set rstcli = dbtmpb.OpenRecordset("select nom from clients where codi=" + atrim(rstc!client))
     If Not rstcli.EOF Then dbllistat.Execute "update  llistatinventari set nomclient='" + treure_apostruf(rstcli!nom) + "' where client=" + atrim(rstc!client)
     Set rstcli = Nothing
   End If
   Set rstcom = dbtmp.OpenRecordset("select amplereb,migelaboratsol,cantitatsol,amplesol,longitudsol,solapasol,cantitatex,materialex,rebmtrs,tubolam,simulteneitatsol from comandes where comanda=" + atrim(rstc!comanda))
   Set rstcomextres = dbtmp.OpenRecordset("select  solpesgrmcm2 from comandes_extres where comanda=" + atrim(rstc!comanda))
   metrescomanda = 0
   dessarrollcomanda = calculardesarrollcomanda(rstc!comanda)
   If Not rstcom.EOF Then
      If InStr(1, rstc!ruta, "R") > 0 And cadbl(rstcom!rebmtrs) > 0 Then
         metrescomanda = cadbl(rstcom!rebmtrs)
           Else: metrescomanda = cadbl(rstcom!cantitatex) '* IIf(atrim(rstcom!tubolam) = "T", 2, 1)
      End If
      Set rstmat = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(rstcom!materialex))
       Else: Exit Sub
   End If
   kilos = 0
   metres = 0
   metres = calcularmetrescomandaproduits(rstc!comanda, rstc!ruta, ultimaseccio)
   If InStr(1, rstc!ruta, "S") = 0 Then kilos = calcularkiloscomandaproduits(rstc!comanda, metres, InStr(1, rstc!ruta, "R"))
   kilosentregats = format(calcularkilosentregats(rstc!comanda), "#,##0.0")
   rstc.Edit
   If InStr(1, rstc!ruta, "R") > 0 Then rstc!ampleesq = rstcom!amplereb
   rstc!pecesvenudes = 0
   rstc!metresentregats = calcularmetresentregats(rstc!comanda)
   rstc!kilosentregats = format(cadbl(kilosentregats), "#,##0.0")
   rstc!nomfamilies = descripciomaterial(rstmat)
   rstc!metrescomanda = format(metrescomanda, "#,##0")
   rstc!metresfabricats = format(cadbl(metres), "#,##0")
   rstc!kilosfabricats = format(cadbl(kilos), "#,##0.0")
   rstc!dessarroll = dessarrollcomanda
   If dessarrollcomanda > 0 Then
      rstc!unitatsxmetre = 1 / dessarrollcomanda
     Else: rstc!unitatsxmetre = 0
   End If
   rstc!pecescomanda = 0
   rstc!rutareal = comprovarlaseccioenruta(rstc!comanda)
   
   
   If ultimaseccio = "S" Then
       ' que son les peces de baixes d'entrega de soldadora ho guardo per operar x metres tres linies a sota
      rstc!unitatsxmetre = rstc!metresfabricats '/ IIf(rstcom!simulteneitatsol = 0, 1, rstcom!simulteneitatsol)
      rstc!pecesfabricades = rstc!metresfabricats ' a soldadora els metres entregats son peces
      rstc!pecesvenudes = calcularmetresentregats(rstc!comanda)  'a soldadora els kilos entregats son peces
      rstc!pecescomanda = cadbl(rstcom!cantitatsol)
      rstc!pes1000mtrs = calcularpescomanda(rstc!comanda)
      rstc!metresfabricats = 0
      rstc!metresentregats = 0
      rstc!unitatsxmetre = 0
      solpesgrm2 = IIf(Not rstcomextres.EOF, rstcomextres!solpesgrmcm2, 0)
      rstc!kilosfabricats = (solpesgrm2 * ((rstcom!amplesol + rstcom!solapasol) * rstcom!longitudsol))
      rstc!kilosfabricats = (IIf(rstcom!migelaboratsol = "L", 1, 2) * rstc!pecesfabricades) * rstc!kilosfabricats
      If rstc!pecesfabricades > 0 Then
         rstc!kilosentregats = rstc!pecesvenudes * rstc!kilosfabricats / rstc!pecesfabricades
        Else: rstc!kilosentregats = 0
      End If
        Else:
           If dessarrollcomanda > 0 Then
              rstc!pecescomanda = rstc!metrescomanda / dessarrollcomanda
             Else: rstc!pecescomanda = 0
           End If
           If dessarrollcomanda > 0 Then
              rstc!pecesfabricades = rstc!metresfabricats / dessarrollcomanda
                Else: rstc!pecesfabricades = 0
           End If
           
   End If
   '''*''
   If ultimaseccio <> "S" Then
     If rstc!metresfabricats > 0 Then
       pesde1000 = (1000 * rstc!kilosfabricats) / rstc!metresfabricats
        Else: pesde1000 = 0
     End If
     rstc!pes1000mtrs = pesde1000
   End If
   If rstc!metresentregats = 0 And ultimaseccio <> "S" Then
    If rstc!pes1000mtrs > 0 Then
       rstc!metresentregats = format((rstc!kilosentregats / rstc!pes1000mtrs) * 1000, "#,##0")
      Else: rstc!metresentregats = 0
    End If
   End If
   If ultimaseccio <> "S" Then
        If dessarrollcomanda > 0 Then
              rstc!pecesvenudes = rstc!metresentregats / dessarrollcomanda
                Else: rstc!pecesvenudes = 0
        End If
   End If
   rstc!kiloscomanda = format((rstc!metrescomanda / 1000) * rstc!pes1000mtrs, "#,##0")
   
   If rstc!unitatinterna = "€/U" Then rstc!unitatspvp = rstc!pecesfabricades - cadbl(rstc!pecesvenudes)
   If rstc!unitatinterna = "€/1000U" Then rstc!unitatspvp = (rstc!pecesfabricades - cadbl(rstc!pecesvenudes)) / 1000
   If rstc!unitatinterna = "€/K" Then rstc!unitatspvp = cadbl(rstc!kilosfabricats) - cadbl(rstc!kilosentregats)
   If pesde1000 > 0 Then
    If rstc!unitatinterna = "€/KM" Then rstc!unitatspvp = (((cadbl(rstc!kilosfabricats) - cadbl(rstc!kilosentregats)) * 1000) / pesde1000) / 1000
    If rstc!unitatinterna = "€/M" Then rstc!unitatspvp = ((cadbl(rstc!kilosfabricats) - cadbl(rstc!kilosentregats)) * 1000) / pesde1000
   End If
   If rstc!unitatinterna = "€/FIX" Then rstc!unitatspvp = 1
   If rstc!unitatinterna = "€/M2" Then rstc!unitatspvp = (rstc!metresfabricats - rstc!metresentregats) * (rstc!ampleesq / 100) '(((rstc!kilosfabricats) - cadbl(rstc!kilosentregats)) * 1000) / (pesde1000 / 1000) * (rstc!ampleesq / 100)
   rstc!unitatinterna = LCase(rstc!unitatinterna)
   'falta € bobina
   rstc.Update
   Set rstmat = Nothing
   Set rstcli = Nothing
   Set rstcom = Nothing
End Sub
Function calcularmetrescomanda(numc As Double) As Double
  Dim rstm As Recordset
  Set rstm = dbtmp.OpenRecordset("select sum(metres) as summtrs from parcials where utilitzada and comanda='" + atrim(numc) + "'")
  If Not rstm.EOF Then calcularmetrescomanda = cadbl(rstm!summtrs)
End Function
Function calcularpescomanda(numc As Double) As Double
   Dim rstc As Recordset
   Set rstc = dbtmp.OpenRecordset("select linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(numc))
   If Not rstc.EOF Then
     Set rstc = dbtmp.OpenRecordset("select pes1000mtrs,comanda from comandes where comanda=" + atrim(numc) + " or comanda=" + atrim(cadbl(rstc!linkcomanda1)) + " or comanda=" + atrim(cadbl(rstc!linkcomanda2)))
     calcularpescomanda = 0
     While Not rstc.EOF
        If rstc!comanda > 0 Then
           calcularpescomanda = calcularpescomanda + cadbl(rstc!pes1000mtrs)
        End If
        rstc.MoveNext
     Wend
   End If
   Set rstc = Nothing
End Function
Function calculardesarrollcomanda(numc As Double) As Double
  Dim rstu As Recordset
  Set rstu = dbtmp.OpenRecordset("SELECT productes.ruta, comandes.dessarroll, comandes.longitudsol FROM comandes INNER JOIN productes ON comandes.producte = productes.codi where comanda=" + atrim(numc) + ";")
  If Not rstu.EOF Then
     If InStr(1, rstu!ruta, "I") Then
        If cadbl(rstu!dessarroll) > 0 Then
           calculardesarrollcomanda = (cadbl(rstu!dessarroll) / 1000)
        End If
      'Else:
      '  If InStr(1, rstu!ruta, "S") Then
      '     If cadbl(rstu!longitudsol) > 0 Then calcularunitatsxmetre = 100 / cadbl(rstu!longitudsol)
      '  End If
     End If
  End If
  
End Function

Function calcularkilosentregats(numc As Double) As Double
  Dim rstk As Recordset
  Set rstk = dbbaixes.OpenRecordset("Select sum(kilosiunitats) as tkilos from bobinesent where entregat='S' and data<>null and FORMAT(data,'dd/mm/yy')<>format(now,'dd/mm/yy') and comanda=" + atrim(numc))
  calcularkilosentregats = 0
  If Not rstk.EOF Then
      calcularkilosentregats = cadbl(rstk!tkilos)
  End If
  Set rstk = Nothing
  
End Function

Function calcularmetresentregats(numc As Double, Optional tots As Boolean) As Double
  Dim rstk As Recordset
  Set rstk = dbbaixes.OpenRecordset("Select sum(metresisacs) as tmetres from bobinesent where " + IIf(Not tots, "entregat='S' and  data<>null and FORMAT(data,'dd/mm/yy')<>format(now,'dd/mm/yy') and ", "") + " comanda=" + atrim(numc))
  'Set rstk = dbbaixes.OpenRecordset("Select sum(metresisacs) as tmetres from bobinesent where " + IIf(Not tots, "entregat='S' and  data<>null and ", "") + " comanda=" + atrim(numc))
  calcularmetresentregats = 0
  If Not rstk.EOF Then
      calcularmetresentregats = cadbl(rstk!tmetres)
  End If
  Set rstk = Nothing
  
End Function

Function calcularkiloscomandaproduits(numc As Double, metres As Double, hihar As Byte) As Double
   Dim rstm As Recordset
   Dim rstl As Recordset
   Dim sim As Byte
   If hihar > 0 Then
     Set rstl = dbbaixes.OpenRecordset("select tkilos from rebobinadorestot where comanda=" + atrim(numc))
     If Not rstl.EOF Then calcularkiloscomandaproduits = rstl!tkilos
     GoTo fi
   End If
   Set rstl = dbtmpb.OpenRecordset("select linkcomanda1,linkcomanda2,simulteneitatreb,simulteneitatlam from comandes where comanda=" + atrim(numc))
   If Not rstl.EOF Then
      sim = IIf(cadbl(rstl!simulteneitatreb) = 0, cadbl(rstl!simulteneitatlam), cadbl(rstl!simulteneitatreb))
      If sim = 0 Then sim = 1
      Set rstm = dbtmpb.OpenRecordset("select sum(pes1000mtrs) as kilos from comandes where comanda=" + atrim(numc) + " or comanda=" + atrim(cadbl(rstl!linkcomanda1)) + " or comanda=" + atrim(cadbl(rstl!linkcomanda2)))
      If Not rstm.EOF Then calcularkiloscomandaproduits = ((metres / sim) / 1000) * cadbl(rstm!kilos)
   End If
   Set rstl = Nothing
   Set rstm = Nothing
fi:
End Function
Function calcularmetrescomandaproduits(numc As Double, ruta As String, ultimaseccio As String) As Double
 
  Dim seccions As Variant
  Dim seccionsbob As Variant
  Dim ordre As String
  Dim nomtaula As String
  Dim nomsubtaula As String
  Dim idscontrol As String
  Dim rstbob As Recordset
  Dim metres
  
  
  ordre = "EILRS"
  seccions = Array("extrussores", "impressores", "laminadores", "rebobinadores", "soldadores")
  seccionsbob = Array("Bobinesext", "bobinesimp", "bobineslam", "bobinesreb", "bobinessol")
  If ruta <> "" Then ultimaseccio = Mid(ruta, Len(ruta), 1)
  nomtaula = seccions(InStr(1, ordre, ultimaseccio) - 1)
  nomsubtaula = seccionsbob(InStr(1, ordre, ultimaseccio) - 1)
  nomordre = "numerodebobina"
  If nomsubtaula = "bobinessol" Then nomordre = "numerodesac"
  'Miro tots els registres de la taula principal per fer la busqueda a les bobines(subtaula)
  Set rsttmp = dbbaixes.OpenRecordset("select * from " + nomtaula + " where comanda=" + atrim(numc))
  While Not rsttmp.EOF
     If idscontrol <> "" Then
         idscontrol = idscontrol + " or controlid=" + atrim(cadbl(rsttmp!id))
        Else: idscontrol = " controlid=" + atrim(cadbl(rsttmp!id))
     End If
     rsttmp.MoveNext
  Wend
  
  'Faig la busqueda de la subtaula i les entro a bobinesent
  If idscontrol <> "" Then
   Set rsttmp = dbbaixes.OpenRecordset("select * from " + nomsubtaula + " where " + idscontrol + " order by " + nomordre + " ASC")
   While Not rsttmp.EOF
     If ultimaseccio = "S" Then
            metres = metres + cadbl(rsttmp!unitatsxsac)
         Else
              metres = metres + cadbl(rsttmp!metres)
     End If
      rsttmp.MoveNext
   Wend
  End If
  
  calcularmetrescomandaproduits = metres
  
End Function
Private Sub m_grupdepalets_Click()
  ' Load formaltarep
  'formaltarep.Caption = "Manteniment de Descripcio grup de palets"
  'formaltarep.Data1.DatabaseName = camistock
  'formaltarep.Data1.RecordSource = "select * from grupsdepalets order by numerogrup"
  'formaltarep.refrescar
  'formaltarep.DBGrid1.Refresh
  'formaltarep.DBGrid1.Columns(0).Width = 1250
  'formaltarep.DBGrid1.Columns(1).Width = 6500
  'formaltarep.Width = 8800
  'formaltarep.Show
  grupdepalets.Show 1
End Sub

Sub imprimiretpalet(numpalet As Double, finumpalet As Double, Optional numbobina As Double, Optional noimpres As Boolean, Optional vnodemanarperprepararimpresora As Boolean)
  Dim rstpalet As Recordset
  Dim rstpro As Recordset
  Dim rstbobina As Recordset
  Dim rstmaterial As Recordset
  Dim rstparcials As Recordset
  Dim nomimpresora As String
  Dim X As Printer
  Dim i As Byte
  
  nomimpresora = llegir_ini("Expedicions", "nomimpresoraetiquetes", fitxerini)
  For Each X In Printers
     If nomimpresora = X.DeviceName Then GoTo cont
  Next
  MsgBox "Has d'escullir primer la impresora de etiquetes, en el menu etiqueta-palet", vbCritical, "Error"
  Exit Sub
cont:
  If nomimpresora <> X.DeviceName Then MsgBox "Impresora no trobada": Exit Sub
  'numpalet = cadbl(InputBox("Entra el numero de palet que vols imprimir", "Etiqueta Palet"))
  'If numpalet < 1 Then Exit Sub
'  numpalet = palets.Recordset!idpalet

  obrir_dbllistats
  
  crear_taules_tmp
  
  Set rstllistat = dbllistat.OpenRecordset("etiquetapalet")
  
  If noimpres Then
  
    Set rstpalet = dbtmp.OpenRecordset("SELECT Palets.*, Bobines.impres FROM Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet WHERE (((Bobines.impres)=False));")
   Else: Set rstpalet = dbtmp.OpenRecordset("select * from Palets where idpalet>=" + atrim(numpalet) + " and idpalet<=" + atrim(finumpalet))
  End If
  
  While Not rstpalet.EOF
     Set rstbobina = dbtmp.OpenRecordset("select * from bobines where idpalet=" + atrim(cadbl(rstpalet!idpalet)) + IIf(numbobina > 0, " and idbobina=" + atrim(numbobina), ""))
     Set rstmaterial = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstpalet!codimatprognou)))
     If Not rstmaterial.EOF Then Set rstpro = dbtmpb.OpenRecordset("select nom from proveidors where codi=" + atrim(cadbl(rstmaterial!proveidor)))
     While Not rstbobina.EOF
       guardar_registre_taulatmp3 rstpalet, rstpro, rstbobina, rstmaterial
       rstbobina.MoveNext
     Wend
     rstpalet.MoveNext
     
  Wend
  

  dbllistat.Close
  
  If vnodemanarperprepararimpresora Then If MsgBox("Prepara la impresora i prem Sí per començar la impresió", vbInformation + vbYesNo, "La impresora està apunt?") <> vbYes Then r = "noimpres": Exit Sub
   'imprimir llistat
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "novaetiquetapaletssensetalls.rpt" '"NOVAetiquetapalets.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 llistat.Formulas(1) = ""
 llistat.Formulas(0) = ""
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 llistat.Formulas(4) = ""
 llistat.Formulas(5) = ""
 llistat.Formulas(6) = ""
 llistat.Formulas(7) = ""
 llistat.Formulas(8) = ""
 llistat.Formulas(9) = ""
 llistat.Formulas(10) = ""
 llistat.Formulas(11) = ""
 llistat.Formulas(12) = ""
 llistat.Formulas(13) = ""

 'llistat.PrinterDriver = X.DriverName
 'llistat.PrinterName = X.DeviceName
 'llistat.PrinterPort = X.Port
 llistat.PrinterSelect
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 'Set dbllistat = Nothing
 Set rstllistat = Nothing
  'Set dbtmp = Nothing
  'Set dbtmpb = Nothing
End Sub

Function copiafoto(foto As String, fldTO As Field)

'This function takes the source field image and copies it
'into the destination field.
'The function first saves the image in the source field to a
'temp file on disc. Then reads this temp file into
'the destination field.
'The temp file is then deleted
'On Error Resume Next

Dim iFieldSize  As Long
Dim varChunk    As Variant
Dim baData()    As Byte
Dim iOffset     As Long
Dim sFName      As String
Dim iFileNum    As Long
Dim cnt         As Long
Dim z()         As Byte

Const CONCHUNKSIZE As Long = 16384

Dim iChunks As Long
Dim iFragmentSize As Long
    
    'Get a unique random filename
    If Not existeix(foto) Then Exit Function
    sFName = foto
    
    Open sFName For Binary Access Read As #1
    ReDim z(FileLen(sFName))
    Get #1, , z()
     fldTO.AppendChunk z
    Close #1
    
    'Delete the file
    'Kill (sFName)
    
End Function



Private Sub mllistarperpantalla_Click(Index As Integer)

End Sub

Private Sub m_grupsdecompatibles_Click()
   grupmaterialscompatibles.Show 1
End Sub

Private Sub m_recepciomat_Click()
  If palets.Recordset.EditMode > 0 Then MsgBox "Primer acaba la edició d'aquest palet": Exit Sub
    If Not comprovaraccessabip Then MsgBox "No tens access al servidor del SAP no es pujarant les linies"
  comprespalets.Show 1
End Sub

Private Sub mactivaciodelspalets_Click()
  comprovarpaletssensedatadactivacio
End Sub

Private Sub mactivaciodepalets_Click()
    comprovarpaletsnodisponibles
  ' Dim numalbc As String
  ' Dim rstbob As Recordset
  ' Dim datarecepcio As String
   '
   'numalbc = InputBox("Entra el numero d'albarà del proveïdor", "Activació de Palets")
   'If atrim(numalbc) = "" Then Exit Sub
   'datarecepcio = (InputBox("Entra la data de recepció del material.", "Activació de Palets"))
 '  If Not IsDate(datarecepcio) Then MsgBox "Format de la data incorrecte.", vbCritical + vbOKOnly, "Atenció": Exit Sub
 '  Set rstbob = dbtmp.OpenRecordset("SELECT Count(Bobines.Idbobina) AS bobs, Palets.Numalb, Palets.Disponible FROM Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet GROUP BY Palets.Numalb, Palets.Disponible HAVING (((Palets.Numalb)='" + numalbc + "') AND ((Palets.Disponible)=False));")
 '  If Not rstbob.EOF Then
 '     If MsgBox("S'han trobat " + atrim(rstbob!bobs) + " amb el numero d'albarà " + numalbc + Chr(10) + Chr(13) + "Vols passar les bobines a disponibles?", vbExclamation + vbYesNo, "Activar bobines") = vbYes Then
 '        dbtmp.Execute "UPDATE Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet SET Palets.Disponible = True, datarec=#" + Format(datarecepcio, "yy/mm/dd") + "#  WHERE (((Palets.Numalb)='" + numalbc + "') AND ((Palets.Disponible)=False));"
 '     End If
 '     palets.RecordSource = "select * from palets where numalb='" + atrim(numalbc) + "'"
 '     palets.Refresh
 '       Else: MsgBox "No s'ha trobat cap coincidencia", vbCritical, "Atenció": GoTo fi
 '  End If
'fi:
'   Set rstbob = Nothing
End Sub

Private Sub matproveidor_Click()
   
End Sub

Private Sub matarpicos_Click()
   Dim rst As Recordset
   Dim nump As Double
   Dim numb As Double
   Dim numc As Double
   Dim mtrsb As Double
   Set dbstocks = dbtmp
   If UCase(InputBox("Escriu [matarpicos] per eliminar tots els picos mes <500.", "Atenció")) <> "MATARPICOS" Then Exit Sub
   Set rst = dbstocks.OpenRecordset("SELECT Bobines.Idpalet, Bobines.Idbobina, Bobines.disponible From bobines WHERE (((Bobines.disponible)<500 And (Bobines.disponible)>0));")
   While Not rst.EOF
    nump = rst!idpalet
    numb = rst!idbobina
    numc = 100
    mtrsb = bobinesdentrada.calcular_mtrsdispreals(nump, numb)
    If mtrsb < 500 And mtrsb > 0 Then
        dbstocks.Execute "insert into parcials (idpalet,idbobina,metres,comanda,orcomassignacio,operari,data,seccio,utilitzada) values (" + atrim(nump) + "," + atrim(numb) + "," + atrim(mtrsb) + "," + atrim(numc) + "," + atrim(numc) + ",0,#" + format(Now, "mm/dd/yy") + "#,'I',true)"
        bobinesdentrada.actualitzar_metres_disponibles nump, numb
    End If
    Me.Caption = atrim(nump) + "-" + atrim(numb)
    DoEvents
    rst.MoveNext
   Wend
  MsgBox "Acavat."
End Sub

Private Sub matclient_Click()
If matclient.Value <> 0 And Not buscant Then preucompra = 0: txtFields(14) = 0
End Sub
Sub borrartaulallistatperpujar()
   On Error Resume Next
   dbllistat.Execute "drop table llistatperpujar"
   dbstocks.Execute "drop table llistatperbaixar"
   On Error GoTo 0
End Sub
Sub possarcomandesalallistaperbaixar()
   Dim rutaxrbaixar As String
   rutaxrbaixar = " llistatperbaixar IN '" + camistock + "'"
   dbbaixes.Execute "SELECT TOP 40 muntadoratot.comanda, Max(muntadores.datafi) AS mdatafi, muntadoratot.acabada into " + rutaxrbaixar + " FROM muntadoratot INNER JOIN muntadores ON muntadoratot.comanda = muntadores.comanda GROUP BY muntadoratot.comanda, muntadoratot.acabada Having (((muntadoratot.acabada) = True)) ORDER BY Max(muntadores.datafi) DESC;"
   dbbaixes.Execute "insert into " + rutaxrbaixar + " select comanda from muntadora_ordremuntatge"
End Sub

Private Sub mbaixaranivellar_Click()
   Dim rst As Recordset
   Dim i As Long
   If sestaactualitzanttorerus Then Exit Sub
   Set rst = dbtmp.OpenRecordset("SELECT Parcials.idpalet, Parcials.idbobina, First(Bobines.Sit) AS PrimeroDeSit FROM Bobines RIGHT JOIN Parcials ON (Bobines.Idpalet = Parcials.idpalet) AND (Bobines.Idbobina = Parcials.idbobina) Where (((parcials.comanda) = '2500')) GROUP BY Parcials.idpalet, Parcials.idbobina HAVING (((First(Bobines.Sit))<>'IMP' Or (First(Bobines.Sit))<>'imp'));")
   r = ""
   i = 0
   While Not rst.EOF
      imprimiretpalet rst!idpalet, rst!idpalet, rst!idbobina, , IIf(i = 0, True, False)
      i = i + 1
      If r = "noimpres" Then GoTo fi
      rst.MoveNext
   Wend
fi:
   escriure_ini "Torerus", "horaultimaactualitzacio", " ", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
End Sub

Private Sub mbobineslam_Click()
   If sestaactualitzanttorerus Then Exit Sub
   formmourebobines.Show 1
   escriure_ini "Torerus", "horaultimaactualitzacio", " ", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
End Sub
Sub llistat_parcials(Optional vnomessenseregularitzar As Boolean)
Dim rstforats As Recordset
   Dim rstbob As Recordset
   Dim rstparcials As Recordset
   Dim rstregularitzats As Recordset
   Dim bob As String
   Dim pres As String
   Dim metresdis As Double
   Dim vdataultimparcial As String
   Dim vdataultimaregularitzacio As String
   Dim contador As Double
   vtipus = UCase(InputBox("Quin tipus de llistat vols:" + vbNewLine + "[D] Detallat" + vbNewLine + "[R] Reduit", "Tipus de llistat", "R"))
   If vtipus <> "D" And vtipus <> "R" Then Exit Sub
   ratoli "espera"
   Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "Palets.mdb")
   obrir_dbllistats
   borrartaulallistatinventari
   dbtmp.Execute "select numlleixa+'A' as nlleixa,mid(numlleixa,1,1) as lletralleixa,cdbl(mid(numlleixa,2)) as lloclleixa, 'A' as foratlleixa into llistatinventari in '" + nomfitxertemporal + "' from prestatges"
   dbllistat.Execute "alter table llistatinventari add column bobines text"
   dbtmp.Execute "insert into llistatinventari in '" + nomfitxertemporal + "' select numlleixa+'B' as nlleixa,mid(numlleixa,1,1) as lletralleixa,cdbl(mid(numlleixa,2)) as lloclleixa, 'B' as foratlleixa from prestatges"
   dbtmp.Execute "insert into llistatinventari in '" + nomfitxertemporal + "' select numlleixa+'C' as nlleixa,mid(numlleixa,1,1) as lletralleixa,cdbl(mid(numlleixa,2)) as lloclleixa, 'C' as foratlleixa  from prestatges"
   Set rstforats = dbllistat.OpenRecordset("SELECT * from llistatinventari order by lletralleixa,lloclleixa,foratlleixa")
   contador = 0
   Set rstforats = dbtmp.OpenRecordset("select * from prestatgesnous")
   Open "c:\temp\llistatpicus.csv" For Output As #1
   If vtipus = "R" Then Print #1, "FORAT;BOBINES DINS EL FORAT"
   If vtipus = "D" Then Print #1, "FORAT;PALET;BOBINA;DIAMETRE(cm);DATA COMPRA;PROVEIDOR;FAMILIA MAT;SUBFAMILIA MAT;FAMILIA COLOR;SUBFAMILIA COLOR;AMPLE BOBINA(cm);Metres disponibles;Micres"
   If Not rstforats.EOF Then rstforats.MoveLast: rstforats.MoveFirst
   Set rstregularitzats = dbtmp.OpenRecordset("select * from comprovacio_diametres_picus")
   While Not rstforats.EOF
     Me.Caption = atrim(rstforats.AbsolutePosition) + "/" + atrim(rstforats.RecordCount)
     Set rstbob = dbtmp.OpenRecordset("select * from bobines where sit='" + atrim(rstforats!estanteria) + format(atrim(rstforats!columna), "00") + atrim(rstforats!fila) + "'")
     vdataultimparcial = ""
     vdataultimaregularitzacio = Date
     If Not rstbob.EOF Then
         Set rstparcials = dbtmp.OpenRecordset("select * from parcials where idpalet=" + atrim(rstbob!idpalet) + " and idbobina=" + atrim(rstbob!idbobina) + " order by data desc")
         If Not rstparcials.EOF Then vdataultimparcial = atrim(rstparcials!data)
         rstregularitzats.FindFirst "numpalet=" + atrim(rstbob!idpalet) + " and bobina=" + atrim(rstbob!idbobina)
         If Not rstregularitzats.NoMatch Then vdataultimaregularitzacio = atrim(rstregularitzats!data)
     End If
     bob = ""
     If vdataultimparcial <> "" And vdataultimaregularitzacio <> "" Then If DateDiff("d", vdataultimparcial, vdataultimaregularitzacio) <= 0 Then GoTo proxim
     While Not rstbob.EOF
       metresdis = bobinesdentrada.calcular_mtrsdispreals(rstbob!idpalet, rstbob!idbobina)
       If metresdis < rstbob!mts Then
            bob = bob + "[" + atrim(rstbob!idpalet) + "/" + atrim(rstbob!idbobina) + "](" + atrim(metresdis) + ") "
            If vtipus = "D" Then
             Print #1, DetallBobinaiPalet(rstbob!idpalet, rstbob!idbobina, metresdis)
            End If
            contador = contador + 1
       End If
       rstbob.MoveNext
     Wend
     If bob <> "" And vtipus = "R" Then Print #1, atrim(rstforats!estanteria) + format(atrim(rstforats!columna), "00") + atrim(rstforats!fila) + ";" + bob
     'rstforats.Edit
     'rstforats!bobines = Mid(bob, 1, 250)
     'rstforats.Update
proxim:
     rstforats.MoveNext
     DoEvents
   Wend
   Print #1, " "
   Print #1, "TOTAL PARCIALS:;" + atrim(contador) + ";BOBINES"
   Close 1
   Me.Caption = "Manteniment de Palets"
   If existeix("c:\temp\llistatpicus.csv") Then obrir_document "c:\temp\llistatpicus.csv"
   GoTo fi
   'borro tots els forats que no hi ha bobines
   dbllistat.Execute "delete * from llistatinventari where bobines=''"
   wait 2
     'faig el llistat
  llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "inventariperforat.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 
 llistat.Formulas(0) = "hora='" + format(Now, "dd/mm/yy  hh:nn") + "'"
 llistat.Formulas(1) = "titol='Llistat de bobines parcials per forat.'"
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 dbllistat.Close
 obrir_dbllistats
 Set rstbob = Nothing
 Set rstforats = Nothing
 If contador > 0 Then MsgBox "He trobat incidencies en els forats hi han " + atrim(contador) + " bobines marcades per revisar-les", vbCritical + vbOKOnly, "Atenció"
fi:
 ratoli "normal"
End Sub
Private Sub mbobinesparcials_Click()
   llistat_parcials False

End Sub
Function DetallBobinaiPalet(vpalet As Double, vbobina As Double, vmetres As Double) As String
   Dim rst As Recordset
   Dim vsql As String
   Dim vdiametre As Double
   vdiametre = bobinesdentrada.calcular_diametre(vpalet, vbobina)
   vsql = "SELECT bobines.sit,palets.ample,palets.micres,Palets.Idpalet, Bobines.Idbobina, Palets.dataaltapalet, proveidors.nom, materials.descripcio as DescripcioMat, familiesmaterials.descripcio as DescripcioFam, subfamiliesmaterials.descripcio as DescripcioSubFam, familiescolorants.descripcio as DescripcioFamCol, subfamiliescolorants.descripcio as DescripcioSubFamCol "
   vsql = vsql + " FROM (((((materials RIGHT JOIN (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) ON materials.codi = Palets.codimatprognou) LEFT JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi) LEFT JOIN familiescolorants ON materials.familiacol = familiescolorants.codi) LEFT JOIN subfamiliescolorants ON materials.subfamiliacol = subfamiliescolorants.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi) LEFT JOIN familiesmaterials ON materials.familia = familiesmaterials.codi "
   vsql = vsql + " where palets.idpalet=" + atrim(vpalet) + " and bobines.idbobina=" + atrim(vbobina)
   Set rst = dbtmp.OpenRecordset(vsql)
   DetallBobinaiPalet = atrim(rst!sit) + ";" + atrim(vpalet) + ";" + atrim(vbobina) + ";" + atrim(vdiametre) + ";" + atrim(rst!dataaltapalet) + ";" + atrim(rst!nom) + ";"
   DetallBobinaiPalet = DetallBobinaiPalet + atrim(rst!DescripcioFam) + ";" + atrim(rst!DescripcioSubFam) + ";" + atrim(rst!DescripcioFamCol) + ";" + atrim(rst!DescripcioSubFamCol) + ";" + atrim(rst!ample) + ";" + atrim(vmetres) + ";" + atrim(rst!micres)
   Set rst = Nothing
End Function
Sub generartaulallistatperbaixarIMP()
   Dim rst As Recordset
   Dim rstp As Recordset
   Dim mtrs As Double
   Dim data As Date
   Dim comandaactual As Double
   Dim borrar As Boolean
   Dim esurgent As Boolean
   
   data = Now
   Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", fitxerini))
   Set dbstocks = palets.Database
   borrartaulallistatperpujar
   possarcomandesalallistaperbaixar
   
   dbstocks.Execute "SELECT Parcials.idpalet, Parcials.idbobina, Bobines.Sit AS ample, Bobines.disponible AS metres, Bobines.Sit, CDbl(parcials.comanda) AS comanda INTO llistatperpujar IN '" + nomfitxertemporal + "' FROM (Bobines INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)) INNER JOIN comandes ON cdbl(Parcials.comanda) = comandes.comanda WHERE comandes.proximaseccio='I' and (((CDbl([parcials].[comanda])) In (select comanda from llistatperbaixar)));"
   dbllistat.Execute "alter table llistatperpujar add column nomproveidor text(30)"
   If arguments(2) = "temporalTORERUS" Then
      dbllistat.Execute "delete * from llistatperpujar where comanda not in (" + llistacomandesvalides + ")"
   End If
   Set rst = dbllistat.OpenRecordset("select * from llistatperpujar order by comanda")
   borrar = False
   If Not rst.EOF Then dbllistat.Execute "update llistatperpujar set ample='',metres=0"
principi:
   While Not rst.EOF
     If comandaactual <> rst!comanda Then
        comandaactual = rst!comanda
        borrar = True
     End If
    'If (Mid(UCase(rst!sit), 1, 2)) <> "IM" Then
    '   borrar = False
    '     Else
           ''Set rstp = dbstocks.OpenRecordset("SELECT Bobines.Idpalet, Parcials.comanda, Bobines.Sit FROM Bobines INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet) where parcials.idpalet=" + atrim(rst!idpalet) + " and cdbl(parcials.comanda)=" + atrim(rst!comanda))
           Set rstp = dbstocks.OpenRecordset("SELECT Bobines.Idpalet, Parcials.comanda, Bobines.Sit FROM Bobines INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet) where cdbl(parcials.comanda)=" + atrim(rst!comanda))
           While Not rstp.EOF
              If Mid(UCase(rstp!sit), 1, 2) <> "IM" Or atrim(rstp!sit) = "" Then borrar = False
              rstp.MoveNext
           Wend
    'End If
    Set rstp = dbstocks.OpenRecordset("select * from parcials where idpalet=" + atrim(rst!idpalet) + " and idbobina=" + atrim(rst!idbobina) + " and cdbl(comanda)=" + atrim(rst!comanda) + " and not utilitzada  ")
    If rstp.EOF Then
       rst.Delete
       GoTo proxima
    End If
    
    'If UCase(Mid(rst!sit, 1, 2)) = "IM" Or UCase(Mid(rst!sit, 1, 2)) = "LA" Or UCase(Mid(rst!sit, 1, 2)) = "RE" Or UCase(Mid(rst!sit, 1, 2)) = "SO" Then
    '   rst.Delete
    '   GoTo proxima
    'End If
    Set rstp = dbstocks.OpenRecordset("select * from palets where idpalet=" + atrim(rst!idpalet))
    If rstp.EOF Then GoTo proxima
    If borrar Then
       dbllistat.Execute "delete * from llistatperpujar where comanda=" + atrim(comandaactual)
       Set rst = dbllistat.OpenRecordset("select * from llistatperpujar order by comanda")
       rst.FindFirst "comanda>=" + atrim(comandaactual)
       GoTo principi
    End If
    mtrs = bobinesdentrada.calcular_mtrsdispreals(rst!idpalet, rst!idbobina)
    If mtrs < 500 Then
      rst.Delete
      GoTo proxima
    End If
    esurgent = mirarsishareclamatcomurgent(comandaactual)
    rst.Edit
    rst!ample = atrim(rstp!ample)
    rst!metres = mtrs
    rst!nomproveidor = IIf(esurgent, "*", "") + Mid(nomproveidor(rst!idpalet), 1, 19)
    rst.Update
    
proxima:
    rst.MoveNext
   Wend
   
   Set rst = Nothing
   Set rstp = Nothing
End Sub
Function mirarsishareclamatcomurgent(vnumc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbbaixes.OpenRecordset("select * from impresores_ordreimpresio where urgent=true and comanda=" + atrim(vnumc))
   If Not rst.EOF Then mirarsishareclamatcomurgent = True
   Set rst = Nothing
End Function
Private Sub mbobperbaixarimp_Click()
 If sestaactualitzanttorerus Then Exit Sub
   generartaulallistatperbaixarIMP
   wait (2)
      'imprimir llistat
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistatbobinesperpujar.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.SortFields(0) = "+{llistatperpujar.comanda}"
 llistat.SortFields(1) = "+{llistatperpujar.sit}"
 llistat.DiscardSavedData = True
 llistat.Formulas(1) = "titol='LListat de bobines per BAIXAR a IMP.'"
 llistat.Formulas(0) = ""
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 llistat.Formulas(4) = ""
 llistat.Formulas(5) = ""
 llistat.Formulas(6) = ""
 llistat.Formulas(7) = ""
 llistat.Formulas(8) = ""
 llistat.Formulas(9) = ""
 llistat.Formulas(10) = ""
 llistat.Formulas(11) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If Form1.mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 Set dbbaixes = Nothing
 llistat.SortFields(0) = ""
 llistat.SortFields(1) = ""
 escriure_ini "Torerus", "horaultimaactualitzacio", " ", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
End Sub
Function llistacomandesvalides() As String
   Dim rst As Recordset
   Dim vdies As Byte
   vdies = cadbl(llegir_ini("Torerus", "diesactualitzaciobobs", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"))
   If vdies < 2 Then vdies = 200
   Set rst = dbbaixes.OpenRecordset("select * from impresores_ordreimpresio where dataprevistaimpresiocalculada<dateadd('d'," + atrim(vdies) + ",now)")
   While Not rst.EOF
      llistacomandesvalides = llistacomandesvalides + IIf(llistacomandesvalides <> "", ",", "") + atrim(rst!comanda)
      rst.MoveNext
   Wend
   If llistacomandesvalides = "" Then llistacomandesvalides = "0"
   Set rst = Nothing
End Function
Sub generartaulallistatperpujarIMP()
   Dim rst As Recordset
   Dim rstp As Recordset
   Dim mtrs As Double
   Set dbstocks = palets.Database
   borrartaulallistatperpujar
   dbstocks.Execute "SELECT Parcials.Idpalet, Parcials.Idbobina, First(Bobines.Sit) AS ample, First(Bobines.disponible) AS metres,'' as sit, cdbl(first(parcials.comanda)) as comanda into llistatperpujar IN '" + nomfitxertemporal + "' FROM Bobines INNER JOIN Parcials ON (Bobines.Idpalet = Parcials.idpalet) AND (Bobines.Idbobina = Parcials.idbobina) GROUP BY Parcials.Idpalet, Parcials.Idbobina HAVING (((First(Bobines.Sit)) Like '*IMP*') AND ((First(Bobines.disponible))>150));"
   dbllistat.Execute "alter table llistatperpujar add column nomproveidor text(30)"
   
   Set rst = dbllistat.OpenRecordset("select * from llistatperpujar")
   If Not rst.EOF Then
      dbllistat.Execute "update llistatperpujar set ample='',metres=0"
   End If
   While Not rst.EOF
    Set rstp = dbstocks.OpenRecordset("select * from parcials where idpalet=" + atrim(rst!idpalet) + " and idbobina=" + atrim(rst!idbobina) + " and utilitzada=false  ")
    If Not rstp.EOF Then
       rst.Delete
       GoTo proxima
    End If
    Set rstp = dbstocks.OpenRecordset("select * from palets where idpalet=" + atrim(rst!idpalet))
    If rstp.EOF Then GoTo proxima
    mtrs = bobinesdentrada.calcular_mtrsdispreals(rst!idpalet, rst!idbobina)
    If mtrs < 500 Then rst.Delete: GoTo proxima
    rst.Edit
     rst!ample = atrim(rstp!ample)
     rst!metres = mtrs
     rst!nomproveidor = Mid(nomproveidor(rst!idpalet), 1, 20)
    rst.Update
    
proxima:
    rst.MoveNext
   Wend
   Set rst = Nothing
   Set rstp = Nothing
End Sub

Sub generartaulallistatperpujarLAM()
   Dim rst As Recordset
   Dim rstp As Recordset
   Dim mtrs As Double
   Set dbstocks = palets.Database
   borrartaulallistatperpujar
   dbstocks.Execute "SELECT Parcials.Idpalet, Parcials.Idbobina, First(Bobines.Sit) AS ample, First(Bobines.disponible) AS metres,'' as sit, cdbl(first(parcials.comanda)) as comanda into llistatperpujar IN '" + nomfitxertemporal + "' FROM Bobines INNER JOIN Parcials ON (Bobines.Idpalet = Parcials.idpalet) AND (Bobines.Idbobina = Parcials.idbobina) GROUP BY Parcials.Idpalet, Parcials.Idbobina HAVING (((First(Bobines.Sit)) Like '*LAM*') AND ((First(Bobines.disponible))>150));"
   dbllistat.Execute "alter table llistatperpujar add column nomproveidor text(30)"
   Set rst = dbllistat.OpenRecordset("select * from llistatperpujar")
   If Not rst.EOF Then dbllistat.Execute "update llistatperpujar set ample='',metres=0"
   While Not rst.EOF
    Set rstp = dbstocks.OpenRecordset("select * from parcials where idpalet=" + atrim(rst!idpalet) + " and idbobina=" + atrim(rst!idbobina) + " and utilitzada=false  ")
    If Not rstp.EOF Then
       rst.Delete
       GoTo proxima
    End If
    Set rstp = dbstocks.OpenRecordset("select * from palets where idpalet=" + atrim(rst!idpalet))
    If rstp.EOF Then GoTo proxima
    mtrs = bobinesdentrada.calcular_mtrsdispreals(rst!idpalet, rst!idbobina)
    If mtrs < 500 Then rst.Delete: GoTo proxima
    rst.Edit
     rst!ample = atrim(rstp!ample)
     rst!metres = mtrs
     rst!nomproveidor = Mid(nomproveidor(rst!idpalet), 1, 20)
    rst.Update
    
proxima:
    rst.MoveNext
   Wend
   Set rst = Nothing
   Set rstp = Nothing
End Sub


Function sestaactualitzanttorerus() As Boolean
   Dim resp As String
   resp = atrim(llegir_ini("Torerus", "horaultimaactualitzacio", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"))
   If resp = "" Or resp = "{[}]" Then
      GoTo fi
   End If
   If DateDiff("n", resp, Now) < 4 Then
      If arguments(2) <> "temporalTORERUS" Then
         MsgBox "S'està actualitzant una tablet espera un minut i torna-ho a provar.", vbCritical, "Actualitzant Torerus"
      End If
      sestaactualitzanttorerus = True
      Exit Function
   End If
fi:
   escriure_ini "Torerus", "horaultimaactualitzacio", atrim(Now), rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
End Function
Private Sub mbobperpujarimp_Click()
  If sestaactualitzanttorerus Then Exit Sub
   generartaulallistatperpujarIMP
   wait (2)
      'imprimir llistat
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistatbobinesperpujar.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 llistat.Formulas(1) = "titol='LListat de bobines per PUJAR.'"
 llistat.Formulas(0) = ""
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 llistat.Formulas(4) = ""
 llistat.Formulas(5) = ""
 llistat.Formulas(6) = ""
 llistat.Formulas(7) = ""
 llistat.Formulas(8) = ""
 llistat.Formulas(9) = ""
 llistat.Formulas(10) = ""
 llistat.Formulas(11) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If Form1.mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 escriure_ini "Torerus", "horaultimaactualitzacio", " ", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
End Sub

Private Sub mbuscarbobproveidor_Click()
   Dim vsql As String
   Dim vnumbob As String
   Dim rst As Recordset
   If palets.Recordset.EditMode > 0 Then MsgBox "Estas editant. Primer finalitza l'edicio.": Exit Sub
   vnumbob = InputBox("Escriu el numero de bobina del proveidor que vols buscar", "Buscar Nº bobina proveidor")
   vsql = "SELECT Bobines.Idpalet, Bobines.Idbobina, Bobines.Numbobina From bobines WHERE (((Bobines.Numbobina)='" + vnumbob + "'));"
   Set rst = dbtmp.OpenRecordset(vsql)
   If Not rst.EOF Then
      palets.Recordset.FindFirst "idpalet=" + atrim(rst!idpalet)
      If Not palets.Recordset.NoMatch Then
          bobines.Recordset.FindFirst "idbobina=" + atrim(rst!idbobina)
      End If
        Else: MsgBox "No he localitzat aquest numero de bobina"
   End If
   Set rst = Nothing
End Sub

Private Sub MCANVISSITUACIO_Click()
   If InputBoxEx("Entra la contrasenya d'acces a canvis de situació.", "Canvis de situació", , , , , , SPassword) = "inplacSA123" Then
      ensenyarcanvisdesituacio
       Else: MsgBox "Contrasenya equivocada.", vbCritical, "Error"
   End If
End Sub
Private Sub membpalets_Click()
   If InputBoxEx("Entra la contrasenya d'acces a control embolicar.", "Control embolicar", , , , , , SPassword) = "inplacSA123" Then
      ensenyarcontrolembolicar
       Else: MsgBox "Contrasenya equivocada.", vbCritical, "Error"
   End If
End Sub

Sub ensenyarcontrolembolicar()
   Load formseleccio
      formseleccio.Caption = "Control embolicar bobines."
      formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "Vendes.mdb"
      formseleccio.Data1.RecordSource = "SELECT numcomanda,numpalet,data,operari from embolicarpalets order by data desc,operari DESC"
      formseleccio.refrescar
      
      formseleccio.Width = 10000
      formseleccio.DBGrid2.Columns(0).Width = 1000
      formseleccio.DBGrid2.Columns(1).Width = 800
      formseleccio.DBGrid2.Columns(2).Width = 2000
      formseleccio.DBGrid2.Columns(3).Width = 800
      formseleccio.Command3.Tag = "filtre"
      formseleccio.Command2.Tag = "3"
      formseleccio.bexportar.Visible = True
      formseleccio.Left = Screen.Width / 2 - (formseleccio.Width / 2)
      formseleccio.Show 1
      
      Unload formseleccio
End Sub
Sub ensenyarcanvisdesituacio()
   Load formseleccio
      formseleccio.Caption = "Canvis de situació de les Bobines"
      formseleccio.Data1.DatabaseName = camistock
      formseleccio.Data1.RecordSource = "SELECT Bobina,sitorigen,sitdesti,data,operari from canvissituacio order by data DESC"
      formseleccio.refrescar
      
      formseleccio.Width = 10000
      formseleccio.DBGrid2.Columns(0).Width = 1000
      formseleccio.DBGrid2.Columns(1).Width = 1000
      formseleccio.DBGrid2.Columns(2).Width = 1000
      formseleccio.DBGrid2.Columns(3).Width = 2000
      formseleccio.DBGrid2.Columns(4).Width = 2000
      formseleccio.Command3.Tag = "filtre"
      formseleccio.Command2.Tag = "1"
      formseleccio.bexportar.Visible = True
      formseleccio.Left = Screen.Width / 2 - (formseleccio.Width / 2)
      formseleccio.Show 1
      
      Unload formseleccio
End Sub
Private Sub mcontrolestocdeseguretat_Click()
   formestocseguretat.Show 1
End Sub

Private Sub mdevmat_Click()
  formdevolucions.Show 1
End Sub

Private Sub mdisponiblegrups_Click()
  llistat_inventari_disponible
  llistat_inventari_disponibleNOMESDEGRUPS
  llistat_inventari_disponible_PALETSNODISPONIBLES
End Sub

Private Sub mdisponiblegrupsdetallcompres_Click()
   llistat_inventari_disponible "compres"
End Sub

Private Sub mllistaperpantalla_Click()
  mllistaperpantalla.Checked = Not mllistaperpantalla.Checked
End Sub

Private Sub mllistatalbaranscompres_Click()
   albaranscompres.llistatalbaranscompresentredates
End Sub

Private Sub mllistatalbaransvsfacturesSAP_Click()
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim datainici As String
  Dim datafi As String
  Dim vcodimaterial As Double
  datainici = InputBox("Entra la data d'inici del llistat", "Inici")
  If Not IsDate(datainici) Then MsgBox "Data no valida", vbCritical, "Error": Exit Sub
  datafi = InputBox("Entra la data fi del llistat", "Inici")
  If Not IsDate(datafi) Then MsgBox "Data no valida", vbCritical, "Error": Exit Sub
  vcodimaterial = cadbl(escullirmaterial)
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "Llistat Albarans vs FacturesSAP.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "connexiosap.mdb"
  oreport.FormulaFields.GetItemByName("dates").Text = "'Inici: " + format(datainici, "dd/mm/yy") + " -> Fi: " + format(datafi, "dd/mm/yy") + "'"
  oreport.RecordSelectionFormula = "{@datadeldocument}>=#" + format(datainici, "mm/dd/yy") + "# and {@datadeldocument}<=#" + format(datafi, "mm/dd/yy") + "#"
 'oreport.RecordSelectionFormula = "{AlbaransVsFacturesSAP.DataAlbara}>=#" + Format(datainici, "mm/dd/yy") + "# and {AlbaransVsFacturesSAP.DataAlbara}<=#" + Format(datafi, "mm/dd/yy") + "#"
  oreport.DiscardSavedData
  Load veurereport
  veurereport.CRViewer.ReportSource = oreport

  veurereport.CRViewer.ViewReport
  veurereport.Show 1
End Sub

Private Sub mllistatdiametre_Click()
  Dim rst As Recordset
  Set dbstocks = dbtmp
  If MsgBox("Imprimiré totes les etiquetes de bobina parcials que s'ha ajustat els metres mitjantçant el diametre.", vbExclamation + vbDefaultButton2 + vbYesNo, "Imprimir etiqueta del parcial") = vbNo Then Exit Sub
  Set rst = dbstocks.OpenRecordset("select * from comprovacio_diametres_picus where etiquetareimpresa=false")
  While Not rst.EOF
    bobinesdentrada.imprimir_bobinaparcial rst!numpalet, rst!bobina, True
    wait 2
    rst.MoveNext
  Wend
  If MsgBox("S'han impres totes les etiquetes de bobina?", vbExclamation + vbDefaultButton2 + vbYesNo, "Etiquetes bobines parcials") = vbYes Then
      dbstocks.Execute "update comprovacio_diametres_picus set etiquetareimpresa=true where etiquetareimpresa=false"
  End If
  Set rst = Nothing
End Sub

Private Sub mnpapersfrontals_Click()
    Dim horaentrada As Date
  ' On Error Resume Next
   horaentrada = Now
   paperfrontal.Show 1
   If DateDiff("s", horaentrada, Now) < 2 Then
     Shell ("msiexec /i \\serverprodu\dades\progcomandes\aplicacio\instalaciotbarcode11\TBarCode_Setup.msi INSTALLDIR=C:\TBarCode11 ADDLOCAL=FeatTBarCode,FeatOCX /qn")
     MsgBox "Es la primera vegada que s'utilitza els Codi de Barres i s'instal.larà el programa per fer-ho espera uns segons i torna-ho a provar.", vbCritical + vbOKOnly, "Atenció": Exit Sub
     Exit Sub
   End If
   
  
End Sub

 Sub modificar_Click()
    editarpalet
End Sub
Sub editarpalet()
  On Error GoTo errors
  If palets.Recordset.EOF Then Exit Sub
  palets.Recordset.Edit
  activarframes True
  If Screen.ActiveForm.Name = "Form1" Then txtFields(0).SetFocus
  Exit Sub
errors:
   MsgBox "Registre bloquejat per algun altre usuari.", vbCritical, "Error"
End Sub

Private Sub mparcialssenseregularitzardiametre_Click()
   llistat_parcials True
End Sub

Private Sub mprestatges_Click()
  mantenimentprestatges.Show 1
End Sub

Private Sub mprestNOUS_Click()
  Formprestatgesnous.Show 1
End Sub

Private Sub msincronitzacionstorerus_Click()
    Load formseleccio
      formseleccio.Caption = "Sincronitzacions de la TABLET"
      formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "palets.mdb"
      ordre = " order by FORMAT(data,'yymmdd') desc ,usuari,FORMAT(data,'hhnn')"
      formseleccio.Data1.RecordSource = "SELECT data,usuari,canvis_bobines from torerus_sincronitzacions " + ordre
      formseleccio.refrescar
      formseleccio.Width = formseleccio.Width + ((formseleccio.Width / 100) * 20)
      formseleccio.DBGrid2.Columns(0).Width = 2000
      formseleccio.DBGrid2.Columns(1).Width = 1500
      formseleccio.DBGrid2.Columns(2).Width = 2000
      formseleccio.Command2.Tag = "2"
      formseleccio.Show 1
   Unload formseleccio
End Sub

Private Sub munabobina_Click()
Dim paletinici As Double
   Dim paletfi As Double
   Dim numbobina As Double
   paletinici = cadbl(InputBox("Escriu el palet.", "Palet", atrim(palets.Recordset!idpalet)))
   paletfi = paletinici
   numbobina = cadbl(InputBox("Escriu el numero de bobina.", "NºBobina", "1"))
   If paletinici > 0 And paletfi > 0 Then imprimiretpalet paletinici, paletfi, numbobina
End Sub

Private Sub packinglistcomanda_Click()
  Dim numcomanda As Double
  numcomanda = cadbl(InputBox("Entra la Comanda que vols imprimir el PACKING-LIST", "Packing-List"))
  If numcomanda > 0 Then imprimir_packinglist numcomanda, llistat, False
End Sub

Sub guardar_registre_taulatmp3(rstpalet As Recordset, rstpro As Recordset, rstbobina As Recordset, rstmaterial As Recordset)
   Dim rstp As Recordset
   
   Set rstp = dbtmp.OpenRecordset("select comanda from parcials where comanda='2500' and idpalet=" + atrim(rstbobina!idpalet) + " and idbobina=" + atrim(rstbobina!idbobina))
   
   crearbmpnumpalet format(rstpalet!idpalet, "#,##0") + "·" + atrim(rstbobina!idbobina)
   'poso el codidebarres
   paperfrontal.codidebarres.Enabled = True
   paperfrontal.codidebarres.BackStyle = BKS_Transparent
   paperfrontal.codidebarres.Text = atrim(rstpalet!idpalet) + "/" + atrim(rstbobina!idbobina)
   paperfrontal.codidebarres.SaveImage "c:\temp\codidebarrespalet", eIMBmp, 12000, 1000, 1200, 1200
   

   rstllistat.AddNew
   copiafoto "c:\temp\numpalet.bmp", rstllistat!numpaletgrosbmp
   copiafoto "c:\temp\codidebarrespalet.bmp", rstllistat!codidebarres
   If Not rstp.EOF Then rstllistat!reserva = "2500"   'faig servir per passar el grup d'anivellar a l'etiqueta
   rstllistat!palet = rstpalet!idpalet
   rstllistat!bobina = rstbobina!idbobina
   rstllistat!ample = rstpalet!ample
   rstllistat!plegat = rstpalet!plegat
   rstllistat!solapa = rstpalet!solapa
   rstllistat!datarec = rstpalet!datarev
   rstllistat!numpaletprov = rstbobina!numpaletpro
   rstllistat!numbobinaprov = atrim(rstbobina!numbobina)
   rstllistat!numlot = rstpalet!numlot
   If cadbl(rstpalet!micres) > 0 Then
       rstllistat!espesor = atrim(Redondejar(rstpalet!micres, 1)) + " " + Chr(181)
        Else: rstllistat!espesor = atrim(Redondejar(cadbl(rstpalet!grmsm2), 1)) + " g/m2"
   End If
   If Not rstmaterial.EOF Then
     rstllistat!material = rstmaterial!descripcio
     rstllistat!materialdelicat = IIf(cabool(rstmaterial!materialdelicat) = True, 1, 0)
     rstllistat!proveidor = rstpro!nom
   End If
   rstllistat!metres = format(rstbobina!mts, "#,##0")
   rstllistat.Update
   
   rstbobina.Edit
   rstbobina!impres = True
   rstbobina.Update
   Set rstp = Nothing
   
End Sub




Private Sub palets_Reposition()
  If Not palets.Recordset.EOF Then
     palets.Caption = "Palets " + atrim(palets.Recordset.AbsolutePosition + 1) + " / " + atrim(palets.Recordset.RecordCount)
     actualitzar_vinculats
     
     calcular_kilos cadbl(palets.Recordset!idpalet)
     bobines.RecordSource = "select * from bobines where idpalet=" + atrim(cadbl(palets.Recordset!idpalet)) + " order by idbobina"
     bobines.Refresh
     
       sumarmetresikilos
  End If
End Sub
Function calcular_kilos_disponibles_palet(palet As Double) As Double
   Dim rstp As Recordset
   Dim kilos As Double
   Dim espesor As Double
   Dim pesprov As Double
   Dim tkilos As Double
   If palet < 1 Then Exit Function
   
   Set rstp = dbtmp.OpenRecordset("select * from bobines where idpalet=" + atrim(palet) + " and disponible>0")
   
   While Not rstp.EOF
      'assignarmat.actualitzar_metres_disponibles rstp!idpalet, rstp!idbobina
      kilos = Redondejar((rstp!kilos / rstp!mts) * rstp!disponible, 0)
      tkilos = tkilos + kilos
      rstp.MoveNext
   Wend
   Set rstp = Nothing
   calcular_kilos_disponibles_palet = tkilos
End Function
Sub calcular_kilos(palet As Double)
   Dim rstp As Recordset
   Dim kilos As Double
   Dim espesor As Double
   Dim pesprov As Double
   Dim tanx100 As Double
   If palet < 1 Then Exit Sub
   
   Set rstp = dbtmp.OpenRecordset("select * from bobines where idpalet=" + atrim(palet))
   
   While Not rstp.EOF
      assignarmat.actualitzar_metres_disponibles rstp!idpalet, rstp!idbobina
      espesor = cadbl(palets.Recordset!micres)
      If grmm2 > 0 Then espesor = grmm2 * -1
      kilos = compramat.conversiokilos(cadbl(palets.Recordset!codimatprognou), cadbl(palets.Recordset!ample), cadbl(rstp!mts), cadbl(espesor), atrim(palets.Recordset!semielaborat), cadbl(palets.Recordset!solapa))
      kilos = Redondejar(kilos, 0)
      pesprov = cadbl(rstp!pesdelproveidor)
      tanx100 = 0
      If kilos > 0 Then
         ' aquest es el metode miralles
         tanx100 = Redondejar((pesprov - kilos) * 100 / kilos, 1)
         'aquest es el metode miquel rabassedas
        ' If kilos > 0 Then tanx100 = Redondejar(((pesprov * 100) / kilos) - 100, 1)
      End If
      dbtmp.Execute ("update bobines set tanx100variaciopes=" + passaradecimalpunt(atrim(tanx100)) + ",kilos=" + passaradecimalpunt(Redondejar(kilos, 0)) + " where idpalet=" + atrim(palet) + " and idbobina=" + atrim(rstp!idbobina))
      rstp.MoveNext
   Wend
   Set rstp = Nothing
   Set rstp = dbtmp.OpenRecordset("select * from bobines where idpalet=" + atrim(palet))
   While Not rstp.EOF
      rstp.MoveNext
   Wend
   Set rstp = Nothing
      DoEvents
End Sub
Sub actualitzar_vinculats()
   Dim rst As Recordset
   Dim codimat As Long
   
   
   If palets.Recordset.EOF Or palets.Recordset.BOF Then Exit Sub
   If palets.Recordset!preucompraavg And txtFields(14).Visible Then
      avg.Caption = "Preu mig."
    Else: avg.Caption = ""
   End If
   txtFields_Change (14)
   codimat = cadbl(palets.Recordset!codimatprognou)
   codimat = cadbl(txtFields(1))
   numerodepalet = format(palets.Recordset!idpalet, "#,##0")
   If IsDate(palets.Recordset!dataaltapalet) Then
      datacreacio = "Data Creació:" + Chr(10) + format(palets.Recordset!dataaltapalet, "dd/mm/yy hh:nn")
     Else: datacreacio = ""
   End If
   Set rst = dbtmpb.OpenRecordset("select descripcio,proveidor,grmcm3,grmm2,refproducte from materials where codi=" + atrim(codimat))
   proveidor = ""
   refprov = ""
   nommaterial = ""
   grmm3 = 0
   grmm2 = 0
   If Not rst.EOF Then
      grmcm3 = cadbl(rst!grmcm3)
      grmm2 = cadbl(rst!grmm2)
      nommaterial = atrim(rst!descripcio)
      refprod = atrim(rst!refproducte)
      Set rst = dbtmpb.OpenRecordset("select codi,nom from proveidors where codi=" + atrim(cadbl(rst!proveidor)))
      If Not rst.EOF Then proveidor = " Proveïdor: " + UCase(rst!nom): proveidor.Tag = atrim(rst!codi)
   End If
   txtFields(13) = atrim(grmm2)
   If txtFields(15) <> "" Then
      If txtFields(9) <> "" Then bdesactivarpalets.Visible = False Else bdesactivarpalets.Visible = True
        Else: bdesactivarpalets.Visible = False
   End If
   If palets.Recordset!teimpost Then
        etimpostenv.Visible = True
        etimpostenv = "Impost Envasos"
         Else: etimpostenv.Visible = False
   End If
   If cadbl(palets.Recordset!idpalet) > 0 Then dbtmp.Execute "update palets set grmsm2=" + atrim(cadbl(grmm2)) + " where idpalet=" + atrim(cadbl(palets.Recordset!idpalet))
End Sub

Private Sub preucompra_Change()
On Error Resume Next
  If Form1.ActiveControl.Name = "preucompra" Then txtFields(14) = preucompra
End Sub

Private Sub preucompra_KeyPress(KeyAscii As Integer)
If matclient.Value <> 0 Then
      MsgBox "Aquest material es del client i no es pot possar preu de cost.", vbCritical + vbOKOnly, "Atenció"
      preucompra = 0
      KeyAscii = 0
   End If
  
End Sub

Private Sub reixaparcials_AfterDelete()
  'assignarmat.actualitzar_metres_disponibles palets.Recordset!idpalet, bobines.Recordset!idbobina
End Sub

Private Sub reixaparcials_BeforeDelete(Cancel As Integer)
  Dim vant As String
  Dim vnou As String
  Dim vid As String
  Dim vpa As String
  Dim vbo As String
  Dim vcaf As String
  Dim vusr As String
  Dim i As Integer
  vid = atrim(parcials.Recordset!id)
  vpa = atrim(palets.Recordset!idpalet)
  vbo = atrim(bobines.Recordset!idbobina)
  vusr = nomordinador

   DBGrid1.Columns("disponible") = cadbl(DBGrid1.Columns("mts")) - (cadbl(metrespartits) - cadbl(reixaparcials.Columns("mtrs")))
   vcaf = "comanda": vant = atrim(parcials.Recordset!comanda): vnou = "Eliminat"
   palets.Database.Execute "insert into Parcials_controlcanvis (idparcial,palet,bobina,campafectat,valoranterior,valoractual,usuari) values (" + atrim(vid) + "," + atrim(vpa) + "," + atrim(vbo) + ",'" + atrim(vcaf) + "','" + atrim(vant) + "','" + treure_apostruf(atrim(vnou)) + "','" + vusr + "')"
   vcaf = "metres": vant = atrim(parcials.Recordset!metres): vnou = "Eliminat"
   palets.Database.Execute "insert into Parcials_controlcanvis (idparcial,palet,bobina,campafectat,valoranterior,valoractual,usuari) values (" + atrim(vid) + "," + atrim(vpa) + "," + atrim(vbo) + ",'" + atrim(vcaf) + "','" + atrim(vant) + "','" + treure_apostruf(atrim(vnou)) + "','" + vusr + "')"
   vcaf = "data": vant = atrim(parcials.Recordset!data): vnou = "Eliminat"
   palets.Database.Execute "insert into Parcials_controlcanvis (idparcial,palet,bobina,campafectat,valoranterior,valoractual,usuari) values (" + atrim(vid) + "," + atrim(vpa) + "," + atrim(vbo) + ",'" + atrim(vcaf) + "','" + atrim(vant) + "','" + treure_apostruf(atrim(vnou)) + "','" + vusr + "')"
   vcaf = "operari": vant = atrim(parcials.Recordset!operari): vnou = "Eliminat"
   palets.Database.Execute "insert into Parcials_controlcanvis (idparcial,palet,bobina,campafectat,valoranterior,valoractual,usuari) values (" + atrim(vid) + "," + atrim(vpa) + "," + atrim(vbo) + ",'" + atrim(vcaf) + "','" + atrim(vant) + "','" + treure_apostruf(atrim(vnou)) + "','" + vusr + "')"
End Sub

Private Sub reixaparcials_BeforeUpdate(Cancel As Integer)
 If atrim(reixaparcials.Columns("orcomassignacio")) = "" Then
   reixaparcials.Columns("orcomassignacio") = reixaparcials.Columns("Comanda")
 '  Else: MsgBox "Pensa que si canvies el numero de comanda assignat no es canvia el numero de comanda assignat inicialment. "
 End If
 If cadbl(reixaparcials.Columns("comanda")) = 0 Then parcials.Recordset!data = Null: parcials.Recordset!comanda = "0": reixaparcials.Columns("comanda") = "0"
 If cadbl(reixaparcials.Columns("comanda")) = 300 Or cadbl(reixaparcials.Columns("comanda")) = 100 Or cadbl(reixaparcials.Columns("comanda")) = 400 Then parcials.Recordset!data = format(Now, "dd/mm/yy"): parcials.Recordset!utilitzada = True
 guardar_modificacions_parcials
End Sub
Sub guardar_modificacions_parcials()
  Dim vant As String
  Dim vnou As String
  Dim vid As String
  Dim vpa As String
  Dim vbo As String
  Dim vcaf As String
  Dim vusr As String
  Dim i As Integer
  vid = atrim(parcials.Recordset!id)
  vpa = atrim(palets.Recordset!idpalet)
  vbo = atrim(bobines.Recordset!idbobina)
  vusr = nomordinador
  For i = 0 To reixaparcials.Columns.Count - 1
     vcaf = treure_apostruf(atrim(reixaparcials.Columns(i).DataField))
     vant = atrim(parcials.Recordset.Fields(reixaparcials.Columns(i).DataField))
     If parcials.Recordset.Fields(reixaparcials.Columns(i).DataField).Type = 8 Then vant = format(vant, "dd/mm/yy")
     If parcials.Recordset.Fields(reixaparcials.Columns(i).DataField).Type = 1 Then vant = IIf(parcials.Recordset.Fields(reixaparcials.Columns(i).DataField), "Sí", "No")
     vnou = atrim(reixaparcials.Columns(reixaparcials.Columns(i).DataField))
     If vant <> vnou And Mid(vcaf + "   ", 1, 2) <> "id" Then
        palets.Database.Execute "insert into Parcials_controlcanvis (idparcial,palet,bobina,campafectat,valoranterior,valoractual,usuari) values (" + atrim(vid) + "," + atrim(vpa) + "," + atrim(vbo) + ",'" + atrim(vcaf) + "','" + atrim(vant) + "','" + treure_apostruf(atrim(vnou)) + "','" + vusr + "')"
     End If
  Next i
End Sub
Private Sub reixaparcials_Change()
'If reixaparcials.Columns(reixaparcials.col).DataField = "comanda" Then
'   If cadbl(reixaparcials.Columns(reixaparcials.col)) = 0 Then reixaparcials.Columns(reixaparcials.col) = "": MsgBox "Comanda no pot ser zero"
'End If
  If reixaparcials.Columns(reixaparcials.col).DataField = "metres" Then
     If comprovar_quenopasidemetres Then MsgBox "Massa metres, superes el total de bobina disponible": reixaparcials.Columns("metres") = atrim(cadbl(bobines.Recordset!mts) - (cadbl(metrespartits) - cadbl(reixaparcials.Columns("metres"))))
     DBGrid1.Columns("disponible") = cadbl(DBGrid1.Columns("mts")) - cadbl(metrespartits)
  End If
  If reixaparcials.Columns(reixaparcials.col).DataField = "seccio" Then
     If InStr(1, "EILRSTV", UCase(reixaparcials.Columns("seccio"))) = 0 Then
         reixaparcials.Columns("seccio") = ""
        Else: reixaparcials.Columns("seccio") = UCase(reixaparcials.Columns("seccio"))
     End If
     If reixaparcials.Columns("seccio") = "" Then
       reixaparcials.Columns("utilitzada") = "False"
         Else: reixaparcials.Columns("utilitzada") = "True"
     End If
  End If
End Sub
Function comprovar_quenopasidemetres() As Boolean
  If metrespartits > cadbl(bobines.Recordset!mts) Then
     comprovar_quenopasidemetres = True
  End If
End Function

Private Sub reixaparcials_OnAddNew()
 If DBGrid1.EditActive Then DBGrid1.EditActive = False
  If Not bobines.Recordset.EOF Then
   reixaparcials.Columns("idpalet") = palets.Recordset!idpalet
   reixaparcials.Columns("idbobina") = bobines.Recordset!idbobina
  End If
End Sub

Function metrespartits() As String
  Dim clonepar As Recordset
  Dim metres As Double
  metres = 0
  'parcials.UpdateRecord
  Set clonepar = parcials.Recordset.Clone
  If Not clonepar.EOF Then
    clonepar.MoveFirst
    While Not clonepar.EOF
      If parcials.Recordset!id <> clonepar!id Then
        metres = metres + cadbl(clonepar!metres)
      End If
      clonepar.MoveNext
    Wend
  End If
  metrespartits = metres + cadbl(reixaparcials.Columns("metres"))
  Set clonepar = Nothing
End Function

Private Sub sortir_Click()
  End
End Sub

Private Sub Timer1_Timer()
  If palets.Recordset.EditMode = 1 Then estatedicio.Caption = "Editant..."
  If palets.Recordset.EditMode = 2 Then estatedicio.Caption = "Afegint..."
  If buscant Then estatedicio.Caption = "Buscant..."
  If palets.Recordset.EditMode = 0 Then
     estatedicio.Caption = ""
     If framepalets.Enabled Then activarframes False
  End If
  mirarsiparar
End Sub

Private Sub tractat_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub tractat_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub txtFields_Change(Index As Integer)
   On Error Resume Next
  If Form1.ActiveControl.Name <> "preucompra" Then If Index = 14 Then preucompra = format(txtFields(14), "0.000")
  If Form1.ActiveControl.Name = "txtFields" And Index = 9 Then
     If IsDate(txtFields(9)) Then
        chkFields(14).Value = 1
       Else: chkFields(14).Value = 0
     End If
  End If
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
 ' If Index = 9 And Len(txtFields(8)) < 8 Then
 '    MsgBox "Falta la data d'Albarà de proveidor.", vbCritical + vbOKOnly, "Atenció"
 '    txtFields(9).Locked = True
 '      Else: txtFields(9).Locked = False
 ' End If
End Sub

Private Sub txtFields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim ordre As String
  If KeyCode = 113 And Index = 1 Then
  
      Load formseleccio
      formseleccio.Caption = "Escull Material del Palet"
      formseleccio.Data1.DatabaseName = cami
      ordre = " order by proveidor,descripcio"
      formseleccio.Data1.RecordSource = "SELECT materials.codi as [Codi], materials.descripcio as [Descripcio], materials.refproducte as [RefProducte], proveidors.nom as [Proveidor] FROM materials LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE (((materials.codi)>499)) " + ordre
      formseleccio.refrescar
      formseleccio.Width = formseleccio.Width + ((formseleccio.Width / 100) * 20)
      formseleccio.DBGrid2.Columns(0).Width = 500
      formseleccio.DBGrid2.Columns(1).Width = 2500
      formseleccio.DBGrid2.Columns(2).Width = 1000
      formseleccio.DBGrid2.Columns(3).Width = 1500
      formseleccio.Command2.Tag = "2"
      formseleccio.Show 1
      noucodimat = 0
      If seleccioret = 1 Then
         txtFields(1) = atrim(formseleccio.Data1.Recordset!codi)
         nommaterial = atrim(formseleccio.Data1.Recordset!descripcio)
      End If
      Unload formseleccio
  End If
   
  If Index = 9 And KeyCode < 110 Then
    KeyCode = 0
    MsgBox "La data de recepció s'ha d'entrar activant els palets desde el menu Compres-Activació de palets", vbInformation, "Atenció"
    
    
  End If
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  If Chr(KeyAscii) = "." And (Index = 2 Or Index = 3 Or Index = 4 Or Index = 11 Or Index = 14) Then
     KeyAscii = 0
     SendKeys (",")
  End If
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
  If Index = 14 And palets.Recordset.EditMode > 0 Then txtFields(14).Text = format(txtFields(14).Text, "0.000")
  actualitzar_vinculats
End Sub

